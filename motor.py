"""
Motor de cálculo do Simulador de Margens (réplica da aba SIMULADOR em Python).

Cascata (R$/t), igual à planilha:
  Preço Bruto
  - Impostos - Descontos - Devoluções - Rapel         (Incentivo é exibido, mas NÃO soma: "s/ Incentivo")
  = Receita Líquida
  - Frete Vendas - Frete Devolução - Frete Transferência - Chapa - Take or Pay - Outros Custos Log.
  = Preço FOB (FOB Líquido)
  - Custo Variável
  = Margem de Contribuição (MCU)
  - Custo Fixo - Custo de Servir (Força de Vendas + Comissões + Marketing + Logísticos + Movimentação CDs)
  = Margem de Servir (MGS)
  - Capex
  = Margem Caixa (MGCX)

Valores ausentes ficam como None ("n/e" na planilha) e propagam para os totais.
"""

from __future__ import annotations

VERSAO_APP = "2026-11-11.2"

from dataclasses import dataclass, field
from typing import Optional

import pandas as pd

from bases import Bases, norm
from resultados import CANAL_VELHO, info_produto

LINHAS = [
    # (nome, tipo)  tipo: "linha" = editável | "total" = calculado | "info" = só exibido
    ("Preço Bruto", "linha"),
    ("Impostos", "linha"),
    ("Descontos", "linha"),
    ("Devoluções", "linha"),
    ("Rapel", "linha"),
    ("Incentivo", "info"),
    ("Receita Líquida", "total"),
    ("Frete de Vendas", "linha"),
    ("Frete Devolução", "linha"),
    ("Frete Transferência", "linha"),
    ("Chapa", "linha"),
    ("Take or Pay", "linha"),
    ("Outros Custos Log.", "linha"),
    ("Preço FOB", "total"),
    ("Custo Variável", "linha"),
    ("Margem de Contribuição", "total"),
    ("Custo Fixo", "linha"),
    ("Custo de Servir", "total"),
    ("Força de Vendas", "linha"),
    ("Comissões", "linha"),
    ("Marketing", "linha"),
    ("Logísticos: Paletes e Strech", "linha"),
    ("Movimentação de CDs", "linha"),
    ("Margem de Servir", "total"),
    ("Capex", "linha"),
    ("Margem Caixa", "total"),
]
EDITAVEIS = [n for n, t in LINHAS if t != "total"]
CANAIS = ["Fábrica", "CD", "DT", "Broker"]


@dataclass
class Entrada:
    fabrica: str
    cidade: str                    # UF_Cidade, ex. "PR_Andirá"
    canal: str                     # canal (nomes da base nova ou da jun/23)
    produto: Optional[int]         # código do produto (bases jun/23); pode ser None
    produto_nome: str = ""         # nome do produto na base de resultados (ex. "CIMENTO CAUE ... 40KG")
    meses: Optional[list] = None   # meses da base de resultados (AAAAMM); None = todos
    incoterm: str = "CIF"          # CIF ou FOB
    sacos: float = 0
    zona: str = ""                 # Zona de Transporte (padrão BI); vazio = usa a cidade
    preco_negociado: Optional[float] = None   # R$/saco (ensacado) ou R$/t (granel/big bag)
    chapa: bool = False
    chapa_valor: Optional[float] = None       # R$/saco; None = valor da fábrica (DE-PARA)
    fracionamento: str = "Não"
    cd: str = ""
    taxa_filial: Optional[float] = None       # R$/saco; None = tabela Taxa Filial
    preco_broker: Optional[float] = None      # R$/saco; None = histórico
    # ajustes manuais (células cinza "ALTERAR"): {"CD": {"Custo Fixo": 70.0}, ...} em R$/t
    ajustes: dict = field(default_factory=dict)
    completar: bool = True    # completa lacunas sem histórico com médias / frete cadastrado
    modo_excel: bool = False  # reproduz a planilha Excel jun/23 (inclusive as fórmulas que zeram linhas), p/ conferência


# ---------------------------------------------------------------- helpers
_NORM_CACHE = {}


def _coln(df: pd.DataFrame, col: str) -> pd.Series:
    """Coluna normalizada (maiúsculas, sem espaços nas pontas), calculada uma vez por tabela."""
    k = (id(df), col)
    if k not in _NORM_CACHE:
        _NORM_CACHE[k] = df[col].map(norm)
    return _NORM_CACHE[k]


def _v(df: pd.DataFrame, chave: str, coluna: str) -> Optional[float]:
    try:
        x = df.at[chave, coluna]
    except KeyError:
        return None
    if isinstance(x, pd.Series):
        x = x.iloc[0]
    try:
        x = float(x)
    except (TypeError, ValueError):
        return None
    return None if pd.isna(x) else x


def _primeira(df: pd.DataFrame, coluna_filtro: str, valor: str, coluna: str) -> Optional[float]:
    m = df[_coln(df, coluna_filtro) == norm(valor)]
    if m.empty:
        return None
    x = m.iloc[0][coluna]
    return None if pd.isna(x) else float(x)


def _frete_planilha(df: pd.DataFrame, chave_sem_inc: str) -> Optional[float]:
    """Como o PROCV da planilha: 1ª linha da tabela de frete com a chave sem incoterm (às vezes a FOB = 0)."""
    for k in df.index:
        if k.startswith(chave_sem_inc) and k[len(chave_sem_inc):] in ("CIF", "FOB"):
            return _v(df, k, "Total")
    return None


def _soma(*xs):
    if any(x is None for x in xs):
        return None
    return sum(xs)


def _ou(*xs):
    """Primeiro valor não vazio (equivale aos SEERRO/SE "n/e" em cadeia)."""
    for x in xs:
        if x is not None:
            return x
    return None


def _desc(df, k):
    """Descontos = Descontos + Provisões + Estorno Provisão (conforme Glossário)."""
    d = _v(df, k, "Descontos R$/t")
    if d is None:
        return None
    return d + (_v(df, k, "Provisões R$/t") or 0) + (_v(df, k, "Estorno Provisão R$/t") or 0)


def _mkt_col(df):
    return "Marketing e Propaganda R$/t" if "Marketing e Propaganda R$/t" in df.columns else "Marketing R$/t"


def cascata(l: dict) -> dict:
    """Recebe as linhas (R$/t) e calcula os totais."""
    r = dict(l)
    r["Receita Líquida"] = _soma(l["Preço Bruto"], *[-x if x is not None else None for x in
                                  (l["Impostos"], l["Descontos"], l["Devoluções"], l["Rapel"])])
    fretes = [l[n] for n in ("Frete de Vendas", "Frete Devolução", "Frete Transferência",
                             "Chapa", "Take or Pay", "Outros Custos Log.")]
    r["Preço FOB"] = None if r["Receita Líquida"] is None or None in fretes else r["Receita Líquida"] - sum(fretes)
    r["Margem de Contribuição"] = None if None in (r["Preço FOB"], l["Custo Variável"]) else r["Preço FOB"] - l["Custo Variável"]
    serv = [l[n] for n in ("Força de Vendas", "Comissões", "Marketing",
                           "Logísticos: Paletes e Strech", "Movimentação de CDs")]
    r["Custo de Servir"] = None if None in serv else sum(serv)
    r["Margem de Servir"] = (None if None in (r["Margem de Contribuição"], l["Custo Fixo"], r["Custo de Servir"])
                             else r["Margem de Contribuição"] - l["Custo Fixo"] - r["Custo de Servir"])
    r["Margem Caixa"] = None if None in (r["Margem de Servir"], l["Capex"]) else r["Margem de Servir"] - l["Capex"]
    return r


def _linhas_hist(df, k, k_cv=None, df_cv=None):
    """Lê as linhas históricas de uma base de custos."""
    g = lambda c: _v(df, k, c)
    return {
        "Preço Bruto": g("Receita Bruta R$/t"),
        "Impostos": g("Impostos R$/t"),
        "Descontos": _desc(df, k),
        "Devoluções": g("Devoluções R$/t"),
        "Rapel": g("Rapel R$/t"),
        "Incentivo": g("Incentivo R$/t"),
        "Frete de Vendas": g("Frete Vendas R$/t"),
        "Frete Devolução": g("Frete Devolução R$/t"),
        "Frete Transferência": g("Frete Transferência R$/t"),
        "Chapa": g("Chapa R$/t"),
        "Take or Pay": g("Take or Pay R$/t"),
        "Outros Custos Log.": g("Outros Custos Log. (SOMA)"),
        "Custo Variável": g("Custo Variável R$/t"),
        "Custo Fixo": g("Custo Fixo R$/t"),
        "Força de Vendas": g("Comercial R$/t"),
        "Comissões": 0.0,
        "Marketing": g(_mkt_col(df)),
        "Logísticos: Paletes e Strech": g("Logísticos: Paletes e Strech (SOMA)"),
        "Movimentação de CDs": g("Movimentação CDs R$/t"),
        "Capex": g("Capex R$/t"),
    }


def _media(df: pd.DataFrame, **filtros) -> Optional[dict]:
    """Média das linhas da base que batem com os filtros (coluna=valor). None se não houver."""
    m = df
    for col, val in filtros.items():
        m = m[_coln(df, col).loc[m.index] == norm(val)]
    if m.empty:
        return None
    linha = m.apply(pd.to_numeric, errors="coerce").mean().to_frame("_MEDIA").T
    return _linhas_hist(linha, "_MEDIA")


def _completar(s: dict, origem: dict, medias: list, nao_completar=("Preço Bruto", "Impostos", "Frete de Vendas")):
    """Preenche linhas vazias com a primeira média disponível (do filtro mais específico ao mais amplo)."""
    for rotulo, med in medias:
        if med is None:
            continue
        for n in EDITAVEIS:
            if s.get(n) is None and n not in nao_completar and med.get(n) is not None:
                s[n] = med[n]
                origem[n] = rotulo


# ---------------------------------------------------------------- simulação
def referencias(b: Bases, e: Entrada) -> dict:
    """Dados do cabeçalho: peso, embalagem, preços históricos, frete comercial, preço lista."""
    cvel = CANAL_VELHO.get(e.canal, e.canal)
    prod = norm(e.produto)
    peso = _v(b.produtos, prod, "Peso") or 50.0
    emb = b.produtos.at[prod, "Embalagem"] if prod in b.produtos.index else "Ensacado"
    nome = b.produtos.at[prod, "Nomenclatura"] if prod in b.produtos.index else str(e.produto)
    if e.produto_nome:
        ip = info_produto(e.produto_nome)
        peso, emb, nome = ip["peso"], ip["embalagem"], e.produto_nome
    ensacado = emb == "Ensacado"
    fab = norm(e.fabrica)
    expedicao = b.fabricas.at[fab, "Expedição"] if fab in b.fabricas.index else e.fabrica
    uf_fab = b.fabricas.at[fab, "UF"] if fab in b.fabricas.index else ""
    chapa_t = _v(b.fabricas, fab, "Valor Chapa R$/t") or 0.0
    zona = e.zona or e.cidade
    cid = norm(e.cidade)
    info_cid = b.cidades.loc[cid] if cid in b.cidades.index else None
    ap = str(info_cid["AP_Unidade"]) if info_cid is not None else "-"
    broker = str(info_cid["Broker"]) if info_cid is not None else "-"
    micro = str(info_cid["Micro IBGE"]) if info_cid is not None else ""

    def preco_hist(origem):
        k = norm(origem, e.cidade, cvel, e.produto, e.incoterm)
        fob = _v(b.preco, k, "Preço FOB R$/saco")
        cif = _v(b.preco, k, "RB R$/saco")
        return fob, cif

    def frete_com(origem):
        # a tabela tem duas colunas "Valor"; a 2ª (positiva, R$/t) é a usada na planilha
        v = _v(b.frete_com, norm(origem, zona, emb), "Valor.1")
        return None if v is None else (v * peso / 1000 if ensacado else v)

    fob_f, cif_f = preco_hist(expedicao)
    fob_c, cif_c = preco_hist(e.cd) if e.cd else (None, None)
    fc_f, fc_c = frete_com(e.fabrica), (frete_com(e.cd) if e.cd else None)
    if e.incoterm == "FOB":
        cif_f = None if fob_f is None or fc_f is None else fob_f + fc_f
        cif_c = None if fob_c is None or fc_c is None else fob_c + fc_c
    else:
        fob_f = None if cif_f is None or fc_f is None else cif_f - fc_f
        fob_c = None if cif_c is None or fc_c is None else cif_c - fc_c

    lista = _v(b.preco_lista, norm(micro, e.produto), "Preço de lista Microrregião")
    taxa = e.taxa_filial if e.taxa_filial is not None else (_v(b.taxa_filial, norm(e.cd, e.produto), "Taxa Filial") or 0.0)
    premio = _v(b.premio_dt, norm(ap), "Prêmio DT (Modelagem)")
    return dict(peso=peso, embalagem=emb, nome=nome, ensacado=ensacado, expedicao=expedicao,
                uf_fabrica=uf_fab, chapa_padrao=chapa_t * peso / 1000, zona=zona, ap=ap, broker=broker,
                preco_fob_hist={"Fábrica": fob_f, "CD": fob_c}, preco_cif_hist={"Fábrica": cif_f, "CD": cif_c},
                frete_comercial={"Fábrica": fc_f, "CD": fc_c}, preco_lista=lista, taxa_filial=taxa,
                premio_ap=premio)


def _sobrepor(h: dict, novo: dict):
    """Usa os valores da base de resultados onde existirem; mantém a base jun/23 no resto (ex. Capex)."""
    for k in h:
        if novo.get(k) is not None:
            h[k] = novo[k]


def simular(b: Bases, e: Entrada, r=None) -> dict:
    """Retorna {canal: DataFrame com colunas Histórico R$/t, Simulação R$/t, Simulação R$/saco}.
    r = base de resultados (resultados.Resultados), opcional; quando informada é a fonte principal do histórico."""
    if e.modo_excel:          # só bases jun/23 e sem preencher lacunas, como na planilha
        from dataclasses import replace as _rp
        r = None
        e = _rp(e, completar=False, produto_nome="")
    ref = referencias(b, e)
    cvel = CANAL_VELHO.get(e.canal, e.canal)
    fonte = {c: "base jun/23" for c in CANAIS}
    origens = {c: {} for c in CANAIS}

    def novo(broker=False, **f):
        if r is None or not e.produto_nome:
            return None
        return r.hist(meses=e.meses, broker=broker, produto=e.produto_nome, **f)

    def rotulo(h):
        return f"base resultados · {h['_vol']:,.0f} t".replace(",", ".")

    peso, zona = ref["peso"], ref["zona"]
    por_t = (lambda p: p * 1000 / peso) if ref["ensacado"] else (lambda p: p)
    chapa_sc = e.chapa_valor if e.chapa_valor is not None else ref["chapa_padrao"]
    chapa_t = por_t(chapa_sc) if e.chapa else 0.0
    adic = _v(b.fracionamento, norm(e.fracionamento), "Adicional") or 0.0
    fob = e.incoterm == "FOB"

    # ---------------- FÁBRICA
    kf = norm(e.fabrica, e.cidade, cvel, e.produto, e.incoterm)
    hf = _linhas_hist(b.fabrica, kf)
    # CV / CF por fábrica+produto; Capex e Transferência por fábrica (como na planilha)
    cv_cf = b.fabrica[_coln(b.fabrica, "Chave 2") == norm(e.fabrica, e.produto)]
    if not cv_cf.empty:
        hf["Custo Variável"] = float(cv_cf.iloc[0]["Custo Variável R$/t"])
        hf["Custo Fixo"] = float(cv_cf.iloc[0]["Custo Fixo R$/t"])
    hf["Capex"] = _primeira(b.fabrica, "Fábrica", e.fabrica, "Capex R$/t")
    if hf["Frete Transferência"] is None:
        hf["Frete Transferência"] = _primeira(b.fabrica, "Fábrica", e.fabrica, "Frete Transferência R$/t")
    frete_f = 0.0 if fob else (_frete_planilha(b.frete_log_fab, norm(e.fabrica, zona, cvel)) if e.modo_excel
                               else _v(b.frete_log_fab, norm(e.fabrica, zona, cvel, e.incoterm), "Total"))
    hf["Frete de Vendas"] = frete_f
    nf = novo(exp=ref["expedicao"], cidade=e.cidade, canal=e.canal, incoterm=e.incoterm)
    if nf:
        _sobrepor(hf, nf)
        frete_f = 0.0 if fob else hf["Frete de Vendas"]
        hf["Frete de Vendas"] = frete_f
        fonte["Fábrica"] = rotulo(nf)
    aliq = None
    if hf["Preço Bruto"] and hf["Impostos"] is not None:
        aliq = hf["Impostos"] / hf["Preço Bruto"]
    if aliq is None:
        uf_cli = e.cidade[:2]
        aliq = _v(b.impostos, norm(ref["uf_fabrica"], uf_cli, cvel), "Alíquota")
    if aliq is None and r is not None:
        # alíquota efetiva da base de resultados: expedição→cidade, depois expedição→UF, depois canal→cidade
        for filtros in ({"exp": ref["expedicao"], "cidade": e.cidade}, {"exp": ref["expedicao"], "canal": e.canal},
                        {"cidade": e.cidade, "canal": e.canal}):
            h = r.hist(meses=e.meses, **filtros)
            if h and h["Preço Bruto"]:
                aliq = h["Impostos"] / h["Preço Bruto"]
                origens["Fábrica"]["Impostos"] = "alíquota efetiva da base resultados"
                break

    sf = dict(hf)
    sf["Preço Bruto"] = None if e.preco_negociado is None else por_t(e.preco_negociado) + chapa_t
    sf["Impostos"] = None if sf["Preço Bruto"] is None or aliq is None else sf["Preço Bruto"] * aliq
    for n in ("Descontos", "Devoluções", "Rapel", "Incentivo", "Frete Devolução", "Take or Pay", "Outros Custos Log."):
        sf[n] = sf[n] or 0.0
    sf["Frete de Vendas"] = 0.0 if fob else (None if frete_f is None else frete_f * (1 + adic))
    sf["Chapa"] = chapa_t
    sf["Frete Devolução"] = 0.0          # planilha: "valor fixado em zero, alterável"

    # ---------------- CD
    hc = {n: None for n in hf}
    sc = None
    if e.cd:
        kc = norm(e.cd, e.cidade, cvel, e.produto, e.incoterm)
        hc = _linhas_hist(b.cd, kc)
        frete_c = 0.0 if fob else (_frete_planilha(b.frete_log_cd, norm(e.cd, zona, cvel)) if e.modo_excel
                                   else _v(b.frete_log_cd, norm(e.cd, zona, cvel, e.incoterm), "Total"))
        hc["Frete de Vendas"] = frete_c
        transf = b.cd[_coln(b.cd, "Chave 2") == norm(e.fabrica, e.cd)]
        hc["Frete Transferência"] = (float(transf.iloc[0]["Frete Transferência R$/t"]) if not transf.empty
                                     else _primeira(b.cd, "Fábrica", e.fabrica, "Frete Transferência R$/t"))
        # histórico do CD abastecido por ESTA fábrica; se não houver, o do CD com qualquer fábrica,
        # mantendo os custos que dependem da fábrica (transferência, CV, CF, Capex)
        nc = novo(exp=e.cd, fab=e.fabrica, cidade=e.cidade, canal=e.canal, incoterm=e.incoterm)
        excl = ()
        if not nc:
            nc = novo(exp=e.cd, cidade=e.cidade, canal=e.canal, incoterm=e.incoterm)
            excl = ("Frete Transferência", "Custo Variável", "Custo Fixo", "Capex")
        if nc:
            _sobrepor(hc, {k: v for k, v in nc.items() if k not in excl})
            frete_c = 0.0 if fob else hc["Frete de Vendas"]
            hc["Frete de Vendas"] = frete_c
            fonte["CD"] = rotulo(nc)
        aliq_c = hc["Impostos"] / hc["Preço Bruto"] if hc["Preço Bruto"] and hc["Impostos"] is not None else aliq
        sc = dict(hc)
        sc["Preço Bruto"] = (None if e.preco_negociado is None
                             else por_t(e.preco_negociado + ref["taxa_filial"]) + chapa_t)
        sc["Impostos"] = None if sc["Preço Bruto"] is None or aliq_c is None else sc["Preço Bruto"] * aliq_c
        # sem histórico no CD → usa o da fábrica (mesma lógica da planilha)
        for n in ("Descontos", "Devoluções", "Rapel", "Incentivo", "Custo Variável", "Custo Fixo",
                  "Força de Vendas", "Marketing", "Logísticos: Paletes e Strech", "Capex"):
            sc[n] = _ou(hc[n], sf[n])
            if hc[n] is None and sf[n] is not None:
                origens["CD"][n] = "igual à Fábrica (CD sem histórico)"
        for n in ("Take or Pay", "Outros Custos Log."):
            sc[n] = hc[n] or 0.0
        sc["Frete Devolução"] = 0.0
        sc["Frete de Vendas"] = 0.0 if fob else (None if frete_c is None else frete_c * (1 + adic))
        sc["Chapa"] = chapa_t
        rem = _v(b.contrato_cd, norm(e.cd), "Remuneração Total (R$/t)")
        sc["Movimentação de CDs"] = _ou(rem, hc["Movimentação de CDs"], 0.0)
        if rem is not None:
            origens["CD"]["Movimentação de CDs"] = "contrato do CD (remuneração)"

    # ---------------- DT / AP (somente histórico do distribuidor)
    kd = norm(e.fabrica, ref["ap"], e.produto)
    hd = _linhas_hist(b.dt, kd)
    nd = novo(exp=ref["expedicao"], cidade=e.cidade, canal="Distribuidor")
    if nd:
        _sobrepor(hd, nd)
        fonte["DT"] = rotulo(nd)
    tem_dt = ref["ap"] not in ("-", "", "nan") or nd is not None
    sd = dict(hd)
    sd["Capex"] = _ou(hd["Capex"], sf["Capex"])

    # ---------------- BROKER
    kb = norm(ref["broker"], e.cidade, cvel, e.produto)
    hb = _linhas_hist(b.broker, kb)
    nb = novo(broker=True, cidade=e.cidade, canal=e.canal)
    if nb:
        _sobrepor(hb, nb)
        fonte["Broker"] = rotulo(nb)
    tem_broker = ref["broker"] not in ("-", "", "nan") or nb is not None
    base_b = sc or sf
    sb = {n: _ou(hb[n], base_b[n]) for n in hb}
    for n in hb:
        if hb[n] is None and base_b[n] is not None:
            origens["Broker"][n] = f"igual ao {'CD' if sc else 'Fábrica'} (Broker sem histórico)"
    if e.preco_broker is not None:
        sb["Preço Bruto"] = por_t(e.preco_broker) + chapa_t
    aliq_b = hb["Impostos"] / hb["Preço Bruto"] if hb["Preço Bruto"] and hb["Impostos"] is not None else aliq
    sb["Impostos"] = None if sb["Preço Bruto"] is None or aliq_b is None else sb["Preço Bruto"] * aliq_b
    origens["Broker"]["Comissões"] = "contrato do broker"
    sb["Comissões"] = _ou(_v(b.contrato_broker, norm(ref["broker"]), "R$/t Comissão"), 0.0)
    sb["Chapa"] = chapa_t if e.chapa else (hb["Chapa"] or 0.0)
    if not tem_broker:
        sb = {n: None for n in hb}

    # ---------------- modo Excel: replica as fórmulas da aba SIMULADOR, inclusive as que apontam para colunas
    # inexistentes (PROCV fora do intervalo devolve 0 ou n/e)
    if e.modo_excel:
        nota = "igual à planilha Excel"
        fonte = {c: "réplica da planilha Excel jun/23" for c in CANAIS}
        sf["Descontos"] = 0.0                       # H40: PROCV col. 41 de C:AM -> erro -> 0
        sf["Outros Custos Log."] = 0.0              # H50: PROCV col. 39 de C:MX -> célula vazia -> 0
        origens["Fábrica"].update({"Descontos": nota, "Outros Custos Log.": nota})
        if sc:
            sc["Descontos"] = _ou(_v(b.cd, kc, "Descontos R$/t"), sf["Descontos"])   # M40 usa só 'Descontos'
            sc["Outros Custos Log."] = 0.0          # M50: coluna vazia -> 0
            sc["Logísticos: Paletes e Strech"] = sf["Logísticos: Paletes e Strech"]   # M59: erro -> usa fábrica
            sc["Movimentação de CDs"] = _ou(_v(b.contrato_cd, norm(e.cd), "Remuneração Total (R$/t)"), 0.0)  # M60
            # M62: com histórico do CD, PROCV col. 38 de C:MW -> célula vazia -> 0; sem histórico -> Capex da fábrica
            sc["Capex"] = 0.0 if kc in b.cd.index else sf["Capex"]
            for n in ("Descontos", "Outros Custos Log.", "Logísticos: Paletes e Strech", "Movimentação de CDs", "Capex"):
                origens["CD"][n] = nota
        sd["Outros Custos Log."] = 0.0
        origens["DT"]["Outros Custos Log."] = nota
        # na planilha, linha de custo sem histórico ("n/e") entra como 0 na simulação
        for canal_s, s_ in (("Fábrica", sf), ("CD", sc)):
            if not s_:
                continue
            for n in EDITAVEIS:
                if n not in ("Preço Bruto", "Impostos", "Incentivo") and s_.get(n) is None:
                    s_[n] = 0.0
                    origens[canal_s][n] = "n/e na planilha → 0"

    # ---------------- completar lacunas (o que no Excel ficava "n/e")
    sem_frete = (lambda origem_frete: None if fob or ref["frete_comercial"][origem_frete] is None
                 else por_t(ref["frete_comercial"][origem_frete]))
    if e.completar:
        _completar(sf, origens["Fábrica"], [
            ("média base resultados: fábrica/canal/produto",
             novo(exp=ref["expedicao"], canal=e.canal, incoterm=e.incoterm)),
            ("média base resultados: fábrica/produto", novo(exp=ref["expedicao"])),
            ("média fábrica/canal/produto", _media(b.fabrica, **{"Fábrica": e.fabrica, "Canal": cvel,
                                                                  "Produto": e.produto, "Modelo de Frete": e.incoterm})),
            ("média fábrica/produto", _media(b.fabrica, **{"Fábrica": e.fabrica, "Produto": e.produto}))])
        fr = None if fob else novo(exp=ref["expedicao"], cidade=e.cidade)
        if sf["Frete de Vendas"] is None and fr and fr["Frete de Vendas"]:
            sf["Frete de Vendas"] = fr["Frete de Vendas"] * (1 + adic)
            origens["Fábrica"]["Frete de Vendas"] = "base resultados: frete médio p/ cidade"
        if sf["Frete de Vendas"] is None and sem_frete("Fábrica") is not None:
            sf["Frete de Vendas"] = sem_frete("Fábrica") * (1 + adic)
            origens["Fábrica"]["Frete de Vendas"] = "frete comercial cadastrado"
        if sc:
            _completar(sc, origens["CD"], [
                ("média base resultados: CD/canal/produto",
                 novo(exp=e.cd, fab=e.fabrica, canal=e.canal, incoterm=e.incoterm)),
                ("média base resultados: CD/produto", novo(exp=e.cd, fab=e.fabrica)),
                ("média CD/canal/produto", _media(b.cd, **{"Expedição": e.cd, "Canal": cvel, "Produto": e.produto})),
                ("média CD/produto", _media(b.cd, **{"Expedição": e.cd, "Produto": e.produto}))])
            fr = None if fob else novo(exp=e.cd, cidade=e.cidade)
            if sc["Frete de Vendas"] is None and fr and fr["Frete de Vendas"]:
                sc["Frete de Vendas"] = fr["Frete de Vendas"] * (1 + adic)
                origens["CD"]["Frete de Vendas"] = "base resultados: frete médio p/ cidade"
            if sc["Frete de Vendas"] is None and sem_frete("CD") is not None:
                sc["Frete de Vendas"] = sem_frete("CD") * (1 + adic)
                origens["CD"]["Frete de Vendas"] = "frete comercial cadastrado"
        if tem_dt:
            _completar(sd, origens["DT"], [("média base resultados: distribuidores/produto",
                                            novo(exp=ref["expedicao"], canal="Distribuidor")),
                                           ("média DTs fábrica/produto",
                                            _media(b.dt, **{"Fábrica": e.fabrica, "Produto": e.produto}))], ())
        if tem_broker:
            _completar(sb, origens["Broker"], [("média base resultados: brokers/canal/produto",
                                                novo(broker=True, canal=e.canal)),
                                               ("média carteira/produto",
                                                _media(b.broker, **{"Carteira": ref["broker"], "Produto": e.produto}))],
                       ("Impostos",))
            if sb["Impostos"] is None and sb["Preço Bruto"] is not None and aliq is not None:
                sb["Impostos"] = sb["Preço Bruto"] * aliq

    motivo_canal = {
        "CD": None if e.cd else "selecione um CD",
        "DT": None if tem_dt else "cidade sem AP/DT cadastrado",
        "Broker": None if tem_broker else "cidade sem carteira Broker",
        "Fábrica": None,
    }

    resultado = {}
    for canal, h, s in (("Fábrica", hf, sf), ("CD", hc, sc), ("DT", hd, sd), ("Broker", hb, sb)):
        s = dict(s) if s else {n: None for n in hf}
        orig = origens[canal]
        for n, v in e.ajustes.get(canal, {}).items():   # ajustes manuais do usuário
            s[n] = v
            orig[n] = "ajuste manual"
        H, S = cascata(h), cascata(s)

        def origem(n, tipo):
            if tipo == "total":
                return "" if S.get(n) is not None else "n/e – falta linha acima"
            if n in orig:
                return orig[n]
            if S.get(n) is not None:
                return fonte[canal] if h.get(n) is not None and abs((h.get(n) or 0) - S.get(n)) < 1e-9 else "regra do simulador"
            if motivo_canal[canal]:
                return "n/e – " + motivo_canal[canal]
            if n == "Preço Bruto" and canal in ("Fábrica", "CD") and e.preco_negociado is None:
                return "n/e – informe o preço negociado"
            if n == "Frete de Vendas":
                return "n/e – zona sem frete cadastrado"
            return "n/e – sem histórico: preencha"

        df = pd.DataFrame({
            "Linha": [n for n, _ in LINHAS],
            "Tipo": [t for _, t in LINHAS],
            "Histórico R$/t": pd.to_numeric(pd.Series([H.get(n) for n, _ in LINHAS], dtype="object"), errors="coerce"),
            "Simulação R$/t": pd.to_numeric(pd.Series([S.get(n) for n, _ in LINHAS], dtype="object"), errors="coerce"),
            "Origem": [origem(n, t) for n, t in LINHAS],
        })
        df["Simulação R$/saco"] = (df["Simulação R$/t"] * peso / 1000) if ref["ensacado"] else float("nan")
        df = df[["Linha", "Tipo", "Histórico R$/t", "Simulação R$/t", "Simulação R$/saco", "Origem"]]
        resultado[canal] = df
    return {"ref": ref, "canais": resultado, "volume_t": e.sacos * peso / 1000, "fonte": fonte}


def resumo(res: dict) -> pd.DataFrame:
    """Quadro 'COMPARATIVO DE MARGENS' (FOB Líquido, MCU, MGS, MGCX)."""
    linhas = {"FOB Líquido": "Preço FOB", "MCU": "Margem de Contribuição",
              "MGS": "Margem de Servir", "MGCX": "Margem Caixa"}
    out = {}
    for canal, df in res["canais"].items():
        d = df.set_index("Linha")["Simulação R$/t"]
        out[canal] = {k: d[v] for k, v in linhas.items()}
    return pd.DataFrame(out)


def melhor_margem(res: dict) -> Optional[str]:
    mg = resumo(res).loc["MGCX"].dropna()
    return None if mg.empty else mg.astype(float).idxmax()


def codigo_equivalente(b: Bases, fabrica: str, nome: str) -> Optional[int]:
    """Código do produto nas bases jun/23 equivalente ao nome da base de resultados
    (mesma família e peso), usado só para buscar o que não existe na base nova (taxa filial, preço lista...)."""
    ip = info_produto(nome)
    cand = b.listas["produtos_por_fabrica"].get(fabrica, []) + [int(c) for c in b.produtos["Código do Produto"].dropna()]
    for c in cand:
        k = norm(c)
        if k in b.produtos.index:
            p = b.produtos.loc[k]
            if p["Família"] == ip["familia"] and float(p["Peso"]) == ip["peso"]:
                return int(c)
    return None


# ---------------------------------------------------------------- comparação de rotas (fábrica × CD)
def _produto_na_fabrica(b: Bases, r, fabrica: str, e: Entrada):
    """Produto equivalente (mesma família/peso/embalagem) vendido pela fábrica. Retorna (nome, código)."""
    alvo = info_produto(e.produto_nome) if e.produto_nome else None
    if r is not None and alvo:
        nomes = r.produtos(fabrica)
        if e.produto_nome in nomes:
            return e.produto_nome, codigo_equivalente(b, fabrica, e.produto_nome)
        for n in nomes:
            ip = info_produto(n)
            if ip == alvo:
                return n, codigo_equivalente(b, fabrica, n)
        return None, None
    # sem base de resultados: código da mesma família/peso nas bases jun/23
    k = norm(e.produto)
    if k not in b.produtos.index:
        return None, None
    fam, peso = b.produtos.at[k, "Família"], float(b.produtos.at[k, "Peso"])
    for c in b.listas["produtos_por_fabrica"].get(fabrica, []):
        kc = norm(c)
        if kc in b.produtos.index and b.produtos.at[kc, "Família"] == fam and float(b.produtos.at[kc, "Peso"]) == peso:
            return "", int(c)
    return None, None


def cds_da_fabrica(b: Bases, r, fabrica: str) -> list:
    cds = set(b.listas["cds_por_fabrica"].get(fabrica, []))
    if r is not None:
        cds |= set(r.expedicoes_cd(fabrica))
    vistos, out = set(), []
    for c in sorted(cds, key=str.upper):
        if norm(c) not in vistos:
            vistos.add(norm(c))
            out.append(c)
    return out


def rotas(b: Bases, e: Entrada, r=None) -> pd.DataFrame:
    """Simula a mesma venda saindo de cada fábrica: direto (Fábrica), por cada CD dela (CD) e pelo
    distribuidor (DT); ordena pela margem caixa."""
    from dataclasses import replace
    linhas = []
    for fab in b.listas["produtos_por_fabrica"].keys():
        nome, cod = _produto_na_fabrica(b, r, fab, e)
        if nome is None and cod is None:
            continue
        cds = cds_da_fabrica(b, r, fab)
        for i, cd in enumerate(cds or [""]):
            e2 = replace(e, fabrica=fab, produto_nome=nome or "", produto=cod, cd=cd, ajustes={},
                         taxa_filial=None, chapa_valor=None)
            res = simular(b, e2, r)
            alvos = (["Fábrica", "DT"] if i == 0 else []) + (["CD"] if cd else [])
            for formato in alvos:
                d = res["canais"][formato].set_index("Linha")
                faltam = d[(d["Tipo"] != "total") & d["Origem"].str.startswith("n/e")]
                est = d[(d["Tipo"] != "total") & d["Origem"].str.startswith(("média", "igual", "base resultados: frete",
                                                                               "frete"))]
                linhas.append({
                    "Fábrica": fab, "Formato": formato,
                    "Expedição": res["ref"]["expedicao"] if formato in ("Fábrica", "DT") else cd,
                    "Rota": f"{fab.title()} → " + {"Fábrica": "direto", "DT": "distribuidor"}.get(formato, cd),
                    "Produto": nome or str(cod),
                    "MGCX": d.at["Margem Caixa", "Simulação R$/t"],
                    "MGS": d.at["Margem de Servir", "Simulação R$/t"],
                    "MCU": d.at["Margem de Contribuição", "Simulação R$/t"],
                    "FOB Líquido": d.at["Preço FOB", "Simulação R$/t"],
                    "Frete vendas": d.at["Frete de Vendas", "Simulação R$/t"],
                    "Frete transf.": d.at["Frete Transferência", "Simulação R$/t"],
                    "Fonte": res["fonte"][formato],
                    "Estimadas": len(est),
                    "Pendência": "" if faltam.empty else ", ".join(faltam.index[:3]),
                })
    df = pd.DataFrame(linhas)
    if df.empty:
        return df
    df["_ok"] = df["MGCX"].notna()
    df = df.sort_values(["_ok", "MGCX"], ascending=[False, False]).drop(columns="_ok").reset_index(drop=True)
    return df


def preco_minimo(b: Bases, e: Entrada, r=None, alvo: float = 0.0) -> dict:
    """Preço negociado (R$/saco ou R$/t) a partir do qual a Margem Caixa atinge `alvo` R$/t, por formato.
    A margem é linear no preço, então bastam duas simulações."""
    from dataclasses import replace
    if e.preco_negociado is None:
        return {}
    r1 = resumo(simular(b, e, r)).loc["MGCX"]
    r2 = resumo(simular(b, replace(e, preco_negociado=e.preco_negociado + 1.0), r)).loc["MGCX"]
    out = {}
    for c in ("Fábrica", "CD"):
        m1, m2 = r1.get(c), r2.get(c)
        if m1 is None or m2 is None or pd.isna(m1) or pd.isna(m2) or m2 == m1:
            out[c] = None
        else:
            out[c] = e.preco_negociado + (alvo - m1) / (m2 - m1)
    return out
