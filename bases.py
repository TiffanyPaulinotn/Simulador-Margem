"""
Leitura das bases da planilha Simulador_de_Margens_Canais_BR (.xlsb).

Cada aba vira um DataFrame limpo. Na primeira leitura é gerado um cache
(.pkl) ao lado da planilha para as próximas execuções serem instantâneas.
Para atualizar as bases (novo mês), basta trocar a planilha: o cache é
refeito automaticamente quando a data do arquivo muda.
"""

from __future__ import annotations

VERSAO_APP = "2026-11-11.2"

import os
import pickle
from dataclasses import dataclass, field

import pandas as pd

ABAS = [
    "Base de Custos Fábrica", "Base de Custos CD", "Base de Custos DT",
    "Base de Custos Broker", "Preço", "Impostos", "Frete Logístico",
    "Frete Comercial", "Taxa Filial", "Contratos", "Listas", "DE-PARA",
    "Prêmio DT+Preço Lista",
]


def norm(*partes) -> str:
    """Monta chave igual ao Excel (concatenação), sem diferenciar maiúsc./minúsc.
    (o PROCV do Excel também não diferencia)."""
    out = []
    for p in partes:
        if p is None or (isinstance(p, float) and pd.isna(p)):
            p = ""
        if isinstance(p, float) and p.is_integer():
            p = int(p)
        out.append(str(p).strip())
    return "".join(out).upper()


def _tabela(raw: pd.DataFrame, linha_cab: int, col_ini: int, col_fim: int) -> pd.DataFrame:
    """Recorta uma tabela de uma aba (cabeçalho na linha `linha_cab`)."""
    bloco = raw.iloc[linha_cab:, col_ini:col_fim + 1]
    cab, vistos = [], {}
    for i, c in enumerate(bloco.iloc[0]):
        nome = str(c).strip() if pd.notna(c) else f"col{i}"
        if nome in vistos:                       # cabeçalhos repetidos -> "Valor.1"
            vistos[nome] += 1
            nome = f"{nome}.{vistos[nome]}"
        else:
            vistos[nome] = 0
        cab.append(nome)
    df = bloco.iloc[1:].copy()
    df.columns = cab
    df = df.dropna(how="all")
    return df.reset_index(drop=True)


def _indexar(df: pd.DataFrame, chave: str) -> pd.DataFrame:
    df = df[df[chave].notna()].copy()
    df["_k"] = df[chave].map(norm)
    # PROCV pega a primeira ocorrência
    return df.drop_duplicates("_k").set_index("_k")


@dataclass
class Bases:
    fabrica: pd.DataFrame
    cd: pd.DataFrame
    dt: pd.DataFrame
    broker: pd.DataFrame
    preco: pd.DataFrame
    impostos: pd.DataFrame
    frete_log_fab: pd.DataFrame
    frete_log_cd: pd.DataFrame
    frete_com: pd.DataFrame
    taxa_filial: pd.DataFrame
    contrato_cd: pd.DataFrame
    contrato_broker: pd.DataFrame
    fabricas: pd.DataFrame       # Fábrica, UF, chapa, expedição
    produtos: pd.DataFrame       # código, peso, embalagem, nome
    cidades: pd.DataFrame        # cidade IBGE -> AP, broker, micro
    fracionamento: pd.DataFrame
    clientes: pd.DataFrame
    premio_dt: pd.DataFrame
    preco_lista: pd.DataFrame
    listas: dict = field(default_factory=dict)


def _ler(caminho: str) -> Bases:
    from arquivos import _abas_e_cabecalhos, chave_col, engine
    reais = {chave_col(a): a for a in _abas_e_cabecalhos(caminho)}
    faltam = [a for a in ABAS if chave_col(a) not in reais]
    if faltam:
        raise ValueError("A planilha de bases não tem as abas: " + ", ".join(faltam))
    lidas = pd.read_excel(caminho, sheet_name=[reais[chave_col(a)] for a in ABAS], header=None,
                          engine=engine(caminho))
    raw = {a: lidas[reais[chave_col(a)]] for a in ABAS}
    # o pyxlsb devolve uma linha vazia antes da linha 1 da planilha
    for aba, df in raw.items():
        while len(df) and df.iloc[0].isna().all():
            df = df.iloc[1:]
        raw[aba] = df.reset_index(drop=True)

    def base_custos(aba):
        r = raw[aba]
        return _tabela(r, 0, 2, r.shape[1] - 1)

    fab = base_custos("Base de Custos Fábrica")
    cd = base_custos("Base de Custos CD")
    dt = base_custos("Base de Custos DT")
    bro = base_custos("Base de Custos Broker")

    fl = raw["Frete Logístico"]
    frete_fab = _tabela(fl, 1, 2, 11)
    frete_cd = _tabela(fl, 1, 13, 22)
    # a chave da planilha não tem o incoterm e o PROCV pegava a 1ª linha (às vezes FOB = 0);
    # aqui a chave inclui o modelo de frete
    frete_fab["CHAVE Fábrica"] = frete_fab["CHAVE Fábrica"].astype(str) + frete_fab["Modelo de Frete"].astype(str)
    frete_cd["CHAVE Expedição"] = frete_cd["CHAVE Expedição"].astype(str) + frete_cd["Modelo de Frete"].astype(str)

    ct = raw["Contratos"]
    contrato_broker = _tabela(ct, 2, 1, 2)
    contrato_cd = _tabela(ct, 2, 4, 13)

    dp = raw["DE-PARA"]
    fabricas = _tabela(dp, 0, 0, 4).rename(columns={"col4": "Chapa R$/sc (50kg)"})
    produtos = _tabela(dp, 0, 5, 9)
    cidades = _tabela(dp, 0, 11, 17)
    fracion = _tabela(dp, 0, 23, 24)
    clientes = _tabela(dp, 0, 26, 31)

    pr = raw["Prêmio DT+Preço Lista"]
    premio = _tabela(pr, 0, 2, 3)
    lista = _tabela(pr, 0, 6, 10)

    # Listas suspensas (CDs e produtos por fábrica, canais, zonas...)
    li = raw["Listas"]
    cab = li.iloc[0]
    listas = {"canais": li.iloc[1:, 1].dropna().astype(str).tolist(),
              "zonas": li.iloc[1:, 16].dropna().astype(str).tolist(),
              "cds_por_fabrica": {}, "produtos_por_fabrica": {}}
    for c in range(18, 28):
        listas["cds_por_fabrica"][str(cab[c])] = li.iloc[1:, c].dropna().astype(str).tolist()
    for c in range(35, 48):
        listas["produtos_por_fabrica"][str(cab[c])] = [
            int(v) for v in li.iloc[1:, c].dropna() if isinstance(v, (int, float))]
    seg = {}
    for c in range(29, 34):
        seg[str(cab[c])] = li.iloc[1:, c].dropna().astype(str).tolist()
    listas["segmentos"] = seg

    return Bases(
        fabrica=_indexar(fab, "Chave"),
        cd=_indexar(cd, "Chave"),
        dt=_indexar(dt, "Chave"),
        broker=_indexar(bro, "Chave"),
        preco=_indexar(_tabela(raw["Preço"], 0, 2, 15), "Chave"),
        impostos=_indexar(_tabela(raw["Impostos"], 0, 2, 8), "CHAVE"),
        frete_log_fab=_indexar(frete_fab, "CHAVE Fábrica"),
        frete_log_cd=_indexar(frete_cd, "CHAVE Expedição"),
        frete_com=_indexar(_tabela(raw["Frete Comercial"], 0, 2, 14), "CHAVE"),
        taxa_filial=_indexar(_tabela(raw["Taxa Filial"], 0, 2, 5), "CHAVE"),
        contrato_cd=_indexar(contrato_cd, "CENTRO"),
        contrato_broker=_indexar(contrato_broker, "Carteira Broker"),
        fabricas=_indexar(fabricas, "Fábrica"),
        produtos=_indexar(produtos, "Código do Produto"),
        cidades=_indexar(cidades, "Cidade IBGE"),
        fracionamento=_indexar(fracion, "Fracionamento"),
        clientes=_indexar(clientes, "Cód. Cliente"),
        premio_dt=_indexar(premio, "Unidade Rede"),
        preco_lista=_indexar(lista, "Chave"),
        listas=listas,
    )


VERSAO = 3


def carregar(caminho: str) -> Bases:
    cache = os.path.splitext(caminho)[0] + ".cache.pkl"
    mtime = os.path.getmtime(caminho)
    if os.path.exists(cache):
        try:
            with open(cache, "rb") as f:
                v, m, b = pickle.load(f)
            if m == mtime and v == VERSAO:
                return b
        except Exception:
            pass
    b = _ler(caminho)
    with open(cache, "wb") as f:
        pickle.dump((VERSAO, mtime, b), f)
    return b


EMBUTIDAS = "bases_jun23_embutidas.xlsx"
