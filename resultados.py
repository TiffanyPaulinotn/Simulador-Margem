"""
Base de resultados mensal (export do BI, .xlsx com colunas 'Ano/Mês', 'Expedição', 'Cidade', 'Canal',
'Produto', 'Incoterm', 'Volume (t)', 'Receita Bruta (R$)' ...).

Agrega os valores em R$ pela chave escolhida e divide pelo volume (R$/t ponderado por volume).
As linhas são montadas para fechar com os subtotais da própria base:
  Descontos          = Receita Bruta - Impostos - Devoluções - Receita Líquida (inclui provisões)
  Outros Custos Log. = (Receita Líquida - FOB s/Inc.) - fretes detalhados
  Custo Fixo         = MC s/Inc. - Margem Bruta
  Logísticos         = (Margem Bruta - Margem de Servir) - Comercial - Marketing - Despesas CDs
  Capex              = Capex (linhas com Capex/t fora de 0..500 R$/t são descartadas: valor inválido na base)
"""

from __future__ import annotations

VERSAO_APP = "2026-11-11.2"

import glob
import os
import pickle
import re
from typing import Optional

import pandas as pd

from bases import norm

COLS = {
    "vol": "Volume (t)", "rb": "Receita Bruta (R$)", "imp": "Impostos (R$)", "inc": "Incentivo Fiscal (R$)",
    "dev": "Devoluções (R$)", "rl": "Receita Líquida (R$)", "rapel": "Rapel (R$)",
    "fv": "Frete de Vendas (R$)", "fdev": "Frete de Devolução (R$)", "chapa": "Despesas de Chapa/Descarga (R$)",
    "top": "Take or Pay (R$)", "transf": "Frete Transf. CD's (R$)", "fob": "FOB Líquido S/Inc. (R$)",
    "cv": "Custo Variável (R$)", "mb": "Margem Bruta (R$)", "ms": "Margem de Servir (R$)",
    "com": "Comercial (R$)", "mkt": "Propaganda e Marketing (R$)", "cds": "Despesas com CD's (R$)",
    "capex": "Capex (R$)",
}
INCOTERM = {"CUSTO, SEGURO & FRETE": "CIF", "FRANCO A BORDO": "FOB"}
# canais da base nova -> canais das bases antigas (jun/23)
CANAL_VELHO = {"Construt. Predial": "Construtora", "Construt. Infra": "Construtora",
               "Distribuidor": "Mat.Con.", "Funcionário": "Consum.Final"}


def eh_base_resultados(caminho: str) -> bool:
    from arquivos import identificar
    return identificar(caminho)[0] == "resultados"


def achar_base(pasta: str) -> Optional[str]:
    from arquivos import achar
    return achar(pasta, "resultados")[0]


def info_produto(nome: str) -> dict:
    """Peso, embalagem e família a partir do nome (ex. 'CIMENTO CAUE ESTRUTURA CPV ARI 40KG')."""
    n = nome.upper()
    if "BIG BAG" in n:
        peso, emb = 1500.0, "Big Bag"
    elif "GRANEL" in n:
        peso, emb = 1000.0, "Granel"
    else:
        m = re.search(r"(\d+(?:[,.]\d+)?)\s*KG", n)
        peso, emb = (float(m.group(1).replace(",", ".")) if m else 50.0), "Ensacado"
    t = re.sub(r"[\s\-]", "", n)
    if "CPV" in t:
        fam = "CP V"
    elif "CPIV" in t:
        fam = "CP IV"
    elif "CPIII" in t:
        fam = "CP III"
    elif "CPII" in t:
        fam = "CP II 40" if re.search(r"CPII[EFZ]?40", t) else "CP II"
    else:
        fam = None
    return {"peso": peso, "embalagem": emb, "familia": fam}


class Resultados:
    def __init__(self, df: pd.DataFrame):
        self.df = df

    @property
    def meses(self):
        return sorted(self.df["mes"].unique())

    def produtos(self, fabrica: str):
        d = self.df[self.df["_fab"] == norm(fabrica)]
        return d.groupby("Produto")["vol"].sum().sort_values(ascending=False).index.tolist()

    def expedicoes_cd(self, fabrica: str):
        d = self.df[(self.df["_fab"] == norm(fabrica)) & self.df["Expedição"].str.upper().str.startswith("CD")]
        return sorted(d["Expedição"].unique())

    def volumes(self, por: str, meses=None, **filtros) -> pd.Series:
        """Volume (t) agrupado pela coluna `por` (ex. 'Cidade'), já filtrado; ordem decrescente."""
        import numpy as np
        self._prep()
        m = np.ones(len(self.df), dtype=bool)
        if meses:
            m &= np.isin(self._mes, list(meses))
        for k, v in filtros.items():
            if v is None or v == "":
                continue
            m &= self._chaves[k] == norm(v)
        d = self.df.loc[m, [por, "vol"]]
        s = d.groupby(por)["vol"].sum()
        return s[s > 0].sort_values(ascending=False)

    def agrupar(self, por, meses=None, min_vol: float = 0.0, **filtros) -> pd.DataFrame:
        """Resultado real (R$/t, ponderado por volume) agrupado por `por` (coluna ou lista), no destino filtrado.
        Colunas: Volume (t), Preço bruto, FOB líquido, Frete+log., MCU, MGS, Capex, MGCX — todas em R$/t."""
        import numpy as np
        self._prep()
        m = np.ones(len(self.df), dtype=bool)
        if meses:
            m &= np.isin(self._mes, list(meses))
        for k, v in filtros.items():
            if v is None or v == "" or (isinstance(v, (list, tuple, set)) and not v):
                continue
            if isinstance(v, (list, tuple, set)):
                m &= np.isin(self._chaves[k], [norm(x) for x in v])
            else:
                m &= self._chaves[k] == norm(v)
        por = [por] if isinstance(por, str) else list(por)
        d = self.df.loc[m, por + list(COLS) + ["capex_ok"]].copy()
        if d.empty:
            return pd.DataFrame()
        d["cap_v"] = d["capex"].where(d["capex_ok"], 0.0)
        d["vol_cap"] = d["vol"].where(d["capex_ok"], 0.0)
        s = d.groupby(por)[list(COLS) + ["cap_v", "vol_cap"]].sum()
        s = s[s["vol"] > max(min_vol, 0.0)]
        if s.empty:
            return pd.DataFrame()
        v = s["vol"]
        out = pd.DataFrame({
            "Volume (t)": v,
            "Preço bruto": s["rb"] / v,
            "FOB líquido": s["fob"] / v,
            "Frete+log.": (s["rl"] - s["fob"]) / v,
            "MCU": (s["fob"] - s["cv"]) / v,
            "MGS": s["ms"] / v,
            "Capex": (s["cap_v"] / s["vol_cap"].where(s["vol_cap"] > 0)).fillna(0.0),
        })
        out["MGCX"] = out["MGS"] - out["Capex"]
        out["Receita (R$)"] = s["rb"]
        out["MGCX total (R$)"] = out["MGCX"] * v
        return out.sort_values("MGCX", ascending=False).reset_index()

    def ufs(self):
        self._prep()
        return self.df.groupby("_uf")["vol"].sum().sort_values(ascending=False).index.tolist()

    def canais(self):
        return self.df.groupby("Canal")["vol"].sum().sort_values(ascending=False).index.tolist()

    def _prep(self):
        if getattr(self, "_chaves", None) is None:
            d = self.df
            self._chaves = {k: d["_" + k].to_numpy() for k in ("exp", "cidade", "canal", "produto", "incoterm", "fab")}
            self._mes = d["mes"].to_numpy()
            self._brk = d["_broker"].to_numpy()
            self._num = d[list(COLS) + ["capex_ok"]].copy()
            self._cache = {}
            # formato de atendimento de cada linha e UF do destino
            import numpy as np
            exp_cd = d["Expedição"].astype(str).str.upper().str.startswith("CD").to_numpy()
            dist = (d["_canal"] == norm("Distribuidor")).to_numpy()
            d["_formato"] = np.where(dist, "Distribuidor", np.where(exp_cd, "CD", "Fábrica"))
            d["_uf"] = d["Cidade"].astype(str).str[:2].str.upper()
            self._chaves["formato"] = d["_formato"].map(norm).to_numpy()
            self._chaves["uf"] = d["_uf"].to_numpy()

    def __getstate__(self):
        return {"df": self.df}

    def __setstate__(self, st_):
        self.df = st_["df"]

    def hist(self, meses=None, broker=False, min_vol=1.0, **filtros) -> Optional[dict]:
        """Linhas em R$/t para os filtros (exp, cidade, canal, produto, incoterm). None se sem volume."""
        self._prep()
        chave = (tuple(meses) if meses else None, broker, min_vol, tuple(sorted(filtros.items())))
        if chave not in self._cache:
            self._cache[chave] = self._hist(meses, broker, min_vol, filtros)
        return self._cache[chave]

    def _hist(self, meses, broker, min_vol, filtros):
        import numpy as np
        m = np.ones(len(self.df), dtype=bool)
        if meses:
            m &= np.isin(self._mes, list(meses))
        if broker:
            m &= self._brk
        for k, v in filtros.items():
            if v is None or v == "":
                continue
            m &= self._chaves[k] == norm(v)
        if not m.any():
            return None
        d = self._num[m]
        d = d.assign(mes=self._mes[m])
        s = d[list(COLS)].sum()
        vol = s["vol"]
        if vol < min_vol:
            return None
        cap = d[d["capex_ok"]]
        capex = cap["capex"].sum() / cap["vol"].sum() if cap["vol"].sum() > 0 else None
        t = lambda x: x / vol
        fretes = s["rl"] - s["fob"]
        fdet = s["fv"] + s["fdev"] + s["chapa"] + s["top"] + s["transf"]
        mc = s["fob"] - s["cv"]
        cs = s["mb"] - s["ms"]
        return {
            "Preço Bruto": t(s["rb"]), "Impostos": t(s["imp"]),
            "Descontos": t(s["rb"] - s["imp"] - s["dev"] - s["rl"]), "Devoluções": t(s["dev"]),
            "Rapel": t(s["rapel"]), "Incentivo": t(s["inc"]),
            "Frete de Vendas": t(s["fv"]), "Frete Devolução": t(s["fdev"]),
            "Frete Transferência": t(s["transf"]), "Chapa": t(s["chapa"]), "Take or Pay": t(s["top"]),
            "Outros Custos Log.": t(fretes - fdet), "Custo Variável": t(s["cv"]), "Custo Fixo": t(mc - s["mb"]),
            "Força de Vendas": t(s["com"]), "Comissões": 0.0, "Marketing": t(s["mkt"]),
            "Logísticos: Paletes e Strech": t(cs - s["com"] - s["mkt"] - s["cds"]),
            "Movimentação de CDs": t(s["cds"]), "Capex": capex,
            "_vol": vol, "_meses": d["mes"].nunique(),
        }


CHAVES = ["Ano/Mês", "Fábrica", "Expedição", "Cidade", "Canal", "Produto", "Incoterm", "Carteira"]


def _ler(caminho: str) -> Resultados:
    from arquivos import chave_col, engine, identificar, ler_csv
    _, aba = identificar(caminho)
    if caminho.lower().endswith(".csv"):
        d = ler_csv(caminho)
    else:
        d = pd.read_excel(caminho, sheet_name=aba or 0, engine=engine(caminho))
    # reconhece as colunas mesmo com acento/maiúsculas/espaços diferentes
    por_chave = {chave_col(c): c for c in d.columns}
    ren = {}
    for nome in CHAVES:
        if chave_col(nome) in por_chave:
            ren[por_chave[chave_col(nome)]] = nome
    for k, nome in COLS.items():
        if chave_col(nome) in por_chave:
            ren[por_chave[chave_col(nome)]] = k
    d = d.rename(columns=ren)
    faltam = [c for c in ("Ano/Mês", "Expedição", "Cidade", "Canal", "Produto") if c not in d.columns]
    if faltam or "vol" not in d.columns:
        raise ValueError("A base de resultados não tem as colunas: " + ", ".join(faltam + ([] if "vol" in d.columns
                                                                                         else ["Volume (t)"])))
    for c in ("Fábrica", "Incoterm", "Carteira"):
        if c not in d.columns:
            d[c] = ""
    if not d["Fábrica"].astype(str).str.strip().any():
        d["Fábrica"] = d["Expedição"].astype(str).str.replace("Cimentos ", "", regex=False).str.upper()
    for c in COLS:
        if c not in d.columns:
            d[c] = 0.0
        if not pd.api.types.is_numeric_dtype(d[c]) and d[c].astype(str).str.contains(",", regex=False).any():
            # números como texto no padrão brasileiro "1.234,56"
            d[c] = d[c].astype(str).str.replace(".", "", regex=False).str.replace(",", ".", regex=False)
        d[c] = pd.to_numeric(d[c], errors="coerce").fillna(0.0)
    d = d[d["Expedição"].notna() & d["Produto"].notna()].copy()
    for c in ("Fábrica", "Expedição", "Cidade", "Canal", "Produto", "Carteira"):
        d[c] = d[c].astype(str).str.strip()
    d["mes"] = d["Ano/Mês"].astype(str).str.replace(r"\D", "", regex=True).str[:6]
    d["Incoterm"] = d["Incoterm"].astype(str).str.upper().map(INCOTERM).fillna(d["Incoterm"])
    cpt = d["capex"] / d["vol"].where(d["vol"] > 0)
    d["capex_ok"] = cpt.between(0, 500)
    d["_fab"] = d["Fábrica"].map(norm)
    d["_exp"] = d["Expedição"].map(norm)
    d["_cidade"] = d["Cidade"].map(norm)
    d["_canal"] = d["Canal"].map(norm)
    d["_produto"] = d["Produto"].map(norm)
    d["_incoterm"] = d["Incoterm"].map(norm)
    d["_broker"] = d["Carteira"].astype(str).str.upper().str.startswith("BROKER")
    return Resultados(d)


def carregar_resultados(caminho: str) -> Resultados:
    cache = os.path.splitext(caminho)[0] + ".cache.pkl"
    mtime = os.path.getmtime(caminho)
    if os.path.exists(cache):
        try:
            with open(cache, "rb") as f:
                m, r = pickle.load(f)
            if m == mtime:
                return r
        except Exception:
            pass
    r = _ler(caminho)
    with open(cache, "wb") as f:
        pickle.dump((mtime, r), f)
    return r
