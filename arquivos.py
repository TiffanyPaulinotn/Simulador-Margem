"""
Identificação automática dos arquivos enviados (qualquer formato: .xlsb, .xlsx, .xlsm, .csv).

- "bases":      planilha do simulador (abas 'Base de Custos Fábrica', 'DE-PARA', 'Frete Logístico'...)
- "resultados": export do BI com colunas 'Ano/Mês', 'Expedição', 'Cidade', 'Canal', 'Produto', 'Volume (t)'...
"""

from __future__ import annotations

VERSAO_APP = "2026-11-11.2"

import glob
import os
import re
import unicodedata
from typing import Optional

EXTENSOES = (".xlsb", ".xlsx", ".xlsm", ".csv")
ABAS_BASES = {"BASE DE CUSTOS FABRICA", "DE-PARA", "FRETE LOGISTICO"}
COLS_RES = {"ANO/MES", "EXPEDICAO", "CIDADE", "CANAL", "PRODUTO"}
_cache: dict = {}
ERROS: dict = {}   # arquivo -> motivo de não ter conseguido abrir


def chave_col(nome) -> str:
    """Normaliza nome de coluna/aba: sem acento, maiúsculas, espaços simples."""
    s = unicodedata.normalize("NFKD", str(nome)).encode("ascii", "ignore").decode()
    return re.sub(r"\s+", " ", s).strip().upper()


def engine(caminho: str) -> Optional[str]:
    ext = os.path.splitext(caminho)[1].lower()
    return "pyxlsb" if ext == ".xlsb" else None


def ler_csv(caminho: str, nrows=None):
    import pandas as pd
    for enc in ("utf-8-sig", "latin-1"):
        try:
            return pd.read_csv(caminho, sep=None, engine="python", encoding=enc, nrows=nrows)
        except UnicodeDecodeError:
            continue
    raise ValueError("Não consegui ler o CSV (codificação desconhecida).")


def _abas_e_cabecalhos(caminho: str) -> dict:
    """{nome_aba: [cabeçalho da 1ª linha não vazia]} lendo o mínimo possível do arquivo."""
    ext = os.path.splitext(caminho)[1].lower()
    out = {}
    if ext == ".csv":
        out["csv"] = list(ler_csv(caminho, nrows=0).columns)
    elif ext == ".xlsb":
        from pyxlsb import open_workbook
        with open_workbook(caminho) as wb:
            for aba in wb.sheets:
                cab = []
                with wb.get_sheet(aba) as sh:
                    for i, row in enumerate(sh.rows()):
                        vals = [c.v for c in row if c.v not in (None, "")]
                        if vals:
                            cab = vals
                            break
                        if i > 5:
                            break
                out[aba] = cab
    else:
        from openpyxl import load_workbook
        wb = load_workbook(caminho, read_only=True, data_only=True)
        for ws in wb.worksheets:
            cab = []
            for i, row in enumerate(ws.iter_rows(max_row=6, values_only=True)):
                vals = [v for v in row if v not in (None, "")]
                if vals:
                    cab = vals
                    break
            out[ws.title] = cab
        wb.close()
    return out


def identificar(caminho: str) -> tuple[Optional[str], Optional[str]]:
    """Retorna (tipo, aba): tipo = 'bases' | 'resultados' | None."""
    try:
        chave = (caminho, os.path.getmtime(caminho), os.path.getsize(caminho))
    except OSError:
        return None, None
    if chave in _cache:
        return _cache[chave]
    tipo, aba_res = None, None
    try:
        abas = _abas_e_cabecalhos(caminho)
        if ABAS_BASES <= {chave_col(a) for a in abas}:
            tipo = "bases"
        else:
            for aba, cab in abas.items():
                if COLS_RES <= {chave_col(c) for c in cab}:
                    tipo, aba_res = "resultados", aba
                    break
    except Exception as ex:
        tipo, aba_res = None, None
        ERROS[caminho] = f"{type(ex).__name__}: {ex}"
    _cache[chave] = (tipo, aba_res)
    return tipo, aba_res


def pastas_de_busca(pasta: str) -> list:
    """Pasta do app, a pasta acima dela e as subpastas diretas (cobre zip descompactado em subpasta)."""
    pastas = [pasta, os.path.dirname(pasta)]
    try:
        pastas += [os.path.join(pasta, d) for d in os.listdir(pasta)
                   if os.path.isdir(os.path.join(pasta, d)) and not d.startswith((".", "__"))]
    except OSError:
        pass
    return pastas


def arquivos_da_pasta(pasta: str) -> list:
    fs = []
    for p in pastas_de_busca(pasta):
        for ext in EXTENSOES:
            fs += glob.glob(os.path.join(p, "*" + ext))
    fs = [f for f in set(fs) if not os.path.basename(f).startswith(("~$", "simulacao_", ".", "bases_jun23_embutidas"))]
    return sorted(fs, key=os.path.getmtime, reverse=True)


def achar(pasta: str, tipo: str) -> tuple[Optional[str], Optional[str]]:
    """Arquivo mais recente da pasta do tipo pedido → (caminho, aba)."""
    for f in arquivos_da_pasta(pasta):
        t, aba = identificar(f)
        if t == tipo:
            return f, aba
    return None, None
