"""Gera o Excel da simulação, formatado (capa/resumo, DRE comparativa, ranking de rotas, premissas)."""

from __future__ import annotations

VERSAO_APP = "2026-11-11.2"

import io
from datetime import datetime

import pandas as pd
from openpyxl import Workbook
from openpyxl.formatting.rule import DataBarRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from motor import CANAIS

NAVY, AZUL, CINZA, CINZA2, BRANCO = "12305C", "2A78D6", "F3F5F8", "E3E7ED", "FFFFFF"
VERDE, VERMELHO, AMBAR, TEXTO, MUTED = "1A7F37", "C0392B", "B7791F", "1F2933", "6B7889"
COR = {"Fábrica": "2A78D6", "CD": "16956A", "DT": "D9662B", "Broker": "7A5CD6"}
NUM = '#,##0.00;[Red]-#,##0.00;"–"'
MOEDA = '"R$" #,##0.00;[Red]-"R$" #,##0.00;"–"'
MOEDA0 = '"R$" #,##0;[Red]-"R$" #,##0;"–"'
fino = Side(style="thin", color=CINZA2)
BORDA = Border(bottom=fino)


def _f(cor):
    return PatternFill("solid", start_color=cor, end_color=cor)


def _v(x):
    return None if x is None or (isinstance(x, float) and pd.isna(x)) else float(x)


def _titulo(ws, texto, sub, ncols):
    ws.sheet_view.showGridLines = False
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=ncols)
    c = ws.cell(1, 1, texto)
    c.font = Font(name="Calibri", size=16, bold=True, color=BRANCO)
    c.alignment = Alignment(vertical="center", indent=1)
    s = ws.cell(2, 1, sub)
    s.font = Font(name="Calibri", size=10, color="C9D6EA")
    s.alignment = Alignment(vertical="center", indent=1)
    for rr in (1, 2):
        for cc in range(1, ncols + 1):
            ws.cell(rr, cc).fill = _f(NAVY)
    ws.row_dimensions[1].height = 30
    ws.row_dimensions[2].height = 20


def _cab(ws, row, valores, cores=None, inicio=1):
    for i, v in enumerate(valores):
        c = ws.cell(row, inicio + i, v)
        c.font = Font(bold=True, color=BRANCO, size=10)
        c.fill = _f((cores or {}).get(v, NAVY) if cores else NAVY)
        c.alignment = Alignment(horizontal="center" if i else "left", vertical="center", indent=0 if i else 1,
                                wrap_text=True)
    ws.row_dimensions[row].height = 22


def _secao(ws, row, texto, ncols):
    c = ws.cell(row, 1, texto.upper())
    c.font = Font(bold=True, size=10, color=AZUL)
    for cc in range(1, ncols + 1):
        ws.cell(row, cc).border = Border(bottom=Side(style="medium", color=AZUL))


def gerar_excel(e, res: dict, melhor, rotas_df: pd.DataFrame | None = None, ajustes: dict | None = None) -> bytes:
    ref = res["ref"]
    canais = res["canais"]
    vol_t = res["volume_t"]
    wb = Workbook()

    # ============================================================ RESUMO
    ws = wb.active
    ws.title = "Resumo"
    ncol = 6
    _titulo(ws, "Simulação de Resultado de Vendas",
            f"{e.fabrica.title()} → {e.cidade}  ·  gerado em {datetime.now():%d/%m/%Y %H:%M}", ncol)
    _secao(ws, 4, "Dados da simulação", ncol)
    dados = [("Fábrica", e.fabrica), ("Cidade (UF_Cidade)", e.cidade), ("Canal", e.canal),
             ("Produto", str(ref["nome"])), ("Incoterm", e.incoterm), ("CD", e.cd or "—"),
             ("Quantidade", f"{e.sacos:,.0f} sacos · {vol_t:,.1f} t".replace(",", ".")),
             ("Preço negociado", e.preco_negociado), ("Preço lista microrregião", _v(ref["preco_lista"])),
             ("Chapa", "Sim" if e.chapa else "Não"), ("Fracionamento", e.fracionamento),
             ("Período do histórico", f"{len(e.meses)} meses" if e.meses else "todos")]
    for i, (k, v) in enumerate(dados):
        r = 5 + i
        ws.cell(r, 1, k).font = Font(color=MUTED, size=10)
        c = ws.cell(r, 2, v)
        c.font = Font(bold=True, color=TEXTO, size=10)
        if k.startswith("Preço") and isinstance(v, (int, float)):
            c.number_format = MOEDA
        c.alignment = Alignment(horizontal="left")
        for cc in (1, 2):
            ws.cell(r, cc).border = BORDA

    # recomendação
    rr = 5
    ws.merge_cells(start_row=rr, start_column=4, end_row=rr, end_column=6)
    ws.cell(rr, 4, "FORMATO RECOMENDADO · MAIOR MARGEM CAIXA").font = Font(bold=True, size=9, color=MUTED)
    resumo_mg = {c: _v(canais[c].set_index("Linha").at["Margem Caixa", "Simulação R$/t"]) for c in CANAIS}
    mg = resumo_mg.get(melhor) if melhor else None
    ws.merge_cells(start_row=rr + 1, start_column=4, end_row=rr + 2, end_column=6)
    c = ws.cell(rr + 1, 4, melhor or "Sem dados suficientes")
    c.font = Font(bold=True, size=20, color=VERDE if (mg or 0) >= 0 else VERMELHO)
    c.alignment = Alignment(vertical="center", indent=1)
    for i, (k, v, fmt) in enumerate([("Margem caixa (R$/t)", mg, NUM),
                                     ("Margem caixa por saco", None if mg is None or not ref["ensacado"]
                                      else mg * ref["peso"] / 1000, MOEDA),
                                     ("Margem caixa no volume", None if mg is None else mg * vol_t, MOEDA0)]):
        r = rr + 3 + i
        ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=5)
        ws.cell(r, 4, k).font = Font(color=MUTED, size=10)
        cv = ws.cell(r, 6, v)
        cv.number_format = fmt
        cv.font = Font(bold=True, size=11, color=TEXTO)
    for r in range(rr, rr + 6):
        for cc in range(4, 7):
            ws.cell(r, cc).fill = _f("EAF6EE" if (mg or 0) >= 0 else "FCECEA")

    # comparativo
    r0 = 5 + len(dados) + 2
    _secao(ws, r0, "Comparativo de margens por formato (R$/t, sem incentivo)", ncol)
    _cab(ws, r0 + 1, ["Indicador"] + CANAIS, COR)
    for i, c_ in enumerate(CANAIS):
        if c_ == melhor:
            ws.cell(r0 + 1, 2 + i, f"{c_} ★")
    inds = [("Preço bruto", "Preço Bruto"), ("Receita líquida", "Receita Líquida"), ("FOB líquido", "Preço FOB"),
            ("Margem de contribuição (MCU)", "Margem de Contribuição"), ("Margem de servir (MGS)", "Margem de Servir"),
            ("Margem caixa (MGCX)", "Margem Caixa")]
    for i, (rot, lin) in enumerate(inds):
        r = r0 + 2 + i
        ws.cell(r, 1, rot).font = Font(size=10, color=TEXTO, bold=lin == "Margem Caixa")
        for j, c_ in enumerate(CANAIS):
            cel = ws.cell(r, 2 + j, _v(canais[c_].set_index("Linha").at[lin, "Simulação R$/t"]))
            cel.number_format = NUM
            cel.font = Font(size=10, bold=lin == "Margem Caixa")
            if c_ == melhor:
                cel.fill = _f("EAF6EE")
        for cc in range(1, 6):
            ws.cell(r, cc).border = BORDA
        if lin == "Margem Caixa":
            for j, c_ in enumerate([None] + CANAIS):
                ws.cell(r, 1 + j).fill = _f("C9EBD3" if c_ == melhor and c_ else "DCE8F8")
    r_fonte = r0 + 2 + len(inds)
    ws.cell(r_fonte, 1, "Fonte do histórico").font = Font(size=9, color=MUTED, italic=True)
    for j, c_ in enumerate(CANAIS):
        cf = ws.cell(r_fonte, 2 + j, res["fonte"][c_])
        cf.font = Font(size=8, color=MUTED, italic=True)
        cf.alignment = Alignment(wrap_text=True, horizontal="center")
    ws.row_dimensions[r_fonte].height = 26
    nota = r_fonte + 2
    ws.merge_cells(start_row=nota, start_column=1, end_row=nota, end_column=ncol)
    ws.cell(nota, 1, "O apontamento de melhor margem não implica direcionamento da venda; analise receitas e custos "
                     "com o seu gestor e com Pricing.").font = Font(size=9, italic=True, color=MUTED)
    for col, w in zip("ABCDEF", (30, 20, 20, 20, 20, 20)):
        ws.column_dimensions[col].width = w
    ws.column_dimensions["B"].width = 44

    # ============================================================ DRE
    wd = wb.create_sheet("DRE comparativa")
    nc = 1 + len(CANAIS) * 2
    _titulo(wd, "DRE comparativa por formato", "R$/t · Simulação e histórico · células em laranja = valor estimado "
                                              "(média / frete médio / ajuste)", nc)
    # cabeçalho em dois níveis
    wd.cell(4, 1, "")
    for j, c_ in enumerate(CANAIS):
        col = 2 + j * 2
        wd.merge_cells(start_row=4, start_column=col, end_row=4, end_column=col + 1)
        h = wd.cell(4, col, c_ + (" ★" if c_ == melhor else ""))
        h.font = Font(bold=True, color=BRANCO)
        h.alignment = Alignment(horizontal="center")
        for cc in (col, col + 1):
            wd.cell(4, cc).fill = _f(COR[c_])
    _cab(wd, 5, ["Linha (R$/t)"] + ["Simulação", "Histórico"] * len(CANAIS))
    dfs = {c_: canais[c_].set_index("Linha") for c_ in CANAIS}
    for i, lin in enumerate(dfs["Fábrica"].index):
        r = 6 + i
        tipo = dfs["Fábrica"].at[lin, "Tipo"]
        nome = lin + (" (não soma)" if tipo == "info" else "")
        c = wd.cell(r, 1, nome if tipo == "total" else "   " + nome)
        c.font = Font(size=10, bold=tipo == "total", italic=tipo == "info", color=TEXTO if tipo != "info" else MUTED)
        for j, c_ in enumerate(CANAIS):
            sim = _v(dfs[c_].at[lin, "Simulação R$/t"])
            hist = _v(dfs[c_].at[lin, "Histórico R$/t"])
            org = str(dfs[c_].at[lin, "Origem"])
            cs = wd.cell(r, 2 + j * 2, sim)
            ch = wd.cell(r, 3 + j * 2, hist)
            for cel in (cs, ch):
                cel.number_format = NUM
            est = org.startswith(("média", "igual", "ajuste", "frete", "base resultados: frete"))
            cs.font = Font(size=10, bold=tipo == "total", color=AMBAR if est and tipo != "total" else TEXTO)
            ch.font = Font(size=9, color=MUTED)
            if sim is None and tipo != "total":
                cs.value = "n/e"
                cs.font = Font(size=9, color=VERMELHO)
                cs.alignment = Alignment(horizontal="right")
        for cc in range(1, nc + 1):
            cel = wd.cell(r, cc)
            cel.border = BORDA
            if tipo == "total":
                cel.fill = _f("DCE8F8" if lin == "Margem Caixa" else CINZA)
    wd.column_dimensions["A"].width = 34
    for cc in range(2, nc + 1):
        wd.column_dimensions[get_column_letter(cc)].width = 13
    wd.freeze_panes = "B6"

    # ============================================================ ROTAS
    if rotas_df is not None and not rotas_df.empty:
        wr = wb.create_sheet("Rotas (fábrica × CD)")
        cols = ["Rota", "Formato", "MGCX", "MGS", "MCU", "FOB Líquido", "Frete vendas", "Frete transf.",
                "Estimadas", "Fonte", "Pendência"]
        _titulo(wr, "Onde atender: ranking de rotas",
                f"Mesma venda ({e.cidade} · {e.canal} · preço {e.preco_negociado or '—'}) saindo de cada fábrica, "
                f"direto ou via CD · R$/t", len(cols) + 1)
        _cab(wr, 4, ["#"] + cols)
        for i, (_, row) in enumerate(rotas_df.iterrows()):
            r = 5 + i
            wr.cell(r, 1, i + 1).font = Font(size=9, color=MUTED)
            for j, k in enumerate(cols):
                v = row.get(k)
                v = _v(v) if isinstance(v, float) else v
                c = wr.cell(r, 2 + j, v)
                c.font = Font(size=10, color=TEXTO, bold=k == "MGCX")
                if k in ("MGCX", "MGS", "MCU", "FOB Líquido", "Frete vendas", "Frete transf."):
                    c.number_format = NUM
                if k == "Formato":
                    c.font = Font(size=10, bold=True, color=COR.get(v, TEXTO))
                if k in ("Fonte", "Pendência"):
                    c.font = Font(size=9, color=MUTED)
            for cc in range(1, len(cols) + 2):
                wr.cell(r, cc).border = BORDA
            if i == 0 and _v(row.get("MGCX")) is not None:
                for cc in range(1, len(cols) + 2):
                    wr.cell(r, cc).fill = _f("EAF6EE")
        fim = 4 + len(rotas_df)
        wr.conditional_formatting.add(f"D5:D{fim}", DataBarRule(start_type="min", end_type="max", color="4C8DFF"))
        widths = [5, 36, 11, 12, 11, 11, 13, 13, 13, 11, 30, 34]
        for i, w in enumerate(widths):
            wr.column_dimensions[get_column_letter(i + 1)].width = w
        wr.freeze_panes = "C5"
        wr.auto_filter.ref = f"A4:{get_column_letter(len(cols) + 1)}{fim}"

    # ============================================================ PREMISSAS
    wp = wb.create_sheet("Premissas e origem")
    _titulo(wp, "Premissas e origem de cada valor", "De onde veio cada linha da simulação (R$/t)", 5)
    _cab(wp, 4, ["Formato", "Linha", "Histórico", "Simulação", "Origem"])
    r = 5
    for c_ in CANAIS:
        d = canais[c_]
        for _, row in d[d["Tipo"] != "total"].iterrows():
            wp.cell(r, 1, c_).font = Font(size=10, bold=True, color=COR[c_])
            wp.cell(r, 2, row["Linha"]).font = Font(size=10, color=TEXTO)
            for j, k in enumerate(("Histórico R$/t", "Simulação R$/t")):
                cel = wp.cell(r, 3 + j, _v(row[k]))
                cel.number_format = NUM
                cel.font = Font(size=10)
            o = str(row["Origem"])
            wp.cell(r, 5, o).font = Font(size=9, color=VERMELHO if o.startswith("n/e") else
                                         AMBAR if o.startswith(("média", "igual", "ajuste", "frete")) else MUTED)
            for cc in range(1, 6):
                wp.cell(r, cc).border = BORDA
            r += 1
    for col, w in zip("ABCDE", (12, 32, 13, 13, 48)):
        wp.column_dimensions[col].width = w
    wp.freeze_panes = "C5"
    wp.auto_filter.ref = f"A4:E{r - 1}"
    if ajustes:
        r += 1
        wp.cell(r, 1, "Ajustes manuais").font = Font(bold=True, color=AZUL)
        for c_, aj in ajustes.items():
            for k, v in aj.items():
                r += 1
                wp.cell(r, 1, c_)
                wp.cell(r, 2, k)
                cel = wp.cell(r, 4, v)
                cel.number_format = NUM

    for w in wb.worksheets:
        w.page_setup.orientation = "landscape"
        w.page_setup.fitToWidth = 1
        w.sheet_properties.pageSetUpPr.fitToPage = True
        w.page_setup.fitToHeight = 0
        w.sheet_properties.tabColor = NAVY if w.title == "Resumo" else AZUL
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
