"""
Simulador de Margens por Formato de Atendimento — versão Python (Streamlit).

Rodar:   python app.py        (ou: streamlit run app.py)
"""
import html
import io
import os
import subprocess
import sys

import altair as alt
import pandas as pd
import streamlit as st
from streamlit import runtime

# Se rodar com "python app.py", reabre automaticamente com "streamlit run" (abre o navegador)
if __name__ == "__main__" and not runtime.exists():
    sys.exit(subprocess.call([sys.executable, "-m", "streamlit", "run", os.path.abspath(__file__)]))

# ---- garante que os módulos do simulador são da mesma versão do app (e recarrega se foram trocados)
import importlib  # noqa: E402

VERSAO_APP = "2026-11-11.2"
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
_desatualizados = []
for _nome in ("arquivos", "bases", "resultados", "motor", "exportar", "login"):
    try:
        _m = importlib.import_module(_nome)
        _mt = os.path.getmtime(_m.__file__)
        if getattr(_m, "_mtime_carregado", None) not in (None, _mt) or getattr(_m, "VERSAO_APP", None) != VERSAO_APP:
            _m = importlib.reload(_m)
        _m._mtime_carregado = _mt
        if getattr(_m, "VERSAO_APP", None) != VERSAO_APP:
            _desatualizados.append(f"{_nome}.py")
    except Exception as _ex:  # arquivo faltando ou com erro
        _desatualizados.append(f"{_nome}.py ({type(_ex).__name__})")
if _desatualizados:
    st.error("Alguns arquivos do simulador estão desatualizados ou faltando: **" + ", ".join(_desatualizados)
             + f"**. Substitua TODOS os arquivos .py da pasta {os.path.dirname(os.path.abspath(__file__))} "
               "pelos do zip mais recente e reabra o simulador.")
    st.stop()

from bases import EMBUTIDAS, carregar, norm  # noqa: E402
from motor import (CANAIS, EDITAVEIS, Entrada, cds_da_fabrica, codigo_equivalente, melhor_margem,  # noqa: E402
                   preco_minimo, resumo, rotas, simular)
from exportar import gerar_excel  # noqa: E402
from arquivos import ERROS, achar, arquivos_da_pasta, identificar  # noqa: E402
from resultados import achar_base, carregar_resultados, eh_base_resultados  # noqa: E402

st.set_page_config(page_title="Simulador de Margens · Canais BR", page_icon="📊", layout="wide",
                   initial_sidebar_state="expanded")
PASTA = os.path.dirname(os.path.abspath(__file__))
PADRAO = os.path.join(PASTA, "Simulador_de_Margens_Canais_BR_jun23.xlsb")

# ------------------------------------------------------------ tokens de cor (tema escuro)
BG, SUP, SUP2, BORDA = "#0b0f14", "#121821", "#18202b", "#243041"
TXT, TXT2, TXT3 = "#e8edf3", "#9aa7b6", "#6b7889"
AZUL, VERDE, VERMELHO, AMBAR = "#4c8dff", "#2fbf71", "#ff5d5d", "#f0b429"
DEDUCAO = "#e07a3f"
# cor de cada formato de atendimento (identidade fixa em cartões, tabelas e gráficos)
COR = {"Fábrica": "#4c8dff", "CD": "#22b07d", "DT": "#f08a3c", "Broker": "#a48bff"}
GRADE = "#1e2733"

st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
html, body, [class*="css"], .stMarkdown, .stText, button, input, select, textarea {{
  font-family: 'Inter', 'Segoe UI', system-ui, sans-serif !important; }}
/* esconder elementos padrão do Streamlit */
#MainMenu, footer, [data-testid="stDecoration"], [data-testid="stStatusWidget"],
[data-testid="stAppDeployButton"], [data-testid="stMainMenu"] {{display:none !important;}}
[data-testid="stHeader"] {{background: transparent;}}
[data-testid="stSidebarCollapsedControl"], [data-testid="stExpandSidebarButton"],
[data-testid="stSidebarCollapseButton"] {{display:flex !important; visibility:visible !important;}}
.block-container {{padding: 1.2rem 2.2rem 3rem; max-width: 1600px;}}
.stApp {{background: {BG};}}
/* barra lateral */
[data-testid="stSidebar"] {{background: {SUP}; border-right: 1px solid {BORDA};}}
[data-testid="stSidebar"] .block-container, [data-testid="stSidebarContent"] {{padding-top: .6rem;}}
.sb-brand {{display:flex; align-items:center; gap:10px; padding: 4px 0 14px; border-bottom:1px solid {BORDA};
  margin-bottom: 6px;}}
.sb-logo {{width:34px; height:34px; border-radius:9px; background: linear-gradient(135deg,{AZUL},#7a5cff);
  display:flex; align-items:center; justify-content:center; font-weight:800; color:#fff; font-size:.95rem;}}
.sb-brand b {{display:block; color:{TXT}; font-size:.95rem; line-height:1.1;}}
.sb-brand span {{color:{TXT3}; font-size:.72rem; letter-spacing:.04em;}}
.sb-sec {{font-size:.68rem; font-weight:700; letter-spacing:.12em; text-transform:uppercase; color:{TXT3};
  margin: 18px 0 4px; display:flex; align-items:center; gap:8px;}}
.sb-sec:after {{content:""; flex:1; height:1px; background:{BORDA};}}
[data-testid="stSidebar"] label p {{font-size:.8rem !important; color:{TXT2} !important;}}
.sb-ok {{font-size:.74rem; color:{TXT2}; background:{SUP2}; border:1px solid {BORDA}; border-radius:8px;
  padding:8px 10px; margin-top:4px;}}
.sb-ok b {{color:{VERDE};}}
/* topo */
.top {{display:flex; justify-content:space-between; align-items:flex-end; padding: 6px 0 18px;
  border-bottom:1px solid {BORDA}; margin-bottom: 20px;}}
.top .eyebrow {{font-size:.7rem; font-weight:700; letter-spacing:.14em; text-transform:uppercase; color:{AZUL};}}
.top h1 {{font-size:1.6rem; font-weight:800; color:{TXT}; margin:2px 0 6px; letter-spacing:-.01em; padding:0;}}
.chips {{display:flex; flex-wrap:wrap; gap:6px;}}
.chip {{background:{SUP2}; border:1px solid {BORDA}; color:{TXT2}; border-radius:999px; padding:3px 11px;
  font-size:.76rem; white-space:nowrap;}}
.chip b {{color:{TXT}; font-weight:600;}}
.top .right {{text-align:right;}}
.top .right .k {{font-size:.7rem; color:{TXT3}; text-transform:uppercase; letter-spacing:.08em;}}
.top .right .v {{font-size:1.5rem; font-weight:800; color:{TXT};}}
.top .right .v small {{font-size:.8rem; color:{TXT2}; font-weight:500;}}
/* destaque */
.hero {{display:grid; grid-template-columns: 1.25fr 2fr; gap:16px; margin-bottom: 18px;}}
.panel {{background:{SUP}; border:1px solid {BORDA}; border-radius:14px; padding:18px 20px;}}
.rec {{background: radial-gradient(120% 140% at 0% 0%, rgba(47,191,113,.16) 0%, {SUP} 55%);
  border-color: rgba(47,191,113,.35);}}
.rec.neg {{background: radial-gradient(120% 140% at 0% 0%, rgba(255,93,93,.14) 0%, {SUP} 55%);
  border-color: rgba(255,93,93,.35);}}
.rec .k {{font-size:.7rem; font-weight:700; letter-spacing:.12em; text-transform:uppercase; color:{TXT3};}}
.rec .ch {{font-size:1.45rem; font-weight:800; color:{TXT}; margin-top:6px;}}
.rec .val {{font-size:2.6rem; font-weight:800; letter-spacing:-.02em; color:{VERDE}; line-height:1.1; margin-top:4px;}}
.rec.neg .val {{color:{VERMELHO};}}
.rec .val small {{font-size:.95rem; color:{TXT2}; font-weight:500;}}
.rec .foot {{display:flex; gap:22px; margin-top:12px; font-size:.8rem; color:{TXT2};}}
.rec .foot b {{color:{TXT}; font-size:.95rem; display:block;}}
.kpis {{display:grid; grid-template-columns: repeat(3, 1fr); gap:0;}}
.kpi {{padding: 6px 16px; border-left:1px solid {BORDA};}}
.kpi:nth-child(3n+1) {{border-left:none; padding-left:0;}}
.kpi:nth-child(n+4) {{margin-top:14px;}}
.kpi .k {{font-size:.7rem; color:{TXT3}; text-transform:uppercase; letter-spacing:.08em;}}
.kpi .v {{font-size:1.15rem; font-weight:700; color:{TXT}; margin-top:3px; font-variant-numeric: tabular-nums;}}
.kpi .s {{font-size:.74rem; color:{TXT2}; margin-top:1px;}}
.up {{color:{VERDE} !important;}} .dn {{color:{VERMELHO} !important;}}
/* cartões dos formatos */
.sec {{display:flex; align-items:baseline; gap:10px; margin: 8px 0 10px;}}
.sec h3 {{font-size:1rem; font-weight:700; color:{TXT}; margin:0; padding:0;}}
.sec span {{font-size:.78rem; color:{TXT3};}}
.cards {{display:grid; grid-template-columns: repeat(4, 1fr); gap:14px; margin-bottom: 8px;}}
.card {{background:{SUP}; border:1px solid {BORDA}; border-radius:14px; padding:16px 18px 14px; position:relative;
  border-top: 3px solid var(--c); background: linear-gradient(180deg, color-mix(in srgb, var(--c) 9%, {SUP}) 0%, {SUP} 45%);}}
.card .dot {{display:inline-block; width:9px; height:9px; border-radius:50%; background:var(--c); margin-right:8px;
  box-shadow: 0 0 10px var(--c);}}
.card:hover {{transform: translateY(-2px); transition: .15s; box-shadow: 0 12px 30px rgba(0,0,0,.35);}}
.route {{background: linear-gradient(110deg, rgba(76,141,255,.18) 0%, rgba(34,176,125,.14) 50%, {SUP} 100%);
  border:1px solid {BORDA}; border-radius:14px; padding:18px 22px; display:grid; grid-template-columns: 1.4fr 1fr 1fr 1fr;
  gap:18px; align-items:center; margin-bottom:14px;}}
.route .k {{font-size:.68rem; font-weight:700; letter-spacing:.12em; text-transform:uppercase; color:{TXT3};}}
.route .v {{font-size:1.25rem; font-weight:800; color:{TXT}; margin-top:4px;}}
.route .v.big {{font-size:1.6rem;}}
.route .s {{font-size:.76rem; color:{TXT2}; margin-top:2px;}}
.tag {{display:inline-block; font-size:.66rem; font-weight:700; padding:2px 8px; border-radius:999px; margin-left:6px;
  vertical-align: middle;}}
.card.best {{border-color: rgba(47,191,113,.55); box-shadow: 0 0 0 1px rgba(47,191,113,.25), 0 10px 30px rgba(0,0,0,.25);}}
.card .hd {{display:flex; justify-content:space-between; align-items:center;}}
.card .t {{font-weight:700; font-size:.98rem; color:{TXT};}}
.pill {{font-size:.64rem; font-weight:700; letter-spacing:.06em; padding:3px 8px; border-radius:999px;}}
.pill.best {{background: rgba(47,191,113,.15); color:{VERDE};}}
.pill.neg {{background: rgba(255,93,93,.13); color:{VERMELHO};}}
.pill.na {{background: {SUP2}; color:{TXT3};}}
.card .sub {{font-size:.76rem; color:{TXT2}; margin-top:2px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;}}
.card .src {{font-size:.7rem; color:{TXT3}; margin-top:2px;}}
.card .big {{font-size:1.9rem; font-weight:800; color:{TXT}; margin-top:10px; letter-spacing:-.01em;
  font-variant-numeric: tabular-nums;}}
.card .big small {{font-size:.72rem; font-weight:600; color:{TXT3}; letter-spacing:.06em; margin-left:4px;}}
.card .big.neg {{color:{VERMELHO};}} .card .big.na {{color:{TXT3};}}
.card .saco {{font-size:.8rem; color:{TXT2};}}
.bar {{height:5px; background:{SUP2}; border-radius:99px; margin:10px 0 8px; overflow:hidden;}}
.bar i {{display:block; height:100%; border-radius:99px;}}
.card table {{width:100%; font-size:.8rem; border-collapse:collapse; border:none; margin:0;}}
.card tr, .card td {{border:none !important; background: transparent !important;}}
.card td {{padding:4px 0; color:{TXT2};}}
.card td:last-child {{text-align:right; color:{TXT}; font-variant-numeric: tabular-nums; font-weight:500;}}
.card .warn {{margin-top:8px; font-size:.72rem; color:{AMBAR};}}
.na-txt {{color:{TXT3} !important;}}
.note {{font-size:.74rem; color:{TXT3}; margin: 6px 0 18px;}}
/* DRE */
.dre {{width:100%; border-collapse:separate; border-spacing:0; font-size:.84rem; background:{SUP};
  border:1px solid {BORDA}; border-radius:14px; overflow:hidden;}}
.dre th {{background:{SUP2}; color:{TXT2}; font-weight:600; font-size:.72rem; letter-spacing:.08em;
  text-transform:uppercase; padding:11px 14px; text-align:right; border-bottom:1px solid {BORDA};}}
.dre th:first-child {{text-align:left;}}
.dre th.best {{color:{VERDE};}}
.dre td {{padding:7px 14px; border-bottom:1px solid {GRADE}; color:{TXT}; text-align:right;
  font-variant-numeric: tabular-nums; border-left:none; border-right:none; border-top:none;}}
.dre td:first-child {{text-align:left; color:{TXT2}; padding-left:28px;}}
.dre tr.tot td {{background: rgba(76,141,255,.07); font-weight:700; color:{TXT};}}
.dre tr.tot td:first-child {{color:{TXT}; padding-left:14px;}}
.dre tr.final td {{background: rgba(76,141,255,.14); font-size:.92rem;}}
.dre tr.info td {{color:{TXT3}; font-style:italic;}}
.dre td.est {{color:{AMBAR};}}
.dre td.ne {{color:{TXT3};}}
.dre td.neg {{color:{VERMELHO};}}
.dre td .h {{display:block; font-size:.7rem; color:{TXT3}; font-weight:400;}}
.legend {{display:flex; gap:16px; font-size:.74rem; color:{TXT3}; margin-top:8px;}}
.legend i {{display:inline-block; width:8px; height:8px; border-radius:2px; margin-right:6px;}}
/* abas */
.stTabs [data-baseweb="tab-list"] {{gap: 4px; border-bottom:1px solid {BORDA};}}
.stTabs [data-baseweb="tab"] {{height: 40px; padding: 0 16px; font-weight:600; color:{TXT2};}}
.stTabs [aria-selected="true"] {{color:{TXT} !important;}}
div[data-testid="stExpander"] {{background:{SUP}; border:1px solid {BORDA}; border-radius:12px;}}
</style>
""", unsafe_allow_html=True)


# ------------------------------------------------------------ login (só e-mail @intercement.com)
from login import exigir_login  # noqa: E402

USUARIO = exigir_login()

# ------------------------------------------------------------ modo de cálculo (escolhido no topo da tela)
PLAN, CORR = "📗 Igual à planilha Excel", "✅ Corrigido (dados 2026)"
st.session_state.setdefault("modo_calc", PLAN)
MODO_PLANILHA = st.session_state.get("modo_calc") == PLAN

# ------------------------------------------------------------ helpers
@st.cache_resource(show_spinner="Carregando bases da planilha...")
def bases(caminho, _mtime):
    return carregar(caminho)


@st.cache_resource(show_spinner="Carregando base de resultados (1ª vez ~20 s)...")
def base_resultados(caminho, _mtime):
    return carregar_resultados(caminho)


def vazio(x):
    return x is None or (isinstance(x, float) and pd.isna(x))


def fmt(x, casas=2):
    if vazio(x):
        return "—"
    return f"{x:,.{casas}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def moeda(x, casas=0):
    return "—" if vazio(x) else ("-R$ " if x < 0 else "R$ ") + fmt(abs(x), casas)


def pct(x):
    return "—" if vazio(x) else f"{x:+.1%}".replace(".", ",")


def H(t: str) -> str:
    """Remove a indentação do HTML (linhas indentadas viram bloco de código no markdown)."""
    return "".join(linha.strip() for linha in t.splitlines())


def _k_opcao(x) -> str:
    """Chave para achar duplicadas: sem acento, sem pontuação/espaços, maiúsculas."""
    import re as _re
    import unicodedata as _ud
    t = _ud.normalize("NFKD", str(x)).encode("ascii", "ignore").decode().upper()
    return _re.sub(r"[^A-Z0-9]", "", t)


def _k_ordem(x) -> str:
    """Ordem alfabética como no Excel: ignora acento e maiúsculas, mantém espaços e pontuação."""
    import unicodedata as _ud
    return _ud.normalize("NFKD", str(x)).encode("ascii", "ignore").decode().upper()


def opcoes(serie_vol, *listas):
    """Todas as opções das listas, sem duplicadas (acento/espaço/maiúscula), com as que têm volume primeiro
    (ordem de volume) e o resto em ordem alfabética. Retorna (lista, {opção: volume})."""
    vol = {}
    if serie_vol is not None:
        for x, v in serie_vol.items():
            vol[str(x)] = float(v)
    escolhido = {}
    for x in list(vol) + [str(i) for lst in listas for i in (lst or []) if i is not None and str(i) not in ("", "nan")]:
        k = _k_opcao(x)
        if k and k not in escolhido:
            escolhido[k] = x
    return sorted(escolhido.values(), key=_k_ordem), vol


def ix_top(lista, vol, chave_ss=None):
    """Índice inicial: a opção com mais volume (só quando o filtro ainda não foi escolhido)."""
    if chave_ss and chave_ss in st.session_state:
        return None
    melhores = [x for x in sorted(vol, key=lambda k: -vol[k]) if x in lista]
    return lista.index(melhores[0]) if melhores else 0


def sb_sec(t):
    st.markdown(f'<div class="sb-sec">{t}</div>', unsafe_allow_html=True)


# ------------------------------------------------------------ barra lateral
with st.sidebar:
    st.markdown('<div class="sb-brand"><div class="sb-logo">SM</div><div><b>Simulador de Margens</b>'
                '<span>CANAIS BR</span></div></div>', unsafe_allow_html=True)
    sb_sec("Bases de dados")
    with st.expander("Enviar / atualizar planilhas", expanded=False):
        ups = st.file_uploader("Planilha de bases ou base de resultados do BI",
                               type=["xlsb", "xlsx", "xlsm", "csv"], accept_multiple_files=True,
                               help="Pode enviar qualquer formato. O simulador identifica sozinho se é a planilha "
                                    "de bases (abas Base de Custos, DE-PARA, Frete...) ou a base de resultados "
                                    "(colunas Ano/Mês, Expedição, Cidade, Canal, Produto, Volume...).")
        for up in ups or []:
            destino = os.path.join(PASTA, os.path.basename(up.name))
            dados = up.getbuffer()
            # só grava se mudou (o envio continua "anexado" a cada atualização da tela)
            if not os.path.exists(destino) or os.path.getsize(destino) != len(dados):
                with open(destino, "wb") as f:
                    f.write(dados)
            tipo, _ = identificar(destino)
            if tipo == "bases":
                st.success(f"✓ {up.name}: planilha de bases", icon=None)
            elif tipo == "resultados":
                st.success(f"✓ {up.name}: base de resultados", icon=None)
            else:
                st.error(f"✗ {up.name}: não reconheci este arquivo. A planilha de bases precisa das abas "
                         "'Base de Custos Fábrica', 'DE-PARA' e 'Frete Logístico'; a base de resultados precisa "
                         "das colunas Ano/Mês, Expedição, Cidade, Canal, Produto e Volume (t).")
        usar_r = st.toggle("Usar base de resultados", value=True)

    caminho, _ = achar(PASTA, "bases")
    if caminho is None and os.path.exists(PADRAO):
        caminho = PADRAO
    if caminho is None and os.path.exists(os.path.join(PASTA, EMBUTIDAS)):
        caminho = os.path.join(PASTA, EMBUTIDAS)   # bases jun/23 que vêm com o simulador
if not caminho:
    st.markdown(H(f"""<div class="top"><div><div class="eyebrow">Simulador de margens</div>
        <h1>Falta a planilha de bases</h1></div></div>"""), unsafe_allow_html=True)
    st.warning("Não encontrei a planilha de bases (a que tem as abas Base de Custos Fábrica, DE-PARA, Frete "
               f"Logístico…). Envie aqui ou copie para a pasta **{PASTA}**.")
    ups2 = st.file_uploader("Enviar planilhas", type=["xlsb", "xlsx", "xlsm", "csv"], accept_multiple_files=True,
                            key="up_main")
    for up in ups2 or []:
        destino = os.path.join(PASTA, os.path.basename(up.name))
        if not os.path.exists(destino) or os.path.getsize(destino) != len(up.getbuffer()):
            with open(destino, "wb") as f:
                f.write(up.getbuffer())
    if ups2:
        st.rerun() if achar(PASTA, "bases")[0] else None
    vistos = arquivos_da_pasta(PASTA)
    if vistos:
        st.markdown("**Arquivos que encontrei:**")
        for f in vistos[:15]:
            t, _ = identificar(f)
            st.markdown(f"- `{f}` → **{t or 'não reconhecido'}**" + (f" — {ERROS[f]}" if f in ERROS else ""))
    st.stop()
try:
    b = bases(caminho, os.path.getmtime(caminho))
except Exception as ex:
    st.error(f"Não consegui ler a planilha de bases **{os.path.basename(caminho)}**: {ex}")
    st.stop()
L = b.listas

with st.sidebar:
    cam_r = achar_base(PASTA)
    r = None
    erro_r = None
    if cam_r and usar_r:
        try:
            r = base_resultados(cam_r, os.path.getmtime(cam_r))
        except Exception as ex:
            erro_r = str(ex)
    nome_b = "bases jun/23 (embutidas)" if caminho.endswith(EMBUTIDAS) else os.path.basename(caminho)
    if r is not None:
        ms = r.meses
        st.markdown(f'<div class="sb-ok"><b>●</b> Resultados {ms[0][4:]}/{ms[0][:4]} – {ms[-1][4:]}/{ms[-1][:4]} · '
                    f'{len(r.df):,} linhas<br><span style="color:{TXT3}">{html.escape(os.path.basename(cam_r))}'
                    f'<br>+ {html.escape(nome_b)}</span></div>'.replace(",", "."), unsafe_allow_html=True)
    else:
        msg = (f"erro na base de resultados: {html.escape(erro_r)}" if erro_r else
               "base de resultados desligada" if cam_r else "envie a base de resultados do BI")
        st.markdown(f'<div class="sb-ok"><b style="color:{AMBAR}">●</b> Só {html.escape(nome_b)}'
                    f'<br><span style="color:{TXT3}">{msg}</span></div>', unsafe_allow_html=True)

cidades = sorted(b.cidades["Cidade IBGE"].astype(str).unique())
if r is not None:
    cidades = list(r.df["Cidade"].astype(str).unique()) + list(cidades)   # nomes da base nova têm preferência
    canais = r.canais()
else:
    canais = sorted(["Mat.Con.", "Construtora", "Industria", "Concreteira", "Consum.Final"], key=_k_ordem)
CANAL_NOVO = {"Construtora": "Construt. Predial"}

with st.sidebar:
    sb_sec("Cliente")
    cod = st.text_input("Código do cliente (emissor)", placeholder="opcional – preenche cidade e canal")
    if cod:
        k = norm(float(cod)) if cod.strip().isdigit() else norm(cod)
        if k in b.clientes.index:
            c = b.clientes.loc[k]
            cn = CANAL_NOVO.get(c["Canal"], c["Canal"]) if r is not None else c["Canal"]
            if st.session_state.get("_ult_cliente") != k:
                st.session_state.update(cidade=c["Cidade"], canal=cn, zona=c["Zona de Transporte"],
                                        _cid_cliente=True, _cid_cliente_canal=True, _ult_cliente=k)
            st.caption(f"✓ {c['Nome do Cliente']} · {c['Segmento Canal']}")
        else:
            st.caption("Cliente não encontrado")
    fabs, vol_fab = opcoes(r.volumes("Fábrica") if r is not None else None, list(L["produtos_por_fabrica"].keys()))
    if st.session_state.get("w_fab") not in fabs:
        st.session_state.pop("w_fab", None)
    fabrica = st.selectbox("Fábrica", fabs, key="w_fab", index=ix_top(fabs, vol_fab, "w_fab"))

    def com_hist(serie, todos):
        return opcoes(serie, todos)

    def rot_vol(vols):
        return lambda x: f"{x}   ·  {fmt(vols[x], 0)} t" if x in vols else str(x)

    v_cid = r.volumes("Cidade", fab=fabrica) if r is not None else None
    cidades, vol_cid = com_hist(v_cid, cidades)
    if st.session_state.get("cidade") in cidades and st.session_state.get("w_cid") != st.session_state["cidade"] \
            and st.session_state.get("_cid_cliente"):
        st.session_state["w_cid"] = st.session_state.pop("cidade")
        st.session_state.pop("_cid_cliente", None)
    if st.session_state.get("w_cid") not in cidades:
        st.session_state.pop("w_cid", None)
    cidade = st.selectbox("Cidade (UF_Cidade)", cidades, key="w_cid", format_func=rot_vol(vol_cid),
                          index=ix_top(cidades, vol_cid, "w_cid"),
                          help="Ordem alfabética; as cidades com venda desta fábrica mostram as toneladas do período")
    v_can = r.volumes("Canal", fab=fabrica, cidade=cidade) if r is not None else None
    canais, vol_can = com_hist(v_can, canais)
    if st.session_state.get("canal") in canais and st.session_state.get("_cid_cliente_canal"):
        st.session_state["w_can"] = st.session_state.pop("canal")
        st.session_state.pop("_cid_cliente_canal", None)
    if st.session_state.get("w_can") not in canais:
        st.session_state.pop("w_can", None)
    canal = st.selectbox("Canal", canais, key="w_can", format_func=rot_vol(vol_can), index=ix_top(canais, vol_can, "w_can"))
    zona = st.text_input("Zona de transporte (padrão BI)", st.session_state.get("zona", cidade))

    sb_sec("Venda")
    produto_nome = ""
    modo_planilha = MODO_PLANILHA
    if r is not None and r.produtos(fabrica) and not modo_planilha:
        v_pr = r.volumes("Produto", fab=fabrica, cidade=cidade, canal=canal)
        opts_p, vol_pr = com_hist(v_pr, r.produtos(fabrica) + sorted(r.df["Produto"].unique()))
        if st.session_state.get("w_prod") not in opts_p:
            st.session_state.pop("w_prod", None)
        produto_nome = st.selectbox("Produto", opts_p, key="w_prod", format_func=rot_vol(vol_pr),
                                    index=ix_top(opts_p, vol_pr, "w_prod"))
        produto = codigo_equivalente(b, fabrica, produto_nome)
    else:
        prods = sorted(L["produtos_por_fabrica"].get(fabrica) or [int(p) for p in b.produtos["Código do Produto"].dropna()])
        produto = st.selectbox("Produto", prods,
                               format_func=lambda p: f"{p} – {b.produtos.at[norm(p), 'Nomenclatura']}"
                               if norm(p) in b.produtos.index else str(p))
    # peso do saco do produto escolhido (para converter toneladas <-> sacos)
    if produto_nome:
        from resultados import info_produto as _ipq
        peso_sel = _ipq(produto_nome)["peso"]
    else:
        peso_sel = float(b.produtos.at[norm(produto), "Peso"]) if norm(produto) in b.produtos.index else 50.0
    c1, c2 = st.columns(2)
    unid_vol = c1.radio("Volume em", ["Toneladas", "Sacos"], horizontal=True, key="unid_vol")
    if unid_vol == "Toneladas" or peso_sel >= 1000:
        ton = c1.number_input("Volume (t)", 0.0, value=30.0, step=5.0, format="%.1f")
        sacos = ton * 1000 / peso_sel
        if peso_sel < 1000:
            c1.caption(f"= {fmt(sacos, 0)} sacos de {fmt(peso_sel, 0)} kg")
    else:
        sacos = c1.number_input("Quantidade (sacos)", 0, value=600, step=50)
        c1.caption(f"= {fmt(sacos * peso_sel / 1000, 1)} t")
    inc_opts = ["CIF", "FOB"]
    ix_inc = 0
    if r is not None and produto_nome:
        v_inc = r.volumes("Incoterm", fab=fabrica, cidade=cidade, canal=canal, produto=produto_nome)
        if len(v_inc) and v_inc.index[0] in inc_opts:
            ix_inc = inc_opts.index(v_inc.index[0])
    incoterm = c2.selectbox("Incoterm", inc_opts, index=ix_inc, key=f"inc|{fabrica}|{cidade}|{canal}|{produto_nome}")

    # preço sugerido = preço médio praticado (base de resultados) para esta combinação
    def preco_medio(inc):
        """Preço médio praticado (R$/saco; granel R$/t) na base de resultados para o incoterm."""
        if r is None or not produto_nome:
            return None
        from resultados import info_produto as _ip
        _peso = _ip(produto_nome)["peso"]
        exp_f = b.fabricas.at[norm(fabrica), "Expedição"] if norm(fabrica) in b.fabricas.index else None
        for filtros in ({"exp": exp_f, "cidade": cidade, "canal": canal, "incoterm": inc},
                        {"fab": fabrica, "cidade": cidade, "canal": canal, "incoterm": inc},
                        {"fab": fabrica, "cidade": cidade, "incoterm": inc},
                        {"fab": fabrica, "canal": canal, "incoterm": inc},
                        {"cidade": cidade, "canal": canal, "incoterm": inc},
                        {"fab": fabrica, "incoterm": inc}):
            hh = r.hist(produto=produto_nome, **filtros)
            if hh and hh["Preço Bruto"]:
                return hh["Preço Bruto"] * (_peso / 1000 if _peso < 1000 else 1)
        return None

    preco_hist = preco_medio(incoterm)
    preco = st.number_input(f"Preço negociado {incoterm} · R\\$/saco (granel R\\$/t)", 0.0,
                            value=round(float(preco_hist), 2) if preco_hist else 0.0, step=0.10, format="%.2f",
                            key=f"preco|{fabrica}|{cidade}|{canal}|{produto_nome or produto}")
    if preco_hist:
        st.caption(f"Preço médio praticado: {moeda(preco_hist, 2)} (base de resultados)")
    meses = None
    if r is not None:
        rot = {m: f"{m[4:]}/{m[:4]}" for m in r.meses}
        meses = st.multiselect("Período do histórico", r.meses, default=r.meses, format_func=rot.get,
                               help="Meses da base de resultados usados no histórico (média ponderada por volume)")

    sb_sec("Condições")
    cds = cds_da_fabrica(b, r, fabrica)
    if st.session_state.get("w_cd") not in [""] + cds:
        st.session_state.pop("w_cd", None)
    cds = sorted(cds, key=_k_ordem)
    cd = st.selectbox("CD", [""] + cds, format_func=lambda x: x or "— sem CD —", key="w_cd")
    c1, c2 = st.columns(2)
    chapa = c1.selectbox("Chapa", ["Não", "Sim"]) == "Sim"
    frac = c2.selectbox("Fracionamento", sorted([str(x) for x in b.fracionamento["Fracionamento"]], key=_k_ordem))
    preco_broker = st.number_input("Preço Broker · R$/saco (opcional)", 0.0, value=0.0, step=0.10, format="%.2f")

e = Entrada(fabrica=fabrica, cidade=cidade, canal=canal, produto=produto, incoterm=incoterm, sacos=sacos,
            zona=zona, preco_negociado=preco or None, chapa=chapa, fracionamento=frac, cd=cd,
            preco_broker=preco_broker or None, produto_nome=produto_nome, meses=meses or None)
ref0 = simular(b, e, r)["ref"]
with st.sidebar:
    with st.expander("Parâmetros avançados"):
        c1, c2 = st.columns(2)
        e.chapa_valor = c1.number_input("Chapa R$/sc", 0.0,
                                        value=0.0 if MODO_PLANILHA else float(ref0["chapa_padrao"]),
                                        step=0.05, format="%.2f", key=f"chapa_v|{MODO_PLANILHA}",
                                        help="Na planilha o custo de chapa fica 0; fora do modo planilha usa o valor "
                                             "da fábrica (DE-PARA)")
        if cd:
            e.taxa_filial = c2.number_input("Taxa filial R$/sc", 0.0, value=float(ref0["taxa_filial"]), step=0.05,
                                            format="%.2f")
        e.completar = st.toggle("Completar dados sem histórico", value=True,
                                help="Usa médias e frete médio da cidade onde não há histórico exato. "
                                     "Desligado = igual ao Excel (n/e).")
e.modo_excel = MODO_PLANILHA

# ------------------------------------------------------------ cálculo (com ajustes manuais)
base = simular(b, e, r)
ajustes = {}
for canal_ in CANAIS:
    estado = st.session_state.get(f"ed_{canal_}")
    if isinstance(estado, dict) and estado.get("edited_rows"):
        dfb = base["canais"][canal_]
        ed_in = dfb[dfb["Linha"].isin(EDITAVEIS)].reset_index(drop=True)
        aj = {}
        for i, mud in estado["edited_rows"].items():
            if "Simulação R$/t" in mud:
                v = mud["Simulação R$/t"]
                aj[ed_in.loc[int(i), "Linha"]] = None if v is None else float(v)
        if aj:
            ajustes[canal_] = aj
e.ajustes = ajustes
res = simular(b, e, r)
ref = res["ref"]
rs = resumo(res)
melhor = melhor_margem(res)
vol_t = res["volume_t"]
unid = "saco" if ref["ensacado"] else "t"

# ------------------------------------------------------------ escolha do modo de cálculo
mc1, mc2 = st.columns([2, 5])
with mc1:
    st.segmented_control("Modo de cálculo", [PLAN, CORR], key="modo_calc", label_visibility="collapsed")
with mc2:
    if MODO_PLANILHA:
        st.markdown(f'<div class="note" style="margin:8px 0 0">Calculando <b style="color:{TXT}">igual à planilha '
                    'Simulador_de_Margens jun/23</b> (mesmas bases e mesmas fórmulas, inclusive as que zeram o frete '
                    'CIF sem histórico, a movimentação de CDs e o Capex do CD). Os resultados devem bater com o Excel.'
                    '</div>', unsafe_allow_html=True)
    else:
        st.markdown(f'<div class="note" style="margin:8px 0 0"><b style="color:{AMBAR}">Resultado corrigido:</b> usa o '
                    'histórico real de 2026, cobra o frete CIF de verdade e a movimentação/Capex do CD. '
                    '<b>Fica diferente da planilha</b> — para conferir com o Excel, escolha "Igual à planilha".'
                    '</div>', unsafe_allow_html=True)

# ------------------------------------------------------------ topo
mgcx_tot = None if not melhor else rs.at["MGCX", melhor] * vol_t
chip_vol = f"{fmt(vol_t, 1)} t" + (f" · {fmt(sacos, 0)} sc" if ref["ensacado"] else "")
st.markdown(H(f"""
<div class="top">
  <div>
    <div class="eyebrow">Simulação de resultado de vendas</div>
    <h1>{html.escape(fabrica.title())} → {html.escape(cidade.replace('_', ' · '))}</h1>
    <div class="chips">
      <span class="chip">Canal <b>{html.escape(canal)}</b></span>
      <span class="chip">Produto <b>{html.escape(str(ref['nome']))}</b></span>
      <span class="chip">Incoterm <b>{incoterm}</b></span>
      <span class="chip">Volume <b>{chip_vol}</b></span>
      <span class="chip">Preço <b>{moeda(preco, 2) if preco else 'não informado'}</b>/{unid}</span>
      {f'<span class="chip">CD <b>{html.escape(cd)}</b></span>' if cd else ''}
    </div>
  </div>
  <div class="right"><div class="k">Margem caixa no volume</div>
    <div class="v {'dn' if mgcx_tot is not None and mgcx_tot < 0 else ''}">{moeda(mgcx_tot)}</div></div>
</div>"""), unsafe_allow_html=True)

if not preco:
    st.info("Informe o **preço negociado** na barra lateral para simular Fábrica e CD.", icon="💡")

# ------------------------------------------------------------ destaque + indicadores
banda = (preco / ref["preco_lista"] - 1) if preco and ref["preco_lista"] else None
premio = ref["premio_ap"]
if melhor:
    v = rs.at["MGCX", melhor]
    rb_m = res["canais"][melhor].set_index("Linha").at["Preço Bruto", "Simulação R$/t"]
    marg_pct = None if vazio(rb_m) or not rb_m else v / rb_m
    seg = sorted([x for x in rs.loc["MGCX"].dropna().astype(float)], reverse=True)
    dif = (seg[0] - seg[1]) if len(seg) > 1 else None
    rec = f"""<div class="panel rec {'neg' if v < 0 else ''}">
      <div class="k">Formato recomendado · maior margem caixa</div>
      <div class="ch">{melhor}</div>
      <div class="val">{fmt(v)} <small>R$ por tonelada</small></div>
      <div class="foot"><div>Margem s/ receita bruta<b>{'—' if marg_pct is None else f'{marg_pct:.1%}'.replace('.', ',')}</b></div>
        <div>Vantagem s/ 2º formato<b>{'—' if dif is None else fmt(dif) + ' R$/t'}</b></div>
        <div>Total no volume<b>{moeda(v * vol_t)}</b></div></div></div>"""
else:
    rec = f"""<div class="panel rec neg"><div class="k">Formato recomendado</div>
      <div class="ch" style="color:{TXT3}">Sem dados suficientes</div>
      <div class="foot"><div>Informe o preço e confira as linhas sem histórico.</div></div></div>"""


def kpi(k, v, s="", cls=""):
    return f'<div class="kpi"><div class="k">{k}</div><div class="v {cls}">{v}</div><div class="s">{s}</div></div>'


fob_h, cif_h, fc = ref["preco_fob_hist"], ref["preco_cif_hist"], ref["frete_comercial"]
kpis = "".join([
    kpi("Preço lista microrregião", moeda(ref["preco_lista"], 2),
        f'Banda <span class="{"up" if banda and banda >= 0 else "dn"}">{pct(banda)}</span>' if banda is not None else "sem preço"),
    kpi("Preço FOB histórico", f"{fmt(fob_h['Fábrica'])} <span style='color:{TXT3}'>/</span> {fmt(fob_h['CD'])}",
        "Fábrica / CD · R$/saco"),
    kpi("Preço CIF histórico", f"{fmt(cif_h['Fábrica'])} <span style='color:{TXT3}'>/</span> {fmt(cif_h['CD'])}",
        "Fábrica / CD · R$/saco"),
    kpi("Frete comercial", f"{fmt(fc['Fábrica'])} <span style='color:{TXT3}'>/</span> {fmt(fc['CD'])}",
        "Fábrica / CD · R$/saco"),
    kpi("AP / Distribuidor", html.escape(ref["ap"] if ref["ap"] != "-" else "—"),
        "" if premio is None else f"prêmio {premio:.1%}".replace(".", ",")
        + (f" · venda AP {moeda(preco * (1 + premio), 2)}" if preco else "")),
    kpi("Taxa filial (CD)", moeda(ref["taxa_filial"], 2) if cd else "—", "R$/saco" if cd else "sem CD"),
])
st.markdown(H(f'<div class="hero">{rec}<div class="panel"><div class="kpis">{kpis}</div></div></div>'),
            unsafe_allow_html=True)

# ------------------------------------------------------------ cartões dos formatos
st.markdown('<div class="sec"><h3>Comparativo por formato de atendimento</h3><span>R$/t · sem incentivo</span></div>',
            unsafe_allow_html=True)
nomes = {"Fábrica": ref["expedicao"], "CD": cd or "nenhum CD selecionado",
         "DT": ref["ap"] if ref["ap"] != "-" else ("distribuidores da cidade"
                                                   if res["fonte"]["DT"].startswith("base resultados") else "sem AP/DT na cidade"),
         "Broker": ref["broker"] if ref["broker"] != "-" else ("brokers da cidade"
                                                   if res["fonte"]["Broker"].startswith("base resultados") else "sem carteira na cidade")}
mx = max([abs(float(x)) for x in rs.loc["MGCX"].dropna()] or [1]) or 1
cards = []
for canal_ in CANAIS:
    v = rs[canal_]
    mg = v["MGCX"]
    na = vazio(mg)
    neg = not na and mg < 0
    pill = ('<span class="pill best">MELHOR</span>' if canal_ == melhor else
            '<span class="pill neg">NEGATIVA</span>' if neg else
            '<span class="pill na">SEM DADOS</span>' if na else "")
    larg = 0 if na else min(100, abs(mg) / mx * 100)
    cor = VERMELHO if neg else (VERDE if canal_ == melhor else AZUL)
    faltam = res["canais"][canal_].query("Tipo != 'total' and Origem.str.startswith('n/e')", engine="python")
    aviso = ""
    if len(faltam):
        motivo = faltam["Origem"].iloc[0][6:]
        aviso = f'<div class="warn">⚠ {html.escape(motivo if len(faltam) > 3 else ", ".join(faltam["Linha"]))}</div>'
    linhas = "".join(f"<tr><td>{k} <span style='color:{TXT3};font-size:.72rem'>R$/t</span></td>"
                     f"<td class='{'na-txt' if vazio(v[k]) else ''}'>{fmt(v[k])}</td></tr>"
                     for k in ("FOB Líquido", "MCU", "MGS"))
    saco = "" if na else f"{moeda(mg * vol_t)} em {fmt(vol_t, 1)} t"
    cards.append(f"""<div class="card {'best' if canal_ == melhor else ''}" style="--c:{COR[canal_]}">
      <div class="hd"><span class="t"><span class="dot"></span>{canal_}</span>{pill}</div>
      <div class="sub">{html.escape(str(nomes[canal_]))}</div>
      <div class="src">{html.escape(res['fonte'][canal_])}</div>
      <div class="big {'neg' if neg else 'na' if na else ''}">{fmt(mg)}<small>MGCX R$/t</small></div>
      <div class="saco">{saco}&nbsp;</div>
      <div class="bar"><i style="width:{larg:.0f}%; background:{cor}"></i></div>
      <table>{linhas}</table>{aviso}</div>""")
st.markdown(H(f'<div class="cards">{"".join(cards)}</div>'), unsafe_allow_html=True)
st.markdown('<div class="note">O apontamento de melhor margem não implica direcionamento da venda; '
            'analise receitas e custos com o seu gestor e com Pricing.</div>', unsafe_allow_html=True)

# ------------------------------------------------------------ CIF × FOB na simulação
from dataclasses import replace as _replace  # noqa: E402

if preco:
    outro = "FOB" if incoterm == "CIF" else "CIF"
    st.markdown('<div class="sec"><h3>CIF × FOB</h3><span>a coluna do incoterm escolhido é o mesmo cálculo dos cartões '
                'acima; a outra usa o mesmo preço (pode ajustar)</span></div>', unsafe_allow_html=True)
    cp1, cp2, _cp3 = st.columns([1, 1, 3])
    cp1.markdown(f'<div class="kpi" style="padding:0"><div class="k">Preço {incoterm} (barra lateral)</div>'
                 f'<div class="v">{moeda(preco, 2)} <span style="font-size:.8rem;color:{TXT2}">/{unid}</span></div></div>',
                 unsafe_allow_html=True)
    sug = None if e.modo_excel else preco_medio(outro)
    p_outro = cp2.number_input(f"Preço {outro} (R$/{unid})", 0.0, value=round(float(preco), 2), step=0.10, format="%.2f",
                               help="Começa com o mesmo preço da barra lateral (como na planilha)."
                                    + (f" Preço médio praticado {outro}: {moeda(sug, 2)}." if sug else ""),
                               key=f"p_outro|{outro}|{fabrica}|{cidade}|{canal}|{produto}|{produto_nome}|{preco}")
    precos_inc = {incoterm: (preco, "preço da barra lateral"),
                  outro: (p_outro, "mesmo preço" if abs(p_outro - preco) < 1e-9 else "preço ajustado")}
    res_inc = {incoterm: rs}       # exatamente o resultado dos cartões (inclui ajustes manuais)
    try:
        res_inc[outro] = resumo(simular(b, _replace(e, incoterm=outro, preco_negociado=p_outro, ajustes=ajustes), r))
    except Exception:
        res_inc[outro] = pd.DataFrame()
    COR_INC = {"CIF": "#4c8dff", "FOB": "#a48bff"}
    linhas_inc, melhor_geral = "", None
    for c_ in CANAIS:
        def _g(inc_, lin):
            df_ = res_inc[inc_]
            return None if df_.empty or c_ not in df_.columns else df_.at[lin, c_]
        v_cif, v_fob = _g("CIF", "MGCX"), _g("FOB", "MGCX")
        v_cif = None if vazio(v_cif) else float(v_cif)
        v_fob = None if vazio(v_fob) else float(v_fob)
        val = {k: v for k, v in (("CIF", v_cif), ("FOB", v_fob)) if v is not None}
        venc = max(val, key=val.get) if val else "—"
        if val and (melhor_geral is None or max(val.values()) > melhor_geral[2]):
            melhor_geral = (c_, venc, max(val.values()))
        tds = ""
        for inc_, v in (("CIF", v_cif), ("FOB", v_fob)):
            est = f"background:rgba(47,191,113,.16);color:{VERDE};font-weight:800" if inc_ == venc else ""
            det = (f'<span class="h">FOB líq. {fmt(_g(inc_, "FOB Líquido"))} · MCU {fmt(_g(inc_, "MCU"))} · '
                   f'MGS {fmt(_g(inc_, "MGS"))}</span>') if v is not None else ""
            tds += f'<td class="{"ne" if v is None else ("neg" if v < 0 else "")}" style="{est}">{fmt(v)}{det}</td>'
        dif = "" if v_cif is None or v_fob is None else fmt(abs(v_cif - v_fob))
        linhas_inc += (f'<tr><td style="color:{TXT};font-weight:600;padding-left:14px"><span style="color:{COR[c_]}">●</span> '
                       f'{c_}</td>{tds}<td style="text-align:center;font-weight:700">{venc}</td><td>{dif}</td></tr>')
    sub_p = " · ".join(f"{k}: {moeda(v[0], 2)}/{unid} ({v[1]})" for k, v in precos_inc.items())
    st.markdown(H(f'<div class="note" style="margin:0 0 8px">{html.escape(sub_p)}</div>'
                  f'<table class="dre"><tr><th>MGCX R$/t</th><th><span style="color:{COR_INC["CIF"]}">●</span> CIF</th>'
                  f'<th><span style="color:{COR_INC["FOB"]}">●</span> FOB</th><th>Melhor</th><th>Diferença</th></tr>'
                  f'{linhas_inc}</table>'), unsafe_allow_html=True)
    if melhor_geral:
        st.markdown(f"Melhor combinação: **{melhor_geral[0]} em {melhor_geral[1]}**, com margem caixa de "
                    f"**{fmt(melhor_geral[2])} R\\$/t**.")
    st.markdown('<div class="note"></div>', unsafe_allow_html=True)


# ------------------------------------------------------------ DRE comparativa
def dre_html(por_saco=False, hist=False):
    dfs = {c: res["canais"][c].set_index("Linha") for c in CANAIS}
    linhas = list(dfs["Fábrica"].index)
    tipos = dfs["Fábrica"]["Tipo"]
    fat = (ref["peso"] / 1000) if (por_saco and ref["ensacado"]) else 1
    th = "".join(f'<th class="{"best" if c == melhor else ""}"><span style="color:{COR[c]}">●</span> {c}'
                 f'{" ★" if c == melhor else ""}</th>' for c in CANAIS)
    corpo = []
    for n in linhas:
        tp = tipos[n]
        cls = "tot" + (" final" if n == "Margem Caixa" else "") if tp == "total" else ("info" if tp == "info" else "")
        tds = []
        for c in CANAIS:
            val = dfs[c].at[n, "Simulação R$/t"]
            h = dfs[c].at[n, "Histórico R$/t"]
            org = str(dfs[c].at[n, "Origem"])
            k = []
            if vazio(val):
                k.append("ne")
            elif tp == "total" and val < 0:
                k.append("neg")
            elif org.startswith(("média", "igual", "ajuste", "frete", "base resultados: frete")):
                k.append("est")
            txt = fmt(None if vazio(val) else val * fat)
            sub = f'<span class="h">hist. {fmt(None if vazio(h) else h * fat)}</span>' if hist and tp != "info" else ""
            tds.append(f'<td class="{" ".join(k)}" title="{html.escape(org)}">{txt}{sub}</td>')
        nome = n + (" (não soma)" if tp == "info" else "")
        corpo.append(f'<tr class="{cls}"><td>{nome}</td>{"".join(tds)}</tr>')
    un = "R$/saco" if fat != 1 else "R$/t"
    return (f'<table class="dre"><thead><tr><th>{un}</th>{th}</tr></thead><tbody>{"".join(corpo)}</tbody></table>'
            f'<div class="legend"><span><i style="background:{AMBAR}"></i>estimado (média / frete médio / ajuste)</span>'
            f'<span><i style="background:{TXT3}"></i>sem dado</span>'
            f'<span>passe o mouse sobre um valor para ver a origem</span></div>')


def tema(ch):
    return (ch.configure(background="transparent", font="Inter")
            .configure_view(stroke=None)
            .configure_axis(labelColor=TXT2, titleColor=TXT3, tickColor=GRADE, domainColor=GRADE,
                            gridColor=GRADE, labelFontSize=11, titleFontSize=11, titleFontWeight=500)
            .configure_legend(labelColor=TXT2, labelFontSize=11, symbolType="square"))


tabd, tab0, tab1, tab2, tabp, tab3, tab4 = st.tabs(
    ["🎯  Por destino", "🧭  Simular rotas", "📊  DRE comparativa", "📈  Gráficos", "💲  Preço mínimo",
     "✏️  Ajustar premissas", "⬇  Exportar"])

# ============================================================ ABA: POR DESTINO (resultado real 2026)
FMT_COR = {"Fábrica": COR["Fábrica"], "CD": COR["CD"], "Distribuidor": COR["DT"]}
FORMATOS = ["Fábrica", "CD", "Distribuidor"]
with tabd:
    if r is None:
        st.info("Esta comparação usa a base de resultados do BI. Coloque o arquivo na pasta do simulador.", icon="💡")
    else:
        st.markdown('<div class="sec"><h3>Comparar por destino</h3><span>resultado real da base de resultados · '
                    'margem caixa em R$/t, ponderada pelo volume</span></div>', unsafe_allow_html=True)
        f1_, f2_, f3_, f4_ = st.columns([1, 2, 2, 2])
        ufs = sorted(r.ufs())
        uf_d = f1_.selectbox("Estado", ufs, index=ufs.index(cidade[:2]) if cidade[:2] in ufs else 0, key="d_uf")
        v_c = r.volumes("Cidade", uf=uf_d)
        todas_cid = [c for c in cidades if str(c)[:2].upper() == uf_d]
        lista_c, vol_c = opcoes(v_c, todas_cid)
        cid_d = f2_.selectbox("Cidade", ["Todas as cidades"] + lista_c, key="d_cid",
                              format_func=lambda x: x if x == "Todas as cidades" else
                              (f"{x}  ·  {fmt(vol_c[x], 0)} t" if x in vol_c else f"{x}  ·  sem venda"))
        filtro = {"uf": uf_d} if cid_d == "Todas as cidades" else {"cidade": cid_d}
        lista_can, vol_can_d = opcoes(r.volumes("Canal", **filtro), r.canais())
        can_d = f3_.selectbox("Canal do cliente", ["Todos"] + lista_can, key="d_can",
                              format_func=lambda x: x if x == "Todos" else
                              (f"{x}  ·  {fmt(vol_can_d[x], 0)} t" if x in vol_can_d else f"{x}  ·  sem venda"))
        if can_d != "Todos":
            filtro["canal"] = can_d
        lista_pr, v_pr_d = opcoes(r.volumes("Produto", **filtro), sorted(r.df["Produto"].unique()))
        prod_d = f4_.selectbox("Cimento", ["Todos os cimentos"] + lista_pr, key="d_prod",
                               format_func=lambda x: x if x.startswith("Todos") else
                               (f"{x}  ·  {fmt(v_pr_d[x], 0)} t" if x in v_pr_d else f"{x}  ·  sem venda"))
        if not prod_d.startswith("Todos"):
            filtro["produto"] = prod_d
        g1_, g3_, g2_ = st.columns([5, 2, 2])
        inc_d = g3_.selectbox("Incoterm", ["CIF + FOB", "CIF", "FOB"], key="d_inc",
                              help="CIF + FOB = todas as vendas; escolha um para filtrar")
        if inc_d != "CIF + FOB":
            filtro["incoterm"] = inc_d
        fabs_d, v_fab_d = opcoes(r.volumes("Fábrica", **filtro), list(L["produtos_por_fabrica"].keys()),
                                 list(r.df["Fábrica"].unique()))
        sel_fabs = g1_.multiselect("Fábricas para comparar (fábrica 1, 2, 3…)", fabs_d,
                                   default=sorted([f for f in fabs_d if f in v_fab_d], key=lambda f: -v_fab_d[f])[:3],
                                   key=f"d_fabs|{uf_d}|{cid_d}|{can_d}|{prod_d}",
                                   format_func=lambda x: f"{x.title()} · {fmt(v_fab_d[x], 0)} t" if x in v_fab_d
                                   else f"{x.title()} · sem venda")
        min_v = g2_.number_input("Volume mínimo por item (t)", 0, value=30, step=10,
                                 help="Ignora combinações com pouco volume (resultado pouco confiável)")
        nome_dest = cid_d.replace("_", " · ") if cid_d != "Todas as cidades" else f"estado de {uf_d}"
        mes_d = meses or None

        por_fab = r.agrupar("Fábrica", meses=mes_d, min_vol=min_v, fab=sel_fabs, **filtro)
        por_fmt = r.agrupar("_formato", meses=mes_d, min_vol=min_v, fab=sel_fabs, **filtro)
        por_ff = r.agrupar(["Fábrica", "_formato"], meses=mes_d, min_vol=min_v, fab=sel_fabs, **filtro)
        por_prod = r.agrupar("Produto", meses=mes_d, min_vol=min_v, fab=sel_fabs,
                             **{k: v for k, v in filtro.items() if k != "produto"})
        por_cd = r.agrupar(["Expedição", "Fábrica"], meses=mes_d, min_vol=min_v, fab=sel_fabs, formato="CD", **filtro)
        por_dist = r.agrupar("Fábrica", meses=mes_d, min_vol=min_v, fab=sel_fabs, formato="Distribuidor", **filtro)

        if por_fab.empty:
            st.warning("Sem volume para esse destino/filtro. Escolha outras fábricas, outro canal ou reduza o volume mínimo.")
        else:
            # ---------------- veredito
            def melhor_linha(df, col):
                return None if df is None or df.empty else df.iloc[0]

            bf, bfm = melhor_linha(por_fab, "Fábrica"), melhor_linha(por_fmt, "_formato")
            bp, bc = melhor_linha(por_prod, "Produto"), melhor_linha(por_cd, "Expedição")
            bff = melhor_linha(por_ff, "Fábrica")

            def bloco(k, titulo, sub, val, cor=TXT):
                return (f'<div><div class="k">{k}</div><div class="v" style="color:{cor}">{titulo}</div>'
                        f'<div class="s">{sub}</div><div class="s" style="font-weight:700;color:{VERDE if (val or 0) >= 0 else VERMELHO}">'
                        f'{fmt(val)} R$/t</div></div>')

            st.markdown(H(f"""<div class="route" style="grid-template-columns:1.3fr 1fr 1fr 1fr 1fr">
              <div><div class="k">Melhor caminho para {html.escape(nome_dest)}</div>
                <div class="v big">{html.escape(bff['Fábrica'].title())} → {html.escape(bff['_formato'])}</div>
                <div class="s">{fmt(bff['Volume (t)'], 0)} t vendidas · {moeda(bff['MGCX total (R$)'])} de margem caixa</div>
                <div class="s" style="font-weight:700;color:{VERDE if bff['MGCX'] >= 0 else VERMELHO};font-size:1rem">
                {fmt(bff['MGCX'])} R$/t</div></div>
              {bloco("Melhor fábrica", html.escape(bf['Fábrica'].title()), f"{fmt(bf['Volume (t)'], 0)} t", bf['MGCX'], COR['Fábrica'])}
              {bloco("Melhor formato", html.escape(bfm['_formato']), f"{fmt(bfm['Volume (t)'], 0)} t", bfm['MGCX'], FMT_COR[bfm['_formato']])}
              {bloco("Melhor CD", html.escape(bc['Expedição']) if bc is not None else "—",
                     (html.escape(bc['Fábrica'].title()) + f" · {fmt(bc['Volume (t)'], 0)} t") if bc is not None else "sem volume via CD",
                     bc['MGCX'] if bc is not None else None, COR['CD'])}
              {bloco("Melhor cimento", html.escape(str(bp['Produto']).replace("CIMENTO ", ""))[:40] if bp is not None else "—",
                     f"{fmt(bp['Volume (t)'], 0)} t" if bp is not None else "", bp['MGCX'] if bp is not None else None, TXT)}
              </div>"""), unsafe_allow_html=True)

            # ---------------- 1. fábrica × fábrica
            st.markdown('<div class="sec"><h3>1 · Fábrica × Fábrica</h3><span>todas as vendas de cada fábrica para o '
                        'destino · 🏆 = melhor margem caixa</span></div>', unsafe_allow_html=True)
            mx_f = max(abs(por_fab["MGCX"]).max(), 1)
            cards_f = []
            for pos, (_, x) in enumerate(por_fab.iterrows()):
                sub = por_ff[por_ff["Fábrica"] == x["Fábrica"]].set_index("_formato")
                linhas_ = "".join(
                    f"<tr><td><span style='color:{FMT_COR[f_]}'>●</span> {f_}</td>"
                    f"<td class='{'na-txt' if f_ not in sub.index else ''}'>"
                    f"{fmt(sub.at[f_, 'MGCX']) if f_ in sub.index else '—'}</td></tr>" for f_ in FORMATOS)
                w = abs(x["MGCX"]) / mx_f * 100
                cards_f.append(f"""<div class="card {'best' if pos == 0 else ''}" style="--c:{COR['Fábrica']}">
                  <div class="hd"><span class="t"><span class="dot"></span>{pos + 1}º · {html.escape(x['Fábrica'].title())}</span>
                  {'<span class="pill best">🏆 MELHOR</span>' if pos == 0 else ''}</div>
                  <div class="sub">{fmt(x['Volume (t)'], 0)} t · preço {moeda(x['Preço bruto'], 0)}/t · frete {fmt(x['Frete+log.'], 0)}/t</div>
                  <div class="big {'neg' if x['MGCX'] < 0 else ''}">{fmt(x['MGCX'])}<small>MGCX R$/t</small></div>
                  <div class="saco">{moeda(x['MGCX total (R$)'])} no período</div>
                  <div class="bar"><i style="width:{w:.0f}%;background:{VERMELHO if x['MGCX'] < 0 else (VERDE if pos == 0 else COR['Fábrica'])}"></i></div>
                  <table>{linhas_}</table></div>""")
            ncol_f = min(max(len(cards_f), 1), 4)
            st.markdown(H(f'<div class="cards" style="grid-template-columns:repeat({ncol_f},1fr)">{"".join(cards_f)}</div>'),
                        unsafe_allow_html=True)
            if len(por_fab) > 1:
                a_, b2_ = por_fab.iloc[0], por_fab.iloc[1]
                st.markdown(f"**{a_['Fábrica'].title()}** é a melhor fábrica para {nome_dest}: margem caixa "
                            f"{fmt(a_['MGCX'])} R\\$/t, **{fmt(a_['MGCX'] - b2_['MGCX'])} R\\$/t acima** de "
                            f"{b2_['Fábrica'].title()}.")

            # ---------------- 2. fábrica × CD × distribuidor
            st.markdown('<div class="sec"><h3>2 · Fábrica direto × CD × Distribuidor</h3><span>qual formato de '
                        'atendimento rende mais nesse destino</span></div>', unsafe_allow_html=True)
            ca_, cb_ = st.columns([2, 3], gap="large")
            with ca_:
                pf = por_fmt.set_index("_formato")
                melhor_f = por_fmt.iloc[0]["_formato"]
                cards_fmt = []
                for f_ in FORMATOS:
                    if f_ not in pf.index:
                        cards_fmt.append(f'<div class="card" style="--c:{FMT_COR[f_]}"><div class="hd"><span class="t">'
                                         f'<span class="dot"></span>{f_}</span><span class="pill na">SEM VOLUME</span></div>'
                                         f'<div class="big na">—</div></div>')
                        continue
                    x = pf.loc[f_]
                    cards_fmt.append(f"""<div class="card {'best' if f_ == melhor_f else ''}" style="--c:{FMT_COR[f_]}">
                      <div class="hd"><span class="t"><span class="dot"></span>{f_}</span>
                      {'<span class="pill best">🏆 MELHOR</span>' if f_ == melhor_f else ''}</div>
                      <div class="big {'neg' if x['MGCX'] < 0 else ''}">{fmt(x['MGCX'])}<small>MGCX R$/t</small></div>
                      <div class="sub">{fmt(x['Volume (t)'], 0)} t · preço {moeda(x['Preço bruto'], 0)}/t · FOB {fmt(x['FOB líquido'], 0)}</div></div>""")
                st.markdown(H(f'<div style="display:grid;gap:10px">{"".join(cards_fmt)}</div>'), unsafe_allow_html=True)
            with cb_:
                hm = por_ff.rename(columns={"_formato": "Formato"}).copy()
                hm["Fábrica"] = hm["Fábrica"].str.title()
                hm["Rótulo"] = hm["MGCX"].map(lambda v: fmt(v, 0))
                lim = float(max(abs(hm["MGCX"]).max(), 1))
                ordem_f = [f.title() for f in por_fab["Fábrica"]]
                heat = alt.Chart(hm).mark_rect(cornerRadius=6, stroke=BG, strokeWidth=3).encode(
                    x=alt.X("Formato:N", sort=FORMATOS, title=None, axis=alt.Axis(orient="top", labelFontSize=12,
                                                                                  labelColor=TXT, labelAngle=0)),
                    y=alt.Y("Fábrica:N", sort=ordem_f, title=None, axis=alt.Axis(labelFontSize=12, labelColor=TXT)),
                    color=alt.Color("MGCX:Q", scale=alt.Scale(domain=[-lim, 0, lim], range=[VERMELHO, "#2a3140", VERDE]),
                                    legend=alt.Legend(title="MGCX R$/t", orient="bottom", gradientLength=260)),
                    tooltip=["Fábrica:N", "Formato:N", alt.Tooltip("MGCX:Q", format=",.1f", title="MGCX R$/t"),
                             alt.Tooltip("Volume (t):Q", format=",.0f"), alt.Tooltip("Preço bruto:Q", format=",.0f"),
                             alt.Tooltip("Frete+log.:Q", format=",.1f")])
                tx = alt.Chart(hm).mark_text(fontSize=13, fontWeight=700, color="#ffffff").encode(
                    x=alt.X("Formato:N", sort=FORMATOS), y=alt.Y("Fábrica:N", sort=ordem_f), text="Rótulo:N")
                st.altair_chart(tema((heat + tx).properties(height=alt.Step(58))), width="stretch")
                st.caption("Mapa de calor: cada célula é a margem caixa (R$/t) da fábrica naquele formato. "
                           "Verde = melhor, vermelho = negativa.")
            if len(por_fmt) > 1:
                a_, b2_ = por_fmt.iloc[0], por_fmt.iloc[1]
                st.markdown(f"Para {nome_dest}, **{a_['_formato']}** é o formato com melhor margem caixa "
                            f"({fmt(a_['MGCX'])} R\\$/t), **{fmt(a_['MGCX'] - b2_['MGCX'])} R\\$/t acima** de {b2_['_formato']}.")

            # ---------------- 3. cimento e CD
            cc1, cc2 = st.columns(2, gap="large")
            with cc1:
                st.markdown('<div class="sec"><h3>3 · Qual cimento</h3><span>margem caixa por produto no destino</span>'
                            '</div>', unsafe_allow_html=True)
                if por_prod.empty:
                    st.caption("Sem produtos com volume suficiente.")
                else:
                    pp = por_prod.head(12).copy()
                    pp["Rótulo"] = pp["MGCX"].map(lambda v: fmt(v, 1))
                    pp["Destaque"] = ["Melhor"] + ["Demais"] * (len(pp) - 1)
                    ordem_p = list(pp["Produto"])
                    bp_ = alt.Chart(pp).encode(y=alt.Y("Produto:N", sort=ordem_p, title=None,
                                                       axis=alt.Axis(labelLimit=300, labelColor=TXT, labelFontSize=11)))
                    st.altair_chart(tema((
                        bp_.mark_bar(cornerRadiusEnd=5, height=18).encode(
                            x=alt.X("MGCX:Q", title="MGCX R$/t", axis=alt.Axis(domain=False)),
                            color=alt.Color("Destaque:N", scale=alt.Scale(domain=["Melhor", "Demais"],
                                                                          range=[VERDE, "#5b8def"]), legend=None),
                            tooltip=["Produto:N", alt.Tooltip("MGCX:Q", format=",.1f"),
                                     alt.Tooltip("Volume (t):Q", format=",.0f"), alt.Tooltip("Preço bruto:Q", format=",.0f")])
                        + bp_.mark_text(align="left", dx=5, fontSize=11, fontWeight=700, color=TXT).encode(
                            x="MGCX:Q", text="Rótulo:N")).properties(height=max(160, 30 * len(pp)))), width="stretch")
            with cc2:
                st.markdown('<div class="sec"><h3>4 · Qual CD</h3><span>margem caixa por CD (e fábrica que abastece)'
                            '</span></div>', unsafe_allow_html=True)
                if por_cd.empty:
                    st.caption("Sem vendas via CD com volume suficiente para esse destino.")
                else:
                    pc = por_cd.head(12).copy()
                    pc["Rota"] = pc["Expedição"] + "  ·  " + pc["Fábrica"].str.title()
                    pc["Rótulo"] = pc["MGCX"].map(lambda v: fmt(v, 1))
                    ordem_c = list(pc["Rota"])
                    bc_ = alt.Chart(pc).encode(y=alt.Y("Rota:N", sort=ordem_c, title=None,
                                                       axis=alt.Axis(labelLimit=300, labelColor=TXT, labelFontSize=11)))
                    st.altair_chart(tema((
                        bc_.mark_bar(cornerRadiusEnd=5, height=18, color=COR["CD"]).encode(
                            x=alt.X("MGCX:Q", title="MGCX R$/t", axis=alt.Axis(domain=False)),
                            tooltip=["Rota:N", alt.Tooltip("MGCX:Q", format=",.1f"),
                                     alt.Tooltip("Volume (t):Q", format=",.0f"), alt.Tooltip("Frete+log.:Q", format=",.1f")])
                        + bc_.mark_text(align="left", dx=5, fontSize=11, fontWeight=700, color=TXT).encode(
                            x="MGCX:Q", text="Rótulo:N")).properties(height=max(160, 30 * len(pc)))), width="stretch")

            # ---------------- 5. CIF × FOB
            f_sem_inc = {k: v for k, v in filtro.items() if k != "incoterm"}
            por_inc = r.agrupar("Incoterm", meses=mes_d, min_vol=min_v, fab=sel_fabs, **f_sem_inc)
            por_fi = r.agrupar(["Fábrica", "Incoterm"], meses=mes_d, min_vol=min_v, fab=sel_fabs, **f_sem_inc)
            por_fmi = r.agrupar(["_formato", "Incoterm"], meses=mes_d, min_vol=min_v, fab=sel_fabs, **f_sem_inc)
            st.markdown('<div class="sec"><h3>5 · CIF × FOB</h3><span>mesmo destino e fábricas · no FOB o cliente '
                        'retira (sem frete de vendas), por isso o preço é menor</span></div>', unsafe_allow_html=True)
            if por_inc.empty:
                st.caption("Sem volume para comparar CIF e FOB.")
            else:
                pi = por_inc.set_index("Incoterm")
                melhor_i = por_inc.iloc[0]["Incoterm"]
                COR_INC = {"CIF": "#4c8dff", "FOB": "#a48bff"}
                ci1, ci2 = st.columns([2, 3], gap="large")
                with ci1:
                    cards_i = []
                    for inc_ in ("CIF", "FOB"):
                        if inc_ not in pi.index:
                            cards_i.append(f'<div class="card" style="--c:{COR_INC[inc_]}"><div class="hd"><span class="t">'
                                           f'<span class="dot"></span>{inc_}</span><span class="pill na">SEM VOLUME</span>'
                                           f'</div><div class="big na">—</div></div>')
                            continue
                        x = pi.loc[inc_]
                        cards_i.append(f"""<div class="card {'best' if inc_ == melhor_i else ''}" style="--c:{COR_INC[inc_]}">
                          <div class="hd"><span class="t"><span class="dot"></span>{inc_}</span>
                          {'<span class="pill best">🏆 MELHOR</span>' if inc_ == melhor_i else ''}</div>
                          <div class="big {'neg' if x['MGCX'] < 0 else ''}">{fmt(x['MGCX'])}<small>MGCX R$/t</small></div>
                          <div class="sub">{fmt(x['Volume (t)'], 0)} t · preço {moeda(x['Preço bruto'], 0)}/t ·
                          frete {fmt(x['Frete+log.'], 0)}/t · FOB líq. {fmt(x['FOB líquido'], 0)}</div></div>""")
                    st.markdown(H(f'<div style="display:grid;gap:10px">{"".join(cards_i)}</div>'), unsafe_allow_html=True)
                    if len(por_inc) > 1:
                        a_, b2_ = por_inc.iloc[0], por_inc.iloc[1]
                        st.markdown(f"**{a_['Incoterm']}** rende **{fmt(a_['MGCX'] - b2_['MGCX'])} R\\$/t a mais** "
                                    f"que {b2_['Incoterm']} nesse destino.")
                with ci2:
                    lin_i = []
                    for nome_g, dfg, col in (("Fábrica", por_fi, "Fábrica"), ("Formato", por_fmi, "_formato")):
                        if dfg.empty:
                            continue
                        for chave_g, g in dfg.groupby(col, sort=False):
                            gi = g.set_index("Incoterm")
                            v_cif = gi.at["CIF", "MGCX"] if "CIF" in gi.index else None
                            v_fob = gi.at["FOB", "MGCX"] if "FOB" in gi.index else None
                            val = {k: v for k, v in (("CIF", v_cif), ("FOB", v_fob)) if v is not None}
                            venc = max(val, key=val.get) if val else "—"
                            lin_i.append((nome_g, chave_g.title() if nome_g == "Fábrica" else chave_g, v_cif, v_fob, venc,
                                          gi["Volume (t)"].sum()))
                    corpo_i = ""
                    grupo_ant = None
                    for nome_g, item, v_cif, v_fob, venc, vol_ in lin_i:
                        if nome_g != grupo_ant:
                            corpo_i += (f'<tr class="tot"><td>{"Por fábrica" if nome_g == "Fábrica" else "Por formato"}'
                                        f'</td><td></td><td></td><td></td></tr>')
                            grupo_ant = nome_g
                        tds = ""
                        for inc_, v in (("CIF", v_cif), ("FOB", v_fob)):
                            est = (f"background:rgba(47,191,113,.16);color:{VERDE};font-weight:800" if inc_ == venc else "")
                            tds += f'<td class="{"ne" if v is None else ("neg" if v < 0 else "")}" style="{est}">{fmt(v)}</td>'
                        corpo_i += f'<tr><td>{html.escape(str(item))}</td>{tds}<td style="text-align:center;font-weight:700">{venc}</td></tr>'
                    st.markdown(H(f'<table class="dre"><tr><th>MGCX R$/t</th><th><span style="color:{COR_INC["CIF"]}">●</span> CIF</th>'
                                  f'<th><span style="color:{COR_INC["FOB"]}">●</span> FOB</th><th>Melhor</th></tr>{corpo_i}</table>'),
                                unsafe_allow_html=True)

            with st.expander("Ver tabela completa (fábrica × formato)"):
                tb = por_ff.rename(columns={"_formato": "Formato"})
                st.dataframe(tb, hide_index=True, width="stretch",
                             column_config={c: st.column_config.NumberColumn(format="%.1f") for c in
                                            ["Volume (t)", "Preço bruto", "FOB líquido", "Frete+log.", "MCU", "MGS",
                                             "Capex", "MGCX"]} | {
                                 "Receita (R$)": st.column_config.NumberColumn(format="R$ %.0f"),
                                 "MGCX total (R$)": st.column_config.NumberColumn(format="R$ %.0f")})
            st.caption("Valores reais da base de resultados (média ponderada pelo volume no período escolhido na barra "
                       "lateral). Ao comparar 'Todos os cimentos', o mix de produtos (granel × saco) influencia o "
                       "resultado — escolha um cimento para uma comparação mais justa.")


@st.cache_data(show_spinner="Comparando todas as fábricas e CDs...", max_entries=30)
def calc_rotas(chave, _b, _r, _e):
    return rotas(_b, _e, _r)


def aplicar_rota(fab, prod, cd_):
    st.session_state["w_fab"] = fab
    if prod:
        st.session_state["w_prod"] = prod
    st.session_state["w_cd"] = cd_


with tab0:
    chave = (caminho, os.path.getmtime(caminho), cam_r if r is not None else None, cidade, canal, produto_nome,
             produto, incoterm, preco, sacos, zona, chapa, frac, tuple(meses or []), e.completar, e.chapa_valor)
    if not preco:
        st.info("Informe o preço negociado na barra lateral para comparar fábricas, CDs e distribuidores.", icon="💡")
        dr = None
    else:
        dr = calc_rotas(chave, b, r, e)
    if dr is not None and dr.empty:
        st.info("Não encontrei o produto equivalente nas outras fábricas.")
    elif dr is not None:
        ROT = {"Fábrica": "Fábrica direto", "CD": "Via CD", "DT": "Distribuidor"}
        so_2026 = st.toggle("Considerar só rotas com histórico de 2026 (mais confiável)", value=r is not None,
                            disabled=r is None,
                            help="Desligado inclui rotas calculadas só com as bases de jun/23 (menos confiáveis).")
        dr = dr.copy()
        dr["Confiança"] = dr.apply(lambda x: "Alta" if str(x["Fonte"]).startswith("base resultados") and x["Estimadas"] == 0
                                   else "Média" if str(x["Fonte"]).startswith("base resultados") else "Baixa", axis=1)
        ok = dr[dr["MGCX"].notna()].copy()
        if so_2026:
            ok = ok[ok["Confiança"] != "Baixa"]
        sem = dr[dr["MGCX"].isna()]
        if ok.empty:
            st.warning("Nenhuma rota com dados suficientes" + (" com histórico de 2026 — desligue a opção acima para "
                                                                "ver as rotas com bases jun/23." if so_2026 else "."))
        else:
            melhor_de = {f: ok[ok["Formato"] == f].iloc[0] if (ok["Formato"] == f).any() else None for f in ROT}
            top = ok.iloc[0]
            atual = {f"{fabrica.title()} → direto", f"{fabrica.title()} → distribuidor"} | (
                {f"{fabrica.title()} → {cd}"} if cd else set())
            mg_atual = ok.loc[ok["Rota"].isin(atual), "MGCX"].max()
            ganho = None if pd.isna(mg_atual) else top["MGCX"] - mg_atual

            # ---- veredito
            cor_top = COR[top["Formato"]]
            st.markdown(H(f"""<div class="route">
              <div><div class="k">Melhor opção para {html.escape(cidade.replace('_', ' · '))} · {html.escape(canal)}</div>
                <div class="v big">{html.escape(top['Rota'])}<span class="tag" style="background:{cor_top}22;
                color:{cor_top}">{ROT[top['Formato']].upper()}</span></div>
                <div class="s">{html.escape(str(top['Produto']))} · confiança {top['Confiança'].lower()}</div></div>
              <div><div class="k">Margem caixa</div><div class="v" style="color:{VERDE if top['MGCX'] >= 0 else VERMELHO}">
                {fmt(top['MGCX'])} <span style="font-size:.8rem;color:{TXT2}">R$/t</span></div>
                <div class="s">{moeda(top['MGCX'] * vol_t)} no volume</div></div>
              <div><div class="k">vs. sua rota atual</div><div class="v {'up' if ganho and ganho > 0 else ''}">
                {'—' if ganho is None else ('+' if ganho >= 0 else '') + fmt(ganho)} <span style="font-size:.8rem;color:{TXT2}">R$/t</span></div>
                <div class="s">{'' if ganho is None else moeda(ganho * vol_t) + ' no volume'}</div></div>
              <div><div class="k">Rotas comparadas</div><div class="v">{len(ok)} <span style="font-size:.8rem;color:{TXT2}">
                de {len(dr)}</span></div><div class="s">{len(dr) - len(ok)} sem dados ou fora do filtro</div></div></div>"""),
                        unsafe_allow_html=True)

            # ---- três comparações
            def duelo(titulo, a, b_, rot_a, rot_b, frase):
                if a is None and b_ is None:
                    corpo = f'<div class="s" style="margin-top:10px">Sem dados para comparar.</div>'
                    return f'<div class="panel"><div class="k" style="font-size:.7rem;color:{TXT3};letter-spacing:.1em">{titulo}</div>{corpo}</div>'
                itens = [(rot_a, a), (rot_b, b_)]
                vals = [abs(float(x["MGCX"])) for _, x in itens if x is not None] or [1]
                mx_ = max(vals) or 1
                venc = max([x for x in itens if x[1] is not None], key=lambda t: t[1]["MGCX"])
                linhas_ = ""
                for rot, x in itens:
                    if x is None:
                        linhas_ += (f'<div style="margin-top:12px"><div style="display:flex;justify-content:space-between;'
                                    f'font-size:.8rem;color:{TXT3}"><span>{rot}</span><span>sem dados</span></div></div>')
                        continue
                    w = abs(float(x["MGCX"])) / mx_ * 100
                    c_ = COR[x["Formato"]]
                    win = x is venc[1]
                    linhas_ += (f'<div style="margin-top:12px"><div style="display:flex;justify-content:space-between;'
                                f'font-size:.8rem;color:{TXT2}"><span>{rot}: <b style="color:{TXT}">{html.escape(x["Rota"])}'
                                f'</b>{" 🏆" if win else ""}</span><span style="color:{VERMELHO if x["MGCX"] < 0 else TXT};'
                                f'font-weight:700">{fmt(x["MGCX"])}</span></div><div class="bar" style="margin:6px 0 0">'
                                f'<i style="width:{w:.0f}%;background:{VERMELHO if x["MGCX"] < 0 else c_}"></i></div></div>')
                return (f'<div class="panel"><div style="font-size:.7rem;font-weight:700;color:{TXT3};letter-spacing:.1em;'
                        f'text-transform:uppercase">{titulo}</div>{linhas_}<div class="s" style="margin-top:12px;'
                        f'font-size:.8rem;color:{TXT2}">{frase}</div></div>')

            def dif_txt(a, b_):
                if a is None or b_ is None:
                    return ""
                d_ = a["MGCX"] - b_["MGCX"]
                return f"{fmt(abs(d_))} R$/t ({moeda(abs(d_) * vol_t)} no volume)"

            dirs = ok[ok["Formato"] == "Fábrica"]
            f1 = dirs.iloc[0] if len(dirs) > 0 else None
            f2 = dirs.iloc[1] if len(dirs) > 1 else None
            c1_ = melhor_de["CD"]
            d1 = melhor_de["DT"]
            fc = max([x for x in (f1, c1_) if x is not None], key=lambda x: x["MGCX"]) if (f1 is not None or c1_ is not None) else None
            t1 = (f"Direto, <b>{html.escape(f1['Rota'].split(' →')[0])}</b> é a melhor fábrica"
                  + (f", {dif_txt(f1, f2)} acima de {html.escape(f2['Rota'].split(' →')[0])}." if f2 is not None else ".")
                  ) if f1 is not None else "Nenhuma fábrica com dados para venda direta."
            if f1 is not None and c1_ is not None:
                t2 = (f"<b>{'Atender via CD' if c1_['MGCX'] > f1['MGCX'] else 'Vender direto da fábrica'}</b> rende "
                      f"{dif_txt(c1_, f1)} a mais.")
            else:
                t2 = "Falta dado de um dos lados para comparar."
            if d1 is not None and fc is not None:
                t3 = (f"<b>{'O distribuidor' if d1['MGCX'] > fc['MGCX'] else ('O CD' if fc['Formato'] == 'CD' else 'A venda direta')}"
                      f"</b> é melhor, por {dif_txt(d1, fc)}.")
            else:
                t3 = "Falta dado de um dos lados para comparar."
            st.markdown(H('<div class="sec"><h3>Comparações</h3><span>margem caixa em R$/t · mesma venda, mesmo preço'
                          '</span></div><div style="display:grid;grid-template-columns:repeat(3,1fr);gap:14px;'
                          'margin-bottom:16px">'
                          + duelo("1 · Fábrica × Fábrica", f1, f2, "1ª", "2ª", t1)
                          + duelo("2 · CD × Fábrica", c1_, f1, "Melhor CD", "Melhor direto", t2)
                          + duelo("3 · Distribuidor × Fábrica/CD", d1, fc, "Melhor distribuidor", "Melhor fábrica/CD", t3)
                          + '</div>'), unsafe_allow_html=True)

            # ---- matriz por fábrica
            linhas_m = []
            for fab_, g in dr.groupby("Fábrica", sort=False):
                g_ok = g[g.index.isin(ok.index)]
                dire = g_ok[g_ok["Formato"] == "Fábrica"]
                cds_ = g_ok[g_ok["Formato"] == "CD"]
                dts_ = g_ok[g_ok["Formato"] == "DT"]
                vals = {"Direto": dire.iloc[0]["MGCX"] if len(dire) else None,
                        "CD": cds_.iloc[0]["MGCX"] if len(cds_) else None,
                        "Distribuidor": dts_.iloc[0]["MGCX"] if len(dts_) else None}
                nome_cd = cds_.iloc[0]["Expedição"] if len(cds_) else ""
                validos = {k: v for k, v in vals.items() if v is not None}
                if not validos:
                    continue
                venc = max(validos, key=validos.get)
                linhas_m.append((max(validos.values()), fab_, vals, nome_cd, venc))
            linhas_m.sort(key=lambda x: -x[0])
            cab = (f'<tr><th>Fábrica</th><th><span style="color:{COR["Fábrica"]}">●</span> Direto</th>'
                   f'<th><span style="color:{COR["CD"]}">●</span> Melhor CD</th>'
                   f'<th><span style="color:{COR["DT"]}">●</span> Distribuidor</th><th>Melhor caminho</th></tr>')
            corpo = ""
            for _, fab_, vals, nome_cd, venc in linhas_m:
                tds = ""
                for k in ("Direto", "CD", "Distribuidor"):
                    v = vals[k]
                    sub = f'<span class="h">{html.escape(nome_cd)}</span>' if k == "CD" and nome_cd else ""
                    estilo = (f'background:rgba(47,191,113,.16);color:{VERDE};font-weight:800' if k == venc else "")
                    cls = "ne" if v is None else ("neg" if v < 0 else "")
                    tds += f'<td class="{cls}" style="{estilo}">{fmt(v)}{sub}</td>'
                corpo += (f'<tr><td style="color:{TXT};font-weight:600;padding-left:14px">{html.escape(fab_.title())}</td>'
                          f'{tds}<td style="text-align:center;font-weight:700">{venc}</td></tr>')
            st.markdown(H(f'<div class="sec"><h3>Por fábrica</h3><span>melhor resultado de cada caminho · verde = '
                          f'vencedor da linha</span></div><table class="dre">{cab}{corpo}</table>'), unsafe_allow_html=True)

            # ---- ranking + seleção
            ok["Atual"] = ok["Rota"].isin(atual)
            ok["Rótulo"] = ok["MGCX"].map(fmt)
            ok["Caminho"] = ok["Formato"].map(ROT)
            topn = ok.head(15)
            ordem = list(topn["Rota"])
            base_ch = alt.Chart(topn).encode(
                y=alt.Y("Rota:N", sort=ordem, title=None, axis=alt.Axis(labelFontSize=12, labelColor=TXT, labelLimit=280)))
            barras = base_ch.mark_bar(cornerRadiusEnd=5, height=20).encode(
                x=alt.X("MGCX:Q", title="Margem caixa (R$/t)", axis=alt.Axis(domain=False, tickCount=6)),
                color=alt.Color("Caminho:N", scale=alt.Scale(domain=list(ROT.values()),
                                                             range=[COR["Fábrica"], COR["CD"], COR["DT"]]),
                                legend=alt.Legend(orient="top", title=None)),
                stroke=alt.condition("datum.Atual", alt.value("#ffffff"), alt.value(None)),
                strokeWidth=alt.condition("datum.Atual", alt.value(2), alt.value(0)),
                tooltip=["Rota:N", "Caminho:N", alt.Tooltip("Rótulo:N", title="MGCX R$/t"),
                         alt.Tooltip("Frete vendas:Q", format=",.2f"), alt.Tooltip("Frete transf.:Q", format=",.2f"),
                         alt.Tooltip("FOB Líquido:Q", format=",.2f"), "Fonte:N", "Confiança:N"])
            txt = base_ch.mark_text(align="left", dx=6, fontSize=11, fontWeight=700, color=TXT).encode(
                x=alt.X("MGCX:Q"), text="Rótulo:N")
            zero = alt.Chart(pd.DataFrame({"x": [0]})).mark_rule(color=TXT3).encode(x="x:Q")
            cg, ct = st.columns([3, 2], gap="large")
            with cg:
                st.markdown('<div class="sec"><h3>Ranking de rotas</h3><span>contorno branco = rota atual</span></div>',
                            unsafe_allow_html=True)
                st.altair_chart(tema((barras + txt + zero).properties(height=max(220, 34 * len(topn)))), width="stretch")
            with ct:
                st.markdown('<div class="sec"><h3>Detalhe</h3><span>clique numa linha para simular a rota</span></div>',
                            unsafe_allow_html=True)
                tab_r = ok[["Rota", "Caminho", "MGCX", "Frete vendas", "Frete transf.", "Confiança", "Fonte"]] \
                    .reset_index(drop=True)
                sel = st.dataframe(
                    tab_r, hide_index=True, width="stretch", height=min(560, 38 + 35 * len(tab_r)),
                    on_select="rerun", selection_mode="single-row", key="sel_rota",
                    column_config={
                        "MGCX": st.column_config.ProgressColumn("MGCX R$/t", format="%.1f",
                                                                min_value=float(min(0, tab_r["MGCX"].min())),
                                                                max_value=float(max(1, tab_r["MGCX"].max()))),
                        "Frete vendas": st.column_config.NumberColumn(format="%.1f"),
                        "Frete transf.": st.column_config.NumberColumn(format="%.1f"),
                    })
                linhas_sel = sel.selection.rows if sel and sel.selection else []
                if linhas_sel:
                    lr = ok.reset_index(drop=True).iloc[linhas_sel[0]]
                    st.button(f"▶  Simular {lr['Rota']}", type="primary", width="stretch",
                              on_click=aplicar_rota,
                              args=(lr["Fábrica"], lr["Produto"] if r is not None else None,
                                    lr["Expedição"] if lr["Formato"] == "CD" else ""))
        if len(sem):
            with st.expander(f"Rotas sem dados suficientes ({len(sem)})"):
                st.dataframe(sem[["Rota", "Formato", "Pendência", "Fonte"]], hide_index=True, width="stretch")
                st.caption("Normalmente falta o frete de vendas daquela expedição para esta cidade.")
        st.caption("Mesmo preço negociado e produto equivalente (mesma família e peso) em todas as rotas. "
                   "Distribuidor = venda para o distribuidor/AP da cidade. Valide a rota escolhida com Logística e Pricing.")

with tab1:
    c1, c2, _ = st.columns([1, 1, 4])
    por_saco = c1.toggle("Mostrar em R$/saco", value=False, disabled=not ref["ensacado"])
    mostrar_h = c2.toggle("Mostrar histórico", value=False)
    st.markdown(dre_html(por_saco, mostrar_h), unsafe_allow_html=True)


with tab2:
    g1, g2 = st.columns([2, 3], gap="large")
    with g1:
        st.markdown('<div class="sec"><h3>Margem caixa por formato</h3><span>R$/t</span></div>', unsafe_allow_html=True)
        dfm = pd.DataFrame({"Formato": CANAIS, "MGCX": [rs.at["MGCX", c] for c in CANAIS]})
        dfm["MGCX"] = pd.to_numeric(dfm["MGCX"], errors="coerce")
        dfm["Rótulo"] = dfm["MGCX"].map(fmt)
        dfm["Formato★"] = dfm["Formato"].map(lambda c: c + (" ★" if c == melhor else ""))
        dfm["y"] = dfm["MGCX"].fillna(0)
        bars = alt.Chart(dfm).mark_bar(cornerRadiusEnd=4, size=42).encode(
            x=alt.X("Formato:N", sort=CANAIS, title=None, axis=alt.Axis(labelAngle=0, labelFontSize=12, labelColor=TXT)),
            y=alt.Y("y:Q", title="R$/t", axis=alt.Axis(domain=False, tickCount=5)),
            color=alt.Color("Formato:N", scale=alt.Scale(domain=CANAIS, range=[COR[c] for c in CANAIS]), legend=None),
            tooltip=[alt.Tooltip("Formato:N"), alt.Tooltip("Rótulo:N", title="MGCX R$/t")])
        lp = alt.Chart(dfm[(dfm["y"] >= 0) & dfm["MGCX"].notna()]).mark_text(dy=-9, fontSize=12, fontWeight=700, color=TXT).encode(
            x=alt.X("Formato:N", sort=CANAIS), y="y:Q", text="Rótulo:N")
        ln = alt.Chart(dfm[dfm["y"] < 0]).mark_text(dy=11, fontSize=12, fontWeight=700, color=VERMELHO).encode(
            x=alt.X("Formato:N", sort=CANAIS), y="y:Q", text="Rótulo:N")
        zero = alt.Chart(pd.DataFrame({"y": [0]})).mark_rule(color=TXT3).encode(y="y:Q")
        st.altair_chart(tema((bars + lp + ln + zero).properties(height=360)), width="stretch")
    with g2:
        st.markdown('<div class="sec"><h3>Cascata de margem</h3><span>do preço bruto à margem caixa · R$/t</span></div>',
                    unsafe_allow_html=True)
        opcoes = [c for c in CANAIS if not vazio(rs.at["FOB Líquido", c])] or CANAIS
        canal_g = st.segmented_control("Formato", opcoes, default=melhor if melhor in opcoes else opcoes[0],
                                       label_visibility="collapsed") or opcoes[0]
        d = res["canais"][canal_g].set_index("Linha")["Simulação R$/t"]
        g = lambda n: 0.0 if vazio(d.get(n)) else float(d.get(n))  # noqa: E731
        passos = [("Preço Bruto", None, "Preço Bruto"),
                  ("Impostos e desc.", -(g("Impostos") + g("Descontos") + g("Devoluções") + g("Rapel")), None),
                  ("Receita Líq.", None, "Receita Líquida"),
                  ("Fretes e log.", -sum(g(n) for n in ("Frete de Vendas", "Frete Devolução", "Frete Transferência",
                                                       "Chapa", "Take or Pay", "Outros Custos Log.")), None),
                  ("FOB Líq.", None, "Preço FOB"),
                  ("Custo var.", -g("Custo Variável"), None),
                  ("MCU", None, "Margem de Contribuição"),
                  ("Custo fixo", -g("Custo Fixo"), None),
                  ("Custo servir", -g("Custo de Servir"), None),
                  ("MGS", None, "Margem de Servir"),
                  ("Capex", -g("Capex"), None),
                  ("MGCX", None, "Margem Caixa")]
        rows, nivel = [], 0.0
        for nome, delta, tot in passos:
            if tot:
                nivel = g(tot)
                rows.append(dict(Etapa=nome, ini=0, fim=nivel, Tipo="Resultado", Valor=nivel))
            else:
                rows.append(dict(Etapa=nome, ini=nivel, fim=nivel + delta, Tipo="Dedução", Valor=delta))
                nivel += delta
        dfw = pd.DataFrame(rows)
        dfw["Rótulo"] = dfw["Valor"].map(fmt)
        ordem = list(dfw["Etapa"])
        wf = alt.Chart(dfw).mark_bar(size=34, cornerRadius=3).encode(
            x=alt.X("Etapa:N", sort=ordem, title=None, axis=alt.Axis(labelAngle=0, labelFontSize=11)),
            y=alt.Y("ini:Q", title="R$/t", axis=alt.Axis(domain=False, tickCount=5)), y2="fim:Q",
            color=alt.Color("Tipo:N", scale=alt.Scale(domain=["Resultado", "Dedução"], range=[COR.get(canal_g, AZUL), DEDUCAO]),
                            legend=alt.Legend(orient="top", title=None)),
            tooltip=["Etapa:N", alt.Tooltip("Rótulo:N", title="R$/t")])
        wt = alt.Chart(dfw[dfw["Tipo"] == "Resultado"]).mark_text(dy=-9, fontSize=11, fontWeight=700, color=TXT).encode(
            x=alt.X("Etapa:N", sort=ordem), y="fim:Q", text="Rótulo:N")
        st.altair_chart(tema((wf + wt).properties(height=320)), width="stretch")

with tab3:
    st.markdown('<div class="sec"><h3>Ajustar premissas</h3><span>edite a coluna Simulação (R$/t) para substituir '
                'um valor — equivale às células "ALTERAR" da planilha</span></div>', unsafe_allow_html=True)
    sub = st.tabs([f"{c}" for c in CANAIS])
    for canal_, tab in zip(CANAIS, sub):
        with tab:
            dfb = base["canais"][canal_]
            ed_in = dfb[dfb["Linha"].isin(EDITAVEIS)][["Linha", "Histórico R$/t", "Simulação R$/t", "Origem"]] \
                .reset_index(drop=True)
            ed_in["Histórico R$/t"] = ed_in["Histórico R$/t"].map(fmt)
            st.data_editor(ed_in, key=f"ed_{canal_}", hide_index=True, width="stretch", height=740,
                           disabled=["Linha", "Histórico R$/t", "Origem"],
                           column_config={"Linha": st.column_config.TextColumn(width="medium"),
                                          "Histórico R$/t": st.column_config.TextColumn(width="small"),
                                          "Simulação R$/t": st.column_config.NumberColumn(format="%.2f", width="small"),
                                          "Origem": st.column_config.TextColumn(width="large")})
            if canal_ in ajustes:
                st.caption("✏️ Ajustado manualmente: " + ", ".join(ajustes[canal_]))

with tabp:
    st.markdown('<div class="sec"><h3>Preço mínimo</h3><span>a partir de qual preço negociado cada formato '
                'atinge a margem caixa desejada</span></div>', unsafe_allow_html=True)
    if not preco:
        st.info("Informe o preço negociado para calcular.", icon="💡")
    else:
        cA, cB = st.columns([1, 3], gap="large")
        with cA:
            alvo = st.slider("Margem caixa desejada (R$/t)", -50, 200, 0, 5)
            pm = preco_minimo(b, e, r, float(alvo))
            for c_ in ("Fábrica", "CD"):
                p = pm.get(c_)
                if p is None:
                    st.markdown(f'<div class="panel" style="margin-bottom:10px;border-left:4px solid {COR[c_]}">'
                                f'<div class="kpi" style="padding:0"><div class="k">{c_}</div><div class="v">—</div>'
                                f'<div class="s">sem dados</div></div></div>', unsafe_allow_html=True)
                    continue
                folga = preco - p
                st.markdown(f'<div class="panel" style="margin-bottom:10px;border-left:4px solid {COR[c_]}">'
                            f'<div class="kpi" style="padding:0"><div class="k">{c_} · preço mínimo</div>'
                            f'<div class="v">{moeda(p, 2)} <span style="font-size:.8rem;color:{TXT2}">/{unid}</span></div>'
                            f'<div class="s"><span class="{"up" if folga >= 0 else "dn"}">{"+" if folga >= 0 else ""}'
                            f'{fmt(folga)}</span> vs. preço atual</div></div></div>', unsafe_allow_html=True)
        with cB:
            base1 = resumo(res).loc["MGCX"]
            from dataclasses import replace as _rep
            base2 = resumo(simular(b, _rep(e, preco_negociado=preco + 1), r)).loc["MGCX"]
            precos = [preco * (0.8 + 0.01 * i) for i in range(41)]
            linhas = []
            for c_ in ("Fábrica", "CD"):
                m1, m2 = base1.get(c_), base2.get(c_)
                if vazio(m1) or vazio(m2):
                    continue
                for p in precos:
                    linhas.append({"Formato": c_, "Preço": p, "MGCX": m1 + (m2 - m1) * (p - preco)})
            if linhas:
                dl = pd.DataFrame(linhas)
                hover = alt.selection_point(fields=["Preço"], nearest=True, on="pointerover", empty=False)
                lin_ = alt.Chart(dl).mark_line(strokeWidth=3).encode(
                    x=alt.X("Preço:Q", title=f"Preço negociado (R$/{unid})", scale=alt.Scale(zero=False)),
                    y=alt.Y("MGCX:Q", title="Margem caixa (R$/t)"),
                    color=alt.Color("Formato:N", scale=alt.Scale(domain=["Fábrica", "CD"],
                                                                 range=[COR["Fábrica"], COR["CD"]]),
                                    legend=alt.Legend(orient="top", title=None)))
                pts = lin_.mark_point(size=80, filled=True).encode(
                    opacity=alt.condition(hover, alt.value(1), alt.value(0)),
                    tooltip=[alt.Tooltip("Formato:N"), alt.Tooltip("Preço:Q", format=",.2f"),
                             alt.Tooltip("MGCX:Q", format=",.2f")]).add_params(hover)
                regra_alvo = alt.Chart(pd.DataFrame({"y": [alvo]})).mark_rule(color=AMBAR, strokeDash=[6, 4]).encode(y="y:Q")
                regra_preco = alt.Chart(pd.DataFrame({"x": [preco]})).mark_rule(color=TXT3, strokeDash=[2, 3]).encode(x="x:Q")
                st.altair_chart(tema((lin_ + pts + regra_alvo + regra_preco).properties(height=380)), width="stretch")
                st.caption("Linha tracejada amarela = margem desejada · linha cinza = preço atual. "
                           "Passe o mouse sobre as linhas para ver os valores.")

with tab4:
    st.markdown('<div class="sec"><h3>Exportar simulação</h3><span>Excel formatado com resumo, DRE comparativa, '
                'ranking de rotas e premissas</span></div>', unsafe_allow_html=True)
    rot_x = None
    if preco:
        try:
            rot_x = calc_rotas(chave, b, r, e)
        except Exception:
            rot_x = None
    arq = gerar_excel(e, res, melhor, rot_x, ajustes)
    st.download_button("⬇  Baixar simulação (.xlsx)", arq,
                       f"simulacao_{norm(fabrica).lower()}_{norm(cidade).lower()}.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")
    st.caption("Abas: Resumo · DRE comparativa · Rotas (fábrica × CD) · Premissas e origem")
