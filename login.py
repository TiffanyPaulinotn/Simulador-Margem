"""
Login do simulador: só entra quem tem e-mail corporativo (@intercement.com).

Usa o login da Microsoft (a mesma conta do Outlook/Teams) via st.login do Streamlit.
A configuração fica em .streamlit/secrets.toml (veja secrets.toml.exemplo). Sem essa configuração:
- rodando no próprio computador (localhost), o simulador abre normalmente, com um aviso;
- acessado por outro endereço (internet/rede), o acesso é bloqueado.
"""
from __future__ import annotations

VERSAO_APP = "2026-11-11.2"

import html

import streamlit as st

DOMINIOS_PADRAO = ["intercement.com"]


def _cfg():
    try:
        auth = st.secrets.get("auth")
        acesso = st.secrets.get("acesso", {})
    except Exception:          # sem secrets.toml
        return None, {}
    return auth, acesso


def _usuario():
    u = getattr(st, "user", None) or getattr(st, "experimental_user", None)
    return u


def _email(u) -> str:
    for campo in ("email", "preferred_username", "upn", "unique_name"):
        try:
            v = u.get(campo) if hasattr(u, "get") else getattr(u, campo, None)
        except Exception:
            v = None
        if v and "@" in str(v):
            return str(v).strip().lower()
    return ""


def _host() -> str:
    try:
        return str(st.context.headers.get("Host", "") or "").lower()
    except Exception:
        return ""


def _local() -> bool:
    h = _host()
    return h == "" or h.startswith(("localhost", "127.0.0.1", "[::1]"))


def _tela(titulo: str, texto: str, botao: str | None = None, erro: bool = False) -> bool:
    st.markdown(f"""
    <div style="max-width:460px;margin:9vh auto 0;background:#121821;border:1px solid #243041;border-radius:18px;
         padding:34px 34px 26px;text-align:center;box-shadow:0 20px 60px rgba(0,0,0,.35)">
      <div style="width:56px;height:56px;border-radius:14px;margin:0 auto 14px;display:flex;align-items:center;
           justify-content:center;font-weight:800;color:#fff;font-size:1.2rem;
           background:linear-gradient(135deg,#4c8dff,#7a5cff)">SM</div>
      <div style="font-size:.72rem;font-weight:700;letter-spacing:.14em;color:#4c8dff">SIMULADOR DE MARGENS · CANAIS BR</div>
      <div style="font-size:1.45rem;font-weight:800;color:#e8edf3;margin:6px 0 10px">{titulo}</div>
      <div style="font-size:.9rem;color:{'#ff5d5d' if erro else '#9aa7b6'};line-height:1.5">{texto}</div>
    </div>""", unsafe_allow_html=True)
    if botao:
        _, c, _ = st.columns([1, 1, 1])
        return c.button(botao, type="primary", width="stretch")
    return False


def exigir_login() -> str:
    """Bloqueia o app até a pessoa entrar com e-mail corporativo. Retorna o e-mail (ou '' no modo local)."""
    auth, acesso = _cfg()
    dominios = [d.lower().lstrip("@") for d in (acesso.get("dominios") or DOMINIOS_PADRAO)]
    rotulo_dom = " ou ".join("@" + d for d in dominios)

    if not auth:
        if _local():
            st.sidebar.caption("🔓 Modo local (sem login). Na internet, o acesso exige login — veja o README.")
            return ""
        _tela("Acesso bloqueado", "O login corporativo ainda não foi configurado neste servidor. "
                                  "Peça para a TI configurar o arquivo .streamlit/secrets.toml.", erro=True)
        st.stop()

    provedor = acesso.get("provedor") or next((k for k in auth if isinstance(auth.get(k), dict)), None)
    u = _usuario()
    if not getattr(u, "is_logged_in", False):
        if _tela("Entrar", f"Acesso restrito a colaboradores com e-mail {html.escape(rotulo_dom)}.<br>"
                           "Use a mesma conta do Outlook/Teams.", "Entrar com a conta InterCement"):
            st.login(provedor) if provedor else st.login()
        st.stop()

    email = _email(u)
    if not any(email.endswith("@" + d) for d in dominios):
        if _tela("Acesso não permitido",
                 f"A conta <b>{html.escape(email or 'sem e-mail')}</b> não é {html.escape(rotulo_dom)}.<br>"
                 "Saia e entre com a sua conta corporativa.", "Sair e trocar de conta", erro=True):
            st.logout()
        st.stop()

    with st.sidebar:
        c1, c2 = st.columns([3, 1])
        c1.caption(f"👤 {email}")
        if c2.button("Sair", key="btn_sair"):
            st.logout()
    return email
