from __future__ import annotations

import base64
import time

import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt

# =========================================================
# CONFIG / CSS
# =========================================================
st.set_page_config(page_title="Login", page_icon="🔐", layout="wide")

st.markdown(
    """
    <style>
        #MainMenu {visibility: hidden;}
        .stDeployButton {display:none;}
        footer {visibility: hidden;}
        #stDecoration {display:none;}

        .stAppHeader { background-color: transparent !important; box-shadow: none !important; }
        .block-container { padding-top: 0.2rem; padding-bottom: 1.5rem; }

        .stTextInput input, .stSelectbox div[data-baseweb="select"] > div {
            border: 2px solid rgba(0,0,0,0.15) !important;
            border-radius: 10px !important;
        }

        .stButton button {
            width: 100% !important;
            background-color: #6B7280 !important;
            color: white !important;
            font-size: 16px !important;
            padding: 10px 12px !important;
            border-radius: 10px !important;
            border: none !important;
            font-weight: 600 !important;
        }

        .stButton button:hover { filter: brightness(0.95); }

        div[data-testid="stVerticalBlockBorderWrapper"] {
            border: 1px solid rgba(0,0,0,0.08);
            border-radius: 14px;
            padding: 12px;
            box-shadow: 0 8px 22px rgba(0,0,0,0.06);
            background: white;
        }

        .status-topo {
            background: #f8fafc;
            border: 1px solid #cbd5e1;
            border-left: 5px solid #2563eb;
            border-radius: 10px;
            padding: 10px 14px;
            min-height: 66px;
            display: flex;
            flex-direction: column;
            justify-content: center;
        }

        .status-topo-ok {
            background: #ecfdf5;
            border: 1px solid #bbf7d0;
            border-left: 5px solid #16a34a;
            border-radius: 10px;
            padding: 10px 14px;
            min-height: 66px;
            display: flex;
            flex-direction: column;
            justify-content: center;
        }

        .status-topo-titulo {
            font-size: 0.95rem;
            font-weight: 700;
            color: #1e293b;
            margin-bottom: 4px;
        }

        .status-topo-ok .status-topo-titulo {
            color: #166534;
        }

        .status-topo-texto {
            font-size: 0.85rem;
            color: #475569;
            line-height: 1.35;
        }

        .status-topo-ok .status-topo-texto {
            color: #166534;
        }
    </style>
    """,
    unsafe_allow_html=True,
)


def set_background(image_path: str = "Auditor.png", alpha: float = 0.90) -> None:
    try:
        with open(image_path, "rb") as f:
            b64 = base64.b64encode(f.read()).decode()

        st.markdown(
            f"""
            <style>
            .stApp {{
                background: linear-gradient(rgba(255,255,255,{alpha}), rgba(255,255,255,{alpha})),
                            url('data:image/png;base64,{b64}') no-repeat center center fixed;
                background-size: cover;
            }}
            </style>
            """,
            unsafe_allow_html=True,
        )
    except FileNotFoundError:
        st.warning(f"Imagem de fundo '{image_path}' não encontrada. O app seguirá sem fundo.")


set_background("Auditor.png", alpha=0.90)

# =========================================================
# SESSION STATE
# =========================================================
st.session_state.setdefault("authenticated", False)
st.session_state.setdefault("username", "")
st.session_state.setdefault("lk_grupo", "")
st.session_state.setdefault("stage", "login")  # login | welcome | app

# status do carregamento no topo
st.session_state.setdefault("loading_message", "")
st.session_state.setdefault("loading_elapsed", 0.0)
st.session_state.setdefault("loading_done", False)

# =========================================================
# AUTH (CSV)
# =========================================================
@st.cache_data(ttl=300)
def load_users_csv(path: str = "data/Lk-grupo.csv") -> pd.DataFrame:
    df = pd.read_csv(path, sep=";")
    df.columns = df.columns.str.strip()

    required = {"LK_GRUPO", "username", "password"}
    if not required.issubset(df.columns):
        raise ValueError(f"CSV precisa ter as colunas: {', '.join(sorted(required))}")

    df["LK_GRUPO"] = df["LK_GRUPO"].astype(str).str.strip()
    df["username"] = df["username"].astype(str).str.strip()
    df["password"] = df["password"].astype(str).str.strip()
    return df


def authenticate(df: pd.DataFrame, lk_grupo: str, username: str, password: str) -> bool:
    lk_grupo = str(lk_grupo).strip()
    username = str(username).strip()
    password = str(password).strip()

    mask = (
        (df["LK_GRUPO"] == lk_grupo)
        & (df["username"] == username)
        & (df["password"] == password)
    )
    return bool(mask.any())


def logout() -> None:
    st.session_state["authenticated"] = False
    st.session_state["username"] = ""
    st.session_state["lk_grupo"] = ""
    st.session_state["stage"] = "login"
    st.session_state["loading_message"] = ""
    st.session_state["loading_elapsed"] = 0.0
    st.session_state["loading_done"] = False
    st.rerun()


# =========================================================
# LOGIN PAGE
# =========================================================
def login_page() -> None:
    col1, col2, col3 = st.columns([1, 1, 1])

    with col2:
        try:
            st.image(plt.imread("Controller.png"))
        except Exception:
            st.info("Logo 'Controller.png' não encontrada.")

        st.markdown("<p style='margin-top: -30px; margin-bottom: 10px;'></p>", unsafe_allow_html=True)

        try:
            df_users = load_users_csv("data/Lk-grupo.csv")
        except Exception as e:
            st.error(f"Erro ao carregar data/Lk-grupo.csv: {e}")
            st.stop()

        empresas = sorted(df_users["LK_GRUPO"].dropna().unique().tolist())
        empresa_default = "LK-GRUPO Adex"
        default_index = empresas.index(empresa_default) if empresa_default in empresas else 0

        with st.container(border=True):
            st.subheader("Acesso ao Programa")

            lk_grupo = st.selectbox("Empresa:", options=empresas, index=default_index)
            username = st.text_input("Usuário:", max_chars=50).strip()
            password = st.text_input("Senha:", type="password", max_chars=50).strip()

            if st.button("ACESSAR", use_container_width=True):
                if authenticate(df_users, lk_grupo, username, password):
                    st.session_state["authenticated"] = True
                    st.session_state["username"] = username
                    st.session_state["lk_grupo"] = lk_grupo
                    st.session_state["stage"] = "welcome"
                    st.session_state["loading_message"] = ""
                    st.session_state["loading_elapsed"] = 0.0
                    st.session_state["loading_done"] = False
                    st.rerun()
                else:
                    st.error("Login falhou. Verifique Empresa/Usuário/Senha.")


# =========================================================
# HEADER DO APP LOGADO
# =========================================================
def header_bar():
    c1, c2, c3 = st.columns([8, 5, 2], vertical_alignment="center")

    with c1:
        st.title("Carga | Painel de Indicadores")
        st.caption(f"Usuário: {st.session_state.username} — Empresa: {st.session_state.lk_grupo}")

    with c2:
        status_placeholder = st.empty()

    with c3:
        if st.button("🚪 Sair", use_container_width=True, key="btn_logout_top"):
            logout()

    st.markdown("---")
    return status_placeholder


# =========================================================
# FLUXO PRINCIPAL
# =========================================================
if not st.session_state["authenticated"]:
    login_page()

elif st.session_state["stage"] == "welcome":

    html_msg = f"""
<div style="margin-top:80px;
padding:28px 32px;
border-radius:18px;
background:rgba(255,255,255,0.84);
border:1px solid rgba(0,0,0,0.08);
box-shadow:0 8px 24px rgba(0,0,0,0.08);
max-width:760px;">

<div style="font-size:32px;font-weight:700;color:#1F2A44;margin-bottom:16px;">
Olá {st.session_state["username"]},
</div>

<div style="font-size:20px;color:#374151;margin-bottom:12px;">
Aguarde enquanto executamos o <b>Programa</b>.
</div>

<div style="font-size:18px;color:#4B5563;line-height:1.7;">
Durante o processo de execução estamos realizando a releitura das informações
dos setores da empresa <b>{st.session_state["lk_grupo"]}</b>.
</div>

</div>
"""

    st.markdown(html_msg, unsafe_allow_html=True)

    with st.spinner("Preparando painel..."):
        time.sleep(2)

    st.session_state["stage"] = "app"
    st.rerun()

elif st.session_state["stage"] == "app":
    status_placeholder = header_bar()

    import runpy
    import traceback

    try:
        runpy.run_path("app_laboratorio.py", run_name="__main__")
    except Exception as e:
        st.error("Erro ao carregar o programa app_laboratorio.py")
        st.exception(e)
        st.code(traceback.format_exc())