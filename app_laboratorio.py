from __future__ import annotations

import os
import re
import time
import base64
from pathlib import Path
from datetime import datetime
from io import BytesIO
from st_aggrid import JsCode

import pandas as pd
import streamlit as st
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode, JsCode
from services.laboratorio_service import carregar_ofs_laboratorio


st.set_page_config(
    page_title="Painel Laboratório",
    page_icon="🧪",
    layout="wide",
    initial_sidebar_state="collapsed",
)

ARQUIVO_BANCO = "banco_laboratorio.txt"
ARQUIVO_LOGO_SIDEBAR = "Controller.png"
ARQUIVO_LOGO_DIREITA = "Logo_ADXW.bmp"
ARQUIVO_FUNDO = "Auditor.png"


def inicializar_estado_app():
    defaults = {
        "mensagem_salvo": "",
        "mostrar_loading": True,
        "forcar_atualizacao": True,
        "df_base": None,
        "df_grid_editado": None,
        "df_sd_auditoria": None,
        "grid_key": 0,
    }

    filtros_defaults = {
        "status_of": ["A"],
        "cliente": [],
        "codigo_produto": [],
        "codigo_original": [],
        "codigo_grupo": [],
        "sequencia_atual": [],
        "operador_atual": [],
        "nw_data_inicio": None,
        "nw_data_fim": None,
        "prev_entrega_inicio": None,
        "prev_entrega_fim": None,
        "abertura_inicio": None,
        "abertura_fim": None,
    }

    for chave, valor in defaults.items():
        if chave not in st.session_state:
            st.session_state[chave] = valor

    if "filtros_aplicados" not in st.session_state or not isinstance(st.session_state["filtros_aplicados"], dict):
        st.session_state["filtros_aplicados"] = filtros_defaults.copy()
    else:
        for chave, valor in filtros_defaults.items():
            if chave not in st.session_state["filtros_aplicados"]:
                st.session_state["filtros_aplicados"][chave] = valor


def imagem_para_base64(caminho_arquivo: str) -> str:
    caminho = Path(caminho_arquivo)
    if not caminho.exists():
        return ""

    with open(caminho, "rb") as f:
        return base64.b64encode(f.read()).decode("utf-8")


def aplicar_estilo_visual():
    fundo_base64 = imagem_para_base64(ARQUIVO_FUNDO)

    background_css = ""
    if fundo_base64:
        background_css = f"""
        .stApp {{
            background-image:
                linear-gradient(rgba(255,255,255,0.92), rgba(255,255,255,0.92)),
                url("data:image/png;base64,...");

            background-size: 60%;
            background-position: right center;
            background-repeat: no-repeat;
        }}
        """

    st.markdown(
        f"""
        <style>
        {background_css}

        /* Sidebar */
        [data-testid="stSidebar"] {{
            background: linear-gradient(180deg, rgba(255,255,255,0.96) 0%, rgba(248,250,252,0.96) 100%);
            border-right: 1px solid rgba(148,163,184,0.18);
        }}

        .block-container {{
            padding-top: 0.35rem;
            padding-bottom: 0.35rem;
        }}

        /* =========================
           CARD KPI
        ========================== */
        .card-kpi {{
            background: linear-gradient(180deg, #FDF5E6 0%, #FDF5E6 100%); 
            border: 1px solid #d6e4f0;
            border-radius: 16px;
            padding: 16px;
            min-height: 120px;

            display: flex;
            flex-direction: column;
            justify-content: center;
            align-items: center;

            text-align: center;

            box-shadow: 
                0 6px 18px rgba(0, 0, 0, 0.05),
                0 2px 6px rgba(0, 0, 0, 0.04);

            margin-bottom: 14px;
            transition: all 0.2s ease-in-out;
        }}

        .card-kpi:hover {{
            transform: translateY(-3px);
            box-shadow: 
                0 12px 26px rgba(0, 0, 0, 0.08),
                0 4px 10px rgba(0, 0, 0, 0.06);
        }}

        /* Título */
        .card-kpi-titulo {{
            font-size: 15px;
            font-weight: 700;
            color: #475569;
            margin-bottom: 6px;
        }}

        /* Valor principal */
        .card-kpi-valor {{
            font-size: 34px;
            font-weight: 900;
            color: #0f172a;
            line-height: 1.1;
        }}

        /* Delta */
        .card-kpi-delta {{
            font-size: 14px;
            font-weight: 700;
            margin-top: 6px;
        }}

        .delta-up {{
            color: #16a34a;
        }}

        .delta-down {{
            color: #dc2626;
        }}

        .delta-neutral {{
            color: #64748b;
        }}

        /* Ajuste colunas */
        div[data-testid="stHorizontalBlock"] > div {{
            padding-left: 0.28rem;
            padding-right: 0.28rem;
        }}

        /* Legenda */
        .legenda-box {{
            border-radius: 10px;
            padding: 10px 12px;
            font-weight: 600;
            text-align: center;
            box-shadow: 0 2px 8px rgba(0,0,0,0.04);
        }}

        </style>
        """,
        unsafe_allow_html=True,
    )

def card_kpi(titulo, valor, delta=None):
    icone = ""
    classe = "delta-neutral"

    if delta is not None:
        if delta > 0:
            icone = "↑"
            classe = "delta-up"
        elif delta < 0:
            icone = "↓"
            classe = "delta-down"
        else:
            icone = "→"

    delta_html = ""
    if delta is not None:
        delta_html = f'<div class="card-kpi-delta {classe}">{icone} {abs(delta):.1f}%</div>'

    st.markdown(f"""
    <div class="card-kpi">
        <div class="card-kpi-titulo">{titulo}</div>
        <div class="card-kpi-valor">{valor}</div>
        {delta_html}
    </div>
    """, unsafe_allow_html=True)


inicializar_estado_app()
aplicar_estilo_visual()


def normalizar_texto(valor):
    if pd.isna(valor):
        return ""
    return str(valor).strip()

def normalizar_nome_base_coluna(nome):
    if nome is None:
        return ""
    nome = str(nome).strip().upper()
    nome = re.sub(r"\.\d+$", "", nome)  # remove sufixo .1, .2, .3...
    return nome


def localizar_coluna_por_nome_base(colunas, nome_base):
    nome_base = normalizar_nome_base_coluna(nome_base)
    for i, col in enumerate(colunas):
        if normalizar_nome_base_coluna(col) == nome_base:
            return i, col
    return None, None


def localizar_coluna_direita(colunas, nome_base):
    idx, _ = localizar_coluna_por_nome_base(colunas, nome_base)
    if idx is None:
        return None
    if idx + 1 >= len(colunas):
        return None
    return colunas[idx + 1]


def juntar_textos(a, b, separador=" - "):
    a = normalizar_texto(a)
    b = normalizar_texto(b)

    if a and b:
        return f"{a}{separador}{b}"
    if a:
        return a
    if b:
        return b
    return ""


def normalizar_nro_of_auditoria(valor):
    """
    Normaliza NRO-OF da planilha SD para comparar com o nro_of do painel.
    Ex.: 23.679-00-001 -> 023679-00-001
    """
    if pd.isna(valor):
        return ""

    texto = str(valor).strip().replace(" ", "")
    if texto == "":
        return ""

    partes = texto.split("-")
    if len(partes) == 3:
        p1, p2, p3 = partes

        try:
            p1 = str(int(float(str(p1).replace(".", "").replace(",", ".")))).zfill(6)
        except Exception:
            p1 = str(p1).replace(".", "").zfill(6)

        try:
            p2 = str(int(float(str(p2).replace(".", "").replace(",", ".")))).zfill(2)
        except Exception:
            p2 = str(p2).zfill(2)

        try:
            p3 = str(int(float(str(p3).replace(".", "").replace(",", ".")))).zfill(3)
        except Exception:
            p3 = str(p3).zfill(3)

        return f"{p1}-{p2}-{p3}"

    return texto.replace(".", "")


def ler_planilha_auditoria_sd(uploaded_file):
    nome = uploaded_file.name.lower()

    if nome.endswith(".csv"):
        try:
            return pd.read_csv(uploaded_file, sep=";", dtype=str, encoding="utf-8").fillna("")
        except Exception:
            uploaded_file.seek(0)
            try:
                return pd.read_csv(uploaded_file, sep=";", dtype=str, encoding="latin1").fillna("")
            except Exception:
                uploaded_file.seek(0)
                return pd.read_csv(uploaded_file, dtype=str).fillna("")

    if nome.endswith(".xlsx") or nome.endswith(".xls"):
        return pd.read_excel(uploaded_file, dtype=str).fillna("")

    raise ValueError("Formato não suportado. Envie CSV, XLSX ou XLS.")


def montar_base_auditoria_sd(df_sd):
    if df_sd is None or df_sd.empty:
        return pd.DataFrame(columns=[
            "nro_of_auditoria",
            "Auditoria SD",
            "Cliente SD",
            "Resultado SD",
            "Observações SD",
        ])

    colunas = list(df_sd.columns)

    _, col_ano = localizar_coluna_por_nome_base(colunas, "ANO-SOLIC")
    _, col_nro_solic = localizar_coluna_por_nome_base(colunas, "NRO-SOLIC")
    _, col_cliente = localizar_coluna_por_nome_base(colunas, "CLIENTE")
    _, col_resultado = localizar_coluna_por_nome_base(colunas, "RESULTADO")
    _, col_obs = localizar_coluna_por_nome_base(colunas, "OBSERVACOES")
    _, col_nro_of = localizar_coluna_por_nome_base(colunas, "NRO-OF")

    col_desc_cliente = localizar_coluna_direita(colunas, "CLIENTE")
    col_desc_resultado = localizar_coluna_direita(colunas, "RESULTADO")

    obrigatorias = {
        "ANO-SOLIC": col_ano,
        "NRO-SOLIC": col_nro_solic,
        "CLIENTE": col_cliente,
        "RESULTADO": col_resultado,
        "OBSERVACOES": col_obs,
        "NRO-OF": col_nro_of,
    }

    faltando = [k for k, v in obrigatorias.items() if v is None]
    if faltando:
        raise ValueError(
            "A planilha SD não contém as colunas obrigatórias: " + ", ".join(faltando)
        )

    if col_desc_cliente is None:
        raise ValueError("Não foi possível localizar a DESCRICAO imediatamente à direita de CLIENTE.")

    if col_desc_resultado is None:
        raise ValueError("Não foi possível localizar a DESCRICAO imediatamente à direita de RESULTADO.")

    df = df_sd.copy()

    for c in [col_ano, col_nro_solic, col_cliente, col_desc_cliente, col_resultado, col_desc_resultado, col_obs, col_nro_of]:
        df[c] = df[c].fillna("").astype(str).str.strip()

    df["nro_of_auditoria"] = df[col_nro_of].apply(normalizar_nro_of_auditoria)

    df["Auditoria SD"] = (
        df[col_nro_solic].astype(str).str.strip()
        + " / "
        + df[col_ano].astype(str).str.strip()
    ).str.strip(" /")

    df["Cliente SD"] = df.apply(
        lambda row: juntar_textos(row[col_cliente], row[col_desc_cliente]),
        axis=1
    )

    df["Resultado SD"] = df.apply(
        lambda row: juntar_textos(row[col_resultado], row[col_desc_resultado]),
        axis=1
    )

    df["Observações SD"] = df[col_obs].astype(str).str.strip()

    df = df[df["nro_of_auditoria"] != ""].copy()

    def agregar_textos(series):
        valores = []
        vistos = set()

        for valor in series.fillna("").astype(str).str.strip():
            if valor and valor not in vistos:
                vistos.add(valor)
                valores.append(valor)

        return " | ".join(valores)

    df_final = (
        df.groupby("nro_of_auditoria", as_index=False)
          .agg({
              "Auditoria SD": agregar_textos,
              "Cliente SD": agregar_textos,
              "Resultado SD": agregar_textos,
              "Observações SD": agregar_textos,
          })
    )

    return df_final


def aplicar_auditoria_sd_no_dataframe(df_base, df_auditoria_sd):
    df = df_base.copy()

    if df_auditoria_sd is None or df_auditoria_sd.empty:
        for col in ["Auditoria SD", "Cliente SD", "Resultado SD", "Observações SD"]:
            if col not in df.columns:
                df[col] = ""
        return df

    df["nro_of_join_sd"] = df["nro_of"].apply(normalizar_nro_of_auditoria)

    df = df.merge(
        df_auditoria_sd,
        how="left",
        left_on="nro_of_join_sd",
        right_on="nro_of_auditoria"
    )

    for col in ["Auditoria SD", "Cliente SD", "Resultado SD", "Observações SD"]:
        if col not in df.columns:
            df[col] = ""
        df[col] = df[col].fillna("").astype(str)

    for col in ["nro_of_join_sd", "nro_of_auditoria"]:
        if col in df.columns:
            df = df.drop(columns=[col])

    return df


def render_upload_auditoria_sd():
    with st.expander("Auditoria SD - Upload da planilha"):
        arquivo_sd = st.file_uploader(
            "Envie a planilha da SD",
            type=["csv", "xlsx", "xls"],
            key="upload_sd_auditoria"
        )

        if arquivo_sd is not None:
            try:
                df_sd = ler_planilha_auditoria_sd(arquivo_sd)
                df_sd_aud = montar_base_auditoria_sd(df_sd)
                st.session_state["df_sd_auditoria"] = df_sd_aud.copy()

                st.success(
                    f"Planilha carregada com sucesso. {len(df_sd_aud)} OF(s) distintas preparadas para auditoria."
                )

                with st.expander("Visualizar base consolidada da auditoria SD"):
                    st.dataframe(df_sd_aud, use_container_width=True, hide_index=True)

            except Exception as e:
                st.error(f"Erro ao processar a planilha de auditoria SD: {e}")

def formatar_numero_br(valor, casas=0):
    if pd.isna(valor):
        return ""
    try:
        valor = float(valor)
        texto = f"{valor:,.{casas}f}"
        texto = texto.replace(",", "X").replace(".", ",").replace("X", ".")
        return texto
    except Exception:
        return str(valor)


def validar_data_br(valor):
    if pd.isna(valor):
        return True

    valor = str(valor).strip()
    if valor == "":
        return True

    if not re.match(r"^\d{2}/\d{2}/\d{4}$", valor):
        return False

    try:
        datetime.strptime(valor, "%d/%m/%Y")
        return True
    except ValueError:
        return False


def converter_serie_data(valor):
    if pd.isna(valor):
        return pd.NaT

    valor = str(valor).strip()
    if valor == "":
        return pd.NaT

    dt = pd.to_datetime(valor, format="%d/%m/%Y", dayfirst=True, errors="coerce")
    if pd.notna(dt):
        return dt

    return pd.to_datetime(valor, errors="coerce")


def converter_data_br_para_ordenacao(valor):
    if pd.isna(valor):
        return pd.NaT

    valor = str(valor).strip()
    if valor == "":
        return pd.NaT

    return pd.to_datetime(valor, format="%d/%m/%Y", dayfirst=True, errors="coerce")


def prioridade_para_ordem(valor):
    if pd.isna(valor):
        return 999999

    valor = str(valor).strip().upper()
    if valor == "":
        return 999999

    if re.fullmatch(r"\d+", valor):
        return int(valor)

    match = re.search(r"(\d+)", valor)
    if match:
        return int(match.group(1))

    return 999999


def garantir_colunas_novas(df):
    defaults = {
        "Nw_Data": "",
        "Prioridade": "",
        "Responsavel": "",
        "Alteracao": "",
        "Remover": False,
    }

    for col, valor in defaults.items():
        if col not in df.columns:
            df[col] = valor

    return df


def lista_multiselect(df, coluna):
    if coluna not in df.columns:
        return []

    valores = df[coluna].fillna("").astype(str).str.strip()
    valores = valores[valores != ""]
    return sorted(valores.unique().tolist())


def aplicar_filtro_multiselect(df, coluna, selecionados):
    if coluna not in df.columns or not selecionados:
        return df

    serie = df[coluna].fillna("").astype(str).str.strip()
    return df[serie.isin(selecionados)]


def aplicar_filtro_data(df, coluna, data_inicio=None, data_fim=None):
    if coluna not in df.columns:
        return df

    serie_data = df[coluna].apply(converter_serie_data)

    if data_inicio is not None:
        df = df[serie_data >= pd.to_datetime(data_inicio)]
        serie_data = df[coluna].apply(converter_serie_data)

    if data_fim is not None:
        df = df[serie_data <= pd.to_datetime(data_fim)]

    return df


def render_card_kpi(titulo, valor):
    st.markdown(
        f"""
        <div class="card-kpi">
            <div class="card-kpi-titulo">{titulo}</div>
            <div class="card-kpi-valor">{valor}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def iniciar_loading():
    status_placeholder = st.empty()
    progress_bar = st.progress(0)
    start_time = time.time()
    return status_placeholder, progress_bar, start_time


def atualizar_loading(status_placeholder, progress_bar, mensagem, etapa, total_etapas, start_time):
    percentual = int((etapa / total_etapas) * 100)
    tempo_decorrido = time.time() - start_time

    status_placeholder.markdown(
        f"""
        <div style="
            background-color:#f8fafc;
            border:1px solid #e2e8f0;
            padding:12px 16px;
            border-radius:10px;
            margin-bottom:10px;
        ">
            <b>🧪 Olá, estamos processando os dados do Painel Laboratório.</b><br>
            Aguarde um momento.<br>
            <hr style="margin:8px 0">
            <b>{mensagem}</b><br>
            📊 Progresso: <b>{percentual}%</b><br>
            ⏱ Tempo: <b>{tempo_decorrido:.1f}s</b>
        </div>
        """,
        unsafe_allow_html=True,
    )

    progress_bar.progress(percentual)


def carregar_banco_txt(caminho=ARQUIVO_BANCO):
    colunas = ["Nro_OF", "Codigo_Produto", "Nw_Data", "Prioridade", "Responsavel", "Alteracao"]

    if not os.path.exists(caminho):
        return pd.DataFrame(columns=colunas)

    try:
        df_banco = pd.read_csv(caminho, sep=";", dtype=str, encoding="utf-8").fillna("")

        for col in colunas:
            if col not in df_banco.columns:
                df_banco[col] = ""

        df_banco = df_banco[colunas].copy()
        df_banco["Nro_OF"] = df_banco["Nro_OF"].apply(normalizar_texto)
        df_banco["Codigo_Produto"] = df_banco["Codigo_Produto"].apply(normalizar_texto)
        df_banco["Nw_Data"] = df_banco["Nw_Data"].apply(normalizar_texto)
        df_banco["Prioridade"] = df_banco["Prioridade"].apply(normalizar_texto).str.upper()
        df_banco["Responsavel"] = df_banco["Responsavel"].apply(normalizar_texto)
        df_banco["Alteracao"] = df_banco["Alteracao"].apply(normalizar_texto)

        df_banco = df_banco.drop_duplicates(subset=["Nro_OF", "Codigo_Produto"], keep="last")
        return df_banco

    except Exception as e:
        st.error(f"Erro ao ler o arquivo {caminho}: {e}")
        return pd.DataFrame(columns=colunas)


def montar_texto_alteracao(valor_antigo, valor_novo, campo, timestamp):
    antigo = normalizar_texto(valor_antigo)
    novo = normalizar_texto(valor_novo)

    if campo.lower() == "prioridade":
        antigo = antigo.upper()
        novo = novo.upper()

    if antigo == novo:
        return ""

    antigo_exib = antigo if antigo != "" else "(vazio)"
    novo_exib = novo if novo != "" else "(vazio)"

    return f"{timestamp} | {campo} | De: {antigo_exib} | Para: {novo_exib}"


def consolidar_alteracoes(texto_existente, novas_linhas):
    texto_existente = normalizar_texto(texto_existente)
    novas_linhas = [normalizar_texto(x) for x in novas_linhas if normalizar_texto(x) != ""]

    if not novas_linhas:
        return texto_existente

    partes = []
    if texto_existente:
        partes.append(texto_existente)

    partes.extend(novas_linhas)
    return " || ".join(partes)


def salvar_banco_txt(df_tabela, caminho=ARQUIVO_BANCO):
    colunas_saida = ["Nro_OF", "Codigo_Produto", "Nw_Data", "Prioridade", "Responsavel", "Alteracao"]

    df_banco_atual = carregar_banco_txt(caminho).copy()

    for col in colunas_saida:
        if col not in df_banco_atual.columns:
            df_banco_atual[col] = ""

    if df_banco_atual.empty:
        df_banco_atual = pd.DataFrame(columns=colunas_saida)

    df_edit = df_tabela.copy()

    rename_map = {
        "Nro OF": "Nro_OF",
        "Código Produto": "Codigo_Produto",
        "Responsavel": "Responsavel",
    }
    df_edit = df_edit.rename(columns=rename_map)

    for col in colunas_saida:
        if col not in df_edit.columns:
            df_edit[col] = ""

    if "Remover" not in df_edit.columns:
        df_edit["Remover"] = False

    df_edit["Remover"] = df_edit["Remover"].fillna(False).astype(bool)

    for col in ["Nro_OF", "Codigo_Produto", "Nw_Data", "Prioridade", "Responsavel", "Alteracao"]:
        df_edit[col] = df_edit[col].apply(normalizar_texto)

    df_edit["Prioridade"] = df_edit["Prioridade"].str.upper()
    

    df_edit = df_edit[
        (df_edit["Nro_OF"].astype(str).str.strip() != "")
        & (df_edit["Codigo_Produto"].astype(str).str.strip() != "")
    ].copy()

    df_edit = df_edit.drop_duplicates(subset=["Nro_OF", "Codigo_Produto"], keep="last")

    timestamp = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

    mapa_banco = {}
    if not df_banco_atual.empty:
        for _, row in df_banco_atual.iterrows():
            chave = (row["Nro_OF"], row["Codigo_Produto"])
            mapa_banco[chave] = {
                "Nw_Data": normalizar_texto(row.get("Nw_Data", "")),
                "Prioridade": normalizar_texto(row.get("Prioridade", "")),
                "Responsavel": normalizar_texto(row.get("Responsavel", "")),
                "Alteracao": normalizar_texto(row.get("Alteracao", "")),
            }

    alteracoes_finais = []
    for _, row in df_edit.iterrows():
        chave = (row["Nro_OF"], row["Codigo_Produto"])
        base_antiga = mapa_banco.get(
            chave,
            {"Nw_Data": "", "Prioridade": "", "Responsavel": "", "Alteracao": ""}
        )

        # 🚀 PROTEÇÃO: não sobrescrever com vazio
        if row["Responsavel"] == "":
            row["Responsavel"] = base_antiga.get("Responsavel", "")

        novas_linhas = [
            montar_texto_alteracao(base_antiga.get("Prioridade", ""), row["Prioridade"], "Prioridade", timestamp),
            montar_texto_alteracao(base_antiga.get("Nw_Data", ""), row["Nw_Data"], "Nw_Data", timestamp),
            montar_texto_alteracao(base_antiga.get("Responsavel", ""), row["Responsavel"], "Responsavel", timestamp),
        ]

        texto_alteracao = consolidar_alteracoes(base_antiga.get("Alteracao", ""), novas_linhas)
        alteracoes_finais.append(texto_alteracao)

    df_edit["Alteracao"] = alteracoes_finais

    chaves_remover = set(
        zip(
            df_edit.loc[df_edit["Remover"], "Nro_OF"],
            df_edit.loc[df_edit["Remover"], "Codigo_Produto"],
        )
    )

    chaves_editadas = set(zip(df_edit["Nro_OF"], df_edit["Codigo_Produto"]))

    if not df_banco_atual.empty:
        df_banco_atual = df_banco_atual[
            ~df_banco_atual.apply(
                lambda row: (row["Nro_OF"], row["Codigo_Produto"]) in chaves_editadas,
                axis=1,
            )
        ].copy()

    df_edit_manter = df_edit[~df_edit["Remover"]].copy()
    df_edit_manter = df_edit_manter[colunas_saida].copy()

    df_final = pd.concat([df_banco_atual[colunas_saida], df_edit_manter], ignore_index=True)

    df_final = df_final[
        (df_final["Prioridade"].astype(str).str.strip() != "")
        | (df_final["Nw_Data"].astype(str).str.strip() != "")
        | (df_final["Responsavel"].astype(str).str.strip() != "")
        | (df_final["Alteracao"].astype(str).str.strip() != "")
    ].copy()

    if chaves_remover:
        df_final = df_final[
            ~df_final.apply(
                lambda row: (row["Nro_OF"], row["Codigo_Produto"]) in chaves_remover,
                axis=1,
            )
        ].copy()

    df_final = df_final.drop_duplicates(subset=["Nro_OF", "Codigo_Produto"], keep="last")

    df_final["_ord_prioridade"] = df_final["Prioridade"].apply(prioridade_para_ordem)
    df_final["_ord_data"] = df_final["Nw_Data"].apply(converter_data_br_para_ordenacao)

    df_final = df_final.sort_values(
        by=["_ord_prioridade", "_ord_data", "Nro_OF", "Codigo_Produto"],
        ascending=[True, True, True, True],
        na_position="last",
    ).drop(columns=["_ord_prioridade", "_ord_data"])

    df_final.to_csv(caminho, sep=";", index=False, encoding="utf-8")


def aplicar_banco_txt(df_principal, df_banco):
    df = df_principal.copy()
    df["nro_of"] = df["nro_of"].apply(normalizar_texto)
    df["codigo_produto"] = df["codigo_produto"].apply(normalizar_texto)

    if df_banco.empty:
        return garantir_colunas_novas(df)

    df_merge = df.merge(
        df_banco,
        how="left",
        left_on=["nro_of", "codigo_produto"],
        right_on=["Nro_OF", "Codigo_Produto"],
    )

    for col_drop in ["Nro_OF", "Codigo_Produto"]:
        if col_drop in df_merge.columns:
            df_merge = df_merge.drop(columns=[col_drop])

    df_merge = garantir_colunas_novas(df_merge)
    df_merge["Prioridade"] = df_merge["Prioridade"].fillna("").astype(str).str.upper()
    df_merge["Nw_Data"] = df_merge["Nw_Data"].fillna("").astype(str)
    df_merge["Alteracao"] = df_merge["Alteracao"].fillna("").astype(str)
    df_merge["Remover"] = False

    return df_merge


def ordenar_dataframe(df):
    df = df.copy()
    df["_ord_prioridade"] = df["Prioridade"].apply(prioridade_para_ordem)
    df["_ord_data"] = df["Nw_Data"].apply(converter_data_br_para_ordenacao)

    df = df.sort_values(
        by=["_ord_prioridade", "_ord_data", "nro_of", "codigo_produto"],
        ascending=[True, True, True, True],
        na_position="last",
    ).drop(columns=["_ord_prioridade", "_ord_data"])

    return df


@st.cache_data(ttl=300, show_spinner=False)
def carregar_base_do_banco(status_list):
    return carregar_ofs_laboratorio(status_list=status_list)


def carregar_base_principal_com_controle():
    filtros = st.session_state["filtros_aplicados"]
    status_list = filtros["status_of"] if filtros["status_of"] else ["A"]

    precisa_recarregar = (
        st.session_state["forcar_atualizacao"]
        or st.session_state["df_base"] is None
    )

    if not precisa_recarregar:
        return st.session_state["df_base"].copy()

    status_placeholder = None
    progress_bar = None
    start_time = None

    if st.session_state["mostrar_loading"]:
        status_placeholder, progress_bar, start_time = iniciar_loading()
        total_etapas = 4
        atualizar_loading(status_placeholder, progress_bar, "Consultando ORDEM_FABRIC...", 1, total_etapas, start_time)

    df_origem = carregar_base_do_banco(tuple(status_list))

    if st.session_state["mostrar_loading"]:
        atualizar_loading(status_placeholder, progress_bar, "Lendo banco_laboratorio.txt...", 2, total_etapas, start_time)

    df_banco = carregar_banco_txt()

    if st.session_state["mostrar_loading"]:
        atualizar_loading(status_placeholder, progress_bar, "Aplicando colunas locais...", 3, total_etapas, start_time)

    df_base = aplicar_banco_txt(df_origem, df_banco)
    df_base = garantir_colunas_novas(df_base)
    df_base = ordenar_dataframe(df_base)

    if st.session_state["mostrar_loading"]:
        atualizar_loading(status_placeholder, progress_bar, "Finalizando painel...", 4, total_etapas, start_time)
        status_placeholder.empty()
        progress_bar.empty()

    st.session_state["df_base"] = df_base.copy()
    st.session_state["forcar_atualizacao"] = False
    st.session_state["mostrar_loading"] = False

    return df_base.copy()


def reaplicar_banco_sem_reconsultar_base():
    if st.session_state["df_base"] is None:
        return

    df_banco = carregar_banco_txt()
    df_base_atual = st.session_state["df_base"].copy()

    colunas_remover = ["Nw_Data", "Prioridade", "Responsavel", "Alteracao", "Remover"]
    for col in colunas_remover:
        if col in df_base_atual.columns:
            df_base_atual = df_base_atual.drop(columns=[col])

    df_base_novo = aplicar_banco_txt(df_base_atual, df_banco)
    df_base_novo = garantir_colunas_novas(df_base_novo)
    df_base_novo = ordenar_dataframe(df_base_novo)

    st.session_state["df_base"] = df_base_novo.copy()
    st.session_state["df_grid_editado"] = None
    st.session_state["grid_key"] += 1


def montar_visao_filtrada(df_base, filtros):
    df_filtrado = df_base.copy()

    df_filtrado = aplicar_filtro_multiselect(df_filtrado, "status_of", filtros["status_of"])
    df_filtrado = aplicar_filtro_multiselect(df_filtrado, "desc_cliente", filtros["cliente"])
    df_filtrado = aplicar_filtro_multiselect(df_filtrado, "codigo_produto", filtros["codigo_produto"])
    df_filtrado = aplicar_filtro_multiselect(df_filtrado, "codigo_original", filtros["codigo_original"])
    df_filtrado = aplicar_filtro_multiselect(df_filtrado, "GP_codigo_grupo", filtros["codigo_grupo"])
    df_filtrado = aplicar_filtro_multiselect(df_filtrado, "desc_operador_atual", filtros["operador_atual"])

    if filtros["sequencia_atual"]:
        seqs = [int(x) for x in filtros["sequencia_atual"]]
        df_filtrado = df_filtrado[df_filtrado["sequencia_atual"].isin(seqs)]

    if filtros["nw_data_inicio"] is not None:
        df_filtrado = aplicar_filtro_data(df_filtrado, "Nw_Data", data_inicio=filtros["nw_data_inicio"])

    if filtros["nw_data_fim"] is not None:
        df_filtrado = aplicar_filtro_data(df_filtrado, "Nw_Data", data_fim=filtros["nw_data_fim"])

    if filtros["prev_entrega_inicio"] is not None:
        df_filtrado = aplicar_filtro_data(df_filtrado, "data_prev_entrega", data_inicio=filtros["prev_entrega_inicio"])

    if filtros["prev_entrega_fim"] is not None:
        df_filtrado = aplicar_filtro_data(df_filtrado, "data_prev_entrega", data_fim=filtros["prev_entrega_fim"])

    if filtros["abertura_inicio"] is not None:
        df_filtrado = aplicar_filtro_data(df_filtrado, "data_abertura", data_inicio=filtros["abertura_inicio"])

    if filtros["abertura_fim"] is not None:
        df_filtrado = aplicar_filtro_data(df_filtrado, "data_abertura", data_fim=filtros["abertura_fim"])

    df_filtrado = ordenar_dataframe(df_filtrado)
    return df_filtrado


def render_sidebar_filtros(df):
    with st.sidebar:
        if os.path.exists(ARQUIVO_LOGO_SIDEBAR):
            c1, c2, c3 = st.columns([1, 8, 1])
            with c2:
                st.image(ARQUIVO_LOGO_SIDEBAR, width=320)

        st.markdown("## Filtros")

        filtros_atuais = st.session_state["filtros_aplicados"]

        opcoes_status = ["A", "F"]
        opcoes_cliente = lista_multiselect(df, "desc_cliente")
        opcoes_codigo_produto = lista_multiselect(df, "codigo_produto")
        opcoes_codigo_original = lista_multiselect(df, "codigo_original")
        opcoes_codigo_grupo = lista_multiselect(df, "GP_codigo_grupo")
        opcoes_operador = lista_multiselect(df, "desc_operador_atual")
        opcoes_seq = sorted(
            pd.to_numeric(df["sequencia_atual"], errors="coerce")
            .dropna()
            .astype(int)
            .astype(str)
            .unique()
            .tolist()
        ) if "sequencia_atual" in df.columns else []

        with st.form("form_filtros_sidebar", clear_on_submit=False):
            filtro_status = st.multiselect("Status OF", opcoes_status, default=filtros_atuais.get("status_of", ["A"]) or ["F"])
            filtro_cliente = st.multiselect("Cliente", opcoes_cliente, default=filtros_atuais.get("cliente", []))
            filtro_codigo_produto = st.multiselect("Código Produto", opcoes_codigo_produto, default=filtros_atuais.get("codigo_produto", []))
            filtro_codigo_original = st.multiselect("Código Original", opcoes_codigo_original, default=filtros_atuais.get("codigo_original", []))
            filtro_codigo_grupo = st.multiselect("Grupo", opcoes_codigo_grupo, default=filtros_atuais.get("codigo_grupo", []))
            filtro_operador = st.multiselect("Operador Atual", opcoes_operador, default=filtros_atuais.get("operador_atual", []))
            filtro_seq = st.multiselect("Sequência Atual", opcoes_seq, default=[str(x) for x in filtros_atuais.get("sequencia_atual", [])])

            st.markdown("### Nw_Data")
            filtro_nw_data_inicio = st.date_input("Nw_Data início", value=filtros_atuais["nw_data_inicio"], format="DD/MM/YYYY")
            filtro_nw_data_fim = st.date_input("Nw_Data final", value=filtros_atuais["nw_data_fim"], format="DD/MM/YYYY")

            st.markdown("### Previsão Entrega")
            filtro_prev_entrega_inicio = st.date_input("Prev. Entrega início", value=filtros_atuais["prev_entrega_inicio"], format="DD/MM/YYYY")
            filtro_prev_entrega_fim = st.date_input("Prev. Entrega final", value=filtros_atuais["prev_entrega_fim"], format="DD/MM/YYYY")

            hoje = datetime.today()
            inicio_ano = datetime(hoje.year, 1, 1)

            st.markdown("### Abertura")
            filtro_abertura_inicio = st.date_input("Abertura início", value=filtros_atuais.get("abertura_inicio") or inicio_ano, format="DD/MM/YYYY")
            filtro_abertura_fim = st.date_input("Abertura final", value=filtros_atuais.get("abertura_fim") or hoje, format="DD/MM/YYYY")

            col_btn1, col_btn2, col_btn3 = st.columns(3)
            aplicar_filtros = col_btn1.form_submit_button("Aplicar", use_container_width=True)
            limpar = col_btn2.form_submit_button("Limpar", use_container_width=True)
            atualizar = col_btn3.form_submit_button("Atualizar", use_container_width=True)

        if aplicar_filtros:
            novo_status = filtro_status if filtro_status else ["A"]

            st.session_state["filtros_aplicados"] = {
                "status_of": novo_status,
                "cliente": filtro_cliente,
                "codigo_produto": filtro_codigo_produto,
                "codigo_original": filtro_codigo_original,
                "codigo_grupo": filtro_codigo_grupo,
                "operador_atual": filtro_operador,
                "sequencia_atual": [int(x) for x in filtro_seq],
                "nw_data_inicio": filtro_nw_data_inicio,
                "nw_data_fim": filtro_nw_data_fim,
                "prev_entrega_inicio": filtro_prev_entrega_inicio,
                "prev_entrega_fim": filtro_prev_entrega_fim,
                "abertura_inicio": filtro_abertura_inicio,
                "abertura_fim": filtro_abertura_fim,
            }

            if novo_status != filtros_atuais["status_of"]:
                st.session_state["forcar_atualizacao"] = True
                st.rerun()

        if limpar:
            st.session_state["filtros_aplicados"] = {
                "status_of": ["A"],
                "cliente": [],
                "codigo_produto": [],
                "codigo_original": [],
                "codigo_grupo": [],
                "operador_atual": [],
                "sequencia_atual": [],
                "nw_data_inicio": None,
                "nw_data_fim": None,
                "prev_entrega_inicio": None,
                "prev_entrega_fim": None,
                "abertura_inicio": None,
                "abertura_fim": None,
            }
            st.session_state["forcar_atualizacao"] = True
            st.rerun()

        if atualizar:
            st.session_state["forcar_atualizacao"] = True
            st.session_state["mostrar_loading"] = True
            st.rerun()


date_mask_editor = JsCode(
    """
class DateMaskEditor {
    init(params) {
        this.eInput = document.createElement('input');
        this.eInput.value = params.value || '';
        this.eInput.style.width = '100%';

        this.eInput.addEventListener('input', (e) => {
            let v = e.target.value.replace(/\\D/g, '');
            if (v.length > 2) v = v.slice(0,2) + '/' + v.slice(2);
            if (v.length > 5) v = v.slice(0,5) + '/' + v.slice(5,9);
            e.target.value = v;
        });
    }

    getGui() { return this.eInput; }
    afterGuiAttached() { this.eInput.focus(); }
    getValue() { return this.eInput.value; }
}
"""
)

prioridade_editor = JsCode(
    """
class PrioridadeEditor {
    init(params) {
        this.eInput = document.createElement('input');
        this.eInput.value = params.value || '';
        this.eInput.style.width = '100%';
        this.eInput.maxLength = 3;

        this.eInput.addEventListener('input', (e) => {
            let v = e.target.value.toUpperCase();
            v = v.replace(/[^A-Z0-9]/g, '');
            v = v.substring(0, 3);
            e.target.value = v;
        });
    }

    getGui() { return this.eInput; }
    afterGuiAttached() { this.eInput.focus(); }
    getValue() { return this.eInput.value; }
}
"""
)

prioridade_style = JsCode(
    """
function(params) {
    const value = (params.value || '').toString();
    if (!value) return {};
    const regex = /^[A-Z0-9]{1,3}$/;
    if (!regex.test(value)) {
        return {backgroundColor: '#ffe5e5'};
    }
    return {};
}
"""
)

cell_style_date = JsCode(
    """
function(params) {
    const value = params.value;
    if (!value) return {};
    const regex = /^\\d{2}\\/\\d{2}\\/\\d{4}$/;
    if (!regex.test(value)) {
        return {backgroundColor: '#ffe5e5'};
    }
    return {};
}
"""
)
row_style = JsCode("""
function(params) {
    const status = (params.data["Status OF"] || "").toString().trim();
    const prioridade = (params.data["Prioridade"] || "").toString().trim();
    const semaforo = (params.data["Semáforo"] || "").toString().trim();

    // 🔴 ATRASADO = prioridade máxima visual
    if (semaforo === "🔴") {
        return {
            'backgroundColor': '#ffe5e5',
            'color': '#7f1d1d',
            'fontWeight': '700'
        };
    }

    // 🟡 PRIORIDADE (sem estar atrasado)
    if (prioridade !== "") {
        return {
            'backgroundColor': '#FFE4B5',
            'color': '#92400e',
            'fontWeight': '600'
        };
    }

    // 🟢 STATUS
    if (status === "A") {
        return {
            'backgroundColor': '#F0FFF0'
        };
    }

    if (status === "F") {
        return {
            'backgroundColor': '#f1fbf5'
        };
    }

    return {};
}
""")


tooltip_js = JsCode("""
function(params) {
    if (params.value == null || params.value === "") {
        return "";
    }
    return params.value.toString();
}
""")

def preparar_dataframe_exibicao(df_filtrado):
    df_exibicao = df_filtrado.copy()

    df_exibicao["data_abertura"] = pd.to_datetime(
        df_exibicao["data_abertura"], errors="coerce"
    ).dt.strftime("%d/%m/%Y")

    df_exibicao["data_prev_entrega"] = pd.to_datetime(
        df_exibicao["data_prev_entrega"], errors="coerce"
    ).dt.strftime("%d/%m/%Y")

    df_exibicao["data_final_apontamento"] = pd.to_datetime(
        df_exibicao["data_final_apontamento"], errors="coerce"
    ).dt.strftime("%d/%m/%Y")

    hoje = pd.Timestamp.today().normalize()

    prev_dt = pd.to_datetime(df_filtrado["data_prev_entrega"], errors="coerce")
    prev_dt = prev_dt.dt.normalize()

    nw_dt = pd.to_datetime(
        df_filtrado["Nw_Data"].astype(str).str.strip(),
        format="%d/%m/%Y",
        dayfirst=True,
        errors="coerce"
    )
    nw_dt = nw_dt.dt.normalize()

    data_referencia = nw_dt.where(nw_dt.notna(), prev_dt)
    dias_atraso = (hoje - data_referencia).dt.days

    df_exibicao["Dias Atraso"] = dias_atraso.apply(
        lambda x: str(int(x)) if pd.notna(x) and x > 0 else ""
    ).astype(str)

    df_exibicao["Semáforo"] = dias_atraso.apply(
        lambda x:
            "🔴" if pd.notna(x) and x > 0 else
            "🟡" if pd.notna(x) and -3 <= x <= 0 else
            "🟢" if pd.notna(x) and x < -3 else
            ""
    ).astype(str)

    for col in [
        "Nw_Data",
        "Prioridade",
        "Responsavel",
        "Alteracao",
        "desc_operacao_atual",
        "desc_operador_atual",
        "operacoes_percorridas",
        "codigo_original",
        "GP_codigo_grupo",
        "SGP_codigo_subgrupo",
    ]:
        if col in df_exibicao.columns:
            df_exibicao[col] = df_exibicao[col].fillna("").astype(str)

    df_exibicao["sequencia_atual"] = pd.to_numeric(
        df_exibicao["sequencia_atual"], errors="coerce"
    ).fillna(0).astype(int)

    df_exibicao["Remover"] = df_exibicao["Remover"].fillna(False).astype(bool)

    colunas_exibicao = [
        "Remover",
        "Prioridade",
        "Responsavel",
        "Nw_Data",
        "Dias Atraso",
        "Semáforo",
        "data_abertura",
        "data_prev_entrega",
        "origem",
        "nro_of",
        "status_of",
        "codigo_produto",
        "desc_cliente",
        "Auditoria SD",
        "Cliente SD",
        "Resultado SD",
        "Observações SD",
        "codigo_original",
        "GP_codigo_grupo",
        "SGP_codigo_subgrupo",
        "sequencia_atual",
        "data_final_apontamento",
        "desc_operacao_atual",
        "desc_operador_atual",
        "operacoes_percorridas",
        "Alteracao",
    ]

    colunas_exibicao = [c for c in colunas_exibicao if c in df_exibicao.columns]
    df_exibicao = df_exibicao[colunas_exibicao]

    return df_exibicao.rename(
        columns={
            "data_abertura": "Abertura",
            "data_prev_entrega": "Prev. Entrega",
            "origem": "Origem",
            "nro_of": "Nro OF",
            "status_of": "Status OF",
            "codigo_produto": "Código Produto",
            "desc_cliente": "Cliente",
            "codigo_original": "Código Original",
            "GP_codigo_grupo": "Grupo",
            "SGP_codigo_subgrupo": "Subgrupo",
            "sequencia_atual": "Sequência Atual",
            "data_final_apontamento": "Data Final",
            "desc_operacao_atual": "Operação Atual",
            "desc_operador_atual": "Operador Atual",
            "Responsavel": "Responsavel",
            "operacoes_percorridas": "Operações Percorridas",
            "Alteracao": "Alteração",
        }
    )


def render_grid(df_exibicao):
    st.markdown("### Posição das SDs / OFs do Laboratório")

    df_grid = df_exibicao.copy()
    
    if "Responsavel" in df_grid.columns:
        df_grid["Responsavel"] = df_grid["Responsavel"].fillna("").astype(str).str.strip()

    gb = GridOptionsBuilder.from_dataframe(df_grid)

    gb.configure_default_column(
        editable=False,
        filter=True,
        sortable=True,
        resizable=True
    )

    for col in df_grid.columns:
        gb.configure_column(col, tooltipValueGetter=tooltip_js)

    gb.configure_grid_options(
        domLayout="normal",
        headerHeight=42,
        getRowStyle=row_style,
        suppressHorizontalScroll=False
    )

    gb.configure_column(
    "Remover",
    editable=True,
    cellRenderer="agCheckboxCellRenderer",
    cellEditor="agCheckboxCellEditor",
    width=90,
    minWidth=80,
    maxWidth=100
    )
    gb.configure_column("Prioridade", editable=True, cellEditor=prioridade_editor, cellStyle=prioridade_style, width=100)
    gb.configure_column("Operações Percorridas", width=400, wrapText=True, autoHeight=True)
    gb.configure_column("Responsavel",editable=True,cellEditor="agTextCellEditor",width=220)
    gb.configure_column("Nw_Data", editable=True, cellEditor=date_mask_editor, cellStyle=cell_style_date, width=110)
    gb.configure_column("Dias Atraso", width=110, type=["textColumn"])
    gb.configure_column("Semáforo", width=90)
    gb.configure_column("Alteração", width=800, wrapText=True, autoHeight=True)
    gb.configure_column("Abertura", width=110)
    gb.configure_column("Prev. Entrega", width=120)
    gb.configure_column("Origem", width=90)
    gb.configure_column("Nro OF", width=130)
    gb.configure_column("Status OF", width=90)
    gb.configure_column("Código Produto", width=130)
    gb.configure_column("Cliente", width=240, wrapText=True, autoHeight=True)
    gb.configure_column("Código Original", width=130)
    gb.configure_column("Grupo", width=90)
    gb.configure_column("Subgrupo", width=90)
    gb.configure_column("Sequência Atual", width=110)
    gb.configure_column("Data Final", width=120)
    gb.configure_column("Operação Atual", width=220)
    gb.configure_column("Operador Atual", width=180,wrapText=True, autoHeight=True)
    gb.configure_column("Cliente SD", width=260)
    gb.configure_column("Resultado SD", width=260, wrapText=True, autoHeight=True)
    gb.configure_column("Observações SD", width=300, wrapText=True, autoHeight=True)
    gb.configure_column("Auditoria SD", width=140, wrapText=True, autoHeight=True)

    grid_response = AgGrid(
        df_grid,
        gridOptions=gb.build(),
        update_mode=GridUpdateMode.MODEL_CHANGED,
        data_return_mode="AS_INPUT",
        fit_columns_on_grid_load=False,
        allow_unsafe_jscode=True,
        enable_enterprise_modules=False,
        theme="streamlit",
        custom_css={
            ".ag-theme-streamlit .ag-header-cell-text": {
                "font-size": "15px",
                "font-weight": "700"
            },
            ".ag-theme-streamlit .ag-cell": {
                "font-size": "16px",
                "display": "flex",
                "align-items": "center"
            }
        },
        height=950,
        reload_data=True,
        key=f"grid_lab_{st.session_state['grid_key']}",
    )
    if "Responsável" in df_grid.columns:
        df_grid["Responsável"] = df_grid["Responsável"].fillna("").astype(str).str.strip()
        
    df_editado = pd.DataFrame(grid_response["data"])

    if (
        df_editado is not None
        and not df_editado.empty
        and "Responsavel" in df_editado.columns
        and "Responsavel" in df_grid.columns
    ):
        serie_editado = df_editado["Responsavel"].fillna("").astype(str).str.strip()
        serie_grid = df_grid["Responsavel"].fillna("").astype(str).str.strip()

        mask_vazio = serie_editado.eq("")

        if len(df_editado) == len(df_grid):
            df_editado.loc[mask_vazio, "Responsavel"] = serie_grid.values[mask_vazio.values]

    if df_editado is not None and not df_editado.empty:
        st.session_state["df_grid_editado"] = df_editado.copy()
    else:
        df_editado = df_grid.copy()

    return df_editado


def render_metricas(df_filtrado):
    hoje = pd.Timestamp.today().normalize()
    inicio_mes = hoje.replace(day=1)

    df = df_filtrado.copy()
    df["data_abertura_dt"] = pd.to_datetime(df["data_abertura"], errors="coerce")
    df["data_prev_entrega_dt"] = pd.to_datetime(df["data_prev_entrega"], errors="coerce")
    df["nw_data_dt"] = df["Nw_Data"].apply(converter_serie_data)

    status_upper = df["status_of"].fillna("").astype(str).str.strip().str.upper()
    grupo = df["GP_codigo_grupo"].fillna("").astype(str).str.strip()

    total_abertas = int((status_upper == "A").sum())
    total_fechadas = int((status_upper == "F").sum())

    abertas_mes = int(((status_upper == "A") & (df["data_abertura_dt"] >= inicio_mes)).sum())
    fechadas_mes = int(((status_upper == "F") & (df["data_abertura_dt"] >= inicio_mes)).sum())

    em_andamento = int(((status_upper == "A") & (df["sequencia_atual"] > 0)).sum())
    paradas = int(((status_upper == "A") & (df["sequencia_atual"] <= 0)).sum())

    urgentes = int(
        (
            (status_upper == "A")
            & (
                (df["data_prev_entrega_dt"] < hoje)
                | (df["nw_data_dt"].notna() & (df["nw_data_dt"] < hoje))
            )
        ).sum()
    )

    abertas_com_prioridade = int(
        (
            (status_upper == "A")
            & (df["Prioridade"].fillna("").astype(str).str.strip() != "")
        ).sum()
    )

    operador_series = df.loc[
        (status_upper == "A") & (df["desc_operador_atual"].fillna("").astype(str).str.strip() != ""),
        "desc_operador_atual"
    ].fillna("").astype(str).str.strip()

    qtd_operadores_abertos = int(operador_series.nunique()) if not operador_series.empty else 0

    grupo_510 = int((grupo == "510").sum())
    grupo_800 = int((grupo == "800").sum())
    grupo_801 = int((grupo == "801").sum())

    with st.expander("Indicadores do Laboratório"):
        st.title("Indicadores do Laboratório")
        r1 = st.columns(6, gap="large")
        r2 = st.columns(6, gap="large")

        with r1[0]:
            render_card_kpi("Número de SD abertas Mês", abertas_mes)
        with r1[1]:
            render_card_kpi("Número de SD fechadas Mês", fechadas_mes)
        with r1[2]:
            render_card_kpi("Número de SD abertas Total", total_abertas)
        with r1[3]:
            render_card_kpi("Número de SD fechadas Total", total_fechadas)
        with r1[4]:
            render_card_kpi("Número de SD em andamento", em_andamento)

        with r1[5]:
            render_card_kpi("Número de SD paradas", paradas)
        with r2[0]:
            render_card_kpi("Número de SD urgentes", urgentes)
        with r2[1]:
            render_card_kpi("SD abertas com Prioridade", abertas_com_prioridade)
        with r2[2]:
            render_card_kpi("Operadores com OF aberta", qtd_operadores_abertos)
        with r2[3]:
            render_card_kpi("Número de SD do Grupo 800", grupo_800)
        with r2[4]:
            render_card_kpi("Número de SD do Grupo 510", grupo_510)
        with r2[5]:
            render_card_kpi("Número de SD do Grupo 801", grupo_801)

    st.markdown("""
    <style>

    /* 🔥 aumenta fonte do st.table */
    div[data-testid="stTable"] table {
        font-size: 20px !important;
    }

    /* cabeçalho */
    div[data-testid="stTable"] th {
        font-size: 25px !important;
        font-weight: 700 !important;
    }

    /* célula */
    div[data-testid="stTable"] td {
        font-size: 25px !important;
    }

    /* coluna específica (2ª coluna) */
    div[data-testid="stTable"] td:nth-child(2) {
        font-weight: 800 !important;
        font-size: 24px !important;
        color: #0f172a;
    }
    td:nth-child(2) {
        background-color: #eef2ff;
        border-radius: 6px;
    }
    </style>
    """, unsafe_allow_html=True)

    with st.expander("OFs abertas por operador "):
        
        if operador_series.empty:
            st.info("Não há operador atual vinculado às OFs abertas.")
        else:
            df_operadores = (
                operador_series.value_counts()
                .reset_index()
            )
            df_operadores.columns = ["Operador", "Qtde OFs Abertas"]

            max_qtde = int(df_operadores["Qtde OFs Abertas"].max()) if not df_operadores.empty else 1
            largura_barra = 18

            st.markdown(
                """
                <style>
                .ranking-wrap {
                    border: 1px solid #dbe4f0;
                    border-radius: 10px;
                    background: linear-gradient(180deg, #FDF5E6 0%, #FDF5E6 100%);
                    padding: 12px 14px;
                    margin-top: 8px;
                }
                .ranking-row {
                    display: grid;
                    grid-template-columns: minmax(280px, 1.8fr) minmax(220px, 1.2fr) 70px;
                    gap: 12px;
                    align-items: center;
                    padding: 8px 4px;
                    border-bottom: 1px solid #e5e7eb;
                }
                .ranking-row:last-child {
                    border-bottom: none;
                }
                .ranking-operador {
                    font-size: 22px;
                    font-weight: 700;
                    color: #111827;
                    line-height: 1.2;
                }
                .ranking-barra {
                    font-family: Consolas, 'Courier New', monospace;
                    font-size: 24px;
                    font-weight: 700;
                    color: #000000;
                    letter-spacing: 1px;
                    white-space: pre;
                }
                .ranking-qtde {
                    font-size: 28px;
                    font-weight: 900;
                    color: #0f172a;
                    text-align: right;
                }
                </style>
                """,
                unsafe_allow_html=True,
            )

            linhas_html = ""

            for _, row in df_operadores.iterrows():
                operador = str(row["Operador"])
                qtde = int(row["Qtde OFs Abertas"])

                tamanho = max(1, round((qtde / max_qtde) * largura_barra)) if max_qtde > 0 else 1
                barra = "█" * tamanho

                linhas_html += f"""
                <div class="ranking-row">
                    <div class="ranking-operador">{operador}</div>
                    <div class="ranking-barra">{barra}</div>
                    <div class="ranking-qtde">{qtde}</div>
                </div>
                """

            st.markdown(
                f"""
                <div class="ranking-wrap">
                    {linhas_html}
                </div>
                """,
                unsafe_allow_html=True,
            )
                    
def validar_grid_para_salvar(df_para_salvar):
    erros = []

    for idx, row in df_para_salvar.iterrows():
        remover = bool(row.get("Remover", False))
        prioridade = str(row.get("Prioridade", "")).strip().upper()
        nw_data = str(row.get("Nw_Data", "")).strip()
        nro_of = str(row.get("Nro OF", "")).strip()
        codigo_produto = str(row.get("Código Produto", "")).strip()

        if remover:
            continue

        if prioridade != "" and not re.match(r"^[A-Z0-9]{1,3}$", prioridade):
            erros.append(
                f"Linha {idx + 1}: Prioridade inválida para OF {nro_of} / Produto {codigo_produto}. Use até 3 caracteres alfanuméricos."
            )

        if not validar_data_br(nw_data):
            erros.append(
                f"Linha {idx + 1}: Nw_Data inválida para OF {nro_of} / Produto {codigo_produto}. Use dd/mm/yyyy."
            )

    return erros


def render_cabecalho():
    col1, col2 = st.columns([8, 2])

    with col1:
        st.markdown(
            """
            <div style="margin-top: 30px;">
                <div style="font-size: 30px; font-weight: 800; color: #1f2937;">
                    🧪 Painel Laboratório
                </div>
                <div style="font-size: 13px; color: #64748b; margin-top: 4px;">
                    Desenvolvimento de Produto e Amostras
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with col2:
        if os.path.exists(ARQUIVO_LOGO_DIREITA):
            st.markdown(
                f"""
                <div style="display:flex; justify-content:flex-end; align-items:center; height:100%;">
                    <img src="data:image/png;base64,{imagem_para_base64(ARQUIVO_LOGO_DIREITA)}"
                         style="max-height:80px; margin-top:30px; margin-right:10px;">
                </div>
                """,
                unsafe_allow_html=True,
            )

def main():
    render_cabecalho()

    try:
        df = carregar_base_principal_com_controle()

        colunas_esperadas = [
            "data_abertura",
            "data_prev_entrega",
            "origem",
            "nro_of",
            "status_of",
            "codigo_produto",
            "desc_cliente",
            "Nw_Data",
            "codigo_original",
            "GP_codigo_grupo",
            "SGP_codigo_subgrupo",
            "Prioridade",
            "sequencia_atual",
            "data_final_apontamento",
            "desc_operacao_atual",
            "desc_operador_atual",
            "Responsavel",
            "operacoes_percorridas",
            "Alteracao",
        ]

        for col in colunas_esperadas:
            if col not in df.columns:
                df[col] = ""

        render_sidebar_filtros(df)
        render_upload_auditoria_sd()

        filtros = st.session_state["filtros_aplicados"]
        df_filtrado = montar_visao_filtrada(df, filtros)

        if st.session_state.get("df_sd_auditoria") is not None:
            df_filtrado = aplicar_auditoria_sd_no_dataframe(
                df_filtrado,
                st.session_state["df_sd_auditoria"]
            )
        else:
            for col in ["Auditoria SD", "Cliente SD", "Resultado SD", "Observações SD"]:
                if col not in df_filtrado.columns:
                    df_filtrado[col] = ""
        
        st.markdown("")
        render_metricas(df_filtrado)
        st.markdown("")
        

        df_exibicao = preparar_dataframe_exibicao(df_filtrado)
        df_grid_editado = render_grid(df_exibicao)
        st.session_state["df_grid_editado"] = df_grid_editado.copy()

        c1, c2 = st.columns([1.2, 5])

        with c1:
            if st.button("💾 Salvar banco_laboratorio.txt", use_container_width=True):
                df_para_salvar = df_grid_editado.copy()

                if "Prioridade" in df_para_salvar.columns:
                    df_para_salvar["Prioridade"] = (
                        df_para_salvar["Prioridade"].fillna("").astype(str).str.strip().str.upper()
                    )

                if "Nw_Data" in df_para_salvar.columns:
                    df_para_salvar["Nw_Data"] = (
                        df_para_salvar["Nw_Data"].fillna("").astype(str).str.strip()
                    )
                if "Responsavel" in df_para_salvar.columns:
                    df_para_salvar["Responsavel"] = (
                        df_para_salvar["Responsavel"].fillna("").astype(str).str.strip()
                    )
                if "Alteração" in df_para_salvar.columns and "Alteracao" not in df_para_salvar.columns:
                    df_para_salvar = df_para_salvar.rename(columns={"Alteração": "Alteracao"})

                if "Remover" in df_para_salvar.columns:
                    df_para_salvar["Remover"] = df_para_salvar["Remover"].fillna(False).astype(bool)

                erros = validar_grid_para_salvar(df_para_salvar)

                if erros:
                    st.error("Foram encontrados erros de preenchimento.")
                    for erro in erros:
                        st.write(f"- {erro}")
                else:
                    try:
                        salvar_banco_txt(df_para_salvar, ARQUIVO_BANCO)
                        reaplicar_banco_sem_reconsultar_base()
                        st.session_state["mensagem_salvo"] = (
                            f"Dados gravados com sucesso em {ARQUIVO_BANCO}."
                        )
                        st.rerun()
                    except Exception as e:
                        st.error(f"Erro ao salvar o arquivo {ARQUIVO_BANCO}: {e}")

        with c2:
            st.write("")

        if st.session_state["mensagem_salvo"]:
            st.info(st.session_state["mensagem_salvo"])

        with st.expander("Visualizar conteúdo atual do banco_laboratorio.txt"):
            df_banco_atual = carregar_banco_txt()
            if df_banco_atual.empty:
                st.write("O arquivo banco_laboratorio.txt ainda não possui registros.")
            else:
                df_banco_atual = df_banco_atual.copy()
                df_banco_atual["_ord_prioridade"] = df_banco_atual["Prioridade"].apply(prioridade_para_ordem)
                df_banco_atual["_ord_data"] = df_banco_atual["Nw_Data"].apply(converter_data_br_para_ordenacao)
                df_banco_atual = df_banco_atual.sort_values(
                    by=["_ord_prioridade", "_ord_data", "Nro_OF", "Codigo_Produto"],
                    ascending=[True, True, True, True],
                    na_position="last",
                ).drop(columns=["_ord_prioridade", "_ord_data"])
                st.dataframe(df_banco_atual, use_container_width=True, hide_index=True)

    except Exception as e:
        import traceback
        st.error(f"Erro ao carregar os dados: {e}")
        st.code(traceback.format_exc())


if __name__ == "__main__":
    main()
