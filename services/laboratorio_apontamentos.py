import pandas as pd
from database import get_connection


def limpar_texto(valor) -> str:
    if pd.isna(valor):
        return ""
    return str(valor).strip()


def extrair_codigo_base(produto) -> str:
    texto = limpar_texto(produto)
    if texto == "":
        return ""

    texto = texto.split(" ")[0].strip()

    if "." in texto:
        texto = texto.split(".", 1)[0].strip()

    return texto


def normalizar_numero_of(numero_of) -> str:
    texto = limpar_texto(numero_of).replace(" ", "")
    if texto == "":
        return ""

    partes = texto.split("-")

    if len(partes) == 3:
        p1, p2, p3 = partes

        try:
            p1 = str(int(float(p1))).zfill(6)
        except Exception:
            p1 = p1.zfill(6)

        try:
            p2 = str(int(float(p2))).zfill(2)
        except Exception:
            p2 = p2.zfill(2)

        try:
            p3 = str(int(float(p3))).zfill(3)
        except Exception:
            p3 = p3.zfill(3)

        return f"{p1}-{p2}-{p3}"

    return texto


def montar_chave_apontamento(numero_of, produto) -> str:
    numero_normalizado = normalizar_numero_of(numero_of)
    produto_base = extrair_codigo_base(produto)

    if not numero_normalizado and not produto_base:
        return ""

    return f"{numero_normalizado}-{produto_base}"


def carregar_apontamentos() -> pd.DataFrame:
    conn = get_connection()

    sql = """
        SELECT
            numero_of,
            produto,
            sequencia_of,
            desc_operacao,
            desc_operador,
            data_final
        FROM APONTAMENTO
    """

    try:
        df = pd.read_sql(sql, conn)
    finally:
        conn.close()

    return df


def consolidar_apontamentos(df_apontamento: pd.DataFrame) -> pd.DataFrame:
    df = df_apontamento.copy()

    colunas_minimas = ["numero_of", "produto", "sequencia_of", "desc_operacao"]
    for col in colunas_minimas:
        if col not in df.columns:
            raise KeyError(f"A coluna '{col}' não existe na tabela APONTAMENTO.")

    if "desc_operador" not in df.columns:
        df["desc_operador"] = ""
    
    if "data_final" not in df.columns:
        df["data_final"] = None

    df["numero_of"] = df["numero_of"].fillna("").astype(str).str.strip()
    df["produto"] = df["produto"].fillna("").astype(str).str.strip()
    df["desc_operacao"] = df["desc_operacao"].fillna("").astype(str).str.strip()
    df["desc_operador"] = df["desc_operador"].fillna("").astype(str).str.strip()
    df["sequencia_of"] = pd.to_numeric(df["sequencia_of"], errors="coerce").fillna(0).astype(int)
    df["data_final"] = pd.to_datetime(df["data_final"], errors="coerce")

    df["chave_of"] = df.apply(
        lambda row: montar_chave_apontamento(
            numero_of=row["numero_of"],
            produto=row["produto"]
        ),
        axis=1
    )

    df = df[df["chave_of"] != ""].copy()

    df = df.sort_values(
        by=["chave_of", "sequencia_of"],
        ascending=[True, True]
    ).reset_index(drop=True)

    def montar_fluxo_operacoes(subdf: pd.DataFrame) -> str:
        pares = []
        vistos = set()

        for _, row in subdf.iterrows():
            seq = int(row["sequencia_of"]) if pd.notna(row["sequencia_of"]) else 0
            oper = limpar_texto(row["desc_operacao"])
            if not oper:
                continue

            chave = (seq, oper)
            if chave not in vistos:
                vistos.add(chave)
                pares.append(f"{seq} - {oper}")

        return " → ".join(pares)

    df_fluxo = (
        df.groupby("chave_of", as_index=False)
        .apply(lambda g: pd.Series({
            "operacoes_percorridas": montar_fluxo_operacoes(g)
        }))
        .reset_index(drop=True)
    )

    idx_ultimos = df.groupby("chave_of")["sequencia_of"].idxmax()
    df_ultimos = df.loc[
        idx_ultimos,
        ["chave_of", "sequencia_of", "desc_operacao", "desc_operador", "data_final"]
    ].copy()

    df_ultimos = df_ultimos.rename(columns={
        "sequencia_of": "sequencia_atual",
        "desc_operacao": "desc_operacao_atual",
        "desc_operador": "desc_operador_atual",
        "data_final": "data_final_apontamento",
    })

    df_final = pd.merge(
        df_fluxo,
        df_ultimos,
        how="left",
        on="chave_of"
    )
    if "data_final_apontamento" not in df_final.columns:
        df_final["data_final_apontamento"] = pd.NaT

    df_final["data_final_apontamento"] = pd.to_datetime(
        df_final["data_final_apontamento"],
        errors="coerce"
    )
    df_final["sequencia_atual"] = pd.to_numeric(df_final["sequencia_atual"], errors="coerce").fillna(0).astype(int)
    df_final["desc_operacao_atual"] = df_final["desc_operacao_atual"].fillna("").astype(str)
    df_final["desc_operador_atual"] = df_final["desc_operador_atual"].fillna("").astype(str)
    df_final["operacoes_percorridas"] = df_final["operacoes_percorridas"].fillna("").astype(str)

    return df_final


def enriquecer_com_apontamentos(df_base: pd.DataFrame) -> pd.DataFrame:
    df = df_base.copy()

    if "chave_of" not in df.columns:
        df["sequencia_atual"] = 0
        df["desc_operacao_atual"] = ""
        df["desc_operador_atual"] = ""
        df["operacoes_percorridas"] = ""
        return df

    df_apontamento = carregar_apontamentos()
    df_apontamento = consolidar_apontamentos(df_apontamento)

    df = pd.merge(
        df,
    df_apontamento[
        [
            "chave_of",
            "sequencia_atual",
            "desc_operacao_atual",
            "desc_operador_atual",
            "operacoes_percorridas",
            "data_final_apontamento",
        ]
    ],
        how="left",
        on="chave_of"
    )
    
    if "data_final_apontamento" not in df.columns:
        df["data_final_apontamento"] = pd.NaT

    df["data_final_apontamento"] = pd.to_datetime(
        df["data_final_apontamento"],
        errors="coerce"
    )

    if "sequencia_atual" not in df.columns:
        df["sequencia_atual"] = 0

    if "desc_operacao_atual" not in df.columns:
        df["desc_operacao_atual"] = ""

    if "desc_operador_atual" not in df.columns:
        df["desc_operador_atual"] = ""

    if "operacoes_percorridas" not in df.columns:
        df["operacoes_percorridas"] = ""



    df["sequencia_atual"] = pd.to_numeric(df["sequencia_atual"], errors="coerce").fillna(0).astype(int)
    df["desc_operacao_atual"] = df["desc_operacao_atual"].fillna("").astype(str)
    df["desc_operador_atual"] = df["desc_operador_atual"].fillna("").astype(str)
    df["operacoes_percorridas"] = df["operacoes_percorridas"].fillna("").astype(str)

    return df