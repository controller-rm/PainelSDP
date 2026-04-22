import pandas as pd
from database import get_connection
from services.laboratorio_apontamentos import enriquecer_com_apontamentos


def limpar_texto(valor) -> str:
    if pd.isna(valor):
        return ""
    return str(valor).strip()


def extrair_codigo_base(produto) -> str:
    """
    Usado somente para montar a chave de apontamento.
    Não usar para JOIN com PRODUTO.
    """
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


def montar_chave_of(nro_of, produto) -> str:
    """
    A chave de apontamento continua usando OF normalizada + produto base,
    para manter compatibilidade com APONTAMENTO.
    """
    nro_normalizado = normalizar_numero_of(nro_of)
    produto_base = extrair_codigo_base(produto)

    if nro_normalizado == "" and produto_base == "":
        return ""

    return f"{nro_normalizado}-{produto_base}"


def carregar_ofs_laboratorio(status_list=None) -> pd.DataFrame:
    """
    Carrega OFs da origem 997.
    Default: somente status A.
    """
    if not status_list:
        status_list = ["A"]

    status_list = [str(x).strip().upper() for x in status_list if str(x).strip() != ""]
    if not status_list:
        status_list = ["A"]

    status_sql = ", ".join([f"'{s}'" for s in status_list])

    conn = get_connection()

    sql_of = f"""
        SELECT
            codigo_filial,
            numero_da_of,
            data_abertura,
            data_fechamento,
            produto,
            desc_produto,
            qtde,
            qtde_reprovada,
            custo_reprovado,
            qtde_produzida,
            custo_mps,
            total_horas,
            custos_mob,
            custo_despesa,
            vlr_requisicoes,
            custo_unitario,
            status_of,
            data_prev_entrega,
            cod_cliente,
            desc_cliente,
            origem,
            desc_origem,
            nro_of
        FROM ORDEM_FABRIC
        WHERE TRIM(COALESCE(origem, '')) = '997'
          AND TRIM(COALESCE(status_of, '')) IN ({status_sql})
    """

    sql_produto = """
        SELECT
            codigo_produto_material,
            GP_codigo_grupo,
            SGP_codigo_subgrupo,
            codigo_original,
            data_inclusao
        FROM PRODUTO
    """

    try:
        df_of = pd.read_sql(sql_of, conn)
        df_prod = pd.read_sql(sql_produto, conn)
    finally:
        conn.close()

    return preparar_dataframe_laboratorio(df_of, df_prod)


def preparar_dataframe_laboratorio(df_of: pd.DataFrame, df_prod: pd.DataFrame) -> pd.DataFrame:
    df = df_of.copy()
    df_prod = df_prod.copy()

    # ======================================================
    # ORDEM_FABRIC
    # ======================================================
    if "produto" not in df.columns:
        raise KeyError("A coluna 'produto' não existe na tabela ORDEM_FABRIC.")

    if "nro_of" not in df.columns:
        raise KeyError("A coluna 'nro_of' não existe na tabela ORDEM_FABRIC.")

    df["produto"] = df["produto"].fillna("").astype(str).str.strip()
    df["nro_of"] = df["nro_of"].fillna("").astype(str).str.strip()
    df["origem"] = df.get("origem", "").fillna("").astype(str).str.strip()
    df["status_of"] = df.get("status_of", "").fillna("").astype(str).str.strip()
    df["desc_cliente"] = df.get("desc_cliente", "").fillna("").astype(str).str.strip()

    # codigo_produto exibido no painel = produto da ORDEM_FABRIC
    df["codigo_produto"] = df["produto"].fillna("").astype(str).str.strip()

    # chave para APONTAMENTO
    df["chave_of"] = df.apply(
        lambda row: montar_chave_of(
            nro_of=row["nro_of"],
            produto=row["produto"]
        ),
        axis=1
    )

    df["data_abertura"] = pd.to_datetime(df["data_abertura"], errors="coerce")
    df["data_prev_entrega"] = pd.to_datetime(df["data_prev_entrega"], errors="coerce")

    # ======================================================
    # PRODUTO
    # JOIN EXATO:
    # ORDEM_FABRIC.produto = PRODUTO.codigo_produto_material
    # ======================================================
    if "codigo_produto_material" not in df_prod.columns:
        raise KeyError("A coluna 'codigo_produto_material' não existe na tabela PRODUTO.")

    df_prod["codigo_produto_material"] = df_prod["codigo_produto_material"].fillna("").astype(str).str.strip()
    df_prod["GP_codigo_grupo"] = df_prod.get("GP_codigo_grupo", "").fillna("").astype(str).str.strip()
    df_prod["SGP_codigo_subgrupo"] = df_prod.get("SGP_codigo_subgrupo", "").fillna("").astype(str).str.strip()
    df_prod["codigo_original"] = df_prod.get("codigo_original", "").fillna("").astype(str).str.strip()

    df_prod = df_prod.drop_duplicates(subset=["codigo_produto_material"], keep="first")

    df = pd.merge(
        df,
        df_prod[
            [
                "codigo_produto_material",
                "GP_codigo_grupo",
                "SGP_codigo_subgrupo",
                "codigo_original",
                "data_inclusao",
            ]
        ],
        how="left",
        left_on="produto",
        right_on="codigo_produto_material"
    )

    df["codigo_original"] = df["codigo_original"].fillna("").astype(str).str.strip()
    df["GP_codigo_grupo"] = df["GP_codigo_grupo"].fillna("").astype(str).str.strip()
    df["SGP_codigo_subgrupo"] = df["SGP_codigo_subgrupo"].fillna("").astype(str).str.strip()

    # ======================================================
    # COLUNAS FINAIS
    # ======================================================
    df = enriquecer_com_apontamentos(df)

    colunas_finais = [
        "data_abertura",
        "data_prev_entrega",
        "origem",
        "nro_of",
        "status_of",
        "codigo_produto",
        "desc_cliente",
        "codigo_original",
        "GP_codigo_grupo",
        "SGP_codigo_subgrupo",
        "chave_of",
        "sequencia_atual",
        "desc_operacao_atual",
        "desc_operador_atual",
        "operacoes_percorridas",
        "data_final_apontamento",
    ]

    for col in colunas_finais:
        if col not in df.columns:
            df[col] = None

    df = df[colunas_finais].copy()

    df["sequencia_atual"] = pd.to_numeric(df["sequencia_atual"], errors="coerce").fillna(0).astype(int)

    df = df.sort_values(
        by=["data_prev_entrega", "data_abertura", "nro_of"],
        ascending=[True, True, True],
        na_position="last"
    ).reset_index(drop=True)

    return df