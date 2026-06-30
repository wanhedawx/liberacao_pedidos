import json
from datetime import datetime
from io import BytesIO

import pandas as pd
import streamlit as st
from sqlalchemy import create_engine, text
from sqlalchemy.exc import SQLAlchemyError


# =========================================================
# CONFIGURAÇÃO DA PÁGINA
# =========================================================

st.set_page_config(
    page_title="Liberação de Pedidos",
    page_icon="📦",
    layout="wide"
)


# =========================================================
# CONFIGURAÇÃO PADRÃO DA SUA PLANILHA
# Ajuste aqui se o nome da coluna na sua planilha for diferente.
# =========================================================

COLUNA_PEDIDO = "PEDIDO"
COLUNA_FORNECEDOR = "FORNECEDOR"
COLUNA_LOJA = "LOJA"
COLUNA_VALOR = "VALOR"
COLUNA_STATUS_PLANILHA = "STATUS"

# Colunas mínimas que a planilha precisa ter para o sistema funcionar.
COLUNAS_OBRIGATORIAS = [
    COLUNA_PEDIDO,
]

# Colunas que serão removidas automaticamente quando existirem na planilha.
# Exemplo: ["OBS", "USUARIO", "DATA_EXPORTACAO"]
COLUNAS_REMOVER_PADRAO = []

# Valores que serão considerados como liberados.
# Aqui deixei alguns exemplos comuns.
VALORES_STATUS_LIBERADO = {
    "L",
    "LIBERADO",
    "LIBERADA",
    "APROVADO",
    "APROVADA",
    "S",
    "SIM",
}


# =========================================================
# CSS VISUAL
# =========================================================

st.markdown(
    """
    <style>
        .block-container {
            padding-top: 1.5rem;
        }
        .titulo-principal {
            font-size: 34px;
            font-weight: 800;
            margin-bottom: 4px;
        }
        .subtitulo {
            color: #A0A0A0;
            font-size: 15px;
            margin-bottom: 20px;
        }
        .card-info {
            border: 1px solid rgba(255,255,255,0.10);
            border-radius: 14px;
            padding: 16px;
            background: rgba(255,255,255,0.03);
        }
    </style>
    """,
    unsafe_allow_html=True,
)


# =========================================================
# FUNÇÕES DE CONEXÃO COM O NEON / POSTGRESQL
# =========================================================


def get_postgres_url():
    """Lê a URL do PostgreSQL no .streamlit/secrets.toml."""
    try:
        return st.secrets["connections"]["postgresql"]["url"]
    except Exception:
        pass

    try:
        return st.secrets["DATABASE_URL"]
    except Exception:
        return None


@st.cache_resource(show_spinner=False)
def get_engine():
    url = get_postgres_url()
    if not url:
        return None

    return create_engine(url, pool_pre_ping=True)


def testar_conexao_postgres():
    engine = get_engine()
    if engine is None:
        return False, "String do Neon não encontrada no secrets.toml."

    try:
        with engine.begin() as conn:
            conn.execute(text("SELECT 1"))
        return True, "Conexão OK"
    except Exception as e:
        return False, str(e)


def criar_tabelas_se_nao_existirem():
    engine = get_engine()
    if engine is None:
        return False

    sql = """
    CREATE TABLE IF NOT EXISTS cargas_importacao (
        id BIGSERIAL PRIMARY KEY,
        nome_carga TEXT NOT NULL,
        nome_arquivo TEXT,
        usuario_importacao TEXT,
        data_importacao TIMESTAMP DEFAULT NOW(),
        qtd_pedidos INTEGER DEFAULT 0,
        qtd_liberados INTEGER DEFAULT 0,
        qtd_nao_liberados INTEGER DEFAULT 0
    );

    CREATE TABLE IF NOT EXISTS pedidos_importados (
        id BIGSERIAL PRIMARY KEY,
        carga_id BIGINT REFERENCES cargas_importacao(id) ON DELETE CASCADE,
        pedido TEXT NOT NULL,
        fornecedor TEXT,
        loja TEXT,
        valor NUMERIC(14,2),
        status_banco TEXT,
        cod_status_banco TEXT,
        dados_planilha JSONB,
        data_consulta TIMESTAMP DEFAULT NOW()
    );

    CREATE TABLE IF NOT EXISTS pedidos_liberados_historico (
        id BIGSERIAL PRIMARY KEY,
        pedido TEXT NOT NULL,
        fornecedor TEXT,
        loja TEXT,
        valor NUMERIC(14,2),
        status_banco TEXT,
        cod_status_banco TEXT,
        primeira_carga TEXT,
        ultima_carga TEXT,
        primeira_liberacao TIMESTAMP DEFAULT NOW(),
        ultima_atualizacao TIMESTAMP DEFAULT NOW(),
        dados_ultima_planilha JSONB,
        UNIQUE (pedido, loja)
    );
    """

    try:
        with engine.begin() as conn:
            conn.execute(text(sql))
        return True
    except SQLAlchemyError as e:
        st.error("Erro ao criar/verificar tabelas no Neon.")
        st.code(str(e))
        return False


# =========================================================
# FUNÇÕES DE ARQUIVO / TRATAMENTO
# =========================================================


def ler_arquivo(arquivo):
    nome = arquivo.name.lower()

    if nome.endswith(".xlsx") or nome.endswith(".xls"):
        return pd.read_excel(arquivo)

    if nome.endswith(".csv"):
        return pd.read_csv(arquivo, sep=None, engine="python")

    raise ValueError("Formato não suportado. Use .xlsx, .xls ou .csv")


def limpar_texto(valor):
    if pd.isna(valor):
        return ""
    return str(valor).strip()


def limpar_valor(valor):
    if pd.isna(valor) or valor == "":
        return None

    if isinstance(valor, (int, float)):
        return float(valor)

    valor_txt = str(valor).strip()
    valor_txt = valor_txt.replace("R$", "").replace(" ", "")

    # Trata padrão brasileiro: 1.234,56
    if "," in valor_txt:
        valor_txt = valor_txt.replace(".", "").replace(",", ".")

    try:
        return float(valor_txt)
    except Exception:
        return None


def normalizar_status(valor):
    status = limpar_texto(valor).upper()

    if not status:
        return "NÃO INFORMADO"

    if status in VALORES_STATUS_LIBERADO:
        return "LIBERADO"

    if status in {"B", "BLOQUEADO", "BLOQUEADA"}:
        return "BLOQUEADO"

    return "NÃO LIBERADO"


def gerar_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Pedidos")
    output.seek(0)
    return output


def row_json(row):
    dados = {}
    for k, v in row.items():
        if pd.isna(v):
            dados[k] = None
        elif isinstance(v, (pd.Timestamp, datetime)):
            dados[k] = str(v)
        else:
            dados[k] = v
    return json.dumps(dados, ensure_ascii=False, default=str)


def tratar_planilha(df_original, colunas_remover_extra=None):
    colunas_remover_extra = colunas_remover_extra or []

    df = df_original.copy()
    df.columns = [str(c).strip() for c in df.columns]

    colunas_faltando = [c for c in COLUNAS_OBRIGATORIAS if c not in df.columns]
    if colunas_faltando:
        return None, colunas_faltando

    colunas_remover = list(set(COLUNAS_REMOVER_PADRAO + colunas_remover_extra))
    df = df.drop(columns=colunas_remover, errors="ignore")

    df[COLUNA_PEDIDO] = df[COLUNA_PEDIDO].apply(limpar_texto)
    df = df[df[COLUNA_PEDIDO] != ""].copy()

    if COLUNA_FORNECEDOR not in df.columns:
        df[COLUNA_FORNECEDOR] = ""

    if COLUNA_LOJA not in df.columns:
        df[COLUNA_LOJA] = "GERAL"

    if COLUNA_VALOR not in df.columns:
        df[COLUNA_VALOR] = None

    df[COLUNA_FORNECEDOR] = df[COLUNA_FORNECEDOR].apply(limpar_texto)
    df[COLUNA_LOJA] = df[COLUNA_LOJA].apply(limpar_texto)
    df[COLUNA_LOJA] = df[COLUNA_LOJA].replace("", "GERAL")
    df[COLUNA_VALOR] = df[COLUNA_VALOR].apply(limpar_valor)

    return df, []


# =========================================================
# CONSULTA AO BANCO DA EMPRESA
# =========================================================


def consultar_status_pela_coluna_da_planilha(df):
    df_status = df[[COLUNA_PEDIDO]].drop_duplicates().copy()

    if COLUNA_STATUS_PLANILHA in df.columns:
        temp = df[[COLUNA_PEDIDO, COLUNA_STATUS_PLANILHA]].drop_duplicates().copy()
        temp["STATUS_BANCO"] = temp[COLUNA_STATUS_PLANILHA].apply(normalizar_status)
        temp["COD_STATUS_BANCO"] = temp[COLUNA_STATUS_PLANILHA].apply(limpar_texto)
        return temp[[COLUNA_PEDIDO, "STATUS_BANCO", "COD_STATUS_BANCO"]]

    df_status["STATUS_BANCO"] = "PENDENTE CONSULTA"
    df_status["COD_STATUS_BANCO"] = ""
    return df_status


def consultar_status_por_arquivo_status(arquivo_status):
    df_status = ler_arquivo(arquivo_status)
    df_status.columns = [str(c).strip() for c in df_status.columns]

    if COLUNA_PEDIDO not in df_status.columns:
        st.error(f"O arquivo de status precisa ter a coluna {COLUNA_PEDIDO}.")
        return pd.DataFrame()

    if "STATUS_BANCO" not in df_status.columns:
        st.error("O arquivo de status precisa ter a coluna STATUS_BANCO.")
        return pd.DataFrame()

    df_status[COLUNA_PEDIDO] = df_status[COLUNA_PEDIDO].apply(limpar_texto)
    df_status["STATUS_BANCO"] = df_status["STATUS_BANCO"].apply(normalizar_status)

    if "COD_STATUS_BANCO" not in df_status.columns:
        df_status["COD_STATUS_BANCO"] = df_status["STATUS_BANCO"]

    return df_status[[COLUNA_PEDIDO, "STATUS_BANCO", "COD_STATUS_BANCO"]].drop_duplicates()


def consultar_status_oracle_empresa(pedidos):
    """
    Consulta o banco da empresa.

    Para usar, configure no .streamlit/secrets.toml:

    [empresa_oracle]
    enabled = true
    user = "SEU_USUARIO"
    password = "SUA_SENHA"
    dsn = "host:1521/service"
    query_status = '''
        SELECT
            C7_NUM AS PEDIDO,
            CASE
                WHEN C7_CONAPRO = 'L' THEN 'LIBERADO'
                WHEN C7_CONAPRO = 'B' THEN 'BLOQUEADO'
                ELSE 'NÃO LIBERADO'
            END AS STATUS_BANCO,
            C7_CONAPRO AS COD_STATUS_BANCO
        FROM TOTVS.SC7010
        WHERE D_E_L_E_T_ = ' '
          AND C7_NUM IN ({placeholders})
    '''
    """

    try:
        cfg = st.secrets["empresa_oracle"]
        enabled = bool(cfg.get("enabled", False))
    except Exception:
        enabled = False
        cfg = None

    if not enabled:
        st.warning("Consulta ao banco da empresa ainda não configurada no secrets.toml.")
        return pd.DataFrame()

    try:
        import oracledb
    except ImportError:
        st.error("Biblioteca oracledb não instalada. Adicione oracledb no requirements.txt.")
        return pd.DataFrame()

    query_modelo = cfg.get("query_status", "")
    if "{placeholders}" not in query_modelo:
        st.error("A query_status precisa conter {placeholders} no IN dos pedidos.")
        return pd.DataFrame()

    usuario = cfg.get("user")
    senha = cfg.get("password")
    dsn = cfg.get("dsn")

    if not usuario or not senha or not dsn:
        st.error("Configuração empresa_oracle incompleta no secrets.toml.")
        return pd.DataFrame()

    pedidos = [limpar_texto(p) for p in pedidos if limpar_texto(p)]
    resultados = []
    tamanho_lote = 900

    try:
        conn = oracledb.connect(user=usuario, password=senha, dsn=dsn)

        for i in range(0, len(pedidos), tamanho_lote):
            lote = pedidos[i:i + tamanho_lote]
            binds = {}
            placeholders = []

            for idx, pedido in enumerate(lote):
                chave = f"p{idx}"
                placeholders.append(f":{chave}")
                binds[chave] = pedido

            sql = query_modelo.replace("{placeholders}", ", ".join(placeholders))
            df_lote = pd.read_sql(sql, conn, params=binds)
            resultados.append(df_lote)

        conn.close()

    except Exception as e:
        st.error("Erro ao consultar o banco da empresa.")
        st.code(str(e))
        return pd.DataFrame()

    if not resultados:
        return pd.DataFrame()

    df_status = pd.concat(resultados, ignore_index=True)
    df_status.columns = [str(c).strip().upper() for c in df_status.columns]

    if "PEDIDO" not in df_status.columns or "STATUS_BANCO" not in df_status.columns:
        st.error("A consulta precisa retornar as colunas PEDIDO e STATUS_BANCO.")
        return pd.DataFrame()

    if "COD_STATUS_BANCO" not in df_status.columns:
        df_status["COD_STATUS_BANCO"] = df_status["STATUS_BANCO"]

    df_status["PEDIDO"] = df_status["PEDIDO"].apply(limpar_texto)
    df_status["STATUS_BANCO"] = df_status["STATUS_BANCO"].apply(normalizar_status)
    df_status["COD_STATUS_BANCO"] = df_status["COD_STATUS_BANCO"].apply(limpar_texto)

    return df_status[["PEDIDO", "STATUS_BANCO", "COD_STATUS_BANCO"]].drop_duplicates()


# =========================================================
# SALVAR NO NEON
# =========================================================


def salvar_carga_no_neon(df_final, nome_carga, nome_arquivo, usuario_importacao):
    engine = get_engine()
    if engine is None:
        st.error("Neon/PostgreSQL não configurado. Verifique o secrets.toml.")
        return False

    qtd_pedidos = len(df_final)
    qtd_liberados = int((df_final["STATUS_BANCO"] == "LIBERADO").sum())
    qtd_nao_liberados = qtd_pedidos - qtd_liberados

    try:
        with engine.begin() as conn:
            carga_id = conn.execute(
                text(
                    """
                    INSERT INTO cargas_importacao
                    (nome_carga, nome_arquivo, usuario_importacao, qtd_pedidos, qtd_liberados, qtd_nao_liberados)
                    VALUES (:nome_carga, :nome_arquivo, :usuario_importacao, :qtd_pedidos, :qtd_liberados, :qtd_nao_liberados)
                    RETURNING id
                    """
                ),
                {
                    "nome_carga": nome_carga,
                    "nome_arquivo": nome_arquivo,
                    "usuario_importacao": usuario_importacao,
                    "qtd_pedidos": qtd_pedidos,
                    "qtd_liberados": qtd_liberados,
                    "qtd_nao_liberados": qtd_nao_liberados,
                },
            ).scalar_one()

            for _, row in df_final.iterrows():
                pedido = limpar_texto(row.get(COLUNA_PEDIDO))
                fornecedor = limpar_texto(row.get(COLUNA_FORNECEDOR))
                loja = limpar_texto(row.get(COLUNA_LOJA)) or "GERAL"
                valor = limpar_valor(row.get(COLUNA_VALOR))
                status_banco = limpar_texto(row.get("STATUS_BANCO"))
                cod_status_banco = limpar_texto(row.get("COD_STATUS_BANCO"))
                dados = row_json(row)

                conn.execute(
                    text(
                        """
                        INSERT INTO pedidos_importados
                        (carga_id, pedido, fornecedor, loja, valor, status_banco, cod_status_banco, dados_planilha)
                        VALUES
                        (:carga_id, :pedido, :fornecedor, :loja, :valor, :status_banco, :cod_status_banco, CAST(:dados_planilha AS JSONB))
                        """
                    ),
                    {
                        "carga_id": carga_id,
                        "pedido": pedido,
                        "fornecedor": fornecedor,
                        "loja": loja,
                        "valor": valor,
                        "status_banco": status_banco,
                        "cod_status_banco": cod_status_banco,
                        "dados_planilha": dados,
                    },
                )

                if status_banco == "LIBERADO":
                    conn.execute(
                        text(
                            """
                            INSERT INTO pedidos_liberados_historico
                            (pedido, fornecedor, loja, valor, status_banco, cod_status_banco, primeira_carga, ultima_carga, dados_ultima_planilha)
                            VALUES
                            (:pedido, :fornecedor, :loja, :valor, :status_banco, :cod_status_banco, :nome_carga, :nome_carga, CAST(:dados_planilha AS JSONB))
                            ON CONFLICT (pedido, loja)
                            DO UPDATE SET
                                fornecedor = EXCLUDED.fornecedor,
                                valor = EXCLUDED.valor,
                                status_banco = EXCLUDED.status_banco,
                                cod_status_banco = EXCLUDED.cod_status_banco,
                                ultima_carga = EXCLUDED.ultima_carga,
                                ultima_atualizacao = NOW(),
                                dados_ultima_planilha = EXCLUDED.dados_ultima_planilha
                            """
                        ),
                        {
                            "pedido": pedido,
                            "fornecedor": fornecedor,
                            "loja": loja,
                            "valor": valor,
                            "status_banco": status_banco,
                            "cod_status_banco": cod_status_banco,
                            "nome_carga": nome_carga,
                            "dados_planilha": dados,
                        },
                    )

        return True

    except Exception as e:
        st.error("Erro ao salvar no Neon/PostgreSQL.")
        st.code(str(e))
        return False


# =========================================================
# CONSULTAS HISTÓRICO
# =========================================================


def carregar_cargas():
    engine = get_engine()
    if engine is None:
        return pd.DataFrame()

    try:
        with engine.begin() as conn:
            return pd.read_sql(
                """
                SELECT
                    id,
                    nome_carga,
                    nome_arquivo,
                    usuario_importacao,
                    data_importacao,
                    qtd_pedidos,
                    qtd_liberados,
                    qtd_nao_liberados
                FROM cargas_importacao
                ORDER BY data_importacao DESC
                """,
                conn,
            )
    except Exception:
        return pd.DataFrame()


def carregar_historico_liberados():
    engine = get_engine()
    if engine is None:
        return pd.DataFrame()

    try:
        with engine.begin() as conn:
            return pd.read_sql(
                """
                SELECT
                    pedido,
                    fornecedor,
                    loja,
                    valor,
                    status_banco,
                    cod_status_banco,
                    primeira_carga,
                    ultima_carga,
                    primeira_liberacao,
                    ultima_atualizacao
                FROM pedidos_liberados_historico
                ORDER BY ultima_atualizacao DESC
                """,
                conn,
            )
    except Exception:
        return pd.DataFrame()


# =========================================================
# APP
# =========================================================

st.markdown('<div class="titulo-principal">📦 Liberação de Pedidos</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="subtitulo">Importe a planilha padrão, consulte o status e mantenha o histórico dos pedidos liberados no Neon/PostgreSQL.</div>',
    unsafe_allow_html=True,
)

ok_conexao, msg_conexao = testar_conexao_postgres()

with st.sidebar:
    st.header("Configuração")

    if ok_conexao:
        st.success("Neon conectado")
        criar_tabelas_se_nao_existirem()
    else:
        st.error("Neon não conectado")
        st.caption(msg_conexao)

    usuario_importacao = st.text_input("Usuário", value="Vitoria")

    metodo_status = st.selectbox(
        "Como buscar o status?",
        options=[
            "Status pela coluna da planilha",
            "Arquivo complementar de status",
            "Banco da empresa / Oracle",
        ],
        index=0,
    )

    modo_avancado = st.checkbox("Modo avançado: remover colunas manualmente", value=False)

aba_importar, aba_historico, aba_config = st.tabs([
    "📥 Importar pedidos",
    "📚 Histórico",
    "⚙️ Configuração esperada",
])


with aba_importar:
    c1, c2 = st.columns([2, 1])

    with c1:
        nome_carga = st.text_input(
            "Nome da carga",
            placeholder="Exemplo: Pedidos liberados - Julho",
        )

    with c2:
        st.write("")
        st.write("")
        st.caption("Esse nome será salvo no histórico do Neon.")

    arquivo = st.file_uploader(
        "Selecione a planilha padrão",
        type=["xlsx", "xls", "csv"],
    )

    arquivo_status = None
    if metodo_status == "Arquivo complementar de status":
        arquivo_status = st.file_uploader(
            "Selecione o arquivo complementar de status",
            type=["xlsx", "xls", "csv"],
            key="arquivo_status",
        )

    if arquivo is not None:
        try:
            df_original = ler_arquivo(arquivo)
            df_original.columns = [str(c).strip() for c in df_original.columns]

            st.subheader("Colunas encontradas")
            st.dataframe(pd.DataFrame({"colunas": list(df_original.columns)}), use_container_width=True, hide_index=True)

            colunas_remover_extra = []
            if modo_avancado:
                colunas_remover_extra = st.multiselect(
                    "Colunas adicionais para remover nesta carga",
                    options=list(df_original.columns),
                    default=[],
                )

            df_tratado, colunas_faltando = tratar_planilha(df_original, colunas_remover_extra)

            if colunas_faltando:
                st.error("A planilha não está no padrão esperado.")
                st.write("Colunas faltando:")
                st.code("\n".join(colunas_faltando))
                st.stop()

            # Busca status
            df_status = pd.DataFrame()

            if metodo_status == "Status pela coluna da planilha":
                df_status = consultar_status_pela_coluna_da_planilha(df_tratado)

            elif metodo_status == "Arquivo complementar de status":
                if arquivo_status is None:
                    st.warning("Envie o arquivo complementar de status para continuar.")
                    st.stop()
                df_status = consultar_status_por_arquivo_status(arquivo_status)

            elif metodo_status == "Banco da empresa / Oracle":
                pedidos = df_tratado[COLUNA_PEDIDO].dropna().astype(str).unique().tolist()
                with st.spinner("Consultando banco da empresa..."):
                    df_status = consultar_status_oracle_empresa(pedidos)

            if df_status.empty:
                st.warning("Nenhum status foi encontrado. Os pedidos ficarão como PENDENTE CONSULTA.")
                df_status = df_tratado[[COLUNA_PEDIDO]].drop_duplicates().copy()
                df_status["STATUS_BANCO"] = "PENDENTE CONSULTA"
                df_status["COD_STATUS_BANCO"] = ""

            df_status.columns = [str(c).strip().upper() for c in df_status.columns]
            df_status["PEDIDO"] = df_status["PEDIDO"].apply(limpar_texto)

            df_final = df_tratado.merge(
                df_status,
                left_on=COLUNA_PEDIDO,
                right_on="PEDIDO",
                how="left",
            )

            if "PEDIDO_y" in df_final.columns:
                df_final = df_final.drop(columns=["PEDIDO_y"])

            if "PEDIDO_x" in df_final.columns:
                df_final = df_final.rename(columns={"PEDIDO_x": COLUNA_PEDIDO})

            df_final["STATUS_BANCO"] = df_final["STATUS_BANCO"].fillna("NÃO ENCONTRADO")
            df_final["COD_STATUS_BANCO"] = df_final["COD_STATUS_BANCO"].fillna("")

            total_pedidos = len(df_final)
            total_liberados = int((df_final["STATUS_BANCO"] == "LIBERADO").sum())
            total_nao_liberados = total_pedidos - total_liberados
            total_valor = df_final[COLUNA_VALOR].fillna(0).sum() if COLUNA_VALOR in df_final.columns else 0

            st.subheader("Resumo da carga")
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Total de pedidos", total_pedidos)
            m2.metric("Liberados", total_liberados)
            m3.metric("Não liberados / pendentes", total_nao_liberados)
            m4.metric("Valor total", f"R$ {total_valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))

            st.subheader("Resultado final")

            st.dataframe(
                df_final,
                use_container_width=True,
                hide_index=True,
            )

            col_baixar, col_salvar = st.columns([1, 1])

            with col_baixar:
                excel = gerar_excel(df_final)
                st.download_button(
                    "📥 Baixar Excel final",
                    data=excel,
                    file_name=f"{nome_carga or 'pedidos'}_final.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

            with col_salvar:
                if st.button("💾 Salvar carga no Neon", type="primary", use_container_width=True):
                    if not nome_carga:
                        st.warning("Informe o nome da carga antes de salvar.")
                    elif not ok_conexao:
                        st.error("Neon não conectado. Verifique o secrets.toml.")
                    else:
                        sucesso = salvar_carga_no_neon(
                            df_final=df_final,
                            nome_carga=nome_carga,
                            nome_arquivo=arquivo.name,
                            usuario_importacao=usuario_importacao,
                        )
                        if sucesso:
                            st.success("Carga salva no Neon e histórico de liberados atualizado.")
                            st.cache_data.clear()

        except Exception as e:
            st.error("Erro ao processar a planilha.")
            st.code(str(e))


with aba_historico:
    if not ok_conexao:
        st.warning("Conecte o Neon para visualizar o histórico.")
    else:
        st.subheader("Cargas importadas")
        df_cargas = carregar_cargas()

        if df_cargas.empty:
            st.info("Nenhuma carga importada ainda.")
        else:
            st.dataframe(df_cargas, use_container_width=True, hide_index=True)

        st.divider()
        st.subheader("Pedidos liberados mantidos no histórico")
        df_hist = carregar_historico_liberados()

        if df_hist.empty:
            st.info("Nenhum pedido liberado salvo ainda.")
        else:
            f1, f2, f3 = st.columns(3)
            filtro_pedido = f1.text_input("Filtrar pedido")
            filtro_fornecedor = f2.text_input("Filtrar fornecedor")
            filtro_loja = f3.text_input("Filtrar loja")

            df_view = df_hist.copy()

            if filtro_pedido:
                df_view = df_view[df_view["pedido"].astype(str).str.contains(filtro_pedido, case=False, na=False)]

            if filtro_fornecedor:
                df_view = df_view[df_view["fornecedor"].astype(str).str.contains(filtro_fornecedor, case=False, na=False)]

            if filtro_loja:
                df_view = df_view[df_view["loja"].astype(str).str.contains(filtro_loja, case=False, na=False)]

            st.dataframe(df_view, use_container_width=True, hide_index=True)

            excel_hist = gerar_excel(df_view)
            st.download_button(
                "📥 Baixar histórico de liberados",
                data=excel_hist,
                file_name="historico_pedidos_liberados.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )


with aba_config:
    st.subheader("Configuração esperada da planilha")

    st.markdown(
        f"""
        O app está configurado para procurar estas colunas:

        - Pedido: `{COLUNA_PEDIDO}`
        - Fornecedor: `{COLUNA_FORNECEDOR}`
        - Loja: `{COLUNA_LOJA}`
        - Valor: `{COLUNA_VALOR}`
        - Status na planilha, se existir: `{COLUNA_STATUS_PLANILHA}`
        """
    )

    st.warning(
        "Se a sua planilha usa outro nome de coluna, altere as variáveis no começo do arquivo app_liberacao_pedidos.py."
    )

    st.subheader("Exemplo do secrets.toml")
    st.code(
        '''[connections.postgresql]
url = "postgresql://usuario:senha@host.neon.tech/neondb?sslmode=require"

# Opcional: somente quando a empresa liberar acesso ao Oracle/TOTVS
[empresa_oracle]
enabled = false
user = "SEU_USUARIO"
password = "SUA_SENHA"
dsn = "host:1521/service"
query_status = """
SELECT
    C7_NUM AS PEDIDO,
    CASE
        WHEN C7_CONAPRO = 'L' THEN 'LIBERADO'
        WHEN C7_CONAPRO = 'B' THEN 'BLOQUEADO'
        ELSE 'NÃO LIBERADO'
    END AS STATUS_BANCO,
    C7_CONAPRO AS COD_STATUS_BANCO
FROM TOTVS.SC7010
WHERE D_E_L_E_T_ = ' '
  AND C7_NUM IN ({placeholders})
"""
''',
        language="toml",
    )

    st.subheader("requirements.txt")
    st.code(
        """streamlit
pandas
openpyxl
sqlalchemy
psycopg2-binary
oracledb
""",
        language="txt",
    )
