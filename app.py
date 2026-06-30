import json
import re
from datetime import date, datetime, timedelta
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
    layout="wide",
)


# =========================================================
# NOMES PADRÃO ESPERADOS NO RETORNO DA QUERY DA EMPRESA
# A query da empresa precisa retornar pelo menos: PEDIDO
# Recomendo retornar também: FORNECEDOR, LOJA, VALOR, DEPARTAMENTO,
# NOME_ARQUIVO, DATA_IMPORTACAO, STATUS_BANCO, COD_STATUS_BANCO
# =========================================================

COLUNA_PEDIDO = "PEDIDO"
COLUNA_FORNECEDOR = "FORNECEDOR"
COLUNA_LOJA = "LOJA"
COLUNA_VALOR = "VALOR"
COLUNA_DEPARTAMENTO = "DEPARTAMENTO"
COLUNA_NOME_ARQUIVO = "NOME_ARQUIVO"
COLUNA_DATA_IMPORTACAO = "DATA_IMPORTACAO"

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
# ANALISTAS E DEPARTAMENTOS FIXOS
# =========================================================

ANALISTAS = {
    "Cleviton": [
        "Portas e Janelas", "Ferramentas", "Ferragens", "Automotivos"
    ],
    "Alec": [
        "Eletrica", "Iluminacao", "Hidraulica"
    ],
    "Jonatas": [
        "Moveis e Colchoes", "Decoracao", "Cama Mesa e Banho", "Lazer",
        "Casa e UD", "Jardim",
    ],
    "Beatriz": [
        "Eletro", "Tecnologia", "Climatizacao"
    ],
    "Ruan": [
        "Tintas", "Organizacao da Casa"
    ],
    "Jessica": [
        "Materiais de Construcao", "Banho e Cozinha"
    ],
    "Rose": [
        "Pisos e Revestimento"
    ],
}

# =========================================================
# CSS VISUAL
# =========================================================

st.markdown(
    """
    <style>
        .block-container { padding-top: 1.5rem; }
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
# FUNÇÕES GERAIS
# =========================================================


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
        elif isinstance(v, (pd.Timestamp, datetime, date)):
            dados[k] = str(v)
        else:
            dados[k] = v
    return json.dumps(dados, ensure_ascii=False, default=str)



def formatar_moeda(valor):
    try:
        return f"R$ {float(valor):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return "R$ 0,00"



def detectar_binds(sql):
    # Captura binds Oracle no formato :data_ini, :data_fim, :departamento, :nome_arquivo
    return set(re.findall(r":([A-Za-z_][A-Za-z0-9_]*)", sql or ""))



def filtrar_params_sql(sql, params):
    binds = detectar_binds(sql)
    return {k: v for k, v in params.items() if k in binds}


# =========================================================
# CONEXÃO COM NEON / POSTGRESQL DO APP
# =========================================================


def get_postgres_url():
    try:
        return st.secrets["connections"]["postgresql"]["url"]
    except Exception:
        pass

    try:
        return st.secrets["DATABASE_URL"]
    except Exception:
        return None


@st.cache_resource(show_spinner=False)
def get_engine_neon():
    url = get_postgres_url()
    if not url:
        return None
    return create_engine(url, pool_pre_ping=True)



def testar_conexao_neon():
    engine = get_engine_neon()
    if engine is None:
        return False, "String do Neon não encontrada nos Secrets."

    try:
        with engine.begin() as conn:
            conn.execute(text("SELECT 1"))
        return True, "Conexão OK"
    except Exception as e:
        return False, str(e)



def criar_tabelas_neon():
    engine = get_engine_neon()
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

    ALTER TABLE cargas_importacao
    ADD COLUMN IF NOT EXISTS departamento TEXT;

    ALTER TABLE cargas_importacao
    ADD COLUMN IF NOT EXISTS filtro_data_ini DATE;

    ALTER TABLE cargas_importacao
    ADD COLUMN IF NOT EXISTS filtro_data_fim DATE;

    ALTER TABLE pedidos_importados
    ADD COLUMN IF NOT EXISTS departamento TEXT;

    ALTER TABLE pedidos_importados
    ADD COLUMN IF NOT EXISTS nome_arquivo_origem TEXT;

    ALTER TABLE pedidos_liberados_historico
    ADD COLUMN IF NOT EXISTS departamento TEXT;
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
# CONEXÃO COM BANCO DA EMPRESA / ORACLE
# =========================================================


def get_empresa_cfg():
    try:
        return st.secrets["empresa_oracle"]
    except Exception:
        return None



def empresa_habilitada():
    cfg = get_empresa_cfg()
    if cfg is None:
        return False, "Configuração [empresa_oracle] não encontrada nos Secrets."

    if not bool(cfg.get("enabled", False)):
        return False, "Banco da empresa está com enabled = false nos Secrets."

    campos = ["user", "password", "dsn", "query_arquivos", "query_pedidos"]
    faltando = [c for c in campos if not cfg.get(c)]
    if faltando:
        return False, f"Configuração incompleta em [empresa_oracle]. Faltando: {', '.join(faltando)}"

    return True, "Configuração OK"


@st.cache_resource(show_spinner=False)
def get_oracle_connection(user, password, dsn):
    import oracledb
    return oracledb.connect(user=user, password=password, dsn=dsn)



def testar_conexao_empresa():
    ok, msg = empresa_habilitada()
    if not ok:
        return False, msg

    try:
        import oracledb  # noqa: F401
    except ImportError:
        return False, "Biblioteca oracledb não instalada. Confira o requirements.txt."

    cfg = get_empresa_cfg()

    try:
        conn = get_oracle_connection(cfg.get("user"), cfg.get("password"), cfg.get("dsn"))
        cursor = conn.cursor()
        cursor.execute("SELECT 1 FROM DUAL")
        cursor.fetchone()
        cursor.close()
        return True, "Conexão OK"
    except Exception as e:
        return False, str(e)



def executar_query_empresa(sql, params):
    cfg = get_empresa_cfg()
    if cfg is None:
        st.error("Configuração [empresa_oracle] não encontrada nos Secrets.")
        return pd.DataFrame()

    try:
        conn = get_oracle_connection(cfg.get("user"), cfg.get("password"), cfg.get("dsn"))
        params_usados = filtrar_params_sql(sql, params)
        df = pd.read_sql(sql, conn, params=params_usados)
        df.columns = [str(c).strip().upper() for c in df.columns]
        return df
    except Exception as e:
        st.error("Erro ao executar consulta no banco da empresa.")
        st.code(str(e))
        return pd.DataFrame()



def carregar_departamentos(data_ini, data_fim):
    cfg = get_empresa_cfg()
    if cfg is None or not cfg.get("query_departamentos"):
        return []

    params = {
        "data_ini": data_ini.strftime("%Y-%m-%d"),
        "data_fim": data_fim.strftime("%Y-%m-%d"),
    }
    df = executar_query_empresa(cfg.get("query_departamentos"), params)

    if df.empty:
        return []

    col = "DEPARTAMENTO" if "DEPARTAMENTO" in df.columns else df.columns[0]
    valores = sorted([limpar_texto(v) for v in df[col].dropna().unique().tolist() if limpar_texto(v)])
    return valores



def carregar_arquivos_empresa(data_ini, data_fim, departamento):
    cfg = get_empresa_cfg()
    params = {
        "data_ini": data_ini.strftime("%Y-%m-%d"),
        "data_fim": data_fim.strftime("%Y-%m-%d"),
        "departamento": departamento,
    }
    df = executar_query_empresa(cfg.get("query_arquivos"), params)

    if df.empty:
        return df

    if COLUNA_NOME_ARQUIVO not in df.columns:
        st.error("A query_arquivos precisa retornar a coluna NOME_ARQUIVO.")
        return pd.DataFrame()

    df[COLUNA_NOME_ARQUIVO] = df[COLUNA_NOME_ARQUIVO].apply(limpar_texto)
    df = df[df[COLUNA_NOME_ARQUIVO] != ""].drop_duplicates().copy()
    return df



def carregar_pedidos_empresa(data_ini, data_fim, departamento, nome_arquivo):
    cfg = get_empresa_cfg()
    params = {
        "data_ini": data_ini.strftime("%Y-%m-%d"),
        "data_fim": data_fim.strftime("%Y-%m-%d"),
        "departamento": departamento,
        "nome_arquivo": nome_arquivo,
    }
    df = executar_query_empresa(cfg.get("query_pedidos"), params)

    if df.empty:
        return df

    if COLUNA_PEDIDO not in df.columns:
        st.error("A query_pedidos precisa retornar a coluna PEDIDO.")
        return pd.DataFrame()

    df[COLUNA_PEDIDO] = df[COLUNA_PEDIDO].apply(limpar_texto)
    df = df[df[COLUNA_PEDIDO] != ""].copy()

    if COLUNA_FORNECEDOR not in df.columns:
        df[COLUNA_FORNECEDOR] = ""

    if COLUNA_LOJA not in df.columns:
        df[COLUNA_LOJA] = "GERAL"

    if COLUNA_VALOR not in df.columns:
        df[COLUNA_VALOR] = None

    if COLUNA_DEPARTAMENTO not in df.columns:
        df[COLUNA_DEPARTAMENTO] = departamento

    if COLUNA_NOME_ARQUIVO not in df.columns:
        df[COLUNA_NOME_ARQUIVO] = nome_arquivo

    if "STATUS_BANCO" not in df.columns:
        if "COD_STATUS_BANCO" in df.columns:
            df["STATUS_BANCO"] = df["COD_STATUS_BANCO"].apply(normalizar_status)
        else:
            df["STATUS_BANCO"] = "PENDENTE CONSULTA"

    if "COD_STATUS_BANCO" not in df.columns:
        df["COD_STATUS_BANCO"] = df["STATUS_BANCO"]

    df[COLUNA_FORNECEDOR] = df[COLUNA_FORNECEDOR].apply(limpar_texto)
    df[COLUNA_LOJA] = df[COLUNA_LOJA].apply(limpar_texto).replace("", "GERAL")
    df[COLUNA_VALOR] = df[COLUNA_VALOR].apply(limpar_valor)
    df[COLUNA_DEPARTAMENTO] = df[COLUNA_DEPARTAMENTO].apply(limpar_texto)
    df[COLUNA_NOME_ARQUIVO] = df[COLUNA_NOME_ARQUIVO].apply(limpar_texto)
    df["STATUS_BANCO"] = df["STATUS_BANCO"].apply(normalizar_status)
    df["COD_STATUS_BANCO"] = df["COD_STATUS_BANCO"].apply(limpar_texto)

    return df


# =========================================================
# SALVAR NO NEON
# =========================================================


def salvar_carga_no_neon(df_final, nome_carga, nome_arquivo, usuario_importacao, departamento, data_ini, data_fim):
    engine = get_engine_neon()
    if engine is None:
        st.error("Neon/PostgreSQL não configurado. Verifique os Secrets.")
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
                    (nome_carga, nome_arquivo, usuario_importacao, qtd_pedidos, qtd_liberados, qtd_nao_liberados,
                     departamento, filtro_data_ini, filtro_data_fim)
                    VALUES
                    (:nome_carga, :nome_arquivo, :usuario_importacao, :qtd_pedidos, :qtd_liberados, :qtd_nao_liberados,
                     :departamento, :filtro_data_ini, :filtro_data_fim)
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
                    "departamento": departamento,
                    "filtro_data_ini": data_ini,
                    "filtro_data_fim": data_fim,
                },
            ).scalar_one()

            for _, row in df_final.iterrows():
                pedido = limpar_texto(row.get(COLUNA_PEDIDO))
                fornecedor = limpar_texto(row.get(COLUNA_FORNECEDOR))
                loja = limpar_texto(row.get(COLUNA_LOJA)) or "GERAL"
                valor = limpar_valor(row.get(COLUNA_VALOR))
                depto_row = limpar_texto(row.get(COLUNA_DEPARTAMENTO)) or departamento
                nome_arq_row = limpar_texto(row.get(COLUNA_NOME_ARQUIVO)) or nome_arquivo
                status_banco = limpar_texto(row.get("STATUS_BANCO"))
                cod_status_banco = limpar_texto(row.get("COD_STATUS_BANCO"))
                dados = row_json(row)

                conn.execute(
                    text(
                        """
                        INSERT INTO pedidos_importados
                        (carga_id, pedido, fornecedor, loja, valor, status_banco, cod_status_banco,
                         dados_planilha, departamento, nome_arquivo_origem)
                        VALUES
                        (:carga_id, :pedido, :fornecedor, :loja, :valor, :status_banco, :cod_status_banco,
                         CAST(:dados_planilha AS JSONB), :departamento, :nome_arquivo_origem)
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
                        "departamento": depto_row,
                        "nome_arquivo_origem": nome_arq_row,
                    },
                )

                if status_banco == "LIBERADO":
                    conn.execute(
                        text(
                            """
                            INSERT INTO pedidos_liberados_historico
                            (pedido, fornecedor, loja, valor, status_banco, cod_status_banco,
                             primeira_carga, ultima_carga, dados_ultima_planilha, departamento)
                            VALUES
                            (:pedido, :fornecedor, :loja, :valor, :status_banco, :cod_status_banco,
                             :nome_carga, :nome_carga, CAST(:dados_planilha AS JSONB), :departamento)
                            ON CONFLICT (pedido, loja)
                            DO UPDATE SET
                                fornecedor = EXCLUDED.fornecedor,
                                valor = EXCLUDED.valor,
                                status_banco = EXCLUDED.status_banco,
                                cod_status_banco = EXCLUDED.cod_status_banco,
                                ultima_carga = EXCLUDED.ultima_carga,
                                ultima_atualizacao = NOW(),
                                dados_ultima_planilha = EXCLUDED.dados_ultima_planilha,
                                departamento = EXCLUDED.departamento
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
                            "departamento": depto_row,
                        },
                    )

        return True

    except Exception as e:
        st.error("Erro ao salvar no Neon/PostgreSQL.")
        st.code(str(e))
        return False


# =========================================================
# CONSULTAS HISTÓRICO NEON
# =========================================================


def carregar_cargas_neon():
    engine = get_engine_neon()
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
                    departamento,
                    filtro_data_ini,
                    filtro_data_fim,
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



def carregar_historico_liberados_neon():
    engine = get_engine_neon()
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
                    departamento,
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
    '<div class="subtitulo">Leia os pedidos direto do banco da empresa por data, departamento e nome do arquivo importado. Depois salve o histórico no Neon/PostgreSQL.</div>',
    unsafe_allow_html=True,
)

ok_neon, msg_neon = testar_conexao_neon()
ok_empresa, msg_empresa = testar_conexao_empresa()

with st.sidebar:
    st.header("Conexões")

    if ok_neon:
        st.success("Neon conectado")
        criar_tabelas_neon()
    else:
        st.error("Neon não conectado")
        st.caption(msg_neon)

    if ok_empresa:
        st.success("Banco da empresa conectado")
    else:
        st.error("Banco da empresa não conectado")
        st.caption(msg_empresa)

    st.divider()
    usuario_importacao = st.text_input("Usuário", value="Vitoria")

aba_consulta, aba_historico, aba_config = st.tabs([
    "🔎 Consultar banco da empresa",
    "📚 Histórico Neon",
    "⚙️ Configuração",
])


with aba_consulta:
    st.subheader("1. Filtros para buscar os arquivos importados")

    hoje = date.today()
    col_data1, col_data2, col_analista, col_depto = st.columns([1, 1, 1.1, 1.6])

    with col_data1:
        data_ini = st.date_input("Data inicial", value=hoje - timedelta(days=30), format="DD/MM/YYYY")

    with col_data2:
        data_fim = st.date_input("Data final", value=hoje, format="DD/MM/YYYY")

    with col_analista:
        analista = st.selectbox("Analista", options=list(ANALISTAS.keys()))

    with col_depto:
        departamentos_analista = ANALISTAS.get(analista, [])
        departamento = st.selectbox("Departamento", options=departamentos_analista)

    if data_ini > data_fim:
        st.error("A data inicial não pode ser maior que a data final.")
        st.stop()

    buscar = st.button("🔎 Buscar arquivos importados", type="primary", disabled=not ok_empresa)

    if buscar:
        with st.spinner("Buscando arquivos importados no banco da empresa..."):
            st.session_state["df_arquivos_empresa"] = carregar_arquivos_empresa(data_ini, data_fim, departamento)
            st.session_state["filtros_arquivos"] = {
                "data_ini": data_ini,
                "data_fim": data_fim,
                "analista": analista,
                "departamento": departamento,
            }

    df_arquivos = st.session_state.get("df_arquivos_empresa", pd.DataFrame())

    if not df_arquivos.empty:
        st.subheader("2. Escolha o arquivo importado")
        st.dataframe(df_arquivos, use_container_width=True, hide_index=True)

        arquivos = df_arquivos[COLUNA_NOME_ARQUIVO].dropna().astype(str).drop_duplicates().tolist()
        nome_arquivo = st.selectbox("Nome do arquivo importado", options=arquivos)

        nome_carga_sugerido = f"{analista} - {departamento} - {nome_arquivo} - {data_ini.strftime('%d/%m/%Y')} a {data_fim.strftime('%d/%m/%Y')}"
        nome_carga = st.text_input("Nome para salvar no histórico do Neon", value=nome_carga_sugerido)

        if st.button("📥 Carregar pedidos desse arquivo", type="primary"):
            with st.spinner("Carregando pedidos do arquivo selecionado..."):
                df_pedidos = carregar_pedidos_empresa(data_ini, data_fim, departamento, nome_arquivo)
                st.session_state["df_pedidos_empresa"] = df_pedidos
                st.session_state["nome_arquivo_selecionado"] = nome_arquivo
                st.session_state["nome_carga"] = nome_carga
                st.session_state["data_ini"] = data_ini
                st.session_state["data_fim"] = data_fim
                st.session_state["analista"] = analista
                st.session_state["departamento"] = departamento

    elif buscar:
        st.warning("Nenhum arquivo encontrado para os filtros selecionados.")

    df_final = st.session_state.get("df_pedidos_empresa", pd.DataFrame())

    if not df_final.empty:
        st.divider()
        st.subheader("3. Pedidos encontrados")

        total_pedidos = len(df_final)
        total_liberados = int((df_final["STATUS_BANCO"] == "LIBERADO").sum()) if "STATUS_BANCO" in df_final.columns else 0
        total_nao_liberados = total_pedidos - total_liberados
        total_valor = df_final[COLUNA_VALOR].fillna(0).sum() if COLUNA_VALOR in df_final.columns else 0

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Total de pedidos", total_pedidos)
        m2.metric("Liberados", total_liberados)
        m3.metric("Não liberados / pendentes", total_nao_liberados)
        m4.metric("Valor total", formatar_moeda(total_valor))

        st.dataframe(df_final, use_container_width=True, hide_index=True)

        col_baixar, col_salvar = st.columns([1, 1])

        with col_baixar:
            excel = gerar_excel(df_final)
            st.download_button(
                "📥 Baixar Excel final",
                data=excel,
                file_name=f"{st.session_state.get('nome_carga', 'pedidos')}.xlsx".replace("/", "-"),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        with col_salvar:
            if st.button("💾 Salvar no Neon", type="primary", use_container_width=True):
                if not ok_neon:
                    st.error("Neon não conectado. Verifique os Secrets.")
                else:
                    sucesso = salvar_carga_no_neon(
                        df_final=df_final,
                        nome_carga=st.session_state.get("nome_carga", "Carga sem nome"),
                        nome_arquivo=st.session_state.get("nome_arquivo_selecionado", ""),
                        usuario_importacao=st.session_state.get("analista", usuario_importacao),
                        departamento=st.session_state.get("departamento", ""),
                        data_ini=st.session_state.get("data_ini", data_ini),
                        data_fim=st.session_state.get("data_fim", data_fim),
                    )
                    if sucesso:
                        st.success("Pedidos salvos no Neon e histórico de liberados atualizado.")


with aba_historico:
    if not ok_neon:
        st.warning("Conecte o Neon para visualizar o histórico.")
    else:
        st.subheader("Cargas salvas no Neon")
        df_cargas = carregar_cargas_neon()

        if df_cargas.empty:
            st.info("Nenhuma carga salva ainda.")
        else:
            st.dataframe(df_cargas, use_container_width=True, hide_index=True)

        st.divider()
        st.subheader("Pedidos liberados mantidos no histórico")
        df_hist = carregar_historico_liberados_neon()

        if df_hist.empty:
            st.info("Nenhum pedido liberado salvo ainda.")
        else:
            f1, f2, f3, f4 = st.columns(4)
            filtro_pedido = f1.text_input("Filtrar pedido")
            filtro_fornecedor = f2.text_input("Filtrar fornecedor")
            filtro_depto = f3.text_input("Filtrar departamento")
            filtro_loja = f4.text_input("Filtrar loja")

            df_view = df_hist.copy()
            if filtro_pedido:
                df_view = df_view[df_view["pedido"].astype(str).str.contains(filtro_pedido, case=False, na=False)]
            if filtro_fornecedor:
                df_view = df_view[df_view["fornecedor"].astype(str).str.contains(filtro_fornecedor, case=False, na=False)]
            if filtro_depto:
                df_view = df_view[df_view["departamento"].astype(str).str.contains(filtro_depto, case=False, na=False)]
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
    st.subheader("Como esse app funciona agora")
    st.markdown(
        """
        Agora o app **não pede upload de planilha**.

        O fluxo é:

        1. Usuário escolhe **data inicial**, **data final** e **analista**.
        2. O app mostra somente os **departamentos vinculados ao analista escolhido**.
        3. O app consulta o banco da empresa e lista os **nomes dos arquivos importados**.
        4. Usuário escolhe o **nome do arquivo**.
        5. O app busca os pedidos daquele arquivo no banco da empresa.
        6. O app salva a carga e mantém o histórico dos liberados no Neon/PostgreSQL.
        """
    )

    st.subheader("Secrets necessários no Streamlit")
    st.code(
        '''[connections.postgresql]
url = "postgresql://neondb_owner:SUA_SENHA@SEU_HOST.neon.tech/neondb?sslmode=require"

[empresa_oracle]
enabled = true
user = "SEU_USUARIO_ORACLE"
password = "SUA_SENHA_ORACLE"
dsn = "host:1521/service"

# Opcional: se não configurar, o app deixa você digitar o departamento.
query_departamentos = """
SELECT DISTINCT
    DEPARTAMENTO
FROM SUA_TABELA_DE_IMPORTACAO
WHERE DATA_IMPORTACAO >= TO_DATE(:data_ini, 'YYYY-MM-DD')
  AND DATA_IMPORTACAO <  TO_DATE(:data_fim, 'YYYY-MM-DD') + 1
ORDER BY DEPARTAMENTO
"""

# Obrigatório: precisa retornar NOME_ARQUIVO.
query_arquivos = """
SELECT DISTINCT
    NOME_ARQUIVO,
    MAX(DATA_IMPORTACAO) AS DATA_IMPORTACAO,
    COUNT(DISTINCT PEDIDO) AS QTD_PEDIDOS
FROM SUA_TABELA_DE_IMPORTACAO
WHERE DATA_IMPORTACAO >= TO_DATE(:data_ini, 'YYYY-MM-DD')
  AND DATA_IMPORTACAO <  TO_DATE(:data_fim, 'YYYY-MM-DD') + 1
  AND (:departamento = 'TODOS' OR DEPARTAMENTO = :departamento)
GROUP BY NOME_ARQUIVO
ORDER BY MAX(DATA_IMPORTACAO) DESC
"""

# Obrigatório: precisa retornar pelo menos PEDIDO.
# Recomendado retornar também FORNECEDOR, LOJA, VALOR, DEPARTAMENTO, NOME_ARQUIVO, STATUS_BANCO e COD_STATUS_BANCO.
query_pedidos = """
SELECT
    PEDIDO,
    FORNECEDOR,
    LOJA,
    VALOR,
    DEPARTAMENTO,
    NOME_ARQUIVO,
    DATA_IMPORTACAO,
    CASE
        WHEN COD_STATUS = 'L' THEN 'LIBERADO'
        WHEN COD_STATUS = 'B' THEN 'BLOQUEADO'
        ELSE 'NÃO LIBERADO'
    END AS STATUS_BANCO,
    COD_STATUS AS COD_STATUS_BANCO
FROM SUA_TABELA_DE_IMPORTACAO
WHERE DATA_IMPORTACAO >= TO_DATE(:data_ini, 'YYYY-MM-DD')
  AND DATA_IMPORTACAO <  TO_DATE(:data_fim, 'YYYY-MM-DD') + 1
  AND (:departamento = 'TODOS' OR DEPARTAMENTO = :departamento)
  AND NOME_ARQUIVO = :nome_arquivo
ORDER BY PEDIDO
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

    st.warning(
        "Importante: se o banco da empresa for interno/VPN, o Streamlit Cloud só vai conseguir acessar se a empresa liberar acesso externo. Caso contrário, esse app precisa rodar dentro da rede da empresa ou atrás de uma API interna."
    )
