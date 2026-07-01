import json
from datetime import date, datetime, timedelta
from io import BytesIO

import pandas as pd
import requests
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
# COLUNAS PADRÃO QUE A API RETORNA
# =========================================================

COLUNA_PEDIDO = "PEDIDO"
COLUNA_FORNECEDOR = "FORNECEDOR"
COLUNA_LOJA = "FILIAL"
COLUNA_VALOR = "VALOR_TOTAL"
COLUNA_DEPARTAMENTO = "DEPARTAMENTO"
COLUNA_NOME_ARQUIVO = "NOME_ARQUIVO"
COLUNA_DATA = "DT_EMISSAO"

VALORES_STATUS_LIBERADO = {"L", "LIBERADO", "LIBERADA", "APROVADO", "APROVADA", "S", "SIM"}


# =========================================================
# ANALISTAS E DEPARTAMENTOS FIXOS
# =========================================================

ANALISTAS = {
    "Cleviton": [
        "PORTAS E JANELAS", "FERRAMENTAS", "FERRAGENS", "AUTOMOTIVOS"
    ],
    "Alec": [
        "ELETRICA", "ILUMINACAO", "HIDRAULICA"
    ],
    "Jonatas": [
        "MOVEIS E COLCHOES", "DECORACAO", "CAMA MESA E BANHO", "LAZER",
        "CASA E UD", "JARDIM",
    ],
    "Beatriz": [
        "ELETRO", "TECNOLOGIA", "CLIMATIZACAO"
    ],
    "Ruan": [
        "TINTAS", "ORGANIZACAO DA CASA"
    ],
    "Jessica": [
        "MATERIAIS DE CONSTRUCAO", "BANHO E COZINHA"
    ],
    "Rose": [
        "PISOS E REVESTIMENTO"
    ],
}

TODOS_DEPARTAMENTOS = sorted({
    departamento
    for departamentos in ANALISTAS.values()
    for departamento in departamentos
})


# =========================================================
# VISUAL
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
        div[data-testid="stMetric"] {
            border: 1px solid rgba(255,255,255,0.10);
            border-radius: 12px;
            padding: 12px;
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

    ALTER TABLE cargas_importacao ADD COLUMN IF NOT EXISTS departamento TEXT;
    ALTER TABLE cargas_importacao ADD COLUMN IF NOT EXISTS filtro_data_ini DATE;
    ALTER TABLE cargas_importacao ADD COLUMN IF NOT EXISTS filtro_data_fim DATE;
    ALTER TABLE pedidos_importados ADD COLUMN IF NOT EXISTS departamento TEXT;
    ALTER TABLE pedidos_importados ADD COLUMN IF NOT EXISTS nome_arquivo_origem TEXT;
    ALTER TABLE pedidos_liberados_historico ADD COLUMN IF NOT EXISTS departamento TEXT;
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
# CONEXÃO COM API DA EMPRESA / FASTAPI
# =========================================================


def get_empresa_api_cfg():
    try:
        return st.secrets["empresa_api"]
    except Exception:
        return None



def get_api_base_url():
    cfg = get_empresa_api_cfg()
    if cfg is None:
        return None
    return str(cfg.get("url", "")).rstrip("/")



def get_api_token():
    cfg = get_empresa_api_cfg()
    if cfg is None:
        return None
    return str(cfg.get("token", ""))



def api_habilitada():
    cfg = get_empresa_api_cfg()
    if cfg is None:
        return False, "Configuração [empresa_api] não encontrada nos Secrets."

    url = get_api_base_url()
    token = get_api_token()

    if not url:
        return False, "URL da API não encontrada em [empresa_api]."

    if not token:
        return False, "Token da API não encontrado em [empresa_api]."

    return True, "Configuração OK"



def chamar_api(method, path, payload=None, timeout=180):
    ok, msg = api_habilitada()
    if not ok:
        raise RuntimeError(msg)

    base_url = get_api_base_url()
    token = get_api_token()
    url = f"{base_url}{path}"

    headers = {
        "x-api-key": token,
        "Content-Type": "application/json",
    }

    if method.upper() == "GET":
        response = requests.get(url, headers=headers, timeout=timeout)
    else:
        response = requests.post(url, headers=headers, json=payload or {}, timeout=timeout)

    if response.status_code >= 400:
        try:
            detalhe = response.json()
        except Exception:
            detalhe = response.text
        raise RuntimeError(f"Erro API {response.status_code}: {detalhe}")

    return response.json()



def testar_conexao_api():
    ok, msg = api_habilitada()
    if not ok:
        return False, msg

    try:
        retorno = chamar_api("GET", "/health", timeout=30)
        status = retorno.get("status", "") if isinstance(retorno, dict) else ""
        if str(status).lower() == "ok":
            return True, "Conexão OK"
        return False, f"API respondeu, mas status inesperado: {retorno}"
    except Exception as e:
        return False, str(e)


# =========================================================
# CONSULTAS VIA API
# =========================================================


def carregar_arquivos_empresa(data_ini, data_fim, departamentos):
    if not departamentos:
        return pd.DataFrame()

    payload = {
        "data_inicial": data_ini.isoformat(),
        "data_final": data_fim.isoformat(),
        "departamentos": [limpar_texto(d).upper() for d in departamentos],
    }

    try:
        retorno = chamar_api("POST", "/arquivos", payload=payload)
        dados = retorno.get("dados", []) if isinstance(retorno, dict) else []
        df = pd.DataFrame(dados)
        df.columns = [str(c).strip().upper() for c in df.columns]
    except Exception as e:
        st.error("Erro ao buscar arquivos na API da empresa.")
        st.code(str(e))
        return pd.DataFrame()

    if df.empty:
        return df

    if COLUNA_NOME_ARQUIVO not in df.columns:
        st.error("A API precisa retornar a coluna NOME_ARQUIVO.")
        return pd.DataFrame()

    df[COLUNA_NOME_ARQUIVO] = df[COLUNA_NOME_ARQUIVO].apply(limpar_texto)
    df = df[df[COLUNA_NOME_ARQUIVO] != ""].drop_duplicates().copy()
    return df



def carregar_pedidos_empresa(data_ini, data_fim, departamentos, nome_arquivo):
    if not departamentos or not nome_arquivo:
        return pd.DataFrame()

    payload = {
        "data_inicial": data_ini.isoformat(),
        "data_final": data_fim.isoformat(),
        "departamentos": [limpar_texto(d).upper() for d in departamentos],
        "nome_arquivo": limpar_texto(nome_arquivo),
    }

    try:
        retorno = chamar_api("POST", "/pedidos", payload=payload)
        dados = retorno.get("dados", []) if isinstance(retorno, dict) else []
        df = pd.DataFrame(dados)
        df.columns = [str(c).strip().upper() for c in df.columns]
    except Exception as e:
        st.error("Erro ao buscar pedidos na API da empresa.")
        st.code(str(e))
        return pd.DataFrame()

    if df.empty:
        return df

    if COLUNA_PEDIDO not in df.columns:
        st.error("A API precisa retornar a coluna PEDIDO.")
        return pd.DataFrame()

    df[COLUNA_PEDIDO] = df[COLUNA_PEDIDO].apply(limpar_texto)
    df = df[df[COLUNA_PEDIDO] != ""].copy()

    # Padronizações para o Neon e métricas.
    if COLUNA_FORNECEDOR not in df.columns:
        df[COLUNA_FORNECEDOR] = ""
    if COLUNA_LOJA not in df.columns:
        df[COLUNA_LOJA] = "GERAL"
    if COLUNA_VALOR not in df.columns:
        df[COLUNA_VALOR] = None
    if COLUNA_DEPARTAMENTO not in df.columns:
        df[COLUNA_DEPARTAMENTO] = ""
    if COLUNA_NOME_ARQUIVO not in df.columns:
        df[COLUNA_NOME_ARQUIVO] = nome_arquivo
    if "STATUS_BANCO" not in df.columns:
        df["STATUS_BANCO"] = df.get("COD_STATUS_BANCO", "").apply(normalizar_status)
    if "COD_STATUS_BANCO" not in df.columns:
        df["COD_STATUS_BANCO"] = df["STATUS_BANCO"]

    df[COLUNA_FORNECEDOR] = df[COLUNA_FORNECEDOR].apply(limpar_texto)
    df[COLUNA_LOJA] = df[COLUNA_LOJA].apply(limpar_texto).replace("", "GERAL")
    df[COLUNA_VALOR] = df[COLUNA_VALOR].apply(limpar_valor)
    df[COLUNA_DEPARTAMENTO] = df[COLUNA_DEPARTAMENTO].apply(limpar_texto)
    df[COLUNA_NOME_ARQUIVO] = df[COLUNA_NOME_ARQUIVO].apply(limpar_texto)
    df["STATUS_BANCO"] = df["STATUS_BANCO"].apply(normalizar_status)
    df["COD_STATUS_BANCO"] = df["COD_STATUS_BANCO"].apply(limpar_texto)

    # Remove apenas linhas 100% repetidas, sem perder produtos diferentes do mesmo pedido.
    df = df.drop_duplicates().copy()
    return df


# =========================================================
# SALVAR NO NEON
# =========================================================


def salvar_carga_no_neon(df_final, nome_carga, nome_arquivo, usuario_importacao, departamento, data_ini, data_fim):
    engine = get_engine_neon()
    if engine is None:
        st.error("Neon/PostgreSQL não configurado. Verifique os Secrets.")
        return False

    qtd_pedidos = int(df_final[COLUNA_PEDIDO].nunique()) if COLUNA_PEDIDO in df_final.columns else len(df_final)
    pedidos_liberados = df_final.loc[df_final["STATUS_BANCO"] == "LIBERADO", COLUNA_PEDIDO].nunique()
    qtd_liberados = int(pedidos_liberados)
    qtd_nao_liberados = max(qtd_pedidos - qtd_liberados, 0)

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
# HISTÓRICO NEON
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
    '<div class="subtitulo">Consulta via API FastAPI por data, analista, múltiplos departamentos e nome do arquivo importado. Depois salva o histórico no Neon/PostgreSQL.</div>',
    unsafe_allow_html=True,
)

ok_neon, msg_neon = testar_conexao_neon()
ok_api, msg_api = testar_conexao_api()

with st.sidebar:
    st.header("Conexões")

    if ok_neon:
        st.success("Neon conectado")
        criar_tabelas_neon()
    else:
        st.error("Neon não conectado")
        st.caption(msg_neon)

    if ok_api:
        st.success("API da empresa conectada")
    else:
        st.error("API da empresa não conectada")
        st.caption(msg_api)

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
    col_data1, col_data2, col_analista = st.columns([1, 1, 1.2])

    with col_data1:
        data_ini = st.date_input("Data inicial", value=hoje - timedelta(days=30), format="DD/MM/YYYY")

    with col_data2:
        data_fim = st.date_input("Data final", value=hoje, format="DD/MM/YYYY")

    with col_analista:
        analista = st.selectbox("Analista", options=list(ANALISTAS.keys()))

    departamentos_analista = ANALISTAS.get(analista, [])
    departamentos_selecionados = st.multiselect(
        "Departamento",
        options=TODOS_DEPARTAMENTOS,
        default=departamentos_analista,
        help="Multi seleção: os departamentos do analista já vêm marcados, mas você pode incluir/remover qualquer outro departamento.",
    )

    departamento_carga = "; ".join(departamentos_selecionados)

    if data_ini > data_fim:
        st.error("A data inicial não pode ser maior que a data final.")
        st.stop()

    if not departamentos_selecionados:
        st.warning("Selecione pelo menos um departamento para buscar os arquivos.")

    buscar = st.button(
        "🔎 Buscar arquivos importados",
        type="primary",
        disabled=(not ok_api or not departamentos_selecionados),
    )

    if buscar:
        with st.spinner("Buscando arquivos importados pela API da empresa..."):
            st.session_state["df_arquivos_empresa"] = carregar_arquivos_empresa(
                data_ini,
                data_fim,
                departamentos_selecionados,
            )
            st.session_state["filtros_arquivos"] = {
                "data_ini": data_ini,
                "data_fim": data_fim,
                "analista": analista,
                "departamentos": departamentos_selecionados,
                "departamento": departamento_carga,
            }

    df_arquivos = st.session_state.get("df_arquivos_empresa", pd.DataFrame())

    if not df_arquivos.empty:
        st.subheader("2. Escolha o arquivo importado")
        st.dataframe(df_arquivos, use_container_width=True, hide_index=True)

        arquivos = df_arquivos[COLUNA_NOME_ARQUIVO].dropna().astype(str).drop_duplicates().tolist()
        nome_arquivo = st.selectbox("Nome do arquivo importado", options=arquivos)

        nome_carga_sugerido = (
            f"{analista} - {departamento_carga} - {nome_arquivo} - "
            f"{data_ini.strftime('%d/%m/%Y')} a {data_fim.strftime('%d/%m/%Y')}"
        )
        nome_carga = st.text_input("Nome para salvar no histórico do Neon", value=nome_carga_sugerido)

        if st.button("📥 Carregar pedidos desse arquivo", type="primary"):
            with st.spinner("Carregando pedidos do arquivo selecionado pela API..."):
                df_pedidos = carregar_pedidos_empresa(
                    data_ini,
                    data_fim,
                    departamentos_selecionados,
                    nome_arquivo,
                )
                st.session_state["df_pedidos_empresa"] = df_pedidos
                st.session_state["nome_arquivo_selecionado"] = nome_arquivo
                st.session_state["nome_carga"] = nome_carga
                st.session_state["data_ini"] = data_ini
                st.session_state["data_fim"] = data_fim
                st.session_state["analista"] = analista
                st.session_state["departamentos_selecionados"] = departamentos_selecionados
                st.session_state["departamento"] = departamento_carga

    elif buscar:
        st.warning("Nenhum arquivo encontrado para os filtros selecionados.")

    df_final = st.session_state.get("df_pedidos_empresa", pd.DataFrame())

    if not df_final.empty:
        st.divider()
        st.subheader("3. Pedidos encontrados")

        total_linhas = len(df_final)
        total_pedidos = df_final[COLUNA_PEDIDO].nunique() if COLUNA_PEDIDO in df_final.columns else len(df_final)
        total_liberados = df_final.loc[df_final["STATUS_BANCO"] == "LIBERADO", COLUNA_PEDIDO].nunique() if "STATUS_BANCO" in df_final.columns else 0
        total_bloqueados = df_final.loc[df_final["STATUS_BANCO"] == "BLOQUEADO", COLUNA_PEDIDO].nunique() if "STATUS_BANCO" in df_final.columns else 0
        total_valor = df_final[COLUNA_VALOR].fillna(0).sum() if COLUNA_VALOR in df_final.columns else 0

        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric("Linhas", total_linhas)
        m2.metric("Pedidos únicos", total_pedidos)
        m3.metric("Liberados", int(total_liberados))
        m4.metric("Bloqueados", int(total_bloqueados))
        m5.metric("Valor total", formatar_moeda(total_valor))

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
                        usuario_importacao=usuario_importacao or st.session_state.get("analista", ""),
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
            filtro_loja = f4.text_input("Filtrar filial")

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
    st.subheader("O que foi ajustado")
    st.markdown(
        """
        - O Streamlit agora consulta o banco da empresa por meio da **API FastAPI**.
        - O Streamlit não conecta mais direto no Oracle.
        - O campo **Departamento** é **multiselect**.
        - Os departamentos do analista já vêm **pré-selecionados**.
        - O usuário pode marcar/remover qualquer outro departamento.
        - O histórico continua sendo salvo no **Neon/PostgreSQL**.
        """
    )

    st.subheader("Secrets necessários no Streamlit local")
    st.code(
        '''[connections.postgresql]
url = "postgresql://neondb_owner:SUA_SENHA@SEU_HOST.neon.tech/neondb?sslmode=require"

[empresa_api]
url = "http://localhost:8000"
token = "SEU_API_TOKEN"
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
requests
""",
        language="txt",
    )

    st.warning(
        "Para testar localmente, deixe a API FastAPI rodando em outro terminal. Para usar no Streamlit Cloud, a API precisa estar publicada em uma URL acessível pelo Cloud."
    )
