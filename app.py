import json
import hashlib
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
COLUNA_DEMANDA = "DEMANDA"
COLUNA_COD_TIPO_PEDIDO_TOTVS = "COD_TIPO_PEDIDO_TOTVS"
COLUNA_OBSERVACAO = "OBSERVACAO"

VALORES_STATUS_LIBERADO = {"L", "LIBERADO", "LIBERADA", "APROVADO", "APROVADA", "S", "SIM", "OK"}
DEMANDAS = ["DEMANDA PUXADA", "DEMANDA EMPURRADA", "DEMANDA VENDA JURIDICA"]

# A escolha Puxada/Empurrada/Venda Jurídica é classificação manual do analista.
# C7_XTPPED não é usado para filtrar essa classificação.
COLUNA_VALOR_LOJA_NAOMISTO_TOTAL = "VALOR_LOJA_NAOMISTO_TOTAL"
COLUNA_QTD_LOJA_NAOMISTO = "QTD_PEDIDOS_LOJA_NAOMISTO"


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
        .block-container { padding-top: 2.5rem; }
        .titulo-principal {
            font-size: 34px;
            font-weight: 800;
            margin-bottom: 18px;
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

    if status in {"NÃO ENCONTRADO", "NAO ENCONTRADO"}:
        return "NÃO ENCONTRADO"

    return "NÃO LIBERADO"



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


def status_icone(valor):
    status = limpar_texto(valor).upper()
    if "LIBERADO" in status or status in {"L", "OK", "✅"}:
        return "✅"
    if "BLOQUEADO" in status or status in {"B", "❌", "X"}:
        return "❌"
    return status or "-"


def valor_loja_naomisto_total(linha_arquivo):
    """Retorna o valor total do arquivo calculado pela query, sem duplicar itens."""
    for coluna in ["VALOR_TOTAL", "VALOR_ARQUIVO", COLUNA_VALOR_LOJA_NAOMISTO_TOTAL, "VALOR_TIPO2_TOTAL"]:
        if coluna in linha_arquivo.index:
            return limpar_valor(linha_arquivo.get(coluna)) or 0.0
    return 0.0


def valor_pedido_tipo_fornecedor(grupo):
    """
    Calcula o valor do pedido para exibição/histórico.

    Regra atual: quando existir linha TIPO_FORNECEDOR no pedido,
    o valor do pedido deve ser a soma somente das linhas TIPO_FORNECEDOR.
    Se não existir TIPO no retorno da API, usa a soma normal como fallback.
    """
    if grupo is None or grupo.empty or COLUNA_VALOR not in grupo.columns:
        return 0.0

    base = grupo.copy()
    if "TIPO" in base.columns:
        tipo = base["TIPO"].astype(str).str.strip().str.upper()
        base_fornecedor = base[tipo == "TIPO_FORNECEDOR"].copy()
        if not base_fornecedor.empty:
            base = base_fornecedor

    return float(sum((limpar_valor(v) or 0) for v in base[COLUNA_VALOR]))


def qtd_loja_naomisto_total(linha_arquivo):
    for coluna in ["QTD_PEDIDOS", COLUNA_QTD_LOJA_NAOMISTO, "QTD_PEDIDOS_TIPO2"]:
        if coluna in linha_arquivo.index:
            try:
                return int(float(linha_arquivo.get(coluna) or 0))
            except Exception:
                return 0
    return 0


def formatar_data_protheus(valor):
    valor_txt = limpar_texto(valor)
    if not valor_txt:
        return ""
    if len(valor_txt) >= 10 and "-" in valor_txt:
        try:
            return pd.to_datetime(valor_txt).strftime("%d/%m/%Y")
        except Exception:
            return valor_txt
    if len(valor_txt) == 8 and valor_txt.isdigit():
        return f"{valor_txt[6:8]}/{valor_txt[4:6]}/{valor_txt[0:4]}"
    return valor_txt


def converter_para_data(valor):
    """Converte datas do Oracle/Neon para date, aceitando YYYYMMDD, timestamp e datetime."""
    if valor is None or pd.isna(valor):
        return None

    if isinstance(valor, datetime):
        return valor.date()

    if isinstance(valor, date):
        return valor

    valor_txt = limpar_texto(valor)
    if not valor_txt:
        return None

    try:
        if len(valor_txt) == 8 and valor_txt.isdigit():
            return datetime.strptime(valor_txt, "%Y%m%d").date()

        convertido = pd.to_datetime(valor_txt, errors="coerce")
        if pd.isna(convertido):
            return None
        return convertido.date()
    except Exception:
        return None


def formatar_data_ddmmyyyy(valor):
    dt = converter_para_data(valor)
    if not dt:
        return ""
    return dt.strftime("%d/%m/%Y")


def formatar_data_hora_ddmmyyyy(valor):
    if valor is None or pd.isna(valor):
        return ""
    try:
        dt = pd.to_datetime(valor, errors="coerce")
        if pd.isna(dt):
            return ""
        return dt.strftime("%d/%m/%Y %H:%M")
    except Exception:
        return limpar_texto(valor)


def obter_data_historico(row):
    """Prioriza a data de emissão do pedido; se não existir, usa a data em que entrou no histórico."""
    for coluna in ["data_emissao_pedido", "data_historico", "ultima_atualizacao", "primeira_liberacao"]:
        if coluna in row.index:
            dt = converter_para_data(row.get(coluna))
            if dt:
                return dt
    return None


def montar_tabela_arquivos_display(df_arquivos):
    if df_arquivos.empty:
        return pd.DataFrame()

    df = df_arquivos.copy()
    for coluna in ["FILIAL", COLUNA_NOME_ARQUIVO, "VALOR_TOTAL", "DATA_EMISSAO"]:
        if coluna not in df.columns:
            df[coluna] = ""

    saida = pd.DataFrame({
        "Filial": df["FILIAL"].apply(limpar_texto),
        "Nome do arquivo": df[COLUNA_NOME_ARQUIVO].apply(limpar_texto),
        "Valor": df["VALOR_TOTAL"].apply(lambda v: formatar_moeda(limpar_valor(v) or 0)),
        "Data de emissão": df["DATA_EMISSAO"].apply(formatar_data_protheus),
    })
    return saida


def lista_fornecedores(df):
    if df is None or df.empty or COLUNA_FORNECEDOR not in df.columns:
        return ["Todos"]

    fornecedores = sorted({limpar_texto(v) for v in df[COLUNA_FORNECEDOR].dropna().tolist() if limpar_texto(v)})
    return ["Todos"] + fornecedores


def normalizar_nome_aba(nome):
    proibidos = ["/", "\\", "?", "*", "[", "]", ":"]
    nome_limpo = str(nome)
    for c in proibidos:
        nome_limpo = nome_limpo.replace(c, "-")
    return nome_limpo[:31] or "Aba"



def gerar_excel_abas(dfs_por_aba, df_removidos=None, df_substituicoes=None):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        escreveu = False
        for nome_aba, df in dfs_por_aba.items():
            if df is None or df.empty:
                continue
            df.to_excel(writer, index=False, sheet_name=normalizar_nome_aba(nome_aba))
            escreveu = True

        if df_removidos is not None and not df_removidos.empty:
            df_removidos.to_excel(writer, index=False, sheet_name="Removidos")
            escreveu = True

        if df_substituicoes is not None and not df_substituicoes.empty:
            df_substituicoes.to_excel(writer, index=False, sheet_name="Substituicoes")
            escreveu = True

        if not escreveu:
            pd.DataFrame({"INFO": ["Sem dados"]}).to_excel(writer, index=False, sheet_name="Pedidos")

    output.seek(0)
    return output



def definir_demanda(row):
    """A demanda é definida manualmente no Streamlit antes da importação."""
    demanda = limpar_texto(row.get(COLUNA_DEMANDA, "")).upper()
    if demanda in DEMANDAS:
        return demanda
    return "OUTROS"


def preparar_df_pedidos(df):
    if df is None or df.empty:
        return pd.DataFrame()

    df = df.copy()
    df.columns = [str(c).strip().upper() for c in df.columns]

    if COLUNA_PEDIDO in df.columns:
        df[COLUNA_PEDIDO] = df[COLUNA_PEDIDO].apply(limpar_texto)
        df = df[df[COLUNA_PEDIDO] != ""].copy()

    for coluna, padrao in [
        (COLUNA_FORNECEDOR, ""),
        (COLUNA_LOJA, "GERAL"),
        (COLUNA_DEPARTAMENTO, ""),
        (COLUNA_NOME_ARQUIVO, ""),
        ("MOTIVO", ""),
        (COLUNA_OBSERVACAO, ""),
        ("STATUS_BANCO", "NÃO INFORMADO"),
        ("COD_STATUS_BANCO", ""),
        (COLUNA_VALOR, None),
    ]:
        if coluna not in df.columns:
            df[coluna] = padrao

    df[COLUNA_FORNECEDOR] = df[COLUNA_FORNECEDOR].apply(limpar_texto)
    df[COLUNA_LOJA] = df[COLUNA_LOJA].apply(limpar_texto).replace("", "GERAL")
    df[COLUNA_DEPARTAMENTO] = df[COLUNA_DEPARTAMENTO].apply(limpar_texto)
    df[COLUNA_NOME_ARQUIVO] = df[COLUNA_NOME_ARQUIVO].apply(limpar_texto)
    df[COLUNA_VALOR] = df[COLUNA_VALOR].apply(limpar_valor)
    df["STATUS_BANCO"] = df["STATUS_BANCO"].apply(normalizar_status)
    df["COD_STATUS_BANCO"] = df["COD_STATUS_BANCO"].apply(limpar_texto)
    df["MOTIVO"] = df["MOTIVO"].apply(limpar_texto)
    df[COLUNA_OBSERVACAO] = df[COLUNA_OBSERVACAO].apply(limpar_texto)
    # Compatibilidade: se vier só MOTIVO, usa como OBSERVACAO.
    df.loc[df[COLUNA_OBSERVACAO] == "", COLUNA_OBSERVACAO] = df.loc[df[COLUNA_OBSERVACAO] == "", "MOTIVO"]
    df[COLUNA_DEMANDA] = df.apply(definir_demanda, axis=1)

    return df.drop_duplicates().copy()



def montar_tabela_planilha(df):
    if df is None or df.empty:
        return pd.DataFrame(columns=["DEPARTAMENTO", "FILIAL", "FORNECEDOR", "PEDIDO", "VALOR", "OBSERVAÇÃO", "Status"])

    df = df.copy()
    for col in [COLUNA_DEPARTAMENTO, COLUNA_LOJA, COLUNA_FORNECEDOR, COLUNA_PEDIDO, COLUNA_VALOR, COLUNA_OBSERVACAO, "MOTIVO", "STATUS_BANCO", "TIPO"]:
        if col not in df.columns:
            df[col] = ""

    # Compatibilidade: se a observação ainda estiver em MOTIVO, aproveita ela.
    df[COLUNA_OBSERVACAO] = df[COLUNA_OBSERVACAO].apply(limpar_texto)
    df.loc[df[COLUNA_OBSERVACAO] == "", COLUNA_OBSERVACAO] = df.loc[df[COLUNA_OBSERVACAO] == "", "MOTIVO"].apply(limpar_texto)
    df[COLUNA_DEPARTAMENTO] = df[COLUNA_DEPARTAMENTO].apply(limpar_texto).replace("", "SEM DEPARTAMENTO")

    def juntar_unicos(serie):
        valores = [limpar_texto(v) for v in serie if limpar_texto(v)]
        return "; ".join(sorted(set(valores)))

    linhas = []
    for chaves, grupo in df.groupby([COLUNA_DEPARTAMENTO, COLUNA_LOJA, COLUNA_FORNECEDOR, COLUNA_PEDIDO], dropna=False):
        departamento, filial, fornecedor, pedido = chaves
        status = "LIBERADO" if "LIBERADO" in set(grupo["STATUS_BANCO"]) else "BLOQUEADO"
        linhas.append({
            "DEPARTAMENTO": departamento,
            "FILIAL": filial,
            "FORNECEDOR": fornecedor,
            "PEDIDO": pedido,
            # Só o VALOR usa a regra do TIPO_FORNECEDOR.
            "VALOR": valor_pedido_tipo_fornecedor(grupo),
            "OBSERVAÇÃO": juntar_unicos(grupo[COLUNA_OBSERVACAO]),
            "Status": status,
        })

    tabela = pd.DataFrame(linhas)
    if tabela.empty:
        return pd.DataFrame(columns=["DEPARTAMENTO", "FILIAL", "FORNECEDOR", "PEDIDO", "VALOR", "OBSERVAÇÃO", "Status"])

    tabela["VALOR"] = tabela["VALOR"].apply(formatar_moeda)
    tabela = tabela[["DEPARTAMENTO", "FILIAL", "FORNECEDOR", "PEDIDO", "VALOR", "OBSERVAÇÃO", "Status"]]
    return tabela



def montar_resumo_departamento(df):
    if df is None or df.empty:
        return pd.DataFrame(columns=["Departamento", "Pedidos", "Liberados", "Bloqueados", "Valor Total"])

    df = df.copy()
    for col in [COLUNA_DEPARTAMENTO, COLUNA_LOJA, COLUNA_FORNECEDOR, COLUNA_PEDIDO, COLUNA_VALOR, "STATUS_BANCO", "TIPO"]:
        if col not in df.columns:
            df[col] = ""

    df[COLUNA_DEPARTAMENTO] = df[COLUNA_DEPARTAMENTO].apply(limpar_texto).replace("", "SEM DEPARTAMENTO")

    linhas = []
    for departamento, grupo_depto in df.groupby(COLUNA_DEPARTAMENTO, dropna=False):
        valor_total = 0.0
        for _, grupo_pedido in grupo_depto.groupby([COLUNA_LOJA, COLUNA_FORNECEDOR, COLUNA_PEDIDO], dropna=False):
            valor_total += valor_pedido_tipo_fornecedor(grupo_pedido)

        linhas.append({
            "Departamento": departamento,
            "Pedidos": grupo_depto[COLUNA_PEDIDO].nunique(),
            "Liberados": grupo_depto.loc[grupo_depto["STATUS_BANCO"] == "LIBERADO", COLUNA_PEDIDO].nunique(),
            "Bloqueados": grupo_depto.loc[grupo_depto["STATUS_BANCO"] == "BLOQUEADO", COLUNA_PEDIDO].nunique(),
            "Valor Total": formatar_moeda(valor_total),
        })

    resumo = pd.DataFrame(linhas)
    return resumo.sort_values("Departamento").reset_index(drop=True)


def montar_opcoes_pedido(df, filtro_fornecedor=""):
    if df is None or df.empty:
        return {}, []

    base = df[[COLUNA_PEDIDO, COLUNA_FORNECEDOR, COLUNA_LOJA, "STATUS_BANCO"]].drop_duplicates().copy()
    if filtro_fornecedor:
        base = base[base[COLUNA_FORNECEDOR].astype(str).str.contains(filtro_fornecedor, case=False, na=False)]

    base = base.sort_values([COLUNA_FORNECEDOR, COLUNA_PEDIDO])
    mapa = {}
    for _, row in base.iterrows():
        pedido = limpar_texto(row[COLUNA_PEDIDO])
        fornecedor = limpar_texto(row[COLUNA_FORNECEDOR])
        loja = limpar_texto(row[COLUNA_LOJA])
        status = limpar_texto(row["STATUS_BANCO"])
        label = f"{pedido} | {fornecedor} | {loja} | {status}"
        mapa[label] = pedido

    return mapa, list(mapa.keys())


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

    CREATE TABLE IF NOT EXISTS pedidos_removidos_importacao (
        id BIGSERIAL PRIMARY KEY,
        carga_id BIGINT REFERENCES cargas_importacao(id) ON DELETE CASCADE,
        nome_carga TEXT,
        nome_arquivo TEXT,
        usuario_importacao TEXT,
        pedido TEXT,
        fornecedor TEXT,
        loja TEXT,
        departamento TEXT,
        status_banco TEXT,
        cod_status_banco TEXT,
        motivo_remocao TEXT DEFAULT 'REMOVIDO DA IMPORTAÇÃO PELO ANALISTA',
        dados_planilha JSONB,
        data_remocao TIMESTAMP DEFAULT NOW()
    );

    CREATE TABLE IF NOT EXISTS substituicoes_pedidos (
        id BIGSERIAL PRIMARY KEY,
        nome_carga TEXT,
        nome_arquivo TEXT,
        usuario_registro TEXT,
        analista TEXT,
        departamento TEXT,
        pedido_antigo TEXT NOT NULL,
        fornecedor_antigo TEXT,
        loja_antiga TEXT,
        status_antigo TEXT,
        cod_status_antigo TEXT,
        pedido_novo TEXT NOT NULL,
        fornecedor_novo TEXT,
        loja_nova TEXT,
        status_novo TEXT,
        cod_status_novo TEXT,
        motivo TEXT NOT NULL,
        observacao TEXT,
        dados_pedido_antigo JSONB,
        dados_pedido_novo JSONB,
        data_registro TIMESTAMP DEFAULT NOW()
    );

    CREATE TABLE IF NOT EXISTS usuarios_sistema (
        id BIGSERIAL PRIMARY KEY,
        usuario TEXT NOT NULL UNIQUE,
        senha_hash TEXT NOT NULL,
        nome TEXT NOT NULL,
        perfil TEXT NOT NULL DEFAULT 'ANALISTA',
        analista TEXT,
        ativo BOOLEAN DEFAULT TRUE,
        criado_em TIMESTAMP DEFAULT NOW(),
        atualizado_em TIMESTAMP DEFAULT NOW()
    );

    ALTER TABLE cargas_importacao ADD COLUMN IF NOT EXISTS departamento TEXT;
    ALTER TABLE cargas_importacao ADD COLUMN IF NOT EXISTS filtro_data_ini DATE;
    ALTER TABLE cargas_importacao ADD COLUMN IF NOT EXISTS filtro_data_fim DATE;
    ALTER TABLE cargas_importacao ADD COLUMN IF NOT EXISTS analista TEXT;
    ALTER TABLE cargas_importacao ADD COLUMN IF NOT EXISTS demanda_importacao TEXT;
    ALTER TABLE cargas_importacao ADD COLUMN IF NOT EXISTS valor_informado NUMERIC(14,2);
    ALTER TABLE cargas_importacao ADD COLUMN IF NOT EXISTS valor_total_calculado NUMERIC(14,2);
    ALTER TABLE cargas_importacao ADD COLUMN IF NOT EXISTS diferenca_valor NUMERIC(14,2);
    ALTER TABLE pedidos_importados ADD COLUMN IF NOT EXISTS departamento TEXT;
    ALTER TABLE pedidos_importados ADD COLUMN IF NOT EXISTS nome_arquivo_origem TEXT;
    ALTER TABLE pedidos_importados ADD COLUMN IF NOT EXISTS demanda TEXT;
    ALTER TABLE pedidos_liberados_historico ADD COLUMN IF NOT EXISTS departamento TEXT;
    ALTER TABLE pedidos_liberados_historico ADD COLUMN IF NOT EXISTS demanda TEXT;
    ALTER TABLE pedidos_liberados_historico ADD COLUMN IF NOT EXISTS usuario_adicao TEXT;
    ALTER TABLE pedidos_liberados_historico ADD COLUMN IF NOT EXISTS usuario_ultima_modificacao TEXT;
    ALTER TABLE pedidos_liberados_historico ADD COLUMN IF NOT EXISTS ultima_modificacao_pedido TIMESTAMP DEFAULT NOW();
    ALTER TABLE pedidos_importados ADD COLUMN IF NOT EXISTS usuario_importacao_pedido TEXT;
    ALTER TABLE pedidos_removidos_importacao ADD COLUMN IF NOT EXISTS usuario_remocao TEXT;
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
# USUÁRIOS / LOGIN
# =========================================================


def hash_senha(senha):
    senha_txt = str(senha or "")
    return hashlib.sha256(("liberacao_pedidos|" + senha_txt).encode("utf-8")).hexdigest()


def garantir_admin_padrao():
    """Cria um ADMIN inicial se ainda não houver usuário cadastrado."""
    engine = get_engine_neon()
    if engine is None:
        return False

    try:
        with engine.begin() as conn:
            qtd = conn.execute(text("SELECT COUNT(*) FROM usuarios_sistema")).scalar() or 0
            if qtd == 0:
                conn.execute(
                    text(
                        """
                        INSERT INTO usuarios_sistema (usuario, senha_hash, nome, perfil, analista, ativo)
                        VALUES (:usuario, :senha_hash, :nome, 'ADMIN', NULL, TRUE)
                        """
                    ),
                    {
                        "usuario": "ADMIN",
                        "senha_hash": hash_senha("123"),
                        "nome": "Administrador",
                    },
                )
        return True
    except Exception as e:
        st.error("Erro ao criar usuário ADMIN padrão.")
        st.code(str(e))
        return False


def autenticar_usuario(usuario, senha):
    engine = get_engine_neon()
    if engine is None:
        return None

    try:
        with engine.begin() as conn:
            row = conn.execute(
                text(
                    """
                    SELECT id, usuario, nome, perfil, analista, ativo
                    FROM usuarios_sistema
                    WHERE UPPER(usuario) = UPPER(:usuario)
                      AND senha_hash = :senha_hash
                      AND ativo = TRUE
                    """
                ),
                {"usuario": limpar_texto(usuario), "senha_hash": hash_senha(senha)},
            ).mappings().first()
            return dict(row) if row else None
    except Exception:
        return None


def listar_usuarios_sistema():
    engine = get_engine_neon()
    if engine is None:
        return pd.DataFrame()

    try:
        with engine.begin() as conn:
            return pd.read_sql(
                """
                SELECT id, usuario, nome, perfil, analista, ativo, criado_em, atualizado_em
                FROM usuarios_sistema
                ORDER BY nome
                """,
                conn,
            )
    except Exception:
        return pd.DataFrame()


def criar_usuario_sistema(usuario, senha, nome, perfil, analista=None):
    engine = get_engine_neon()
    if engine is None:
        return False, "Neon não conectado."

    usuario = limpar_texto(usuario).upper()
    nome = limpar_texto(nome)
    perfil = limpar_texto(perfil).upper()
    analista = limpar_texto(analista) or None

    if not usuario or not senha or not nome:
        return False, "Preencha usuário, nome e senha."

    try:
        with engine.begin() as conn:
            conn.execute(
                text(
                    """
                    INSERT INTO usuarios_sistema (usuario, senha_hash, nome, perfil, analista, ativo)
                    VALUES (:usuario, :senha_hash, :nome, :perfil, :analista, TRUE)
                    """
                ),
                {
                    "usuario": usuario,
                    "senha_hash": hash_senha(senha),
                    "nome": nome,
                    "perfil": perfil,
                    "analista": analista,
                },
            )
        return True, "Usuário criado."
    except Exception as e:
        return False, str(e)


def alterar_senha_usuario(usuario_id, nova_senha):
    engine = get_engine_neon()
    if engine is None:
        return False, "Neon não conectado."
    if not nova_senha:
        return False, "Informe a nova senha."

    try:
        with engine.begin() as conn:
            conn.execute(
                text(
                    """
                    UPDATE usuarios_sistema
                    SET senha_hash = :senha_hash, atualizado_em = NOW()
                    WHERE id = :id
                    """
                ),
                {"senha_hash": hash_senha(nova_senha), "id": int(usuario_id)},
            )
        return True, "Senha alterada."
    except Exception as e:
        return False, str(e)


def definir_status_usuario(usuario_id, ativo):
    engine = get_engine_neon()
    if engine is None:
        return False, "Neon não conectado."

    try:
        with engine.begin() as conn:
            conn.execute(
                text(
                    """
                    UPDATE usuarios_sistema
                    SET ativo = :ativo, atualizado_em = NOW()
                    WHERE id = :id
                    """
                ),
                {"ativo": bool(ativo), "id": int(usuario_id)},
            )
        return True, "Usuário atualizado."
    except Exception as e:
        return False, str(e)


def usuario_e_admin(usuario_logado):
    return limpar_texto(usuario_logado.get("perfil", "")).upper() == "ADMIN"


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



def carregar_pedidos_empresa(data_ini, data_fim, departamentos, nome_arquivo, data_emissao=None):
    if not departamentos or not nome_arquivo:
        return pd.DataFrame()

    payload = {
        "data_inicial": data_ini.isoformat(),
        "data_final": data_fim.isoformat(),
        "departamentos": [limpar_texto(d).upper() for d in departamentos],
        "nome_arquivo": limpar_texto(nome_arquivo),
    }
    if data_emissao:
        data_emissao_txt = limpar_texto(data_emissao)
        if len(data_emissao_txt) == 8 and data_emissao_txt.isdigit():
            payload["data_emissao"] = f"{data_emissao_txt[0:4]}-{data_emissao_txt[4:6]}-{data_emissao_txt[6:8]}"
        else:
            payload["data_emissao"] = data_emissao_txt

    try:
        retorno = chamar_api("POST", "/pedidos", payload=payload)
        dados = retorno.get("dados", []) if isinstance(retorno, dict) else []
        df = pd.DataFrame(dados)
    except Exception as e:
        st.error("Erro ao buscar pedidos na API da empresa.")
        st.code(str(e))
        return pd.DataFrame()

    return preparar_df_pedidos(df)



def consultar_status_pedidos_api(pedidos):
    pedidos_limpos = [limpar_texto(p) for p in pedidos if limpar_texto(p)]
    if not pedidos_limpos:
        return pd.DataFrame()

    payload = {"pedidos": pedidos_limpos}
    try:
        retorno = chamar_api("POST", "/status-pedidos", payload=payload, timeout=120)
        dados = retorno.get("dados", []) if isinstance(retorno, dict) else []
        df = pd.DataFrame(dados)
        if not df.empty:
            df.columns = [str(c).strip().upper() for c in df.columns]
            if "STATUS_BANCO" in df.columns:
                df["STATUS_BANCO"] = df["STATUS_BANCO"].apply(normalizar_status)
        return df
    except Exception as e:
        st.error("Erro ao consultar status do pedido na API.")
        st.code(str(e))
        return pd.DataFrame()


# =========================================================
# SALVAR NO NEON
# =========================================================


def salvar_carga_no_neon(
    df_final,
    df_removidos,
    nome_carga,
    nome_arquivo,
    usuario_importacao,
    analista,
    departamento,
    data_ini,
    data_fim,
    demanda_importacao="",
    valor_informado=None,
):
    engine = get_engine_neon()
    if engine is None:
        st.error("Neon/PostgreSQL não configurado. Verifique os Secrets.")
        return False

    qtd_pedidos = int(df_final[COLUNA_PEDIDO].nunique()) if COLUNA_PEDIDO in df_final.columns else len(df_final)
    pedidos_liberados = df_final.loc[df_final["STATUS_BANCO"] == "LIBERADO", COLUNA_PEDIDO].nunique() if not df_final.empty else 0
    qtd_liberados = int(pedidos_liberados)
    qtd_nao_liberados = max(qtd_pedidos - qtd_liberados, 0)
    if df_final is not None and not df_final.empty and COLUNA_VALOR in df_final.columns:
        valor_total_calculado = 0.0
        for _, grupo_pedido in df_final.groupby([COLUNA_LOJA, COLUNA_FORNECEDOR, COLUNA_PEDIDO], dropna=False):
            valor_total_calculado += valor_pedido_tipo_fornecedor(grupo_pedido)
    else:
        valor_total_calculado = 0.0
    valor_informado_num = limpar_valor(valor_informado)
    diferenca_valor = None if valor_informado_num is None else float(valor_total_calculado - valor_informado_num)

    try:
        with engine.begin() as conn:
            carga_id = conn.execute(
                text(
                    """
                    INSERT INTO cargas_importacao
                    (nome_carga, nome_arquivo, usuario_importacao, qtd_pedidos, qtd_liberados, qtd_nao_liberados,
                     departamento, filtro_data_ini, filtro_data_fim, analista, demanda_importacao,
                     valor_informado, valor_total_calculado, diferenca_valor)
                    VALUES
                    (:nome_carga, :nome_arquivo, :usuario_importacao, :qtd_pedidos, :qtd_liberados, :qtd_nao_liberados,
                     :departamento, :filtro_data_ini, :filtro_data_fim, :analista, :demanda_importacao,
                     :valor_informado, :valor_total_calculado, :diferenca_valor)
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
                    "analista": analista,
                    "demanda_importacao": demanda_importacao,
                    "valor_informado": valor_informado_num,
                    "valor_total_calculado": valor_total_calculado,
                    "diferenca_valor": diferenca_valor,
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
                demanda = limpar_texto(row.get(COLUNA_DEMANDA))
                dados = row_json(row)

                conn.execute(
                    text(
                        """
                        INSERT INTO pedidos_importados
                        (carga_id, pedido, fornecedor, loja, valor, status_banco, cod_status_banco,
                         dados_planilha, departamento, nome_arquivo_origem, demanda, usuario_importacao_pedido)
                        VALUES
                        (:carga_id, :pedido, :fornecedor, :loja, :valor, :status_banco, :cod_status_banco,
                         CAST(:dados_planilha AS JSONB), :departamento, :nome_arquivo_origem, :demanda, :usuario_importacao_pedido)
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
                        "demanda": demanda,
                        "usuario_importacao_pedido": usuario_importacao,
                    },
                )



            # Histórico de liberados: salva 1 linha por pedido/filial.
            # Só o VALOR usa a regra do TIPO_FORNECEDOR; os demais campos continuam normais.
            df_liberados_hist = df_final[df_final["STATUS_BANCO"] == "LIBERADO"].copy()
            if not df_liberados_hist.empty:
                for (pedido_hist, loja_hist), grupo_hist in df_liberados_hist.groupby([COLUNA_PEDIDO, COLUNA_LOJA], dropna=False):
                    primeira_linha = grupo_hist.iloc[0]
                    pedido = limpar_texto(pedido_hist)
                    fornecedor = limpar_texto(primeira_linha.get(COLUNA_FORNECEDOR))
                    loja = limpar_texto(loja_hist) or "GERAL"
                    valor = valor_pedido_tipo_fornecedor(grupo_hist)
                    depto_row = limpar_texto(primeira_linha.get(COLUNA_DEPARTAMENTO)) or departamento
                    status_banco = limpar_texto(primeira_linha.get("STATUS_BANCO"))
                    cod_status_banco = limpar_texto(primeira_linha.get("COD_STATUS_BANCO"))
                    demanda = limpar_texto(primeira_linha.get(COLUNA_DEMANDA))
                    dados = row_json(primeira_linha)

                    conn.execute(
                        text(
                            """
                            INSERT INTO pedidos_liberados_historico
                            (pedido, fornecedor, loja, valor, status_banco, cod_status_banco,
                             primeira_carga, ultima_carga, dados_ultima_planilha, departamento, demanda,
                             usuario_adicao, usuario_ultima_modificacao, ultima_modificacao_pedido)
                            VALUES
                            (:pedido, :fornecedor, :loja, :valor, :status_banco, :cod_status_banco,
                             :nome_carga, :nome_carga, CAST(:dados_planilha AS JSONB), :departamento, :demanda,
                             :usuario_adicao, :usuario_ultima_modificacao, NOW())
                            ON CONFLICT (pedido, loja)
                            DO UPDATE SET
                                fornecedor = EXCLUDED.fornecedor,
                                valor = EXCLUDED.valor,
                                status_banco = EXCLUDED.status_banco,
                                cod_status_banco = EXCLUDED.cod_status_banco,
                                ultima_carga = EXCLUDED.ultima_carga,
                                ultima_atualizacao = NOW(),
                                ultima_modificacao_pedido = NOW(),
                                dados_ultima_planilha = EXCLUDED.dados_ultima_planilha,
                                departamento = EXCLUDED.departamento,
                                demanda = EXCLUDED.demanda,
                                usuario_adicao = COALESCE(pedidos_liberados_historico.usuario_adicao, EXCLUDED.usuario_adicao),
                                usuario_ultima_modificacao = EXCLUDED.usuario_ultima_modificacao
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
                            "demanda": demanda,
                            "usuario_adicao": usuario_importacao,
                            "usuario_ultima_modificacao": usuario_importacao,
                        },
                    )


            if df_removidos is not None and not df_removidos.empty:
                for _, row in df_removidos.iterrows():
                    conn.execute(
                        text(
                            """
                            INSERT INTO pedidos_removidos_importacao
                            (carga_id, nome_carga, nome_arquivo, usuario_importacao, usuario_remocao, pedido, fornecedor, loja,
                             departamento, status_banco, cod_status_banco, dados_planilha)
                            VALUES
                            (:carga_id, :nome_carga, :nome_arquivo, :usuario_importacao, :usuario_remocao, :pedido, :fornecedor, :loja,
                             :departamento, :status_banco, :cod_status_banco, CAST(:dados_planilha AS JSONB))
                            """
                        ),
                        {
                            "carga_id": carga_id,
                            "nome_carga": nome_carga,
                            "nome_arquivo": nome_arquivo,
                            "usuario_importacao": usuario_importacao,
                            "usuario_remocao": usuario_importacao,
                            "pedido": limpar_texto(row.get(COLUNA_PEDIDO)),
                            "fornecedor": limpar_texto(row.get(COLUNA_FORNECEDOR)),
                            "loja": limpar_texto(row.get(COLUNA_LOJA)) or "GERAL",
                            "departamento": limpar_texto(row.get(COLUNA_DEPARTAMENTO)) or departamento,
                            "status_banco": limpar_texto(row.get("STATUS_BANCO")),
                            "cod_status_banco": limpar_texto(row.get("COD_STATUS_BANCO")),
                            "dados_planilha": row_json(row),
                        },
                    )

        return True

    except Exception as e:
        st.error("Erro ao salvar no Neon/PostgreSQL.")
        st.code(str(e))
        return False



def salvar_substituicao_neon(dados):
    engine = get_engine_neon()
    if engine is None:
        st.error("Neon/PostgreSQL não configurado. Verifique os Secrets.")
        return False

    try:
        with engine.begin() as conn:
            conn.execute(
                text(
                    """
                    INSERT INTO substituicoes_pedidos
                    (nome_carga, nome_arquivo, usuario_registro, analista, departamento,
                     pedido_antigo, fornecedor_antigo, loja_antiga, status_antigo, cod_status_antigo,
                     pedido_novo, fornecedor_novo, loja_nova, status_novo, cod_status_novo,
                     motivo, observacao, dados_pedido_antigo, dados_pedido_novo)
                    VALUES
                    (:nome_carga, :nome_arquivo, :usuario_registro, :analista, :departamento,
                     :pedido_antigo, :fornecedor_antigo, :loja_antiga, :status_antigo, :cod_status_antigo,
                     :pedido_novo, :fornecedor_novo, :loja_nova, :status_novo, :cod_status_novo,
                     :motivo, :observacao, CAST(:dados_pedido_antigo AS JSONB), CAST(:dados_pedido_novo AS JSONB))
                    """
                ),
                dados,
            )
        return True
    except Exception as e:
        st.error("Erro ao salvar substituição no Neon.")
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
                    analista,
                    departamento,
                    demanda_importacao,
                    valor_informado,
                    valor_total_calculado,
                    diferenca_valor,
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
                    demanda,
                    valor,
                    status_banco,
                    COALESCE(usuario_adicao, '') AS usuario_adicao,
                    COALESCE(usuario_ultima_modificacao, usuario_adicao, '') AS usuario_ultima_modificacao,
                    COALESCE(ultima_modificacao_pedido, ultima_atualizacao) AS ultima_modificacao_pedido,
                    COALESCE(
                        dados_ultima_planilha ->> 'DATA_EMISSAO_ARQUIVO',
                        dados_ultima_planilha ->> 'DT_EMISSAO',
                        ''
                    ) AS data_emissao_pedido,
                    ultima_atualizacao AS data_historico,
                    COALESCE(
                        dados_ultima_planilha ->> 'OBSERVACAO',
                        dados_ultima_planilha ->> 'OBSERVAÇÃO',
                        dados_ultima_planilha ->> 'MOTIVO',
                        ''
                    ) AS observacao
                FROM pedidos_liberados_historico
                ORDER BY ultima_atualizacao DESC
                """,
                conn,
            )
    except Exception:
        return pd.DataFrame()



def carregar_historico_substituicoes_neon(nome_arquivo=None):
    engine = get_engine_neon()
    if engine is None:
        return pd.DataFrame()

    try:
        with engine.begin() as conn:
            if nome_arquivo:
                return pd.read_sql(
                    text(
                        """
                        SELECT
                            data_registro,
                            usuario_registro,
                            analista,
                            departamento,
                            nome_arquivo,
                            pedido_antigo,
                            fornecedor_antigo,
                            loja_antiga,
                            status_antigo,
                            cod_status_antigo,
                            pedido_novo,
                            fornecedor_novo,
                            loja_nova,
                            status_novo,
                            cod_status_novo,
                            motivo,
                            observacao
                        FROM substituicoes_pedidos
                        WHERE nome_arquivo = :nome_arquivo
                        ORDER BY data_registro DESC
                        """
                    ),
                    conn,
                    params={"nome_arquivo": nome_arquivo},
                )

            return pd.read_sql(
                """
                SELECT
                    data_registro,
                    usuario_registro,
                    analista,
                    departamento,
                    nome_arquivo,
                    pedido_antigo,
                    fornecedor_antigo,
                    loja_antiga,
                    status_antigo,
                    cod_status_antigo,
                    pedido_novo,
                    fornecedor_novo,
                    loja_nova,
                    status_novo,
                    cod_status_novo,
                    motivo,
                    observacao
                FROM substituicoes_pedidos
                ORDER BY data_registro DESC
                """,
                conn,
            )
    except Exception:
        return pd.DataFrame()




def montar_historico_liberados_display(df):
    if df is None or df.empty:
        return pd.DataFrame(columns=["Pedido", "Data", "Fornecedor", "Filial", "Departamento", "Tipo", "Valor", "Status", "Observação", "Adicionado por", "Última modificação"] )

    df = df.copy()
    for col in ["pedido", "fornecedor", "loja", "departamento", "demanda", "valor", "status_banco", "observacao", "data_emissao_pedido", "data_historico", "usuario_adicao", "usuario_ultima_modificacao", "ultima_modificacao_pedido"]:
        if col not in df.columns:
            df[col] = ""

    display = pd.DataFrame({
        "Pedido": df["pedido"].apply(limpar_texto),
        "Data": df.apply(lambda row: formatar_data_ddmmyyyy(obter_data_historico(row)), axis=1),
        "Fornecedor": df["fornecedor"].apply(limpar_texto),
        "Filial": df["loja"].apply(limpar_texto),
        "Departamento": df["departamento"].apply(limpar_texto),
        "Tipo": df["demanda"].apply(lambda v: limpar_texto(v).replace("DEMANDA ", "").title().replace("Juridica", "Jurídica")),
        "Valor": df["valor"].apply(lambda v: formatar_moeda(limpar_valor(v) or 0)),
        "Status": df["status_banco"].apply(status_icone),
        "Observação": df["observacao"].apply(limpar_texto),
        "Adicionado por": df["usuario_adicao"].apply(limpar_texto),
        "Última modificação": df["ultima_modificacao_pedido"].apply(lambda v: formatar_data_hora_ddmmyyyy(v)),
    })

    return display


def aplicar_filtros_historico(df, filtro_pedido="", filtro_fornecedor="", filtro_depto="", filtro_filial=""):
    df_view = df.copy()

    if filtro_pedido:
        df_view = df_view[df_view["pedido"].astype(str).str.contains(filtro_pedido, case=False, na=False)]
    if filtro_fornecedor:
        df_view = df_view[df_view["fornecedor"].astype(str).str.contains(filtro_fornecedor, case=False, na=False)]
    if filtro_depto:
        df_view = df_view[df_view["departamento"].astype(str).str.contains(filtro_depto, case=False, na=False)]
    if filtro_filial:
        df_view = df_view[df_view["loja"].astype(str).str.contains(filtro_filial, case=False, na=False)]

    return df_view


def montar_base_substituicao_historico(df_hist):
    """Converte o histórico de liberados em uma base compatível com a tela de substituição."""
    if df_hist is None or df_hist.empty:
        return pd.DataFrame()

    df = df_hist.copy()
    for col in ["pedido", "fornecedor", "loja", "departamento", "demanda", "valor", "status_banco", "observacao"]:
        if col not in df.columns:
            df[col] = ""

    base = pd.DataFrame({
        COLUNA_PEDIDO: df["pedido"].apply(limpar_texto),
        COLUNA_FORNECEDOR: df["fornecedor"].apply(limpar_texto),
        COLUNA_LOJA: df["loja"].apply(limpar_texto).replace("", "GERAL"),
        COLUNA_DEPARTAMENTO: df["departamento"].apply(limpar_texto),
        COLUNA_DEMANDA: df["demanda"].apply(limpar_texto),
        COLUNA_VALOR: df["valor"].apply(limpar_valor),
        "STATUS_BANCO": df["status_banco"].apply(normalizar_status),
        "COD_STATUS_BANCO": "",
        COLUNA_OBSERVACAO: df["observacao"].apply(limpar_texto),
    })

    base = base[base[COLUNA_PEDIDO] != ""].drop_duplicates().copy()
    return base


def montar_historico_substituicoes_display(df_subs):
    if df_subs is None or df_subs.empty:
        return pd.DataFrame(columns=[
            "Data", "Usuário", "Pedido antigo", "Fornecedor antigo", "Filial antiga", "Status antigo",
            "Pedido novo", "Fornecedor novo", "Filial nova", "Status novo", "Motivo", "Observação"
        ])

    df = df_subs.copy()
    for col in [
        "data_registro", "usuario_registro", "pedido_antigo", "fornecedor_antigo", "loja_antiga", "status_antigo",
        "pedido_novo", "fornecedor_novo", "loja_nova", "status_novo", "motivo", "observacao"
    ]:
        if col not in df.columns:
            df[col] = ""

    display = pd.DataFrame({
        "Data": df["data_registro"].apply(formatar_data_hora_ddmmyyyy),
        "Usuário": df["usuario_registro"].apply(limpar_texto),
        "Pedido antigo": df["pedido_antigo"].apply(limpar_texto),
        "Fornecedor antigo": df["fornecedor_antigo"].apply(limpar_texto),
        "Filial antiga": df["loja_antiga"].apply(limpar_texto),
        "Status antigo": df["status_antigo"].apply(status_icone),
        "Pedido novo": df["pedido_novo"].apply(limpar_texto),
        "Fornecedor novo": df["fornecedor_novo"].apply(limpar_texto),
        "Filial nova": df["loja_nova"].apply(limpar_texto),
        "Status novo": df["status_novo"].apply(status_icone),
        "Motivo": df["motivo"].apply(limpar_texto),
        "Observação": df["observacao"].apply(limpar_texto),
    })

    return display


# =========================================================
# APP
# =========================================================

st.markdown('<div class="titulo-principal">📦 Liberação de Pedidos</div>', unsafe_allow_html=True)
ok_neon, msg_neon = testar_conexao_neon()
ok_api, msg_api = testar_conexao_api()

if not ok_neon:
    st.error("Neon não conectado. Não é possível fazer login nem gravar histórico.")
    st.caption(msg_neon)
    st.stop()

criar_tabelas_neon()
garantir_admin_padrao()

if "usuario_logado" not in st.session_state:
    st.subheader("Entrar no sistema")
    st.caption("O primeiro acesso padrão é usuário ADMIN e senha 123. Depois altere a senha na aba Usuários.")
    with st.form("form_login"):
        login_usuario = st.text_input("Usuário", value="ADMIN")
        login_senha = st.text_input("Senha", type="password")
        entrar = st.form_submit_button("Entrar", type="primary")

    if entrar:
        usuario_auth = autenticar_usuario(login_usuario, login_senha)
        if usuario_auth:
            st.session_state["usuario_logado"] = usuario_auth
            st.rerun()
        else:
            st.error("Usuário ou senha inválidos, ou usuário inativo.")
    st.stop()

usuario_logado = st.session_state["usuario_logado"]
usuario_importacao = limpar_texto(usuario_logado.get("nome")) or limpar_texto(usuario_logado.get("usuario"))
perfil_logado = limpar_texto(usuario_logado.get("perfil", "ANALISTA")).upper()
admin_logado = usuario_e_admin(usuario_logado)

col_user_info, col_sair = st.columns([5, 1])
with col_user_info:
    st.caption(f"Logado como: **{usuario_importacao}** | Perfil: **{perfil_logado}**")
with col_sair:
    if st.button("Sair"):
        st.session_state.pop("usuario_logado", None)
        st.rerun()

if not ok_api:
    st.warning("API da empresa não conectada. Você ainda consegue consultar o histórico, mas não consegue buscar novos pedidos.")
    st.caption(msg_api)

nomes_abas = [
    "🔎 Consultar banco da empresa",
    "📚 Histórico Neon",
    "⚙️ Configuração",
]
if admin_logado:
    nomes_abas.append("👥 Usuários")

abas_app = st.tabs(nomes_abas)
aba_consulta = abas_app[0]
aba_historico = abas_app[1]
aba_config = abas_app[2]
aba_usuarios = abas_app[3] if admin_logado else None


with aba_consulta:
    st.subheader("1. Filtros para buscar os arquivos importados")

    hoje = date.today()
    col_data1, col_data2, col_analista = st.columns([1, 1, 1.2])

    with col_data1:
        data_ini = st.date_input("Data inicial", value=hoje - timedelta(days=30), format="DD/MM/YYYY")

    with col_data2:
        data_fim = st.date_input("Data final", value=hoje, format="DD/MM/YYYY")

    with col_analista:
        analista_padrao_usuario = limpar_texto(usuario_logado.get("analista", ""))
        idx_analista = list(ANALISTAS.keys()).index(analista_padrao_usuario) if analista_padrao_usuario in ANALISTAS else 0
        analista = st.selectbox("Analista", options=list(ANALISTAS.keys()), index=idx_analista)

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
            st.session_state.pop("df_pedidos_empresa", None)
            st.session_state.pop("pedidos_remover_labels", None)

    df_arquivos = st.session_state.get("df_arquivos_empresa", pd.DataFrame())

    if not df_arquivos.empty:
        st.subheader("2. Escolha o(s) arquivo(s) importado(s)")

        df_arquivos_opcoes = df_arquivos.reset_index(drop=True).copy()
        tabela_arquivos = montar_tabela_arquivos_display(df_arquivos_opcoes)
        st.dataframe(tabela_arquivos, use_container_width=True, hide_index=True)

        # Usa exatamente a mesma data exibida na tabela também no multiselect.
        # Assim evita aparecer 06/10/2026 no seletor quando a tabela mostra 10/06/2026.
        if not tabela_arquivos.empty and "Data de emissão" in tabela_arquivos.columns:
            df_arquivos_opcoes["DATA_EMISSAO_DISPLAY"] = tabela_arquivos["Data de emissão"].astype(str).tolist()
        else:
            df_arquivos_opcoes["DATA_EMISSAO_DISPLAY"] = df_arquivos_opcoes.get("DATA_EMISSAO", "").apply(formatar_data_protheus)

        opcoes_arquivos = []
        mapa_opcoes_arquivos = {}

        for idx, row in df_arquivos_opcoes.iterrows():
            filial_label = limpar_texto(row.get("FILIAL")) or "SEM FILIAL"
            nome_label = limpar_texto(row.get(COLUNA_NOME_ARQUIVO))
            data_label = limpar_texto(row.get("DATA_EMISSAO_DISPLAY")) or "SEM DATA"
            valor_label = formatar_moeda(valor_loja_naomisto_total(row))
            label = f"{filial_label} | {nome_label} | {data_label} | {valor_label}"
            opcoes_arquivos.append(label)
            mapa_opcoes_arquivos[label] = idx

        arquivos_selecionados_labels = st.multiselect(
            "Nome do(s) arquivo(s) importado(s)",
            options=opcoes_arquivos,
            default=[],
            help="Selecione um ou mais arquivos para carregar na mesma importação.",
        )

        if arquivos_selecionados_labels:
            linhas_selecionadas = [
                df_arquivos_opcoes.iloc[mapa_opcoes_arquivos[label]]
                for label in arquivos_selecionados_labels
                if label in mapa_opcoes_arquivos
            ]
            df_linhas_selecionadas = pd.DataFrame(linhas_selecionadas)

            nomes_arquivos = []
            for _, row in df_linhas_selecionadas.iterrows():
                nome_tmp = limpar_texto(row.get(COLUNA_NOME_ARQUIVO))
                if nome_tmp and nome_tmp not in nomes_arquivos:
                    nomes_arquivos.append(nome_tmp)

            nomes_arquivo_label = "; ".join(nomes_arquivos)
            valor_total_arquivos = float(sum(valor_loja_naomisto_total(row) for _, row in df_linhas_selecionadas.iterrows()))

            datas_raw = []
            datas_display = []
            for _, row in df_linhas_selecionadas.iterrows():
                raw = limpar_texto(row.get("DATA_EMISSAO"))
                display = limpar_texto(row.get("DATA_EMISSAO_DISPLAY")) or formatar_data_protheus(raw)
                if raw and raw not in datas_raw:
                    datas_raw.append(raw)
                if display and display not in datas_display:
                    datas_display.append(display)

            if len(datas_display) == 1:
                data_emissao_label = datas_display[0]
            elif len(datas_display) > 1:
                data_emissao_label = f"{datas_display[0]} a {datas_display[-1]}"
            else:
                data_emissao_label = "-"

            st.subheader("3. Classificação antes de importar")

            col_tipo_demanda, col_data_emissao, col_valor_demanda = st.columns([1.1, 1, 1])
            with col_tipo_demanda:
                tipo_demanda_importacao = st.selectbox(
                    "Esse(s) arquivo(s) é/são de qual tipo?",
                    options=DEMANDAS,
                    format_func=lambda x: x.replace("DEMANDA ", "").title().replace("Juridica", "Jurídica"),
                )

            with col_data_emissao:
                st.text_input("Data de emissão", value=data_emissao_label, disabled=True)

            with col_valor_demanda:
                st.metric("Valor Total", formatar_moeda(valor_total_arquivos))

            tipo_demanda_resumido = tipo_demanda_importacao.replace("DEMANDA ", "")
            observacao_sugerida = nomes_arquivo_label
            observacao_importacao = st.text_input(
                "Observação",
                value=observacao_sugerida,
                help="Essa observação será gravada na coluna OBSERVAÇÃO dos pedidos e no histórico do Neon.",
            )
            nome_carga = limpar_texto(observacao_importacao) or (
                f"{analista} | {tipo_demanda_resumido} | {nomes_arquivo_label} | Emissão: {data_emissao_label}"
            )

            if valor_total_arquivos <= 0:
                st.warning("A query retornou valor zerado para o(s) arquivo(s)/departamento(s) selecionado(s).")

            if st.button("📥 Carregar pedidos dos arquivos selecionados", type="primary"):
                with st.spinner("Carregando pedidos dos arquivos selecionados pela API..."):
                    lista_dfs = []
                    for _, linha_arquivo in df_linhas_selecionadas.iterrows():
                        nome_arquivo = limpar_texto(linha_arquivo.get(COLUNA_NOME_ARQUIVO))
                        data_emissao_api = limpar_texto(linha_arquivo.get("DATA_EMISSAO"))
                        df_pedidos_arquivo = carregar_pedidos_empresa(
                            data_ini,
                            data_fim,
                            departamentos_selecionados,
                            nome_arquivo,
                            data_emissao=data_emissao_api,
                        )
                        if not df_pedidos_arquivo.empty:
                            df_pedidos_arquivo[COLUNA_NOME_ARQUIVO] = nome_arquivo
                            df_pedidos_arquivo["DATA_EMISSAO_ARQUIVO"] = data_emissao_api
                            lista_dfs.append(df_pedidos_arquivo)

                    if lista_dfs:
                        df_pedidos = pd.concat(lista_dfs, ignore_index=True).drop_duplicates().copy()
                    else:
                        df_pedidos = pd.DataFrame()

                    if not df_pedidos.empty:
                        df_pedidos[COLUNA_DEMANDA] = tipo_demanda_importacao
                        df_pedidos[COLUNA_OBSERVACAO] = nome_carga
                        df_pedidos["MOTIVO"] = nome_carga  # compatibilidade com histórico anterior
                        df_pedidos["VALOR_INFORMADO_IMPORTACAO"] = float(valor_total_arquivos)
                        df_pedidos["VALOR_REFERENCIA_ARQUIVO_QUERY"] = float(valor_total_arquivos)
                        df_pedidos["TIPO_IMPORTACAO"] = tipo_demanda_importacao

                    st.session_state["df_pedidos_empresa"] = df_pedidos
                    st.session_state["nome_arquivo_selecionado"] = nomes_arquivo_label
                    st.session_state["nomes_arquivos_selecionados"] = nomes_arquivos
                    st.session_state["nome_carga"] = nome_carga
                    st.session_state["data_ini"] = data_ini
                    st.session_state["data_fim"] = data_fim
                    st.session_state["analista"] = analista
                    st.session_state["departamentos_selecionados"] = departamentos_selecionados
                    st.session_state["departamento"] = departamento_carga
                    st.session_state["tipo_demanda_importacao"] = tipo_demanda_importacao
                    st.session_state["valor_informado_importacao"] = float(valor_total_arquivos)
    elif buscar:
        st.warning("Nenhum arquivo encontrado para os filtros selecionados.")

    df_base = st.session_state.get("df_pedidos_empresa", pd.DataFrame())
    df_base = preparar_df_pedidos(df_base)

    if not df_base.empty:
        tipo_demanda_importacao = st.session_state.get("tipo_demanda_importacao", "")
        valor_informado_importacao = limpar_valor(st.session_state.get("valor_informado_importacao", 0)) or 0.0

        if tipo_demanda_importacao:
            df_base[COLUNA_DEMANDA] = tipo_demanda_importacao
            df_base["TIPO_IMPORTACAO"] = tipo_demanda_importacao
            df_base["VALOR_INFORMADO_IMPORTACAO"] = float(valor_informado_importacao)
            observacao_atual = limpar_texto(st.session_state.get("nome_carga", ""))
            if COLUNA_OBSERVACAO not in df_base.columns:
                df_base[COLUNA_OBSERVACAO] = observacao_atual
            else:
                df_base[COLUNA_OBSERVACAO] = df_base[COLUNA_OBSERVACAO].replace("", observacao_atual)
            df_base["MOTIVO"] = df_base[COLUNA_OBSERVACAO]

        st.divider()
        st.subheader("4. Remover pedidos errados desta importação")
        st.caption("Use esse campo quando algum pedido entrou por engano. Ele será retirado desta importação, mas ficará registrado como removido quando você salvar no Neon.")

        fornecedores_remover = lista_fornecedores(df_base)
        fornecedor_remover_sel = st.selectbox("Filtrar fornecedor para selecionar pedido a remover", options=fornecedores_remover)
        filtro_fornecedor_remover = "" if fornecedor_remover_sel == "Todos" else fornecedor_remover_sel
        mapa_remover, opcoes_remover = montar_opcoes_pedido(df_base, filtro_fornecedor_remover)
        labels_remover = st.multiselect(
            "Tem algum pedido que deseja remover desta importação?",
            options=opcoes_remover,
            default=[],
            key="pedidos_remover_labels",
        )
        pedidos_remover = sorted({mapa_remover[label] for label in labels_remover if label in mapa_remover})

        if pedidos_remover:
            st.warning(f"{len(pedidos_remover)} pedido(s) selecionado(s) para remover desta importação: {', '.join(pedidos_remover)}")

        df_removidos = df_base[df_base[COLUNA_PEDIDO].isin(pedidos_remover)].copy()
        df_final = df_base[~df_base[COLUNA_PEDIDO].isin(pedidos_remover)].copy()

        if not df_removidos.empty:
            st.markdown("**Pedidos removidos desta importação**")
            tabela_removidos = montar_tabela_planilha(df_removidos)
            st.dataframe(tabela_removidos, use_container_width=True, hide_index=True)

        st.divider()
        st.subheader("5. Pedidos encontrados")

        total_pedidos = df_final[COLUNA_PEDIDO].nunique() if COLUNA_PEDIDO in df_final.columns else len(df_final)
        total_liberados = df_final.loc[df_final["STATUS_BANCO"] == "LIBERADO", COLUNA_PEDIDO].nunique() if "STATUS_BANCO" in df_final.columns else 0
        total_bloqueados = df_final.loc[df_final["STATUS_BANCO"] == "BLOQUEADO", COLUNA_PEDIDO].nunique() if "STATUS_BANCO" in df_final.columns else 0
        total_removidos = df_removidos[COLUNA_PEDIDO].nunique() if not df_removidos.empty and COLUNA_PEDIDO in df_removidos.columns else 0
        if COLUNA_VALOR in df_final.columns and not df_final.empty:
            total_valor = 0.0
            for _, grupo_pedido in df_final.groupby([COLUNA_LOJA, COLUNA_FORNECEDOR, COLUNA_PEDIDO], dropna=False):
                total_valor += valor_pedido_tipo_fornecedor(grupo_pedido)
        else:
            total_valor = 0.0

        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric("Pedidos", int(total_pedidos))
        m2.metric("Liberados", int(total_liberados))
        m3.metric("Bloqueados", int(total_bloqueados))
        m4.metric("Removidos", int(total_removidos))
        m5.metric("Valor Total", formatar_moeda(total_valor))

        abas_demandas = st.tabs(["Demanda Puxada", "Demanda Empurrada", "Demanda Venda Jurídica"])
        tabelas_excel = {}

        for demanda_nome, aba in zip(DEMANDAS, abas_demandas[:3]):
            with aba:
                df_demanda = df_final[df_final[COLUNA_DEMANDA] == demanda_nome].copy()
                tabela_planilha_export = montar_tabela_planilha(df_demanda)
                tabelas_excel[demanda_nome] = tabela_planilha_export

                st.markdown(f"### {demanda_nome.title()}")
                if df_demanda.empty:
                    st.info("Nenhum pedido nesta demanda para os filtros selecionados.")
                else:
                    st.markdown("**Resumo por departamento**")
                    resumo_departamento = montar_resumo_departamento(df_demanda)
                    st.dataframe(resumo_departamento, use_container_width=True, hide_index=True)

                    departamentos_demanda = sorted({
                        limpar_texto(v) or "SEM DEPARTAMENTO"
                        for v in df_demanda.get(COLUNA_DEPARTAMENTO, pd.Series(dtype=str)).dropna().tolist()
                    })

                    if len(departamentos_demanda) > 1:
                        departamentos_visiveis = st.multiselect(
                            "Filtrar departamento nesta aba",
                            options=departamentos_demanda,
                            default=departamentos_demanda,
                            key=f"filtro_depto_{demanda_nome}",
                        )
                    else:
                        departamentos_visiveis = departamentos_demanda

                    df_demanda_view = df_demanda.copy()
                    if departamentos_visiveis:
                        depto_norm = df_demanda_view[COLUNA_DEPARTAMENTO].apply(limpar_texto).replace("", "SEM DEPARTAMENTO")
                        df_demanda_view = df_demanda_view[depto_norm.isin(departamentos_visiveis)].copy()

                    tabela_planilha = montar_tabela_planilha(df_demanda_view)
                    st.dataframe(tabela_planilha, use_container_width=True, hide_index=True)

                    with st.expander("Ver detalhes dos produtos"):
                        st.dataframe(df_demanda_view, use_container_width=True, hide_index=True)

        st.divider()
        col_baixar, col_salvar = st.columns([1, 1])

        with col_baixar:
            df_subs_arquivo = carregar_historico_substituicoes_neon(st.session_state.get("nome_arquivo_selecionado", "")) if ok_neon else pd.DataFrame()
            excel = gerar_excel_abas(tabelas_excel, df_removidos=df_removidos, df_substituicoes=df_subs_arquivo)
            st.download_button(
                "📥 Baixar Excel com abas",
                data=excel,
                file_name=f"{st.session_state.get('nome_carga', 'pedidos')}.xlsx".replace("/", "-"),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        with col_salvar:
            if st.button("💾 Salvar importação", type="primary", use_container_width=True):
                if not ok_neon:
                    st.error("Neon não conectado. Verifique os Secrets.")
                else:
                    sucesso = salvar_carga_no_neon(
                        df_final=df_final,
                        df_removidos=df_removidos,
                        nome_carga=st.session_state.get("nome_carga", "Carga sem nome"),
                        nome_arquivo=st.session_state.get("nome_arquivo_selecionado", ""),
                        usuario_importacao=usuario_importacao or st.session_state.get("analista", ""),
                        analista=st.session_state.get("analista", analista),
                        departamento=st.session_state.get("departamento", ""),
                        data_ini=st.session_state.get("data_ini", data_ini),
                        data_fim=st.session_state.get("data_fim", data_fim),
                        demanda_importacao=st.session_state.get("tipo_demanda_importacao", ""),
                        valor_informado=st.session_state.get("valor_informado_importacao", None),
                    )
                    if sucesso:
                        st.success("Importação salva com sucesso.")


with aba_historico:
    if not ok_neon:
        st.warning("Conecte para visualizar o histórico.")
    else:
        st.subheader("Pedidos liberados mantidos no histórico")
        df_hist = carregar_historico_liberados_neon()
        df_hist_base = df_hist.copy()

        if df_hist.empty:
            st.info("Nenhum pedido liberado salvo ainda.")
        else:
            df_hist = df_hist.copy()
            df_hist["_data_filtro"] = df_hist.apply(obter_data_historico, axis=1)

            st.markdown("**Filtros do histórico**")
            datas_validas = [d for d in df_hist["_data_filtro"].dropna().tolist()]
            data_padrao_ini = min(datas_validas) if datas_validas else date.today()
            data_padrao_fim = max(datas_validas) if datas_validas else date.today()

            col_hist_1, col_hist_2, col_hist_3, col_hist_4, col_hist_5, col_hist_6, col_hist_7 = st.columns([1, 1.2, 1, 1, 1, 1, 1])
            with col_hist_1:
                filtro_pedido_hist = st.text_input("Pedido", key="hist_filtro_pedido")
            with col_hist_2:
                filtro_fornecedor_hist = st.text_input("Fornecedor", key="hist_filtro_fornecedor")
            with col_hist_3:
                filtro_usuario_hist = st.text_input("Usuário", key="hist_filtro_usuario")
            with col_hist_4:
                filtro_depto_hist = st.text_input("Departamento", key="hist_filtro_depto")
            with col_hist_5:
                filtro_filial_hist = st.text_input("Filial", key="hist_filtro_filial")
            with col_hist_6:
                data_hist_ini = st.date_input("Data inicial", value=data_padrao_ini, format="DD/MM/YYYY", key="hist_data_ini")
            with col_hist_7:
                data_hist_fim = st.date_input("Data final", value=data_padrao_fim, format="DD/MM/YYYY", key="hist_data_fim")

            if data_hist_ini > data_hist_fim:
                st.error("A data inicial não pode ser maior que a data final.")
                st.stop()

            df_hist = df_hist[
                df_hist["_data_filtro"].notna()
                & (df_hist["_data_filtro"] >= data_hist_ini)
                & (df_hist["_data_filtro"] <= data_hist_fim)
            ].copy()

            if filtro_pedido_hist:
                df_hist = df_hist[df_hist["pedido"].astype(str).str.contains(filtro_pedido_hist, case=False, na=False)]
            if filtro_fornecedor_hist:
                df_hist = df_hist[df_hist["fornecedor"].astype(str).str.contains(filtro_fornecedor_hist, case=False, na=False)]
            if filtro_usuario_hist:
                filtro_usuario_txt = str(filtro_usuario_hist)
                df_hist = df_hist[
                    df_hist.get("usuario_adicao", pd.Series(dtype=str)).astype(str).str.contains(filtro_usuario_txt, case=False, na=False)
                    | df_hist.get("usuario_ultima_modificacao", pd.Series(dtype=str)).astype(str).str.contains(filtro_usuario_txt, case=False, na=False)
                ]
            if filtro_depto_hist:
                df_hist = df_hist[df_hist["departamento"].astype(str).str.contains(filtro_depto_hist, case=False, na=False)]
            if filtro_filial_hist:
                df_hist = df_hist[df_hist["loja"].astype(str).str.contains(filtro_filial_hist, case=False, na=False)]

            abas_hist = st.tabs(["Puxada", "Empurrada", "Venda Jurídica"])
            excel_hist_abas = {}

            for demanda_nome, aba in zip(DEMANDAS, abas_hist):
                with aba:
                    titulo = demanda_nome.replace("DEMANDA ", "").title().replace("Juridica", "Jurídica")
                    st.markdown(f"### {titulo}")

                    df_demanda_hist = df_hist[
                        df_hist["demanda"].astype(str).str.upper().str.strip() == demanda_nome
                    ].copy()

                    if df_demanda_hist.empty:
                        st.info(f"Nenhum pedido liberado salvo em {titulo}.")
                        excel_hist_abas[titulo] = pd.DataFrame()
                        continue

                    df_view = df_demanda_hist.copy()

                    df_display = montar_historico_liberados_display(df_view)
                    excel_hist_abas[titulo] = df_display

                    st.dataframe(df_display, use_container_width=True, hide_index=True)

            excel_hist = gerar_excel_abas(excel_hist_abas)
            st.download_button(
                "📥 Baixar histórico de liberados",
                data=excel_hist,
                file_name="historico_pedidos_liberados.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        st.divider()
        st.markdown("### Registrar substituição")
        st.caption("Use esta área para registrar uma substituição")

        df_base_sub = montar_base_substituicao_historico(df_hist_base if 'df_hist_base' in locals() else df_hist)

        if df_base_sub.empty:
            st.info("Não há pedidos no histórico para usar como pedido antigo.")
        else:
            st.caption("Informe o código do fornecedor, o pedido antigo e o pedido novo. Não precisa selecionar em lista grande.")

            col_fornecedor_cod, col_pedido_antigo_cod = st.columns([1, 1])
            with col_fornecedor_cod:
                codigo_fornecedor_sub_hist = st.text_input(
                    "Código do fornecedor",
                    placeholder="Ex: 000603 ou 01/000603",
                    key="hist_sub_codigo_fornecedor",
                )
            with col_pedido_antigo_cod:
                pedido_antigo_hist = st.text_input(
                    "Pedido antigo",
                    placeholder="Digite somente o código do pedido antigo",
                    key="hist_sub_pedido_antigo_codigo",
                )

            codigo_fornecedor_sub_hist = limpar_texto(codigo_fornecedor_sub_hist)
            pedido_antigo_hist = limpar_texto(pedido_antigo_hist).upper()

            df_antigos_filtrados = df_base_sub.copy()

            if codigo_fornecedor_sub_hist:
                df_antigos_filtrados = df_antigos_filtrados[
                    df_antigos_filtrados[COLUNA_FORNECEDOR].astype(str).str.contains(
                        codigo_fornecedor_sub_hist,
                        case=False,
                        na=False,
                    )
                ].copy()

            if pedido_antigo_hist:
                df_antigos_filtrados = df_antigos_filtrados[
                    df_antigos_filtrados[COLUNA_PEDIDO].astype(str).str.upper().str.strip() == pedido_antigo_hist
                ].copy()

            if codigo_fornecedor_sub_hist or pedido_antigo_hist:
                if df_antigos_filtrados.empty:
                    st.warning("Nenhum pedido antigo encontrado com esse fornecedor/pedido no histórico.")
                else:
                    st.markdown("**Pedido antigo encontrado**")
                    colunas_preview_antigo = [
                        COLUNA_PEDIDO,
                        COLUNA_FORNECEDOR,
                        COLUNA_LOJA,
                        COLUNA_DEPARTAMENTO,
                        COLUNA_VALOR,
                        "STATUS_BANCO",
                        COLUNA_OBSERVACAO,
                    ]
                    colunas_preview_antigo = [c for c in colunas_preview_antigo if c in df_antigos_filtrados.columns]
                    df_preview_antigo = df_antigos_filtrados[colunas_preview_antigo].drop_duplicates().copy()
                    if COLUNA_VALOR in df_preview_antigo.columns:
                        df_preview_antigo[COLUNA_VALOR] = df_preview_antigo[COLUNA_VALOR].apply(lambda v: formatar_moeda(limpar_valor(v) or 0))
                    if "STATUS_BANCO" in df_preview_antigo.columns:
                        df_preview_antigo["STATUS"] = df_preview_antigo["STATUS_BANCO"].apply(status_icone)
                        df_preview_antigo = df_preview_antigo.drop(columns=["STATUS_BANCO"])
                    df_preview_antigo = df_preview_antigo.rename(columns={
                        COLUNA_PEDIDO: "Pedido",
                        COLUNA_FORNECEDOR: "Fornecedor",
                        COLUNA_LOJA: "Filial",
                        COLUNA_DEPARTAMENTO: "Departamento",
                        COLUNA_VALOR: "Valor",
                        COLUNA_OBSERVACAO: "Observação",
                    })
                    st.dataframe(df_preview_antigo, use_container_width=True, hide_index=True)

            col_sub1, col_sub2 = st.columns([1, 1])
            with col_sub1:
                pedido_novo_hist = st.text_input(
                    "Pedido novo",
                    placeholder="Digite somente o código do pedido novo",
                    key="hist_sub_pedido_novo",
                )
            with col_sub2:
                motivo_sub_hist = st.selectbox(
                    "Motivo fixo",
                    options=["ERRO DE CADASTRO", "ERRO DE TABELA", "ERRO DE FAT. FORNECEDOR", "FAT. EM DUPLICIDADE"],
                    key="hist_sub_motivo",
                )

            pedido_novo_hist = limpar_texto(pedido_novo_hist).upper()

            observacao_sub_hist = st.text_area(
                "Observação",
                placeholder="Digite a observação da substituição",
                key="hist_sub_observacao",
            )

            if pedido_novo_hist and st.button("🔎 Consultar status do novo pedido", key="hist_sub_consultar_status"):
                df_status_novo_hist = consultar_status_pedidos_api([pedido_novo_hist])
                st.session_state["df_status_pedido_novo_hist"] = df_status_novo_hist

            df_status_novo_hist = st.session_state.get("df_status_pedido_novo_hist", pd.DataFrame())
            if pedido_novo_hist and not df_status_novo_hist.empty:
                st.write("Status do pedido novo:")
                df_status_display = df_status_novo_hist.copy()
                if "STATUS_BANCO" in df_status_display.columns:
                    df_status_display["STATUS"] = df_status_display["STATUS_BANCO"].apply(status_icone)
                st.dataframe(df_status_display, use_container_width=True, hide_index=True)

            if st.button("💾 Salvar substituição", type="primary", key="hist_sub_salvar"):
                if not ok_neon:
                    st.error("Neon não conectado. Verifique os Secrets.")
                elif not pedido_antigo_hist:
                    st.error("Informe o pedido antigo.")
                elif not pedido_novo_hist:
                    st.error("Informe o pedido novo.")
                elif df_antigos_filtrados.empty:
                    st.error("Pedido antigo não encontrado no histórico. Confira o código do fornecedor e o código do pedido antigo.")
                else:
                    row_antigo_hist = df_antigos_filtrados.iloc[0]

                    df_status_novo_hist = consultar_status_pedidos_api([pedido_novo_hist])
                    row_novo_hist = df_status_novo_hist.iloc[0] if not df_status_novo_hist.empty else pd.Series(dtype="object")

                    demanda_antiga = limpar_texto(row_antigo_hist.get(COLUNA_DEMANDA, ""))
                    observacao_antiga = limpar_texto(row_antigo_hist.get(COLUNA_OBSERVACAO, ""))
                    nome_carga_sub = observacao_antiga or f"Substituição {pedido_antigo_hist} -> {pedido_novo_hist}"

                    dados_salvar = {
                        "nome_carga": nome_carga_sub,
                        "nome_arquivo": observacao_antiga,
                        "usuario_registro": usuario_importacao,
                        "analista": "",
                        "departamento": limpar_texto(row_antigo_hist.get(COLUNA_DEPARTAMENTO, "")),
                        "pedido_antigo": pedido_antigo_hist,
                        "fornecedor_antigo": limpar_texto(row_antigo_hist.get(COLUNA_FORNECEDOR, "")),
                        "loja_antiga": limpar_texto(row_antigo_hist.get(COLUNA_LOJA, "")),
                        "status_antigo": limpar_texto(row_antigo_hist.get("STATUS_BANCO", "")),
                        "cod_status_antigo": limpar_texto(row_antigo_hist.get("COD_STATUS_BANCO", "")),
                        "pedido_novo": pedido_novo_hist,
                        "fornecedor_novo": limpar_texto(row_novo_hist.get("FORNECEDOR", "")),
                        "loja_nova": limpar_texto(row_novo_hist.get("FILIAL", "")),
                        "status_novo": limpar_texto(row_novo_hist.get("STATUS_BANCO", "NÃO ENCONTRADO")),
                        "cod_status_novo": limpar_texto(row_novo_hist.get("COD_STATUS_BANCO", "")),
                        "motivo": motivo_sub_hist,
                        "observacao": observacao_sub_hist,
                        "dados_pedido_antigo": row_json(row_antigo_hist) if not row_antigo_hist.empty else "{}",
                        "dados_pedido_novo": row_json(row_novo_hist) if not row_novo_hist.empty else "{}",
                    }

                    if salvar_substituicao_neon(dados_salvar):
                        st.success("Substituição salva no histórico.")
        st.divider()
        st.markdown("### Substituições registradas")
        df_subs = carregar_historico_substituicoes_neon()
        if df_subs.empty:
            st.info("Nenhuma substituição registrada ainda.")
        else:
            df_subs_display = montar_historico_substituicoes_display(df_subs)
            st.dataframe(df_subs_display, use_container_width=True, hide_index=True)
            excel_subs = gerar_excel_abas({"Substituicoes": df_subs_display})
            st.download_button(
                "📥 Baixar histórico de substituições",
                data=excel_subs,
                file_name="historico_substituicoes_pedidos.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )


with aba_config:
    st.subheader("O que foi ajustado nesta versão")
    st.markdown(
        """
        - Antes de carregar o arquivo, o app pergunta se a importação é **Puxada**, **Empurrada** ou **Venda Jurídica**.
        - Antes de carregar, o app também pede o **valor total informado** e depois compara com a soma dos pedidos.
        - Após carregar o arquivo, o app pergunta se existe pedido para remover da importação.
        - A remoção tem filtro por fornecedor e seleção múltipla de pedidos.
        - Os pedidos ficam separados nas abas **Demanda Puxada**, **Demanda Empurrada** e **Demanda Venda Jurídica**.
        - Existe uma aba de **Substituição**, com pedido antigo, pedido novo, motivo fixo e OBS escrita.
        - O histórico de substituições mantém pedido antigo e novo, mesmo que seja refeito várias vezes.
        - O status do pedido novo vem da API: `L = LIBERADO`, `B = BLOQUEADO`.
        - O Streamlit continua salvando histórico no **Neon/PostgreSQL**.
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
        "Para testar localmente, deixe a API FastAPI rodando em outro terminal. Para a equipe usar na rede, rode o Streamlit com --server.address 0.0.0.0."
    )


if admin_logado and aba_usuarios is not None:
    with aba_usuarios:
        st.subheader("Usuários do sistema")
        st.caption("Somente ADMIN cria usuários, altera senha e ativa/desativa acesso.")

        with st.expander("Criar novo usuário", expanded=True):
            col_u1, col_u2, col_u3 = st.columns([1, 1.4, 1])
            with col_u1:
                novo_usuario = st.text_input("Usuário", key="novo_usuario")
            with col_u2:
                novo_nome = st.text_input("Nome", key="novo_nome")
            with col_u3:
                novo_perfil = st.selectbox("Perfil", options=["ANALISTA", "ADMIN"], key="novo_perfil")

            col_u4, col_u5 = st.columns([1, 1])
            with col_u4:
                novo_analista = st.selectbox("Analista vinculado", options=[""] + list(ANALISTAS.keys()), key="novo_analista")
            with col_u5:
                nova_senha = st.text_input("Senha inicial", type="password", key="nova_senha")

            if st.button("Criar usuário", type="primary"):
                ok_criar, msg_criar = criar_usuario_sistema(novo_usuario, nova_senha, novo_nome, novo_perfil, novo_analista)
                if ok_criar:
                    st.success(msg_criar)
                    st.rerun()
                else:
                    st.error(msg_criar)

        df_usuarios = listar_usuarios_sistema()
        if df_usuarios.empty:
            st.info("Nenhum usuário cadastrado.")
        else:
            df_usuarios_display = df_usuarios.copy()
            if "criado_em" in df_usuarios_display.columns:
                df_usuarios_display["criado_em"] = df_usuarios_display["criado_em"].apply(formatar_data_hora_ddmmyyyy)
            if "atualizado_em" in df_usuarios_display.columns:
                df_usuarios_display["atualizado_em"] = df_usuarios_display["atualizado_em"].apply(formatar_data_hora_ddmmyyyy)
            st.dataframe(df_usuarios_display, use_container_width=True, hide_index=True)

            st.markdown("### Manutenção de usuário")
            opcoes_usuarios = {
                f"{row['id']} | {row['usuario']} | {row['nome']} | {row['perfil']} | {'ATIVO' if row['ativo'] else 'INATIVO'}": int(row["id"])
                for _, row in df_usuarios.iterrows()
            }
            usuario_label = st.selectbox("Escolha o usuário", options=list(opcoes_usuarios.keys()))
            usuario_id_sel = opcoes_usuarios.get(usuario_label)

            col_m1, col_m2, col_m3 = st.columns([1, 1, 1])
            with col_m1:
                senha_redefinir = st.text_input("Nova senha", type="password", key="senha_redefinir")
                if st.button("Alterar senha"):
                    ok_senha, msg_senha = alterar_senha_usuario(usuario_id_sel, senha_redefinir)
                    if ok_senha:
                        st.success(msg_senha)
                    else:
                        st.error(msg_senha)
            with col_m2:
                if st.button("Ativar usuário"):
                    ok_st, msg_st = definir_status_usuario(usuario_id_sel, True)
                    if ok_st:
                        st.success(msg_st)
                        st.rerun()
                    else:
                        st.error(msg_st)
            with col_m3:
                if st.button("Desativar usuário"):
                    if int(usuario_id_sel) == int(usuario_logado.get("id", -1)):
                        st.error("Você não pode desativar o próprio usuário logado.")
                    else:
                        ok_st, msg_st = definir_status_usuario(usuario_id_sel, False)
                        if ok_st:
                            st.success(msg_st)
                            st.rerun()
                        else:
                            st.error(msg_st)
