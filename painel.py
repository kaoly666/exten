
import streamlit as st
import pandas as pd
from pathlib import Path
import subprocess
import sys

# =====================================
# CONFIGURAÇÕES
# =====================================

PASTA = Path(__file__).resolve().parent
PASTA_DADOS = PASTA / "dados"

ARQUIVO_EXCEL = (
    PASTA_DADOS / "relatorio_geral_escola.xlsx"
)

ANALISADOR = PASTA / "analisador.py"

st.set_page_config(
    page_title="Gestão da Escola de Dança",
    page_icon="💃",
    layout="wide"
)

# =====================================
# CABEÇALHO
# =====================================

st.title("💃 Gestão Inteligente")
st.subheader("Painel administrativo da Escola de Dança")

st.write(
    "Acompanhe os indicadores, alunos, "
    "modalidades e resultados financeiros."
)

st.divider()

# =====================================
# ATUALIZAR ANÁLISE
# =====================================

if st.button("Atualizar dados da escola"):

    if not ANALISADOR.exists():
        st.error("Arquivo analisador.py não encontrado.")

    else:
        with st.spinner("Executando análise dos dados..."):

            try:
                subprocess.run(
                    [sys.executable, str(ANALISADOR)],
                    check=True,
                    capture_output=True,
                    text=True
                )

                st.success("Análise atualizada!")

                st.rerun()

            except subprocess.CalledProcessError as erro:

                st.error("O analisador encontrou um erro.")

                st.code(erro.stderr or erro.stdout)

# =====================================
# VERIFICAR EXCEL
# =====================================

if not ARQUIVO_EXCEL.exists():

    st.warning(
        "O relatório Excel ainda não foi encontrado. "
        "Execute primeiro o analisador.py."
    )

    st.stop()

# =====================================
# CARREGAR PLANILHAS
# =====================================

try:

    excel = pd.ExcelFile(ARQUIVO_EXCEL)

except Exception as erro:

    st.error(f"Não foi possível abrir o relatório: {erro}")

    st.stop()

abas = excel.sheet_names

st.sidebar.title("Navegação")

aba_escolhida = st.sidebar.selectbox(
    "Escolha uma análise:",
    abas
)

# =====================================
# EXIBIR DADOS
# =====================================

df = pd.read_excel(
    ARQUIVO_EXCEL,
    sheet_name=aba_escolhida
)

st.header(aba_escolhida.replace("_", " "))

col1, col2 = st.columns(2)

with col1:
    st.metric("Registros", len(df))

with col2:
    st.metric("Colunas", len(df.columns))

st.subheader("Tabela de dados")

st.dataframe(
    df,
    use_container_width=True,
    hide_index=True
)

# =====================================
# DOWNLOAD
# =====================================

with open(ARQUIVO_EXCEL, "rb") as arquivo:

    st.download_button(
        label="Baixar relatório Excel",
        data=arquivo,
        file_name="relatorio_geral_escola.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        )
    )

st.divider()

st.caption(
    "Sistema de gestão e análise de dados da escola."
)
