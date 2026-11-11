
# -*- coding: utf-8 -*-

"""
PAINEL 2 - GESTÃO INTELIGENTE

Painel estratégico da Escola de Dança.

Utiliza os dados produzidos por analisador2.py.
"""

from pathlib import Path
import subprocess
import sys

import pandas as pd
import streamlit as st


# =====================================================
# 1. CONFIGURAÇÕES
# =====================================================

PASTA = Path(__file__).resolve().parent

PASTA_DADOS = PASTA / "dados"

ARQUIVO_RELATORIO = (
    PASTA_DADOS / "relatorio_analise2.xlsx"
)

ARQUIVO_ANALISADOR = (
    PASTA / "analisador2.py"
)


# =====================================================
# 2. CONFIGURAÇÃO VISUAL
# =====================================================

st.set_page_config(
    page_title="Gestão Inteligente - Painel 2",
    page_icon="💃",
    layout="wide"
)

st.title("💃 Gestão Inteligente")

st.subheader(
    "Painel estratégico da Escola de Dança"
)

st.write(
    "Acompanhe o perfil dos alunos, "
    "popularidade dos cursos, receitas, "
    "evasão, frequência e ocupação das turmas."
)

st.divider()


# =====================================================
# 3. ATUALIZAÇÃO DOS DADOS
# =====================================================

col_atualizar, col_status = st.columns([1, 3])

with col_atualizar:

    atualizar = st.button(
        "Atualizar análises",
        type="primary",
        width="stretch"
    )

with col_status:

    st.caption(
        "Executa o analisador2.py e atualiza o relatório Excel."
    )


if atualizar:

    if not ARQUIVO_ANALISADOR.exists():

        st.error(
            "O arquivo analisador2.py não foi encontrado."
        )

    else:

        with st.spinner(
            "Analisando os dados da escola..."
        ):

            resultado = subprocess.run(
                [
                    sys.executable,
                    str(ARQUIVO_ANALISADOR)
                ],
                cwd=PASTA,
                capture_output=True,
                text=True
            )

        if resultado.returncode == 0:

            st.success(
                "Análises atualizadas com sucesso!"
            )

            if resultado.stdout:

                st.code(resultado.stdout)

            st.rerun()

        else:

            st.error(
                "O analisador encontrou um erro."
            )

            st.code(
                resultado.stderr or resultado.stdout
            )


# =====================================================
# 4. VERIFICAR RELATÓRIO
# =====================================================

if not ARQUIVO_RELATORIO.exists():

    st.warning(
        "O relatório ainda não foi encontrado."
    )

    st.info(
        "Clique em 'Atualizar análises' para gerar o relatório."
    )

    st.stop()


try:

    excel = pd.ExcelFile(
        ARQUIVO_RELATORIO
    )

except Exception as erro:

    st.error(
        f"Não foi possível abrir o relatório: {erro}"
    )

    st.stop()


# =====================================================
# 5. FUNÇÃO PARA LER ABAS
# =====================================================

def ler_aba(nome):

    if nome not in excel.sheet_names:

        return pd.DataFrame()

    return pd.read_excel(
        ARQUIVO_RELATORIO,
        sheet_name=nome
    )


# =====================================================
# 6. CARREGAR DADOS
# =====================================================

perfil = ler_aba(
    "Perfil_modalidade"
)

popularidade = ler_aba(
    "Popularidade"
)

receita = ler_aba(
    "Receita_modalidade"
)

evasao = ler_aba(
    "Evasao"
)

frequencia = ler_aba(
    "Frequencia"
)

ocupacao = ler_aba(
    "Ocupacao"
)

diagnostico = ler_aba(
    "Diagnostico"
)

limitacoes = ler_aba(
    "Limitacoes"
)


# =====================================================
# 7. INDICADORES GERAIS
# =====================================================

total_alunos = 0

if (
    not popularidade.empty
    and "alunos_identificados" in popularidade.columns
):

    total_alunos = int(
        popularidade[
            "alunos_identificados"
        ].sum()
    )


total_modalidades = 0

if not popularidade.empty and "modalidade" in popularidade.columns:

    total_modalidades = (
        popularidade["modalidade"].nunique()
    )


total_receita = None

if not receita.empty and "receita" in receita.columns:

    total_receita = receita["receita"].sum()


total_presencas = 0

if not frequencia.empty and "presencas" in frequencia.columns:

    total_presencas = frequencia[
        "presencas"
    ].sum()


def moeda(valor):

    return (
        "R$ "
        + f"{valor:,.2f}"
        .replace(",", "X")
        .replace(".", ",")
        .replace("X", ".")
    )


st.header("Visão geral da escola")

c1, c2, c3, c4 = st.columns(4)

c1.metric(
    "Vínculos aluno-modalidade",
    f"{total_alunos:,}".replace(",", ".")
)

c2.metric(
    "Modalidades",
    total_modalidades
)

c3.metric(
    "Receita por modalidade",
    moeda(total_receita) if total_receita is not None else "Não disponível"
)

c4.metric(
    "Presenças registradas",
    f"{total_presencas:,.0f}".replace(",", ".")
)

st.caption(
    "A soma representa alunos únicos dentro de cada modalidade. "
    "Uma pessoa matriculada em mais de uma turma aparece em cada vínculo; "
    "não é uma contagem global de pessoas únicas."
)

st.divider()


# =====================================================
# 8. NAVEGAÇÃO
# =====================================================

st.sidebar.title(
    "Navegação"
)

pagina = st.sidebar.radio(
    "Escolha uma análise:",
    [
        "Visão geral",
        "Perfil dos alunos",
        "Popularidade dos cursos",
        "Receita por modalidade",
        "Evasão e permanência",
        "Frequência dos alunos",
        "Ocupação das turmas",
        "Diagnóstico dos dados"
    ]
)


# =====================================================
# 9. VISÃO GERAL
# =====================================================

if pagina == "Visão geral":

    st.header(
        "Resumo estratégico"
    )

    coluna1, coluna2 = st.columns(2)

    with coluna1:

        st.subheader(
            "Popularidade dos cursos"
        )

        if not popularidade.empty:

            dados = popularidade.set_index(
                "modalidade"
            )["alunos_identificados"]

            st.bar_chart(dados)

        else:

            st.info(
                "Não foram identificadas matrículas por modalidade."
            )

    with coluna2:

        st.subheader(
            "Receita associada a modalidades"
        )

        if not receita.empty and {"modalidade", "receita"}.issubset(receita.columns):

            dados = receita.set_index(
                "modalidade"
            )["receita"]

            st.bar_chart(dados)

        else:

            st.info(
                "Não existem dados suficientes para "
                "atribuir receitas aos cursos."
            )

    st.divider()

    st.subheader(
        "Resumo de evasão"
    )

    if not evasao.empty:

        st.dataframe(
            evasao,
            width="stretch",
            hide_index=True
        )

    else:

        st.info(
            "Não foram encontrados dados de situação dos contratos."
        )


# =====================================================
# 10. PERFIL DOS ALUNOS
# =====================================================

elif pagina == "Perfil dos alunos":

    st.header(
        "Perfil dos alunos por modalidade"
    )

    if perfil.empty:

        st.warning(
            "Não foram encontradas informações suficientes "
            "de idade ou gênero."
        )

    else:

        modalidades = sorted(
            perfil["modalidade"].dropna().unique()
        )

        modalidade_escolhida = st.selectbox(
            "Selecione uma modalidade:",
            modalidades
        )

        dados = perfil[
            perfil["modalidade"] == modalidade_escolhida
        ]

        st.subheader(
            f"Perfil: {modalidade_escolhida}"
        )

        quantidade = dados.loc[
            dados["indicador"] == "Alunos únicos na lista",
            "valor"
        ]

        if not quantidade.empty:
            st.metric("Alunos nesta lista", int(quantidade.iloc[0]))

        idade = dados.loc[
            dados["indicador"] == "Idade média",
            "valor"
        ]

        if not idade.empty and pd.notna(idade.iloc[0]):
            st.metric("Idade média informada", f"{idade.iloc[0]:.1f} anos")

        generos = dados[
            dados["indicador"].astype(str).str.startswith("Gênero:")
        ].copy()

        if not generos.empty:
            generos["gênero"] = generos["indicador"].str.replace(
                "Gênero: ", "", regex=False
            )
            st.subheader("Distribuição por gênero")
            st.bar_chart(generos.set_index("gênero")["valor"])

        st.subheader("Indicadores disponíveis")
        st.dataframe(dados, width="stretch", hide_index=True)


# =====================================================
# 11. POPULARIDADE DOS CURSOS
# =====================================================

elif pagina == "Popularidade dos cursos":

    st.header(
        "Popularidade das modalidades"
    )

    if popularidade.empty:

        st.info(
            "Não há dados de matrículas."
        )

    else:

        popularidade = popularidade.sort_values(
            "alunos_identificados",
            ascending=False
        )

        st.subheader(
            "Ranking de cursos"
        )

        st.bar_chart(
            popularidade.set_index(
                "modalidade"
            )["alunos_identificados"]
        )

        st.dataframe(
            popularidade,
            width="stretch",
            hide_index=True
        )

        menos_popular = popularidade.iloc[-1]

        st.warning(
            "Modalidade com menos alunos identificados: "
            f"{menos_popular['modalidade']} "
            f"({menos_popular['alunos_identificados']} alunos)."
        )

        st.caption(
            "A contagem é de alunos únicos em cada lista de turma. "
            "A mesma pessoa pode aparecer em mais de uma modalidade; "
            "as planilhas não permitem deduplicar com segurança o total geral."
        )


# =====================================================
# 12. RECEITA POR MODALIDADE
# =====================================================

elif pagina == "Receita por modalidade":

    st.header(
        "Receita associada a modalidades"
    )

    if receita.empty:

        st.warning(
            "Não foi possível identificar receitas diretamente "
            "associadas às modalidades."
        )

    else:

        receita = receita.sort_values(
            "receita",
            ascending=False
        )

        st.metric(
            "Receita identificada",
            moeda(receita["receita"].sum())
        )

        st.bar_chart(
            receita.set_index(
                "modalidade"
            )["receita"]
        )

        st.dataframe(
            receita,
            width="stretch",
            hide_index=True
        )

        st.caption(
            "Os recebimentos só são exibidos quando associados diretamente à modalidade. "
            "Receita não significa lucro; custos por modalidade não estão disponíveis."
        )


# =====================================================
# 13. EVASÃO E PERMANÊNCIA
# =====================================================

elif pagina == "Evasão e permanência":

    st.header(
        "Análise de evasão e permanência"
    )

    if evasao.empty:

        st.info(
            "Não foram identificadas informações de status."
        )

    else:

        st.dataframe(
            evasao,
            width="stretch",
            hide_index=True
        )

        st.subheader(
            "Quantidade por situação"
        )

        resumo = evasao.groupby(
            "situacao"
        )["quantidade"].sum()

        st.bar_chart(resumo)

        st.caption(
            "As quantidades são situações atuais dos contratos. Não há datas de entrada "
            "e saída para calcular taxa de evasão ou tempo de permanência."
        )


# =====================================================
# 14. FREQUÊNCIA DOS ALUNOS
# =====================================================

elif pagina == "Frequência dos alunos":

    st.header(
        "Frequência dos alunos"
    )

    if frequencia.empty:

        st.info(
            "Não foram encontrados registros de frequência."
        )

    else:

        agregacoes = {"presencas": "sum"}

        if "ultimo_checkin" in frequencia.columns:
            agregacoes["ultimo_checkin"] = "max"

        if "situacao" in frequencia.columns:
            agregacoes["situacao"] = "first"

        resumo = frequencia.groupby(
            "aluno",
            as_index=False
        ).agg(agregacoes)

        resumo = resumo.sort_values(
            "presencas",
            ascending=True
        )

        st.subheader(
            "Menor contagem de check-ins no relatório"
        )

        st.dataframe(
            resumo.head(20),
            width="stretch",
            hide_index=True
        )

        st.subheader(
            "Distribuição de frequência"
        )

        st.bar_chart(
            resumo.tail(30).set_index(
                "aluno"
            )["presencas"]
        )

        st.caption(
            "O período coberto e o número de aulas possíveis não foram informados. "
            "A contagem, portanto, não é uma taxa de frequência nem confirma evasão."
        )


# =====================================================
# 15. OCUPAÇÃO DAS TURMAS
# =====================================================

elif pagina == "Ocupação das turmas":

    st.header(
        "Ocupação e capacidade"
    )

    if ocupacao.empty:

        st.warning(
            "Não foram encontradas informações de capacidade."
        )

    else:

        st.dataframe(ocupacao, width="stretch", hide_index=True)

        if (
            "ocupacao_percentual" in ocupacao.columns
            and ocupacao["ocupacao_percentual"].notna().any()
        ):

            dados = ocupacao.dropna(
                subset=["ocupacao_percentual"]
            )

            st.subheader(
                "Percentual de ocupação"
            )

            st.bar_chart(
                dados.set_index(
                    "modalidade"
                )["ocupacao_percentual"]
            )

        else:

            st.info(
                "Para calcular a ocupação percentual, "
                "é necessário cadastrar a capacidade máxima das turmas."
            )


# =====================================================
# 16. DIAGNÓSTICO
# =====================================================

elif pagina == "Diagnóstico dos dados":

    st.header(
        "Diagnóstico das planilhas"
    )

    st.write(
        "Aqui você pode conferir quais arquivos foram lidos "
        "e quais informações foram identificadas."
    )

    if diagnostico.empty:

        st.info(
            "Diagnóstico não disponível."
        )

    else:

        st.dataframe(
            diagnostico,
            width="stretch",
            hide_index=True
        )

    st.subheader("Análises limitadas pelos dados disponíveis")

    if limitacoes.empty:
        st.info("O relatório não contém uma aba de limitações.")
    else:
        st.dataframe(limitacoes, width="stretch", hide_index=True)


# =====================================================
# 17. DOWNLOAD DO RELATÓRIO
# =====================================================

st.divider()

st.subheader(
    "Exportar relatório"
)

with open(
    ARQUIVO_RELATORIO,
    "rb"
) as arquivo:

    st.download_button(
        label="Baixar relatório Excel",
        data=arquivo.read(),
        file_name="relatorio_analise2.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        )
    )


# =====================================================
# 18. RODAPÉ
# =====================================================

st.divider()

st.caption(
    "Gestão Inteligente | Sistema de análise estratégica "
    "da Escola de Dança."
)
