
# -*- coding: utf-8 -*-

"""
ANALISADOR 2 - ESCOLA DE DANÇA

Análises:
1. Perfil dos alunos por modalidade
2. Popularidade dos cursos
3. Receita por modalidade
4. Evasão e permanência
5. Frequência dos alunos
6. Ocupação e capacidade das turmas

Entrada: arquivos Excel e CSV da pasta dados.
Saída: dados/relatorio_analise2.xlsx
"""

from pathlib import Path
import pandas as pd
import unicodedata
import re


# =====================================================
# 1. CONFIGURAÇÕES
# =====================================================

PASTA = Path(__file__).resolve().parent
PASTA_DADOS = PASTA / "dados"

ARQUIVO_SAIDA = PASTA_DADOS / "relatorio_analise2.xlsx"

EXTENSOES = [".xlsx", ".xls", ".xlsm", ".csv"]

ARQUIVOS_IGNORADOS = [
    "relatorio_analise2.xlsx",
    "relatorio_geral_escola.xlsx",
    "analise_modalidades.xlsx",
    "inventario_dados.xlsx",
    "investigacao_matriculas.xlsx",
    "investigacao_vendas.xlsx",
]


# =====================================================
# 2. FUNÇÕES AUXILIARES
# =====================================================

def normalizar(texto):
    """Remove acentos e padroniza nomes de colunas."""

    texto = str(texto).strip().lower()

    texto = unicodedata.normalize("NFKD", texto)

    texto = "".join(
        caractere
        for caractere in texto
        if not unicodedata.combining(caractere)
    )

    texto = re.sub(r"[^a-z0-9]+", "_", texto)

    return texto.strip("_")


def encontrar_coluna(df, alternativas):
    """Procura uma coluna por nomes conhecidos."""

    colunas = {
        normalizar(coluna): coluna
        for coluna in df.columns
    }

    for alternativa in alternativas:

        nome = normalizar(alternativa)

        if nome in colunas:
            return colunas[nome]

    for nome_normalizado, nome_original in colunas.items():

        for alternativa in alternativas:

            if normalizar(alternativa) in nome_normalizado:
                return nome_original

    return None


def converter_numero(serie):
    """Converte valores numéricos, inclusive moeda brasileira."""

    if pd.api.types.is_numeric_dtype(serie):
        return pd.to_numeric(serie, errors="coerce")

    serie = serie.astype(str)

    serie = serie.str.replace(
        r"R\$|\s",
        "",
        regex=True
    )

    possui_virgula = serie.str.contains(",", na=False)

    serie.loc[possui_virgula] = (
        serie.loc[possui_virgula]
        .str.replace(".", "", regex=False)
        .str.replace(",", ".", regex=False)
    )

    return pd.to_numeric(serie, errors="coerce")


def ler_planilha(arquivo):
    """Lê todas as abas de um arquivo Excel ou um CSV."""

    if arquivo.suffix.lower() == ".csv":

        try:
            df = pd.read_csv(
                arquivo,
                sep=None,
                engine="python",
                encoding="utf-8-sig"
            )

        except UnicodeDecodeError:

            df = pd.read_csv(
                arquivo,
                sep=None,
                engine="python",
                encoding="latin1"
            )

        return {"CSV": df}

    excel = pd.ExcelFile(arquivo)

    resultado = {}

    for aba in excel.sheet_names:

        if "matricula" in normalizar(arquivo.stem):

            bruto = pd.read_excel(
                arquivo,
                sheet_name=aba,
                header=None
            )

            linha_cabecalho = None

            for indice, linha in bruto.iterrows():

                valores = {
                    normalizar(valor)
                    for valor in linha
                    if pd.notna(valor)
                }

                if "nome" in valores and "idade" in valores:
                    linha_cabecalho = indice
                    break

            if linha_cabecalho is None:
                resultado[aba] = pd.DataFrame()
                continue

            cabecalhos = [
                str(valor).strip()
                if pd.notna(valor)
                else f"coluna_{indice + 1}"
                for indice, valor in enumerate(
                    bruto.iloc[linha_cabecalho].tolist()
                )
            ]

            participantes = bruto.iloc[
                linha_cabecalho + 1:
            ].copy()

            participantes.columns = cabecalhos
            participantes = participantes.dropna(how="all")

            modalidade = None

            for _, linha in bruto.iterrows():

                if (
                    len(linha) > 1
                    and normalizar(linha.iloc[0]) == "modalidade"
                    and pd.notna(linha.iloc[1])
                ):
                    modalidade = str(linha.iloc[1]).strip()
                    break

            participantes.attrs["modalidade"] = (
                modalidade or identificar_modalidade_arquivo(arquivo.name)
            )

            resultado[aba] = participantes

        else:

            resultado[aba] = pd.read_excel(
                arquivo,
                sheet_name=aba
            )

    return resultado


def identificar_modalidade_arquivo(nome):
    """
    Tenta identificar a modalidade pelo nome
    dos arquivos de matrícula.
    """

    nome = Path(nome).stem

    nome = re.sub(
        r"^matriculas?[-_ ]*",
        "",
        nome,
        flags=re.IGNORECASE
    )

    nome = nome.replace("_", " ")
    nome = nome.replace("-", " ")

    nome = re.sub(
        r"^\s*\d{1,2}\s*\d{2}\s+",
        "",
        nome
    )

    return nome.strip()


# =====================================================
# 3. LEITURA DOS DADOS
# =====================================================

def analisar():

    PASTA_DADOS.mkdir(
        parents=True,
        exist_ok=True
    )

    arquivos = [
        arquivo
        for arquivo in PASTA_DADOS.iterdir()
        if arquivo.is_file()
        and arquivo.suffix.lower() in EXTENSOES
        and arquivo.name not in ARQUIVOS_IGNORADOS
        and not arquivo.name.startswith("~$")
    ]

    if not arquivos:

        print("Nenhuma planilha encontrada na pasta dados.")

        return

    print("=" * 55)
    print("INICIANDO ANALISADOR 2")
    print("=" * 55)

    diagnostico = []

    dados_perfil = []
    dados_popularidade = []
    dados_receita = []
    dados_evasao = []
    dados_frequencia = []
    dados_ocupacao = []

    # Evita contar o mesmo aluno duas vezes
    # dentro da mesma modalidade.
    matriculas = {}

    # =================================================
    # 4. EXAMINAR ARQUIVOS
    # =================================================

    for arquivo in arquivos:

        print(f"Lendo: {arquivo.name}")

        try:

            abas = ler_planilha(arquivo)

        except Exception as erro:

            diagnostico.append({
                "arquivo": arquivo.name,
                "aba": "",
                "linhas": 0,
                "observacao": str(erro)
            })

            continue

        for nome_aba, df in abas.items():

            if df.empty:
                continue

            nome_col = encontrar_coluna(
                df,
                ["nome", "nome do aluno", "aluno", "cliente"]
            )

            curso_col = encontrar_coluna(
                df,
                ["modalidade", "curso", "turma", "atividade"]
            )

            idade_col = encontrar_coluna(
                df,
                ["idade"]
            )

            genero_col = encontrar_coluna(
                df,
                ["sexo", "genero"]
            )

            status_col = encontrar_coluna(
                df,
                ["status", "situacao", "situacao do contrato", "estado"]
            )

            situacao_contrato_col = encontrar_coluna(
                df,
                ["situacao do contrato", "status do contrato"]
            )

            valor_col = encontrar_coluna(
                df,
                ["valor", "valor pago", "total", "receita", "valor recebido"]
            )

            data_col = encontrar_coluna(
                df,
                ["data", "data da venda", "data pagamento", "data matricula"]
            )

            presenca_col = encontrar_coluna(
                df,
                [
                    "checkin",
                    "check ins",
                    "qtd check ins",
                    "quantidade de check ins",
                    "presenca",
                    "presencas",
                    "frequencia"
                ]
            )

            ultimo_checkin_col = encontrar_coluna(
                df,
                [
                    "ultimo checkin",
                    "ultimo check in",
                    "data do ultimo checkin"
                ]
            )

            capacidade_col = encontrar_coluna(
                df,
                ["capacidade", "vagas", "limite de alunos"]
            )

            modalidade_arquivo = (
                df.attrs.get("modalidade")
                or identificar_modalidade_arquivo(arquivo.name)
            )

            diagnostico.append({
                "arquivo": arquivo.name,
                "aba": nome_aba,
                "linhas": len(df),
                "linhas_aproveitadas": 0,
                "colunas": ", ".join(map(str, df.columns)),
                "nome_identificado": nome_col or "",
                "modalidade_identificada": curso_col or (
                    modalidade_arquivo
                    if "matricula" in normalizar(arquivo.name)
                    else ""
                ),
                "valor_identificado": valor_col or "",
                "status_identificado": status_col or "",
                "situacao_contrato_identificada": situacao_contrato_col or "",
                "frequencia_identificada": presenca_col or "",
                "observacao": "Leitura concluída"
            })

            # =========================================
            # 5. PERFIL E POPULARIDADE
            # =========================================

            eh_matricula = (
                "matricula" in normalizar(arquivo.name)
                or "matricula" in normalizar(nome_aba)
            )

            if eh_matricula and nome_col:

                if curso_col:

                    modalidades = df[curso_col].astype("string").str.strip()

                else:

                    modalidades = pd.Series(
                        [modalidade_arquivo] * len(df),
                        index=df.index
                    )

                alunos = df[nome_col].astype("string").str.strip()

                base = pd.DataFrame({
                    "aluno": alunos,
                    "modalidade": modalidades
                })

                base = base.dropna(subset=["aluno", "modalidade"])
                base = base[
                    base["aluno"].ne("")
                    & ~base["aluno"].str.casefold().isin(["nan", "none"])
                    & base["modalidade"].ne("")
                ]

                diagnostico[-1]["linhas_aproveitadas"] = len(base)

                for modalidade, grupo in base.groupby("modalidade"):

                    alunos_unicos = grupo["aluno"].nunique()

                    dados_perfil.append({
                        "modalidade": modalidade,
                        "indicador": "Alunos únicos na lista",
                        "valor": alunos_unicos
                    })

                    matriculas.setdefault(
                        modalidade,
                        set()
                    )

                    matriculas[modalidade].update(
                        grupo["aluno"].unique()
                    )

                    # Perfil por idade
                    if idade_col:

                        idades = converter_numero(
                            df.loc[grupo.index, idade_col]
                        )

                        dados_perfil.append({
                            "modalidade": modalidade,
                            "indicador": "Idade média",
                            "valor": round(idades.mean(), 1)
                            if idades.notna().any()
                            else None
                        })

                    # Perfil por gênero
                    if genero_col:

                        generos = df.loc[
                            grupo.index,
                            genero_col
                        ].fillna("Não informado")

                        for genero, quantidade in generos.value_counts().items():

                            dados_perfil.append({
                                "modalidade": modalidade,
                                "indicador": f"Gênero: {genero}",
                                "valor": int(quantidade)
                            })

                    dados_ocupacao.append({
                        "modalidade": modalidade,
                        "alunos": alunos_unicos,
                        "capacidade": None,
                        "ocupacao_percentual": None,
                        "observacao": "Capacidade da turma não informada nas planilhas."
                    })

            # =========================================
            # 6. RECEITA POR MODALIDADE
            # =========================================

            if valor_col and curso_col:

                valores = converter_numero(
                    df[valor_col]
                )

                base_receita = pd.DataFrame({
                    "modalidade": df[curso_col].astype(str),
                    "valor": valores
                })

                base_receita = base_receita.dropna(
                    subset=["valor"]
                )

                for modalidade, grupo in base_receita.groupby("modalidade"):

                    dados_receita.append({
                        "modalidade": modalidade,
                        "receita": grupo["valor"].sum(),
                        "quantidade_registros": len(grupo),
                        "arquivo_origem": arquivo.name
                    })

            # =========================================
            # 7. EVASÃO E PERMANÊNCIA
            # =========================================

            if situacao_contrato_col:

                status = df[situacao_contrato_col].astype("string").str.strip()
                status_normalizado = status.map(normalizar)

                situacoes = {
                    "Ativos": {"ativo", "ativa", "active"},
                    "Cancelados": {"cancelado", "cancelada", "evasao", "desistente"},
                    "Encerrados": {"encerrado", "encerrada", "finalizado", "concluido"}
                }

                diagnostico[-1]["linhas_aproveitadas"] = max(
                    diagnostico[-1]["linhas_aproveitadas"],
                    int(status.notna().sum())
                )

                for situacao, valores in situacoes.items():

                    quantidade = status_normalizado.isin(valores).sum()

                    if quantidade > 0:

                        dados_evasao.append({
                            "arquivo": arquivo.name,
                            "situacao": situacao,
                            "quantidade": int(quantidade)
                        })

            # =========================================
            # 8. FREQUÊNCIA
            # =========================================

            eh_checkin = bool(
                re.search(
                    r"check[_]?ins?|frequencia",
                    normalizar(arquivo.name)
                )
            )

            if nome_col and eh_checkin:

                if presenca_col:

                    valores = converter_numero(
                        df[presenca_col]
                    )

                    frequencia = pd.DataFrame({
                        "aluno": df[nome_col].astype("string").str.strip(),
                        "presencas": valores
                    })

                    if ultimo_checkin_col:
                        frequencia["ultimo_checkin"] = df[ultimo_checkin_col]

                    if status_col:
                        frequencia["situacao"] = df[status_col]

                    frequencia = frequencia.dropna(
                        subset=["aluno", "presencas"]
                    )

                    diagnostico[-1]["linhas_aproveitadas"] = max(
                        diagnostico[-1]["linhas_aproveitadas"],
                        len(frequencia)
                    )

                    agregacoes = {"presencas": "sum"}

                    if "ultimo_checkin" in frequencia.columns:
                        agregacoes["ultimo_checkin"] = "max"

                    if "situacao" in frequencia.columns:
                        agregacoes["situacao"] = "first"

                    frequencia = frequencia.groupby(
                        "aluno",
                        as_index=False
                    ).agg(agregacoes)

                else:

                    frequencia = (
                        df.dropna(subset=[nome_col])
                        .groupby(nome_col)
                        .size()
                        .reset_index(name="presencas")
                    )

                    frequencia = frequencia.rename(
                        columns={nome_col: "aluno"}
                    )

                frequencia["arquivo"] = arquivo.name

                dados_frequencia.extend(
                    frequencia.to_dict("records")
                )

            # =========================================
            # 9. OCUPAÇÃO DAS TURMAS
            # =========================================

            if curso_col and capacidade_col and nome_col:

                base_ocupacao = df.copy()

                base_ocupacao["_capacidade"] = converter_numero(
                    base_ocupacao[capacidade_col]
                )

                for modalidade, grupo in base_ocupacao.groupby(curso_col):

                    quantidade_alunos = grupo[nome_col].nunique()

                    capacidades = grupo["_capacidade"].dropna()

                    capacidade = (
                        capacidades.max()
                        if not capacidades.empty
                        else None
                    )

                    percentual = None

                    if capacidade and capacidade > 0:

                        percentual = round(
                            quantidade_alunos / capacidade * 100,
                            1
                        )

                    dados_ocupacao.append({
                        "modalidade": modalidade,
                        "alunos": quantidade_alunos,
                        "capacidade": capacidade,
                        "ocupacao_percentual": percentual
                    })

    # =================================================
    # 10. ORGANIZAÇÃO DOS RESULTADOS
    # =================================================

    popularidade = pd.DataFrame([
        {
            "modalidade": modalidade,
            "alunos_identificados": len(alunos)
        }
        for modalidade, alunos in matriculas.items()
    ])

    if not popularidade.empty:

        popularidade = popularidade.sort_values(
            "alunos_identificados",
            ascending=False
        )

    perfil = pd.DataFrame(dados_perfil)

    receita = pd.DataFrame(dados_receita)

    if not receita.empty:

        receita = receita.groupby(
            "modalidade",
            as_index=False
        ).agg(
            receita=("receita", "sum"),
            quantidade_registros=("quantidade_registros", "sum")
        )

        receita["receita"] = receita["receita"].round(2)

    evasao = pd.DataFrame(dados_evasao)

    frequencia = pd.DataFrame(dados_frequencia)

    ocupacao = pd.DataFrame(dados_ocupacao)

    diagnostico = pd.DataFrame(diagnostico)

    # =================================================
    # 11. EXPORTAÇÃO PARA EXCEL
    # =================================================

    with pd.ExcelWriter(
        ARQUIVO_SAIDA,
        engine="openpyxl"
    ) as escritor:

        diagnostico.to_excel(
            escritor,
            sheet_name="Diagnostico",
            index=False
        )

        perfil.to_excel(
            escritor,
            sheet_name="Perfil_modalidade",
            index=False
        )

        popularidade.to_excel(
            escritor,
            sheet_name="Popularidade",
            index=False
        )

        receita.to_excel(
            escritor,
            sheet_name="Receita_modalidade",
            index=False
        )

        evasao.to_excel(
            escritor,
            sheet_name="Evasao",
            index=False
        )

        frequencia.to_excel(
            escritor,
            sheet_name="Frequencia",
            index=False
        )

        ocupacao.to_excel(
            escritor,
            sheet_name="Ocupacao",
            index=False
        )

        limitacoes = pd.DataFrame([
            {
                "analise": "Receita por modalidade",
                "situacao": "Não calculada",
                "motivo": (
                    "Os arquivos de pagamentos/receitas não contêm uma associação "
                    "confiável entre cada valor pago e uma modalidade. Vendas não "
                    "foram tratadas como recebimentos."
                )
            },
            {
                "analise": "Evasão e permanência",
                "situacao": "Contagens de situação disponíveis; taxas não calculadas",
                "motivo": (
                    "Há situação atual do contrato, mas não há datas de entrada e saída "
                    "que permitam calcular evasão percentual ou tempo de permanência."
                )
            },
            {
                "analise": "Frequência",
                "situacao": "Contagem de check-ins disponível",
                "motivo": (
                    "O relatório contém quantidade de check-ins, mas não define período "
                    "nem número de aulas possíveis; não foi calculada frequência percentual."
                )
            },
            {
                "analise": "Ocupação das turmas",
                "situacao": "Alunos por turma disponíveis; percentual não calculado",
                "motivo": "As planilhas de matrícula não informam capacidade máxima das turmas."
            },
            {
                "analise": "Alunos únicos no total",
                "situacao": "Não calculado",
                "motivo": (
                    "As listas permitem identificar alunos únicos por modalidade, mas "
                    "não oferecem uma chave estável para deduplicação global entre turmas."
                )
            }
        ])

        limitacoes.to_excel(
            escritor,
            sheet_name="Limitacoes",
            index=False
        )

    # =================================================
    # 12. RESUMO NO TERMINAL
    # =================================================

    print("\n" + "=" * 55)
    print("ANÁLISE CONCLUÍDA")
    print("=" * 55)

    print(f"Arquivos examinados: {len(arquivos)}")

    print(f"Modalidades identificadas: {len(popularidade)}")

    print(f"Alunos identificados: {sum(len(x) for x in matriculas.values())}")

    if not receita.empty:

        print(
            "Receita identificada: R$ "
            f"{receita['receita'].sum():,.2f}"
        )

    print(f"\nRelatório salvo em:\n{ARQUIVO_SAIDA}")

    print("\nConfira a aba Diagnostico antes de interpretar os resultados.")


# =====================================================
# 13. EXECUÇÃO
# =====================================================

if __name__ == "__main__":
    analisar()
