
from pathlib import Path
import pandas as pd
import unicodedata

# ==========================================================
# 1. CONFIGURAÇÕES
# ==========================================================

PASTA = Path(__file__).resolve().parent / "dados"

SAIDA = PASTA / "relatorio_geral_escola.xlsx"


# ==========================================================
# 2. FUNÇÕES AUXILIARES
# ==========================================================

def normalizar(texto):
    """Remove acentos e padroniza textos para comparação."""

    texto = str(texto).strip().lower()
    texto = unicodedata.normalize("NFKD", texto)

    return "".join(
        caractere
        for caractere in texto
        if not unicodedata.combining(caractere)
    )


def localizar_arquivo(prefixo):
    """
    Localiza arquivos pelo início do nome.
    Evita confundir arquivos originais com relatórios gerados.
    """

    encontrados = [
        arquivo
        for arquivo in PASTA.glob("*.xlsx")
        if normalizar(arquivo.stem).startswith(
            normalizar(prefixo)
        )
        and arquivo.name != SAIDA.name
    ]

    if len(encontrados) > 1:
        print(
            f"ATENÇÃO: mais de um arquivo encontrado para {prefixo}:"
        )
        for arquivo in encontrados:
            print("-", arquivo.name)

    return encontrados[0] if encontrados else None


def imprimir_titulo(texto):
    print("\n" + "=" * 65)
    print(texto)
    print("=" * 65)


def salvar_resumo(nome, dados, relatorios):
    """Guarda uma tabela para exportação ao Excel."""

    if isinstance(dados, pd.DataFrame):
        relatorios[nome[:31]] = dados.copy()


def limpar_nomes(serie):
    """Remove valores vazios de uma coluna de nomes."""

    nomes = serie.dropna().astype(str).str.strip()

    return nomes[
        (nomes != "")
        & (nomes.str.lower() != "nan")
    ]


# ==========================================================
# 3. VERIFICAR A PASTA
# ==========================================================

if not PASTA.exists():
    raise FileNotFoundError(
        f"Pasta de dados não encontrada: {PASTA}"
    )

arquivos = sorted(
    [
        arquivo
        for arquivo in PASTA.glob("*.xlsx")
        if arquivo.name != SAIDA.name
    ]
)

imprimir_titulo("INVENTÁRIO DOS ARQUIVOS")

print("Pasta:", PASTA)
print("Quantidade de planilhas:", len(arquivos))

for arquivo in arquivos:
    print("-", arquivo.name)


# ==========================================================
# 4. PREPARAR RELATÓRIOS
# ==========================================================

relatorios = {}


# ==========================================================
# 5. ANÁLISE DE CLIENTES
# ==========================================================

imprimir_titulo("ANÁLISE DE CLIENTES")

arquivo_clientes = localizar_arquivo("clientes-")

if arquivo_clientes:

    clientes = pd.read_excel(arquivo_clientes)

    print("Arquivo:", arquivo_clientes.name)
    print("Total de registros:", len(clientes))
    print("Nomes únicos:", clientes["Nome"].nunique())
    print("Nomes vazios:", clientes["Nome"].isna().sum())

    for coluna in [
        "Situação do contrato",
        "Situação do cliente",
        "Contrato"
    ]:

        if coluna in clientes.columns:

            resumo = (
                clientes[coluna]
                .value_counts(dropna=False)
                .rename_axis(coluna)
                .reset_index(name="Quantidade")
            )

            print(f"\n{coluna}:")
            print(resumo.to_string(index=False))

            salvar_resumo(
                "Clientes_" + coluna,
                resumo,
                relatorios
            )

    # Guardar a base original para análises posteriores
    salvar_resumo(
        "Base_clientes",
        clientes,
        relatorios
    )

else:
    print("Arquivo de clientes não encontrado.")


# ==========================================================
# 6. ANÁLISE DE MATRÍCULAS
# ==========================================================

imprimir_titulo("ANÁLISE DE MATRÍCULAS")

arquivos_matriculas = [
    arquivo
    for arquivo in arquivos
    if normalizar(arquivo.stem).startswith("matricula")
]

resumo_matriculas = []

for arquivo in arquivos_matriculas:

    try:

        # Ler as primeiras 15 linhas sem definir cabeçalho
        previa = pd.read_excel(
            arquivo,
            header=None,
            nrows=15
        )

        linha_cabecalho = None

        # Procurar a linha que contém a coluna Nome
        for indice, linha in previa.iterrows():

            valores = [
                normalizar(valor)
                for valor in linha.dropna().tolist()
            ]

            if any(
                valor in [
                    "nome",
                    "nome do aluno",
                    "nome completo"
                ]
                for valor in valores
            ):

                linha_cabecalho = indice
                break

        if linha_cabecalho is None:

            print(
                "Cabeçalho não identificado:",
                arquivo.name
            )

            continue

        # Ler novamente usando o cabeçalho identificado
        df = pd.read_excel(
            arquivo,
            header=linha_cabecalho
        )

        coluna_nome = next(
            (
                coluna
                for coluna in df.columns
                if normalizar(coluna) in [
                    "nome",
                    "nome do aluno",
                    "nome completo"
                ]
            ),
            None
        )

        if coluna_nome is None:

            print(
                "Coluna de nomes não encontrada:",
                arquivo.name
            )

            continue

        nomes = limpar_nomes(df[coluna_nome])

        turma = arquivo.stem

        resumo_matriculas.append({

            "Arquivo": arquivo.name,

            "Turma": turma,

            "Alunos matriculados": len(nomes),

            "Nomes únicos na turma": nomes.nunique()

        })

        print(
            f"{turma}: {len(nomes)} matrículas"
        )

    except Exception as erro:

        print(
            "Erro na matrícula:",
            arquivo.name,
            erro
        )


if resumo_matriculas:

    matriculas = pd.DataFrame(resumo_matriculas)

    print(
        "\nTotal de participações nas turmas:",
        matriculas["Alunos matriculados"].sum()
    )

    print(
        "Atenção: participações não significam necessariamente "
        "alunos únicos na escola."
    )

    salvar_resumo(
        "Resumo_matriculas",
        matriculas,
        relatorios
    )


# ==========================================================
# 7. ANÁLISE DE VENDAS
# ==========================================================

imprimir_titulo("ANÁLISE DE VENDAS")

arquivo_vendas = localizar_arquivo("Vendas-")

if arquivo_vendas:

    vendas = pd.read_excel(arquivo_vendas)

    vendas["Valor"] = pd.to_numeric(
        vendas["Valor"],
        errors="coerce"
    )

    print("Arquivo:", arquivo_vendas.name)
    print("Quantidade de vendas:", len(vendas))

    print(
        "Valor total vendido: R$",
        round(vendas["Valor"].sum(), 2)
    )

    print(
        "Valor médio por venda: R$",
        round(vendas["Valor"].mean(), 2)
    )

    if "Descrição" in vendas.columns:

        resumo = (
            vendas.groupby("Descrição", dropna=False)
            .agg(
                Quantidade=("Valor", "count"),
                Receita=("Valor", "sum"),
                Valor_medio=("Valor", "mean")
            )
            .reset_index()
            .sort_values(
                "Receita",
                ascending=False
            )
        )

        print("\nVendas por descrição:")
        print(resumo.to_string(index=False))

        salvar_resumo(
            "Vendas_por_descricao",
            resumo,
            relatorios
        )

    # Investigar datas
    if "Data da venda" in vendas.columns:

        vendas["Data da venda"] = pd.to_datetime(
            vendas["Data da venda"],
            errors="coerce"
        )

        print(
            "\nPrimeira venda:",
            vendas["Data da venda"].min()
        )

        print(
            "Última venda:",
            vendas["Data da venda"].max()
        )

    salvar_resumo(
        "Base_vendas",
        vendas,
        relatorios
    )

else:
    print("Arquivo de vendas não encontrado.")


# ==========================================================
# 8. ANÁLISE DE RECEITAS RECEBIDAS
# ==========================================================

imprimir_titulo("ANÁLISE DE RECEITAS")

arquivo_receita = localizar_arquivo("Receita-")

if arquivo_receita:

    receita = pd.read_excel(arquivo_receita)

    receita["Valor"] = pd.to_numeric(
        receita["Valor"],
        errors="coerce"
    )

    print("Arquivo:", arquivo_receita.name)
    print("Recebimentos:", len(receita))

    print(
        "Total recebido registrado: R$",
        round(receita["Valor"].sum(), 2)
    )

    if "Método" in receita.columns:

        resumo = (
            receita.groupby("Método", dropna=False)
            .agg(
                Quantidade=("Valor", "count"),
                Total=("Valor", "sum")
            )
            .reset_index()
        )

        print("\nRecebimentos por método:")
        print(resumo.to_string(index=False))

        salvar_resumo(
            "Receitas_por_metodo",
            resumo,
            relatorios
        )

    salvar_resumo(
        "Base_receitas",
        receita,
        relatorios
    )

else:
    print("Arquivo de receitas não encontrado.")


# ==========================================================
# 9. ANÁLISE DE DESCONTOS
# ==========================================================

imprimir_titulo("ANÁLISE DE DESCONTOS")

arquivo_descontos = localizar_arquivo(
    "Descontos concedidos-"
)

if arquivo_descontos:

    descontos = pd.read_excel(arquivo_descontos)

    for coluna in [
        "Valor original",
        "Desconto",
        "Valor final"
    ]:

        if coluna in descontos.columns:

            descontos[coluna] = pd.to_numeric(
                descontos[coluna],
                errors="coerce"
            )

    print("Quantidade de registros:", len(descontos))

    if "Desconto" in descontos.columns:

        print(
            "Total de descontos: R$",
            round(descontos["Desconto"].sum(), 2)
        )

    salvar_resumo(
        "Base_descontos",
        descontos,
        relatorios
    )

else:
    print("Arquivo de descontos não encontrado.")


# ==========================================================
# 10. ANÁLISE DE CHECK-INS
# ==========================================================

imprimir_titulo("ANÁLISE DE CHECK-INS")

arquivo_checkins = localizar_arquivo(
    "Relatório de top check-ins-"
)

if arquivo_checkins:

    checkins = pd.read_excel(arquivo_checkins)

    print("Registros:", len(checkins))

    if "Qtd. check-ins" in checkins.columns:

        checkins["Qtd. check-ins"] = pd.to_numeric(
            checkins["Qtd. check-ins"],
            errors="coerce"
        )

        print(
            "Total de check-ins registrados:",
            checkins["Qtd. check-ins"].sum()
        )

        print("\n10 maiores frequências:")

        print(
            checkins.nlargest(
                10,
                "Qtd. check-ins"
            )[["Qtd. check-ins"]].to_string(index=False)
        )

    salvar_resumo(
        "Base_checkins",
        checkins,
        relatorios
    )

else:
    print("Arquivo de check-ins não encontrado.")


# ==========================================================
# 11. ANÁLISE DAS MOVIMENTAÇÕES DO CAIXA
# ==========================================================

imprimir_titulo("MOVIMENTAÇÕES DO CAIXA")

arquivo_caixa = localizar_arquivo(
    "Movimentações do caixa-"
)

if arquivo_caixa:

    caixa = pd.read_excel(arquivo_caixa)

    if "Valor" in caixa.columns:

        caixa["Valor"] = pd.to_numeric(
            caixa["Valor"],
            errors="coerce"
        )

        print("Movimentações:", len(caixa))

        print(
            "Soma dos valores registrados: R$",
            round(caixa["Valor"].sum(), 2)
        )

    if (
        "Tipo" in caixa.columns
        and "Valor" in caixa.columns
    ):

        resumo = (
            caixa.groupby("Tipo", dropna=False)
            .agg(
                Quantidade=("Valor", "count"),
                Total=("Valor", "sum")
            )
            .reset_index()
        )

        print("\nMovimentações por tipo:")
        print(resumo.to_string(index=False))

        salvar_resumo(
            "Caixa_por_tipo",
            resumo,
            relatorios
        )

    salvar_resumo(
        "Base_movimentacoes",
        caixa,
        relatorios
    )

else:
    print("Arquivo de movimentações não encontrado.")


# ==========================================================
# 12. ANÁLISE DO FLUXO DE CAIXA
# ==========================================================

imprimir_titulo("ANÁLISE DO FLUXO DE CAIXA")

arquivo_fluxo = localizar_arquivo(
    "fluxo-caixa-"
)

if arquivo_fluxo:

    fluxo = pd.read_excel(arquivo_fluxo)

    print("Arquivo:", arquivo_fluxo.name)
    print("Linhas:", len(fluxo))
    print("Colunas:", len(fluxo.columns))

    print("\nCategorias:")
    print(fluxo.iloc[:, 0].to_string(index=False))

    salvar_resumo(
        "Base_fluxo_caixa",
        fluxo,
        relatorios
    )

else:
    print("Arquivo de fluxo de caixa não encontrado.")


# ==========================================================
# 13. GERAR RELATÓRIO GERAL
# ==========================================================

imprimir_titulo("GERAÇÃO DO RELATÓRIO GERAL")

if relatorios:

    with pd.ExcelWriter(
        SAIDA,
        engine="openpyxl"
    ) as escritor:

        for nome, df in relatorios.items():

            df.to_excel(
                escritor,
                sheet_name=nome[:31],
                index=False
            )

    print("Relatório criado com sucesso!")
    print("Arquivo:", SAIDA)
    print("Quantidade de abas:", len(relatorios))

else:

    print("Nenhum relatório foi gerado.")

imprimir_titulo("ANÁLISE CONCLUÍDA")
