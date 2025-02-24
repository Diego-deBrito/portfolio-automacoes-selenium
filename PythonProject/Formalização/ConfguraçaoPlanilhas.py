import pandas as pandas_library
from openpyxl import load_workbook
from openpyxl.styles import Alignment, PatternFill
from datetime import datetime, timedelta
from unidecode import unidecode


# Função para gerar endereço de e-mail
def gerar_endereco_de_email(nome_do_tecnico_responsavel):
    if not isinstance(nome_do_tecnico_responsavel, str) or pandas_library.isna(nome_do_tecnico_responsavel):
        return "email_invalido@esporte.gov.br"

    nome_limpo = unidecode(nome_do_tecnico_responsavel.strip())
    partes_do_nome = nome_limpo.split()

    if len(partes_do_nome) >= 2:
        primeiro_nome = partes_do_nome[0].lower()
        segundo_nome = partes_do_nome[1].lower()
        return f"{primeiro_nome}.{segundo_nome}@esporte.gov.br"

    return "email_invalido@esporte.gov.br"


# Função para ler a primeira planilha
def ler_primeira_planilha():
    caminho = r'C:\Users\diego.brito\OneDrive - Ministério do Desenvolvimento e Assistência Social\Power BI\Python\Consulta Parecer\Consulta Transferegov Parecer Novo.xlsx'
    df = pandas_library.read_excel(caminho)
    df.columns = df.columns.str.strip()
    return df


# Função para ler a segunda planilha
def ler_segunda_planilha():
    caminho = r'C:\Users\diego.brito\OneDrive - Ministério do Desenvolvimento e Assistência Social\Power BI\Python\Consulta Transferegov Requisitos.xlsx'
    try:
        df = pandas_library.read_excel(caminho)
    except Exception as e:
        print(f"Erro ao ler a segunda planilha: {e}")
        return None

    df.columns = df.columns.str.strip()
    colunas_de_interesse = ['Proposta', 'Certidões', 'Declarações', 'Comprovantes de Execução', 'Outros',
                            'Históricos - Data', 'Históricos - Status']

    colunas_presentes = [col for col in colunas_de_interesse if col in df.columns]
    return df[colunas_presentes]


# Função para ler a terceira planilha
def ler_terceira_planilha():
    caminho = r'C:\Users\diego.brito\OneDrive - Ministério do Desenvolvimento e Assistência Social\Power BI\Python\Consulta Convênios\Consulta convenios.xlsx'
    df = pandas_library.read_excel(caminho, usecols=['Proposta', 'DataUploadMaisRecente'])
    df.columns = df.columns.str.strip()

    # Renomear 'DataUploadMaisRecente' para 'Projeto Básico / Termo de Referência'
    df.rename(columns={'DataUploadMaisRecente': 'Projeto Básico / Termo de Referência'}, inplace=True)

    df['Projeto Básico / Termo de Referência'] = pandas_library.to_datetime(df['Projeto Básico / Termo de Referência'], errors='coerce', dayfirst=True)
    return df


# Função para ler a quarta planilha
def ler_quarta_planilha():
    caminho = r'C:\Users\diego.brito\OneDrive - Ministério do Desenvolvimento e Assistência Social\Power BI\Python\Consulta Parecer\propostas_iniciais_Parecer.xlsx'
    try:
        df = pandas_library.read_excel(caminho, usecols=['Entidade', 'ValorDestinado', 'Município', 'UF'])
    except Exception as e:
        print(f"Erro ao ler a quarta planilha: {e}")
        return None

    # Remover espaços extras dos nomes das colunas
    df.columns = df.columns.str.strip()

    # Renomear 'ValorDestinado' para 'Valor de Repasse'
    df.rename(columns={'ValorDestinado': 'Valor de Repasse'}, inplace=True)

    return df


# Função para combinar as planilhas
def combinar_planilhas(df1, df2, df3, df4):
    # Mesclar as planilhas df1, df2 e df3 com base em 'Proposta'
    df_combinado = pandas_library.merge(df1, df2, on='Proposta', how='left')
    df_combinado = pandas_library.merge(df_combinado, df3, on='Proposta', how='left')

    # Adicionar os dados da quarta planilha (df4) diretamente
    df_combinado = pandas_library.concat([df_combinado, df4], axis=1)

    # Gerar a coluna de e-mails
    df_combinado['e-mail'] = df_combinado['TécnicoResponsável'].apply(gerar_endereco_de_email)

    # Ordem obrigatória das colunas
    ordem_colunas = [
        'Data da consulta', 'Proposta', 'Instrumento', 'OrigemRecurso',
        'CoordenaçãoResponsável', 'Processo', 'TécnicoResponsável', 'Entidade', 'Valor de Repasse',
        'AçãoOrçamentária', 'Município', 'UF', 'Certidões', 'Declarações', 'Comprovantes de Execução', 'Outros',
        'Históricos - Data', 'Históricos - Status', 'DataMaisRecenteProposta', 'DataMaisRecentePlanodeTrabalho',
        'Projeto Básico / Termo de Referência', 'Ação Necessária', 'e-mail'
    ]

    # Garantir que todas as colunas existam no DataFrame, criando com valores nulos se necessário
    for coluna in ordem_colunas:
        if coluna not in df_combinado.columns:
            print(f"Coluna ausente detectada: {coluna}. Criando com valores nulos.")
            df_combinado[coluna] = pandas_library.NA

    # Garantir que as colunas estejam na ordem especificada
    df_combinado = df_combinado[ordem_colunas]

    return df_combinado


# Função para salvar a planilha resultante
def salvar_planilha(df):
    caminho = r'C:\Users\diego.brito\OneDrive - Ministério do Desenvolvimento e Assistência Social\Power BI\Python\Consulta Transferegov Resultado Novo.xlsx'
    df.to_excel(caminho, index=False)
    print(f"Planilha combinada criada com sucesso em '{caminho}'!")


# Função principal para executar o processo
def executar_combinacao_das_planilhas():
    df1 = ler_primeira_planilha()
    df2 = ler_segunda_planilha()
    df3 = ler_terceira_planilha()
    df4 = ler_quarta_planilha()

    if any(df is None for df in [df1, df2, df3, df4]):
        print("Erro ao carregar uma ou mais planilhas. Verifique os arquivos e tente novamente.")
        return

    df_combinado = combinar_planilhas(df1, df2, df3, df4)
    salvar_planilha(df_combinado)


# Executar o processo
executar_combinacao_das_planilhas()
