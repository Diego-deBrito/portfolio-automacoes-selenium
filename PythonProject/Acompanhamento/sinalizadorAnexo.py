import time
from datetime import datetime

import pandas as pd
import os
from selenium import webdriver
from selenium.common import TimeoutException
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# Caminhos dos arquivos
CAMINHO_ENTRADA = r"C:\Users\diego.brito\Downloads\robov1\CONTROLE DE PARCERIAS CGAP.xlsx"
CAMINHO_SAIDA = r"C:\Users\diego.brito\Downloads\robov1\saida.xlsx"


def conectar_navegador_existente():
    """Conecta ao navegador Chrome já aberto na porta 9222."""
    options = webdriver.ChromeOptions()
    options.debugger_address = "localhost:9222"
    try:
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
        print("✅ Conectado ao navegador existente!")
        return driver
    except Exception as erro:
        print(f"❌ Erro ao conectar ao navegador: {erro}")
        exit()



def ler_planilha_entrada():
    """Lê a planilha de entrada e retorna apenas os instrumentos com status 'ATIVOS TODOS'."""
    try:
        df = pd.read_excel(CAMINHO_ENTRADA, dtype=str)
        df.columns = df.columns.str.strip()  # Remove espaços extras dos nomes das colunas
    except Exception as erro:
        print(f"❌ Erro ao carregar a planilha de entrada: {erro}")
        exit()

    colunas_esperadas = ["Instrumento nº", "Técnico", "e-mail do Técnico", "Status"]

    # Verifica se todas as colunas esperadas estão presentes
    colunas_faltando = [col for col in colunas_esperadas if col not in df.columns]
    if colunas_faltando:
        raise ValueError(f"🚨 Erro: As colunas {colunas_faltando} não foram encontradas na planilha!")

    # Filtra apenas os instrumentos com Status == "ATIVOS TODOS"
    df_filtrado = df[df["Status"].str.upper() == "ATIVOS TODOS"]

    if df_filtrado.empty:
        print("⚠️ Nenhum instrumento com status 'ATIVOS TODOS' encontrado!")
        exit()

    return df_filtrado[colunas_esperadas]




def salvar_dado_extracao(numero_instrumento, tecnico, email, status, data_upload):
    """Salva os dados extraídos na planilha de saída, garantindo que todas as colunas sejam mantidas."""

    if data_upload is None:
        print(f"⚠️ Nenhuma data extraída para {numero_instrumento}. Pulando...")
        return

    try:
        # 🔥 Definição correta das colunas esperadas
        colunas_necessarias = ["Instrumento nº", "Técnico", "e-mail do Técnico", "Status", "Data Upload"]

        # Se o arquivo de saída já existe, carregamos
        if os.path.exists(CAMINHO_SAIDA):
            df_saida = pd.read_excel(CAMINHO_SAIDA, dtype=str)
            print("📂 Planilha de saída encontrada. Atualizando dados...")
        else:
            print("📂 Criando nova planilha de saída...")
            df_saida = pd.DataFrame(columns=colunas_necessarias)

        # 🚀 Garante que TODAS as colunas necessárias existam na planilha de saída
        for coluna in colunas_necessarias:
            if coluna not in df_saida.columns:
                df_saida[coluna] = ""

        # 🔄 Atualiza ou adiciona os dados
        if numero_instrumento in df_saida["Instrumento nº"].values:
            df_saida.loc[df_saida["Instrumento nº"] == numero_instrumento, "Técnico"] = tecnico
            df_saida.loc[df_saida["Instrumento nº"] == numero_instrumento, "e-mail do Técnico"] = email
            df_saida.loc[df_saida["Instrumento nº"] == numero_instrumento, "Status"] = status
            df_saida.loc[df_saida["Instrumento nº"] == numero_instrumento, "Data Upload"] = data_upload
            print(f"✏️ Atualizando 'Data Upload' para {numero_instrumento}: {data_upload}")
        else:
            # 🆕 Adiciona um novo registro sem sobrescrever o arquivo
            novo_dado = pd.DataFrame([[numero_instrumento, tecnico, email, status, data_upload]],
                                     columns=colunas_necessarias)
            df_saida = pd.concat([df_saida, novo_dado], ignore_index=True)
            print(f"➕ Adicionando novo registro para {numero_instrumento}: {data_upload}")

        # 🔄 Salva a planilha garantindo que todas as colunas sejam preservadas
        df_saida.to_excel(CAMINHO_SAIDA, index=False)
        print("✅ Planilha atualizada com sucesso!")

    except PermissionError:
        print("❌ ERRO: O arquivo está aberto no Excel. Feche-o e tente novamente.")
    except Exception as erro:
        print(f"❌ ERRO ao salvar dados no Excel: {erro}")

def automatizar_navegacao(driver, numero_instrumento):
    """Realiza a automação do site seguindo os passos indicados."""
    wait = WebDriverWait(driver, 5)

    def clicar(xpath, descricao=""):
        """ Aguarda o elemento estar disponível e clica """
        try:
            elemento = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
            elemento.click()
            print(f"✔ {descricao}")
        except Exception as erro:
            print(f"⚠️ Erro ao clicar ({descricao}): {erro}")

    print(f"\n➡️ Processando instrumento {numero_instrumento}...")

    # Acessar menu principal
    clicar("/html/body/div[1]/div[3]/div[1]/div[1]/div[1]/div[4]", "Acessando menu principal")
    clicar("/html/body/div[1]/div[3]/div[2]/div[1]/div[1]/ul/li[6]/a", "Acessando pesquisa")

    # Inserir Número do Instrumento
    try:
        input_field = wait.until(EC.presence_of_element_located(
            (By.XPATH, "/html/body/div[3]/div[15]/div[3]/div/div/form/table/tbody/tr[2]/td[2]/input")))
        input_field.clear()
        input_field.send_keys(numero_instrumento)
        clicar("/html/body/div[3]/div[15]/div[3]/div/div/form/table/tbody/tr[2]/td[2]/span/input",
               "Pesquisando instrumento")
    except Exception as erro:
        print(f"⚠️ Erro ao inserir número do instrumento {numero_instrumento}: {erro}")
        return None

    # Selecionar primeiro resultado da pesquisa
    clicar("/html/body/div[3]/div[15]/div[3]/div[3]/table/tbody/tr/td/div/a", "Selecionando primeiro resultado")
    clicar("/html/body/div[3]/div[15]/div[1]/div/div[1]/a[2]/div", "Abrindo menu")
    clicar("/html/body/div[3]/div[15]/div[1]/div/div[2]/a[8]/div", "Acessando detalhes")
    clicar("/html/body/div[3]/div[15]/div[3]/div[1]/div/form/table/tbody/tr/td[2]/input[2]",
           "Visualizando data de upload")

    # 🔹 Capturar a data do último anexo
    data_upload_extraida = capturar_data_ultimo_anexo(driver)

    # 🔄 Voltar para a pesquisa inicial
    try:
        driver.find_element(By.XPATH, "/html/body/div[3]/div[2]/div[1]/a").click()
        print("🔄 Retornando para a pesquisa inicial...")
    except Exception:
        print("⚠️ Não foi possível voltar, tentando continuar...")

    return data_upload_extraida  # ✅ Agora retorna um único valor corretamente



# Função para capturar a data do último anexo
def capturar_data_ultimo_anexo(driver):
    # Localiza a tabela pelo XPath
    xpath_tabela = "/html/body/div[3]/div[15]/div[4]/div/div[1]/form/div/div[1]/table"
    tabela = driver.find_element(By.XPATH, xpath_tabela)

    # Coletar todas as linhas da tabela
    linhas = tabela.find_elements(By.TAG_NAME, "tr")

    datas = []

    # Iterar sobre as linhas (ignorando o cabeçalho)
    for linha in linhas[1:]:
        colunas = linha.find_elements(By.TAG_NAME, "td")

        # Ajuste conforme a posição da coluna "Data Upload" (exemplo: índice 2)
        data_texto = colunas[2].text.strip()

        try:
            # Converte string para formato de data (ajuste conforme o formato correto da página)
            data_formatada = datetime.strptime(data_texto, "%d/%m/%Y")
            datas.append(data_formatada)
        except ValueError:
            print(f"Formato inválido: {data_texto}")

    # Identifica a data mais recente
    if datas:
        data_mais_recente = max(datas)
        print("Data mais recente:", data_mais_recente.strftime("%d/%m/%Y"))

        # Criar dataframe para salvar no Excel
        df = pd.DataFrame({"Data Mais Recente": [data_mais_recente.strftime("%d/%m/%Y")]})

        # Salvar em Excel
        output_file = r"C:\Users\diego.brito\Downloads\robov1\saida.xlsx"
        df.to_excel(output_file, index=False)

        print(f"Data salva em {output_file} com sucesso! ✅")
    else:
        print("Nenhuma data válida foi encontrada. ❌")


def main():
    """Fluxo principal do código."""
    driver = conectar_navegador_existente()
    df_entrada = ler_planilha_entrada()  # Agora só contém instrumentos ATIVOS

    for _, linha in df_entrada.iterrows():
        numero_do_instrumento = linha["Instrumento nº"]
        tecnico_responsavel = linha["Técnico"]
        email_do_tecnico = linha["e-mail do Técnico"]
        status_instrumento = linha["Status"]  # ✅ Correção: Adicionando o status

        # Obtém a data real do último anexo
        data_mais_recente = automatizar_navegacao(driver, numero_do_instrumento)

        # 🔥 Correção: Agora passamos TODOS os dados corretamente
        salvar_dado_extracao(numero_do_instrumento, tecnico_responsavel, email_do_tecnico, status_instrumento, data_mais_recente)

    print("\n✅ Processamento concluído!")


if __name__ == "__main__":
    main()