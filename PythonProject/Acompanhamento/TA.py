import time
from datetime import datetime, timedelta
import pandas as pd
import os
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# 🔹 Caminhos dos arquivos
CAMINHO_ENTRADA = r"C:\Users\diego.brito\Downloads\robov1\CONTROLE DE PARCERIAS CGAP - Copia.xlsx"
CAMINHO_SAIDA = r"C:\Users\diego.brito\Downloads\robov1\saida1.xlsx"


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
    """Lê a planilha de entrada e retorna os instrumentos com status 'ATIVOS TODOS'."""
    try:
        df = pd.read_excel(CAMINHO_ENTRADA, dtype=str)
        df.columns = df.columns.str.strip()
    except Exception as erro:
        print(f"❌ Erro ao carregar a planilha: {erro}")
        exit()

    colunas_esperadas = ["Instrumento nº", "Técnico", "e-mail do Técnico", "Status"]
    colunas_faltando = [col for col in colunas_esperadas if col not in df.columns]

    if colunas_faltando:
        raise ValueError(f"🚨 Colunas ausentes na planilha: {colunas_faltando}")

    df_filtrado = df[df["Status"].str.upper() == "ATIVOS TODOS"]

    if df_filtrado.empty:
        print("⚠️ Nenhum instrumento com status 'ATIVOS TODOS' encontrado!")
        exit()

    return df_filtrado[colunas_esperadas]


def clicar(driver, xpath, descricao):
    """Tenta clicar em um elemento, se necessário via JavaScript."""
    wait = WebDriverWait(driver, 5)
    try:
        elemento = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
        elemento.click()
        print(f"✔ {descricao}")
    except:
        driver.execute_script("arguments[0].click();", elemento)
        print(f"✔ {descricao} (via JS)")


def inserir_texto(driver, xpath, texto, descricao):
    """Insere texto em um campo."""
    wait = WebDriverWait(driver, 5)
    try:
        campo = wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
        campo.clear()
        campo.send_keys(texto)
        print(f"✔ {descricao}: {texto}")
    except Exception as erro:
        print(f"⚠️ Erro ao inserir {descricao}: {erro}")


def extrair_data(driver, xpath, descricao):
    """Extrai e valida uma data no formato dd/mm/yyyy."""
    wait = WebDriverWait(driver, 5)
    try:
        elemento = wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
        texto = elemento.text.strip()

        if not texto:
            print(f"⚠️ {descricao} está vazia. Pulando instrumento.")
            return None

        try:
            data_formatada = datetime.strptime(texto, "%d/%m/%Y")
            print(f"✔ {descricao}: {data_formatada.strftime('%d/%m/%Y')}")
            return data_formatada
        except ValueError:
            print(f"⚠️ {descricao} não está em formato válido: '{texto}'. Pulando instrumento.")
            return None

    except Exception as erro:
        print(f"⚠️ Erro ao extrair {descricao}: {erro}")
        return None





def extrair_texto(driver, xpath, descricao):
    """Extrai texto de um campo específico."""
    wait = WebDriverWait(driver, 5)
    try:
        elemento = wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
        texto = elemento.text.strip()
        print(f"✔ {descricao}: {texto}")
        return texto
    except Exception:
        print(f"⚠️ Não foi possível extrair {descricao}")
        return ""


def automatizar_navegacao(driver, numero_instrumento):
    """Realiza a automação da pesquisa, análise do instrumento e envio de notificação se necessário."""
    print(f"\n➡️ Processando instrumento {numero_instrumento}...")

    # 🔹 Acessar o menu e a tela de pesquisa
    clicar(driver, "/html/body/div[1]/div[3]/div[1]/div[1]/div[1]/div[4]", "Abrindo menu principal")
    clicar(driver, "/html/body/div[1]/div[3]/div[2]/div[1]/div[1]/ul/li[6]/a", "Acessando pesquisa")

    # 🔹 Inserir o número do instrumento e pesquisar
    inserir_texto(driver, "/html/body/div[3]/div[15]/div[3]/div/div/form/table/tbody/tr[2]/td[2]/input",
                  numero_instrumento, "Pesquisando instrumento")
    clicar(driver, "/html/body/div[3]/div[15]/div[3]/div/div/form/table/tbody/tr[2]/td[2]/span/input",
           "Confirmando pesquisa")

    # 🔹 Selecionar o primeiro resultado da pesquisa
    clicar(driver, "/html/body/div[3]/div[15]/div[3]/div[3]/table/tbody/tr/td/div/a", "Selecionando primeiro resultado")

    # 🔍 **Extrair informações do instrumento**
    tipo = extrair_texto(driver, "/html/body/div[3]/div[15]/div[4]/div[1]/div/form/table/tbody/tr[1]",
                         "Tipo de instrumento")

    data_termino_texto = extrair_texto(driver, "/html/body/div[3]/div[15]/div[4]/div[1]/div/form/table/tbody/tr[33]",
                                       "Data término")

    # 🔹 Limpar o texto capturado para extrair apenas a data
    import re
    match = re.search(r"\d{2}/\d{2}/\d{4}", data_termino_texto)  # Captura a data no formato dd/mm/yyyy
    data_termino = None

    if match:
        try:
            data_termino = datetime.strptime(match.group(), "%d/%m/%Y")
            print(f"✔ Data de término válida: {data_termino.strftime('%d/%m/%Y')}")
        except ValueError:
            print(f"⚠️ Erro ao converter data: {data_termino_texto}")

    # 🔹 Se a data de término não for válida, volta para a pesquisa e pula para o próximo instrumento
    if not data_termino:
        print("⚠️ Data de término inválida. Pulando instrumento.")
        clicar(driver, "/html/body/div[3]/div[2]/div[1]/a", "Retornando para pesquisa")
        return

        # 🔹 Calcular a diferença de dias até o vencimento
    hoje = datetime.today()
    diferenca_dias = (data_termino - hoje).days

    # 🔹 Aplicar a regra dos dias faltando
    if (tipo.lower() == "termo de fomento" and 45 <= diferenca_dias <= 60) or \
            (tipo.lower() == "convênio" and 75 <= diferenca_dias <= 90):

        print("✅ Dentro do período necessário, enviando notificação...")

        clicar(driver, "/html/body/div[3]/div[2]/div[4]/div/div[7]", "Abrindo menu")
        clicar(driver, "/html/body/div[3]/div[2]/div[5]/div/div[2]/ul/li[1]/a", "Selecionando opção")
        clicar(driver, "/html/body/div[3]/div[15]/div[4]/div[1]/div/form/table/tbody/tr[7]/td/input",
               "Confirmando ação")

        # 🔹 Inserir a data do alerta (10 dias à frente)
        data_alerta = (hoje + timedelta(days=10)).strftime("%d/%m/%Y")
        inserir_texto(driver, "/html/body/div[1]/div[3]/form/fieldset/div/table/tbody/tr[9]/td[2]/span/input",
                      data_alerta, "Inserindo data de alerta")

        # 🔹 Definir o texto do alerta de acordo com o tipo do instrumento
        texto_alerta = "XXXXXXXXXXXXXXXXXXXXXXXXXX" if tipo.lower() == "termo de fomento" else "YYYYYYYYYYYYYYYYYYYYYYYY"
        inserir_texto(driver, "/html/body/div[1]/div[3]/form/fieldset/div/table/tbody/tr[10]/td[2]/textarea",
                      texto_alerta, "Inserindo texto de notificação")

    else:
        print("⚠️ Fora do período necessário, pulando instrumento.")

    # 🔄 **Voltar para a pesquisa para continuar o loop**
    clicar(driver, "/html/body/div[3]/div[2]/div[1]/a", "Retornando para pesquisa")


def main():
    driver = conectar_navegador_existente()
    df_entrada = ler_planilha_entrada()

    for _, linha in df_entrada.iterrows():
        automatizar_navegacao(driver, linha["Instrumento nº"])

    print("\n✅ Processamento concluído!")


if __name__ == "__main__":
    main()
