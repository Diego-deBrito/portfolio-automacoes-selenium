import time
import pandas as pd
import os
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# 📌 Caminhos das planilhas
PLANILHA_ENTRADA = r"C:\Users\diego.brito\Downloads\robov1\CONTROLE DE PARCERIAS CGAP.xlsx"
PLANILHA_SAIDA = r"C:\Users\diego.brito\Downloads\robov1\resultado_extracao.xlsx"

# 🛠 Conectar ao navegador já aberto
def conectar_navegador_existente():
    options = webdriver.ChromeOptions()
    options.debugger_address = "localhost:9222"
    try:
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
        print("✅ Conectado ao navegador existente!")
        return driver
    except Exception as erro:
        print(f"❌ Erro ao conectar ao navegador: {erro}")
        exit()

# 📂 Ler planilha de entrada e extrair os números dos instrumentos
def ler_planilha():
    df = pd.read_excel(PLANILHA_ENTRADA, engine="openpyxl")

    # Garantir que a coluna "Instrumento nº" seja tratada corretamente
    df["Instrumento nº"] = df["Instrumento nº"].apply(
        lambda x: str(int(x)) if pd.notna(x) and isinstance(x, float) else str(x))

    return df

# 🔍 Esperar um elemento estar visível e retorná-lo
def esperar_elemento(driver, xpath, tempo=10):
    try:
        return WebDriverWait(driver, tempo).until(EC.presence_of_element_located((By.XPATH, xpath)))
    except:
        print(f"⚠️ Elemento {xpath} não encontrado!")
        return None

# 🔄 Executar a busca e extrair os dados
def extrair_dados(driver, instrumento):
    try:
        print(f"🔎 Pesquisando Instrumento Nº: {instrumento}")

        # 📌 1️⃣ Clicar no menu principal
        elemento = esperar_elemento(driver, "/html/body/div[1]/div[3]/div[1]/div[1]/div[1]/div[4]")
        if elemento:
            elemento.click()
        else:
            return None  # Retorna se o elemento não for encontrado

        # 📌 2️⃣ Clicar na opção do menu
        elemento = esperar_elemento(driver, "/html/body/div[1]/div[3]/div[2]/div[1]/div[1]/ul/li[6]/a")
        if elemento:
            elemento.click()
        else:
            return None

        # 📌 3️⃣ Inserir o número do instrumento no campo de pesquisa
        campo_pesquisa = esperar_elemento(driver, "/html/body/div[3]/div[15]/div[3]/div/div/form/table/tbody/tr[2]/td[2]/input")
        if campo_pesquisa:
            campo_pesquisa.clear()
            campo_pesquisa.send_keys(instrumento)
        else:
            return None

        # 📌 4️⃣ Clicar no botão de pesquisa
        elemento = esperar_elemento(driver, "/html/body/div[3]/div[15]/div[3]/div/div/form/table/tbody/tr[2]/td[2]/span/input")
        if elemento:
            elemento.click()
        else:
            return None

        time.sleep(2)

        # 📌 5️⃣ Clicar no primeiro resultado da pesquisa
        elemento = esperar_elemento(driver, "/html/body/div[3]/div[15]/div[3]/div[3]/table/tbody/tr/td/div/a")
        if elemento:
            elemento.click()
        else:
            return None

        time.sleep(2)

        # 📌 6️⃣ Extrair a data do campo especificado
        data_elemento = esperar_elemento(driver, "/html/body/div[3]/div[15]/div[4]/div[1]/div/form/table/tbody/tr[35]/td[2]")
        if data_elemento:
            data_extraida = data_elemento.text.strip()
            print(f"📅 Data extraída: {data_extraida}")
            return data_extraida
        else:
            print("⚠️ Não foi possível extrair a data.")
            return "Data não encontrada"

    except Exception as e:
        print(f"❌ Erro ao extrair dados do instrumento {instrumento}: {e}")
        return "Erro"

# 📤 Salvar planilha de saída
def salvar_planilha(df):
    try:
        if os.path.exists(PLANILHA_SAIDA):
            os.remove(PLANILHA_SAIDA)
        df.to_excel(PLANILHA_SAIDA, index=False)
        print(f"📂 Planilha salva em: {PLANILHA_SAIDA}")
    except PermissionError:
        print(f"⚠️ Erro: Feche o arquivo {PLANILHA_SAIDA} antes de salvar.")

# 🚀 Fluxo principal do robô
def executar_robo():
    driver = conectar_navegador_existente()
    df_entrada = ler_planilha()

    # Criar DataFrame para saída
    df_saida = pd.DataFrame(columns=["Instrumento nº", "Data Extraída"])

    print("🚀 Iniciando extração de dados...")

    for index, row in df_entrada.iterrows():
        instrumento = row["Instrumento nº"]

        # 🔹 Pular instrumentos inválidos
        if not instrumento or instrumento in ["nan", "None", ""]:
            print(f"⚠️ Instrumento inválido encontrado na linha {index + 1}. Pulando...")
            continue

        # 🔹 Extrair dados do site
        data_extraida = extrair_dados(driver, instrumento)

        # 🔹 Adicionar ao DataFrame de saída
        df_saida = pd.concat([df_saida, pd.DataFrame([[instrumento, data_extraida]], columns=df_saida.columns)])

        # 📌 Salvar a planilha após cada extração
        salvar_planilha(df_saida)

    print("✅ Processamento concluído! Planilha salva com sucesso.")

# 🔥 Executar o robô
executar_robo()
