import time
import pandas as pd
import os
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# 📂 Caminho da planilha
CAMINHO_PLANILHA_ENTRADA = r"C:\Users\diego.brito\Downloads\robov1\resultado_main_2.xlsx"


# 🛠 Conectar ao navegador já aberto
def conectar_navegador_existente():
    options = webdriver.ChromeOptions()
    options.debugger_address = "127.0.0.1:9222"  # Conectar ao Chrome já aberto

    try:
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
        print("✅ Conectado ao navegador existente!")
        return driver
    except Exception as erro:
        print(f"❌ Erro ao conectar ao navegador: {erro}")
        exit()


# 📂 Ler a planilha de entrada e filtrar apenas os registros necessários
def ler_planilha(arquivo):
    df = pd.read_excel(arquivo, engine="openpyxl", dtype={"Instrumento": str})  # Evita problemas com ".0"

    # Garantir que as colunas de saída são do tipo string
    if "Situação" not in df.columns:
        df["Situação"] = ""
    if "Data Término Vigência" not in df.columns:
        df["Data Término Vigência"] = ""

    df["Situação"] = df["Situação"].astype(str)
    df["Data Término Vigência"] = df["Data Término Vigência"].astype(str)

    # Filtrar apenas os instrumentos onde a aba Anexos tem "Sem anexo encontrado"
    df_filtrado = df[df["Aba Anexos"] == "Sem anexo encontrado"]

    if df_filtrado.empty:
        print("⚠️ Nenhum instrumento com 'Sem anexo encontrado' para processar.")
        exit()

    return df, df_filtrado


# 📤 Salvar os dados atualizados na planilha (após cada instrumento)
def salvar_planilha(df, arquivo):
    df.to_excel(arquivo, index=False)
    print(f"📂 Planilha salva: {arquivo}")


# 🔍 Esperar elemento estar visível com timeout menor
def esperar_elemento(driver, xpath, tempo=2):
    """ Aguarda um elemento aparecer na página, reduzindo tempo de espera """
    try:
        return WebDriverWait(driver, tempo).until(EC.presence_of_element_located((By.XPATH, xpath)))
    except:
        return None


# 📌 Capturar Situação (testando múltiplos XPaths)
def pegar_situacao(driver):
    """ Captura a Situação do instrumento na página """
    xpaths_possiveis = [
        "/html[1]/body[1]/div[3]/div[15]/div[4]/div[1]/div[1]/form[1]/table[1]/tbody[1]/tr[4]/td[2]/table[1]/tbody[1]/tr[1]/td[1]/div[1]",
        "/html/body/div[3]/div[15]/div[4]/div[1]/div/form/table/tbody/tr[4]/td[2]/table/tbody/tr[1]/td/div",
        '//*[@id="tr-alterarStatus"]/td[2]/table/tbody/tr[1]/td/div'
    ]

    for xpath in xpaths_possiveis:
        situacao_elemento = esperar_elemento(driver, xpath)
        if situacao_elemento:
            return situacao_elemento.text.strip()

    return "Não encontrado"


# 📌 Capturar Data de Término de Vigência (testando múltiplos XPaths)
def pegar_data_termino(driver):
    """ Captura a Data de Término de Vigência do instrumento na página """
    xpaths_possiveis = [
        "/html/body/div[3]/div[15]/div[4]/div[1]/div/form/table/tbody/tr[38]/td[2]",
        "/html[1]/body[1]/div[3]/div[15]/div[4]/div[1]/div[1]/form[1]/table[1]/tbody[1]/tr[38]/td[2]",
        '//*[@id="tr-alterarTerminoVigencia"]/td[@class="field"]'
    ]

    for xpath in xpaths_possiveis:
        data_elemento = esperar_elemento(driver, xpath)
        if data_elemento:
            data_termino = data_elemento.text.strip()

            # 🔹 **Correção: Formatar a data corretamente**
            try:
                return pd.to_datetime(data_termino, dayfirst=True).strftime("%d/%m/%Y")
            except ValueError:
                print(f"⚠️ Erro ao converter data: {data_termino}")

    return "Data não encontrada"


# 🏁 Executar o fluxo do robô
def executar_robô():
    driver = conectar_navegador_existente()
    df, df_filtrado = ler_planilha(CAMINHO_PLANILHA_ENTRADA)

    for index, row in df_filtrado.iterrows():
        instrumento = row["Instrumento"]
        print(f"\n🔍 Processando Instrumento: {instrumento}")

        try:
            # Acessar menu rapidamente
            driver.find_element(By.XPATH, "/html/body/div[1]/div[3]/div[1]/div[1]/div[1]/div[4]").click()
            driver.find_element(By.XPATH, "/html/body/div[1]/div[3]/div[2]/div[1]/div[1]/ul[1]/li[6]/a[1]").click()

            # Preencher campo do instrumento
            campo_instrumento = driver.find_element(By.XPATH,
                                                    "/html/body/div[3]/div[15]/div[3]/div[1]/div[1]/form/table/tbody/tr[2]/td[2]/input")
            campo_instrumento.clear()
            campo_instrumento.send_keys(instrumento)

            # Submeter busca
            driver.find_element(By.XPATH,
                                "/html/body/div[3]/div[15]/div[3]/div[1]/div[1]/form/table/tbody/tr[2]/td[2]/span/input").click()

            # Clicar no primeiro resultado
            driver.find_element(By.XPATH,
                                "/html[1]/body[1]/div[3]/div[15]/div[3]/div[3]/table[1]/tbody[1]/tr[1]/td[1]/div[1]/a[1]").click()

            # Capturar dados
            situacao = pegar_situacao(driver)
            print(f"📌 Situação: {situacao}")

            data_termino = pegar_data_termino(driver)
            print(f"📅 Data de Término: {data_termino}")

            # Atualizar a planilha
            df.loc[df["Instrumento"] == instrumento, "Situação"] = str(situacao)
            df.loc[df["Instrumento"] == instrumento, "Data Término Vigência"] = str(data_termino)
            salvar_planilha(df, CAMINHO_PLANILHA_ENTRADA)

            # Voltar rapidamente para próxima iteração
            driver.find_element(By.XPATH, "/html/body/div[3]/div[2]/div[1]/a").click()

        except Exception as erro:
            print(f"❌ Erro ao processar o instrumento {instrumento}: {erro}")
            continue

    print("🎉 Processo concluído!")


if __name__ == "__main__":
    executar_robô()










