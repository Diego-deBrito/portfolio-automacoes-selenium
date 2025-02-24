import time
import os
import shutil
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager


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


# 📂 Ler planilha e extrair número de instrumento
def ler_planilha(arquivo=r"C:\Users\diego.brito\Downloads\robov1\CONTROLE DE PARCERIAS CGAP.xlsx"):
    df = pd.read_excel(arquivo, engine="openpyxl")

    # Garantir que a coluna "Instrumento nº" seja tratada corretamente
    df["Instrumento nº"] = df["Instrumento nº"].apply(
        lambda x: str(int(x)) if pd.notna(x) and isinstance(x, float) else str(x))

    # Filtrar apenas os ativos
    return df[df["Status"] == "ATIVOS TODOS"]





# 🔄 Navegar no menu principal
def navegar_menu_principal(driver, instrumento):
    try:
        elemento = esperar_elemento(driver, "/html/body/div[1]/div[3]/div[1]/div[1]/div[1]/div[4]")
        if elemento:
            elemento.click()

        elemento = esperar_elemento(driver, "/html[1]/body[1]/div[1]/div[3]/div[2]/div[1]/div[1]/ul[1]/li[6]/a[1]")
        if elemento:
            elemento.click()

        campo_pesquisa = esperar_elemento(driver,
                                          "/html[1]/body[1]/div[3]/div[15]/div[3]/div[1]/div[1]/form[1]/table[1]/tbody[1]/tr[2]/td[2]/input[1]")
        if campo_pesquisa:
            campo_pesquisa.clear()
            campo_pesquisa.send_keys(instrumento)

        elemento = esperar_elemento(driver,
                                    "/html[1]/body[1]/div[3]/div[15]/div[3]/div[1]/div[1]/form[1]/table[1]/tbody[1]/tr[2]/td[2]/span[1]/input[1]")
        if elemento:
            elemento.click()

        time.sleep(2)

        elemento = esperar_elemento(driver,
                                    "/html[1]/body[1]/div[3]/div[15]/div[3]/div[3]/table[1]/tbody[1]/tr[1]/td[1]/div[1]/a[1]")
        if elemento:
            elemento.click()

        return True
    except:
        print(f"⚠️ Instrumento {instrumento} não encontrado.")
        return False



# 📌 Navegar até a Tabela
def acessar_tabela(driver):
    try:
        print("📂 Acessando Aba da Tabela...")

        # 📌 1️⃣ Clicar na aba principal
        aba_principal = driver.execute_script("return document.querySelector(\"div[id='div_-481524888'] span span\");")
        if aba_principal:
            driver.execute_script("arguments[0].scrollIntoView();", aba_principal)
            driver.execute_script("arguments[0].click();", aba_principal)
            time.sleep(2)
        else:
            print("⚠️ Aba principal não encontrada!")
            return False

        # 📌 2️⃣ Clicar na aba secundária
        aba_secundaria = driver.execute_script(
            "return document.querySelector(\"a[id='menu_link_-481524888_1374656230'] div[class='inactiveTab'] span span\");")
        if aba_secundaria:
            driver.execute_script("arguments[0].scrollIntoView();", aba_secundaria)
            driver.execute_script("arguments[0].click();", aba_secundaria)
            time.sleep(2)
            return True
        else:
            print("⚠️ Sub Aba não encontrada!")
            return False

    except Exception as e:
        print(f"❌ Erro ao acessar a Aba da Tabela: {e}")
        return False










# 🔍 Espera um elemento estar visível
def esperar_elemento(driver, xpath, tempo=4):
    try:
        return WebDriverWait(driver, tempo).until(EC.presence_of_element_located((By.XPATH, xpath)))
    except:
        print(f"⚠️ Elemento {xpath} não encontrado!")
        return None


# 📌 Criar Pasta para cada Instrumento
def criar_pasta_instrumento(instrumento):
    """ Cria uma pasta exclusiva para armazenar os arquivos do instrumento pesquisado. """
    pasta_base = r"C:\Users\diego.brito\OneDrive - Ministério do Desenvolvimento e Assistência Social\Power BI\Python\Prestação de Contas - Contratos"
    pasta_destino = os.path.join(pasta_base, f"Instrumento_{instrumento}")

    if not os.path.exists(pasta_destino):
        os.makedirs(pasta_destino)
        print(f"📁 Criada pasta para o Instrumento {instrumento}")

    return pasta_destino


# ✅ Verificar e clicar no botão "Detalhar Contrato Original" se existir
def clicar_detalhar_contrato_original(driver):
    try:
        xpath_botao = "//input[@id='form_submit']"
        botao_detalhar = WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.XPATH, xpath_botao)))

        if botao_detalhar:
            print("📑 Botão 'Detalhar Contrato Original' encontrado! Clicando...")
            botao_detalhar.click()
            time.sleep(2)  # Aguardar carregamento da página

        return True  # Continua normalmente
    except:
        print("⚠️ Botão 'Detalhar Contrato Original' não encontrado, seguindo em frente...")
        return False  # Seguir sem erro


# 📌 Processar Tabela e Baixar Documentos
def processar_tabela(driver, instrumento):
    try:
        print("📂 Acessando tabela de contratos...")

        # Criar pasta para armazenar os documentos do instrumento
        pasta_destino = criar_pasta_instrumento(instrumento)

        # 📌 1️⃣ Identificar o total de páginas
        try:
            total_paginas = int(driver.execute_script(
                "return document.querySelector('#listaContratos > span').innerText.split(' de ')[1];"))
        except:
            total_paginas = 1  # Caso não consiga extrair o número
            print("⚠️ Não foi possível determinar o número de páginas. Assumindo 1 página.")

        print(f"📄 Total de páginas: {total_paginas}")

        # 📌 2️⃣ Percorrer Todas as Páginas
        for pagina in range(1, total_paginas + 1):
            print(f"➡️ Processando página {pagina}/{total_paginas}...")

            while True:
                try:
                    # 📌 3️⃣ Identificar o tbody correto dentro da tabela
                    tbody_element = driver.execute_script(
                        "return document.evaluate(\"/html/body/div[3]/div[15]/div[3]/div/div/form/div/div[2]/table/tbody\", document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;")

                    if not tbody_element:
                        print("⚠️ Tbody da tabela não encontrado.")
                        break

                    # 📌 4️⃣ Buscar botões "Detalhar" dentro do tbody
                    botoes_detalhar = tbody_element.find_elements(By.XPATH, ".//a[contains(text(),'Detalhar')]")

                    if not botoes_detalhar:
                        print("⚠️ Nenhum botão 'Detalhar' encontrado na página.")
                        break

                    # 📌 5️⃣ Percorrer os botões "Detalhar" e clicar um por um
                    for index in range(len(botoes_detalhar)):
                        try:
                            # Recarregar os botões antes de cada clique para evitar erro de referência inválida
                            tbody_element = driver.execute_script(
                                "return document.evaluate(\"/html/body/div[3]/div[15]/div[3]/div/div/form/div/div[2]/table/tbody\", document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;")
                            botoes_detalhar = tbody_element.find_elements(By.XPATH, ".//a[contains(text(),'Detalhar')]")

                            print(f"🔎 Clicando no botão 'Detalhar' {index + 1} de {len(botoes_detalhar)}...")

                            driver.execute_script("arguments[0].scrollIntoView();", botoes_detalhar[index])
                            driver.execute_script("arguments[0].click();", botoes_detalhar[index])
                            time.sleep(3)

                            # 📌 6️⃣ Baixar Documento
                            botao_download = driver.execute_script("return document.querySelector('.buttonLink');")
                            if botao_download:
                                driver.execute_script("arguments[0].click();", botao_download)
                                print("📂 Documento baixado!")
                                mover_para_pasta_instrumento(instrumento, pasta_destino)
                            else:
                                print("⚠️ Nenhum botão de download encontrado!")

                            # ✅ Verifica se o botão "Detalhar Contrato Original" está presente e clica nele, se necessário
                            clicar_detalhar_contrato_original(driver)

                            # 📌 7️⃣ Voltar para a Lista de Contratos
                            botao_voltar = driver.execute_script(
                                "return document.querySelector('input[value=\"Voltar\"]');")
                            if botao_voltar:
                                driver.execute_script("arguments[0].click();", botao_voltar)
                                time.sleep(3)  # Aguardar a página recarregar
                            else:
                                print("⚠️ Botão de voltar não encontrado!")

                        except Exception as e:
                            print(f"⚠️ Erro ao processar um botão 'Detalhar': {e}")
                            continue  # Pula para o próximo botão se houver erro

                    break  # Sai do loop após processar todos os botões da página

                except Exception as e:
                    print(f"⚠️ Erro ao recarregar os botões: {e}")
                    time.sleep(2)
                    continue  # Tenta recarregar os botões

            # 📌 🔟 Se houver mais páginas, clicar na próxima página
            if pagina < total_paginas:
                try:
                    botao_proxima_pagina = driver.execute_script(
                        f"return document.querySelector(\"a[onclick*='paginar({pagina + 1})']\");")
                    if botao_proxima_pagina:
                        driver.execute_script("arguments[0].scrollIntoView();", botao_proxima_pagina)
                        driver.execute_script("arguments[0].click();", botao_proxima_pagina)
                        time.sleep(3)  # Aguardar a nova página carregar
                    else:
                        print(f"⚠️ Botão para ir à página {pagina + 1} não encontrado.")
                except Exception as e:
                    print(f"⚠️ Erro ao tentar avançar para a página {pagina + 1}: {e}")

    except Exception as e:
        print(f"❌ Erro ao processar tabela: {e}")



# 📂 Mover Arquivo Baixado para a Pasta do Instrumento
def mover_para_pasta_instrumento(instrumento, pasta_destino):
    pasta_download = r"C:\Users\diego.brito\Downloads"

    time.sleep(5)  # Espera para garantir o download

    arquivos = os.listdir(pasta_download)

    for arquivo in arquivos:
        if arquivo.lower().endswith(('.pdf', '.docx', '.xlsx')):  # Suporte a mais tipos
            origem = os.path.join(pasta_download, arquivo)
            destino = os.path.join(pasta_destino, arquivo)
            shutil.move(origem, destino)
            print(f"📂 Documento {arquivo} movido para {pasta_destino}")


# 🚀 Executar o Robô
driver = conectar_navegador_existente()
df_entrada = ler_planilha()

for index, row in df_entrada.iterrows():
    instrumento = row["Instrumento nº"]
    print(f"🔎 Pesquisando Instrumento Nº: {instrumento}")

    if navegar_menu_principal(driver, instrumento):
        if acessar_tabela(driver):  # Certifica que acessamos a aba correta
            processar_tabela(driver, instrumento)
