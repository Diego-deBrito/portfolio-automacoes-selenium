import time
import pandas as pd
import os
from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from datetime import datetime

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

# 📂 Ler planilha de entrada
def ler_planilha(arquivo=r"C:\Users\diego.brito\Downloads\robov1\CONTROLE DE PARCERIAS CGAP.xlsx"):
    df = pd.read_excel(arquivo, engine="openpyxl")
    return df[df["Status"] == "ATIVOS TODOS"]

# 📤 Salvar planilha de saída
def salvar_planilha(df, arquivo=r"C:\Users\diego.brito\Downloads\robov1\resultado_main_2.xlsx"):
    try:
        if os.path.exists(arquivo):
            os.remove(arquivo)
        df.to_excel(arquivo, index=False)
        print(f"📂 Planilha salva em: {arquivo}")
    except PermissionError:
        print(f"⚠️ Erro: Feche o arquivo {arquivo} antes de salvar.")

# 🔍 Espera um elemento estar visível
def esperar_elemento(driver, xpath, tempo=10):
    try:
        return WebDriverWait(driver, tempo).until(EC.presence_of_element_located((By.XPATH, xpath)))
    except:
        print(f"⚠️ Elemento {xpath} não encontrado!")
        return None

# 🔄 Navegar no menu principal
def navegar_menu_principal(driver, instrumento):
    try:
        esperar_elemento(driver, "/html/body/div[1]/div[3]/div[1]/div[1]/div[1]/div[4]").click()
        esperar_elemento(driver, "/html[1]/body[1]/div[1]/div[3]/div[2]/div[1]/div[1]/ul[1]/li[6]/a[1]").click()
        campo_pesquisa = esperar_elemento(driver, "/html[1]/body[1]/div[3]/div[15]/div[3]/div[1]/div[1]/form[1]/table[1]/tbody[1]/tr[2]/td[2]/input[1]")
        campo_pesquisa.clear()
        campo_pesquisa.send_keys(instrumento)
        esperar_elemento(driver, "/html[1]/body[1]/div[3]/div[15]/div[3]/div[1]/div[1]/form[1]/table[1]/tbody[1]/tr[2]/td[2]/span[1]/input[1]").click()
        time.sleep(2)
        esperar_elemento(driver, "/html[1]/body[1]/div[3]/div[15]/div[3]/div[3]/table[1]/tbody[1]/tr[1]/td[1]/div[1]/a[1]").click()
        return True
    except:
        print(f"⚠️ Instrumento {instrumento} não encontrado.")
        return False


def processar_aba_ajustes(driver):
    """Acessa a aba Ajustes do PT e retorna a situação mais recente com base no maior ano da coluna 'Número'."""

    situacao_encontrada = "Sem Registros"  # Valor padrão caso a tabela esteja vazia
    maior_ano = -1  # Inicializa com um valor baixo para comparação

    try:
        print("📂 Acessando Aba Ajustes do PT...")

        # 🏷️ Passo 1: Clicar na aba Ajustes do PT
        esperar_elemento(driver,
                         "/html[1]/body[1]/div[3]/div[15]/div[1]/div[1]/div[1]/a[6]/div[1]/span[1]/span[1]").click()

        # 🏷️ Passo 2: Acessar a segunda aba dentro de Ajustes do PT
        esperar_elemento(driver,
                         "/html[1]/body[1]/div[3]/div[15]/div[1]/div[1]/div[2]/a[16]/div[1]/span[1]/span[1]").click()

        # 📌 Passo 3: Esperar a tabela carregar
        tabela_ajustes = esperar_elemento(driver, "/html/body/div[3]/div[15]/div[3]/div[2]/div[2]/table")

        # 📌 Passo 4: Coletar todas as linhas da tabela
        linhas = tabela_ajustes.find_elements(By.TAG_NAME, "tr")

        for linha in linhas[1:]:  # Ignorar o cabeçalho
            colunas = linha.find_elements(By.TAG_NAME, "td")

            if len(colunas) >= 4:  # Garantir que há pelo menos 4 colunas na linha
                numero_texto = colunas[0].text.strip()  # Coluna "Número" (1ª coluna)
                situacao_texto = colunas[1].text.strip()  # Coluna "Situação" (2ª coluna)

                # 📌 Extrai o ano da coluna "Número" (Ex: "1/2024" -> 2024)
                match = re.search(r'/(\d{4})$', numero_texto)
                if match:
                    ano = int(match.group(1))

                    # 📌 Verifica se este é o maior ano encontrado até agora
                    if ano > maior_ano:
                        maior_ano = ano
                        situacao_encontrada = situacao_texto

        # 📌 Se encontrou uma situação válida, imprime
        print(f"📌 Situação mais recente ({maior_ano}): {situacao_encontrada}")

    except Exception as e:
        print(f"⚠️ Erro ao processar Aba Ajustes do PT: {e}")

    return situacao_encontrada  # Retorna a situação da linha com o maior ano


# 📌 Processar Aba TA
from selenium.webdriver.common.by import By
import re


def processar_aba_TA(driver):
    """Acessa a Aba TA e retorna a situação mais recente com base no maior ano da coluna 'Número'."""

    situacao_encontrada = "Sem Registros"  # Valor padrão caso a tabela esteja vazia
    maior_ano = -1  # Inicializa com um valor baixo para comparação

    try:
        print("📂 Acessando a Aba TA...")

        # 🏷️ Passo 1: Clicar na Aba TA
        aba_TA = esperar_elemento(driver,
                                  "/html[1]/body[1]/div[3]/div[15]/div[1]/div[1]/div[2]/a[20]/div[1]/span[1]/span[1]")
        driver.execute_script("arguments[0].scrollIntoView();", aba_TA)
        aba_TA.click()
        time.sleep(2)  # Pequena espera para carregamento

        # 📌 Passo 2: Esperar a tabela carregar
        print("🔍 Buscando registros na tabela de TA...")
        tabela_TA = esperar_elemento(driver, "/html/body/div[3]/div[15]/div[4]/div/form/div/div[3]")
        linhas = tabela_TA.find_elements(By.TAG_NAME, "tr")

        for linha in linhas[1:]:  # Ignorar o cabeçalho
            colunas = linha.find_elements(By.TAG_NAME, "td")

            if len(colunas) >= 4:  # Garantir que há pelo menos 4 colunas
                numero_texto = colunas[0].text.strip()  # Coluna "Número" (1ª coluna)
                situacao_texto = colunas[1].text.strip()  # Coluna "Situação" (2ª coluna)

                # 📌 Extrai o ano da coluna "Número" (Ex: "1/2024" -> 2024)
                match = re.search(r'/(\d{4})$', numero_texto)
                if match:
                    ano = int(match.group(1))

                    # 📌 Verifica se este é o maior ano encontrado até agora
                    if ano > maior_ano:
                        maior_ano = ano
                        situacao_encontrada = situacao_texto

        # 📌 Se encontrou uma situação válida, imprime
        print(f"📌 Situação mais recente ({maior_ano}): {situacao_encontrada}")

    except Exception as e:
        print(f"⚠️ Erro ao processar a Aba TA: {e}")

    return situacao_encontrada  # Retorna a situação mais recente com base no maior ano


# 📌 Processar Aba Rendimento de Aplicação
def processar_aba_rendimento_aplicacao(driver):
    """ Acessa a aba de Rendimento de Aplicação e verifica o status da solicitação. """
    try:
        print("📂 Acessando Aba Rendimento de Aplicação...")

        # 📌 Passo 1: Clicar na aba correta
        aba_rendimento = esperar_elemento(driver, "/html[1]/body[1]/div[3]/div[15]/div[1]/div[1]/div[2]/a[28]/div[1]/span[1]/span[1]")
        driver.execute_script("arguments[0].scrollIntoView();", aba_rendimento)
        aba_rendimento.click()
        time.sleep(2)  # Pequena espera para carregamento

        # 📌 Passo 2: Aguardar a tabela carregar
        tabela = esperar_elemento(driver, "/html/body/div[3]/div[15]/div[7]")

        # 📌 Passo 3: Procurar pela coluna "Status da Solicitação"
        linhas = tabela.find_elements(By.TAG_NAME, "tr")

        for linha in linhas:
            colunas = linha.find_elements(By.TAG_NAME, "td")

            if len(colunas) >= 1:  # Verifica se há pelo menos uma coluna na linha
                status_solicitacao = colunas[0].text.strip()  # Pegando a primeira coluna

                if "Em análise" in status_solicitacao:
                    print(f"✅ Situação encontrada: {status_solicitacao}")
                    return status_solicitacao  # Retorna o status encontrado

        print("⚠️ Nenhum status 'Em análise' encontrado.")
        return "Sem registro"

    except Exception as e:
        print(f"⚠️ Erro ao processar Aba Rendimento de Aplicação: {e}")
        return "Erro ao processar"



from selenium.webdriver.common.by import By
from datetime import datetime

def processar_aba_anexos(driver):
    """Acessa a aba de Anexos e extrai a Data Upload mais recente. Se não encontrar, registra e continua."""
    try:
        print("📂 Acessando Aba de Anexos...")

        # 📌 Passo 1: Acessar a aba correta
        aba_anexos_primaria = esperar_elemento(driver, "/html[1]/body[1]/div[3]/div[15]/div[1]/div[1]/div[1]/a[2]/div[1]/span[1]/span[1]", 5)
        if not aba_anexos_primaria:
            print("⚠️ Aba de Anexos não encontrada. Registrando como 'Sem anexo encontrado'.")
            return "Sem anexo encontrado"

        driver.execute_script("arguments[0].scrollIntoView();", aba_anexos_primaria)
        aba_anexos_primaria.click()
        time.sleep(1)

        aba_anexos_secundaria = esperar_elemento(driver, "/html[1]/body[1]/div[3]/div[15]/div[1]/div[1]/div[2]/a[8]/div[1]/span[1]/span[1]", 5)
        if not aba_anexos_secundaria:
            print("⚠️ Aba secundária de Anexos não encontrada. Registrando como 'Sem anexo encontrado'.")
            return "Sem anexo encontrado"

        driver.execute_script("arguments[0].scrollIntoView();", aba_anexos_secundaria)
        aba_anexos_secundaria.click()
        time.sleep(1)

        # 📌 Passo 2: Clicar no botão de pesquisa para carregar a lista de anexos
        botao_pesquisar = esperar_elemento(driver, "/html[1]/body[1]/div[3]/div[15]/div[3]/div[1]/div[1]/form[1]/table[1]/tbody[1]/tr[1]/td[2]/input[2]", 5)
        if not botao_pesquisar:
            print("⚠️ Botão de pesquisa de anexos não encontrado. Registrando como 'Sem anexo encontrado'.")
            return "Sem anexo encontrado"

        driver.execute_script("arguments[0].click();", botao_pesquisar)
        time.sleep(2)

        # 📌 Passo 3: Aguardar a tabela carregar
        tabela_anexos = esperar_elemento(driver, "/html/body/div[3]/div[15]/div[4]/div/div[1]/form/div/div[1]/table", 5)
        if not tabela_anexos:
            print("⚠️ Tabela de anexos não encontrada. Registrando como 'Sem anexo encontrado'.")
            return "Sem anexo encontrado"

        # 📌 Passo 4: Coletar todas as linhas da tabela
        linhas = tabela_anexos.find_elements(By.TAG_NAME, "tr")
        datas_upload = []  # Lista para armazenar as datas encontradas

        for linha in linhas[1:]:  # Ignorar cabeçalho
            colunas = linha.find_elements(By.TAG_NAME, "td")

            if len(colunas) >= 3:  # Garante que há pelo menos 3 colunas
                data_texto = colunas[2].text.strip()  # Pegando a coluna 'Data Upload'

                if data_texto:
                    try:
                        data_formatada = datetime.strptime(data_texto, "%d/%m/%Y")
                        datas_upload.append(data_formatada)
                    except ValueError:
                        print(f"⚠️ Data inválida ignorada: {data_texto}")

        # 📌 Passo 5: Se houver datas, pegar a mais recente
        if datas_upload:
            data_upload_recente = max(datas_upload).strftime("%d/%m/%Y")
            print(f"📅 Data mais recente na coluna 'Data Upload': {data_upload_recente}")
            return data_upload_recente
        else:
            print("⚠️ Nenhuma data válida encontrada na coluna 'Data Upload'.")
            return "Sem registro"

    except Exception as e:
        print(f"⚠️ Erro ao processar Aba de Anexos: {e}")
        return "Erro ao processar"



# 🚀 Fluxo principal do robô
def executar_robo():
    """ Executa o robô navegando entre as abas e coletando os dados. """
    driver = conectar_navegador_existente()
    df_entrada = ler_planilha()

    # 🔹 Corrigir possíveis NaNs substituindo por string vazia
    df_entrada["Instrumento nº"] = df_entrada["Instrumento nº"].fillna("").astype(str)

    dados_saida = []
    planilha_criada = False  # Flag para indicar se a planilha já foi gerada

    print("🚀 Iniciando processamento dos instrumentos...")

    for index, row in df_entrada.iterrows():
        instrumento = row["Instrumento nº"].strip()

        # 🔍 Se o campo estiver vazio, pula para o próximo
        if not instrumento:
            print(f"⚠️ Instrumento vazio na linha {index + 1}. Pulando para o próximo...")
            continue

        print(f"\n🔎 Processando Instrumento Nº: {instrumento} ({index + 1}/{len(df_entrada)})")

        try:
            if not navegar_menu_principal(driver, instrumento):
                print(f"⚠️ Instrumento {instrumento} não encontrado. Pulando para o próximo...")
                continue

            # Chamando funções de processamento de cada aba
            data_ajustes = processar_aba_ajustes(driver)
            data_ta = processar_aba_TA(driver)
            status_registro = processar_aba_rendimento_aplicacao(driver)
            data_upload = processar_aba_anexos(driver)

            # Adicionando os dados na lista de saída
            dados_saida.append([
                instrumento, data_ajustes, data_ta, status_registro, data_upload
            ])

            # 📂 Criar ou atualizar a planilha Excel
            df_saida = pd.DataFrame(dados_saida, columns=[
                "Instrumento", "Data Ajustes", "Data TA", "Rendimento Aplicação", "Aba Anexos"
            ])
            salvar_planilha(df_saida)

            # 🔔 Criar a planilha assim que o primeiro instrumento for processado
            if not planilha_criada:
                print("📂 Criando planilha de controle inicial...")
                planilha_criada = True

            print("📂 Planilha atualizada com os dados coletados.")

            # 🔄 **Voltar para pesquisar um novo instrumento**
            print("↩️ Voltando para pesquisa de novo instrumento...")
            botao_voltar = esperar_elemento(driver, "/html/body/div[3]/div[2]/div[1]/a")
            if botao_voltar:
                botao_voltar.click()
                time.sleep(2)  # Pequena pausa para evitar problemas de carregamento

        except Exception as e:
            print(f"❌ Erro ao processar o instrumento {instrumento}: {e}")
            continue  # Continua para o próximo instrumento mesmo em caso de erro

    print("✅ Processamento concluído! Planilha final salva com sucesso.")

executar_robo()
