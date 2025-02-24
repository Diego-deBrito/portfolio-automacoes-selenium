import re
import time
import pandas as pd
import os
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from openpyxl import load_workbook

# 📌 Caminhos dos arquivos
CAMINHO_ENTRADA = r"C:\Users\diego.brito\Downloads\robov1\CONTROLE DE PARCERIAS CGAP.xlsx"
CAMINHO_SAIDA = r"C:\Users\diego.brito\Downloads\robov1\saida_Anexos.xlsx"

# 🛠 Conectar ao navegador já aberto
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

# 📥 Ler os números da coluna "Instrumento nº" na aba "PARCERIAS CGAP"
import pandas as pd

def obter_dados_propostas():
    """Lê os dados da planilha e filtra apenas os instrumentos com Status 'ATIVOS TODOS', ignorando campos vazios."""
    try:
        df = pd.read_excel(CAMINHO_ENTRADA, sheet_name="PARCERIAS CGAP",
                           usecols=["Instrumento nº", "Técnico", "e-mail do Técnico", "Status"])

        df.columns = df.columns.str.strip()  # Remove espaços dos nomes das colunas

        # 🔹 Remover completamente as linhas onde "Instrumento nº" está vazio ou é NaN
        df = df.dropna(subset=["Instrumento nº"])

        # 🔹 Converter para string, remover espaços extras e garantir formato correto
        df["Instrumento nº"] = df["Instrumento nº"].astype(str).str.replace(r"\.0$", "", regex=True).str.strip()

        # 🔹 Filtrar linhas onde "Instrumento nº" ainda pode estar vazio após limpeza
        df = df[df["Instrumento nº"] != ""]

        # 🔹 Preencher valores nulos em outras colunas
        df["Técnico"] = df["Técnico"].fillna("Desconhecido").astype(str).str.strip()
        df["e-mail do Técnico"] = df["e-mail do Técnico"].fillna("Sem e-mail").astype(str).str.strip()
        df["Status"] = df["Status"].fillna("").astype(str).str.strip().str.upper()

        # 🔹 Filtrar apenas os instrumentos ativos
        df_filtrado = df[df["Status"] == "ATIVOS TODOS"]

        if df_filtrado.empty:
            print("⚠️ Nenhum instrumento ativo encontrado na planilha!")
            return pd.DataFrame()

        return df_filtrado

    except Exception as e:
        print(f"❌ Erro ao ler a planilha: {e}")
        return pd.DataFrame()




# 🔍 Função para buscar a data mais recente ou registrar "Nenhum anexo encontrado"
def encontrar_data_mais_recente(driver, tabela_xpath):
    """Busca a data mais recente dentro da tabela especificada. Se não encontrar anexos, registra 'Nenhum anexo encontrado'."""
    wait = WebDriverWait(driver, 5)
    try:
        if verificar_ausencia_de_anexos(driver):
            return "Nenhum anexo encontrado"

        wait.until(EC.presence_of_element_located((By.XPATH, tabela_xpath)))
        elementos_datas = driver.find_elements(By.XPATH, f"{tabela_xpath}/tbody/tr/td[3]")  # Coluna 3 = Data

        datas = []
        for elemento in elementos_datas:
            try:
                data_texto = elemento.text.strip()
                data_formatada = datetime.strptime(data_texto, "%d/%m/%Y")
                datas.append(data_formatada)
            except ValueError:
                continue

        return max(datas).strftime('%d/%m/%Y') if datas else "Nenhum anexo encontrado"
    except Exception:
        return "Nenhum anexo encontrado"


def obter_dados_tecnico(numero_instrumento):
    """Retorna o Técnico e o E-mail do Técnico apenas para instrumentos ativos."""
    try:
        df = pd.read_excel(CAMINHO_ENTRADA, sheet_name="PARCERIAS CGAP",
                           usecols=["Instrumento nº", "Técnico", "e-mail do Técnico", "Status"])

        # 🔍 Remover espaços dos nomes das colunas
        df.columns = df.columns.str.strip()

        # 🔍 Garantir que "Instrumento nº" é string sem ".0"
        df["Instrumento nº"] = df["Instrumento nº"].fillna("").astype(str).str.replace(r"\.0$", "", regex=True).str.strip()
        df["Técnico"] = df["Técnico"].fillna("Desconhecido").astype(str).str.strip()
        df["e-mail do Técnico"] = df["e-mail do Técnico"].fillna("Sem e-mail").astype(str).str.strip()
        df["Status"] = df["Status"].fillna("").astype(str).str.strip().str.upper()

        # 📌 Filtrar apenas os instrumentos ativos e que correspondem ao número pesquisado
        filtro = (df["Instrumento nº"] == str(numero_instrumento)) & (df["Status"] == "ATIVOS TODOS")
        dados = df[filtro]

        if not dados.empty:
            # 📌 Se houver mais de um técnico para o mesmo instrumento, junta os valores com "; "
            tecnico = "; ".join(dados["Técnico"].unique())
            email_tecnico = "; ".join(dados["e-mail do Técnico"].unique())
            return tecnico, email_tecnico
        else:
            return "Desconhecido", "Sem e-mail"

    except Exception as e:
        print(f"❌ Erro ao ler a planilha: {e}")
        return "Erro", "Erro"



from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from datetime import datetime

def processar_proposta(driver, numero_instrumento):
    """Executa a automação para extrair as datas da proposta e execução."""
    wait = WebDriverWait(driver, 10)

    try:
        print(f"🔎 Buscando Instrumento {numero_instrumento}...")

        # 📌 Acessar a pesquisa de propostas
        wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[3]/div[1]/div[1]/div[1]/div[4]"))).click()
        wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[3]/div[2]/div[1]/div[1]/ul/li[6]/a"))).click()

        # 📌 Inserir número do instrumento e submeter
        campo_instrumento = wait.until(EC.visibility_of_element_located(
            (By.XPATH, "/html/body/div[3]/div[15]/div[3]/div/div/form/table/tbody/tr[2]/td[2]/input")))
        campo_instrumento.clear()
        campo_instrumento.send_keys(numero_instrumento)

        wait.until(EC.element_to_be_clickable(
            (By.XPATH, "/html/body/div[3]/div[15]/div[3]/div/div/form/table/tbody/tr[2]/td[2]/span/input"))).click()

        # 📌 Clicar no link do instrumento
        wait.until(EC.element_to_be_clickable(
            (By.XPATH, "/html/body/div[3]/div[15]/div[3]/div[3]/table/tbody/tr/td/div/a"))).click()

        # 📌 Verificar se a página contém apenas o botão "Listar Anexos Proposta"
        try:
            if driver.find_element(By.XPATH, "//button[contains(text(), 'Listar Anexos Proposta')]").is_displayed():
                print(f"⚠️ Nenhum anexo disponível para {numero_instrumento}. Registrando e continuando o loop...")
                return "Nenhuma aba de anexos disponível", "Nenhuma aba de anexos disponível"
        except NoSuchElementException:
            pass  # Se o botão não estiver presente, continua normalmente

        # 📌 Acessar aba de anexos proposta
        try:
            wait.until(EC.element_to_be_clickable(
                (By.XPATH, "/html/body/div[3]/div[15]/div[1]/div/div[1]/a[2]/div/span/span"))).click()
            wait.until(EC.element_to_be_clickable(
                (By.XPATH, "/html/body/div[3]/div[15]/div[1]/div/div[2]/a[8]/div/span/span"))).click()
        except (TimeoutException, NoSuchElementException):
            print(f"⚠️ Erro ao acessar aba de anexos proposta para {numero_instrumento}. Registrando erro e seguindo...")
            return "Nenhuma aba de anexos disponível", "Nenhuma aba de anexos disponível"

        # 📌 Acessar aba de anexos execução
        try:
            wait.until(EC.element_to_be_clickable(
                (By.XPATH, "/html/body/div[3]/div[15]/div[3]/div[1]/div/form/table/tbody/tr/td[2]/input[1]"))).click()
        except (TimeoutException, NoSuchElementException):
            print(f"⚠️ Erro ao acessar aba de anexos execução para {numero_instrumento}. Registrando erro e voltando ao loop...")
            return "Nenhuma aba de anexos disponível", "Nenhuma aba de anexos disponível"

        # 📌 Capturar a data mais recente na aba "Anexos Proposta"
        try:
            data_anexos_proposta = encontrar_data_mais_recente(
                driver, "/html/body/div[3]/div[15]/div[4]/div/div[1]/form/div/div[1]/table"
            )
        except (TimeoutException, NoSuchElementException):
            print(f"⚠️ Erro ao capturar anexos da proposta para {numero_instrumento}.")
            data_anexos_proposta = "Elemento não encontrado"

        # 📌 Voltar para a tela principal
        wait.until(EC.element_to_be_clickable(
            (By.XPATH, "/html/body/div[3]/div[15]/div[4]/div/div[1]/form/table/tbody/tr[1]/td/input"))).click()

        # 📌 Acessar aba de anexos execução (segundo botão)
        try:
            wait.until(EC.element_to_be_clickable(
                (By.XPATH, "/html/body/div[3]/div[15]/div[3]/div[1]/div/form/table/tbody/tr/td[2]/input[2]"))).click()
        except (TimeoutException, NoSuchElementException):
            print(f"⚠️ Erro ao acessar novamente a aba de anexos execução para {numero_instrumento}. Registrando erro e voltando ao loop...")
            return data_anexos_proposta, "Nenhuma aba de anexos disponível"

        # 📌 Capturar a data mais recente na aba "Anexos Execução"
        try:
            data_anexos_execucao = encontrar_data_mais_recente(
                driver, "/html/body/div[3]/div[15]/div[4]/div/div[1]/form/div/div[1]/table"
            )
        except (TimeoutException, NoSuchElementException):
            print(f"⚠️ Erro ao capturar anexos da execução para {numero_instrumento}.")
            data_anexos_execucao = "Elemento não encontrado"

        # 📌 Voltar ao loop
        wait.until(EC.element_to_be_clickable(
            (By.XPATH, "/html/body/div[3]/div[2]/div[1]/a"))).click()

        print(f"✅ Concluído: {numero_instrumento} - Proposta: {data_anexos_proposta} | Execução: {data_anexos_execucao}")
        return data_anexos_proposta, data_anexos_execucao

    except Exception as e:
        print(f"❌ Erro inesperado ao processar {numero_instrumento}: {e}")
        return "Erro inesperado", "Erro inesperado"


# 🛠 Verificar se não há anexos
def verificar_ausencia_de_anexos(driver):
    """Verifica se a mensagem 'Nenhum registro foi encontrado.' está na página."""
    try:
        mensagem_xpath = "//div[contains(text(), 'Nenhum registro foi encontrado.')]"
        return driver.find_element(By.XPATH, mensagem_xpath).is_displayed()
    except:
        return False



# 🚀 Função para salvar os dados no Excel sem sobrescrever o conteúdo existente
def salvar_dados(df_saida):
    if os.path.exists(CAMINHO_SAIDA):
        with pd.ExcelWriter(CAMINHO_SAIDA, engine="openpyxl", mode="a", if_sheet_exists="overlay") as writer:
            df_saida.to_excel(writer, index=False, sheet_name="Dados", header=False, startrow=writer.sheets["Dados"].max_row)
    else:
        df_saida.to_excel(CAMINHO_SAIDA, index=False, sheet_name="Dados")

# 🚀 Executar automação
def executar_automacao():
    driver = conectar_navegador_existente()
    df_propostas = obter_dados_propostas()

    if df_propostas.empty:
        print("⚠️ Nenhum instrumento encontrado na aba 'PARCERIAS CGAP'.")
        return

    for _, row in df_propostas.iterrows():
        numero_instrumento = row["Instrumento nº"]
        tecnico = row["Técnico"]
        email_tecnico = row["e-mail do Técnico"]

        print(f"🔎 Processando Instrumento {numero_instrumento}...")

        # 📌 Buscar datas dos anexos
        data_proposta, data_execucao = processar_proposta(driver, numero_instrumento)

        # Criar DataFrame para salvar no Excel
        df_saida = pd.DataFrame([{
            "Instrumento nº": numero_instrumento,
            "Técnico": tecnico,
            "e-mail do Técnico": email_tecnico,
            "Data Anexos Proposta": data_proposta,
            "Data Anexos Execução": data_execucao
        }])

        salvar_dados(df_saida)

    driver.quit()

# 🔥 Rodar a automação
executar_automacao()