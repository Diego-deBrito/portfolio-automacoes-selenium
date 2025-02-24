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

# 📌 Caminhos dos arquivos
CAMINHO_ENTRADA = r"C:\Users\diego.brito\Downloads\robov1\CONTROLE DE PARCERIAS CGAP.xlsx"
CAMINHO_SAIDA = r"C:\Users\diego.brito\Downloads\robov1\saida.xlsx"

# 📥 Ler os números da coluna "Instrumento nº" (coluna F)
def obter_numeros_instrumento():
    """Lê os números de instrumento da coluna 'Instrumento nº' da planilha de entrada, filtrando apenas os ativos."""
    try:
        df = pd.read_excel(CAMINHO_ENTRADA, usecols=["Instrumento nº", "Status"])

        # Converter "Instrumento nº" para string e remover NaN
        df = df.dropna(subset=["Instrumento nº", "Status"])
        df["Instrumento nº"] = df["Instrumento nº"].astype(int).astype(str)

        # Filtrar apenas os que possuem "ATIVOS TODOS" na coluna "Status"
        df_filtrado = df[df["Status"].str.strip().str.upper() == "ATIVOS TODOS"]

        return df_filtrado["Instrumento nº"].tolist()  # Retorna apenas os instrumentos ativos
    except Exception as e:
        print(f"❌ Erro ao ler a planilha: {e}")
        return []


def obter_total_paginas(driver):
    """Identifica quantas páginas e itens existem na tabela de esclarecimentos."""
    try:
        info_elemento = driver.find_element(By.XPATH, "/html/body/div[3]/div[15]/div[4]/div[1]/div/form/table/tbody/tr[6]/td/div[1]/span[1]")
        texto_info = info_elemento.text.strip()

        # Exemplo de texto: "Página 1 de 4 (69 item(s))"
        match = re.search(r'Página \d+ de (\d+) \((\d+) item', texto_info)
        if match:
            total_paginas = int(match.group(1))
            total_itens = int(match.group(2))
            return total_paginas, total_itens
        else:
            return 1, 0  # Se não encontrar, assume que há apenas 1 página e nenhum item

    except Exception as e:
        print(f"❌ Erro ao obter total de páginas: {e}")
        return 1, 0  # Se houver erro, assume 1 página e 0 itens











# 📥 Função para obter os dados do Técnico e e-mail do Técnico
def obter_dados_tecnico(numero_instrumento):
    """Retorna o Técnico e o E-mail do Técnico para um instrumento ativo."""
    try:
        df = pd.read_excel(CAMINHO_ENTRADA, sheet_name="PARCERIAS CGAP",
                           usecols=["Instrumento nº", "Técnico", "Status", "e-mail do Técnico"])

        # 🔍 Limpar espaços dos nomes das colunas
        df.columns = df.columns.str.strip()

        # 🔍 Converter "Instrumento nº" para string e remover ".0"
        df["Instrumento nº"] = df["Instrumento nº"].fillna("").astype(str).str.replace(r"\.0$", "", regex=True).str.strip()

        # 🔍 Tratar valores nulos
        df["Técnico"] = df["Técnico"].fillna("Desconhecido").astype(str).str.strip()
        df["e-mail do Técnico"] = df["e-mail do Técnico"].fillna("Sem e-mail").astype(str).str.strip()
        df["Status"] = df["Status"].fillna("").astype(str).str.strip().str.upper()

        # 📌 Garantir que estamos comparando strings formatadas corretamente
        numero_instrumento_str = str(numero_instrumento).strip()

        # 🔍 Filtrar apenas instrumentos ativos e que correspondem ao número pesquisado
        filtro = (df["Instrumento nº"] == numero_instrumento_str) & (df["Status"] == "ATIVOS TODOS")
        dados = df[filtro]

        if not dados.empty:
            # 📌 Se houver mais de um técnico, concatena os valores com "; "
            tecnico = "; ".join(dados["Técnico"].unique())
            email_tecnico = "; ".join(dados["e-mail do Técnico"].unique())
            return tecnico, email_tecnico
        else:
            return "Desconhecido", "Sem e-mail"

    except Exception as e:
        print(f"❌ Erro ao ler a planilha: {e}")
        return "Erro", "Erro"

# 🛠 Função principal de automação
def encontrar_dados_esclarecimento(driver, numero_instrumento):
    """Percorre todas as páginas da tabela e retorna as datas da situação 'Resposta Enviada'."""
    wait = WebDriverWait(driver, 10)

    # 📌 Navegação até a seção de esclarecimentos
    menu_button = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[3]/div[1]/div[1]/div[1]/div[4]")))
    menu_button.click()

    menu_option = wait.until(EC.element_to_be_clickable((By.XPATH, "//html/body/div[1]/div[3]/div[2]/div[1]/div[1]/ul/li[6]/a")))
    menu_option.click()

    campo_convenio = wait.until(EC.visibility_of_element_located((By.XPATH, "//input[@id='consultarNumeroConvenio']")))
    campo_convenio.clear()
    campo_convenio.send_keys(numero_instrumento)

    botao_submit = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[@class='FormLinhaBotoes']//input[@id='form_submit']")))
    botao_submit.click()

    convenio_link = wait.until(EC.element_to_be_clickable((By.XPATH, f"//a[normalize-space()='{numero_instrumento}']")))
    convenio_link.click()

    menu_button = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[3]/div[2]/div[4]/div/div[7]")))
    menu_button.click()

    esclarecimentos_link = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[@id='contentMenuInterno']//a[normalize-space()='Esclarecimentos']")))
    esclarecimentos_link.click()

    # 📌 Identificar total de páginas antes de começar
    total_paginas, total_itens = obter_total_paginas(driver)
    print(f"🔍 Tabela possui {total_paginas} páginas e {total_itens} itens.")

    # 📌 Caminho para elementos da tabela
    tabela_xpath = "/html/body/div[3]/div[15]/div[4]/div[1]/div/form/table/tbody/tr[6]/td/div[1]/table"
    botoes_paginacao_xpath = "/html/body/div[3]/div[15]/div[4]/div[1]/div/form/table/tbody/tr[6]/td/div[1]/span[2]"

    datas_resposta = []  # Lista para armazenar todas as datas encontradas vinculadas à "Resposta Enviada"

    for pagina in range(1, total_paginas + 1):
        print(f"📄 Processando página {pagina} de {total_paginas}...")

        # 🔍 Pegando todas as linhas da tabela
        linhas = driver.find_elements(By.XPATH, f"{tabela_xpath}/tbody/tr")

        for linha in linhas:
            try:
                situacao = linha.find_element(By.XPATH, "./td[6]").text.strip()  # Coluna 6 = Situação
                data_solicitacao = linha.find_element(By.XPATH, "./td[1]").text.strip()  # Coluna 1 = Data de Solicitação

                if situacao == "Resposta Enviada":
                    datas_resposta.append(data_solicitacao)  # Somente essa data será salva
            except:
                continue

        # 📌 Avançar para a próxima página, se houver mais páginas disponíveis
        if pagina < total_paginas:
            try:
                proxima_pagina_botao = driver.find_element(By.XPATH, f"{botoes_paginacao_xpath}/a[text()='{pagina + 1}']")
                proxima_pagina_botao.click()
                time.sleep(2)  # Espera um pouco para carregar a nova página
            except Exception as e:
                print(f"⚠️ Erro ao tentar avançar para a página {pagina + 1}: {e}")
                break  # Se não conseguir avançar, interrompe o loop

    return datas_resposta if datas_resposta else ["Sem dados"]

# 🚀 Executar automação para todos os números da planilha

def executar_automacao():
    driver = conectar_navegador_existente()
    numeros_instrumento = obter_numeros_instrumento()

    if not numeros_instrumento:
        print("⚠️ Nenhum instrumento ativo encontrado na planilha.")
        return

    # Criar a planilha se não existir
    if not os.path.exists(CAMINHO_SAIDA):
        df_vazio = pd.DataFrame(columns=["Instrumento nº", "Data de Resposta Enviada", "Técnico", "e-mail do Técnico"])
        df_vazio.to_excel(CAMINHO_SAIDA, index=False)

    for numero in numeros_instrumento:
        print(f"🔎 Buscando dados para instrumento {numero}...")

        # 📌 Buscar todas as datas de "Resposta Enviada"
        datas_resposta = encontrar_dados_esclarecimento(driver, numero)

        # 📌 Buscar Técnico e E-mail do Técnico
        tecnico, email_tecnico = obter_dados_tecnico(numero)

        # Filtrar para não registrar "Sem dados" se não houver respostas enviadas
        datas_resposta_filtradas = [data for data in datas_resposta if data != "Sem dados"]

        if datas_resposta_filtradas:
            print(f"✅ Datas de Resposta Enviada: {', '.join(datas_resposta_filtradas)}")
            print(f"✅ Técnico: {tecnico}, E-mail: {email_tecnico}")

            novos_dados = []
            for data in datas_resposta_filtradas:
                novos_dados.append({
                    "Instrumento nº": numero,
                    "Data de Resposta Enviada": data,
                    "Técnico": tecnico,
                    "e-mail do Técnico": email_tecnico
                })

            # 📌 Atualizar a planilha imediatamente
            df_existente = pd.read_excel(CAMINHO_SAIDA)  # Lê os dados já salvos
            df_novo = pd.DataFrame(novos_dados)  # Criar DataFrame com os novos dados
            df_final = pd.concat([df_existente, df_novo], ignore_index=True)  # Adicionar os novos dados
            df_final.to_excel(CAMINHO_SAIDA, index=False)  # Salvar novamente

            print(f"📂 Planilha atualizada para instrumento {numero}")
        else:
            print(f"⚠️ Nenhuma 'Resposta Enviada' encontrada para {numero}")

        # 📌 Voltar para a tela inicial para continuar o loop
        try:
            voltar_button = driver.find_element(By.XPATH, "/html/body/div[3]/div[2]/div[1]/a")
            voltar_button.click()
            time.sleep(2)
        except:
            print("❌ Erro ao voltar para a tela inicial.")
            driver.quit()
            return

    print(f"✅ Finalizado! Planilha salva em {CAMINHO_SAIDA}")
    driver.quit()

# 🔥 Rodar a automação
executar_automacao()
