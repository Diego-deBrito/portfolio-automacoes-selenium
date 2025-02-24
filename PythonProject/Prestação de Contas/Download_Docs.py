import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, ElementNotInteractableException, NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager
import time


# Função para conectar ao navegador já aberto
def conectar_navegador_existente():
    print("Passo 1: Conectando ao navegador existente na porta de depuração 9222...")
    options = webdriver.ChromeOptions()
    options.debugger_address = "localhost:9222"  # Porta que o Chrome está utilizando para depuração
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    print("Conexão ao navegador estabelecida.")
    return driver


# Função para clicar no primeiro elemento especificado pelo XPath (clicar em Execução)
def clicar_execucao(driver):
    print("Passo 2: Tentando clicar no elemento 'Execução'...")
    try:
        elemento = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[3]/div[1]/div[1]/div[1]/div[4]'))
        )
        elemento.click()
        print("Clique no elemento 'Execução' realizado com sucesso.")
    except TimeoutException:
        print("Erro: O elemento 'Execução' não foi encontrado ou não estava clicável.")


# Função para clicar no elemento 'Consultar Instrumentos/Pré-Instrumentos'
def clicar_consultar_instrumentos(driver):
    print("Passo 3: Tentando clicar no elemento 'Consultar Instrumentos/Pré-Instrumentos'...")
    try:
        elemento = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[3]/div[2]/div[1]/div[1]/ul/li[6]/a'))
        )
        elemento.click()
        print("Clique no elemento 'Consultar Instrumentos/Pré-Instrumentos' realizado com sucesso.")
    except TimeoutException:
        print("Erro: O elemento 'Consultar Instrumentos/Pré-Instrumentos' não foi encontrado ou não estava clicável.")


# Função para inserir o número do instrumento e realizar cliques subsequentes
def inserir_codigo_e_realizar_cliques(driver, codigo_instrumento):
    print(f"Pesquisando Pré-Instrumento {codigo_instrumento}...")
    try:
        # Localiza o campo de entrada e insere o número do instrumento
        campo_input = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(
                (By.XPATH, '/html/body/div[3]/div[15]/div[3]/div/div/form/table/tbody/tr[2]/td[2]/input'))
        )
        driver.execute_script("arguments[0].scrollIntoView();", campo_input)  # Rolagem até o campo, se necessário
        campo_input.clear()  # Limpa o campo antes de inserir o novo valor
        campo_input.send_keys(codigo_instrumento)  # Insere o código do pré-instrumento
        print(f"Código '{codigo_instrumento}' inserido no campo de pré-instrumento com sucesso.")

        # Clica no botão de consulta
        botao_consultar = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(
                (By.XPATH, '/html/body/div[3]/div[15]/div[3]/div/div/form/table/tbody/tr[2]/td[2]/span/input'))
        )
        botao_consultar.click()
        print("Botão de consulta clicado com sucesso.")

        # Clica no número do instrumento
        numero_instrumento = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[3]/div[15]/div[3]/div[3]/table/tbody/tr/td/div/a'))
        )
        numero_instrumento.click()
        print("Clicou no número do instrumento com sucesso.")

        # Clica em 'Execução Concedente' com tentativa de dois XPaths
        try:
            execucao_concedente = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '/html/body/div[3]/div[15]/div[1]/div/div[1]/a[6]/div/span/span'))
            )
            execucao_concedente.click()
            print("Clicou em 'Execução Concedente' usando o primeiro XPath com sucesso.")
            proximo_xpath = '/html/body/div[3]/div[15]/div[1]/div/div[2]/a[25]/div/span/span'
        except TimeoutException:
            print("Primeiro XPath para 'Execução Concedente' falhou. Tentando o segundo XPath...")
            execucao_concedente = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '/html/body/div[3]/div[15]/div[1]/div/div[1]/a[5]/div/span/span'))
            )
            execucao_concedente.click()
            print("Clicou em 'Execução Concedente' usando o segundo XPath com sucesso.")
            proximo_xpath = '/html/body/div[3]/div[15]/div[1]/div/div[2]/a[24]/div/span/span'

        # Clica em 'Documento de Liquidação'
        documento_liquidacao = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, proximo_xpath))
        )
        documento_liquidacao.click()
        print("Clicou em 'Documento de Liquidação' com sucesso.")

        # Processa as linhas na página atual
        processar_linhas(driver)

        # Navega para as próximas páginas até o fim
        navegar_e_processar_paginas(driver, codigo_instrumento)

        return 20  # Assumimos 20 linhas por página, a ser ajustado conforme necessário

    except TimeoutException:
        print("Erro: Um dos elementos não foi encontrado ou não estava clicável.")
        return 0
    except ElementNotInteractableException:
        print("Erro: Não foi possível interagir com um dos elementos.")
        return 0


# Função para processar até 20 linhas em cada página
def processar_linhas(driver):
    for i in range(1, 21):  # Até 20 linhas por página
        try:
            # Gera o XPath para cada linha
            xpath_linha = f'/html/body/div[3]/div[15]/div[4]/div[3]/table/tbody/tr[{i}]/td[1]/a'
            linha_elemento = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, xpath_linha))
            )
            linha_elemento.click()
            print(f"Clicou na linha {i}.")

            # Clica no botão de baixar
            try:
                botao_baixar = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH,
                                                '/html/body/div[3]/div[15]/div[4]/div[1]/div/form/table/tbody/tr[24]/td/div[1]/table/tbody/tr/td[3]/nobr/a'))
                )
                botao_baixar.click()
                print(f"Clicou no botão de baixar para a linha {i}.")
            except TimeoutException:
                print(f"Erro: O botão de baixar não foi encontrado ou não estava clicável para a linha {i}.")

            # Volta para a página anterior
            driver.back()
            print(f"Voltou para a página anterior após clicar na linha {i}.")
            time.sleep(2)  # Aguarda o carregamento

        except TimeoutException:
            print(f"Erro ao tentar clicar na linha {i}. Prosseguindo para a próxima linha.")
        except NoSuchElementException:
            print(f"Elemento na linha {i} não encontrado. Prosseguindo para a próxima linha.")
        except ElementNotInteractableException:
            print(f"Elemento na linha {i} não pode ser interagido. Prosseguindo para a próxima linha.")


# Função para navegar para a próxima página e reiniciar a consulta
def navegar_e_processar_paginas(driver, codigo_instrumento):
    pagina_atual = 1
    while True:
        try:
            pagina_atual += 1
            print(f"Tentando navegar para a página {pagina_atual}...")

            # Clica no link para a próxima página
            link_pagina = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.LINK_TEXT, str(pagina_atual)))
            )
            link_pagina.click()
            time.sleep(2)  # Aguarda a página carregar

            # Processa as linhas da nova página
            processar_linhas(driver)

            # Retorna para a etapa inicial de consulta
            print("Retornando à etapa inicial para nova consulta...")
            elemento_voltar = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '/html/body/div[3]/div[2]/div[6]/a[2]'))
            )
            elemento_voltar.click()
            clicar_consultar_instrumentos(driver)  # Retoma a consulta desde o início
            inserir_codigo_e_realizar_cliques(driver, codigo_instrumento)  # Reinicia a pesquisa com o mesmo instrumento

        except TimeoutException:
            print(
                f"Erro ao tentar acessar a página {pagina_atual}. Presumivelmente todas as páginas foram processadas.")
            break


# Função principal para carregar a planilha e processar os instrumentos
def processar_instrumentos():
    caminho_planilha = r"D:\Users\andrei.rodrigues\OneDrive - Ministério do Desenvolvimento e Assistência Social\Power BI\Python\Consulta Prestação de contas\Instrumentos Iniciais.xlsx"
    print(f"Carregando a planilha '{caminho_planilha}'...")

    try:
        df_instrumentos = pd.read_excel(caminho_planilha, usecols=["Instrumento"])
        df_instrumentos["Linhas"] = 0  # Cria a coluna "Linhas" com valor inicial 0
        print("Planilha carregada com sucesso.")
    except FileNotFoundError:
        print("Erro: Arquivo não encontrado. Verifique o caminho e o nome do arquivo.")
        return
    except KeyError:
        print("Erro: A coluna 'Instrumento' não foi encontrada na planilha.")
        return

    driver = conectar_navegador_existente()

    clicar_execucao(driver)
    clicar_consultar_instrumentos(driver)

    for index, row in df_instrumentos.iterrows():
        codigo_instrumento = str(row['Instrumento']).strip()
        print(f"\nProcessando o Pré-Instrumento {codigo_instrumento}")
        num_linhas = inserir_codigo_e_realizar_cliques(driver, codigo_instrumento)
        df_instrumentos.at[index, "Linhas"] = num_linhas

    caminho_saida = r"D:\Users\andrei.rodrigues\OneDrive - Ministério do Desenvolvimento e Assistência Social\Power BI\Python\Consulta Prestação de contas\Instrumentos Iniciais Atualizado.xlsx"
    df_instrumentos.to_excel(caminho_saida, index=False)
    print(f"Planilha atualizada salva em '{caminho_saida}'.")

    print("Processamento dos instrumentos concluído.")
    driver.quit()


# Executa o processamento dos instrumentos
processar_instrumentos()
