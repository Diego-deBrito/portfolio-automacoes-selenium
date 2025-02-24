import time
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, ElementNotInteractableException
import pandas as pd
import os
from datetime import datetime


# 1. Função para conectar ao navegador já aberto
def conectar_navegador_existente():
    options = webdriver.ChromeOptions()
    options.debugger_address = "localhost:9222"  # Porta de depuração do Chrome
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    return driver


# 2. Função auxiliar para verificar se um elemento existe
def elemento_existe(driver, xpath, tempo_espera=2):
    try:
        WebDriverWait(driver, tempo_espera).until(EC.presence_of_element_located((By.XPATH, xpath)))
        return True
    except TimeoutException:
        return False


# 3. Função para cada script individual
def executar_requisitos(driver):
    print("Executando script: Consulta Requisitos...")

    # 3. Função para garantir que o elemento esteja visível e clicável antes de interagir


def clicar_elemento(driver, xpath):
    try:
        # Esperar o elemento estar visível e clicável
        elemento = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, xpath)))
        driver.execute_script("arguments[0].scrollIntoView();", elemento)  # Scroll até o elemento
        elemento.click()
        print(f"Elemento clicado com sucesso: {xpath}")
    except (TimeoutException, ElementNotInteractableException) as e:
        print(f"Erro ao clicar no elemento: {xpath}, Erro: {e}")
        raise e


# 4. Função para reiniciar a navegação a partir da proposta seguinte
def reiniciar_navegacao(driver, max_tentativas=3):
    for tentativa in range(max_tentativas):
        try:
            print(f"[INFO] Tentando reiniciar navegação (tentativa {tentativa + 1}/{max_tentativas})...")
            driver.refresh()
            WebDriverWait(driver, 1).until(
                EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[3]/div[1]/div[1]/div[1]/div[3]'))
            ).click()
            WebDriverWait(driver, 1).until(
                EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[3]/div[2]/div[1]/div[1]/ul/li[3]/a'))
            ).click()
            print("[INFO] Navegação reiniciada com sucesso.")
            return
        except TimeoutException:
            print(f"[ERRO] Falha ao reiniciar navegação na tentativa {tentativa + 1}.")
    print("[ERRO] Não foi possível reiniciar a navegação após várias tentativas.")


# 5. Função para buscar a data mais recente em uma tabela (Certidões, Declarações, etc.)
def buscar_data_mais_recente(driver, xpath_coluna_datas, secao):
    try:
        if elemento_existe(driver, xpath_coluna_datas, tempo_espera=0):  # Sem tempo de espera
            datas_upload = driver.find_elements(By.XPATH, xpath_coluna_datas)
            if datas_upload:
                lista_datas = []
                for data_element in datas_upload:
                    data_text = data_element.text
                    try:
                        data_convertida = datetime.strptime(data_text, "%d/%m/%Y %H:%M:%S")
                        lista_datas.append(data_convertida)
                    except Exception as e:
                        print(f"Erro ao converter data na seção '{secao}': {data_text}, Erro: {e}")
                if lista_datas:
                    data_mais_recente = max(lista_datas)
                    print(
                        f"Data mais recente na seção '{secao}': {data_mais_recente.strftime('%d/%m/%Y %H:%M:%S')}")
                    return data_mais_recente.strftime('%d/%m/%Y %H:%M:%S')
            print(f"Nenhuma data encontrada na seção '{secao}'.")
        else:
            print(f"Seção '{secao}' não localizada.")
        return None
    except Exception as e:
        print(f"Erro ao buscar a data mais recente na seção '{secao}': {e}")
        return None

    # 6. Função para buscar o valor do status, sem tentar converter para data


def buscar_status(driver, xpath_coluna_status, secao):
    try:
        if elemento_existe(driver, xpath_coluna_status, tempo_espera=0):  # Sem tempo de espera
            status_valor = driver.find_element(By.XPATH, xpath_coluna_status).text
            print(f"Status na seção '{secao}': {status_valor}")
            return status_valor
        else:
            print(f"Seção '{secao}' não localizada.")
        return None
    except Exception as e:
        print(f"Erro ao buscar o status na seção '{secao}': {e}")
        return None


# 7. Função para processar uma proposta
def processar_proposta(driver, proposta_numero):
    print(f"Iniciando o processamento da proposta: {proposta_numero}")

    # Digitar o número da proposta e consultar
    driver.find_element(By.XPATH,
                        '/html/body/div[3]/div[15]/div[3]/div/div/form/table/tbody/tr[1]/td[2]/input').clear()
    driver.find_element(By.XPATH,
                        '/html/body/div[3]/div[15]/div[3]/div/div/form/table/tbody/tr[1]/td[2]/input').send_keys(
        proposta_numero)

    if elemento_existe(driver, '/html/body/div[3]/div[15]/div[3]/div/div/form/table/tbody/tr[1]/td[2]/span/input',
                       tempo_espera=2):
        clicar_elemento(driver, '/html/body/div[3]/div[15]/div[3]/div/div/form/table/tbody/tr[1]/td[2]/span/input')

    # Clicar no número da proposta
    if elemento_existe(driver, '/html/body/div[3]/div[15]/div[3]/div[3]/table/tbody/tr/td[1]/div/a',
                       tempo_espera=2):
        clicar_elemento(driver, '/html/body/div[3]/div[15]/div[3]/div[3]/table/tbody/tr/td[1]/div/a')

    # Clicar na aba "Requisitos"
    if elemento_existe(driver, '/html/body/div[3]/div[15]/div[1]/div/div[1]/a[3]/div/span/span', tempo_espera=2):
        clicar_elemento(driver, '/html/body/div[3]/div[15]/div[1]/div/div[1]/a[3]/div/span/span')
        print("Aba 'Requisitos' clicada com sucesso.")
    else:
        print("Aba 'Requisitos' não localizada.")
        return False

    # Tentar clicar em "Requisitos para celebração" com cinco XPaths
    xpaths = [
        '/html/body/div[3]/div[15]/div[1]/div/div[2]/a[12]/div/span/span',  # Primeiro XPath
        '/html/body/div[3]/div[15]/div[1]/div/div[1]/a[3]/div/span/span',  # Segundo XPath
        '/html/body/div[3]/div[16]/div[1]/div/div[2]/a[29]/div/span/span',  # Terceiro XPath
        '/html/body/div[3]/div[16]/div[1]/div/div[2]/a[13]/div/span/span',  # Quarto XPath
        '/html/body/div[3]/div[15]/div[1]/div/div[2]/a[13]/div/span/span'  # Quinto XPath
    ]

    clicado = False
    for xpath in xpaths:
        if elemento_existe(driver, xpath, tempo_espera=1):
            try:
                clicar_elemento(driver, xpath)
                clicado = True
                print(f"Clicado em 'Requisitos para celebração' usando XPath: {xpath}")
                break
            except Exception as e:
                print(
                    f"Erro ao tentar clicar em 'Requisitos para celebração' usando XPath: {xpath}. Tentando o próximo...")

    if not clicado:
        print("Requisitos para celebração não localizado.")
        return False

    # Consultar valores nas seções, agora com pulos rápidos em caso de não localização
    dados_proposta = {
        "Proposta": proposta_numero,
        "Certidões": buscar_data_mais_recente(driver,
                                              "/html/body/div[3]/div[16]/div[2]/div[2]/form/div[1]/div[1]/table/tbody/tr[1]/td[2]",
                                              "Certidões"),
        "Declarações": buscar_data_mais_recente(driver,
                                                "/html/body/div[3]/div[16]/div[2]/div[2]/form/div[1]/div[2]/table/tbody/tr[1]/td[2]",
                                                "Declarações"),
        "Comprovantes de Execução": buscar_data_mais_recente(driver,
                                                             "/html/body/div[3]/div[16]/div[2]/div[2]/form/div[1]/div[3]/table/tbody/tr[1]/td[2]",
                                                             "Comprovantes de Execução"),
        "Outros": buscar_data_mais_recente(driver,
                                           "/html/body/div[3]/div[16]/div[2]/div[2]/form/div[1]/div[4]/table/tbody/tr[1]/td[2]",
                                           "Outros"),
        "Históricos - Data": buscar_data_mais_recente(driver,
                                                      "/html/body/div[3]/div[16]/div[2]/div[2]/form/div[1]/div[5]/table/tbody/tr[1]/td[3]",
                                                      "Históricos"),
        "Históricos - Status": buscar_status(driver,
                                             "/html/body/div[3]/div[16]/div[2]/div[2]/form/div[1]/div[5]/table/tbody/tr[1]/td[1]",
                                             "Históricos Status")
    }

    return dados_proposta


# 8. Função para salvar o progresso no Excel, agora limpando a planilha no início
def salvar_progresso(resultado):
    caminho_arquivo = r"C:\Users\diego.brito\OneDrive - Ministério do Desenvolvimento e Assistência Social\Power BI\Python\Consulta Transferegov Requisitos.xlsx"

    try:
        if not os.path.exists(caminho_arquivo):
            # Se o arquivo não existir, criar pela primeira vez
            pd.DataFrame([resultado]).to_excel(caminho_arquivo, index=False)
        else:
            # Concatenar novos dados ao arquivo existente
            df_existente = pd.read_excel(caminho_arquivo)
            df_atualizado = pd.concat([df_existente, pd.DataFrame([resultado])], ignore_index=True)
            df_atualizado.to_excel(caminho_arquivo, index=False)
        print(f"[INFO] Progresso salvo no arquivo '{caminho_arquivo}'.")
    except PermissionError:
        print(f"[ERRO] O arquivo '{caminho_arquivo}' está aberto. Feche o arquivo e tente novamente.")
    except Exception as e:
        print(f"[ERRO] Falha ao salvar progresso no arquivo Excel: {e}")


# 9. Função para clicar no botão "Nova Pesquisa" com fallback para XPaths diferentes
def clicar_nova_pesquisa(driver):
    xpaths_nova_pesquisa = [
        '/html/body/div[3]/div[3]/div[6]/a[2]',  # Primeiro XPath
        '/html/body/div[3]/div[2]/div[6]/a[2]',  # Segundo XPath
        '/html/body[1]/div[3]/div[3]/div[6]/a[2]'  # Terceiro XPath
    ]

    botao_encontrado = False

    for xpath in xpaths_nova_pesquisa:
        if elemento_existe(driver, xpath, tempo_espera=2):
            try:
                WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, xpath))).click()
                print(f"Botão 'Nova Pesquisa' clicado com sucesso usando XPath: {xpath}")
                botao_encontrado = True
                break
            except (NoSuchElementException, TimeoutException):
                print(f"Falha ao clicar no botão 'Nova Pesquisa' com o XPath: {xpath}. Tentando o próximo...")
        else:
            print(f"Botão 'Nova Pesquisa' não encontrado com o XPath: {xpath}. Tentando o próximo...")

    # Se nenhum dos XPaths funcionar, recarregar a página e reiniciar a navegação
    if not botao_encontrado:
        print(f"Botão 'Nova Pesquisa' não encontrado após tentar todos os XPaths. Recarregando a página.")
        reiniciar_navegacao(driver)


# 10. Função principal para processar todas as propostas
def processar_todas_abas(driver, proposta_numero):
    """
    Processa as abas Requisitos, Parecer e Convênios para uma proposta específica.
    """
    try:
        print(f"\n[INFO] Processando proposta {proposta_numero}...")

        # Inserir o número da proposta e consultar
        input_xpath = '/html/body/div[3]/div[15]/div[3]/div/div/form/table/tbody/tr[1]/td[2]/input'
        consultar_xpath = '/html/body/div[3]/div[15]/div[3]/div/div/form/table/tbody/tr[1]/td[2]/span/input'
        proposta_xpath = '/html/body/div[3]/div[15]/div[3]/div[3]/table/tbody/tr/td[1]/div/a'

        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, input_xpath))).clear()
        driver.find_element(By.XPATH, input_xpath).send_keys(proposta_numero)
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, consultar_xpath))).click()

        # Clicar no número da proposta
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, proposta_xpath))).click()

        # Processar cada aba
        dados_completos = {}

        # 1. Aba Requisitos
        print("[INFO] Navegando para a aba 'Requisitos'...")
        aba_requisitos_xpath = '/html/body/div[3]/div[15]/div[1]/div/div[1]/a[3]/div/span/span'
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, aba_requisitos_xpath))).click()
        dados_completos["Requisitos"] = executar_requisitos(driver)  # Substitua por sua lógica de coleta

        # 2. Aba Parecer
        print("[INFO] Navegando para a aba 'Parecer'...")
        aba_parecer_xpath = '/html/body/div[3]/div[15]/div[1]/div/div[2]/a[10]/div/span/span'
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, aba_parecer_xpath))).click()
        dados_completos["Parecer"] = executar_parecer(driver)  # Substitua por sua lógica de coleta

        # 3. Aba Convênios
        print("[INFO] Navegando para a aba 'Convênios'...")
        aba_convenios_xpath = '/html/body/div[3]/div[15]/div[1]/div/div[1]/a[4]/div/span/span'
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, aba_convenios_xpath))).click()
        dados_completos["Convênios"] = executar_convenios(driver)  # Substitua por sua lógica de coleta

        print(f"[INFO] Dados coletados para proposta {proposta_numero}: {dados_completos}")
        return dados_completos

    except Exception as e:
        print(f"[ERRO] Falha ao processar a proposta {proposta_numero}: {e}")
        return {"Proposta": proposta_numero, "Erro": str(e)}


def executar_parecer(driver):
    print("Executando script: Consulta Parecer...")

    # Auxiliar para converter datas em diferentes formatos
    def converter_data(data_text):
        formatos = ["%d/%m/%Y %H:%M:%S", "%d/%m/%Y"]
        for formato in formatos:
            try:
                return datetime.strptime(data_text, formato)
            except ValueError:
                continue
        print(f"Erro ao converter data: {data_text}, formatos esperados: {formatos}")
        return None

    # Processar uma proposta


def processar_proposta(driver, proposta_numero, info_adicional):
    try:
        print(f"Processando proposta: {proposta_numero}")

        # Inserir número da proposta e consultar
        input_xpath = '/html/body/div[3]/div[15]/div[3]/div/div/form/table/tbody/tr[1]/td[2]/input'
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, input_xpath))).clear()
        driver.find_element(By.XPATH, input_xpath).send_keys(proposta_numero)

        consultar_xpath = '/html/body/div[3]/div[15]/div[3]/div/div/form/table/tbody/tr[1]/td[2]/span/input'
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, consultar_xpath))).click()

        # Clicar no número da proposta
        proposta_xpath = '/html/body/div[3]/div[15]/div[3]/div[3]/table/tbody/tr/td[1]/div/a'
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, proposta_xpath))).click()

        # Navegar para a aba Plano de Trabalho
        aba_plano_xpath = '/html/body/div[3]/div[15]/div[1]/div/div[1]/a[2]/div/span/span'
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, aba_plano_xpath))).click()

        # Navegar para a aba Pareceres
        try:
            parecer_xpath = '/html/body/div[3]/div[15]/div[1]/div/div[2]/a[10]/div/span/span'
            WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, parecer_xpath))).click()
        except TimeoutException:
            print("Falha ao acessar aba Pareceres.")

        # Consultar informações da proposta
        data_mais_recente_1 = None
        data_mais_recente_2 = None

        # Consultar primeira coluna
        coluna1_xpath = '/html/body/div[3]/div[15]/div[3]/div[2]/table/tbody/tr/td[1]'
        if elemento_existe(driver, coluna1_xpath):
            elementos = driver.find_elements(By.XPATH, coluna1_xpath)
            datas = [converter_data(e.text) for e in elementos if e.text.strip()]
            data_mais_recente_1 = max(datas).strftime('%d/%m/%Y %H:%M:%S') if datas else None

        # Consultar segunda coluna
        coluna2_xpath = '/html/body/div[3]/div[15]/div[3]/div[4]/div/div/form/div[1]/table/tbody/tr/td[1]'
        if elemento_existe(driver, coluna2_xpath):
            elementos = driver.find_elements(By.XPATH, coluna2_xpath)
            datas = [converter_data(e.text) for e in elementos if e.text.strip()]
            data_mais_recente_2 = max(datas).strftime('%d/%m/%Y %H:%M:%S') if datas else None

        return {
            "Proposta": proposta_numero,
            "Instrumento": info_adicional.get("Instrumento"),
            "AçãoOrçamentária": info_adicional.get("AçãoOrçamentária"),
            "OrigemRecurso": info_adicional.get("OrigemRecurso"),
            "CoordenaçãoResponsável": info_adicional.get("CoordenaçãoResponsável"),
            "Processo": info_adicional.get("Processo"),
            "TécnicoResponsável": info_adicional.get("TécnicoResponsável"),
            "DataMaisRecenteColuna1": data_mais_recente_1,
            "DataMaisRecenteColuna2": data_mais_recente_2
        }
    except Exception as e:
        print(f"Erro ao processar proposta {proposta_numero}: {e}")
        return {"Proposta": proposta_numero, "Erro": str(e)}


def executar_processo_completo():
    """
    Executa o processamento completo das abas Requisitos, Parecer e Convênios para todas as propostas.
    """
    caminho_planilha = r"C:\caminho\para\propostas.xlsx"  # Substitua pelo caminho correto
    caminho_resultados = r"C:\caminho\para\resultados.xlsx"  # Substitua pelo caminho correto

    # Carregar propostas
    if not os.path.exists(caminho_planilha):
        print(f"[ERRO] Planilha de propostas não encontrada: {caminho_planilha}")
        return

    propostas = pd.read_excel(caminho_planilha).fillna("")
    resultados = []

    # Inicializar o navegador
    driver = conectar_navegador_existente()

    try:
        for _, row in propostas.iterrows():
            proposta_numero = row["NºProposta"].strip()
            print(f"\n[INFO] Iniciando processamento da proposta {proposta_numero}...")

            # Processar as três abas para a proposta atual
            dados = processar_todas_abas(driver, proposta_numero)
            resultados.append(dados)

            # Salvar incrementalmente
            salvar_resultados(resultados, caminho_resultados)

            # Clicar no botão "Nova Pesquisa" para ir para a próxima proposta
            clicar_nova_pesquisa(driver)

    finally:
        if driver:
            driver.quit()

    print("[INFO] Processamento completo para todas as propostas.")


def clicar_nova_pesquisa(driver):
    """
    Clica no botão 'Nova Pesquisa' para retornar à tela inicial e processar a próxima proposta.
    """
    xpaths_nova_pesquisa = [
        '/html/body/div[3]/div[3]/div[6]/a[2]',  # Primeiro XPath
        '/html/body/div[3]/div[2]/div[6]/a[2]',  # Segundo XPath
        '/html/body[1]/div[3]/div[3]/div[6]/a[2]'  # Terceiro XPath
    ]

    for xpath in xpaths_nova_pesquisa:
        if elemento_existe(driver, xpath, tempo_espera=2):
            try:
                WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, xpath))).click()
                print("[INFO] Botão 'Nova Pesquisa' clicado com sucesso.")
                return
            except Exception as e:
                print(f"[ERRO] Falha ao clicar no botão 'Nova Pesquisa' com XPath {xpath}: {e}")

    print("[ERRO] Botão 'Nova Pesquisa' não encontrado. Recarregando a página...")
    driver.refresh()


# Salvar resultados em um arquivo Excel


def salvar_resultados(resultados, caminho_arquivo):
    """
    Salva os resultados no Excel incrementalmente.
    """
    try:
        df = pd.DataFrame(resultados)
        if os.path.exists(caminho_arquivo):
            df_existente = pd.read_excel(caminho_arquivo)
            df = pd.concat([df_existente, df], ignore_index=True)
        df.to_excel(caminho_arquivo, index=False)
        print(f"[INFO] Resultados salvos em {caminho_arquivo}")
    except Exception as e:
        print(f"[ERRO] Falha ao salvar resultados no Excel: {e}")

    # Caminhos de entrada e saída
    caminho_planilha = r"C:\Users\diego.brito\OneDrive - Ministério do Desenvolvimento e Assistência Social\Power BI\Python\Consulta Parecer\propostas_iniciais_Parecer.xlsx"
    caminho_resultados = r"C:\Users\diego.brito\OneDrive - Ministério do Desenvolvimento e Assistência Social\Power BI\Python\Consulta Parecer\Consulta_Parecer_Resultados.xlsx"

    # Carregar as propostas
    if not os.path.exists(caminho_planilha):
        print(f"Arquivo de entrada {caminho_planilha} não encontrado.")
        return

    propostas = pd.read_excel(caminho_planilha).fillna("")
    resultados = []

    # Processar cada proposta
    for _, row in propostas.iterrows():
        proposta_numero = row["NºProposta"].strip()
        info_adicional = {
            "Instrumento": row.get("Instrumento", ""),
            "AçãoOrçamentária": row.get("AçãoOrçamentária", ""),
            "OrigemRecurso": row.get("OrigemRecurso", ""),
            "CoordenaçãoResponsável": row.get("CoordenaçãoResponsável", ""),
            "Processo": row.get("Processo", ""),
            "TécnicoResponsável": row.get("TécnicoResponsável", "")
        }
        resultado = processar_proposta(driver, proposta_numero, info_adicional)
        resultados.append(resultado)

    # Salvar resultados
    salvar_resultados(resultados, caminho_resultados)
    print("Consulta Parecer concluída.")


def executar_convenios(driver):
    print("Executando script: Consulta Convênios...")

    # Auxiliar para converter datas em diferentes formatos


def converter_data(data_text):
    formatos = ["%d/%m/%Y %H:%M:%S", "%d/%m/%Y"]
    for formato in formatos:
        try:
            return datetime.strptime(data_text, formato)
        except ValueError:
            continue
    print(f"Erro ao converter data: {data_text}, formatos esperados: {formatos}")
    return None


# Salvar os resultados em um arquivo Excel
def salvar_resultados(resultados, caminho_arquivo):
    try:
        df = pd.DataFrame(resultados)
        if os.path.exists(caminho_arquivo):
            df_existente = pd.read_excel(caminho_arquivo)
            df = pd.concat([df_existente, df], ignore_index=True)
        df.to_excel(caminho_arquivo, index=False)
        print(f"Resultados salvos com sucesso em {caminho_arquivo}")
    except Exception as e:
        print(f"Erro ao salvar resultados: {e}")

    # Caminhos de entrada e saída


caminho_planilha = r"C:\Users\diego.brito\OneDrive - Ministério do Desenvolvimento e Assistência Social\Power BI\Python\Consulta Convênios\propostas_iniciais_convenios.xlsx"
caminho_resultados = r"C:\Users\diego.brito\OneDrive - Ministério do Desenvolvimento e Assistência Social\Power BI\Python\Consulta Convênios\Consulta Convênios\Consulta convenios.xlsx"

# Verificar se o arquivo base existe
if not os.path.exists(caminho_planilha):
    print(f"Arquivo de entrada {caminho_planilha} não encontrado.")
    return

# Carregar as propostas
propostas = pd.read_excel(caminho_planilha).fillna("")
resultados = []

# Processar cada proposta
for _, row in propostas.iterrows():
    proposta_numero = row["NºProposta"].strip()
    info_adicional = {
        "Instrumento": row.get("Instrumento", ""),
        "AçãoOrçamentária": row.get("AçãoOrçamentária", ""),
        "OrigemRecurso": row.get("OrigemRecurso", ""),
        "CoordenaçãoResponsável": row.get("CoordenaçãoResponsável", ""),
        "Processo": row.get("Processo", ""),
        "TécnicoResponsável": row.get("TécnicoResponsável", "")
    }
    resultado = processar_proposta(driver, proposta_numero, info_adicional)
    resultados.append(resultado)

    # Salvar incrementalmente
    salvar_resultados(resultados, caminho_resultados)

print("Consulta Convênios concluída.")


# 4. Função principal para orquestrar a execução
def executar_todos_os_scripts():
    """
    Executa os scripts de consulta sequencialmente (Requisitos, Parecer e Convênios)
    e registra informações de progresso e erros.
    """
    import time
    start_time = time.time()  # Monitorar tempo total de execução

    driver = None  # Inicialização do driver
    try:
        print("Iniciando conexão com o navegador...")
        driver = conectar_navegador_existente()
        print("Navegador conectado com sucesso!")

        # Executar os scripts sequencialmente com logs claros
        try:
            print("\nIniciando execução do script de Requisitos...")
            executar_requisitos(driver)
            print("Script de Requisitos concluído com sucesso.")
        except Exception as e:
            print(f"Erro ao executar script de Requisitos: {e}")

        try:
            print("\nIniciando execução do script de Parecer...")
            executar_parecer(driver)
            print("Script de Parecer concluído com sucesso.")
        except Exception as e:
            print(f"Erro ao executar script de Parecer: {e}")

        try:
            print("\nIniciando execução do script de Convênios...")
            executar_convenios(driver)
            print("Script de Convênios concluído com sucesso.")
        except Exception as e:
            print(f"Erro ao executar script de Convênios: {e}")

    except Exception as e:
        print(f"Erro durante a inicialização ou execução dos scripts: {e}")
    finally:
        # Fechar o driver se estiver ativo
        if driver:
            try:
                print("\nFinalizando navegador...")
                driver.quit()
                print("Navegador finalizado com sucesso.")
            except Exception as e:
                print(f"Erro ao finalizar o navegador: {e}")

        # Exibir tempo total de execução
        end_time = time.time()
        tempo_total = end_time - start_time
        minutos = int(tempo_total // 60)
        segundos = int(tempo_total % 60)
        print(f"\nTempo total de execução: {minutos}m:{segundos:02d}s.")


if __name__ == "__main__":
    executar_todos_os_scripts()
