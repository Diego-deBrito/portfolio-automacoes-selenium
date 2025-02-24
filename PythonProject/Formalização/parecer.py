import time
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from datetime import datetime
import pandas as pd
import os

from SEI.chatbot import resultados

# Variável para rastrear as propostas já processadas e evitar duplicação
propostas_processadas = set()


# 1. Função para conectar ao navegador já aberto
def conectar_navegador_existente():
    options = webdriver.ChromeOptions()
    options.debugger_address = "localhost:9222"  # Porta que o Chrome está utilizando para depuração
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    return driver


# 2. Função auxiliar para verificar se um elemento existe com tempo de espera reduzido
def elemento_existe(driver, xpath, tempo_espera=0.5):
    try:
        WebDriverWait(driver, tempo_espera).until(EC.presence_of_element_located((By.XPATH, xpath)))
        return True
    except TimeoutException:
        return False


# 3. Função para reiniciar a navegação a partir da proposta seguinte
def reiniciar_navegacao(driver):
    try:
        # Clicar no menu Propostas (aumentei o tempo de espera para garantir o carregamento)
        WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[3]/div[1]/div[1]/div[1]/div[3]'))).click()
        WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[3]/div[2]/div/div[1]/ul/li[3]/a'))).click()
    except TimeoutException:
        print("Elemento não encontrado ou levou muito tempo para carregar durante a navegação inicial.")
        raise


# 4. Função para tentar converter data com múltiplos formatos
def converter_data(data_text):
    formatos = ["%d/%m/%Y %H:%M:%S", "%d/%m/%Y"]  # Lista de formatos possíveis
    for formato in formatos:
        try:
            return datetime.strptime(data_text, formato)
        except ValueError:
            continue
    print(f"Erro ao converter data '{data_text}', formatos esperados: {formatos}")
    return None


# 5. Função para processar uma única proposta
def processar_proposta(driver, proposta_numero, info_adicional):
    # Digitar o número da proposta
    driver.find_element(By.XPATH, '/html/body/div[3]/div[15]/div[3]/div/div/form/table/tbody/tr[1]/td[2]/input').clear()
    driver.find_element(By.XPATH,
                        '/html/body/div[3]/div[15]/div[3]/div/div/form/table/tbody/tr[1]/td[2]/input').send_keys(
        proposta_numero)

    # Verificar se o botão "Consultar" está presente e clicá-lo
    if elemento_existe(driver, '/html/body/div[3]/div[15]/div[3]/div/div/form/table/tbody/tr[1]/td[2]/span/input',
                       tempo_espera=0.5):
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
            (By.XPATH, '/html/body/div[3]/div[15]/div[3]/div/div/form/table/tbody/tr[1]/td[2]/span/input'))).click()
    else:
        print(f"Botão 'Consultar' não encontrado para a proposta {proposta_numero}")
        return "não localizada"

    # Verificar se o número da proposta aparece e clicar nele
    try:
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
            (By.XPATH, '/html/body/div[3]/div[15]/div[3]/div[3]/table/tbody/tr/td[1]/div/a'))).click()
    except TimeoutException:
        print(f"Proposta {proposta_numero} não encontrada!")
        return "não localizada"

    # Clicar na aba Plano de trabalho
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
        (By.XPATH, '/html/body/div[3]/div[15]/div[1]/div/div[1]/a[2]/div/span/span'))).click()

    # Clicar na aba Pareceres, tentando múltiplos XPaths
    try:
        # Primeira tentativa de clicar
        WebDriverWait(driver, 5).until(EC.element_to_be_clickable(
            (By.XPATH, '/html/body/div[3]/div[15]/div[1]/div/div[2]/a[10]/div/span/span'))).click()
    except TimeoutException:
        # Segunda tentativa de clicar
        try:
            driver.find_element(By.XPATH, '/html/body/div[3]/div[15]/div[1]/div/div[2]/a[11]/div/span/span').click()
        except NoSuchElementException:
            print(f"Elemento 'Pareceres' não encontrado para a proposta {proposta_numero}")
            return "não localizada"

    # Realizar a consulta de data mais recente nos elementos das colunas fornecidas
    data_mais_recente_1 = ""
    data_mais_recente_2 = ""

    try:
        # Consulta na primeira coluna
        if elemento_existe(driver, '/html/body/div[3]/div[15]/div[3]/div[2]/table/thead/tr/th[1]', tempo_espera=0.5):
            datas_coluna_1 = driver.find_elements(By.XPATH,
                                                  '/html/body/div[3]/div[15]/div[3]/div[2]/table/tbody/tr/td[1]')
            lista_datas_1 = [converter_data(data_element.text.strip()) for data_element in datas_coluna_1 if
                             data_element.text.strip()]
            if lista_datas_1:
                data_mais_recente_1 = max(lista_datas_1).strftime('%d/%m/%Y %H:%M:%S')
                print(f"Data mais recente na primeira coluna: {data_mais_recente_1}")
            else:
                print(f"Nenhuma data válida encontrada na primeira coluna.")
        else:
            print(f"Elemento da primeira coluna não encontrado, prosseguindo para a próxima etapa.")

        # Consulta na segunda coluna
        if elemento_existe(driver, '/html/body/div[3]/div[15]/div[3]/div[4]/div/div/form/div[1]/table/thead/tr/th[1]',
                           tempo_espera=0.5):
            datas_coluna_2 = driver.find_elements(By.XPATH,
                                                  '/html/body/div[3]/div[15]/div[3]/div[4]/div/div/form/div[1]/table/tbody/tr/td[1]')
            lista_datas_2 = [converter_data(data_element.text.strip()) for data_element in datas_coluna_2 if
                             data_element.text.strip()]
            if lista_datas_2:
                data_mais_recente_2 = max(lista_datas_2).strftime('%d/%m/%Y %H:%M:%S')
                print(f"Data mais recente na segunda coluna: {data_mais_recente_2}")
            else:
                print(f"Nenhuma data válida encontrada na segunda coluna.")
        else:
            print(f"Elemento da segunda coluna não encontrado, prosseguindo para a próxima etapa.")

    except Exception as e:
        print(f"Erro ao buscar as datas mais recentes: {e}")

    # Retornar os dados da proposta processada
    return {
        "Proposta": proposta_numero,
        "Instrumento": info_adicional["Instrumento"],
        "AçãoOrçamentária": info_adicional["AçãoOrçamentária"],
        "OrigemRecurso": info_adicional["OrigemRecurso"],
        "CoordenaçãoResponsável": info_adicional["CoordenaçãoResponsável"],
        "Processo": info_adicional["Processo"],
        "TécnicoResponsável": info_adicional["TécnicoResponsável"],
        "DataMaisRecenteProposta": data_mais_recente_1,
        "DataMaisRecentePlanodeTrabalho": data_mais_recente_2
    }


# 6. Função para salvar o progresso no Excel usando ExcelWriter
def salvar_progresso(resultado):
    caminho_arquivo = r"C:\Users\diego.brito\OneDrive - Ministério do Desenvolvimento e Assistência Social\Power BI\Python\Consulta Parecer\Consulta Transferegov Parecer Novo.xlsx"

    try:
        # Se a planilha já existir, carregue o DataFrame atual e concatene os resultados
        if os.path.exists(caminho_arquivo):
            with pd.ExcelWriter(caminho_arquivo, mode='a', engine='openpyxl', if_sheet_exists='overlay') as writer:
                df_existente = pd.read_excel(caminho_arquivo)
                df_atualizado = pd.concat([df_existente, pd.DataFrame([resultado])], ignore_index=True)
                df_atualizado.to_excel(writer, index=False, sheet_name='Sheet1')
        else:
            # Caso não exista, crie uma nova planilha
            with pd.ExcelWriter(caminho_arquivo, engine='openpyxl') as writer:
                df_atualizado = pd.DataFrame([resultado])
                df_atualizado.to_excel(writer, index=False)

        print(f"Progresso salvo com sucesso no arquivo '{caminho_arquivo}'!")
    except Exception as e:
        print(f"Erro ao salvar o arquivo Excel: {e}")


# 7. Função para limpar a planilha antes de começar a processar
def limpar_planilha():
    caminho_arquivo = r"C:\Users\diego.brito\OneDrive - Ministério do Desenvolvimento e Assistência Social\Power BI\Python\Consulta Parecer\Consulta Transferegov Parecer Novo.xlsx"
    try:
        # Se o arquivo existir, será apagado
        if os.path.exists(caminho_arquivo):
            os.remove(caminho_arquivo)
            print(f"Arquivo '{caminho_arquivo}' foi apagado para iniciar uma nova consulta.")
    except Exception as e:
        print(f"Erro ao tentar apagar o arquivo Excel: {e}")


# 8. Laço principal para processar todas as propostas
def processar_todas_propostas():
    propostas_consultadas = 0
    tempo_acumulado = 0

    try:
        driver = conectar_navegador_existente()
        reiniciar_navegacao(driver)

        for index, row in df_propostas.iterrows():
            proposta_numero = row['NºProposta']

            # Verificar se a proposta já foi processada
            if proposta_numero in propostas_processadas:
                print(f"Proposta {proposta_numero} já foi processada. Pulando...")
                continue

            print(f"Processando proposta: {proposta_numero}")

            try:
                tempo_inicio = time.time()

                info_adicional = {
                    "Instrumento": row["Instrumento"],
                    "AçãoOrçamentária": row["AçãoOrçamentária"],
                    "OrigemRecurso": row["OrigemRecurso"],
                    "CoordenaçãoResponsável": row["CoordenaçãoResponsável"],
                    "Processo": row["Processo"],
                    "TécnicoResponsável": row["TécnicoResponsável"]
                }

                dados_proposta = processar_proposta(driver, proposta_numero, info_adicional)

                # Marcar esta proposta como processada
                propostas_processadas.add(proposta_numero)

                tempo_fim = time.time()
                tempo_consulta = tempo_fim - tempo_inicio
                tempo_acumulado += tempo_consulta

                propostas_consultadas += 1

                print(f"Proposta {proposta_numero} consultada.")
                print(f"Tempo para consulta desta proposta: {tempo_consulta:.2f} segundos.")
                print(f"Tempo acumulado: {int(tempo_acumulado // 60)}m:{int(tempo_acumulado % 60)}s.")
                print(f"Total de propostas consultadas: {propostas_consultadas}\n")

                salvar_progresso(dados_proposta)

                # Clicar no botão "Nova Pesquisa" para pesquisar a próxima proposta
                xpaths_nova_pesquisa = [
                    '/html/body/div[3]/div[3]/div[6]/a[2]',  # Primeira tentativa
                    '/html/body/div[3]/div[2]/div[6]/a[2]',  # Segunda tentativa
                    '/html/body[1]/div[3]/div[3]/div[6]/a[2]'  # Terceira tentativa
                ]

                botao_encontrado = False

                for xpath in xpaths_nova_pesquisa:
                    if elemento_existe(driver, xpath, tempo_espera=0.5):
                        try:
                            WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, xpath))).click()
                            print(f"Botão 'Nova Pesquisa' clicado com sucesso usando XPath: {xpath}")
                            botao_encontrado = True
                            break
                        except (NoSuchElementException, TimeoutException):
                            print(
                                f"Falha ao clicar no botão 'Nova Pesquisa' com o XPath: {xpath}. Tentando o próximo...")
                    else:
                        print(f"Botão 'Nova Pesquisa' não encontrado com o XPath: {xpath}. Tentando o próximo...")

                # Se nenhum dos XPaths funcionar, recarregar a página e reiniciar a navegação
                if not botao_encontrado:
                    print(f"Botão 'Nova Pesquisa' não encontrado após tentar todos os XPaths. Recarregando a página.")
                    reiniciar_navegacao(driver)

            except Exception as e:
                print(f"Erro ao processar a proposta {proposta_numero}: {e}")
                salvar_progresso(resultados)
                driver.quit()
                driver = conectar_navegador_existente()
                reiniciar_navegacao(driver)

    finally:
        driver.quit()


# 9. Carregar a planilha de propostas iniciais
caminho_planilha = r"C:\Users\diego.brito\OneDrive - Ministério do Desenvolvimento e Assistência Social\Power BI\Python\Consulta Parecer\propostas_iniciais_Parecer.xlsx"
df_propostas = pd.read_excel(caminho_planilha)
df_propostas.columns = df_propostas.columns.str.strip()
df_propostas['NºProposta'] = df_propostas['NºProposta'].str.strip()

# 10. Limpar a planilha antes de começar
limpar_planilha()

# 11. Executar o processamento de todas as propostas
processar_todas_propostas()

print("Execução concluída e dados salvos com sucesso!")


def processar_aba_parecer(driver, numero_proposta):
    """
    Processa a aba Parecer para uma proposta específica.

    Args:
        driver: Instância do navegador Selenium.
        numero_proposta: Número da proposta a ser processada.

    Returns:
        Um dicionário contendo os dados processados na aba Parecer.
    """
    print(f"Iniciando o processamento da aba Parecer para a proposta {numero_proposta}...")

    try:
        # Navegar para a aba Parecer
        print("Navegando para a aba Parecer...")
        xpath_aba_parecer = '/html/body/div[3]/div[15]/div[1]/div/div[2]/a[10]/div/span/span'  # Atualizar com o XPath real
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, xpath_aba_parecer))).click()

        # Garantir que a aba foi carregada
        print("Verificando o carregamento da aba Parecer...")
        xpath_verificacao_carregamento = "//h1[contains(text(), 'Parecer')]"  # Atualizar com o título correto da aba, se necessário
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, xpath_verificacao_carregamento)))

        # Extração dos dados específicos da aba Parecer
        print("Extraindo dados da aba Parecer...")

        # Exemplo de campos para extração (atualizar XPaths conforme necessário)
        campo_parecer_data_xpath = '/html/body/div[3]/div[15]/div[3]/div[2]/table/tbody/tr/td[1]'
        campo_parecer_status_xpath = '/html/body/div[3]/div[15]/div[3]/div[2]/table/tbody/tr/td[2]'

        # Obter a data mais recente no campo de parecer
        try:
            elementos_parecer_data = driver.find_elements(By.XPATH, campo_parecer_data_xpath)
            lista_datas_parecer = [
                datetime.strptime(el.text.strip(), "%d/%m/%Y %H:%M:%S") for el in elementos_parecer_data if el.text.strip()
            ]
            data_mais_recente_parecer = max(lista_datas_parecer).strftime('%d/%m/%Y %H:%M:%S') if lista_datas_parecer else ""
        except Exception as e:
            print(f"Erro ao extrair a data mais recente na aba Parecer: {e}")
            data_mais_recente_parecer = ""

        # Obter o status mais recente no campo de parecer
        try:
            elementos_parecer_status = driver.find_elements(By.XPATH, campo_parecer_status_xpath)
            status_mais_recente_parecer = elementos_parecer_status[-1].text.strip() if elementos_parecer_status else ""
        except Exception as e:
            print(f"Erro ao extrair o status mais recente na aba Parecer: {e}")
            status_mais_recente_parecer = ""

        # Criar o dicionário com os dados coletados
        dados_coletados = {
            "DataMaisRecenteParecer": data_mais_recente_parecer,
            "StatusMaisRecenteParecer": status_mais_recente_parecer,
        }

        print(f"Dados extraídos da aba Parecer para a proposta {numero_proposta}: {dados_coletados}")
        return dados_coletados

    except Exception as erro:
        print(f"Ocorreu um erro ao processar a aba Parecer para a proposta {numero_proposta}: {erro}")
        return {}
