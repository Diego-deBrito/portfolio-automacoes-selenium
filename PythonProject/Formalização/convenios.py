import os
import time
from datetime import datetime

import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager

# Funções específicas para processar cada aba (importadas de outros scripts)
from requisitos import processar_aba_requisitos
from parecer import processar_aba_parecer
from convenios import processar_aba_convenios


def conectar_navegador_existente():
    """
    Conecta ao navegador Chrome já aberto para reutilização durante o processo.
    Isso utiliza a depuração remota na porta configurada (9222).
    """
    opcoes_navegador = webdriver.ChromeOptions()
    opcoes_navegador.debugger_address = "localhost:9222"  # Porta configurada para depuração remota
    navegador = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=opcoes_navegador)
    return navegador


def processar_proposta_completa(navegador, numero_proposta, informacoes_adicionais):
    """
    Processa todas as abas (Requisitos, Parecer e Convênios) para uma única proposta.

    Args:
        navegador: Instância do navegador Selenium.
        numero_proposta: Número da proposta a ser processada.
        informacoes_adicionais: Dicionário com informações adicionais sobre a proposta.

    Returns:
        Um dicionário com os dados consolidados das abas processadas.
    """
    print(f"Iniciando o processamento da proposta: {numero_proposta}")

    try:
        # Processar a aba Requisitos
        print("==> Processando a aba Requisitos...")
        dados_requisitos = processar_aba_requisitos(navegador, numero_proposta)

        # Processar a aba Parecer
        print("==> Processando a aba Parecer...")
        dados_parecer = processar_aba_parecer(navegador, numero_proposta)

        # Processar a aba Convênios
        print("==> Processando a aba Convênios...")
        dados_convenios = processar_aba_convenios(navegador, numero_proposta)

        # Consolidar os dados da proposta em um único dicionário
        dados_consolidados = {
            "Proposta": numero_proposta,
            "Instrumento": informacoes_adicionais.get("Instrumento", ""),
            "AçãoOrçamentária": informacoes_adicionais.get("AçãoOrçamentária", ""),
            "OrigemRecurso": informacoes_adicionais.get("OrigemRecurso", ""),
            "CoordenaçãoResponsável": informacoes_adicionais.get("CoordenaçãoResponsável", ""),
            "Processo": informacoes_adicionais.get("Processo", ""),
            "TécnicoResponsável": informacoes_adicionais.get("TécnicoResponsável", ""),
            **dados_requisitos,
            **dados_parecer,
            **dados_convenios
        }

        return dados_consolidados

    except Exception as erro:
        print(f"Erro ao processar a proposta {numero_proposta}: {erro}")
        return None


def processar_todas_propostas():
    """
    Processa todas as propostas presentes em uma planilha, passando por todas as abas (Requisitos, Parecer e Convênios)
    para cada proposta, antes de avançar para a próxima.
    """
    # Caminho para a planilha de entrada com as propostas e para o arquivo de saída consolidado
    caminho_planilha_entrada = "propostas_iniciais.xlsx"
    caminho_planilha_saida = "Consolidado_Consulta.xlsx"

    # Verificar se o arquivo de entrada existe
    if not os.path.exists(caminho_planilha_entrada):
        print(f"Erro: O arquivo de entrada '{caminho_planilha_entrada}' não foi encontrado.")
        return

    # Carregar a planilha com os números das propostas
    try:
        planilha_propostas = pd.read_excel(caminho_planilha_entrada)
        planilha_propostas['NºProposta'] = planilha_propostas['NºProposta'].astype(str).str.strip()
    except Exception as erro:
        print(f"Erro ao carregar a planilha: {erro}")
        return

    # Inicializar o navegador
    navegador = conectar_navegador_existente()

    # Lista para armazenar os resultados de cada proposta processada
    resultados_propostas = []
    total_propostas_processadas = 0

    try:
        # Loop para processar cada proposta na planilha
        for indice, linha in planilha_propostas.iterrows():
            numero_proposta = linha['NºProposta']
            informacoes_adicionais = {
                "Instrumento": linha.get("Instrumento", ""),
                "AçãoOrçamentária": linha.get("AçãoOrçamentária", ""),
                "OrigemRecurso": linha.get("OrigemRecurso", ""),
                "CoordenaçãoResponsável": linha.get("CoordenaçãoResponsável", ""),
                "Processo": linha.get("Processo", ""),
                "TécnicoResponsável": linha.get("TécnicoResponsável", "")
            }

            print(f"\nProcessando a proposta {numero_proposta} ({indice + 1}/{len(planilha_propostas)})...")

            # Processar a proposta completa (Requisitos, Parecer e Convênios)
            dados_proposta = processar_proposta_completa(navegador, numero_proposta, informacoes_adicionais)

            if dados_proposta:
                resultados_propostas.append(dados_proposta)
                total_propostas_processadas += 1

            # Navegar para a próxima proposta
            try:
                xpath_nova_pesquisa = '/html/body/div[3]/div[2]/div[6]/a[2]'
                WebDriverWait(navegador, 10).until(EC.element_to_be_clickable((By.XPATH, xpath_nova_pesquisa))).click()
                print("Botão 'Nova Pesquisa' clicado com sucesso.")
            except TimeoutException:
                print("Erro ao clicar em 'Nova Pesquisa'. Tentando recarregar a página...")
                navegador.refresh()
                time.sleep(5)

        # Salvar os resultados processados em uma planilha Excel
        if resultados_propostas:
            try:
                dataframe_resultados = pd.DataFrame(resultados_propostas)
                dataframe_resultados.to_excel(caminho_planilha_saida, index=False)
                print(f"\nResultados consolidados salvos em '{caminho_planilha_saida}'.")
            except Exception as erro:
                print(f"Erro ao salvar os resultados em Excel: {erro}")
        else:
            print("Nenhum resultado foi processado. Verifique se as propostas foram corretamente consultadas.")

    except Exception as erro_principal:
        print(f"Erro durante o processamento de todas as propostas: {erro_principal}")
    finally:
        navegador.quit()
        print(f"\nProcessamento concluído. Total de propostas processadas: {total_propostas_processadas}")


if __name__ == "__main__":
    processar_todas_propostas()


def processar_aba_convenios(driver, numero_proposta):
    """
    Processa a aba Convênios para uma proposta específica.

    Args:
        driver: Instância do navegador Selenium.
        numero_proposta: Número da proposta a ser processada.

    Returns:
        Um dicionário contendo os dados processados na aba Convênios.
    """
    print(f"Iniciando o processamento da aba Convênios para a proposta {numero_proposta}...")

    try:
        # Navegar para a aba Convênios
        print("Navegando para a aba Convênios...")
        xpath_aba_convenios = '/html/body/div[3]/div[15]/div[1]/div/div[2]/a[12]/div/span/span'  # Atualizado conforme o script principal
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, xpath_aba_convenios))).click()

        # Garantir que a aba foi carregada
        print("Verificando o carregamento da aba Convênios...")
        xpath_verificacao_carregamento = "//h1[contains(text(), 'Convênios')]"  # Atualizar conforme necessário
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, xpath_verificacao_carregamento)))

        # Extração dos dados específicos na aba Convênios
        print("Extraindo dados da aba Convênios...")

        # Exemplos de XPaths para os campos a serem extraídos
        campo_convenio_data_xpath = '/html/body/div[3]/div[15]/div[3]/div[2]/table/tbody/tr/td[1]'
        campo_convenio_status_xpath = '/html/body/div[3]/div[15]/div[3]/div[2]/table/tbody/tr/td[2]'

        # Obter a data mais recente no campo de Convênios
        try:
            elementos_convenio_data = driver.find_elements(By.XPATH, campo_convenio_data_xpath)
            lista_datas_convenios = [
                datetime.strptime(el.text.strip(), "%d/%m/%Y %H:%M:%S") for el in elementos_convenio_data if el.text.strip()
            ]
            data_mais_recente_convenios = max(lista_datas_convenios).strftime('%d/%m/%Y %H:%M:%S') if lista_datas_convenios else ""
        except Exception as e:
            print(f"Erro ao extrair a data mais recente na aba Convênios: {e}")
            data_mais_recente_convenios = ""

        # Obter o status mais recente no campo de Convênios
        try:
            elementos_convenio_status = driver.find_elements(By.XPATH, campo_convenio_status_xpath)
            status_mais_recente_convenios = elementos_convenio_status[-1].text.strip() if elementos_convenio_status else ""
        except Exception as e:
            print(f"Erro ao extrair o status mais recente na aba Convênios: {e}")
            status_mais_recente_convenios = ""

        # Criar o dicionário com os dados coletados
        dados_coletados = {
            "DataMaisRecenteConvenios": data_mais_recente_convenios,
            "StatusMaisRecenteConvenios": status_mais_recente_convenios,
        }

        print(f"Dados extraídos da aba Convênios para a proposta {numero_proposta}: {dados_coletados}")
        return dados_coletados

    except Exception as erro:
        print(f"Ocorreu um erro ao processar a aba Convênios para a proposta {numero_proposta}: {erro}")
        return {}
