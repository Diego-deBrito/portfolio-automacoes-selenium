import time
import os
import pandas as pd
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, ElementNotInteractableException
from webdriver_manager.chrome import ChromeDriverManager

# Importação de funções específicas de processamento de abas
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


def clicar_nova_pesquisa(driver):
    """
    Tenta clicar no botão "Nova Pesquisa" para avançar para a próxima proposta.
    """
    xpaths_botoes_nova_pesquisa = [
        '/html/body/div[3]/div[3]/div[6]/a[2]',  # Primeiro XPath
        '/html/body/div[3]/div[2]/div[6]/a[2]',  # Segundo XPath
        '/html/body[1]/div[3]/div[3]/div[6]/a[2]'  # Terceiro XPath
    ]

    botao_encontrado = False

    for xpath in xpaths_botoes_nova_pesquisa:
        if elemento_existe(driver, xpath, tempo_espera=2):
            try:
                WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, xpath))).click()
                print(f"Botão 'Nova Pesquisa' clicado com sucesso usando o XPath: {xpath}")
                botao_encontrado = True
                break
            except (NoSuchElementException, TimeoutException):
                print(f"Falha ao clicar no botão 'Nova Pesquisa' com o XPath: {xpath}. Tentando o próximo...")

    # Se nenhum dos XPaths funcionar, recarregar a página e tentar novamente
    if not botao_encontrado:
        print(f"Botão 'Nova Pesquisa' não encontrado após tentar todos os XPaths. Recarregando a página.")
        driver.refresh()
        time.sleep(3)  # Aguardar recarregamento
        return clicar_nova_pesquisa(driver)  # Tentar novamente após recarregar



def elemento_existe(driver, xpath, tempo_espera=1):
    """
    Verifica se um elemento existe no DOM com um tempo de espera especificado.
    """
    try:
        WebDriverWait(driver, tempo_espera).until(EC.presence_of_element_located((By.XPATH, xpath)))
        return True
    except TimeoutException:
        print(f"Elemento não encontrado: {xpath}")
        return False




def processar_proposta_completa(driver, numero_proposta, informacoes_adicionais):
    """
    Processa todas as abas (Requisitos → Parecer → Convênios) para uma única proposta,
    garantindo que todas sejam concluídas antes de passar para a próxima.

    Args:
        driver: Instância do navegador Selenium.
        numero_proposta: Número da proposta a ser processada.
        informacoes_adicionais: Dicionário com informações adicionais sobre a proposta.

    Returns:
        Um dicionário consolidando os dados processados de todas as abas.
    """
    print(f"\nIniciando o processamento completo da proposta {numero_proposta}...")

    # Consolidar os dados de todas as abas em um único dicionário
    dados_consolidados = {
        "Proposta": numero_proposta,
        "Instrumento": informacoes_adicionais.get("Instrumento", ""),
        "AçãoOrçamentária": informacoes_adicionais.get("AçãoOrçamentária", ""),
        "OrigemRecurso": informacoes_adicionais.get("OrigemRecurso", ""),
        "CoordenaçãoResponsável": informacoes_adicionais.get("CoordenaçãoResponsável", ""),
        "Processo": informacoes_adicionais.get("Processo", ""),
        "TécnicoResponsável": informacoes_adicionais.get("TécnicoResponsável", ""),
    }

    try:
        # Etapa 1: Processar a aba Requisitos
        print("==> Processando a aba Requisitos...")
        if ir_para_aba(driver, "Requisitos"):
            try:
                dados_requisitos = processar_aba_requisitos(driver, numero_proposta)
                dados_consolidados.update(dados_requisitos)
                print("==> Aba Requisitos processada com sucesso!")
            except Exception as e:
                print(f"Erro ao processar a aba Requisitos: {e}")
                dados_consolidados["Erro_Requisitos"] = str(e)
        else:
            dados_consolidados["Erro_Requisitos"] = "Falha ao navegar para a aba Requisitos"

        # Etapa 2: Processar a aba Parecer
        print("==> Processando a aba Parecer...")
        if ir_para_aba(driver, "Parecer"):
            try:
                dados_parecer = processar_aba_parecer(driver, numero_proposta)
                dados_consolidados.update(dados_parecer)
                print("==> Aba Parecer processada com sucesso!")
            except Exception as e:
                print(f"Erro ao processar a aba Parecer: {e}")
                dados_consolidados["Erro_Parecer"] = str(e)
        else:
            dados_consolidados["Erro_Parecer"] = "Falha ao navegar para a aba Parecer"

        # Etapa 3: Processar a aba Convênios
        print("==> Processando a aba Convênios...")
        if ir_para_aba(driver, "Convênios"):
            try:
                dados_convenios = processar_aba_convenios(driver, numero_proposta)
                dados_consolidados.update(dados_convenios)
                print("==> Aba Convênios processada com sucesso!")
            except Exception as e:
                print(f"Erro ao processar a aba Convênios: {e}")
                dados_consolidados["Erro_Convenios"] = str(e)
        else:
            dados_consolidados["Erro_Convenios"] = "Falha ao navegar para a aba Convênios"

    except Exception as erro_geral:
        print(f"Erro geral ao processar a proposta {numero_proposta}: {erro_geral}")
        dados_consolidados["Erro_Geral"] = str(erro_geral)

    print(f"\nProcessamento completo da proposta {numero_proposta}.")
    print(f"Dados consolidados: {dados_consolidados}")

    return dados_consolidados


def ir_para_aba(driver, nome_aba):
    """
    Navega para uma aba específica no sistema com base no nome da aba.

    Args:
        driver: Instância do navegador Selenium.
        nome_aba: Nome da aba a ser acessada ("Requisitos", "Parecer", "Convênios").

    Returns:
        True se a navegação foi bem-sucedida, False caso contrário.
    """
    try:
        # Mapeamento dos XPaths para as abas principais e subabas
        abas = {
            "Requisitos": "/html/body/div[3]/div[15]/div[1]/div/div[1]/a[3]/div/span/span",
            "Parecer": "/html/body/div[3]/div[15]/div[1]/div/div[2]/a[10]/div/span/span",
            "Convênios": "/html/body/div[3]/div[15]/div[1]/div/div[1]/a[4]/div/span/span"
        }

        subabas = {
            "Requisitos": [
                "/html/body/div[3]/div[15]/div[1]/div/div[2]/a[12]/div/span/span",
                "/html/body/div[3]/div[15]/div[1]/div/div[1]/a[3]/div/span/span",
                "/html/body/div[3]/div[16]/div[1]/div/div[2]/a[29]/div/span/span"
            ],
            "Parecer": [
                "/html/body/div[3]/div[15]/div[1]/div/div[2]/a[11]/div/span/span"
            ],
            "Convênios": [
                "/html/body/div[3]/div[15]/div[1]/div/div[2]/a[9]/div/span/span"
            ]
        }

        if nome_aba not in abas:
            print(f"Nome da aba inválido: {nome_aba}")
            return False

        # Navegar para a aba principal
        xpath_aba = abas[nome_aba]
        print(f"Tentando acessar a aba principal: {nome_aba}...")
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, xpath_aba))).click()

        # Verificar se há subabas a serem acessadas
        if nome_aba in subabas:
            for xpath_subaba in subabas[nome_aba]:
                if elemento_existe(driver, xpath_subaba, tempo_espera=2):
                    print(f"Tentando acessar subaba de {nome_aba}...")
                    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, xpath_subaba))).click()
                    break

        print(f"Aba '{nome_aba}' acessada com sucesso.")
        return True
    except TimeoutException:
        print(f"Timeout ao tentar acessar a aba '{nome_aba}'.")
        return False
    except Exception as erro:
        print(f"Erro ao navegar para a aba '{nome_aba}': {erro}")
        return False




def salvar_progresso(resultado, caminho_arquivo):
    """
    Salva os dados processados em um arquivo Excel, adicionando os novos resultados ao arquivo existente.

    Args:
        resultado: Dicionário com os dados da proposta processada.
        caminho_arquivo: Caminho completo do arquivo Excel para salvar os dados.
    """
    try:
        if not os.path.exists(caminho_arquivo):
            # Se o arquivo não existir, cria um novo
            pd.DataFrame([resultado]).to_excel(caminho_arquivo, index=False)
        else:
            # Caso exista, adiciona os novos dados ao final
            df_existente = pd.read_excel(caminho_arquivo)
            df_atualizado = pd.concat([df_existente, pd.DataFrame([resultado])], ignore_index=True)
            df_atualizado.to_excel(caminho_arquivo, index=False)
        print(f"Progresso salvo com sucesso no arquivo '{caminho_arquivo}'!")
    except Exception as erro:
        print(f"Erro ao salvar os dados no arquivo Excel: {erro}")


def processar_todas_propostas():
    """
    Processa todas as propostas presentes na planilha, garantindo que cada proposta passe pelas três abas
    (Requisitos, Parecer e Convênios) antes de avançar para a próxima proposta.
    """
    caminho_planilha_entrada = r"C:\Users\diego.brito\OneDrive - Ministério do Desenvolvimento e Assistência Social\Power BI\Python\propostas_iniciais.xlsx"
    caminho_planilha_saida = r"C:\Users\diego.brito\OneDrive - Ministério do Desenvolvimento e Assistência Social\Power BI\Python\Consulta Transferegov Consolidado.xlsx"

    # Verificar se o arquivo de entrada existe
    if not os.path.exists(caminho_planilha_entrada):
        print(f"Erro: O arquivo de entrada '{caminho_planilha_entrada}' não foi encontrado.")
        return

    # Carregar a planilha com os números das propostas
    try:
        planilha_propostas = pd.read_excel(caminho_planilha_entrada)
        planilha_propostas['NºProposta'] = planilha_propostas['NºProposta'].astype(str).str.strip()
    except Exception as erro_carregamento:
        print(f"Erro ao carregar a planilha: {erro_carregamento}")
        return

    # Inicializar o navegador
    driver = conectar_navegador_existente()

    try:
        # Processar cada proposta da planilha
        for indice, linha in planilha_propostas.iterrows():
            numero_proposta = linha['NºProposta']
            informacoes_adicionais = linha.to_dict()  # Informações adicionais da linha
            print(f"\n=== Processando a proposta {numero_proposta} ({indice + 1}/{len(planilha_propostas)}) ===")

            try:
                # Processar a proposta completa (Requisitos → Parecer → Convênios)
                dados_proposta = processar_proposta_completa(driver, numero_proposta, informacoes_adicionais)

                # Salvar os dados após processar todas as abas
                salvar_progresso(dados_proposta, caminho_planilha_saida)

                # Avançar para a próxima proposta
                clicar_nova_pesquisa(driver)

            except Exception as erro_proposta:
                print(f"Erro ao processar a proposta {numero_proposta}: {erro_proposta}")
                # Registrar erro e continuar com a próxima proposta

    except Exception as erro_principal:
        print(f"Erro durante o processamento: {erro_principal}")

    finally:
        driver.quit()
        print("Processamento concluído. Navegador encerrado.")
if __name__ == "__main__":
    processar_todas_propostas()
