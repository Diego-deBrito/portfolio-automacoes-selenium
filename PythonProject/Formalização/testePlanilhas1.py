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

# Configuração do Navegador
def conectar_navegador_existente():
    """
    Conecta ao navegador Chrome já aberto para reutilização durante o processo.
    Isso utiliza a depuração remota na porta configurada (9222).
    """
    try:
        opcoes_navegador = webdriver.ChromeOptions()
        opcoes_navegador.debugger_address = "localhost:9222"  # Porta configurada para depuração remota
        navegador = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=opcoes_navegador)
        return navegador
    except Exception as e:
        print(f"Erro ao conectar ao navegador existente: {e}")
        raise

# Verificar se o elemento existe no DOM
def elemento_existe(driver, xpath, tempo_espera=2):
    """
    Verifica se um elemento existe no DOM com um tempo de espera especificado.
    """
    try:
        WebDriverWait(driver, tempo_espera).until(EC.presence_of_element_located((By.XPATH, xpath)))
        return True
    except TimeoutException:
        return False

# Clicar em um elemento
def clicar_elemento(driver, xpath):
    """
    Garante que o elemento está visível e clicável antes de realizar a interação.
    """
    try:
        elemento = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, xpath)))
        driver.execute_script("arguments[0].scrollIntoView();", elemento)
        elemento.click()
        print(f"Elemento clicado com sucesso: {xpath}")
    except Exception as e:
        print(f"Erro ao clicar no elemento {xpath}: {e}")
        raise


# Navegar para uma aba específica
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
        abas = {
            "Requisitos": "/html/body/div[3]/div[15]/div[1]/div/div[1]/a[3]/div/span/span",
            "Parecer": "/html/body/div[3]/div[15]/div[1]/div/div[2]/a[10]/div/span/span",
            "Convênios": "/html/body/div[3]/div[15]/div[1]/div/div[1]/a[4]/div/span/span"
        }

        if nome_aba not in abas:
            print(f"Nome da aba inválido: {nome_aba}")
            return False

        xpath_aba = abas[nome_aba]
        clicar_elemento(driver, xpath_aba)
        print(f"Aba '{nome_aba}' acessada com sucesso.")
        return True
    except Exception as e:
        print(f"Erro ao navegar para a aba '{nome_aba}': {e}")
        return False

# Processar a aba Requisitos
def processar_aba_requisitos(driver, numero_proposta):
    """
    Processa os dados na aba Requisitos para uma proposta específica.
    """
    try:
        print("Processando a aba Requisitos...")
        data_certidoes_xpath = "/html/body/div[3]/div[16]/div[2]/div[2]/form/div[1]/div[1]/table/tbody/tr[1]/td[2]"
        data_declaracoes_xpath = "/html/body/div[3]/div[16]/div[2]/div[2]/form/div[1]/div[2]/table/tbody/tr[1]/td[2]"

        data_certidoes = buscar_data_mais_recente(driver, data_certidoes_xpath)
        data_declaracoes = buscar_data_mais_recente(driver, data_declaracoes_xpath)

        return {"DataCertidoes": data_certidoes, "DataDeclaracoes": data_declaracoes}
    except Exception as e:
        print(f"Erro ao processar a aba Requisitos para a proposta {numero_proposta}: {e}")
        return {}

# Processar a aba Parecer
def processar_aba_parecer(driver, numero_proposta):
    """
    Processa os dados na aba Parecer para uma proposta específica.
    """
    try:
        print("Processando a aba Parecer...")
        data_parecer_xpath = "/html/body/div[3]/div[15]/div[3]/div[2]/table/tbody/tr/td[1]"

        data_parecer = buscar_data_mais_recente(driver, data_parecer_xpath)

        return {"DataParecer": data_parecer}
    except Exception as e:
        print(f"Erro ao processar a aba Parecer para a proposta {numero_proposta}: {e}")
        return {}

# Buscar a data mais recente em uma tabela
def buscar_data_mais_recente(driver, xpath):
    """
    Busca a data mais recente presente em elementos localizados pelo XPath.
    """
    try:
        elementos = driver.find_elements(By.XPATH, xpath)
        datas = [datetime.strptime(e.text.strip(), "%d/%m/%Y %H:%M:%S") for e in elementos if e.text.strip()]
        return max(datas).strftime('%d/%m/%Y %H:%M:%S') if datas else None
    except Exception as e:
        print(f"Erro ao buscar datas: {e}")
        return None

# Salvar os resultados em um arquivo Excel
def salvar_progresso(resultado, caminho_arquivo):
    """
    Salva os dados processados em um arquivo Excel, adicionando os novos resultados ao arquivo existente.
    """
    try:
        if not os.path.exists(caminho_arquivo):
            pd.DataFrame([resultado]).to_excel(caminho_arquivo, index=False)
        else:
            df_existente = pd.read_excel(caminho_arquivo)
            df_atualizado = pd.concat([df_existente, pd.DataFrame([resultado])], ignore_index=True)
            df_atualizado.to_excel(caminho_arquivo, index=False)
        print(f"Progresso salvo com sucesso no arquivo '{caminho_arquivo}'!")
    except Exception as e:
        print(f"Erro ao salvar os dados no arquivo Excel: {e}")

# Processar a proposta completa
def processar_proposta_completa(driver, numero_proposta, informacoes_adicionais):
    """
    Processa todas as abas (Requisitos → Parecer) para uma única proposta.
    """
    print(f"Processando a proposta {numero_proposta}...")
    dados_consolidados = {"Proposta": numero_proposta}

    try:
        if ir_para_aba(driver, "Requisitos"):
            dados_requisitos = processar_aba_requisitos(driver, numero_proposta)
            dados_consolidados.update(dados_requisitos)

        if ir_para_aba(driver, "Parecer"):
            dados_parecer = processar_aba_parecer(driver, numero_proposta)
            dados_consolidados.update(dados_parecer)

    except Exception as e:
        print(f"Erro geral ao processar a proposta {numero_proposta}: {e}")
        dados_consolidados["Erro"] = str(e)

    return dados_consolidados

# Processar todas as propostas
def processar_todas_propostas():
    """
    Processa todas as propostas presentes na planilha de entrada.
    """
    caminho_planilha_entrada = r"C:\Users\diego.brito\OneDrive - Ministério do Desenvolvimento e Assistência Social\Power BI\Python\propostas_iniciais.xlsx"
    caminho_planilha_saida = r"C:\Users\diego.brito\OneDrive - Ministério do Desenvolvimento e Assistência Social\Power BI\Python\Consulta Transferegov Consolidado.xlsx"

    if not os.path.exists(caminho_planilha_entrada):
        print(f"Erro: O arquivo de entrada '{caminho_planilha_entrada}' não foi encontrado.")
        return

    driver = conectar_navegador_existente()

    try:
        planilha_propostas = pd.read_excel(caminho_planilha_entrada)
        for indice, linha in planilha_propostas.iterrows():
            numero_proposta = linha['NºProposta']
            informacoes_adicionais = linha.to_dict()

            print(f"=== Processando a proposta {numero_proposta} ({indice + 1}/{len(planilha_propostas)}) ===")
            dados_proposta = processar_proposta_completa(driver, numero_proposta, informacoes_adicionais)
            salvar_progresso(dados_proposta, caminho_planilha_saida)

    finally:
        driver.quit()
        print("Processamento concluído.")

# Executar o script
if __name__ == "__main__":
    processar_todas_propostas()
