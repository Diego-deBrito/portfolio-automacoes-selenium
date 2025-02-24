import time
from datetime import datetime, timedelta
import pandas as pd
import os
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# 🔹 Caminhos dos arquivos
CAMINHO_ENTRADA = r"C:\Users\diego.brito\Downloads\robov1\Proposta Cláusula Suspensiva 2024.xlsx"
CAMINHO_SAIDA = r"C:\Users\diego.brito\Downloads\robov1\propostaEsclarecimento.xlsx"


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


def registrar_resultado(numero_proposta, status, mensagem=""):
    """Registra o resultado da execução na planilha de saída."""
    dados = {
        "Data e Hora": [datetime.now().strftime("%d/%m/%Y %H:%M:%S")],
        "Número da Proposta": [numero_proposta],
        "Status": [status],
        "Mensagem de Erro": [mensagem]
    }

    df_novo = pd.DataFrame(dados)

    # Se a planilha já existe, adicionamos novos dados sem sobrescrever
    if os.path.exists(CAMINHO_SAIDA):
        df_existente = pd.read_excel(CAMINHO_SAIDA)
        df_final = pd.concat([df_existente, df_novo], ignore_index=True)
    else:
        df_final = df_novo

    df_final.to_excel(CAMINHO_SAIDA, index=False)
    print(f"📄 Registro salvo para a proposta {numero_proposta}: {status}")


def clicar_enviar_solicitacao(driver):
    """Clica no botão 'Enviar Solicitação' corretamente, tratando possíveis erros."""
    try:
        # 🔹 XPath atualizado para encontrar o botão correto
        XPATH_BOTAO_ENVIAR = "//a[contains(@class, 'buttonLink') and contains(text(), 'Enviar Solicitação')]"

        # 🔹 Esperar o botão estar visível e interativo
        botao = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, XPATH_BOTAO_ENVIAR))
        )

        # 🔹 Tentar clicar normalmente
        botao.click()
        print("✔ Botão 'Enviar Solicitação' clicado com sucesso!")

    except Exception as erro:
        print(f"⚠️ Erro ao clicar normalmente. Tentando via JavaScript... {erro}")

        try:
            # 🔹 Forçar clique via JavaScript
            botao = driver.find_element(By.XPATH, XPATH_BOTAO_ENVIAR)
            driver.execute_script("arguments[0].click();", botao)
            print("✔ Botão 'Enviar Solicitação' clicado via JavaScript!")

        except Exception as js_erro:
            print(f"❌ Erro ao clicar no botão 'Enviar Solicitação': {js_erro}")
            driver.save_screenshot("erro_enviar_solicitacao.png")  # Captura print da tela



def ler_planilha_entrada():
    """Lê a planilha de entrada e retorna os números das propostas."""
    try:
        df = pd.read_excel(CAMINHO_ENTRADA, dtype=str)
        df.columns = df.columns.str.strip()
    except Exception as erro:
        print(f"❌ Erro ao carregar a planilha: {erro}")
        exit()

    if "PROPOSTA" not in df.columns:
        raise ValueError("🚨 Coluna 'PROPOSTA' não encontrada na planilha!")

    return df["PROPOSTA"].dropna().tolist()


def clicar(driver, xpath, descricao):
    """Tenta clicar em um elemento, se necessário via JavaScript."""
    wait = WebDriverWait(driver, 10)  # Aumenta tempo de espera
    try:
        # Espera o elemento estar presente antes de verificar se é clicável
        elemento = wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
        elemento = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
        elemento.click()
        print(f"✔ {descricao}")
    except Exception as erro:
        try:
            # Se não for clicável diretamente, tenta clicar via JavaScript
            elemento = driver.find_element(By.XPATH, xpath)
            driver.execute_script("arguments[0].click();", elemento)
            print(f"✔ {descricao} (via JS)")
        except Exception as js_erro:
            print(f"❌ Erro ao clicar em {descricao}: {js_erro}")
            driver.save_screenshot(f"erro_{descricao}.png")  # Captura tela do erro


def inserir_texto(driver, xpath, texto, descricao):
    """Insere texto em um campo."""
    wait = WebDriverWait(driver, 5)
    try:
        campo = wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
        campo.clear()
        campo.send_keys(texto)
        print(f"✔ {descricao}: {texto}")
    except Exception as erro:
        print(f"⚠️ Erro ao inserir {descricao}: {erro}")



def esperar_e_mudar_para_iframe(driver, xpath_iframe, descricao):
    """Aguarda o iframe estar disponível e muda para ele."""
    try:
        wait = WebDriverWait(driver, 10)
        iframe = wait.until(EC.presence_of_element_located((By.XPATH, xpath_iframe)))
        driver.switch_to.frame(iframe)
        print(f"✔ {descricao}")
    except Exception as erro:
        print(f"❌ Erro ao mudar para {descricao}: {erro}")
        driver.save_screenshot(f"erro_iframe_{descricao}.png")




def esperar_e_clicar(driver, xpath, descricao):
    """Espera o elemento estar visível e clicável antes de interagir."""
    try:
        wait = WebDriverWait(driver, 10)
        elemento = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
        elemento.click()
        print(f"✔ {descricao}")
    except Exception as erro:
        print(f"❌ Erro ao clicar em {descricao}: {erro}")
        driver.save_screenshot(f"erro_{descricao}.png")  # Captura tela do erro



def clicar_com_js(driver, xpath, descricao):
    """Força um clique no elemento via JavaScript."""
    try:
        elemento = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, xpath)))
        driver.execute_script("arguments[0].click();", elemento)
        print(f"✔ {descricao} (via JavaScript)")
    except Exception as erro:
        print(f"❌ Erro ao clicar em {descricao}: {erro}")
        driver.save_screenshot(f"erro_{descricao}.png")


def anexar_arquivo(driver, caminho_arquivo):
    """Anexa um arquivo WinRAR na opção de anexar documento."""
    try:
        # Verifica se o arquivo existe antes de tentar anexar
        if not os.path.exists(caminho_arquivo):
            raise FileNotFoundError(f"❌ Arquivo não encontrado: {caminho_arquivo}")

        # Passo 1: Inserir o caminho do arquivo no campo de upload
        campo_upload = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div[3]/form/fieldset/div/table/tbody/tr[11]/td[2]/input"))
        )
        campo_upload.send_keys(caminho_arquivo)
        print(f"✔ Arquivo anexado: {caminho_arquivo}")

        # Passo 2: Clicar no botão de anexar
        clicar(driver, "/html/body/div[1]/div[3]/form/fieldset/div/table/tbody/tr[11]/td[2]/span/button/span", "Anexando documento")

    except Exception as erro:
        print(f"❌ Erro ao anexar o arquivo: {erro}")


def automatizar_navegacao(driver, numero_proposta):
    """Executa os passos da automação para cada proposta."""
    print(f"\n➡️ Processando proposta {numero_proposta}...")

    try:
        # 🔹 Passo 1-12: Fluxo inicial
        clicar(driver, "/html/body/div[1]/div[3]/div[1]/div[1]/div[1]/div[3]", "Acessando menu principal")
        clicar(driver, "/html/body/div[1]/div[3]/div[2]/div[1]/div[1]/ul/li[3]/a", "Acessando pesquisa")
        inserir_texto(driver, "/html/body/div[3]/div[15]/div[3]/div/div/form/table/tbody/tr[1]/td[2]/input",
                      numero_proposta, "Inserindo número da proposta")
        clicar(driver, "/html/body/div[3]/div[15]/div[3]/div/div/form/table/tbody/tr[1]/td[2]/span/input",
               "Pesquisando proposta")
        clicar(driver, "/html/body/div[3]/div[15]/div[3]/div[3]/table/tbody/tr/td[1]/div/a",
               "Selecionando proposta encontrada")
        clicar(driver, "/html/body/div[3]/div[2]/div[4]/div/div[7]", "Acessando menu secundário")
        clicar(driver, "/html/body/div[3]/div[2]/div[5]/div/div[2]/ul/li[1]/a", "Acessando opção adicional")
        clicar(driver, "/html/body/div[3]/div[15]/div[4]/div[1]/div/form/table/tbody/tr[7]/td/input",
               "Confirmando ação")

        # 🔹 Passo 9: Preencher data
        data_futura = (datetime.today() + timedelta(days=15)).strftime("%d/%m/%Y")
        inserir_texto(driver, "/html/body/div[1]/div[3]/form/fieldset/div/table/tbody/tr[9]/td[2]/span/input",
                      data_futura, "Inserindo data futura")

        # 🔹 Passo 10: Inserir TEXTO OFICIAL
        texto_oficial = """
        Prezado Convenente,

        Considerando que a parceria foi celebrada em Cláusula Suspensiva, de acordo com a Subcláusula Primeira, da Cláusula Terceira, do Termo de Convênio assinado entre as partes, cujo prazo encerra 9 (nove) meses após a assinatura do instrumento, solicitamos a inserção na aba “Projeto Básico/Termo de Referência”, no Portal Transferegov, de:

        - Projeto Técnico Pedagógico;
        - Planilha de Custos devidamente preenchida e assinada (modelo em anexo);
        - Comprovações dos custos - 3 (três) cotações para cada item previsto no Plano de Trabalho; e
        - Termo de Referência.

        Esclarecemos que a documentação mencionada é necessária ao saneamento das pendências da Cláusula Suspensiva, visando possibilitar a entidade o início dos trâmites do processo licitatório referente à aquisição e contratação dos bens e serviços pactuados no Plano de Trabalho.

        Para atendimento da demanda, assinalamos o prazo de 15 (quinze) dias, a contar desta solicitação.

        Permanecemos à disposição para prestar os esclarecimentos necessários.
            """

        inserir_texto(driver, "/html/body/div[1]/div[3]/form/fieldset/div/table/tbody/tr[10]/td[2]/textarea",
                      texto_oficial, "Inserindo texto oficial")


        inserir_texto(driver, "/html/body/div[1]/div[3]/form/fieldset/div/table/tbody/tr[10]/td[2]/textarea",
                      texto_oficial, "Inserindo texto oficial")

        # 🔹 Passo 11: Anexar arquivo
        CAMINHO_ARQUIVO_RAR = r"C:\Users\diego.brito\Downloads\robov1\Planilha de Custos.zip"
        anexar_arquivo(driver, CAMINHO_ARQUIVO_RAR)

        # 🔹 Passo 12: Salvar processo
        clicar(driver, '/html/body/div[1]/div[3]/form/div[6]/div/button[1]', 'Botão salvar')
        clicar(driver, '/html/body/div[3]/div[15]/div[4]/div/div/form/table/tbody/tr[16]/td[2]/input[3]',
               'Continuar a salvar')

        # 🔹 Passo 13: Mudar para o iframe (se existir)
        XPATH_IFRAME = "/html/body/div[3]/div[15]/div[4]/div[1]/iframe"
        esperar_e_mudar_para_iframe(driver, XPATH_IFRAME, "Mudando para iframe do botão Enviar")

        # 🔹 Passo 14: Clicar no botão "Enviar Solicitação"
        clicar_enviar_solicitacao(driver)

        # 🔹 Voltar ao conteúdo principal
        driver.switch_to.default_content()
        print("✔ Voltando para o conteúdo principal")

        # 🔹 Passo 15: Voltar para a tela inicial
        clicar(driver, "/html/body/div[3]/div[2]/div[1]/a", "Voltando para a tela inicial")

        # 🔹 Registrar sucesso na planilha
        registrar_resultado(numero_proposta, "Sucesso")

        print("✅ Proposta processada com sucesso!")

    except Exception as erro:
        print(f"❌ Erro ao processar proposta {numero_proposta}: {erro}")
        driver.save_screenshot(f"erro_{numero_proposta}.png")

        # 🔹 Registrar erro na planilha
        registrar_resultado(numero_proposta, "Falha", str(erro))

def main():
    driver = conectar_navegador_existente()
    propostas = ler_planilha_entrada()

    for proposta in propostas:
        automatizar_navegacao(driver, proposta)

    print("\n✅ Processamento concluído!")


if __name__ == "__main__":
    main()
