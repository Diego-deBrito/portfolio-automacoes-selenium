import os
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import WebDriverException, NoSuchElementException, TimeoutException
from webdriver_manager.chrome import ChromeDriverManager

# Caminho do arquivo Excel
ARQUIVO_EXCEL = r"C:\Users\diego.brito\Downloads\robov1\Lista para Teste - Email.xlsx"

# XPath do campo de pesquisa
XPATH_CAMPO_PESQUISA = "//input[@type='text']"

# XPaths dos botões
XPATH_REABRIR_PROCESSO = "//a[contains(@onclick, 'reabrir')]/img"
XPATH_EMAIL = "//a[contains(@onclick, 'enviarEmailProcedimento')]/img"

# XPath da caixa de texto do e-mail
XPATH_CAIXA_TEXTO_EMAIL = "/html/body/div[1]/div/div/form[1]/div[6]/textarea"


# Tempo de espera padrão
TEMPO_ESPERA = 10


def conectar_navegador_existente():
    """Conecta ao navegador Chrome já aberto, utilizando a porta de depuração 9222."""
    try:
        chrome_options = webdriver.ChromeOptions()
        chrome_options.debugger_address = "localhost:9222"
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
        print("✅ Conectado ao navegador existente com sucesso.")
        return driver
    except WebDriverException as e:
        print(f"❌ Erro ao conectar ao navegador existente: {e}")
        return None


def extrair_numeros_propostas():
    """Extrai números de processo do arquivo Excel."""
    try:
        df = pd.read_excel(ARQUIVO_EXCEL, dtype=str)
        if "Processo" not in df.columns:
            print("❌ Coluna 'Processo' não encontrada no Excel.")
            return []
        return df["Processo"].dropna().values.tolist()
    except Exception as e:
        print(f"❌ Erro ao ler o Excel: {e}")
        return []

def listar_iframes(driver):
    """Lista todos os iframes na página e retorna a quantidade."""
    iframes = driver.find_elements(By.TAG_NAME, "iframe")
    print(f"🔍 Encontrados {len(iframes)} iframes na página.")
    return iframes


def trocar_para_iframe(driver):
    """Troca diretamente para o iframe específico 'ifrVisualizacao'."""
    try:
        iframe = WebDriverWait(driver, TEMPO_ESPERA).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="ifrVisualizacao"]'))
        )
        driver.switch_to.frame(iframe)
        print("✅ Mudança para o iframe 'ifrVisualizacao' realizada com sucesso.")
        return True
    except TimeoutException:
        print("❌ Erro: Iframe 'ifrVisualizacao' não encontrado. Continuando na página principal.")
        driver.switch_to.default_content()  # Garante que volta ao conteúdo principal
        return False



def pesquisar_processo(driver, numero_proposta):
    """Realiza a pesquisa pelo número do processo."""
    try:
        campo_pesquisa = WebDriverWait(driver, TEMPO_ESPERA).until(
            EC.presence_of_element_located((By.XPATH, XPATH_CAMPO_PESQUISA))
        )
        campo_pesquisa.clear()
        campo_pesquisa.send_keys(numero_proposta + Keys.RETURN)
        print(f"🔍 Pesquisa realizada para: {numero_proposta}")
    except (NoSuchElementException, TimeoutException):
        print(f"❌ Erro ao pesquisar {numero_proposta}")


def esperar_elemento_visivel(driver, xpath, descricao):
    """Aguarda até que um elemento esteja visível na tela antes de interagir."""
    try:
        elemento = WebDriverWait(driver, TEMPO_ESPERA).until(EC.presence_of_element_located((By.XPATH, xpath)))
        driver.execute_script("arguments[0].scrollIntoView();", elemento)
        time.sleep(2)  # Pequeno delay para garantir que carregou
        print(f"✅ Elemento '{descricao}' está visível.")
        return elemento
    except TimeoutException:
        print(f"⚠️ Elemento '{descricao}' não encontrado ou não visível.")
        return None


def clicar_botao(driver, xpath, descricao):
    """Clica em um botão caso ele esteja visível."""
    try:
        botao = WebDriverWait(driver, TEMPO_ESPERA).until(
            EC.element_to_be_clickable((By.XPATH, xpath))
        )
        botao.click()
        print(f"✅ Botão '{descricao}' clicado.")
        return True
    except TimeoutException:
        print(f"⚠️ Botão '{descricao}' não encontrado.")
        return False





def anexar_arquivo(driver, caminho_arquivo):
    """Seleciona um arquivo para upload no campo de anexo."""
    try:
        # Aguarda o campo de upload estar visível
        input_arquivo = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div/div/form[2]/div/input"))
        )
        print("✅ Campo de upload encontrado.")

        # Usa send_keys() para enviar o caminho do arquivo
        input_arquivo.send_keys(caminho_arquivo)
        print(f"✅ Arquivo '{caminho_arquivo}' anexado com sucesso.")

    except TimeoutException:
        print("❌ Erro: Campo de upload não encontrado.")



def anexar_multiplos_arquivos(driver, caminho_pasta, arquivos):
    """Anexa múltiplos arquivos no e-mail."""
    try:
        input_arquivo = WebDriverWait(driver, TEMPO_ESPERA).until(
            EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div/div/form[2]/div/input"))
        )

        for arquivo in arquivos:
            caminho_arquivo = os.path.join(caminho_pasta, arquivo)
            if os.path.exists(caminho_arquivo):
                input_arquivo.send_keys(caminho_arquivo)
                print(f"✅ Arquivo anexado: {arquivo}")
            else:
                print(f"❌ Arquivo não encontrado: {arquivo}")
    except TimeoutException:
        print("❌ Erro ao anexar arquivos.")

# 🔹 Caminho do arquivo Excel e nome da coluna do email
COLUNA_EMAIL = "email"

def extrair_email_do_excel():
    """Lê o arquivo Excel e retorna o primeiro e-mail encontrado na coluna."""
    try:
        df = pd.read_excel(ARQUIVO_EXCEL, dtype=str)  # Lendo como string para evitar erros
        if COLUNA_EMAIL not in df.columns:
            print(f"❌ A coluna '{COLUNA_EMAIL}' não foi encontrada no arquivo Excel.")
            return None

        email = df[COLUNA_EMAIL].dropna().iloc[0]  # Pega o primeiro email não vazio
        print(f"✅ E-mail extraído do Excel: {email}")
        return email
    except Exception as e:
        print(f"❌ Erro ao ler o arquivo Excel: {e}")
        return None



def preencher_destinatario_email(driver, email):
    """Preenche o campo de destinatário no pop-up e confirma o e-mail pressionando ENTER."""
    try:
        # 🔹 Aguarda o campo de e-mail estar visível e interativo
        campo_email = WebDriverWait(driver, TEMPO_ESPERA).until(
            EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div/div/form[1]/div[4]/p/div/ul/li/input"))
        )
        print("✅ Campo de destinatário encontrado!")

        # 🔹 Garante que o campo está interativo e remove restrições
        driver.execute_script("arguments[0].scrollIntoView();", campo_email)
        driver.execute_script("arguments[0].removeAttribute('readonly');", campo_email)

        # 🔹 Limpa o campo antes de inserir o e-mail
        campo_email.send_keys(Keys.CONTROL, "a", Keys.DELETE)
        time.sleep(1)

        # 🔹 Insere o e-mail
        campo_email.send_keys(email)
        time.sleep(1)  # Pequeno delay para garantir que o sistema reconheça o texto
        print(f"✅ Destinatário preenchido com: {email}")

        # 🔹 Pressiona ENTER para confirmar o destinatário
        campo_email.send_keys(Keys.RETURN)
        print("✅ ENTER pressionado para confirmar o destinatário.")

        # 🔹 Aguarda um pouco para verificar se o alerta aparece
        time.sleep(2)
        try:
            alert = driver.switch_to.alert
            alert_text = alert.text
            print(f"⚠️ Alerta detectado: {alert_text}")
            if "Nenhum destinatário para o email informado." in alert_text:
                alert.accept()
                print("✅ Alerta fechado. Verifique se o e-mail está correto.")
                return False  # Indica falha ao preencher o destinatário
        except:
            print("✅ Nenhum alerta detectado. E-mail parece válido.")

        return True  # Indica sucesso

    except TimeoutException:
        print("❌ Erro: Não foi possível encontrar o campo de e-mail no pop-up.")
        return False


def fechar_popup(driver):
    """Fecha o pop-up após preencher e anexar os arquivos."""
    try:
        botoes_fechar = driver.find_elements(By.XPATH, "//button[contains(text(), 'Fechar') or contains(text(), 'Cancelar')]")
        if botoes_fechar:
            botoes_fechar[0].click()
            print("✅ Pop-up fechado automaticamente.")
        else:
            print("⚠️ Nenhum botão de fechar encontrado.")
    except Exception as e:
        print(f"⚠️ Erro ao tentar fechar o pop-up: {e}")




def preencher_texto_email(driver):
    """Preenche os campos do e-mail, insere anexos e envia o e-mail automaticamente."""
    try:
        # Captura a aba principal antes de abrir o pop-up
        aba_principal = driver.current_window_handle

        # Aguarda até que o pop-up seja detectado
        WebDriverWait(driver, 10).until(lambda d: len(d.window_handles) > 1)

        # Troca para o pop-up
        for janela in driver.window_handles:
            if janela != aba_principal:
                driver.switch_to.window(janela)
                break

        print("✅ Pop-up de e-mail ativado.")

        # Aguarda a página carregar
        WebDriverWait(driver, TEMPO_ESPERA).until(
            EC.presence_of_element_located((By.TAG_NAME, "body"))
        )
        print("✅ Página do pop-up carregada.")

        # 🔹 Trocar para o iframe correto
        if trocar_para_iframe(driver):
            print("✅ Mudança para o iframe 'ifrVisualizacao' bem-sucedida.")
        else:
            print("⚠️ Nenhum iframe detectado. Continuando no contexto padrão.")

        # 🔹 Preencher o campo de destinatário (apenas uma vez)
        email = extrair_email_do_excel()
        if email:
            if not preencher_destinatario_email(driver, email):
                print("❌ Erro ao preencher o destinatário. Abortando envio.")
                return
        else:
            print("⚠️ Nenhum e-mail encontrado no Excel. Continuando...")

        # 🔹 Preencher o campo Assunto
        try:
            campo_assunto = WebDriverWait(driver, TEMPO_ESPERA).until(
                EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div/div/form[1]/div[6]/input[1]"))
            )
            campo_assunto.clear()
            campo_assunto.send_keys("Providências após publicação do Termo de Fomento no Diário Oficial da União.")
            print("✅ Assunto preenchido com sucesso.")
        except TimeoutException:
            print("❌ Erro ao localizar o campo Assunto.")

        # 🔹 Preencher a caixa de texto do e-mail
        caixa_texto = WebDriverWait(driver, TEMPO_ESPERA).until(
            EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div/div/form[1]/div[6]/textarea"))
        )
        caixa_texto.clear()
        caixa_texto.send_keys("""Prezado Dirigente da ,

Em relação ao Ofício nº 42/2025/MESP/SNEAELIS (SEI nº 16426691), encaminhado por email na manhã do dia 16 de janeiro de 2025, retifica-se o item 3:

O encaminhamento da documentação requerida deverá ser realizado por meio da Plataforma TransfereGov, exclusivamente na aba Plano de Trabalho - Anexos - Listar Anexos Propostas, sem a exclusão de documentos inseridos anteriormente.

Este e-mail não deve ser respondido diretamente. Dúvidas devem ser encaminhadas para gabinete.sneaelis@esporte.gov.br.

Atenciosamente,
Andrei Rodrigues
Assessor - SNEAELIS
gabinete.sneaelis@esporte.gov.br""")
        print("✅ Texto do e-mail inserido com sucesso.")

        # 🔹 Anexar os 3 arquivos
        anexar_multiplos_arquivos(driver, r"C:\Users\diego.brito\Downloads\doc_teste", [
            "Modelo de Projeto Tecnico - PROJETO.docx",
            "Modelo de Projeto Tecnico - EVENTO.docx",
            "Oficio n 42-2025-MESP-SNEAELIS.pdf"
        ])
        print("✅ Todos os anexos foram inseridos com sucesso.")

        # 🔹 Tentar clicar no botão "Enviar"
        try:
            botao_enviar = WebDriverWait(driver, TEMPO_ESPERA).until(
                EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div/div/div[3]/button[1]"))
            )
            botao_enviar.click()
            print("✅ Botão 'Enviar' clicado com sucesso!")
        except TimeoutException:
            print("⚠️ Botão 'Enviar' não clicável. Tentando via JavaScript...")
            try:
                botao_enviar = driver.find_element(By.XPATH, "/html/body/div[1]/div/div/div[3]/button[1]")
                driver.execute_script("arguments[0].click();", botao_enviar)
                print("✅ Botão 'Enviar' clicado via JavaScript!")
            except Exception as e:
                print(f"❌ Erro ao tentar clicar no botão de envio: {e}")

        # 🔹 Aguardar envio ser concluído antes de fechar pop-up
        time.sleep(3)  # Tempo extra para garantir que o envio foi processado

        # 🔹 Fechar o pop-up automaticamente se ainda estiver aberto
        try:
            botoes_fechar = driver.find_elements(By.XPATH, "//button[contains(text(), 'Fechar') or contains(text(), 'Cancelar')]")
            if botoes_fechar:
                botoes_fechar[0].click()
                print("✅ Pop-up fechado automaticamente.")
            else:
                print("⚠️ Nenhum botão de fechar encontrado. Tentando fechar via JavaScript...")
                driver.execute_script("window.close();")
        except Exception as e:
            print(f"⚠️ Erro ao tentar fechar o pop-up: {e}")

        # Retorna para a aba principal
        driver.switch_to.window(aba_principal)
        print("🔄 Retornando para a aba principal.")

    except TimeoutException:
        print("❌ Erro: O pop-up não carregou corretamente.")





def selecionar_arquivo(driver, caminho_arquivo):
    """Seleciona um arquivo no input de upload."""
    try:
        # Aguarda o campo de upload estar visível
        input_arquivo = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div/div/form[2]/div/input"))
        )

        # Usa send_keys() para enviar o caminho do arquivo
        input_arquivo.send_keys(caminho_arquivo)
        print(f"✅ Arquivo '{caminho_arquivo}' selecionado com sucesso.")

    except TimeoutException:
        print("❌ Erro: Campo de upload não encontrado.")






def verificar_e_clicar_botoes(driver):
    """Verifica e clica nos botões necessários dentro do iframe 'ifrVisualizacao'."""
    try:
        # 🔹 Troca para o iframe correto
        try:
            iframe = WebDriverWait(driver, TEMPO_ESPERA).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="ifrVisualizacao"]'))
            )
            driver.switch_to.frame(iframe)
            print("✅ Mudança para o iframe 'ifrVisualizacao' realizada com sucesso.")
        except TimeoutException:
            print("❌ Erro: Iframe 'ifrVisualizacao' não encontrado. Continuando na página principal.")
            driver.switch_to.default_content()
            return

        # 🔹 Clicar no botão "Reabrir Processo" (se existir)
        if clicar_botao(driver, XPATH_REABRIR_PROCESSO, "Reabrir Processo"):
            time.sleep(2)  # Tempo para a ação ser processada
            clicar_botao(driver, XPATH_EMAIL, "Abrir E-mail")  # Depois clica no e-mail
        else:
            print("⚠️ Botão 'Reabrir Processo' não encontrado. Tentando 'Abrir E-mail' diretamente...")
            clicar_botao(driver, XPATH_EMAIL, "Abrir E-mail")

        # 🔹 Preencher o e-mail após abrir o pop-up
        preencher_texto_email(driver)

    except Exception as e:
        print(f"❌ Erro ao verificar e clicar nos botões: {e}")


def alternar_para_nova_aba(driver):
    """Troca o controle do Selenium para a nova aba do e-mail."""
    try:
        time.sleep(3)  # Aguarda a nova aba abrir
        abas = driver.window_handles  # Lista todas as abas abertas
        print(f"🔍 Abas abertas no momento: {abas}")

        for aba in abas:
            driver.switch_to.window(aba)
            time.sleep(2)  # Pequeno delay para garantir a mudança
            url_atual = driver.current_url
            print(f"🔄 Verificando aba: {url_atual}")

            # Verifica se a aba corresponde à URL do e-mail
            if "procedimento_enviar_email" in url_atual:
                print("✅ Aba correta do e-mail encontrada!")
                return True

        print("⚠️ Nenhuma aba corresponde à página de e-mail. Verifique se a aba foi aberta corretamente.")
        return False

    except Exception as e:
        print(f"❌ Erro ao alternar para a nova aba: {e}")
        return False



def main():
    """Executa o fluxo de extração e pesquisa no site."""
    driver = conectar_navegador_existente()
    if not driver:
        return

    abas_originais = driver.window_handles  # Salva as abas abertas inicialmente

    numeros_propostas = extrair_numeros_propostas()
    if not numeros_propostas:
        print("❌ Nenhum número de proposta encontrado no Excel.")
        return

    try:
        for numero in numeros_propostas:
            pesquisar_processo(driver, numero)  # Pesquisa o número no site
            verificar_e_clicar_botoes(driver)   # Clica nos botões (inclusive no de e-mail)

            # Após preencher o e-mail, volta para a aba original para continuar o processo
            abas_atualizadas = driver.window_handles  # Atualiza a lista de abas abertas
            if len(abas_atualizadas) > len(abas_originais):
                driver.switch_to.window(abas_originais[0])  # Retorna para a aba original
                print("🔄 Retornando para a aba principal para processar a próxima proposta.")

    finally:
        print("✅ Processo finalizado.")
        driver.quit()  # Fecha o navegador ao final


if __name__ == "__main__":
    main()