import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd

# Caminho do arquivo Excel
caminho_planilha = r"C:\Users\diego.brito\Downloads\robov1\saida_Anexos.xlsx"

# Função para conectar ao navegador já aberto
def conectar_navegador_existente():
    print("🔄 Conectando ao navegador na porta 9222...")
    options = webdriver.ChromeOptions()
    options.debugger_address = "localhost:9222"  # Porta do Chrome para depuração
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    driver.maximize_window()
    print("✅ Conexão estabelecida!")
    return driver

# Função para garantir que a página tenha carregado completamente
def esperar_pagina_carregar(driver, timeout=20):
    try:
        WebDriverWait(driver, timeout).until(lambda d: d.execute_script("return document.readyState") == "complete")

        # Aguarda o desaparecimento de um possível elemento de carregamento
        try:
            WebDriverWait(driver, timeout).until(
                EC.invisibility_of_element_located(
                    (By.XPATH, "//div[contains(@class, 'loading') or contains(@class, 'spinner')]"))
            )
        except Exception:
            pass

        print("✅ Página carregada completamente!")
    except Exception as e:
        print(f"⚠️ Tempo limite ao esperar a página carregar: {e}")

# Função para aguardar um elemento estar visível e clicável, com captura de exceções
def esperar_elemento(driver, xpath, timeout=15):
    try:
        return WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((By.XPATH, xpath)))
    except Exception:
        print(f"⚠️ Elemento não encontrado: {xpath}")
        return None



# Função para retornar ao menu lateral
def retornar_ao_menu(driver):
            try:
                menu = esperar_elemento(driver, "//button[@title='Menu']")
                if menu:
                    menu.click()

                aba_sec = esperar_elemento(driver,
                                           "/html/body/transferencia-especial-root/br-main-layout/div/div/div/div/br-side-menu/nav/div[3]/a/span[2]")
                if aba_sec:
                    aba_sec.click()
                    esperar_pagina_carregar(driver)
            except Exception as e:
                print(f"⚠️ Erro ao tentar voltar pelo menu: {e}")





# Função para extrair números da planilha
def extrair_numeros_planilha(caminho_arquivo):
    df = pd.read_excel(caminho_arquivo, dtype=str)
    if "Instrumento nº" in df.columns:
        numeros = df["Instrumento nº"].dropna().tolist()
        print(f"📄 {len(numeros)} números extraídos da planilha.")
        return numeros
    else:
        print("⚠️ ERRO: Coluna 'Instrumento nº' não encontrada na planilha!")
        return []

# Inicia o driver do Chrome
driver = conectar_navegador_existente()
esperar_pagina_carregar(driver)

# Lê os números da planilha
numeros_processos = extrair_numeros_planilha(caminho_planilha)

try:
    # Acessar menu lateral
    menu = esperar_elemento(driver, "//button[@title='Menu']")
    if menu:
        menu.click()

    aba_sec = esperar_elemento(driver, "/html/body/transferencia-especial-root/br-main-layout/div/div/div/div/br-side-menu/nav/div[3]/a/span[2]")
    if aba_sec:
        aba_sec.click()
        esperar_pagina_carregar(driver)

    # Iterar sobre cada número da planilha
    for numero_processo in numeros_processos:
        print(f"🔍 Pesquisando número do processo: {numero_processo}")


        filtro = esperar_elemento(driver, "//i[@class='fas fa-filter']")
        if filtro:
            filtro.click()
            esperar_pagina_carregar(driver)

        input_field = esperar_elemento(driver, "/html/body/transferencia-especial-root/br-main-layout/div/div/div/main/transferencia-especial-main/transferencia-plano-acao/transferencia-plano-acao-consulta/br-table/div/br-fieldset/fieldset/div[2]/form/div[1]/div[2]/br-input/div/div/input")
        if input_field:
            input_field.clear()
            input_field.send_keys(numero_processo)
            input_field.send_keys(Keys.ENTER)
            esperar_pagina_carregar(driver)

        botao_detalhar = esperar_elemento(driver, "/html/body/transferencia-especial-root/br-main-layout/div/div/div/main/transferencia-especial-main/transferencia-plano-acao/transferencia-plano-acao-consulta/br-table/div/ngx-datatable/div/datatable-body/datatable-selection/datatable-scroller/datatable-row-wrapper/datatable-body-row/div[2]/datatable-body-cell[8]/div/div/button/i")
        if botao_detalhar:
            botao_detalhar.click()
            esperar_pagina_carregar(driver)

        botao_analise = esperar_elemento(driver, "/html/body/transferencia-especial-root/br-main-layout/div/div/div/main/transferencia-especial-main/transferencia-plano-acao/transferencia-cadastro/br-tab-set/div/nav/ul/li[4]/button/span")
        if botao_analise:
            driver.execute_script("arguments[0].scrollIntoView();", botao_analise)
            driver.execute_script("arguments[0].click();", botao_analise)
            esperar_pagina_carregar(driver)

        # Tentar encontrar e clicar no botão "Adicionar"
        botao_adicionar = esperar_elemento(driver,
                                           "/html/body/transferencia-especial-root/br-main-layout/div/div/div/main/transferencia-especial-main/transferencia-plano-acao/transferencia-cadastro/br-tab-set/div/nav/transferencia-consulta/div[1]/div[1]/button")

        if botao_adicionar:
            driver.execute_script("arguments[0].scrollIntoView();", botao_adicionar)
            driver.execute_script("arguments[0].click();", botao_adicionar)
            esperar_pagina_carregar(driver)
            print("✅ Botão 'Adicionar' clicado com sucesso!")

        else:
            print("⚠️ Botão 'Adicionar' não encontrado. Pulando para o próximo processo...")
            retornar_ao_menu(driver)  # 🔄 Garante que volte ao menu antes de continuar
            continue  # Pula para o próximo processo no loop




        # Função para selecionar "Ministério do Esporte"
        def selecionar_ministerio(driver):
            try:
                print("🔍 Tentando selecionar 'Ministério do Esporte'...")
                campo_ministerio = esperar_elemento(driver,
                                                    "/html/body/transferencia-especial-root/br-main-layout/div/div/div/main/transferencia-especial-main/transferencia-plano-acao/transferencia-cadastro/br-tab-set/div/nav/transferencia-cadastro/form/div[1]/div/br-select/div/div/div[1]/ng-select/div/div/div[2]/input")

                if campo_ministerio:
                    driver.execute_script("arguments[0].scrollIntoView();", campo_ministerio)
                    campo_ministerio.click()
                    time.sleep(1)
                    campo_ministerio.send_keys("Ministério do Esporte")
                    time.sleep(2)
                    campo_ministerio.send_keys(Keys.DOWN)
                    campo_ministerio.send_keys(Keys.ENTER)
                    esperar_pagina_carregar(driver)
                    print("✅ 'Ministério do Esporte' selecionado com sucesso!")
                else:
                    print("⚠️ ERRO: Campo de seleção do Ministério não encontrado.")

            except Exception as e:
                print(f"❌ ERRO ao selecionar 'Ministério do Esporte': {e}")



        # Chama a função no fluxo do código
        selecionar_ministerio(driver)

        # Localiza o campo de texto para o parecer
        campo_texto = esperar_elemento(driver,
                                       "/html/body/transferencia-especial-root/br-main-layout/div/div/div/main/transferencia-especial-main/transferencia-plano-acao/transferencia-cadastro/br-tab-set/div/nav/transferencia-cadastro/form/div[4]/div/br-textarea/div/div[1]/div/textarea")

        if campo_texto:
            campo_texto.clear()

            # Novo TEXTO 2 atualizado
            texto_parecer = f"""Declaramos, em resposta ao determinado pelo egrégio Supremo Tribunal Federal nos autos da ADPF-DF nº 854/2024 e em observância à Portaria Conjunta MGI/MF nº 2/2025, que nenhuma das metas mencionadas na Plataforma Transferegov, oriundas da análise deste Plano de Trabalho, decorrente de indicação de “Emenda Pix”, tem aderência à função e subfunção orçamentárias designadas a este Ministério.

            Assim, orientamos ao contemplado ente federado que adeque o Plano de Ação, modificando a função e a subfunção orçamentárias direcionando-o ao órgão setorial responsável.

            Número do processo: {numero_processo}"""

        campo_situacao = esperar_elemento(driver, "/html/body/transferencia-especial-root/br-main-layout/div/div/div/main/transferencia-especial-main/transferencia-plano-acao/transferencia-cadastro/br-tab-set/div/nav/transferencia-cadastro/form/div[2]/div[1]/br-select/div/div/div[1]/ng-select/ng-dropdown-panel/div/div[2]/div[3]/span")
        if campo_situacao:
            campo_situacao.send_keys("Não se aplica")
            campo_situacao.send_keys(Keys.ENTER)
            print("✅ Situação 'Não se aplica' selecionada com sucesso!")
            esperar_pagina_carregar(driver)

            # XPath do botão que precisa ser clicado antes de continuar
            xpath_botao_proximo = "/html/body/transferencia-especial-root/br-main-layout/div/div/div/main/transferencia-especial-main/transferencia-plano-acao/transferencia-cadastro/br-tab-set/div/nav/transferencia-cadastro/div/button[2]"

            # Tenta clicar no botão "Próximo" antes de continuar
            botao_proximo = esperar_elemento(driver, xpath_botao_proximo)

            if botao_proximo:
                driver.execute_script("arguments[0].scrollIntoView();",
                                      botao_proximo)  # Garante que o botão esteja visível
                botao_proximo.click()  # Clica no botão
                esperar_pagina_carregar(driver)  # Aguarda a página carregar completamente
                print("✅ Botão 'Próximo' clicado com sucesso! Continuando para o próximo processo...")
            else:
                print("⚠️ Erro: Botão 'Próximo' não encontrado. Continuando mesmo assim.")

            # Retornar ao menu antes de iniciar o próximo loop
            retornar_ao_menu(driver)



except Exception as e:
    print(f'❌ Erro durante o processo: {e}')

print("🚀 Automação finalizada com sucesso!")
