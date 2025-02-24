from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import time

929834


# Tempo máximo de espera para os elementos carregarem
TEMPO_ESPERA = 10

# Função para conectar ao navegador já aberto
def conectar_navegador_existente():
    print("Passo 1: Conectando ao navegador existente na porta de depuração 9222...")
    options = webdriver.ChromeOptions()
    options.debugger_address = "localhost:9222"  # Porta do Chrome para depuração
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    print("Conexão ao navegador estabelecida.")
    return driver

# Função para capturar todos os números das páginas disponíveis
def obter_paginas_disponiveis(driver):
    try:
        paginas = WebDriverWait(driver, TEMPO_ESPERA).until(
            EC.presence_of_all_elements_located((By.XPATH, "//a[contains(@href, 'ConsultarNotasFiscais') and not(contains(text(), 'Próx'))]"))
        )
        return [pagina.text for pagina in paginas if pagina.text.isdigit()]
    except:
        return []




# Função para clicar no botão "Próximo" múltiplas vezes
def clicar_proximo(driver, vezes):
    """Clica no botão 'Próximo' o número necessário de vezes para alcançar a página correta."""
    for i in range(vezes):
        try:
            print(f"➡️ Clicando em 'Próximo' ({i+1}/{vezes})...")
            botao_proximo = WebDriverWait(driver, TEMPO_ESPERA).until(
                EC.element_to_be_clickable((By.XPATH, "//a[text()='Próx']"))
            )
            botao_proximo.click()
            time.sleep(2)  # Pequena pausa para evitar falhas
        except Exception as e:
            print(f"⚠️ Erro ao clicar no botão 'Próximo': {e}")
            break



# Função para processar todas as páginas e baixar notas fiscais
def baixar_notas_fiscais(driver):
    pagina_atual = 1  # Começa na primeira página
    while True:  # Loop até processar todas as páginas
        paginas = obter_paginas_disponiveis(driver)

        if not paginas:
            print("🚀 Nenhuma página encontrada! Finalizando...")
            break

        print(f"\n📄 Encontradas {len(paginas)} páginas: {paginas}")


        for pagina in paginas:
            print(f"🔹 Acessando página {pagina}...")
            try:
                link_pagina = WebDriverWait(driver, TEMPO_ESPERA).until(
                    EC.element_to_be_clickable((By.XPATH, f"//a[text()='{pagina}']"))
                )
                link_pagina.click()
                WebDriverWait(driver, TEMPO_ESPERA).until(
                    EC.presence_of_element_located((By.XPATH, "//tbody[@id='tbodyrow']"))
                )  # Aguarda o carregamento da tabela
            except Exception as e:
                print(f"⚠️ Erro ao clicar na página {pagina}: {e}")
                continue  # Pula para a próxima página se houver erro

            print("Passo 2: Buscando notas fiscais na tabela...")

            try:
                # Captura todas as linhas da tabela
                linhas = WebDriverWait(driver, TEMPO_ESPERA).until(
                    EC.presence_of_all_elements_located((By.XPATH, "//tbody[@id='tbodyrow']/tr"))
                )

                print(f"Encontradas {len(linhas)} linhas na tabela.")

                for i in range(len(linhas)):
                    # Atualizar a lista de linhas porque a página pode recarregar
                    linhas = WebDriverWait(driver, TEMPO_ESPERA).until(
                        EC.presence_of_all_elements_located((By.XPATH, "//tbody[@id='tbodyrow']/tr"))
                    )

                    # Captura o link dentro da coluna "Tipo", que deve ser "NOTA FISCAL"
                    tipo_doc = WebDriverWait(linhas[i], TEMPO_ESPERA).until(
                        EC.presence_of_element_located((By.XPATH, "./td[3]/a"))
                    ).text.strip()

                    if tipo_doc == "NOTA FISCAL":
                        print(f"\n🔹 Acessando Nota Fiscal {i + 1}...")

                        # Clicar no link para abrir os detalhes
                        link_nota = WebDriverWait(linhas[i], TEMPO_ESPERA).until(
                            EC.element_to_be_clickable((By.XPATH, "./td[3]/a"))
                        )
                        link_nota.click()

                        WebDriverWait(driver, TEMPO_ESPERA).until(
                            EC.presence_of_element_located((By.XPATH, "/html/body/div[3]/div[15]/div[4]/div[1]"))
                        )  # Aguarda o carregamento da página de detalhes

                        try:
                            # Clicar no botão de download
                            print("🔽 Baixando o arquivo...")
                            botao_download = WebDriverWait(driver, TEMPO_ESPERA).until(
                                EC.element_to_be_clickable((By.XPATH, "/html/body/div[3]/div[15]/div[4]/div[1]/div/form/table/tbody/tr[24]/td/div[1]/table/tbody/tr/td[3]/nobr/a"))
                            )
                            botao_download.click()
                        except Exception as e:
                            print(f"⚠️ Erro ao tentar baixar o arquivo: {e}")

                        # Voltar para a lista de documentos
                        print("🔙 Retornando para a lista...")
                        botao_voltar = WebDriverWait(driver, TEMPO_ESPERA).until(
                            EC.element_to_be_clickable((By.XPATH, "/html/body/div[3]/div[15]/div[4]/div[1]/div/form/table/tbody/tr[39]/td[2]/input"))
                        )
                        botao_voltar.click()

                        WebDriverWait(driver, TEMPO_ESPERA).until(
                            EC.presence_of_element_located((By.XPATH, "//tbody[@id='tbodyrow']"))
                        )  # Aguarda recarregamento da tabela

                        # Reabrir a mesma página para continuar
                        print(f"🔄 Voltando para a página {pagina}...")
                        try:
                            link_pagina = WebDriverWait(driver, TEMPO_ESPERA).until(
                                EC.element_to_be_clickable((By.XPATH, f"//a[text()='{pagina}']"))
                            )
                            link_pagina.click()
                            WebDriverWait(driver, TEMPO_ESPERA).until(
                                EC.presence_of_element_located((By.XPATH, "//tbody[@id='tbodyrow']"))
                            )
                        except Exception as e:
                            print(f"⚠️ Erro ao tentar recarregar a página {pagina}: {e}")
                            break  # Sai do loop se não conseguir voltar

            except Exception as e:
                print(f"⚠️ Erro geral ao processar a página {pagina}: {e}")

        # Tentar clicar no botão "Próximo" após todas as páginas numéricas serem processadas
        try:
            botao_proximo = WebDriverWait(driver, TEMPO_ESPERA).until(
                EC.element_to_be_clickable((By.XPATH, "//a[text()='Próx']"))
            )
            print("➡️ Avançando para a próxima série de páginas...")
            botao_proximo.click()
            WebDriverWait(driver, TEMPO_ESPERA).until(
                EC.presence_of_element_located((By.XPATH, "//tbody[@id='tbodyrow']"))
            )
        except:
            print("🚀 Todas as páginas foram processadas.")
            break  # Sai do loop quando não há mais páginas

# Execução do script
if __name__ == "__main__":
    driver = conectar_navegador_existente()
    baixar_notas_fiscais(driver)
    print("✅ Processo finalizado com sucesso!")
