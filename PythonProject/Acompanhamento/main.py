import time
import pandas as pd
from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# 🛠 Conectar ao navegador já aberto
def conectar_navegador_existente():
    options = webdriver.ChromeOptions()
    options.debugger_address = "localhost:9222"
    try:
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
        print("✅ Conectado ao navegador existente!")
        return driver
    except Exception as erro:
        print(f"❌ Erro ao conectar ao navegador: {erro}")
        exit()

# 📂 Ler planilha de entrada respeitando o filtro "Status"
def ler_planilha(arquivo=r"C:\Users\diego.brito\Downloads\robov1\CONTROLE DE PARCERIAS CGAP.xlsx"):
    df = pd.read_excel(arquivo, engine="openpyxl")
    df = df[df["Status"] == "ATIVOS TODOS"]
    return df


# 📤 Salvar planilha de saída em tempo real
def salvar_planilha(df, arquivo=r"C:\Users\diego.brito\Downloads\robov1\resultado_main.xlsx"):
    df.to_excel(arquivo, index=False)
    print(f"📂 Planilha atualizada: {arquivo}")




# 🔍 Espera até que um elemento esteja visível
def esperar_elemento(driver, xpath, tempo=10):
    return WebDriverWait(driver, tempo).until(EC.presence_of_element_located((By.XPATH, xpath)))

# 🚀 Fluxo principal do robô
def executar_robo():
    driver = conectar_navegador_existente()
    df_entrada = ler_planilha()
    dados_saida = []

    for index, row in df_entrada.iterrows():
        tecnico = row["Técnico"]
        instrumento = str(row["Instrumento nº"])
        email_tecnico = row["e-mail do Técnico"]

        print(f"🔎 Pesquisando Instrumento nº {instrumento}...")

        try:
            esperar_elemento(driver, "/html/body/div[1]/div[3]/div[1]/div[1]/div[1]/div[4]").click()
            esperar_elemento(driver, "/html[1]/body[1]/div[1]/div[3]/div[2]/div[1]/div[1]/ul[1]/li[6]/a[1]").click()
            campo_pesquisa = esperar_elemento(driver, "/html[1]/body[1]/div[3]/div[15]/div[3]/div[1]/div[1]/form[1]/table[1]/tbody[1]/tr[2]/td[2]/input[1]")
            campo_pesquisa.clear()
            campo_pesquisa.send_keys(instrumento)
            esperar_elemento(driver, "/html[1]/body[1]/div[3]/div[15]/div[3]/div[1]/div[1]/form[1]/table[1]/tbody[1]/tr[2]/td[2]/span[1]/input[1]").click()
            time.sleep(2)
        except:
            print(f"⚠️ Instrumento {instrumento} não encontrado.")
            dados_saida.append([tecnico, instrumento, email_tecnico, "Instrumento não encontrado", "", "", ""])
            df_saida = pd.DataFrame(dados_saida,
                                    columns=["Técnico", "Instrumento nº", "E-mail", "Data Ajustes", "Data TA",
                                             "Rendimento de Aplicação", "Último Upload"])
            salvar_planilha(df_saida)
            continue

        try:
            esperar_elemento(driver, "/html[1]/body[1]/div[3]/div[15]/div[3]/div[3]/table[1]/tbody[1]/tr[1]/td[1]/div[1]/a[1]").click()
        except:
            continue

        # 🔹 Aba Ajustes do PT
        data_ajustes = "Sem informação"  # Valor padrão caso a situação não seja encontrada

        try:
            print("📂 Acessando Aba Ajustes do PT...")

            # 🏷️ Passo 1: Clicar na aba Ajustes do PT
            esperar_elemento(driver,
                             "/html[1]/body[1]/div[3]/div[15]/div[1]/div[1]/div[1]/a[6]/div[1]/span[1]/span[1]").click()

            # 🏷️ Passo 2: Acessar a segunda aba dentro de Ajustes do PT
            esperar_elemento(driver,
                             "/html[1]/body[1]/div[3]/div[15]/div[1]/div[1]/div[2]/a[16]/div[1]/span[1]/span[1]").click()

            # 📌 Passo 3: Esperar a tabela carregar
            tabela_ajustes = esperar_elemento(driver, "/html/body/div[3]/div[15]/div[3]/div[2]/div[2]/table")

            # 📌 Passo 4: Coletar todas as linhas da tabela
            linhas = tabela_ajustes.find_elements(By.TAG_NAME, "tr")

            encontrou_situacao = False  # Variável para verificar se encontrou a situação correta

            for linha in linhas[1:]:  # Ignorar o cabeçalho
                colunas = linha.find_elements(By.TAG_NAME, "td")

                if len(colunas) >= 4:  # Garantir que há pelo menos 4 colunas na linha
                    situacao_texto = colunas[3].text.strip()  # Coluna 4 contém a situação

                    # 📌 Passo 5: Verifica se a situação é "Em análise" ou "Em Análise (aguardando Parecer)"
                    if "Em análise" in situacao_texto or "Em Análise (aguardando Parecer)" in situacao_texto:
                        encontrou_situacao = True  # Marcamos que encontramos a situação correta

                        print(f"🔍 Situação encontrada: {situacao_texto}. Clicando em 'Detalhar'...")

                        # 📌 Passo 6: Clicar no botão "Detalhar" (última coluna da linha)
                        botao_detalhar = colunas[3].find_element(By.TAG_NAME, "a")
                        botao_detalhar.click()

                        # 📌 Passo 7: Esperar e extrair a Data de Solicitação
                        data_ajustes = esperar_elemento(driver,
                                                        "/html/body/div[3]/div[15]/div[4]/div[1]/div/form/table/tbody/tr[13]").text

                        print(f"📅 Data da solicitação extraída: {data_ajustes}")

                        break  # Já encontramos uma linha válida, podemos parar o loop

            # 📌 Passo 8: Se não encontrar a situação, registrar "Sem informação"
            if not encontrou_situacao:
                print("⚠️ Nenhuma situação válida encontrada. Registrando como 'Sem informação'.")

        except Exception as e:
            print(f"⚠️ Erro ao processar Aba Ajustes do PT: {e}")
            pass

        # 🔹 Aba TA
        data_ta = "Sem informação"  # Valor padrão caso a situação não seja encontrada

        try:
            print("📂 Acessando Aba TA...")

            # 🏷️ Passo 1: Acessar a Aba TA
            esperar_elemento(driver,
                             "/html[1]/body[1]/div[3]/div[15]/div[1]/div[1]/div[1]/a[6]/div[1]/span[1]/span[1]").click()
            esperar_elemento(driver,
                             "/html[1]/body[1]/div[3]/div[15]/div[1]/div[1]/div[2]/a[20]/div[1]/span[1]/span[1]").click()

            # 📌 Passo 2: Esperar a tabela carregar
            tabela_ta = esperar_elemento(driver, "/html/body/div[3]/div[15]/div[4]/div/form/div/div[3]/table")

            # 📌 Passo 3: Coletar todas as linhas da tabela
            linhas = tabela_ta.find_elements(By.TAG_NAME, "tr")

            encontrou_situacao = False  # Variável para verificar se encontrou a situação correta

            for linha in linhas[1:]:  # Ignorar o cabeçalho
                colunas = linha.find_elements(By.TAG_NAME, "td")

                if len(colunas) >= 4:  # Garantir que há pelo menos 4 colunas na linha
                    situacao_texto = colunas[3].text.strip()  # Coluna 4 contém a situação

                    # 📌 Passo 4: Verifica se a situação é "Em análise" ou "Em Análise (aguardando Parecer)"
                    if "Em análise" in situacao_texto or "Em Análise (aguardando Parecer)" in situacao_texto:
                        encontrou_situacao = True  # Marcamos que encontramos a situação correta

                        print(f"🔍 Situação encontrada: {situacao_texto}. Clicando em 'Detalhar'...")

                        # 📌 Passo 5: Clicar no botão "Detalhar" (última coluna da linha)
                        botao_detalhar = colunas[3].find_element(By.TAG_NAME, "a")
                        botao_detalhar.click()

                        # 📌 Passo 6: Esperar e extrair a Data de Solicitação
                        data_ta = esperar_elemento(driver,
                                                   "/html/body/div[3]/div[15]/div[3]/div[1]/div/form/table/tbody/tr[13]").text

                        print(f"📅 Data da solicitação extraída: {data_ta}")

                        break  # Já encontramos uma linha válida, podemos parar o loop

            # 📌 Passo 7: Se não encontrar a situação, registrar "Sem informação"
            if not encontrou_situacao:
                print("⚠️ Nenhuma situação válida encontrada. Registrando como 'Sem informação'.")

        except Exception as e:
            print(f"⚠️ Erro ao processar Aba TA: {e}")

            pass

        # 🔹 Última Aba - Verificar registros
        status_registro = "Sem registro"
        try:
            esperar_elemento(driver, "/html[1]/body[1]/div[3]/div[15]/div[1]/div[1]/div[2]/a[28]/div[1]/span[1]/span[1]").click()
            status_registro = esperar_elemento(driver, "/html/body/div[3]/div[15]/div[7]").text
        except:
            pass

        from datetime import datetime

        # 🔹 Coletar Data Mais Recente na Coluna Data Upload
        data_upload = "Sem registro"

        try:
            print("📂 Acessando aba de anexos para buscar Data Upload...")

            # 🏷️ Acessar aba correta antes de buscar os anexos
            esperar_elemento(driver,
                             "/html[1]/body[1]/div[3]/div[15]/div[1]/div[1]/div[1]/a[2]/div[1]/span[1]/span[1]").click()
            esperar_elemento(driver,
                             "/html[1]/body[1]/div[3]/div[15]/div[1]/div[1]/div[2]/a[8]/div[1]/span[1]/span[1]").click()
            esperar_elemento(driver,
                             "/html[1]/body[1]/div[3]/div[15]/div[3]/div[1]/div[1]/form[1]/table[1]/tbody[1]/tr[1]/td[2]/input[2]").click()

            # 📌 Aguardar a tabela carregar
            tabela_upload = esperar_elemento(driver,
                                             "/html/body/div[3]/div[15]/div[4]/div/div[1]/form/div/div[1]/table")

            # 📌 Verificar se a coluna "Data Upload" realmente existe
            coluna_data_upload = esperar_elemento(driver,
                                                  "/html/body/div[3]/div[15]/div[4]/div/div[1]/form/div/div[1]/table/thead/tr/th[3]").text

            if "Data Upload" in coluna_data_upload:
                print("✅ Coluna 'Data Upload' encontrada!")

                # 📌 Coletar todas as linhas da tabela
                linhas = tabela_upload.find_elements(By.TAG_NAME, "tr")

                # 📌 Lista para armazenar as datas extraídas da coluna "Data Upload"
                datas_upload = []

                for linha in linhas[1:]:  # Ignorar cabeçalho
                    colunas = linha.find_elements(By.TAG_NAME, "td")

                    if len(colunas) >= 3:  # Garante que há pelo menos 3 colunas
                        data_texto = colunas[2].text.strip()  # Pegando a 3ª coluna (th[3] → td[3])

                        if data_texto:
                            try:
                                data_formatada = datetime.strptime(data_texto,
                                                                   "%d/%m/%Y")  # Ajuste para o formato correto
                                datas_upload.append(data_formatada)
                            except ValueError:
                                print(f"⚠️ Data inválida ignorada: {data_texto}")

                # 📌 Se houver datas, pegar a mais recente
                if datas_upload:
                    data_upload = max(datas_upload).strftime("%d/%m/%Y")  # Converter de volta para string
                    print(f"📅 Data mais recente na coluna Data Upload: {data_upload}")
                else:
                    print("⚠️ Nenhuma data válida encontrada na coluna Data Upload.")

            else:
                print("⚠️ O nome da coluna não corresponde a 'Data Upload'. Verifique o XPath!")

        except Exception as e:
            print(f"⚠️ Erro ao coletar Data Upload: {e}")

        # 🔹 Aba Esclarecimentos
        data_esclarecimento = "Sem informação"
        anexo_esclarecimento = "Nenhum anexo encontrado"

        try:
            print("📂 Acessando Aba Esclarecimentos...")

            # 🏷️ Passo 1: Acessar a aba Esclarecimentos
            aba_esclarecimentos = esperar_elemento(driver, "/html[1]/body[1]/div[3]/div[2]/div[4]/div[1]/div[7]")
            driver.execute_script("arguments[0].scrollIntoView();", aba_esclarecimentos)
            aba_esclarecimentos.click()

            aba_esclarecimentos_secundaria = esperar_elemento(driver,
                                                              "/html[1]/body[1]/div[3]/div[2]/div[5]/div[1]/div[2]/ul[1]/li[1]/a[1]")
            driver.execute_script("arguments[0].scrollIntoView();", aba_esclarecimentos_secundaria)
            aba_esclarecimentos_secundaria.click()

            # 📌 Passo 2: Ir até a última página antes de buscar a Data mais recente
            try:
                while True:
                    paginacao = esperar_elemento(driver,
                                                 "/html/body/div[3]/div[15]/div[4]/div[1]/div/form/table/tbody/tr[6]/td/div[1]/span[2]")
                    botoes_pagina = paginacao.find_elements(By.TAG_NAME, "a")  # Todos os botões de página

                    # Encontrar a página atual
                    pagina_atual = paginacao.find_element(By.XPATH,
                                                          ".//span[contains(@class, 'pagina-selecionada')]").text

                    if botoes_pagina:
                        ultimo_botao = botoes_pagina[-1]  # Último botão da paginação

                        if pagina_atual != ultimo_botao.text:
                            print(f"➡️ Indo para a página {ultimo_botao.text}...")
                            driver.execute_script("arguments[0].scrollIntoView();", ultimo_botao)
                            ultimo_botao.click()
                            time.sleep(3)  # Esperar carregar a página
                        else:
                            print("✅ Já estamos na última página.")
                            break  # Sai do loop se já estiver na última página
                    else:
                        print("⚠️ Não há paginação visível.")
                        break  # Sai do loop se não houver botões de página

            except Exception as e:
                print(f"⚠️ Erro ao navegar pela paginação: {e}")

            # 📌 Passo 3: Encontrar a Data de Solicitação mais recente
            try:
                print("🔍 Buscando a Data de Solicitação mais recente...")

                tabela_esclarecimentos = esperar_elemento(driver,
                                                          "/html/body/div[3]/div[15]/div[4]/div[1]/div/form/table/tbody/tr[6]/td/div[1]/table")
                linhas = tabela_esclarecimentos.find_elements(By.TAG_NAME, "tr")

                data_mais_recente = None
                botao_detalhar_associado = None

                for linha in linhas:
                    colunas = linha.find_elements(By.TAG_NAME, "td")

                    if len(colunas) >= 7:  # Garante que há pelo menos 7 colunas
                        data_texto = colunas[6].text.strip()  # Coluna 7 contém a Data de Solicitação
                        botao_detalhar = colunas[6].find_element(By.TAG_NAME, "a")  # Botão Detalhar

                        if data_texto:
                            try:
                                data_formatada = datetime.strptime(data_texto, "%d/%m/%Y")

                                # Se for a data mais recente, atualiza
                                if data_mais_recente is None or data_formatada > data_mais_recente:
                                    data_mais_recente = data_formatada
                                    botao_detalhar_associado = botao_detalhar

                            except ValueError:
                                print(f"⚠️ Data inválida ignorada: {data_texto}")

                # 📌 Passo 4: Clicar no botão "Detalhar" correspondente à data mais recente
                if botao_detalhar_associado:
                    data_esclarecimento = data_mais_recente.strftime("%d/%m/%Y")
                    print(f"📅 Data de Esclarecimento mais recente: {data_esclarecimento}")

                    driver.execute_script("arguments[0].scrollIntoView();", botao_detalhar_associado)
                    ActionChains(driver).move_to_element(botao_detalhar_associado).perform()
                    botao_detalhar_associado.click()
                    print("✅ Clicou no botão 'Detalhar'!")
                    time.sleep(3)  # Pequena espera para carregar a página dos anexos
                else:
                    print("⚠️ Nenhuma Data de Solicitação encontrada.")

            except Exception as e:
                print(f"⚠️ Erro ao buscar a Data de Solicitação: {e}")

        except Exception as e:
            print(f"⚠️ Erro ao processar Aba Esclarecimentos: {e}")

        # 🔹 Retornar os dados para adicionar ao Excel
        print(f"📄 Data Esclarecimento: {data_esclarecimento}, Anexo Esclarecimento: {anexo_esclarecimento}")

        # 📝 Adicionar dados na saída (incluindo Esclarecimento e Anexo Esclarecimento)
        dados_saida.append([
            tecnico, instrumento, email_tecnico, data_ajustes, data_ta, status_registro, data_upload,
            data_esclarecimento, anexo_esclarecimento  # Novas colunas adicionadas
        ])

        # 📤 Atualizar planilha em tempo real com as novas colunas
        df_saida = pd.DataFrame(dados_saida, columns=[
            "Técnico", "Instrumento nº", "E-mail", "Data Ajustes", "Data TA",
            "Rendimento de Aplicação", "Último Upload", "Data Esclarecimento", "Anexo Esclarecimento"  # Colunas atualizadas
        ])

        salvar_planilha(df_saida)

        # 🔄 Retornar ao Menu Principal antes de iniciar o próximo instrumento
        try:
            print("🔄 Retornando ao Menu Principal...")
            botao_menu_principal = esperar_elemento(driver, "/html/body/div[3]/div[2]/div[1]/a")
            driver.execute_script("arguments[0].scrollIntoView();", botao_menu_principal)
            ActionChains(driver).move_to_element(botao_menu_principal).click().perform()
            time.sleep(2)  # Pequena pausa para garantir carregamento
        except Exception as e:
            print(f"⚠️ Erro ao tentar voltar ao menu principal: {e}")

        print("📂 Planilha gerada com sucesso!")


executar_robo()
