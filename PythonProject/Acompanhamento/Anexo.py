import time
from datetime import datetime

import pandas as pd
import os
from selenium import webdriver
from selenium.common import TimeoutException
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# Caminhos dos arquivos
CAMINHO_ENTRADA = r"C:\Users\diego.brito\Downloads\robov1\CONTROLE DE PARCERIAS CGAP - Copia.xlsx"
CAMINHO_SAIDA = r"C:\Users\diego.brito\Downloads\robov1\saida.xlsx"


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



def ler_planilha_entrada():
    """Lê a planilha de entrada e retorna apenas os instrumentos com status 'ATIVOS TODOS'."""
    try:
        df = pd.read_excel(CAMINHO_ENTRADA, dtype=str)
        df.columns = df.columns.str.strip()  # Remove espaços extras dos nomes das colunas
    except Exception as erro:
        print(f"❌ Erro ao carregar a planilha de entrada: {erro}")
        exit()

    colunas_esperadas = ["Instrumento nº", "Técnico", "e-mail do Técnico", "Status"]

    # Verifica se todas as colunas esperadas estão presentes
    colunas_faltando = [col for col in colunas_esperadas if col not in df.columns]
    if colunas_faltando:
        raise ValueError(f"🚨 Erro: As colunas {colunas_faltando} não foram encontradas na planilha!")

    # Filtra apenas os instrumentos com Status == "ATIVOS TODOS"
    df_filtrado = df[df["Status"].str.upper() == "ATIVOS TODOS"]

    if df_filtrado.empty:
        print("⚠️ Nenhum instrumento com status 'ATIVOS TODOS' encontrado!")
        exit()

    return df_filtrado[colunas_esperadas]


def salvar_dado_extracao(numero_instrumento, tecnico, email, status, data_upload):
    """Salva os dados extraídos na planilha de saída."""
    try:
        colunas_necessarias = ["Instrumento nº", "Técnico", "e-mail do Técnico", "Status", "Data Upload"]

        if os.path.exists(CAMINHO_SAIDA):
            df_saida = pd.read_excel(CAMINHO_SAIDA, dtype=str)
        else:
            df_saida = pd.DataFrame(columns=colunas_necessarias)

        for coluna in colunas_necessarias:
            if coluna not in df_saida.columns:
                df_saida[coluna] = ""

        if numero_instrumento in df_saida["Instrumento nº"].values:
            df_saida.loc[df_saida["Instrumento nº"] == numero_instrumento, "Data Upload"] = data_upload
        else:
            novo_dado = pd.DataFrame([[numero_instrumento, tecnico, email, status, data_upload]], columns=colunas_necessarias)
            df_saida = pd.concat([df_saida, novo_dado], ignore_index=True)

        df_saida.to_excel(CAMINHO_SAIDA, index=False)
        print(f"✅ Dados salvos para {numero_instrumento}!")

    except PermissionError:
        print("❌ ERRO: O arquivo está aberto no Excel. Feche-o e tente novamente.")
    except Exception as erro:
        print(f"❌ ERRO ao salvar dados no Excel: {erro}")

def automatizar_navegacao(driver, numero_instrumento):
    """Realiza a automação do site seguindo os passos indicados."""
    wait = WebDriverWait(driver, 5)

    def clicar(xpath, descricao=""):
        """ Aguarda o elemento estar disponível e tenta clicar, se falhar usa JavaScript """
        try:
            elemento = wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
            wait.until(EC.visibility_of(elemento))
            wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))

            try:
                elemento.click()
            except:
                driver.execute_script("arguments[0].click();", elemento)

            print(f"✔ {descricao} (via JS se necessário)")
        except Exception as erro:
            print(f"⚠️ Erro ao clicar ({descricao}): {erro}")

    print(f"\n➡️ Processando instrumento {numero_instrumento}...")

    # Acessar menu principal
    clicar("/html/body/div[1]/div[3]/div[1]/div[1]/div[1]/div[4]", "Acessando menu principal")
    clicar("/html/body/div[1]/div[3]/div[2]/div[1]/div[1]/ul/li[6]/a", "Acessando pesquisa")

    # Inserir Número do Instrumento
    try:
        input_field = wait.until(EC.presence_of_element_located(
            (By.XPATH, "/html/body/div[3]/div[15]/div[3]/div/div/form/table/tbody/tr[2]/td[2]/input")))
        input_field.clear()
        input_field.send_keys(numero_instrumento)
        clicar("/html/body/div[3]/div[15]/div[3]/div/div/form/table/tbody/tr[2]/td[2]/span/input",
               "Pesquisando instrumento")
    except Exception as erro:
        print(f"⚠️ Erro ao inserir número do instrumento {numero_instrumento}: {erro}")
        return None

    # Selecionar primeiro resultado da pesquisa
    clicar("/html/body/div[3]/div[15]/div[3]/div[3]/table/tbody/tr/td/div/a", "Selecionando primeiro resultado")
    clicar("/html/body/div[3]/div[15]/div[1]/div/div[1]/a[2]/div", "Abrindo menu")
    clicar("/html/body/div[3]/div[15]/div[1]/div/div[2]/a[8]/div", "Acessando detalhes")
    clicar("/html/body/div[3]/div[15]/div[3]/div[1]/div/form/table/tbody/tr/td[2]/input[2]",
           "Visualizando data de upload")

    # 🔹 Capturar a data do último anexo
    data_upload_extraida = capturar_data_ultimo_anexo(driver)

    # 🔄 Voltar para a pesquisa inicial
    try:
        driver.find_element(By.XPATH, "/html/body/div[3]/div[2]/div[1]/a").click()
        print("🔄 Retornando para a pesquisa inicial...")
    except Exception:
        print("⚠️ Não foi possível voltar, tentando continuar...")

    return data_upload_extraida  # ✅ Agora retorna um único valor corretamente



# Função para capturar a data do último anexo
def capturar_data_ultimo_anexo(driver):
    """Captura a data mais recente da coluna 'Data Upload'."""
    try:
        tabela = driver.find_element(By.XPATH, "/html/body/div[3]/div[15]/div[4]/div/div[1]/form/div/div[1]/table")
        linhas = tabela.find_elements(By.TAG_NAME, "tr")

        if not linhas:
            print("⚠️ Nenhuma linha encontrada na tabela.")
            return None

        cabecalho = linhas[0].find_elements(By.TAG_NAME, "th")
        colunas_titulos = [coluna.text.strip() for coluna in cabecalho]

        try:
            indice_data_upload = colunas_titulos.index("Data Upload")  # Encontra a posição correta
        except ValueError:
            print("⚠️ A coluna 'Data Upload' não foi encontrada!")
            return None

        datas = []
        for linha in linhas[1:]:
            colunas = linha.find_elements(By.TAG_NAME, "td")
            if len(colunas) > indice_data_upload:
                data_texto = colunas[indice_data_upload].text.strip()
                try:
                    data_formatada = datetime.strptime(data_texto, "%d/%m/%Y")
                    datas.append(data_formatada)
                except ValueError:
                    continue

        return max(datas).strftime("%d/%m/%Y") if datas else None

    except Exception as erro:
        print(f"⚠️ Erro ao capturar data do último anexo: {erro}")
        return None


def main():
    driver = conectar_navegador_existente()
    df_entrada = ler_planilha_entrada()

    for _, linha in df_entrada.iterrows():
        data_mais_recente = automatizar_navegacao(driver, linha["Instrumento nº"])
        salvar_dado_extracao(linha["Instrumento nº"], linha["Técnico"], linha["e-mail do Técnico"], linha["Status"], data_mais_recente)

    print("\n✅ Processamento concluído!")

if __name__ == "__main__":
    main()