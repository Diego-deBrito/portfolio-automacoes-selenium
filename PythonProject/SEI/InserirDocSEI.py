from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import WebDriverException
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
import time


def conectar_navegador_existente():
    """
    Conecta ao navegador Chrome já aberto utilizando a porta de depuração 9222.
    """
    try:
        print("Tentando conectar ao navegador na porta 9222...")

        opcoes_navegador = webdriver.ChromeOptions()
        opcoes_navegador.debugger_address = "localhost:9222"

        navegador = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=opcoes_navegador)
        print("Conectado ao navegador existente com sucesso.")
        return navegador

    except WebDriverException as erro:
        print(f"Erro ao conectar ao navegador existente: {erro}")
        return None


def inserir_documento(navegador, caminho_arquivo, descricao_documento):
    """
    Realiza a inserção de um documento no SEI já logado.

    :param navegador: Objeto do navegador conectado.
    :param caminho_arquivo: Caminho local do arquivo a ser inserido.
    :param descricao_documento: Descrição do documento a ser adicionada.
    """
    try:
        print("Tentando inserir documento no SEI...")

        # Localizar o botão ou campo para adicionar documentos
        botao_adicionar = navegador.find_element(By.ID, "btnAdicionarDocumento")  # Substitua pelo ID correto
        botao_adicionar.click()

        time.sleep(2)  # Aguarda o modal de upload abrir

        # Localizar o campo de upload e enviar o arquivo
        campo_upload = navegador.find_element(By.ID, "inputUploadArquivo")  # Substitua pelo ID correto
        campo_upload.send_keys(caminho_arquivo)

        # Preencher a descrição do documento
        campo_descricao = navegador.find_element(By.ID, "inputDescricao")  # Substitua pelo ID correto
        campo_descricao.send_keys(descricao_documento)

        # Confirmar a inserção
        botao_confirmar = navegador.find_element(By.ID, "btnConfirmarInsercao")  # Substitua pelo ID correto
        botao_confirmar.click()

        time.sleep(3)  # Aguarda a operação concluir
        print("Documento inserido com sucesso!")

    except Exception as e:
        print(f"Erro ao inserir documento: {e}")


# Configuração inicial
if __name__ == "__main__":
    # Conectar ao navegador existente
    navegador = conectar_navegador_existente()

    if navegador:
        # Caminho do arquivo e descrição
        caminho_arquivo = r"C:\Users\diego.brito\OneDrive - Ministério do Desenvolvimento e Assistência Social\Power BI\Mala direta Ofício nova.xlms"  # Substitua pelo caminho real
        descricao_documento = r"C:\Users\diego.brito\OneDrive - Ministério do Desenvolvimento e Assistência Social\Power BI\retorno.xlms"

        # Inserir documento no SEI
        inserir_documento(navegador, caminho_arquivo, descricao_documento)
    else:
        print("Não foi possível iniciar o robô devido à falha na conexão com o navegador.")
