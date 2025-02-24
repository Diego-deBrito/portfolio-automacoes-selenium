import pandas as pd
import win32com.client as win32
from datetime import datetime

# 📂 Caminho da planilha de origem
file_path = r"C:\Users\diego.brito\Downloads\robov1\saida.xlsx"

# 📌 Carregar a planilha
df = pd.read_excel(file_path, engine='openpyxl')

# 📌 Remover espaços extras nos nomes das colunas
df.columns = df.columns.str.strip()

# 📌 Selecionar colunas desejadas
colunas_necessarias = [
    'Instrumento', 'Número Ajustes', 'Situação P.Trabalho', 'Situação TA', 'Número TA',
    'Aba Anexos', 'Data Esclarecimento', 'Resposta Esclarecimento', 'Técnico', 'e-mail do Técnico'
]
df_filtered = df[colunas_necessarias]

# 📌 Lidar com valores nan ou ".0" e convertê-los adequadamente
df_filtered = df_filtered.fillna("")

# 📌 Função para verificar e ajustar as situações
def ajustar_situacoes(row):
    if not ((row['Situação P.Trabalho'] == "Em Análise (aguardando parecer)") or
            (row['Situação TA'] in ["Cadastrada", "Em Análise"])):
        row['Situação P.Trabalho'] = ""
        row['Situação TA'] = ""
        row['Número TA'] = ""
        row['Aba Anexos'] = ""
        row['Data Esclarecimento'] = ""
        row['Resposta Esclarecimento'] = ""
    return row

# 📌 Aplicar a função de ajuste
df_filtered = df_filtered.apply(ajustar_situacoes, axis=1)

# 📌 Remover valores NaN nas colunas essenciais
df_filtered = df_filtered.dropna(subset=['e-mail do Técnico'])

# 📌 Função para enviar e-mails
def send_email(subject, body, recipient):
    try:
        outlook = win32.Dispatch('Outlook.Application')
        mail = outlook.CreateItem(0)
        mail.Subject = subject
        mail.HTMLBody = body
        mail.To = recipient
        mail.Send()
        print(f"📧 E-mail enviado para {recipient}")
    except Exception as e:
        print(f"⚠️ Erro ao enviar e-mail para {recipient}: {e}")

# 📌 Função para gerar tabela HTML no e-mail
def generate_email_table(data):
    if not data:
        return "<p>Não há dados para exibir.</p>"

    table_html = """
    <html>
      <head>
        <style>
          table { width: 100%; border-collapse: collapse; }
          th, td { border: 1px solid black; padding: 8px; text-align: left; }
          th { background-color: #f2f2f2; }
        </style>
      </head>
      <body>
        <h2>Relatório de Ajustes</h2>
        <table>
          <tr>
            <th>Técnico</th>
            <th>Instrumento</th>
            <th>Situação P.Trabalho</th>
            <th>Situação TA</th>
            <th>Aba Anexos</th>
            <th>Resposta Esclarecimento</th>
          </tr>
    """
    for row in data:
        table_html += "<tr>" + "".join(f"<td>{item}</td>" for item in row) + "</tr>"

    table_html += "</table></body></html>"
    return table_html

# 📌 Preparar dados para e-mail
from collections import defaultdict

# 📌 Agrupar dados por técnico e e-mail
grouped_data = defaultdict(list)
for _, row in df_filtered.iterrows():
    grouped_data[(row['Técnico'], row['e-mail do Técnico'])].append([
        row['Técnico'], row['Instrumento'], row['Situação P.Trabalho'], row['Situação TA'],
        row['Aba Anexos'], row['Resposta Esclarecimento']
    ])

# 📧 Enviar e-mails
for (técnico, email_do_tecnico), data in grouped_data.items():
    # Verifica se o técnico é "A DISTRIBUIR - SUSPENSIVA"
    if técnico == "A DISTRIBUIR - SUSPENSIVA":
        email_do_tecnico = ("diego.brito"
                            "@esporte.gov.br")

    # Verifica se o e-mail do técnico está vazio
    if not email_do_tecnico or pd.isna(email_do_tecnico):
        print(f"⚠️ E-mail do técnico {técnico} está vazio ou inválido. Pulando...")
        continue

    # Gera o corpo do e-mail
    email_body = generate_email_table(data)

    # Envia o e-mail
send_email(
    subject=f"Relatório de Ajustes - {técnico}",
    body=(f"<p><strong>Prezado(a) {técnico},</strong></p>"
          f"<p>Segue abaixo o relatório de ajustes para os instrumentos sob sua responsabilidade.</p>"
          f"{email_body}"
          "<p>Atenciosamente,<br><strong>Equipe de Automação</strong></p>"
          "<p>🤖</p>"),
    recipient=email_do_tecnico  # Usa o e-mail do técnico diretamente
)

