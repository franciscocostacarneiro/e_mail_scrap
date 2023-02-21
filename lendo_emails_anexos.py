from imbox import Imbox  # Importa a classe Imbox para acessar as mensagens do email
from datetime import datetime  # Importa a classe datetime para trabalhar com as datas
import pandas as pd  # Importa a biblioteca pandas para trabalhar com dados estruturados
import os  # Importa a biblioteca os para manipulação do sistema operacional
import csv  # Importa a biblioteca csv para trabalhar com arquivos CSV

# Define as credenciais para acessar a conta de email
username = 'cochico01@gmail.com'  # Define o nome do usuário do email
password = open('pass', 'r').read()  # Lê a senha do arquivo 'pass'
host = "imap.gmail.com"  # Define o servidor de email a ser acessado
download_folder = "anexos"  # Define o diretório onde os anexos serão salvos

# Cria uma instância da classe Imbox com as credenciais fornecidas
mail = Imbox(host, username=username, password=password, ssl=True, ssl_context=None, starttls=False)

# Obtém uma lista de todos os emails na caixa de entrada
messages = mail.messages()

# # Loop pelos emails na caixa de entrada
# for (uid, message) in messages:
#     message.subject # Imprime o assunto do email
#     message.body # Imprime o corpo do email
#     message.sent_from # Imprime o remetente do email
#     message.sent_to # Imprime o destinatário do email
#     message.cc # Imprime a lista de destinatários em cópia do email
#     message.headers # Imprime os cabeçalhos do email
#     pd.to_datetime(message.date) # Converte a data de envio do email para o formato datetime

# Loop pelos emails na caixa de entrada
with open('emails.csv', 'w', newline='', encoding='utf-8') as csvfile:
    writer = csv.writer(csvfile)  # Cria um objeto writer para trabalhar com o arquivo CSV
    writer.writerow(['Subject', 'Date'])  # Escreve a primeira linha do arquivo CSV

    for i, (uid, message) in enumerate(messages):
        try:
            sent_date = pd.to_datetime(message.date)  # Converte a data do email para um objeto datetime

            # Imprime o título do email e a data de envio
            print(f"{i + 1}/{len(messages)} - {message.subject} - {sent_date}")
            
            writer.writerow([message.subject, sent_date])  # Escreve as informações do email no arquivo CSV

            # Loop pelos anexos do email
            for attach in message.attachments:
                if attach.get('content') is not None:  # Verifica se o anexo tem conteúdo
                    file_name = attach.get('filename')  # Obtém o nome do arquivo
                    file_path = os.path.join(download_folder, file_name)  # Cria o caminho completo do arquivo

                    with open(file_path, "wb") as fp:
                        fp.write(attach.get('content').read())  # Escreve o conteúdo do anexo no arquivo
        except Exception as e:
            print(f"Erro ao processar email UID {uid}: {e}")  # Imprime uma mensagem de erro caso ocorra uma exceção

mail.logout()  # Fecha a conexão com o email