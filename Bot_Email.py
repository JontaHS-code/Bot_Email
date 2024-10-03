import smtplib
import pandas as pd
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from dotenv import load_dotenv
import os

# Carregar variáveis do arquivo .env
load_dotenv()

# Pegando as credenciais do arquivo .env
REMETENTE = os.getenv('REMETENTE')
SENHA = os.getenv('SENHA')

# Função para enviar email
def enviar_email(destinatario, assunto, mensagem):
    msg = MIMEMultipart()
    msg['From'] = REMETENTE
    msg['To'] = destinatario
    msg['Subject'] = assunto
    msg.attach(MIMEText(mensagem, 'plain'))

    try:
        with smtplib.SMTP('smtp.gmail.com', 587) as servidor:
            servidor.starttls()
            servidor.login(REMETENTE, SENHA)
            servidor.sendmail(REMETENTE, destinatario, msg.as_string())
            print(f"Email enviado para {destinatario}")
    except Exception as e:
        print(f"Erro ao enviar email para {destinatario}: {e}")

# Função para ler a lista de emails do arquivo CSV ou Excel
def ler_emails(arquivo):
    if arquivo.endswith('.csv'):
        df = pd.read_csv(arquivo)
    elif arquivo.endswith('.xlsx'):
        df = pd.read_excel(arquivo)
    else:
        raise ValueError("Arquivo não suportado. Use CSV ou Excel.")
    
    # Considerando que a coluna se chama sempre "email"
    if 'email' not in df.columns:
        raise ValueError("A coluna 'email' não foi encontrada no arquivo.")
    
    return df['email'].tolist()

# Interface simples para coletar informações do usuário
def interface():
    assunto = input("Digite o assunto do email: ")
    mensagem = input("Digite a mensagem do email: ")

    arquivo = input("Digite o nome do arquivo (CSV ou Excel) com a lista de emails: ")

    # Lendo os emails do arquivo
    try:
        lista_emails = ler_emails(arquivo)
    except Exception as e:
        print(f"Erro ao ler o arquivo: {e}")
        return

    # Confirmar antes de enviar
    print(f"\nVocê está prestes a enviar o seguinte email para {len(lista_emails)} destinatários:")
    print(f"Assunto: {assunto}")
    print(f"Mensagem: {mensagem}")
    confirmar = input("\nDeseja continuar? (s/n): ")

    if confirmar.lower() == 's':
        for email in lista_emails:
            enviar_email(email, assunto, mensagem)
    else:
        print("Envio cancelado.")

if __name__ == "__main__":
    interface()
