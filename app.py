from flask import Flask, request, render_template, redirect, url_for, flash
import smtplib
import pandas as pd
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from dotenv import load_dotenv
import os

# Carregar variáveis do arquivo .env
load_dotenv()

# Pegando as credenciais do arquivo .env
REMETENTE = os.getenv('REMETENTE')
SENHA = os.getenv('SENHA')

app = Flask(__name__)
app.secret_key = os.getenv('SECRET_KEY')

# Função para enviar email
def enviar_email(destinatario, assunto, mensagem, cc=None, anexos=[]):
    msg = MIMEMultipart()
    msg['From'] = REMETENTE
    msg['To'] = destinatario  # Para referência, mas não será usado no envio
    msg['Subject'] = assunto
    msg.attach(MIMEText(mensagem, 'plain'))

    # Adicionando os anexos
    for anexo in anexos:
        try:
            with open(anexo, 'rb') as file:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(file.read())
                encoders.encode_base64(part)
                part.add_header('Content-Disposition', f'attachment; filename={os.path.basename(anexo)}')
                msg.attach(part)
        except Exception as e:
            print(f"Erro ao anexar {anexo}: {e}")

    try:
        with smtplib.SMTP('smtp.gmail.com', 587) as servidor:
            servidor.starttls()
            servidor.login(REMETENTE, SENHA)
            # Enviando o email como BCC
            servidor.sendmail(REMETENTE, destinatario, msg.as_string())
            if cc:
                msg['Cc'] = cc
                servidor.sendmail(REMETENTE, cc, msg.as_string())
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

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        assunto = request.form['assunto']
        mensagem = request.form['mensagem']
        cc = request.form['cc'] if request.form['cc'] else None

        # Lendo os emails do arquivo
        arquivo = request.files['arquivo']
        arquivo.save(arquivo.filename)  # Salvar temporariamente
        lista_emails = ler_emails(arquivo.filename)

        # Processando anexos
        anexos = []
        if 'anexos' in request.files:
            for file in request.files.getlist('anexos'):
                if file:
                    # Salvar o arquivo anexo
                    filepath = os.path.join(os.path.dirname(__file__), file.filename)
                    file.save(filepath)
                    anexos.append(filepath)

        # Enviar emails como BCC
        for email in lista_emails:
            enviar_email(email, assunto, mensagem, cc, anexos)
        
        # Remover arquivos temporários após o envio
        for anexo in anexos:
            try:
                os.remove(anexo)  # Remove os anexos salvos
            except Exception as e:
                print(f"Erro ao remover anexo {anexo}: {e}")

        try:
            os.remove(arquivo.filename)  # Remove o arquivo CSV ou Excel
        except Exception as e:
            print(f"Erro ao remover arquivo {arquivo.filename}: {e}")

        # Redirecionar para evitar reenvio do formulário
        return redirect(url_for('index', success_message="Emails enviados com sucesso!"))

    success_message = request.args.get('success_message')
    return render_template('index.html', success_message=success_message)

if __name__ == "__main__":
    app.run(debug=True)
