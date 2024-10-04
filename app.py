import os
import smtplib
import pandas as pd
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from dotenv import load_dotenv
from flask import Flask, render_template, request, redirect, url_for, flash

# Configuração do Flask
app = Flask(__name__)
app.secret_key = 'secret-key'  # chave para as mensagens flash

# Carregar variáveis do arquivo .env
load_dotenv()
REMETENTE = os.getenv('REMETENTE')
SENHA = os.getenv('SENHA')

ALLOWED_EXTENSIONS = {'csv', 'xlsx'}

# Função para verificar se o arquivo é permitido
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# Função para enviar e-mail
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
    except Exception as e:
        print(f"Erro ao enviar e-mail para {destinatario}: {e}")

# Função para ler e-mails diretamente do arquivo carregado
def ler_emails(arquivo):
    if arquivo.filename.endswith('.csv'):
        df = pd.read_csv(arquivo)
    elif arquivo.filename.endswith('.xlsx'):
        df = pd.read_excel(arquivo)
    else:
        raise ValueError("Arquivo não suportado. Use CSV ou Excel.")

    if 'email' not in df.columns:
        raise ValueError("A coluna 'email' não foi encontrada no arquivo.")
    
    return df['email'].tolist()

# Rota principal
@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        assunto = request.form['assunto']
        mensagem = request.form['mensagem']
        arquivo = request.files['arquivo']

        if arquivo and allowed_file(arquivo.filename):
            try:
                lista_emails = ler_emails(arquivo)
            except Exception as e:
                flash(f"Erro ao ler o arquivo: {e}", 'error')
                return redirect(url_for('index'))

            for email in lista_emails:
                enviar_email(email, assunto, mensagem)
            
            flash(f"E-mails enviados com sucesso para {len(lista_emails)} destinatários!", 'success')
            return redirect(url_for('index'))
        else:
            flash("Arquivo não suportado. Envie um arquivo CSV ou Excel.", 'error')
            return redirect(url_for('index'))

    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)
