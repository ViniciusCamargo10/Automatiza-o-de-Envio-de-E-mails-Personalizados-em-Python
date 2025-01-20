import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import pandas as pd
from datetime import datetime
import math

# Função para enviar o e-mail
def enviar_email(destinatario, nome_cliente):
    # Credenciais do remetente
    usuario = "seu@email.com"  # Substitua por seu e-mail
    senha = "senha forte app google"  # Substitua pela sua senha de aplicativo do Gmail

    # Configurações do e-mail
    assunto = " Desconto Especial de Aniversário - 10% em qualquer produto! 🎁"
    corpo_email = f"""
    <html>
    <head>
        <style>
            body {{
                font-family: 'Arial', sans-serif;
                background-color: #f7f7f7;
                padding: 20px;
                color: #333;
            }}
            .container {{
                background-color: #ffffff;
                padding: 30px;
                border-radius: 10px;
                box-shadow: 0 4px 10px rgba(0, 0, 0, 0.1);
                width: 100%;
                max-width: 600px;
                margin: 0 auto;
                text-align: center;
            }}
            .header {{
                font-size: 26px;
                color: #ff6f61;
                margin-bottom: 20px;
            }}
            .content {{
                font-size: 16px;
                line-height: 1.6;
            }}
            .footer {{
                margin-top: 30px;
                font-size: 14px;
                color: #777;
            }}
            .button {{
                background-color: #ff6f61;
                color: white;
                padding: 10px 20px;
                text-align: center;
                font-size: 16px;
                text-decoration: none;
                border-radius: 5px;
                display: inline-block;
                margin-top: 20px;
            }}
            .button:hover {{
                background-color: #ff4f3b;
            }}
        </style>
    </head>
    <body>
        <div class="container">
            <div class="header">
                <p> Parabéns, {nome_cliente}! 🎂</p>
            </div>
            <div class="content">
                <p>Queremos comemorar o seu mês aniversário com você! 🎉</p>
                <p>Para tornar seu mês ainda mais incrível, você tem direito a um <strong>desconto exclusivo de 10%</strong> em qualquer produto da nossa loja.</p>
                <p>Aproveite essa oportunidade para adquirir aquele item que você tanto queria!</p>
                <a href="https://www.seuSite.com/" class="button">Aproveite seu desconto agora</a>
            </div>
            <div class="footer">
                <p>Com carinho,</p>
                <p><strong>Sua Loja</strong></p>
            </div>
        </div>
    </body>
    </html>
    """

    # Criar a estrutura do e-mail
    msg = MIMEMultipart()
    msg['From'] = usuario
    msg['To'] = destinatario
    msg['Subject'] = assunto

    # Anexando o corpo do e-mail
    msg.attach(MIMEText(corpo_email, 'html'))

    try:
        # Conectando ao servidor SMTP do Gmail
        with smtplib.SMTP('smtp.gmail.com', 587) as servidor:
            servidor.starttls()  # Começando uma conexão segura
            servidor.login(usuario, senha)  # Fazendo login no servidor SMTP

            # Enviando o e-mail
            servidor.sendmail(usuario, destinatario, msg.as_string())
            print(f"E-mail enviado com sucesso para {nome_cliente}!")

    except Exception as e:
        print(f"Erro ao enviar e-mail para {nome_cliente}: {e}")

# Função para processar a planilha
def enviar_emails_para_clientes():
    # Carregar a planilha Excel
    arquivo_excel = r"C:\SeuPacoteAqui"  # Usando r para evitar problemas com as barras invertidas
    df = pd.read_excel(arquivo_excel)

    # Obter o mês atual
    mes_atual = datetime.now().month

    # Iterar sobre as linhas do dataframe
    for index, row in df.iterrows():
        # Extrair o nome, e-mail, e as colunas de aniversário (dividido)
        nome_cliente = row['cv1-nomecliv']
        email_cliente = row['cv1-emaicliv']
        dia_nasc = row['aw-dia-nasc']
        mes_nasc = row['aw-mes-nasc']
        ano_nasc = row['aw-ano-nasc']

        # Verificar se o e-mail é válido (não é NaN, não é nulo e é uma string)
        if pd.isna(email_cliente) or not isinstance(email_cliente, str) or email_cliente.strip() == "":
            print(f"Endereço de e-mail inválido para {nome_cliente}. Pulando para o próximo cliente...")
            continue  # Pula para o próximo cliente

        # Compor a data de aniversário
        try:
            aniversario = datetime(ano_nasc, mes_nasc, dia_nasc)
        except ValueError:
            # Caso haja um erro na data (por exemplo, dia 30 em fevereiro), ignorar essa linha
            print(f"Data inválida para {nome_cliente}. Ignorando...")
            continue

        # Verificar se o mês de aniversário é o mesmo que o mês atual
        if aniversario.month == mes_atual:
            enviar_email(email_cliente, nome_cliente)

# Executar a função para enviar os e-mails
enviar_emails_para_clientes()
