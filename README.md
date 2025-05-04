# Envio de E-mails de Desconto de Aniversário

Este projeto tem como objetivo automatizar o envio de e-mails personalizados para clientes com desconto especial de aniversário. O script lê uma planilha Excel contendo informações dos clientes, como nome, e-mail e data de nascimento, e envia um e-mail de parabéns com um código de desconto exclusivo para cada cliente que faz aniversário no mês atual.

## Objetivos

- **Automatizar o envio de e-mails**: Enviar e-mails personalizados de aniversário com um desconto especial para clientes.
- **Integrar com planilha Excel**: Ler dados de clientes a partir de uma planilha Excel.
- **Envio de e-mails via SMTP**: Usar o servidor SMTP do Gmail para enviar e-mails de forma segura.

## Tecnologias Utilizadas

- **Python 3.x**
- **Biblioteca smtplib**: Para enviar os e-mails via SMTP.
- **Biblioteca pandas**: Para manipulação da planilha Excel.
- **Biblioteca email.mime**: Para construir e-mails com HTML.
- **Biblioteca datetime**: Para trabalhar com datas e comparar aniversários.

## Como Rodar o Projeto

1. Clone este repositório:
   ```bash
   git clone https://github.com/seu-usuario/envio-emails-aniversario.git
   cd envio-emails-aniversario
