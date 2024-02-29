import os
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# importar base de dados
caminho_arquivo_excell = "AnaliseDadosERelatorio_RPA/Dados/Vendas.xlsx"

if os.path.exists(caminho_arquivo_excell):
    tabela_vendas = pd.read_excel(caminho_arquivo_excell)
else:
    print("Arquivo não encontrado")

# visualizar base de dados
pd.set_option('display.max_columns', None)
print(tabela_vendas)
print("-"*50)

# faturamento por loja
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(faturamento)
print("-"*50)

# Quantidade de produto vendido por loja
produtoVendidoPorLoja = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(produtoVendidoPorLoja)
print("-"*50)

# Faturamento divido por quantidade - Ticket medio - PS.: Adicionar o .to_Frame() no final para transformar em uma tabela
ticketMedio = (faturamento['Valor Final'] / produtoVendidoPorLoja['Quantidade']).to_frame()
ticketMedio = ticketMedio.rename(columns={0: 'Ticket Médio'})
print(ticketMedio)
print("-"*50)

# enviar email com relatorio

# Dados Config
username = "brunocazelatto@sis-it.com"
password = "LT$Hohyq"
mail_from = username
mail_to = "bcazelatto@gmail.com"
mail_subject = "Relatório de Vendas"
mail_body = f'''
<html>
<head></head>
<body>
    <p>Prezados,</P>

    <p>Segue o relatório de vendas por cada loja.</P>

    <p><strong>Faturamento:</strong></P>
    {faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

    <p><strong>Quantidades Vendidas:</strong></P>
    {produtoVendidoPorLoja.to_html()}

    <p><strong>Ticket Médio dos Produtos de cada loja:</strong></P>
    {ticketMedio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}

    <p>Qualquer dúvida estou a disposição.</P>

    <p>Atenciosamente.</P>
</body>
</html>
'''

mimemsg = MIMEMultipart()
mimemsg['From'] = mail_from
mimemsg['To'] = mail_to
mimemsg['Subject'] = mail_subject
mimemsg.attach(MIMEText(mail_body, 'html'))

connection = smtplib.SMTP('smtp.office365.com', 587)
connection.starttls()
connection.login(username, password)
connection.send_message(mimemsg)
connection.quit()
