import pandas as pd
import smtplib
import email.message


df = pd.read_excel(r'C:\Users\david\PycharmProjects\AutoEmail\Vendas.xlsx')

# Calculando o faturamento de cada loja
faturamento = df[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
faturamento = faturamento.sort_values(by='Valor Final', ascending=False)

# Calculando a quantidade de vendas de cada loja
quantidade = df[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
quantidade = quantidade.sort_values(by='Quantidade', ascending=False)

# Calculando o ticket medio de cada loja
ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Medio'})
ticket_medio = ticket_medio.sort_values(by='Ticket Medio', ascending=False)


def enviar_email(resumo_loja, loja):

    server = smtplib.SMTP('smtp.gmail.com:587')
    email_content = f'''
      <p>Prezados, segue o relatório </p>
      {resumo_loja.to_html()}
      <p>Qualquer dúvida estou à disposição.</p>'''

    msg = email.message.Message()
    msg['Subject'] = f'Relatório de faturamento - Loja: {loja}'

    msg['From'] = 'remetente'
    msg['To'] = 'destinatario'
    password = 'senha do email'
    msg.add_header('Content-Type', 'text/html')
    msg.set_payload(email_content)

    s = smtplib.SMTP('smtp.gmail.com: 587')
    s.starttls()
    # login no email
    s.login(msg['From'], password)
    s.sendmail(msg['From'], [msg['To']], msg.as_string().encode('utf-8'))


tabela_diretoria = faturamento.join(quantidade).join(ticket_medio)
enviar_email(tabela_diretoria, 'Todas as lojas')


lojas = df['ID Loja'].unique()

for loja in lojas:

    tabela_loja = df.loc[df['ID Loja'] == loja, ['ID Loja', 'Quantidade', 'Valor Final']]
    resumo_loja = tabela_loja.groupby('ID Loja').sum()
    resumo_loja['Ticket Medio'] = resumo_loja['Valor Final'] / resumo_loja['Quantidade']
    enviar_email(resumo_loja, loja)
