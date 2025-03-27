import pandas as pd
import  win32com.client as win32

# Essa é a lógica do programa.

# Importar a base de dados.
tabela_vendas = pd.read_excel('Vendas.xlsx')

# Visualizar a base de dados.
pd .set_option('display.max_columns', None)
print (tabela_vendas)

# Importar a base de dados.
tabela_vendas = pd.read_excel('Vendas.xlsx')

# Visualizar a base de dados.
pd.set_option('display.max_columns', None)
print(tabela_vendas)

# Faturamento por loja.
faturamento = tabela_vendas[['ID loja', 'Valor Final']].groupby('ID loja').sum()
print(faturamento)

# Quantidade de produtos vendidos por loja.
quantidade = tabela_vendas[['ID loja', 'Quantidade']].groupby('ID loja').sum()
print(quantidade)

# Ticket médio por produto em cada loja.
ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame(name='Ticket Médio')
print(ticket_medio)

# Enviar um e-mail com o relatório.
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'andersoncubano@hotmail.com'  # Certifique-se de que o e-mail está correto.
mail.Subject = 'Relatório de Vendas por Loja'

# Gerar o corpo do e-mail com os dados formatados.
mail.HTMLBody = '''
<p>Prezados,</p>

<p>Segue o Relatório de Vendas por cada Loja.</p>

<p><b>Faturamento:</b></p>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p><b>Quantidade Vendida:</b></p>
{quantidade.to_html()}

<p><b>Ticket Médio dos Produtos em cada Loja:</b></p>
{ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}

<p>Qualquer dúvida estou à disposição.</p>

<p>Att.,</p>
<p>AndersonDev</p>
'''

mail.Send()
print('Email Enviado')
# Faturamento por loja.
faturamento = tabela_vendas[['ID loja', 'Valor Final']].groupby('ID loja').sum()
print(faturamento)

# Quantidade de produtos vendidos por loja.
Quantidade = tabela_vendas[['ID loja', 'Quantidade']].groupby('ID loja').sum()
print(Quantidade)

# ticket médio por produto em cada loja
ticket_medio = (faturamento['Valor Final']  /  quantidade['Quantidade']).to_frame()
print(ticket_medio)

# enviar um email com o relatório

outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'andersoncubano@hptmail.com'
mail.Subject = 'Relatório de Vendas por Loja'
mail.HTMLBody = '''


<p>Prezados,</p>

<p>Segue o Relatório de Vendas por cada Loja.</p>

<p>Faturamento:</p>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>Quantidade Vendida:</p>
{quantidade.to_html()}

<p>Ticket Médio dos Produtos em cada Loja:</p>
{ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}

<p>Qualquer dúvida estou à disposição.</p>

<p>Att.,</p>
<p>AndersonDev</p>
'''

mail.Send()

print('Email Enviado')






