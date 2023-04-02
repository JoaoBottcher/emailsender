import pandas as pd #as pd (pd sera o apelido de pandas para reduzir escrita)
import win32com.client as win32
#importar a base de dados
tabela_vendas = pd.read_excel('Vendas.xlsx') #armazena a base de dados dentro da tabela_vendas

#visualizar a base de dados
pd.set_option('display.max_columns', None) #mostra o maximo de colunas da tabela, sem ocultar

print(tabela_vendas)

#faturamento por loja (agrupando todas as lojas e somar a coluna do valor final)
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(faturamento)

#quantidade de produtos vendidos por loja
quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(quantidade)


print('-' * 50)

#ticket medio por produto de cada loja
ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()#to_frame transforma em tabela
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'})
print (ticket_medio)

#enviar um email com o relatorio
outlook = win32.Dispatch('Outlook.Application')
mail = outlook.CreateItem(0)
mail.To = 'johnbottcher22@outlook.com'
mail.Subject = 'Relatório de Vendas por Loja'
mail.HTMLBody = f'''
<p>Prezados,</p>

</p>Segue o Relatório de Vendas por cada Loja.</p>

<p>Faturamento:</p>
{faturamento.to_html()}

<p>Quantidade Vendida:</p>
{quantidade.to_html()}

<p>Ticket Médio dos Produtos em cada Loja:</p>
{ticket_medio.to_html()}

<p>Qualquer dúvida estou à disposição</p>

<p>Atenciosamente,</p>
'''
mail.Send()

print('Email enviado!')