import pandas as pd
import win32com.client as win32

tabelaVendas = pd.read_excel('Vendas.xlsx')

faturamento = tabelaVendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(faturamento)

quantidade = tabelaVendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(quantidade)

ticketMedio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()
ticketMedio = ticketMedio.rename(columns={0: 'Ticket Médio'})
print(ticketMedio)

outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'email@email.com'
mail.Subject = 'Relatório'
mail.HTMLBody = f'''
<p> Segue o relatório: </p>

<p> Faturamento:</p>
{faturamento.to_html(formatters={"Valor Final": 'R${:,.2f}'.format})}

<p> Quantidade:</p>
{quantidade.to_html()}

<p> Ticket Médio:</p>
{ticketMedio.to_html(formatters={"Ticket Médio": 'R${:,.2f}'.format})}

'''

mail.Send()

print('E-mail Enviado')
