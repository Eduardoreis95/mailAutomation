import pandas as pd
import win32com.client as win32

# Importação da base de dados

tabela_vendas = pd.read_excel("Vendas.xlsx")

# Visualisação da base de dados

pd.set_option("display.max_columns", None)
print(tabela_vendas)

print("-" * 50)

# Faturamento por loja

faturamento = tabela_vendas[["ID Loja", "Valor Final"]].groupby("ID Loja").sum()
print(faturamento)

print("-" * 50)

# Quantidade de produtos vendidos por loja

quantidade = tabela_vendas[["ID Loja", "Quantidade"]].groupby("ID Loja").sum()
print((quantidade))

print("-" * 50)

# Ticket médio por produto em cada loja

ticket_medio = (faturamento["Valor Final"] / quantidade["Quantidade"]).to_frame()
ticket_medio = ticket_medio.rename(columns={0: "Ticket Médio"})
print(ticket_medio)

# Envio de e-mail com o relatório

outlook = win32.Dispatch("outlook.application")
mail = outlook.CreateItem(0)
mail.To = "Email"
mail.Subject = "Relatório de Vendas por Loja"
mail.HTMLBody = f"""

<p>Prezados,</p>

<p>Segue relatório de vendas por cada loja, sendo:</p>

<p>Faturamento:</p>
{faturamento.to_html(formatters={"Valor Final": "R${:,.2f}".format})}

<p>Quantidade Vendida:</p>
{quantidade.to_html()}

<p>Ticket Médio dos Produtos em cada Loja:</p>
{ticket_medio.to_html(formatters={"Ticket Médio": "R${:,.2f}".format})}

<p>Qualquer dúvida, estamos à disposição.</p>

<p>Cordialmente,</p>
<p>Eduardo Reis.</p>

"""

mail.Send()

print("Email enviado.")