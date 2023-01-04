import win32com.client as win32


#Integração do python com o outlook
outlook  = win32.Dispatch('outlook.application')

#Criar e-mail
email = outlook.CreateItem(0)

faturamento = 1500
qtde_produtos = 10
ticket_medio = faturamento / qtde_produtos

email.To = "wallasscurty@gmail.com; wallas123@gmail.com" #Destino do E-mail
email.Subject = "E-mail autimático do Python" #Assunto 
email.HTMLBody = f""" 
<h3>Boa tarde Wallas !!</h3>

<p>O faturamento da loja foi de R$ {faturamento}</p>
<p>Vendemos {qtde_produtos} produtos</p>
<p>o ticket Médio foi de R${ticket_medio}</p>
"""

#Anexar aquivos
anexo = "C:\Users\walla\Documents\Ciencia de Dados\Jupyter"
email.Attachments.Add(anexo)


email.Send()
print("Email foi enviado com sucesso")