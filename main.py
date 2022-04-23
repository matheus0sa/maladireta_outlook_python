import win32com.client
import pandas as pd
import numpy as np

# Na variável df vamos guardar a tabela principal com os funcionários
# Na variável dest vamos guardar a tabela dos destinatários
df = pd.read_excel('excel.xlsx', 'Planilha1')
dest = pd.read_excel('excel.xlsx', 'Planilha2').values

# Vamos fazer um laço para percorrer cada linha da tabela de destinatários
for i in range(len(dest)):
    # podemos acessar os valores de dest passando [linha][coluna]
    tabela = df.query(f"Lotação == '{dest[i][0]}'")
    tabela = tabela.to_html()

    # Conecção com o Outlook
    o = win32com.client.Dispatch("Outlook.Application")
    Msg = o.CreateItem(0)

    # Destinatário
    Msg.To = dest[i][1]

    # Msg.CC = "Caso queira colocar copia para algum endereço"

    # Assunto
    Msg.Subject = f"Equipe da {dest[i][0]}"

    # Corpo do email, deve ser escrito em html
    Msg.HTMLBody = "Sr(a) Gestor, bom dia.<br><br>" \
               "Segue as informações<br><br>" \
               f"{tabela} <br><br>" \
               f"Atenciosamente<br><br>" \
               f"Fulano de Tal"

    # # Anexos
    # attachment1 = "Caminho do anexo no. 1"
    # attachment2 = "Caminho do anexo no. 2"
    # Msg.Attachments.Add(attachment1)
    # Msg.Attachments.Add(attachment2)

    # Enviar
    Msg.Send()
