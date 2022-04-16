# pip install pywin32
import win32com.client
import pandas as pd
import numpy as np

df = pd.read_excel('excel.xlsx', 'Planilha1').values
dest = pd.read_excel('excel.xlsx', 'Planilha2').values

for i in range(len(dest)):
    linha = ''
    for j in range(len(df)):
        if df[j][0] == dest[i][0]:
            formatando = str(df[j]).replace('[',"")
            formatando = formatando.replace(']',"")
            formatando = formatando.replace("'", "\t")

            linha = (linha + formatando + '\n')

    o = win32com.client.Dispatch("Outlook.Application")

    Msg = o.CreateItem(0)
    Msg.To = dest[i][1]

    # Msg.CC = "more email addresses here"

    Msg.Subject = f"Equipe da {dest[i][0]}"

    Msg.Body = "Sr(a) Gestor, bom dia.\n\n" \
               "Segue as informações\n\n" \
               f"{linha} \n\n" \
               f"Atenciosamente\n" \
               f"Fulano de Tal"

    # # Anexos
    # attachment1 = "Caminho do anexo no. 1"
    # attachment2 = "Caminho do anexo no. 2"
    # Msg.Attachments.Add(attachment1)
    # Msg.Attachments.Add(attachment2)

    Msg.Send()
