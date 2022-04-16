import win32com.client
import pandas as pd
import numpy as np

def tabular(lista):
    global tabela
    tabela = tabela + '<tr>'
    for item in lista:
        tabela = tabela + '<td>' + item + '</td>'
    tabela = tabela + '</tr>'
    return tabela

df = pd.read_excel('excel.xlsx', 'Planilha1').values
dest = pd.read_excel('excel.xlsx', 'Planilha2').values

for i in range(len(dest)):
    tabela = '<table border ="1">' \
             '<tr>' \
             '  <th>Lotação</th>' \
             '  <th>Nome</th>' \
             '  <th>Função</th>' \
             '</tr>'
    for j in range(len(df)):
        if df[j][0] == dest[i][0]:
            tabela= tabular(df[j])

    tabela = tabela + '</table>'
    o = win32com.client.Dispatch("Outlook.Application")

    Msg = o.CreateItem(0)
    Msg.To = dest[i][1]

    # Msg.CC = "more email addresses here"

    Msg.Subject = f"Equipe da {dest[i][0]}"

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

    Msg.Send()
