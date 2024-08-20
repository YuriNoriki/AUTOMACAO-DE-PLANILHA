import pandas as pd
import os
from datetime import datetime
import win32com.client as win32

#criando variavel para o diretorio das bases 
caminho = "bases"
# Lista todos os arquivos e diretórios no caminho especificado
arquivos = os.listdir(caminho)
print(arquivos)

#criamos uma tabela vazia 
tabela_consolidada = pd.DataFrame()

# Fazemos o FOR para percorrer o caminho das bases
for nome_arquivo in arquivos:
    
    # Criamos a tabela e lemos os arquivos.csv(bases) e depois juntamos todos eles 
    tabela_vendas = pd.read_csv(os.path.join(caminho,nome_arquivo))

    # No Excel a data é marcada por dias 1,2,3,4,5.....
    # Converter "Data de Vendas" de número de dias para formato datetime
    tabela_vendas["Data de Venda"] = pd.to_datetime("01/01/1900") + pd.to_timedelta(tabela_vendas["Data de Venda"],unit="d")

    # Juntou a tabela_vendas com a tabela_consolidada que estava fazia 
    tabela_consolidada = pd.concat([tabela_consolidada, tabela_vendas])

tabela_consolidada.to_excel("Vendas.xlsx", index=False)

# Transformamos a DATA que era em dias para data 10/10/2024
tabela_consolidada = tabela_consolidada.sort_values(by="Data de Venda")
# Ordena as linhas  
tabela_consolidada = tabela_consolidada.reset_index(drop=True)

# ENVIO DE EMAIL//COPIA E COLA

outlook = win32.Dispatch('outlook.application')
email = outlook.CreateItem(0)
email.To = "yushinzato36@gmail.com"
data_hoje = datetime.today().strftime("%d/%m/%Y")
email.Subject = f"Relatorio de vendas {data_hoje}"
email.Body = """ Segue em anexo as Bases para sua analise
                    qualquer coisa só chmar"""
cam = os.getcwd()
anexo = os.path.join(cam,"Vendas.xlsx")
email.Attachments.Add(anexo)
email.Send()
print('enviado')



