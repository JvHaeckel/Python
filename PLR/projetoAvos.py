import pandas as pd

# Aqui pegou o caminho do arquivo
arquivo = (r"C:\Users\joaorocha\Desktop\Py\PLR\Projeto Avos - Completo.xlsm")

# Faz a leitura do excel usando o Pandas
table = pd.read_excel(arquivo, sheet_name="Base" )

# Ao imprimir percebi que ele não lê as datas assim teremos que converter
# print(table)

# ***** Convertendo datas *****




