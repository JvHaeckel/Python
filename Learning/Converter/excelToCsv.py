# Como converter uma planilha do Excel para CSV no Python
# https://www.youtube.com/watch?v=PJNba-tHeZY

import pandas as pd

# Pegamos o caminho do arquivo em excel: 

excel_file = (r"C:\Users\joaorocha\Desktop\Py\Learning\Converter\Informações.xlsx")

# Inserir os dados da planilha nessa variável
Data = pd.read_excel(excel_file)

# csv_saida = 




