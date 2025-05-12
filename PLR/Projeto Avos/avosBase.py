
import pandas as pd
from openpyxl import load_workbook

def automate_excel():
    
    #Caminho do excel
    File = r'C:\Users\joaorocha\Desktop\Projeto Avos\Projeto Avos - Completo.xlsm'

# Ler o arquivo Excel
    try: 
     # Nome da aba que você quer ler
      sheet_name = "Base"
      
      avos = pd.read_excel(File, sheet_name=sheet_name)

      print("Arquivo lido com sucesso")
      print(avos[['Chapa', 'Função']])
    
    # Converte as datas
    avos["Admissão"] = pd.to_datetime(avos["Admissão"], errors='coerce')
    avos["Afastamento"] = pd.to_datetime(avos["Admissão"], errors = 'coerce')
    
    # Ano de referência
    ano_referencia = 2025
    inicio_ano = pd.Timestamp(f"{ano_referencia}-01-01")
    fim_ano = pd.Timestamp(f"{ano_referencia}-12-31")
    
    # Função para calcular avôs
   # Função para calcular as datas ajustadas de admissão e afastamento para o cálculo de avôs

def calcular_avos(row):
    if pd.notnull(row["Admissão"]):
        adm = max(row["Admissão"], inicio_ano)
    else:
        adm = inicio_ano

    if pd.notnull(row["Afastamento"]):
        afa = min(row["Afastamento"], fim_ano)
    else:
        afa = fim_ano
    
    
    
    
    
    
    except Exception as e:
       print (f"Erro ao carregar arquivo {e}")
    return





# Criar a Tabela Dinâmica (Pivot Table)



  # Chamar a função para rodar o processo
automate_excel()
    
    
    
    
    
    
    