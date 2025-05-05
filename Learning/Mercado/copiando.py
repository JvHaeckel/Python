# Calcular os avos para o ano de 2025 para a tabela base da planilha: PLR



import pandas as pd
from openpyxl import load_workbook

def automate_excel():
    # Caminho do Excel
    File = r'C:\Users\joaorocha\Desktop\Projeto Avos\Projeto Avos - Completo.xlsm'

    try: 
        # Nome da aba que você quer ler
        sheet_name = "Base"
        
        # Lê o arquivo Excel
        avos = pd.read_excel(File, sheet_name=sheet_name)

        print("Arquivo lido com sucesso")
        print(avos[['Chapa', 'Função']].head())

        # Converte as datas
        avos["Admissão"] = pd.to_datetime(avos["Admissão"], errors="coerce")
        avos["Afastamento"] = pd.to_datetime(avos["Afastamento"], errors="coerce")

        # Ano de referência
        ano_referencia = 2024
        inicio_ano = pd.Timestamp(f"{ano_referencia}-01-01")
        fim_ano = pd.Timestamp(f"{ano_referencia}-12-31")

        # Função para calcular avôs
        def calcular_avos(row):
            adm = max(row["Admissão"], inicio_ano) if not pd.isnull(row["Admissão"]) else inicio_ano
            afa = min(row["Afastamento"], fim_ano) if not pd.isnull(row["Afastamento"]) else fim_ano

            if afa < inicio_ano or adm > fim_ano:
                return 0

            meses = 0
            atual = adm

            while atual <= afa:
                ultimo_dia_mes = atual + pd.offsets.MonthEnd(0)
                dias_trabalhados = (min(afa, ultimo_dia_mes) - atual).days + 1

                if dias_trabalhados >= 15:
                    meses += 1

                atual = ultimo_dia_mes + pd.Timedelta(days=1)

            return meses / 12

        # Aplica o cálculo
        avos["Avos"] = avos.apply(calcular_avos, axis=1)

        print("\nResultado dos avôs:")
        print(avos[["Chapa", "Admissão", "Afastamento", "Avos"]].head())

        # Aqui você pode continuar com sua pivot_table se quiser
        # ...

    except Exception as e:
        print(f"Erro ao carregar arquivo: {e}")
