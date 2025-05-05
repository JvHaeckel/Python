import pandas as pd

# Aqui pegou o caminho do arquivo
arquivo = (r"C:\Users\joaorocha\Desktop\Py\PLR\Projeto Avos - Completo.xlsm")

# Faz a leitura do excel usando o Pandas
table = pd.read_excel(arquivo, sheet_name="Base" )

# Ao imprimir percebi que ele não lê as datas assim teremos que converter
# print(table)

# ***** Convertendo datas *****
table["Afastamento"] = pd.to_datetime(table["Afastamento"], errors= 'coerce')
table["Ultimo dia Ativo"] = pd.to_datetime(table["Ultimo dia Ativo"], errors= 'coerce')

# Filtra o DataFrame table para manter apenas as linhas em que a Data Desligamento seja do ano de 2025.
table_2025 = table[table["Afastamento"].dt.year == 2025]

# Verificando se o filtro funcionou
# print(table_2025[["Afastamento","Ultimo dia Ativo"]]) OK


# Criando nova coluna chamada Avos e atribuindo a ela um valor zerado inicial
table_2025 ["Avos"] = 0 

# Laço simples para calcular os avos
for i, row in table_2025.iterrows():
    data_final = row["Ultimo dia Ativo"]

    # Garante que a data final esteja dentro de 2025
    if pd.notna(data_final):
        if data_final > pd.Timestamp("2025-12-31"):
            data_final = pd.Timestamp("2025-12-31")

        # Conta os meses completos até a data final
        meses = 0
        for mes in range(1, 13):  # Janeiro (1) até Dezembro (12)
            inicio_mes = pd.Timestamp(f"2025-{mes:02d}-01")
            fim_mes = pd.Timestamp(f"2025-{mes:02d}-28") + pd.offsets.MonthEnd(0)

            # Verifica se trabalhou pelo menos 15 dias no mês
            if data_final >= inicio_mes:
                dias_trabalhados = min(data_final, fim_mes) - inicio_mes
                if dias_trabalhados.days + 1 >= 15:
                    meses += 1

        table_2025.at[i, "Avos 2025"] = meses

# Exibe os resultados
print(table_2025[["Nome","Afastamento", "Ultimo dia Ativo", "Avos 2025"]])
