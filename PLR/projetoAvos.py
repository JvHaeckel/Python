# Script que calcula os avos para o ano de 2025 considerando que o que conta mesmo é a data de afastamento ex: a pessoa X
# seu último dia ativo foi 10/02 e o afastamento foi dia 16 assim não conta pois deve ser maior que 15 ou <=16

import pandas as pd

# Aqui pegou o caminho do arquivo
arquivo = (r"C:\Users\joaorocha\Desktop\Py\PLR\Projeto Avos - Completo.xlsm")

# Faz a leitura do excel usando o Pandas e já escolhe a planilha que será feita as contas no caso foi a Base
table = pd.read_excel(arquivo, sheet_name="Base" )


# Ao imprimir percebi que ele não lê as datas assim teremos que converter
# print(table)

# ***** Convertendo datas *****
table["Afastamento"] = pd.to_datetime(table["Afastamento"], errors= 'coerce')
table["Ultimo dia Ativo"] = pd.to_datetime(table["Ultimo dia Ativo"], errors= 'coerce')
table["Retor."] = pd.to_datetime(table["Retor."], errors='coerce')

# Filtra o DataFrame table para manter apenas as linhas em que a Data Desligamento seja do ano de 2025.
table_2025 = table[table["Afastamento"].dt.year == 2025].copy()
#  Você deve forçar a criação de uma cópia ao criar table_2025 elimina o warning e te garante que está trabalhando em uma cópia segura.

# Verificando se o filtro funcionou
# print(table_2025[["Afastamento","Ultimo dia Ativo"]]) OK

# Criando nova coluna chamada Avos e atribuindo a ela um valor zerado inicial
table_2025 ["Avos"] = 0 

# Laço simples para calcular os avos
for i, row in table_2025.iterrows():
    data_final = row["Afastamento"]

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
                if dias_trabalhados.days + 1 >= 16:
                    meses += 1

        table_2025.at[i, "Avos 2025"] = meses

# Calcular o número de dias Afastados
if table_2025["Retor."] == 0:
    table_2025["Dias"] = table_2025["Afastamento"] - table_2025["Ultimo dia Ativo"]
else:
    table_2025["Dias"] = table_2025["Retor."] - table_2025["Ultimo dia Ativo"]

# Exibe os resultados
print(table_2025[["Nome","Afastamento", "Ultimo dia Ativo", "Avos 2025", "Dias"]])

# Salva o resultado em uma nova planilha Excel
saida = r"C:\Users\joaorocha\Desktop\Py\PLR\Resultado_Avos_2025.xlsx"
table_2025.to_excel(saida, index=False)

print(f"Arquivo exportado com sucesso para: {saida}")