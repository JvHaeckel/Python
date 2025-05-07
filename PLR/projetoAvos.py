# Script que calcula os avos para o ano de 2025 considerando que o que conta mesmo é a data de afastamento ex: a pessoa X
# seu último dia ativo foi 10/02 e o afastamento foi dia 16 assim não conta pois deve ser maior que 15 ou <=16. Lembrando
# que o filtro apenas considerou 

import pandas as pd
from datetime import datetime

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

# Filtra o DataFrame table para manter apenas as linhas em que a Data Afastamento seja do ano de 2025.

table_2025 = table[(table["Afastamento"].dt.year == 2025) | 
                   (table["Retor."].dt.year == 2025)].copy()

#  Você deve forçar a criação de uma cópia ao criar table_2025 elimina o warning e te garante que está trabalhando em uma cópia segura.

# Verificando se o filtro funcionou
# print(table_2025[["Afastamento","Ultimo dia Ativo"]]) OK  

# Criando nova coluna chamada Avos e atribuindo a ela um valor zerado inicial
table_2025 ["Avos"] = 0 

# ***********************************************************************************

# Laço simples para calcular os avos

# Data atual
hoje = pd.Timestamp.today()

for i, row in table_2025.iterrows():
    ultimo_ativo = row["Ultimo dia Ativo"]
    afastamento = row["Afastamento"]
    retorno = row["Retor."]

    # Se não há último dia ativo, não dá pra calcular
    if pd.isna(ultimo_ativo):
        continue

    # Definimos o início do período de cálculo: o último dia ativo
    inicio_periodo = pd.Timestamp("2025-01-01")

    # Se retorno existe, usamos ele como reentrada e seguimos até hoje (ou até 31/12/2025)
    if pd.notna(retorno):
        retorno_efetivo = retorno
    else:
        # Se ainda não voltou, ficou afastado até hoje
        retorno_efetivo = hoje

    # Garante que o retorno não passe de 31/12/2025
    if retorno_efetivo > pd.Timestamp("2025-12-31"):
        retorno_efetivo = pd.Timestamp("2025-12-31")

    meses = 0

    for mes in range(1, 13):
        inicio_mes = pd.Timestamp(f"2025-{mes:02d}-01")
        fim_mes = inicio_mes + pd.offsets.MonthEnd(0)

        # Se o mês começa depois do retorno, consideramos como ativo
        if fim_mes < retorno_efetivo:
            continue  # Estava afastado esse mês inteiro

        # Considera os meses entre retorno e hoje
        if inicio_mes > retorno_efetivo:
            inicio_trabalho = inicio_mes
        else:
            inicio_trabalho = retorno_efetivo

        fim_trabalho = min(fim_mes, hoje, pd.Timestamp("2025-12-31"))

        if fim_trabalho < inicio_trabalho:
            continue

        dias_trabalhados = (fim_trabalho - inicio_trabalho).days + 1

        if dias_trabalhados >= 16:
            meses += 1

    table_2025.at[i, "Avos 2025"] = meses


# ***********************************************************************************


# ***** Calcular o número de dias das pessoas Afastadas até a data de hoje *****
dias_afastados = []
hoje = pd.Timestamp.today()

for i, row in table_2025.iterrows():
    retorno = row["Retor."]
    ultimo_ativo = row["Ultimo dia Ativo"]

    if pd.isna(retorno):
        dias = (hoje - ultimo_ativo).days
    else:
        dias = (retorno - ultimo_ativo).days

    dias_afastados.append(dias)

table_2025["Dias"] = dias_afastados

# ***** Exibe os resultados *****

exibir = [
    "Chapa",
    "Divisões",
    "Nome",	
    "Função",	
    "Admis." ,
    "Ultimo dia Ativo",
    "Afastamento",
    "Cid",
    "Retor." ,
    "Dias" 	,
    "Motivo",
    "Avos 2025"	
]
# Para não sobrescrever o table_2025 com uma lista de strings, o que quebra o .to_excel() criamos outra variável chamada resultado
resultado = table_2025[exibir]
print(resultado)

# Salva o resultado em uma nova planilha Excel
saida = r"C:\Users\joaorocha\Desktop\Py\PLR\Resultado_Avos_2025.xlsx"
resultado.to_excel(saida, index=False)

print(f"Arquivo exportado com sucesso para: {saida}")