# Script que calcula os avos para o ano de 2025 considerando que o que conta mesmo é a data de afastamento ex: a pessoa X
# seu último dia ativo foi 10/02 e o afastamento foi dia 16 assim não conta pois deve ser maior que 15 ou <=16. Lembrando
# que o filtro apenas considerou 

import pandas as pd
from pandas.tseries.offsets import MonthEnd

# Caminho do arquivo Excel pegando as duas planilhas
arquivo = r"C:\Users\joaorocha\Desktop\Py\PLR\Projeto Avos - Completo.xlsm"

# Para ler a planilha "Base" vamos fazer a seleção aqui: 
table = pd.read_excel(arquivo, sheet_name="Base")

# Converte colunas de data para o Python poder interpretar
table["Afastamento"] = pd.to_datetime(table["Afastamento"], errors='coerce')
table["Ultimo dia Ativo"] = pd.to_datetime(table["Ultimo dia Ativo"], errors='coerce')
table["Retor."] = pd.to_datetime(table["Retor."], errors='coerce')

# Filtra registros com datas em 2025
table_2025 = table[
    (table["Afastamento"].dt.year == 2025) |
    (table["Retor."].dt.year == 2025) |
    (table["Ultimo dia Ativo"].dt.year == 2025)
].copy()

# Inicializa colunas de Avos
table_2025["Avos Parte 1"] = 0
table_2025["Avos Parte 2"] = 0

# Atribui a variável hoje a data do sistema
hoje = pd.Timestamp.today()


# === Função para contar avos válidos no ano de 2025 ===

# Função auxiliar: calcula número de avos num intervalo (mín. 16 dias por mês)
def contar_avos(inicio, fim):
    if pd.isna(inicio) or pd.isna(fim) or fim < pd.Timestamp("2025-01-01"):
        return 0

    meses = 0
    for mes in range(1, 13):
        inicio_mes = pd.Timestamp(f"2025-{mes:02d}-01")
        fim_mes = inicio_mes + MonthEnd(0)
        if fim < inicio_mes or inicio > fim_mes:
            continue
        real_inicio = max(inicio_mes, inicio)
        real_fim = min(fim_mes, fim)
        dias = (real_fim - real_inicio).days + 1
        if dias >= 16:
            meses += 1
    return meses

# Loop linha a linha
for i, row in table_2025.iterrows():
    # Parte 1 – Até o último dia ativo
    ultimo_ativo = row["Ultimo dia Ativo"]
    if pd.notna(ultimo_ativo) and ultimo_ativo >= pd.Timestamp("2025-01-01"):
        avos1 = contar_avos(pd.Timestamp("2025-01-01"), ultimo_ativo)
        table_2025.at[i, "Avos Parte 1"] = avos1

    # Parte 2 – Após retorno (ou 0 se não retornou)
    retorno = row["Retor."]
    afastamento = row["Afastamento"]

    if pd.notna(retorno):
        avos2 = contar_avos(retorno, hoje)
    else:
        # Não voltou ainda → NÃO conta avos após afastamento
        avos2 = 0

    table_2025.at[i, "Avos Parte 2"] = avos2

# Soma final
table_2025["Avos 2025"] = table_2025["Avos Parte 1"] + table_2025["Avos Parte 2"]


# ***************************************************************************************************************************************************

# Cálculo dos dias afastados
dias_afastados = []

for i, row in table_2025.iterrows():
    retorno = row["Retor."]
    ultimo_ativo = row["Ultimo dia Ativo"]

    if pd.isna(ultimo_ativo):
        dias = None
    elif pd.isna(retorno):
        dias = (hoje - ultimo_ativo).days
    else:
        dias = (retorno - ultimo_ativo).days

    dias_afastados.append(dias)

table_2025["Dias"] = dias_afastados

# Seleção das colunas para exportação
colunas = [
    "Chapa", "Divisões", "Nome", "Função", "Admis.",
    "Ultimo dia Ativo", "Afastamento", "Cid", "Retor.",
    "Dias", "Motivo", "Avos Parte 1", "Avos Parte 2", "Avos 2025"
]

resultado = table_2025[colunas]

# Exporta para Excel
saida = r"C:\Users\joaorocha\Desktop\Py\PLR\Resultado_Avos_2025.xlsx"
resultado.to_excel(saida, index=False)


print(resultado)
print(f"Arquivo exportado com sucesso para: {saida}")
