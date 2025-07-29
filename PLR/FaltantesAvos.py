# Código para aprender a somar as linhas dos faltantes

import pandas as pd
from datetime import datetime

# Função para contar avôs válidos
def calcular_avos_2025(row):
    ano_base = 2025
    avos = 0

    # Se tiver data de retorno válida e estiver no ano de 2025
    if pd.notna(row['Retorno']):
        data_inicio = max(pd.to_datetime(row['Retorno']), pd.Timestamp(f'{ano_base}-01-01'))
    elif pd.notna(row['Admissao']):
        data_inicio = max(pd.to_datetime(row['Admissao']), pd.Timestamp(f'{ano_base}-01-01'))
    else:
        return 0

    # Se tiver desligamento, usa o "Último dia ativo", senão vai até 31/12/2025
    if pd.notna(row['Ultimo_dia_Ativo']):
        data_fim = min(pd.to_datetime(row['Ultimo_dia_Ativo']), pd.Timestamp(f'{ano_base}-12-31'))
    else:
        data_fim = pd.Timestamp(f'{ano_base}-12-31')

    if data_fim < data_inicio:
        return 0

    for mes in range(1, 13):
        inicio_mes = pd.Timestamp(f'{ano_base}-{mes:02d}-01')
        fim_mes = pd.Timestamp(f'{ano_base}-{mes:02d}-28') + pd.offsets.MonthEnd(0)

        dias_trabalhados = (min(data_fim, fim_mes) - max(data_inicio, inicio_mes)).days + 1
        if dias_trabalhados >= 15:
            avos += 1

    return avos

# Exemplo: carregando os dados da sua planilha
df = pd.read_excel('avos_2025.xlsx')  # ou o caminho do seu arquivo
df.rename(columns={
    'Admis.': 'Admissao',
    'Afastamento': 'Afastamento',
    'Ultimo dia Ativo': 'Ultimo_dia_Ativo',
    'Retor.': 'Retorno'
}, inplace=True)

# Aplica a função para cada linha
df['Avos 2025.1 calculado'] = df.apply(calcular_avos_2025, axis=1)

# Se quiser somar os avôs por funcionário (caso tenha várias linhas por pessoa):
df_total = df.groupby(['Chapa', 'Nome'], as_index=False)['Avos 2025.1 calculado'].sum()

# Exibe resultado
print(df_total)
