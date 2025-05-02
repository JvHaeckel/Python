# Integração entre Python e Excel usando Pandas e o Openpyxl
# 
# https://www.youtube.com/watch?v=IT7zPluDADk&t=502s

import pandas as pd
import os  # tive que importar para poder salvar o arquivo na mesma pasta


# Coloquei uma variável para poder usar com mais facilidade
# Perceba que copiei no VS code o Caminho relativo, assim fica mais curto
caminho = "Learning\Hashtag\Produtos.xlsx"

table = pd.read_excel(caminho)

# Sempre teste o que está trazendo: print(table) 

# table.loc [linha, coluna]
# Ele queria mudar nas linhas aonde tem serviço o multiplicador de imposto para 1,5
table.loc[table["Tipo"] == "Serviço", "Multiplicador Imposto"] = 1.5

table["Preço Base Reais"] = table["Multiplicador Imposto"] * table["Preço Base Original"]

saida = "Learning\Hashtag\Produtos.xlsx"

# Para salvar na mesma pasta coloquei essas duas linhas a mais abaixo e ainda tive que importar
# o OS acima
script_dir = os.path.dirname(os.path.abspath(__file__))
saida = os.path.join(script_dir, "Pandas_Produtos.xlsx")
table.to_excel(saida, index=False, engine="openpyxl")








