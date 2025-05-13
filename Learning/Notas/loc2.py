# Exercício proposto pelo chat gpt : 

# 1) Crie o DataFrame
# 2) Aumente o salário do Bruno para 5500 usando .loc.
# 3) Troque o cargo da Carla para "Analista Jr".
# 4) Diminua o salário do Diego em 1000 reais.
# 5) Filtre e mostre os funcionários que ganham mais de 3000 reais.
# 6) Aumente o salário de todos que ganham até 3000 em 10%.

import pandas as pd

# Dicionário criado pelo chat gpt
dados = {
    "Nome": ["Alice", "Bruno", "Carla", "Diego"],
    "Cargo": ["Analista", "Gerente", "Assistente", "Diretor"],
    "Salario": [3000, 5000, 2500, 8000]
}

print("1) Crie o DataFrame")
print("")
df= pd.DataFrame(dados)
print(df)
print("")


print("2) Aumente o salário do Bruno para 5500 usando .loc.")
print("")

aumento = df.loc[1, 'Salario'] = 5500

print(df)
print("")


print("3) Troque o cargo da Carla para Analista Jr")
print("")

trocar_cargo = df.loc[2 , "Cargo"] = " Analista Jr"

print(df)
print("")

print("5) Filtre e mostre os funcionários que ganham mais de 3000 reais.")
print("")


print(df)
print("")






