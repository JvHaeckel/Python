

import pandas as pd

dados = {
    "Nome": ["Ana", "Bruno", "Carlos"],
    "Idade": [25, 30, 28]
}

df = pd.DataFrame(dados)
print(df)


# Exemplo 2: Acessar um valor com .loc
print (" Exemplo 2: Acessar um valor com .loc ")
print(df.loc[1, "Nome"])



print (" Exemplo 3: Modificar um valor com .loc ")
modificar = df.loc[1, "Idade"] = 31
print(df)

# Exemplo 4: Modificar várias colunas ao mesmo tempo
print (" Exemplo 4: Modificar várias colunas ao mesmo tempo ")
modificar_varias = df.loc[2, ["Nome", "Idade"]] = ["Carlos Eduardo", 29]
print(df)

# Exemplo 5: Filtrar linhas com .loc
print ("  Exemplo 5: Filtrar linhas com .loc ")

filtrar = df.loc[df['Idade'] > 25]
print(df)


