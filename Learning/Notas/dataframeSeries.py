# https://medium.com/data-hackers/as-principais-fun%C3%A7%C3%B5es-do-pandas-que-todos-deveriam-saber-60a1899324e6

# Uma tabela de dados no pandas, é conhecida por DataFrame. Apesar do nome ser diferente, “DataFrame”, a estrutura é aquela padrão que conhecemos,
# como a de uma matriz, ou se preferir, como a de alguns SGBD’s (Sistema de Gerenciamento de Banco de Dados) ou planilhas Excel. Um DataFrame é 
# composto por linhas e colunas, e podermos inserir dados de tipos diferente.

import pandas as pd

cientis = { 
           "Nome": ["Marmota", "Sergio"] , 
           "raça": ["Maltes" , "Humano"]   
           }
dataframe = pd.DataFrame(cientis)

print(dataframe)
print("****************")

# Séries: é como uma coluna em uma tabela. Matriz unidimensional que contém dados de qualquer tipo

test = [1,2,3]
test2 = pd.Series(test)

print(test2)
print("****************")


# Crie um DataFrame a partir de duas séries:

dicionario = {
    "Calorias": [420, 380, 350] ,
    "Duração": [50, 45 , 20]
}

dataframe2 = pd.DataFrame(dicionario)


# loc - atributo para retornar uma ou mais linhas especificadas

print(dataframe2.loc[0])
print("****************")


# Adicione uma lista de nomes para dar um nome a cada linha:

data = {
    "Calories": [420 , 500 , 300] ,   # Corresponde a coluna de Calories
    "Duration": [15 , 20 , 10]        # Corresponde a coluna de Duration
}

df = pd.DataFrame( data , index= ['Dia 1', 'Dia 2' , 'Dia 3'])
print(df)
print("****************")

# Localizar índices nomeados

print(df.loc["Dia 2"])
print("****************")


# Carregar arquivos em um DataFrame


csv = pd.read_csv ()







