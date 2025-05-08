# https://medium.com/data-hackers/as-principais-fun%C3%A7%C3%B5es-do-pandas-que-todos-deveriam-saber-60a1899324e6

# Uma tabela de dados no pandas, é conhecida por DataFrame. Apesar do nome ser diferente, “DataFrame”, a estrutura é aquela padrão que conhecemos,
# como a de uma matriz, ou se preferir, como a de alguns SGBD’s (Sistema de Gerenciamento de Banco de Dados) ou planilhas Excel. Um DataFrame é 
# composto por linhas e colunas, e podermos inserir dados de tipos diferente.

import pandas as pd

cientis = { "Nome": ["Marmota", "Sergio"] , "raça": ["Maltes" , "Humano"]   }


dataframe = pd.DataFrame(cientis)

print(dataframe)



test = [1,2,3]

test2 = pd.Series(test)
print(test2)
