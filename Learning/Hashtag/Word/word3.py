# Integração do Python com Word - Como Criar Contratos Automaticamente
#https://www.youtube.com/watch?v=N01MPYL3UVY
# A partir dos 23:54 vamos usar o pandas

from docx import Document
import pandas as pd

# Lê o arquivo em excel e atribui a variável tabela
tabela = pd.read_excel(r"C:\Users\joaorocha\Desktop\Py\Learning\Hashtag\Word\Informações.xlsx") 

# Percorre as linhas do arquivo Excel
for linha in tabela.index:

     # Carrega o modelo de contrato (arquivo .docx)
    documento = Document(r"C:\Users\joaorocha\Desktop\Py\Learning\Hashtag\Word\Contrato.docx")

    name = input ("Digite o nome para o arquivo: ")

    nome = tabela.loc[linha, "Nome"]
    item1 = tabela.loc[linha, "Item1"]
    item2 = tabela.loc[linha, "Item2"]
    item3 = tabela.loc[linha, "Item3"]
    dia = input("Digite o dia:  ")
    mes = input("Digite o mês: ")
    ano = 2025
    
     # Dicionário
    referencias = {
        "XXXX" : nome,
        "YYYY" : item1,
        "ZZZZ" : item2,
        "WWWW" : item3,
        "DD"   : dia,
        "MM"   : mes,
        "AAAA" : str(ano),
    }

    # Explica em 13:30
    # Percorre todos os parágrafos do documento Word
    for paragrafo in documento.paragraphs:
        
        # Dentro de cada parágrafo, pode haver vários pedaços com estilos diferentes (negrito, itálico, etc.)
        for run in paragrafo.runs:
            
            # Para cada código (ex: "XXXX") e seu valor (ex: "João"), O método .items() retorna pares (chave, valor)
            for codigo, valor in referencias.items():
                
                # Se o código estiver nesse pedaço de texto, ele será substituído pelo valor
                if codigo in run.text:
                    run.text = run.text.replace(codigo, valor)

    documento.save('Contrato de ' + name + " .docx")








