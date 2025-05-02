# Importações

import pandas as pd
from docx import Document
import tkinter as tk

caminho = (r"C:\Users\joaorocha\Desktop\Py\Learning\Converter\Informações.xlsx")

# Ler e carrega os dados do excel
tabela = pd.read_excel(caminho)

# Criando um novo documento
documento = Document()

titulo = input("Qual título do documento? ")
documento.add_heading(titulo, level=1)

# Itera pelas linhas da planilha
for index, linha in tabela.iterrows():  # Esse comando itera (passa por) cada linha da tabela do Excel.
    texto = ""                          # 
    for coluna in tabela.columns:
        texto += f"{coluna}: {linha[coluna]}  "
    documento.add_paragraph(texto.strip(" | "))

# Salvar o documento Word com o nome escolhido pela pessoa 
name = input("Qual o nome do arquivo? ")
documento.save(fr"C:\Users\joaorocha\Desktop\Py\Learning\Converter\{name}.docx")


# Criando janela
janela = tk.Tk()
janela.title("Conversor Excel → Word")


