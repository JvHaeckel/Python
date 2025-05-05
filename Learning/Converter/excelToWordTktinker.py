import pandas as pd
from docx import Document
import tkinter as tk
from tkinter import filedialog

# Função que converte Excel em Word
def converter():
    # Aqui está sendo usada a biblioteca tkinter para abrir uma janela de diálogo para que o usuário escolha um arquivo no computador. 
    # filedialog.askopenfilename():Esse método abre uma janela para o usuário selecionar um arquivo e retorna o caminho completo do 
    # arquivo escolhido. 
    # 
    # filetypes=[("Excel files", "*.xlsx")]:Especifica que a janela de diálogo deve filtrar e mostrar apenas arquivos com extensão 
    # .xlsx, ou seja, apensa arquivos em Excel
    
    
    
    arquivo_excel = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    
    if not arquivo_excel:
        return
    
    # Lê a planilha usando Pandas
    tabela = pd.read_excel(arquivo_excel)

    # Cria o documento Word
    documento = Document()
    
    titulo = campo_titulo.get()
    documento.add_heading(titulo, level=1)

    for _, linha in tabela.iterrows():
        texto = ""
        for coluna in tabela.columns:
            texto += f"{coluna}: {linha[coluna]}  "
        documento.add_paragraph(texto)
        documento.add_paragraph("")  # Linha em branco entre os itens

    # Salvar com nome digitado
    nome_arquivo = campo_nome.get()
    if nome_arquivo.strip() == "":
        nome_arquivo = "documento_word"
        
    documento.save(f"{nome_arquivo}.docx")

# Criando janela
janela = tk.Tk()
janela.title("Conversor Excel→Word Powered By MOBI")

# Campos na janela
tk.Label(janela, text="Digite o Título do documento:").pack(pady=(10,0))
campo_titulo = tk.Entry(janela, width=50)
campo_titulo.pack(pady=(0,10))

tk.Label(janela, text="Digite o nome do arquivo Word:").pack()
campo_nome = tk.Entry(janela, width=50,)
campo_nome.pack()

# Botão para converter
botao = tk.Button(janela, text="Selecionar e Converter", command=converter)
botao.pack(pady=20)

# Inicia a janela
janela.mainloop()
