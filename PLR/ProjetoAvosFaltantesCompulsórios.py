import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import os  # Faltava esse import

# Explicação do que o código deve fazer:
# A funcionária ELIZETE SILVA ROCHA AMORIM faltou em duas datas diferentes (exemplo na linha 1000).
# O programa precisa percorrer a coluna 'Chapa' (matrícula) e verificar se o mesmo número aparece mais de uma vez.
# Se não aparecer, ignora. Se aparecer mais de uma vez, deve:
# 1. Exibir essas linhas duplicadas numa nova planilha;
# 2. Depois, calcular os avos para cada data e somar os resultados.

def processamento():
    # filedialog.askopenfilename abre uma janela para o usuário escolher um arquivo Excel
    caminho_arquivo = filedialog.askopenfilename(
        title="Selecione arquivo em Excel apenas",
        filetypes=[("Arquivos Excel", "*.xlsx *.xlsm *.xls")]
    )

    # Se o usuário cancelar e não escolher nenhum arquivo
    if caminho_arquivo == "":
        return

    # Lê o arquivo Excel
    table = pd.read_excel(caminho_arquivo)

    # Garante que a coluna 'Chapa' é tratada como número inteiro
    table["Chapa"] = pd.to_numeric(table["Chapa"], errors="coerce").astype('Int64')

    # Filtra as chapas que aparecem mais de uma vez
    duplicatas = table[table.duplicated(subset='Chapa', keep=False)]

    if duplicatas.empty:
        messagebox.showinfo("Sem duplicatas", "Nenhuma chapa repetida foi encontrada.")
    else:
        # Exibir duplicatas em nova planilha
        novo_arquivo = os.path.splitext(caminho_arquivo)[0] + "_chapas_repetidas.xlsx"
        duplicatas.to_excel(novo_arquivo, index=False)
        messagebox.showinfo("Duplicatas encontradas", f"Chapas repetidas salvas em:\n{novo_arquivo}")

################ Interface gráfica ################

janela = tk.Tk()
janela.geometry("400x220")  # Corrigido: sem espaço no meio
janela.title("Calcular os faltantes compulsórios")

botao = tk.Button(janela, text="Procurar", command=processamento)
botao.pack(pady=10)

janela.mainloop()