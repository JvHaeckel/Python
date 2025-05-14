import pandas as pd
from pandas.tseries.offsets import MonthEnd  # Faltava importar
import tkinter as tk
from tkinter import filedialog, messagebox
import os

# Função para contar avos válidos no ano de 2025
def contar_avos(inicio, fim):
    if pd.isna(inicio) or pd.isna(fim) or fim < pd.Timestamp("2025-01-01"):
        return 0

    meses = 0
    for mes in range(1, 13):
        data_texto = f"2025-{mes:02d}-01"
        inicio_mes = pd.Timestamp(data_texto)
        fim_mes = inicio_mes + MonthEnd(0)

        if fim < inicio_mes or inicio > fim_mes:
            continue

        real_inicio = max(inicio_mes, inicio)
        real_fim = min(fim_mes, fim)

        dias = (real_fim - real_inicio).days + 1
        if dias >= 15:
            meses += 1

    return meses


def processamento():
    caminho_arquivo = filedialog.askopenfilename(
        title="Selecione arquivo em Excel apenas",
        filetypes=[("Arquivos Excel", "*.xlsx *.xlsm *.xls")]
    )

    if caminho_arquivo == "":
        return

    try:
        # Corrigido: faltava dois-pontos
        table = pd.read_excel(caminho_arquivo)
        table.columns = table.columns.str.strip()

        # Conversões e garantias de tipos
        table["Chapa"] = pd.to_numeric(table["Chapa"], errors="coerce").astype('Int64')
        table["Afastamento"] = pd.to_datetime(table["Afastamento"], errors='coerce')
        table["Ultimo dia Ativo"] = pd.to_datetime(table["Ultimo dia Ativo"], errors='coerce')
        table["Retor."] = pd.to_datetime(table["Retor."], errors='coerce')
        table["Admis."] = pd.to_datetime(table["Admis."], errors='coerce')

        # Encontrar chapas duplicadas
        duplicatas = table[table.duplicated(subset='Chapa', keep=False)]

        if duplicatas.empty:
            messagebox.showinfo("Sem duplicatas", "Nenhuma chapa repetida foi encontrada.")
        else:
            novo_arquivo = os.path.splitext(caminho_arquivo)[0] + "_chapas_repetidas.xlsx"
            duplicatas.to_excel(novo_arquivo, index=False)
            messagebox.showinfo("Duplicatas encontradas", f"Chapas repetidas salvas em:\n{novo_arquivo}")

        # Filtro para registros relevantes de 2025
        table_2025 = table[
            (table["Afastamento"].dt.year == 2025) |
            (table["Retor."].dt.year == 2025) |
            (table["Ultimo dia Ativo"].dt.year == 2025) |
            (table["Situação"] == "A")
        ].copy()

        table_2025["Avos Parte 1"] = 0
        table_2025["Avos Parte 2"] = 0
        table_2025["Avos 2025"] = 0

        data_final = pd.Timestamp("2025-12-31")
        data_inicio_ano = pd.Timestamp("2025-01-01")

        for i, row in table_2025.iterrows():
            situacao = row["Situação"]
            retorno = row["Retor."]
            admissao = row["Admis."]
            afastamento = row["Afastamento"]
            ultimo_ativo = row["Ultimo dia Ativo"]

            if situacao == "A":
                if pd.notna(retorno) and pd.notna(ultimo_ativo) and retorno >= data_inicio_ano:
                    parte1 = contar_avos(data_inicio_ano, ultimo_ativo)
                    parte2 = contar_avos(retorno, data_final)
                    table_2025.loc[i, "Avos Parte 1"] = parte1
                    table_2025.loc[i, "Avos Parte 2"] = parte2
                    table_2025.loc[i, "Avos 2025"] = parte1 + parte2
                elif pd.notna(admissao) and admissao >= data_inicio_ano:
                    avos = contar_avos(admissao, data_final)
                    table_2025.loc[i, "Avos Parte 1"] = avos
                    table_2025.loc[i, "Avos Parte 2"] = 0
                    table_2025.loc[i, "Avos 2025"] = avos
                else:
                    avos = contar_avos(data_inicio_ano, data_final)
                    table_2025.loc[i, "Avos Parte 1"] = avos
                    table_2025.loc[i, "Avos Parte 2"] = 0
                    table_2025.loc[i, "Avos 2025"] = avos

            elif situacao == "F":
                avos1 = contar_avos(data_inicio_ano, ultimo_ativo) if pd.notna(ultimo_ativo) and ultimo_ativo >= data_inicio_ano else 0
                avos2 = contar_avos(retorno, data_final) if pd.notna(retorno) else 0
                table_2025.loc[i, "Avos Parte 1"] = avos1
                table_2025.loc[i, "Avos Parte 2"] = avos2
                table_2025.loc[i, "Avos 2025"] = avos1 + avos2

        # Exporta planilha final com os avos
        novo_arquivo_avos = os.path.splitext(caminho_arquivo)[0] + "_avos_calculados.xlsx"
        table_2025.to_excel(novo_arquivo_avos, index=False)
        messagebox.showinfo("Finalizado", f"Cálculo finalizado e salvo em:\n{novo_arquivo_avos}")

    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro:\n{e}")


################ Interface Gráfica ################

janela = tk.Tk()
janela.geometry("400x220")
janela.title("Calcular os faltantes compulsórios")

botao = tk.Button(janela, text="Procurar", command=processamento)
botao.pack(pady=10)

janela.mainloop()
