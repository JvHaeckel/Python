import pandas as pd
from pandas.tseries.offsets import MonthEnd
import tkinter as tk
from tkinter import filedialog, messagebox

# Função para contar avos válidos no ano de 2025
def contar_avos(inicio, fim):
    if pd.isna(inicio) or pd.isna(fim) or fim < pd.Timestamp("2025-01-01"):
        return 0

    meses = 0
    for mes in range(1, 13):
        inicio_mes = pd.Timestamp(f"2025-{mes:02d}-01")
        fim_mes = inicio_mes + MonthEnd(0)
        if fim < inicio_mes or inicio > fim_mes:
            continue
        real_inicio = max(inicio_mes, inicio)
        real_fim = min(fim_mes, fim)
        dias = (real_fim - real_inicio).days + 1
        if dias >= 15:
            meses += 1
    return meses

# Função principal de processamento
def processar():
    data_input = entrada_data.get()

    try:
        data = pd.to_datetime(data_input, dayfirst=True)
    except Exception:
        messagebox.showerror("Erro", "Data inválida. Use o formato DD/MM/AAAA.")
        return

    caminho_arquivo = filedialog.askopenfilename(
        title="Selecione o arquivo Excel",
        filetypes=[("Arquivos Excel", "*.xlsx *.xlsm *.xls")]
    )

    if not caminho_arquivo:
        return

    try:
        table = pd.read_excel(caminho_arquivo, sheet_name="Geral")
        table.columns = table.columns.str.strip()  # Remove espaços extras nos nomes das colunas

        table["Afastamento"] = pd.to_datetime(table["Afastamento"], errors='coerce')
        table["Ultimo dia Ativo"] = pd.to_datetime(table["Ultimo dia Ativo"], errors='coerce')
        table["Retor."] = pd.to_datetime(table["Retor."], errors='coerce')

        table_2025 = table[
            (table["Afastamento"].dt.year == 2025) |
            (table["Retor."].dt.year == 2025) |
            (table["Ultimo dia Ativo"].dt.year == 2025) |
            (table["Situação"] == "A")  # Inclui os ativos
        ].copy()

        table_2025["Avos Parte 1"] = 0
        table_2025["Avos Parte 2"] = 0

        for i, row in table_2025.iterrows():
            situacao = row["Situação"]
            retorno = row["Retor."]

            if situacao == "A" and retorno == 0:
                avos = contar_avos(pd.Timestamp("2025-01-01"), data)
                table_2025.at[i, "Avos 2025"] = avos
                table_2025.at[i, "Avos Parte 1"] = avos
                table_2025.at[i, "Avos Parte 2"] = 0
            elif situacao == "F":
                ultimo_ativo = row["Ultimo dia Ativo"]
                if pd.notna(ultimo_ativo) and ultimo_ativo >= pd.Timestamp("2025-01-01"):
                    avos1 = contar_avos(pd.Timestamp("2025-01-01"), ultimo_ativo)
                    table_2025.at[i, "Avos Parte 1"] = avos1
                else:
                    avos1 = 0

                retorno = row["Retor."]
                if pd.notna(retorno):
                    avos2 = contar_avos(retorno, data)
                else:
                    avos2 = 0
                table_2025.at[i, "Avos Parte 2"] = avos2
                table_2025.at[i, "Avos 2025"] = avos1 + avos2

        dias_afastados = []
        for i, row in table_2025.iterrows():
            retorno = row["Retor."]
            ultimo_ativo = row["Ultimo dia Ativo"]
            if pd.isna(ultimo_ativo):
                dias = None
            elif pd.isna(retorno):
                dias = (data - ultimo_ativo).days
            else:
                dias = (retorno - ultimo_ativo).days
            dias_afastados.append(dias)

        table_2025["Dias"] = dias_afastados

        colunas = [
            "Chapa", "Nome", "Admis.", "Situação",
            "Ultimo dia Ativo", "Afastamento", "Retor.",
            "Dias", "Avos Parte 1", "Avos Parte 2", "Avos 2025"
        ]

        resultado = table_2025[colunas]

        saida = caminho_arquivo.replace(".xlsm", "_RESULTADO.xlsx").replace(".xlsx", "_RESULTADO.xlsx")
        resultado.to_excel(saida, index=False)

        messagebox.showinfo("Sucesso", f"Arquivo exportado para:\n{saida}")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao processar o arquivo:\n{e}")

# === Interface gráfica ===
janela = tk.Tk()
janela.geometry("400x220")
janela.title("Cálculo de Avos 2025")

tk.Label(janela, text="Digite a data de referência (DD/MM/AAAA):").pack(pady=(20, 5))
entrada_data = tk.Entry(janela, width=20)
entrada_data.pack()

tk.Label(janela, text="Depois clique abaixo para escolher o arquivo Excel:").pack(pady=(20, 5))

botao = tk.Button(janela, text="Selecionar Arquivo e Calcular", command=processar)
botao.pack(pady=10)

janela.mainloop()
