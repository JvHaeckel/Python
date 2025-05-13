import pandas as pd 
import tkinter as tk
from tkinter import filedialog, messagebox
import os
from pandas.tseries.offsets import MonthEnd

# Função para contar avos com base nas regras de 15 dias por mês
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

# Função principal
def processamento():
    caminho_arquivo = filedialog.askopenfilename(
        title="Selecione arquivo em Excel apenas",
        filetypes=[("Arquivos Excel", "*.xlsx *.xlsm *.xls")]
    )

    if caminho_arquivo == "":
        return

    try:
        table = pd.read_excel(caminho_arquivo)

        # Garante que a coluna 'Chapa' é numérica
        table["Chapa"] = pd.to_numeric(table["Chapa"], errors="coerce").astype('Int64')

        # Filtra chapas repetidas
        duplicatas = table[table.duplicated(subset='Chapa', keep=False)].copy()

        if duplicatas.empty:
            messagebox.showinfo("Sem duplicatas", "Nenhuma chapa repetida foi encontrada.")
            return

        # Converte campos de data (com .loc para evitar o warning)
        duplicatas.loc[:, "Retor."] = pd.to_datetime(duplicatas["Retor."], errors="coerce")
        duplicatas.loc[:, "Ultimo dia Ativo"] = pd.to_datetime(duplicatas["Ultimo dia Ativo"], errors="coerce")

        # Filtra apenas os registros com datas relevantes para 2025
        duplicatas = duplicatas[
            (duplicatas["Retor."].dt.year >= 2025) |
            (duplicatas["Ultimo dia Ativo"].dt.year >= 2025)
        ].copy()

        # Lista para armazenar os avos por chapa
        avos_total = []

        for chapa, grupo in duplicatas.groupby("Chapa"):
            grupo = grupo.sort_values("Retor.")
            total_avos = 0

            for _, row in grupo.iterrows():
                retorno = row["Retor."]
                ultimo_ativo = row["Ultimo dia Ativo"]

                if pd.notna(retorno) and pd.notna(ultimo_ativo):
                    total_avos += contar_avos(retorno, ultimo_ativo)

            avos_total.append({"Chapa": chapa, "Avos 2025": total_avos})

        # Cria DataFrame com avos somados
        avos_df = pd.DataFrame(avos_total)

        # Junta os dados de duplicatas com os avos
        duplicatas_com_avos = duplicatas.merge(avos_df, on="Chapa", how="left")

        # Define colunas desejadas (ajuste conforme necessário)
        colunas = [
            "Chapa", "Nome", "Admis.", "Situação",
            "Ultimo dia Ativo", "Afastamento", "Retor.",
            "Dias Afastados", "Avos Parte 1", "Avos Parte 2", "Avos 2025"
        ]

        # Exporta tudo em um único arquivo com duas abas
        arquivo_saida = os.path.splitext(caminho_arquivo)[0] + "_resultado_final.xlsx"
        with pd.ExcelWriter(arquivo_saida, engine='xlsxwriter') as writer:
            duplicatas_com_avos.to_excel(writer, sheet_name="Duplicatas + Avos", index=False, columns=colunas)
            avos_df.to_excel(writer, sheet_name="Somatório Avos", index=False)

        messagebox.showinfo("Concluído", f"Arquivo salvo em:\n{arquivo_saida}")

    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro:\n{str(e)}")

# Interface Gráfica
janela = tk.Tk()
janela.geometry("400x220")
janela.title("Calcular os Avos dos faltantes compulsórios")

botao = tk.Button(janela, text="Selecionar arquivo Excel", command=processamento)
botao.pack(pady=20)

janela.mainloop()
