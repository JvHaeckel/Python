import pandas as pd
from pandas.tseries.offsets import MonthEnd
import tkinter as tk
from tkinter import filedialog, messagebox

def contar_avos(inicio, fim):
    if pd.isna(inicio) or pd.isna(fim) or fim < pd.Timestamp("2025-01-01"):
        return 0
    avos = 0
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
            avos += 1
    return avos

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
    except ValueError:
        messagebox.showerror("ERRO", 'A aba "Geral" não foi encontrada na planilha.')

    table.columns = table.columns.str.strip()
    table["Afastamento"] = pd.to_datetime(table["Afastamento"], errors='coerce')
    table["Ultimo dia Ativo"] = pd.to_datetime(table["Ultimo dia Ativo"], errors='coerce')
    table["Retor."] = pd.to_datetime(table["Retor."], errors='coerce')
    table["Admis."] = pd.to_datetime(table["Admis."], errors='coerce')

    table_2025 = table.copy()
    table_2025["Avos Parte 1"] = 0
    table_2025["Avos Parte 2"] = 0
    table_2025["Avos 2025"] = 0

    for i, row in table_2025.iterrows():
        situacao = row["Situação"]
        retorno = row["Retor."]
        admissao = row["Admis."]
        afastamento = row["Afastamento"]
        ultimo_ativo = row["Ultimo dia Ativo"]
        data_final = data
        data_inicio_ano = pd.Timestamp("2025-01-01")
        data_fim_ano = pd.Timestamp("2025-12-31")

        if situacao == "A":
            if admissao >= data_inicio_ano:
                if pd.isna(ultimo_ativo) and pd.isna(afastamento):
                    parte1 = contar_avos(admissao, data_final)
                    parte2 = 0
                    table_2025.loc[i, "Avos Parte 1"] = parte1
                    table_2025.loc[i, "Avos Parte 2"] = parte2
                    table_2025.loc[i, "Avos 2025"] = parte1 + parte2
                elif pd.notna(ultimo_ativo) and pd.notna(afastamento) and pd.notna(retorno) and data_inicio_ano <= retorno <= data_fim_ano:
                    parte1 = contar_avos(admissao, afastamento)
                    parte2 = contar_avos(retorno, data_final)
                    table_2025.loc[i, "Avos Parte 1"] = parte1
                    table_2025.loc[i, "Avos Parte 2"] = parte2
                    table_2025.loc[i, "Avos 2025"] = parte1 + parte2
                elif afastamento >= data_inicio_ano and pd.isna(retorno):
                    parte1 = contar_avos(admissao, afastamento)
                    parte2 = 0
                    table_2025.loc[i, "Avos Parte 1"] = parte1
                    table_2025.loc[i, "Avos Parte 2"] = parte2
                    table_2025.loc[i, "Avos 2025"] = parte1 + parte2
            else:
                if pd.isna(ultimo_ativo) and pd.isna(afastamento):
                    parte1 = contar_avos(data_inicio_ano, data_final)
                    parte2 = 0
                    table_2025.loc[i, "Avos Parte 1"] = parte1
                    table_2025.loc[i, "Avos Parte 2"] = parte2
                    table_2025.loc[i, "Avos 2025"] = parte1 + parte2
                elif pd.notna(ultimo_ativo) and pd.notna(afastamento) and pd.notna(retorno) and retorno >= data_inicio_ano:
                    if afastamento <= data_inicio_ano:
                        parte1 = 0
                        parte2 = contar_avos(retorno, data_final)
                        table_2025.loc[i, "Avos Parte 1"] = parte1
                        table_2025.loc[i, "Avos Parte 2"] = parte2
                        table_2025.loc[i, "Avos 2025"] = parte1 + parte2
                    elif afastamento >= data_inicio_ano:
                        parte1 = contar_avos(data_inicio_ano, afastamento)
                        parte2 = contar_avos(retorno, data_final)
                        table_2025.loc[i, "Avos Parte 1"] = parte1
                        table_2025.loc[i, "Avos Parte 2"] = parte2
                        table_2025.loc[i, "Avos 2025"] = parte1 + parte2

        elif situacao == 'F':
            if pd.notna(admissao) and admissao >= data_inicio_ano:
                if pd.notna(ultimo_ativo) and pd.notna(afastamento) and pd.isna(retorno):
                    parte1 = contar_avos(admissao, afastamento)
                    parte2 = 0
                    table_2025.loc[i, "Avos Parte 1"] = parte1
                    table_2025.loc[i, "Avos Parte 2"] = parte2
                    table_2025.loc[i, "Avos 2025"] = parte1 + parte2
                else:
                    parte1 = contar_avos(admissao, afastamento)
                    parte2 = contar_avos(retorno, data_final)
                    table_2025.loc[i, "Avos Parte 1"] = parte1
                    table_2025.loc[i, "Avos Parte 2"] = parte2
                    table_2025.loc[i, "Avos 2025"] = parte1 + parte2
            else:
                if (pd.notna(ultimo_ativo) and pd.notna(afastamento) and 
                    afastamento >= data_inicio_ano and pd.notna(retorno) and retorno >= data_inicio_ano):
                    parte1 = contar_avos(data_inicio_ano, afastamento)
                    parte2 = contar_avos(retorno, data_final)
                    table_2025.loc[i, "Avos Parte 1"] = parte1
                    table_2025.loc[i, "Avos Parte 2"] = parte2
                    table_2025.loc[i, "Avos 2025"] = parte1 + parte2
