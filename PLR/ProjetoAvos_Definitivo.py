import pandas as pd
from pandas.tseries.offsets import MonthEnd
import tkinter as tk
from tkinter import filedialog, messagebox
import os
from datetime import datetime

# Função para contar avos válidos no ano de 2025
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


################ Função principal de processamento ################

def processar():
    data_input = entrada_data.get()

    try:
        data = pd.to_datetime(data_input, dayfirst=True)
    except Exception:
        messagebox.showerror("Erro", "Data inválida. Use o formato DD/MM/AAAA ou DD-MM-AA")
        return

    caminho_arquivo = filedialog.askopenfilename(
        title="Selecione o arquivo Excel",
        filetypes=[("Arquivos Excel", "*.xlsx *.xlsm *.xls")]
    )

    if not caminho_arquivo:
        messagebox.showwarning("Aviso", "Nenhum arquivo foi selecionado.")
        return

    try:
       
        # Lendo a aba "Geral" para o cálculo de Avos Mudar depois para ficar sem ABA  # Lendo a aba "Geral" para o cálculo de Avos Mudar depois para ficar sem ABA
        
        # Lendo a aba "Geral" para o cálculo de Avos Mudar depois para ficar sem ABA  # Lendo a aba "Geral" para o cálculo de Avos Mudar depois para ficar sem ABA
        
        table = pd.read_excel(caminho_arquivo, sheet_name="Geral")
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
            situacao = str(row["Situação"]).strip().upper()
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
                elif admissao <= data_inicio_ano:
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
                    elif afastamento >= data_inicio_ano and pd.isna(retorno):
                        parte1 = contar_avos(data_inicio_ano, afastamento)
                        parte2 = 0
                        table_2025.loc[i, "Avos Parte 1"] = parte1
                        table_2025.loc[i, "Avos Parte 2"] = parte2
                        table_2025.loc[i, "Avos 2025"] = parte1 + parte2
                    elif ultimo_ativo < data_inicio_ano and afastamento >= data_inicio_ano and retorno >= data_inicio_ano:
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
                    elif pd.notna(ultimo_ativo) and pd.notna(afastamento) and afastamento >= data_inicio_ano and pd.isna(retorno):
                        parte1 = contar_avos(data_inicio_ano, afastamento)
                        parte2 = 0
                        table_2025.loc[i, "Avos Parte 1"] = parte1
                        table_2025.loc[i, "Avos Parte 2"] = parte2
                        table_2025.loc[i, "Avos 2025"] = parte1 + parte2

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

        table_2025["Dias Afastados"] = dias_afastados

        colunas_avos = [
            "Chapa", "Nome", "Admis.", "Situação",
            "Ultimo dia Ativo", "Afastamento", "Retor.",
            "Dias Afastados", "Avos Parte 1", "Avos Parte 2", "Avos 2025"
        ]
        df_avos = table_2025[colunas_avos]

        # --- Parte de Faltantes/Duplicatas ---
        # Garantir que a coluna 'Chapa' é tratada como número inteiro para a verificação de duplicatas
        table["Chapa"] = pd.to_numeric(table["Chapa"], errors="coerce").astype('Int64')
        
        # Filtra as chapas que aparecem mais de uma vez
        duplicatas = table[table.duplicated(subset='Chapa', keep=False)]
        
        if duplicatas.empty:
            messagebox.showinfo("Sem duplicatas", "Nenhuma chapa repetida foi encontrada.")
           # Não há 'else' aqui, pois a lógica de salvar é feita mais abaixo
        
        # Abre a janela para o usuário escolher onde salvar o NOVO arquivo Excel com as duas abas
        nome_base, extensao = os.path.splitext(caminho_arquivo)
        sugestao_saida = f"{nome_base}_Calculado{extensao}" # Sugere um nome para o novo arquivo

        caminho_saida_final = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Arquivos Excel", "*.xlsx")],
            title="Salvar resultado como:",
            initialfile=os.path.basename(sugestao_saida) # Define o nome inicial do arquivo na caixa de diálogo
        )

        if not caminho_saida_final:
            messagebox.showwarning("Cancelado", "Operação de salvar arquivo cancelada.")
            return

        # Usa ExcelWriter para salvar múltiplas abas
        with pd.ExcelWriter(caminho_saida_final, engine='xlsxwriter') as writer:
            df_avos.to_excel(writer, sheet_name="Cálculo de Avos", index=False)

            if not duplicatas.empty:
                duplicatas.to_excel(writer, sheet_name="Faltantes", index=False)
                messagebox.showinfo("Sucesso", f"Arquivo salvo com sucesso em:\n{caminho_saida_final}\n\n"
                                               f"Aba 'Faltantes' criada com chapas repetidas.")
            else:
                messagebox.showinfo("Sucesso", f"Arquivo salvo com sucesso em:\n{caminho_saida_final}\n\n"
                                               f"Nenhuma chapa repetida encontrada, aba 'Faltantes' não criada.")

    except PermissionError:
        messagebox.showerror("Permissão Negada", "Você deve fechar o arquivo Excel de destino se ele estiver aberto.")
    except KeyError as e:
        messagebox.showerror("Erro de Coluna", f"Verifique se as colunas necessárias existem na aba 'Geral' do Excel. Erro: {e}")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro inesperado ao processar o arquivo:\n{e}")


################ Interface gráfica ################

def limpar_placeholder(event):
    
    # Criei variável para colocar a data de Hoje (data do teste)
    data_hoje = datetime.today().strftime('%d/%m/%Y')
    if entrada_data.get() == "dd/mm/aaaa":
        entrada_data.delete(0, tk.END)
        entrada_data.insert(0,data_hoje)

janela = tk.Tk()
janela.geometry("400x220")
janela.title("Cálculo de Avos powered by Mobi")

tk.Label(janela, text="Cálculo de Avos", font=("Helvetica", 14, "bold")).pack(pady=(10, 15))

entrada_data = tk.Entry(janela, width=15, font=("Helvetica", 13))

entrada_data.insert(0, "dd/mm/aaaa")
entrada_data.bind("<FocusIn>", limpar_placeholder)
entrada_data.pack()

tk.Label(janela, text="Escolha o arquivo em Excel:", font=("Helvetica", 12, "bold")).pack(pady=(20, 5))

botao = tk.Button(janela, text="Calcular", command=processar, font=("Helvetica", 10, "bold"))
botao.pack(pady=10)

janela.mainloop()