import pandas as pd
from pandas.tseries.offsets import MonthEnd
import tkinter as tk
from tkinter import filedialog, messagebox

# Função para contar avos válidos no ano de 2025

# Se as datas forem inválidas (vazias) ou a data final for antes de 2025, retorna 0 avos.
def contar_avos(inicio, fim):
    # Primeiro, vamos verificar se as datas são válidas. Se a data de início ou a data de fim estiverem vazias,
    # ou se a data de fim for antes do ano de 2025, então não podemos contar nenhum avo e o resultado é zero (retorna 0 ).
    if pd.isna(inicio) or pd.isna(fim) or fim < pd.Timestamp("2025-01-01"):
        return 0

    avos = 0  # vai contar quantos avos a pessoa tem

    # Verificar todos os 12 meses do ano de 2025
    for mes in range(1, 13):
       
        # Para cada mês, vamos descobrir qual é o primeiro dia, exemplo, para janeiro, seria "2025-01-01"
        data_texto = f"2025-{mes:02d}-01"
        
# Chamada "f-string" (string formatada). Ela permite colocar valores de variáveis dentro de um texto de maneira fácil.
# "2025-": Essa é a primeira parte do texto, o ano de 2025, seguido por um traço.
# {mes:02d}: parte aonde o valor da nossa variável mes vai ser colocado dentro do texto.
# mes: É o valor atual do mês (1, 2, 3, ..., 12) que o nosso for está usando no momento.
# :02d: Instrução de formatação, diz para o Python pegar o valor de mes e formatá-lo como um número inteiro:
# (d). O :02 significa que esse número deve sempre ter dois dígitos. Se o número for menor que 10 (como 1, 2, ..., 9), 
# ele vai adicionar um zero na frente (01, 02, ..., 09). Se for 10, 11 ou 12, ele já tem dois dígitos.
        
        # É uma função do pandas usada para criar objetos de data e hora.  Pega o texto que representa o primeiro dia do mês
        # (que está guardado em data_texto) e o transforma em um formato especial de data que o Python (com a ajuda do pandas) 
        # consegue entender e usar para fazer cálculos com datas.
        inicio_mes = pd.Timestamp(data_texto)
        
         # Vamos descobrir qual é o último dia desse mesmo mês, exemplo, para janeiro seria 31/01, para fevereiro 28/02
        fim_mes = inicio_mes + MonthEnd(0)
        
        # Verificar se o período que estamos analisando (do início ao fim) tem alguma parte dentro desse mês de 2025.
        # Se a data de fim for antes do COMEÇO do mês, OU se a data de início for DEPOIS do fim do mês,
        # significa que esse mês não tem nada a ver com o nosso período, então vamos para o próximo mês.
        if fim < inicio_mes or inicio > fim_mes:
            continue   # Ir para o próximo mês
        
        # Se o nosso período tem alguma parte dentro desse mês, precisamos ver qual é o pedacinho exato que está dentro do mês.
        # O "real_inicio" é o mais tarde entre o começo do mês e a nossa data de início.
        real_inicio = max(inicio_mes, inicio)
        
        # O "real_fim" é o mais cedo entre o fim do mês e a nossa data de fim.
        real_fim = min(fim_mes, fim)
        
         # Agora, vamos contar quantos dias tem esse pedacinho dentro do mês.
        dias = (real_fim - real_inicio).days + 1
        if dias >= 15:
            avos += 1
    return avos


################ Função principal de processamento ################

def processar():
     
     # Pegar a data que o usuário digitou na tela, na parte do Tkinter
    data_input = entrada_data.get()

    # Transformar a data o que o usuário digitou em uma data de verdade.
    # dayfirst - Indica que o dia vem primeiro, avisando que a data está no formato Dia/Mês/Ano (DD/MM/AAAA). E não
    # no formato americano: MM/DD/AAAA
    try:
        data = pd.to_datetime(data_input, dayfirst=True)
        
    # Se o que o usuário digitou não for uma data válida, o programa vai dar um erro. Vamos "pegar" esse 
    # erro e mostrar uma mensagem para o usuário. 
    except Exception:
        messagebox.showerror("Erro", "Data inválida. Use o formato DD/MM/AAAA.")
        return

    # Essa parte do código em Python utiliza a biblioteca tkinter (através do módulo filedialog) para abrir
    # uma janela de diálogo que permite ao usuário selecionar um arquivo Excel.
    caminho_arquivo = filedialog.askopenfilename(
        
    #  filedialog: É uma função do Tkinter que abre a janela "Abrir arquivo".
    # .askopenfilename(): Essa é uma função do módulo filedialog (que faz parte do tkinter) e significa literalmente:
    # "perguntar qual arquivo abrir".
        title="Selecione o arquivo Excel",
        filetypes=[("Arquivos Excel", "*.xlsx *.xlsm *.xls")]
    )
    # filetypes=[("Arquivos Excel", "*.xlsx *.xlsm *.xls")] - argumento opcional e  útil, pois especifica os tipos de arquivos que 
    # serão mostrados no filtro da janela de diálogo. Cria uma lista contendo uma tupla.
    # Este é o padrão de arquivo que será usado para filtrar os arquivos mostrados. O asterisco (*) é um caractere curinga que significa
    # "qualquer coisa". Portanto, *.xlsx significa "qualquer arquivo com a extensão .xlsx"...

    #  se o usuário não selecionou nenhum arquivo, esta linha fará com que a função atual termine sua execução imediatamente
    if not caminho_arquivo:
        return
# Poderia ser assim também: if caminho_arquivo == "":
#                                  return 

################ LENDO O EXCEL  ################
    try:
        
        # CUIDADO COM O SHEET  CUIDADO COM O SHEET  CUIDADO COM O SHEET  CUIDADO COM O SHEET 
        
        # CUIDADO COM O SHEET  CUIDADO COM O SHEET  CUIDADO COM O SHEET  CUIDADO COM O SHEET 
        
        # CUIDADO COM O SHEET  CUIDADO COM O SHEET  CUIDADO COM O SHEET  CUIDADO COM O SHEET 
        
        # CUIDADO COM O SHEET  CUIDADO COM O SHEET  CUIDADO COM O SHEET  CUIDADO COM O SHEET 
        
        # Funçaõ do Pandas que faz a leitura da planilha em excel, e cria uma estrutura de dados chamada DataFrame. Nesse caso seleciona a aba com nome Geral
        table = pd.read_excel(caminho_arquivo, sheet_name="Geral")
        
        # Acessa a lista de nomes das colunas do DataFrame table
        # str - transforma em String
        # strip() - retira espaçoes em branco, pois as vezes as colunas podem ter espaços no início ou no final do nome, e isso pode causar problemas em manipulações futuras.
        table.columns = table.columns.str.strip()

        # Temos que tratar as datas para poder trabalhar com elas, como anteriormente usamos o pandas.read
        # sabemos que virou um DATAFRAME assim sendo temos que usar os módulos/funções do Pandas, logo usaremos
        # pd.to_datetime pq ele trata com colunas diferente do Timestamp que trata valores individuais. 
        
        # Converte os valores das colunas para o tipo datetime do Pandas (representando datas e horas)e tratando
        # de valores que não podem ser convertidos (inválidos) para NAN/NAT
        table["Afastamento"] = pd.to_datetime(table["Afastamento"], errors='coerce')
        table["Ultimo dia Ativo"] = pd.to_datetime(table["Ultimo dia Ativo"], errors='coerce')
        table["Retor."] = pd.to_datetime(table["Retor."], errors='coerce')
        table["Admis."] = pd.to_datetime(table["Admis."], errors='coerce')


       # FILTRAGEM - filtra as linhas do DATAFRAME tabel para as colunas de datas de acordo com o ano de 2025 e a Situação de ATIVO (A). 
        table_2025 = table[
            (table["Afastamento"].dt.year == 2025) |
            (table["Retor."].dt.year == 2025) |
            (table["Ultimo dia Ativo"].dt.year == 2025)
        ].copy()
        # Vamos extrair o ano para cada coluna que tem as datas. 
        # copy() - criamos uma cópia para não modificar o DATAFRAME original evitando o aviso de "SettingWithCopyWarning"

        #  Inicializar três novas colunas no DataFrame table_2025.
        table_2025["Avos Parte 1"] = 0
        table_2025["Avos Parte 2"] = 0
        table_2025["Avos 2025"] = 0

               #########################################################################################################

        # i - número da linha (índice)
        # row -  informações da linha
        for i, row in table_2025.iterrows():
            
            # Essas variáveis estão apenas pegando os valores de cada linha das colunas : Essas variáveis estão apenas pegando os valores das colunas ,
            # Situação, Retor, Admis, Afastamento e Ultimo dia Ativo.
            situacao = row["Situação"]
            retorno = row["Retor."]
            admissao = row["Admis."]
            afastamento = row["Afastamento"]
            ultimo_ativo = row["Ultimo dia Ativo"]
            data_final = data
            data_incio_ano = pd.Timestamp("2025-01-01") # Coloquei para susbtituir nos IFs aninhados deixando mais claro
            
            if situacao == "A":
                # pd.notna(retorno) - verifica se a variável retorno não é nula
                # pd.notna(ultimo_ativo) - verifica se a variável ultimo_ativo não é nula
                # retorno >= data_incio_ano -  verifica se as linhas da coluna retorno são maiores ou iguais a ao inicio do ano 01/01/2025
                if pd.notna(retorno) and pd.notna(ultimo_ativo) and retorno >= data_incio_ano:
                    
                    # A parte1 vai chamar a função contar_avos colocando a data de Início do ano ano 2025 e Último dia Ativo
                    parte1 = contar_avos(data_incio_ano, ultimo_ativo) 
                   
                    # A parte2 vai chamar a função contar_avos colocando a data de Retorno e a data escolhida pelo user
                    parte2 = contar_avos(retorno, data_final)
                 
                    # .loc[i, "coluna"] - estamos dizendo linha i, coluna "coluna"
                    table_2025.loc[i, "Avos Parte 1"] = parte1
                    table_2025.loc[i, "Avos Parte 2"] = parte2
                    table_2025.loc[i, "Avos 2025"] = parte1 + parte2
                # pd.notna(admissao) -  verifica se a variável não é nula
                # admissao >= data_incio_ano -  se a data de admissão for maior que a data de de Início do ano 2025
                elif pd.notna(admissao) and admissao >= data_incio_ano:
                    avos = contar_avos(admissao, data_final)
                    table_2025.loc[i, "Avos Parte 1"] = avos
                    table_2025.loc[i, "Avos Parte 2"] = 0
                    table_2025.loc[i, "Avos 2025"] = avos

                else:
                    avos = contar_avos(data_incio_ano, data_final)
                    table_2025.loc[i, "Avos Parte 1"] = avos
                    table_2025.loc[i, "Avos Parte 2"] = 0
                    table_2025.loc[i, "Avos 2025"] = avos
                    
# Aqui faz as condições para quando a situação for Afastado (F)
            elif situacao == "F":
                if pd.notna(ultimo_ativo) and ultimo_ativo >= data_incio_ano:
                    avos1 = contar_avos(data_incio_ano, ultimo_ativo)
                    table_2025.loc[i, "Avos Parte 1"] = avos1
                else:
                    avos1 = 0

                if pd.notna(retorno):
                    avos2 = contar_avos(retorno, data_final)
                else:
                    avos2 = 0

                table_2025.loc[i, "Avos Parte 2"] = avos2
                table_2025.loc[i, "Avos 2025"] = avos1 + avos2


################ DIAS AFASTADOS ################

        # Criação de Lista ( não foi dicionário e nem tuplas )
        dias_afastados = []
        
        for i, row in table_2025.iterrows():
            retorno = row["Retor."]
            ultimo_ativo = row["Ultimo dia Ativo"]
            
            # Se o último dia Ativo for nulo , caso sim o núemro de dias = none
            if pd.isna(ultimo_ativo):
                dias = None
                # Se o último dia Ativo não for nulo e a data de Retorno for nula logo
                # dias = data escolhida do user - ultimo dia ativo
            elif pd.isna(retorno):
                dias = (data - ultimo_ativo).days
                # Se tiver o último dia Ativo não nulo e a data de retorno não nula calcula entre elas
            else:
                dias = (retorno - ultimo_ativo).days
            dias_afastados.append(dias)

        table_2025["Dias Afastados"] = dias_afastados

        colunas = [
            "Chapa", "Nome", "Admis.", "Situação",
            "Ultimo dia Ativo", "Afastamento", "Retor.",
            "Dias Afastados", "Avos Parte 1", "Avos Parte 2", "Avos 2025"
        ]

        resultado = table_2025[colunas]

        saida = caminho_arquivo.replace(".xlsm", "_RESULTADO.xlsx").replace(".xlsx", "_RESULTADO.xlsx")
        resultado.to_excel(saida, index=False)

        messagebox.showinfo("Sucesso", f"Arquivo exportado para:\n{saida}")

    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao processar o arquivo:\n{e}")


################ Interface gráfica ################

janela = tk.Tk()
janela.geometry("400x220")
janela.title("Cálculo de Avos 2025")

tk.Label(janela, text="Digite a data que deseja (Dia/Mês/Ano):").pack(pady=(20, 5))
entrada_data = tk.Entry(janela, width=20)
entrada_data.pack()

tk.Label(janela, text="Clique abaixo para escolher o arquivo Excel:").pack(pady=(20, 5))

botao = tk.Button(janela, text="Calcular Avos", command=processar)
botao.pack(pady=10)

janela.mainloop()