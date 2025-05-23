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
        messagebox.showwarning("Aviso", "Nenhum arquivo foi selecionado")
        return
  # Poderia ser assim também: if caminho_arquivo == "":
  #                                  return 

   ################ LENDO O EXCEL  ################
    try:
        
        # ATENÇÃO: Verifique se o nome da aba está correto (neste caso, "Geral")

        # A função read_excel do pandas lê a planilha do Excel e retorna um DataFrame
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

        table_2025 = table.copy()
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
            situacao = str(row["Situação"]).strip().upper()  # Converteu para String e deixou em caixa alta
            retorno = row["Retor."]
            admissao = row["Admis."]
            afastamento = row["Afastamento"]
            ultimo_ativo = row["Ultimo dia Ativo"]
            data_final = data
            data_inicio_ano = pd.Timestamp("2025-01-01") # Coloquei para susbtituir nos IFs aninhados deixando mais claro
            data_fim_ano = pd.Timestamp("2025-12-31")

            # Calculando primeiramente para os ATIVOS 
            if situacao == "A":

                # ATIVOS admitidos em 2025
                if admissao >= data_inicio_ano:

                    # ATIVOS admitidos em 2025 que não tiveram Último dia Ativo(nulo/vazio/zero) e Afastamento(nulo/vazio/zero)
                    if pd.isna(ultimo_ativo) and pd.isna(afastamento):

                        # A parte1 calcula da Admissão de 2025 até a data final escolhida pelo user.
                        # A parte2 é nulo/vazio/zero porque não teve retorno.

                        # Avos da admissão de 2025 até a data do input do usuário
                        parte1 = contar_avos(admissao, data_final) 
                        parte2 = 0
                        table_2025.loc[i, "Avos Parte 1"] = parte1
                        table_2025.loc[i, "Avos Parte 2"] = parte2
                        table_2025.loc[i, "Avos 2025"] = parte1 + parte2

                    # ATIVOS admitidos em 2025 que tiveram afastamento com retorno no ano de 2025
                    elif pd.notna(ultimo_ativo) and pd.notna(afastamento) and pd.notna(retorno) and data_inicio_ano <= retorno <= data_fim_ano:

                        parte1 = contar_avos(admissao, afastamento)
                        parte2 = contar_avos(retorno, data_final)
                        table_2025.loc[i, "Avos Parte 1"] = parte1
                        table_2025.loc[i, "Avos Parte 2"] = parte2
                        table_2025.loc[i, "Avos 2025"] = parte1 + parte2

                    # ATIVOS admitidos em 2025 que tiveram Afastamento em 2025 sem Retorno(nulo/vazio/zero) até o momento
                    elif afastamento >= data_inicio_ano and pd.isna(retorno):

                        parte1 = contar_avos(admissao, afastamento)
                        parte2 = 0
                        table_2025.loc[i, "Avos Parte 1"] = parte1
                        table_2025.loc[i, "Avos Parte 2"] = parte2
                        table_2025.loc[i, "Avos 2025"] = parte1 + parte2

                # ATIVOS admitidos ANTES de 2025
                elif  admissao <= data_inicio_ano:

                    # ATIVOS admitidos ANTES de 2025 que não tiveram Último dia Ativo e Afastamento                           ******** OK ********
                    if pd.isna(ultimo_ativo) and pd.isna(afastamento):

                        # Avos da data de Início do ano de 2025 até a data do input do usuário
                        parte1 = contar_avos(data_inicio_ano, data_final) 
                        parte2 = 0
                        table_2025.loc[i, "Avos Parte 1"] = parte1
                        table_2025.loc[i, "Avos Parte 2"] = parte2
                        table_2025.loc[i, "Avos 2025"] = parte1 + parte2
                        
                    # ATIVOS admitidos ANTES de 2025 que tiveram Afastamento com Retorno no ano de 2025 
                    elif pd.notna(ultimo_ativo) and pd.notna(afastamento) and pd.notna(retorno) and retorno >= data_inicio_ano:

                        # ATIVOS admitidos ANTES de 2025 que tiveram Afastamento ANTES de 2025 com Retorno no ano de 2025  ******** Não existiu essa condição : sem afastados antes de 2025********
                        if afastamento <= data_inicio_ano:

                            parte1 = 0
                            parte2 = contar_avos(retorno, data_final)
                            table_2025.loc[i, "Avos Parte 1"] = parte1
                            table_2025.loc[i, "Avos Parte 2"] = parte2
                            table_2025.loc[i, "Avos 2025"] = parte1 + parte2

                        # ATIVOS admitidos ANTES de 2025 que tiveram Afastamento em 2025 com Retorno no ano de 2025         ******** OK ********
                        elif afastamento >= data_inicio_ano:

                            parte1 = contar_avos(data_inicio_ano, afastamento)
                            parte2 = contar_avos(retorno, data_final)
                            table_2025.loc[i, "Avos Parte 1"] = parte1
                            table_2025.loc[i, "Avos Parte 2"] = parte2
                            table_2025.loc[i, "Avos 2025"] = parte1 + parte2
                            
                         # ATIVOS admitidos ANTES de 2025 que tiveram Afastamento em 2025 sem Retorno(0) no ano de 2025      ******** OK ********     
                    elif  afastamento >= data_inicio_ano and pd.isna(retorno):
                        
                        parte1 = contar_avos(data_inicio_ano, afastamento)
                        parte2 = 0
                        table_2025.loc[i, "Avos Parte 1"] = parte1
                        table_2025.loc[i, "Avos Parte 2"] = parte2
                        table_2025.loc[i, "Avos 2025"] = parte1 + parte2
                        
                        # Ativos admitidos antes de 2025 que tiveram seu último dia ativo antes de 2025 com Afastamento em 2025 e Retorno em 2025  ******** OK ********
                    elif ultimo_ativo < data_inicio_ano and afastamento >= data_inicio_ano and retorno >= data_inicio_ano:
                        
                        parte1 = contar_avos(data_inicio_ano, afastamento)
                        parte2 = contar_avos(retorno, data_final)
                        table_2025.loc[i, "Avos Parte 1"] = parte1
                        table_2025.loc[i, "Avos Parte 2"] = parte2
                        table_2025.loc[i, "Avos 2025"] = parte1 + parte2

                ########################### Calculando para os AFASTADOS ############################################         
            
            elif situacao == 'F':

                # AFASTADOS admitidos EM 2025
                
                if pd.notna(admissao) and admissao >= data_inicio_ano:
                    
                    # AFASTADOS admitidos em 2025 com Afastamento em 2025 e sem Retorno                                        ******** OK ********
                    if pd.notna(ultimo_ativo) and pd.notna(afastamento) and pd.isna(retorno):
                        
                        # A parte1 calcula da admissão até o Afastamento 
                        # A parte2 calcula do Retorno até a data que o user pediu, mas como não existe retorno ficará zerado
                        parte1 = contar_avos(admissao, afastamento)
                        parte2 = 0
                        table_2025.loc[i, "Avos Parte 1"] = parte1
                        table_2025.loc[i, "Avos Parte 2"] = parte2
                        table_2025.loc[i, "Avos 2025"] = parte1 + parte2
                        
                    else:  # AFASTADOS admitidos em 2025 com Afastamento e Retorno em 2025            ******** OK - mas até o presente não tinham pessoas nessa condição ********
                        
                        parte1 = contar_avos(admissao, afastamento)
                        parte2 = contar_avos(retorno, data_final)
                        table_2025.loc[i, "Avos Parte 1"] = parte1
                        table_2025.loc[i, "Avos Parte 2"] = parte2
                        table_2025.loc[i, "Avos 2025"] = parte1 + parte2
                        
                   # AFASTADOS admitidos ANTES de 2025     

                   # Poderia colocar isso mais aconselhável também: elif  admissao <= data_inicio_ano:
                else:

                    # AFASTADOS admitidos ANTES de 2025 que tiveram afastamento no ano de 2025 com retorno no ano de 2025   ******** OK ********
                    if (pd.notna(ultimo_ativo) and pd.notna(afastamento) and 
                        afastamento >= data_inicio_ano and pd.notna(retorno) and retorno >= data_inicio_ano):

                        parte1 = contar_avos(data_inicio_ano, afastamento)
                        parte2 = contar_avos(retorno, data_final)
                        table_2025.loc[i, "Avos Parte 1"] = parte1
                        table_2025.loc[i, "Avos Parte 2"] = parte2
                        table_2025.loc[i, "Avos 2025"] = parte1 + parte2
                        
                    # AFASTADOS admitidos ANTES de 2025 que tiveram afastamento no ano de 2025 sem  retorno no ano de 2025       
                    elif pd.notna(ultimo_ativo) and pd.notna(afastamento) and afastamento >= data_inicio_ano and pd.isna(retorno): # ******** OK ********

                        parte1 = contar_avos(data_inicio_ano, afastamento)
                        parte2 = 0
                        table_2025.loc[i, "Avos Parte 1"] = parte1
                        table_2025.loc[i, "Avos Parte 2"] = parte2
                        table_2025.loc[i, "Avos 2025"] = parte1 + parte2
       
    
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
        
        # Abre a janela para salvar o arquivo, forçando extensão .xlsx
        saida = filedialog.asksaveasfilename(
            defaultextension=".xlsx" ,
            filetypes=[("Arquivos Excel", "*.xlsx")],
            title="Salvar arquivo como:"
        )

        
        if saida:
            resultado.to_excel(saida, index=False)
            messagebox.showinfo("Sucesso", f"Arquivo salvo com sucesso em:\n{saida}")
        else:
            messagebox.showwarning("Cancelado", "Operação cancelada pelo usuário.")
          
          # Tratando erro de deixar arquivo do excel a ser salvo em Aberto.   
    except PermissionError:
        messagebox.showerror("Permisssão negada", "Você deve estar mantendo arquivo em excel de destino em aberto")
    
    except Exception as e:
            messagebox.showerror("Erro", f"Erro ao processar o arquivo:\n{e}")


################ Interface gráfica ################

# Função para limpar o placeholder: 
def limpar_placeholder(event):
    if entrada_data.get() == "dd/mm/aaaa":
        entrada_data.delete(0, tk.END)

janela = tk.Tk()
janela.geometry("400x220")
janela.title("Cálculo de Avos")

# Título maior e em negrito
tk.Label(janela, text="Cálculo de Avos 2025", font=("Helvetica", 14, "bold")).pack(pady=(10, 15))
entrada_data = tk.Entry(janela, width=20)
entrada_data.insert(0, "dd/mm/aaaa")  # Preenche o campo com o formato
entrada_data.bind("<FocusIn>", limpar_placeholder)  # Remove o placeholder ao clicar
entrada_data.pack()

tk.Label(janela, text="Escolha o arquivo apenas em Excel:", font=("Helvetica", 12, "bold")).pack(pady=(20, 5))

botao = tk.Button(janela, text="Calcular Avos", command=processar, font=("Helvetica", 10, "bold"))
botao.pack(pady=10)

janela.mainloop()