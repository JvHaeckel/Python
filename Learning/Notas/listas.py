cats = []

while True:
    
    nome = input(f"Digite o nome do gato {len(cats) + 1} : ")
    cats = cats + [nome]
    # ou 
    # cats.append(nome)
    
    resposta = input("Deseja continuar? Sim ou não")
    resposta = resposta.lower()
    
    if resposta == 'sim' :
       continue
    else: print("Acabou")
    break
    
    
    #Exibir os nomes dos gatos: 
    
    print(cats[])