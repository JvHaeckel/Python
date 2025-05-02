# O truque da atribuição múltipla (tecnicamente chamado de desempacotamento de tuplas ) é um atalho que 
# permite atribuir a várias variáveis ​​os valores de uma lista em uma única linha de código. O número de 
# variáveis ​​e o comprimento da lista devem ser exatamente iguais, ou o Python retornará um ValueError 


gato = ['gordo', 'cheiroso', 'marmota']


print(gato[0])

corpo = gato[0]
odor = gato[1]
caracteristica = gato[2]

corpo, odor, caracteristica = gato
print(gato)

