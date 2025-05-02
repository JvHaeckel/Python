
name = input("Qual seu nome?")
idade = input("Qual sua idade? ")

print("Oi " + name + ", sua idade é " + idade)


print("Comprimento do seu nome é: " + str(len(name)))

# O Python apresenta um erro porque o operador + só pode ser usado para somar dois números inteiros 
# ou concatenar duas strings. Você não pode somar um número inteiro a uma string, pois isso é agramatical 
# em Python. Você pode corrigir isso usando uma versão em string do número inteiro

# As funções str(), int() e float() são é úteis quando você tem um inteiro ou float que deseja concatenar
# a outras palavras

