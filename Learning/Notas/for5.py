
# Percorrendo um dicionário (chaves e valores)
# Por que isso é útil?
# Acessar e trabalhar com cada dado dentro de um dicionário sem precisar saber antecipadamente quais são as chaves.

bus = {
    'linha' : ['2050', '2040' , '2060' ],
    'origem': [ 'TI Camaragibe', 'TI Várzea', 'TI Abreu'],
    'destino': ['São Lourenço', 'Federal', 'Litoral']
}

for chave, valor in bus.items():
   
        print(chave , ':' , valor) 





