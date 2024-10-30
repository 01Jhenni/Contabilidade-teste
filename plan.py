import pandas as pd

# Carregando os dados das duas planilhas
funcionarios_sul = pd.read_excel('Z:\\INFORMATICA\\Testes Notas\\aplicar.xlsx')
funcionarios_norte = pd.read_excel('Z:\\INFORMATICA\\Testes Notas\\catho.xlsx')

# Realizando a união dos dados
funcionarios_consolidado = pd.concat([funcionarios_sul, funcionarios_norte])

# Exibe os dados unidos
print(funcionarios_consolidado)

# Salva os dados unidos em uma nova planilha
funcionarios_consolidado.to_excel('Z:\\INFORMATICA\\Testes Notas\\nova.xlsx', index=False)
