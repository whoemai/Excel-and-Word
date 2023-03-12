import pandas as pd

# Pedir ao usuário para inserir o número de dados
n = int(input("Insira o número de dados: "))

# Pedir ao usuário para inserir a quantidade de colunas desejadas
while True:
    quantidadeColuna = int(input("Digite a quantidade de colunas desejadas: "))
    if quantidadeColuna > 0:
        break
    else:
        print("A quantidade de colunas deve ser maior que zero.")

# Criar uma lista vazia para armazenar os nomes das colunas
nomesColunas = []

# Pedir ao usuário para inserir o nome de cada coluna
for i in range(quantidadeColuna):
    nomeColuna = input(f"Digite o nome da coluna {i+1}: ")
    nomesColunas.append(nomeColuna)

# Pedir ao usuário para inserir os dados
dados = {}
for nome in nomesColunas:
    dados[nome] = [input(f"Insira o dado referente a {nome} {i+1}: ") for i in range(n)]

# Criar um dataframe com os dados
df = pd.DataFrame(dados, columns=nomesColunas)

print("sucesso")
