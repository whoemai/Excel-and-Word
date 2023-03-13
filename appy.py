import os
import tkinter as tk
from datetime import datetime
from tkinter import messagebox
import pandas as pd
from docx import Document

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
    dados[nome] = [
        input(f"Insira o dado referente a {nome} {i+1}: ") for i in range(n)
    ]

# Criar um dataframe com os dados
df = pd.DataFrame(dados, columns=nomesColunas)

# Define o caminho e nome do arquivo com a data do dia
today = datetime.today().strftime('%Y-%m-%d')
filepath_excel = f'C:\\Users\\x\\Documents\\Export Python\\appy-PY{today}.xlsx'

# Salva os arquivos permitindo substituir os existentes
if os.path.exists(filepath_excel):
    confirm = messagebox.askyesno("Arquivo existente",
                                  f"O arquivo {filepath_excel} já existe. Deseja substituí-lo?")
    if not confirm:
        messagebox.showinfo("Arquivo não salvo",
                            f"O arquivo {filepath_excel} não foi salvo.")
    else:
        df.to_excel(filepath_excel, index=False)
else:
    df.to_excel(filepath_excel, index=False)

print("Excel salvo com sucesso!")
