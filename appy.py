import os
import pandas as pd
from docx import Document
import tkinter as tk
from tkinter import messagebox
from datetime import datetime

# Pedir ao usuário para inserir o número de dados
n = int(input("Insira o número de dados: "))
colUm = input("Digite o nome da primeira coluna: ")
colDois = input("Digite o nome da segunda coluna: ")

# Pedir ao usuário para inserir as cores e nomes
colUm = [input(f"Insira o dado referente {colUm} {i+1}: ") for i in range(n)]
colDois = [
    input(f"Insira o dado referente {colDois} {i+1}: ") for i in range(n)]

# Criar um dataframe com as cores e nomes
dados = {'Cor': colUm, 'Nome': colDois}
df = pd.DataFrame(dados)

# Define o caminho e nome do arquivo com a data do dia
today = datetime.today().strftime('%Y-%m-%d')
filepath_excel = f'C:\\Users\\x\\Documents\\Excel\\cores_e_nomes_{today}.xlsx'
filepath_word = f'C:\\Users\\x\\Documents\\Word\\cores_e_nomes_{today}.docx'
filepath_csv = f'C:\\Users\\x\\Documents\\csv\\cores_e_nomes_{today}.csv'

# Salva os arquivos sem substituir os existentes
if not os.path.exists(filepath_excel):
    df.to_excel(filepath_excel, index=False)
else:
    messagebox.showinfo(
        "Arquivo existente", "O arquivo " + filepath_excel + " já existe.")

if not os.path.exists(filepath_word):
    document = Document()
    table = document.add_table(rows=1, cols=2)
    for i in range(len(df)):
        cells = table.add_row().cells
        cells[0].text = df.iloc[i]['Cor']
        cells[1].text = df.iloc[i]['Nome']
    document.save(filepath_word)
else:
    messagebox.showinfo(
        "Arquivo existente", "O arquivo " + filepath_word + " já existe.")

print("Salvo com sucesso!")
