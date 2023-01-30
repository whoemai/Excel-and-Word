import os
import pandas as pd
from docx import Document
import tkinter as tk
from tkinter import messagebox

# Pedir ao usuário para inserir 4 cores e 4 nomes
cores = [input(f"Insira a cor {i+1}: ") for i in range(4)]
nomes = [input(f"Insira o nome {i+1}: ") for i in range(4)]

# Criar um dataframe com as cores e nomes
dados = {'Cor': cores, 'Nome': nomes}
df = pd.DataFrame(dados)

# Define o caminho e nome do arquivo
filepath_excel = 'C:\\Users\\x\\Documents\\Excel\\cores_e_nomes.xlsx'
filepath_word = 'C:\\Users\\x\\Documents\\Word\\cores_e_nomes.docx'
filepath_csv = 'C:\\Users\\x\\Documents\\csv\\cores_e_nomes.csv'

if os.path.exists(filepath_excel):
    root = tk.Tk()
    root.withdraw()
    result = messagebox.askyesno(
        "Permitir salvar", "O arquivo " + filepath_excel + " já existe, deseja substituir?")
    if result:
        df.to_excel(filepath_excel, index=False)
else:
    df.to_excel(filepath_excel, index=False)

if os.path.exists(filepath_word):
    root = tk.Tk()
    root.withdraw()
    result = messagebox.askyesno(
        "Permitir salvar", "O arquivo " + filepath_word + " já existe, deseja substituir?")
    if result:
        document = Document()
        table = document.add_table(rows=1, cols=2)
        for i in range(len(df)):
            cells = table.add_row().cells
            cells[0].text = df.iloc[i]['Cor']
            cells[1].text = df.iloc[i]['Nome']
        document.save(filepath_word)
else:
    document = Document()
    table = document.add_table(rows=1, cols=2)
    for i in range(len(df)):
        cells = table.add_row().cells
        cells[0].text = df.iloc[i]['Cor']
        cells[1].text = df.iloc[i]['Nome']
    document