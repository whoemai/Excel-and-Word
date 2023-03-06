import os
import tkinter as tk
from datetime import datetime
from tkinter import messagebox

import pandas as pd
from docx import Document

# Pedir ao usuário para inserir o número de dados
n = int(input("Insira o número de dados: "))
colUm = input("Digite o nome da primeira coluna: ")
colDois = input("Digite o nome da segunda coluna: ")

# Pedir ao usuário para inserir os dados
colUmData = [
    input(f"Insira o dado referente {colUm} {i+1}: ") for i in range(n)]
colDoisData = [
    input(f"Insira o dado referente {colDois} {i+1}: ") for i in range(n)]

# Criar um dataframe com os dados
dados = {colUm: colUmData, colDois: colDoisData}
df = pd.DataFrame(dados)

# Define o caminho e nome do arquivo com a data do dia
today = datetime.today().strftime('%Y-%m-%d')
filepath_excel = f'C:\\Users\\x\\Documents\\Export Python\\appy-PY{today}.xlsx'

# Salva os arquivos permitindo substituir os existentes
if os.path.exists(filepath_excel):
    confirm = messagebox.askyesno(
        "Arquivo existente", f"O arquivo {filepath_excel} já existe. Deseja substituí-lo?")
    if not confirm:
        messagebox.showinfo("Arquivo não salvo",
                            f"O arquivo {filepath_excel} não foi salvo.")
    else:
        df.to_excel(filepath_excel, index=False)
else:
    df.to_excel(filepath_excel, index=False)

print("Excel salvo com sucesso!")
