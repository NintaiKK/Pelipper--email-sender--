from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from tkinter import messagebox
import tkinter as tk
from tkinter import ttk
from tkinter import *
from tkinter.ttk import *
from openpyxl import Workbook
from openpyxl import load_workbook
import openpyxl
import time

print(r'''
 _____        _  _
|  __ \      | |(_)                                 
| |__) |  __ | | _  _ __   _ __    ___  _ __      _______
|  ___/ / _ \| || || '_ \ | '_ \  / _ \| '__|    |==   []|   
| |    |  __/| || || |_) || |_) ||  __/| |       |  ==== |  
|_|     \___||_||_|| .__/ | .__/  \___||_|       '-------'        
                   | |    | |
                   |_|    |_|

        1.0.0 beta            by PakecitusPaki
''')

selected_data = []

def opcoes():
    global selected_data

    wb = load_workbook('pelipper_wb.xlxs')
    ws = wb.active

    janelaOpc = tk.Tk()
    janelaOpc.title('Selecione os Dados')

    frame = tk.Frame(janelaOpc)
    frame.pack(fill = tk.BOTH, expand = True)

    colunas = [cell.value for cell in ws[1]]
    tree = ttk.Treeview(frame, columns =colunas, show = 'headings', selectmode = 'none')
    tree.pack(fill=tk.BOTH, expand = True)

    for col in colunas:
        tree.heading(col, text = col)
        tree.column(col, width = 100, anchor = 'w')

    for row in ws.iter_rows(min_row = 2, values_only = True):
        row = tuple('' if value is None else value for value in row)
        tree.insert("", tk.END, values = row)

    selected_itens = []

    def atualizar_destaque():
        for item in tree.get_children():
            tree.item(item, tags= '')
        for item in selected_itens:
            tree.item(item, tags = 'selected')

    def selecionar_linha(event):
        item = tree.identify_row(event.y)
        if item:
            if item in selected_itens:
                selected_itens.remove(item)
            else:
                selected_itens.append(item)
            atualizar_destaque()

    def selecionar_todas_linhas():
        global selected_data
        selected_itens[:] = [item for item in tree.get_children()]  # Atualiza a lista de itens selecionados
        atualizar_destaque()

    def confirmar_selecao():
        global selected_data
        selected_data = [tree.item(item, 'values') for item in selected_itens]
        janelaOpc.destroy()

    tree.tag_configure('selected', background='lightblue')

    tree.bind("<ButtonRelease-1>", selecionar_linha)

    select_all_btn = tk.Button(janelaOpc, text="Selecionar Todas", command=selecionar_todas_linhas)
    select_all_btn.pack(pady=10)

    confirm_btn = tk.Button(janelaOpc, text="Confirmar Seleção", command=confirmar_selecao)
    confirm_btn.pack(pady=10)

    janelaOpc.mainloop()

def trabaio():
    
    global selected_data

    workbook = load_workbook('cnpj.xlsx')
    sheet = workbook.active
    dados = [{'empresa': row[0], 'numero': row[1], 'validade': row[2]} for row in sheet.iter_rows(min_row=2, values_only=True)]

    for item in selected_data:
        empresa, numero, validade = item

        servico = Service(ChromeDriverManager().install())
        navegador = webdriver.Chrome(service=servico)

        navegador.get("https://web.whatsapp.com/")

        input('Pressione enter depois de entrar')

        try:
            caixaPesquisa = WebDriverWait(navegador, 10).until(
                EC.visibility_of_element_located((By.XPATH, '//*[@id="side"]/div[1]/div/div[2]/div[2]/div'))
            )
            caixaPesquisa.click()
            caixaPesquisa.send_keys(numero)
                
        except Exception as e:
            messagebox.showerror('Erro', 'A página não foi carregada corretamente ou não foi possível fazer o login, tente novamente mais tarde')
                
        try:
            contato = WebDriverWait(navegador, 10).until(
                EC.visibility_of_element_located((By.XPATH, '//*[@id="pane-side"]/div/div/div/div[2]'))
            )
            contato.click()
                
        except Exception as e:
            messagebox.showerror('Erro', 'A página não foi carregada corretamente, tente novamente mais tarde')

        try:
            caixaMensagem = WebDriverWait(navegador, 10).until(
                EC.visibility_of_element_located((By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div[1]/p'))
            )
            caixaMensagem.click()
            caixaMensagem.send_keys(mensagem)
                
        except Exception as e:
            messagebox.showerror('Erro', 'A página não foi carregada corretamente, tente novamente mais tarde')

ababue = input('Digite 1 para ver os dados e 2 para iniciar o processo ')

if ababue == '1':
    opcoes
elif ababue == '2':
    trabaio
else:
    print('Reinicie o programa e tente novamente')

input('pressione enter para fechar')
