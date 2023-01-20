import pandas as pd
from tkinter import *
import tkinter as tk
from tkinter.filedialog import askopenfilename
from tkinter import filedialog
import os


def selecionar_arquivo():
    path = askopenfilename(title="Selecione um arquivo em Excel para abrir")
    return path

def concatenar_tabelas(tabela_1, tabela_2):
    tabela_agrupada = pd.concat([tabela_1, tabela_2])
    return tabela_agrupada

def remover_duplicatas(tabela_agrupada):
    df = tabela_agrupada.value_counts()
    df = tabela_agrupada.drop_duplicates()
    df = tabela_agrupada.drop_duplicates(subset = ['N° DA COLETA', 'COLETOR'], keep='first')
    df = tabela_agrupada.drop_duplicates(subset = ['N° DA COLETA', 'DE: '], keep='first')
    return df

def exportar_tabela(tabela_exportar, path_save):
    path = montar_caminho(path_save)
    tabela_exportar.to_excel(path, index=False, encoding='latin1')
    

def montar_caminho(path):
    caminho = (path.get())
    caminho_com_nome = caminho + f"/Resultadofinal.xlsx"
    return caminho_com_nome
    
def rodando(path_1, path_2, path_save):
    df_1 = pd.read_excel(path_1.get())
    df_2 = pd.read_excel(path_2.get())
    tabela_duplicada = concatenar_tabelas(df_1, df_2)
    tabela_filtrada = remover_duplicatas(tabela_duplicada)
    exportar_tabela(tabela_filtrada, path_save)

def selecionar_pasta_salvamento():
    path_salvar = filedialog.askdirectory()
    return path_salvar
#-------------------------------------------------------FUNÇÕES FRONTEND----------------------------------------------------------------------------------------------------------

#Botão Procurar primeira tabela.
def btn_clicked_2():
    arquivo_1 = selecionar_arquivo()
    path_1.set(arquivo_1)
    if path_1:
        entry0['text'] = f"Arquivo não selecionado: {arquivo_1}"
    
#Botão procurar segunda tabela.
def btn_clicked_3():
    arquivo_2 = selecionar_arquivo()
    path_2.set(arquivo_2)
    if path_1:
        entry2['text'] = f"Arquivo não selecionado: {arquivo_2}"
#Botão procurar pasta de salvamento.
def btn_clicked_4():
    pasta = selecionar_pasta_salvamento()
    path_save.set(pasta)
    if pasta:
        entry1['text'] = f"{pasta}"
    
#Botão Manual de instruções.
def btn_clicked_5():
    os.startfile('venv\data\manual.pdf')
    
#Botão funcionar.
def btn_clicked():
    rodando(path_1, path_2, path_save)
    tk.messagebox.showinfo("Lik.en", "A sua tabela foi enviada, verifique se está tudo certo!")

#-------------------------------------------------------LAYOUT FRONTEND-----------------------------------------------------------------------------------------------------------
window = Tk()
window.title("Pazote - LIK.EN")

#variáveis para o programa
path_1 = tk.StringVar()
path_2 = tk.StringVar()
path_save = tk.StringVar()

window.geometry("1024x768")
window.configure(bg = "#f0fff0")
canvas = Canvas(
    window,
    bg = "#f0fff0",
    height = 768,
    width = 1024,
    bd = 0,
    highlightthickness = 0,
    relief = "ridge")
canvas.place(x = 0, y = 0)

background_img = PhotoImage(file = f"venv\data\dbackbround.png")
background = canvas.create_image(
    419.0, 301.5,
   image=background_img)


#primeiro arquivo select
entry0_img = PhotoImage(file = f"venv\data\img_textBox0.png")
entry0_bg = canvas.create_image(
    682.0, 359.0,
    image = entry0_img)

entry0 = tk.Label(
    text= 'Nenhum arquivo selecionado',
    bd = 0,
    bg = "#d9d9d9",
    highlightthickness = 0)

entry0.place(
    x = 522, y = 349,
    width = 320,
    height = 18)


#Pasta de salvamento
entry1_img = PhotoImage(file = f"venv\data\img_textBox1.png")
entry1_bg = canvas.create_image(
    682.0, 554.0,
    image = entry1_img)

entry1 = tk.Label(
    text= "Selecione uma pasta para salvar",
    bd = 0,
    bg = "#d9d9d9",
    highlightthickness = 0)

entry1.place(
    x = 522, y = 544,
    width = 320,
    height = 18)



#Segundo Arquivo select
entry2_img = PhotoImage(file = f"venv\data\img_textBox2.png")
entry2_bg = canvas.create_image(
    682.0, 391.5,
    image = entry2_img)

entry2 = tk.Label( 
    text= 'Nenhum arquivo selecionado',
    bd = 0,
    bg = "#d9d9d9",
    highlightthickness = 0)

entry2.place(
    x = 522, y = 382,
    width = 320,
    height = 17)



#Botões
img0 = PhotoImage(file = f"venv\data\img0.png")
b0 = Button(
    image = img0,
    borderwidth = 0,
    highlightthickness = 0,
    command = btn_clicked,
    relief = "flat")

b0.place(
    x = 668, y = 612,
    width = 142,
    height = 33)

img1 = PhotoImage(file = f"venv\data\img1.png")
b1 = Button(
    image = img1,
    borderwidth = 0,
    highlightthickness = 0,
    command = btn_clicked_2,
    relief = "flat")

b1.place(
    x = 863, y = 349,
    width = 102,
    height = 20)

img2 = PhotoImage(file = f"venv\data\img2.png")
b2 = Button(
    image = img2,
    borderwidth = 0,
    highlightthickness = 0,
    command = btn_clicked_3,
    relief = "flat")

b2.place(
    x = 863, y = 382,
    width = 102,
    height = 19)

img3 = PhotoImage(file = f"venv\data\img3.png")
b3 = Button(
    image = img3,
    borderwidth = 0,
    highlightthickness = 0,
    command = btn_clicked_4,
    relief = "flat")

b3.place(
    x = 857, y = 544,
    width = 108,
    height = 20)

img4 = PhotoImage(file = f"venv\data\img4.png")
b4 = Button(
    image = img4,
    borderwidth = 0,
    highlightthickness = 0,
    command = btn_clicked_5,
    relief = "flat")

b4.place(
    x = 614, y = 156,
    width = 276,
    height = 55)

window.resizable(False, False)
window.mainloop()