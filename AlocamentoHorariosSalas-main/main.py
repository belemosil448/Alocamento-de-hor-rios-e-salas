import tkinter as tk
from tkinter import filedialog, ttk
from openpyxl import load_workbook
import subprocess
import os
from integralizacaobuttons import alocar_salas, escolher_planilha, abrir_planilha_semestre, salvar_planilha, criar_horarios

def on_button_click(label, text):
    label.config(text=text)

def editar_celula(event):
    selected_item = tree.selection()[0]
    column = tree.identify_column(event.x)
    column_index = int(column[1:]) - 1
    entry = tk.Entry(root)
    entry.place(x=event.x_root - root.winfo_rootx(), y=event.y_root - root.winfo_rooty())
    entry.focus_set()

    def save_edit(event):
        new_value = entry.get()
        tree.set(selected_item, column, new_value)
        row_index = int(tree.index(selected_item)) + 2
        current_sheet.cell(row=row_index, column=column_index + 1).value = new_value
        entry.destroy()

    entry.bind("<Return>", save_edit)

root = tk.Tk()
root.title("Menu")

workbook = None
current_sheet = None
planilha_path = None

button_frame = tk.Frame(root, bg="black")
button_frame.pack(side=tk.LEFT, fill=tk.Y)

label = tk.Label(root, text="Alocação de salas:", font=("Arial", 24))
label.pack(pady=10)

options = ["1° Semestre", "2° Semestre", "3° Semestre", "4° Semestre", "5° Semestre", "6° Semestre", "7° Semestre", "8° Semestre", "9° Semestre", "10° Semestre", "Indefinido"]
selected_option = tk.StringVar()
selected_option.set(options[0])

option_menu = tk.OptionMenu(root, selected_option, *options)
option_menu.pack(pady=10)

tree = ttk.Treeview(root, show="headings")
tree.pack(expand=True, fill=tk.BOTH)
tree.bind("<Double-1>", editar_celula)

button_width = 20
button_height = 5

button1 = tk.Button(button_frame, text="Alocar Salas", command=lambda: alocar_salas(selected_option, tree), bg="green", fg="white", width=button_width, height=button_height)
button1.pack(fill=tk.X, padx=5, pady=10)

button2 = tk.Button(button_frame, text="Abrir planilha", command=lambda: abrir_planilha_semestre(selected_option, tree), bg="green", fg="white", width=button_width, height=button_height)
button2.pack(fill=tk.X, padx=5, pady=10)

button3 = tk.Button(button_frame, text="Selecionar planilha", command=lambda: escolher_planilha(selected_option, tree), bg="green", fg="white", width=button_width, height=button_height)
button3.pack(fill=tk.X, padx=5, pady=10)

button4 = tk.Button(button_frame, text="Salvar Alterações", command=salvar_planilha, bg="green", fg="white", width=button_width, height=button_height)
button4.pack(fill=tk.X, padx=5, pady=10)


root.mainloop()
