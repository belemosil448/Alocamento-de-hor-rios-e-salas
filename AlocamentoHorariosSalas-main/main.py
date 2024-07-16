import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from openpyxl import load_workbook

# Funções existentes do seu código
from integralizacaobuttons import alocar_salas, escolher_planilha, abrir_planilha_semestre, salvar_planilha, current_sheet

workbook = None
current_sheet = None
planilha_path = None

def editar_celula(event):
    item = tree.identify_row(event.y)
    coluna = tree.identify_column(event.x)
    valor_atual = tree.item(item, 'values')[int(coluna.replace('#', '')) - 1]

    entry_editor = tk.Entry(root)
    entry_editor.insert(0, valor_atual)
    entry_editor.pack()
    entry_editor.focus_set()

    def salvar():
        novo_valor = entry_editor.get()
        tree.set(item, coluna, novo_valor)
        entry_editor.pack_forget()
        button_salvar.pack_forget()

    button_salvar = tk.Button(root, text="Salvar", command=salvar)
    button_salvar.pack()

def on_button_click(label, text):
    label.config(text=text)

root = tk.Tk()
root.title("Menu")

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
tree['columns'] = ('col1', 'col2', 'col3')
tree.heading('col1', text='Coluna 1')
tree.heading('col2', text='Coluna 2')
tree.heading('col3', text='Coluna 3')
tree.pack(expand=True, fill=tk.BOTH)

# Exemplo de inserção de dados na Treeview (substitua pelos seus dados)
for i in range(10):
    tree.insert('', tk.END, values=(f'Dado {i}', f'Dado {i+1}', f'Dado {i+2}'))

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
