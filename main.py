import tkinter as tk
from grade import gradefuncao
from intregalizacaobuttons import realocar
from tkinter import filedialog


def on_button_click(label, text):
    label.config(text=text)
def teste():
    print("foi")
def semselecionado():
    semestre = selected_option.get()
    if(semestre == "1° Semestre"): intsemestre = 1
    if(semestre == "2° Semestre"): intsemestre = 2
    if(semestre == "3° Semestre"): intsemestre = 3
    if(semestre == "4° Semestre"): intsemestre = 4
    if(semestre == "5° Semestre"): intsemestre = 5
    if(semestre == "6° Semestre"): intsemestre = 6
    if(semestre == "7° Semestre"): intsemestre = 7
    if(semestre == "8° Semestre"): intsemestre = 8
    if(semestre == "9° Semestre"): intsemestre = 9
    if(semestre == "10° Semestre"):  intsemestre = 10
    if(semestre == "Indefinido"): intsemestre = 0
    print(intsemestre)



def escolherplanilha():
    # Função para abrir a janela do gerenciador de arquivos
    root = tk.Tk()
    root.withdraw()  # Esconde a janela principal
    file_path = filedialog.askopenfilename()
    text_box.delete('1.0', tk.END)
    text_box.insert(tk.END, file_path)
    return file_path

root = tk.Tk()
root.title("Menu")
#root.state('zoomed')  # Maximiza a janela

# Cria um frame para os botões
button_frame = tk.Frame(root, bg="black")
button_frame.pack(side=tk.LEFT, fill=tk.Y)

# Cria um label centralizado
label = tk.Label(root, text="Alocação de salas:", font=("Arial", 24))
label.pack(pady=10)
options = ["1° Semestre", "2° Semestre","3° Semestre","4° Semestre","5° Semestre","6° Semestre","7° Semestre","8° Semestre","9° Semestre","10° Semestre","Indefinido",]
selected_option = tk.StringVar()
selected_option.set(options[0])  # Definir o valor padrão
option_menu = tk.OptionMenu(root, selected_option, *options, command=semselecionado())
option_menu.pack(pady=10)

# Cria um label com planilha
gradelabel = tk.Label(root, text=gradefuncao(root), font=("Arial", 24))
gradelabel.pack(pady=50)

# Tamanho dos botões
button_width = 20
button_height = 5

# Botão 1
button1 = tk.Button(button_frame, text="Alocar", command=lambda: on_button_click(realocar()), bg="green", fg="white", width=button_width, height=button_height)
button1.pack(fill=tk.X, padx=5, pady=10)

# Botão 2
button2 = tk.Button(button_frame, text="Abrir planilha", command=lambda: on_button_click(label, "Redirecionado para Tabela de Horários"), bg="green", fg="white", width=button_width, height=button_height)
button2.pack(fill=tk.X, padx=5, pady=10)

# Botão 3 (adicional)
button3 = tk.Button(button_frame, text="Selecionar planilha", command=lambda: on_button_click(escolherplanilha()), bg="green", fg="white", width=button_width, height=button_height)
button3.pack(fill=tk.X, padx=5, pady=10)

# Botão 4 (adicional)
button4 = tk.Button(button_frame, text="Opção 4", command=lambda: on_button_click(label, "Redirecionado para Opção 4"), bg="green", fg="white", width=button_width, height=button_height)
button4.pack(fill=tk.X, padx=5, pady=10)



# Botão rteste (adicional)
button4 = tk.Button(button_frame, text="botão teste", command=lambda: on_button_click(semselecionado()), bg="green", fg="white", width=button_width, height=button_height)
button4.pack(fill=tk.X, padx=5, pady=10)

# Cria uma caixa de texto
text_box = tk.Text(button_frame, height=3, width=20, bg="white")
text_box.pack(fill=tk.X, padx=5, pady=5)

# Executa a aplicação
root.mainloop()
