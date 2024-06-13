import tkinter as tk
from grade import gradefuncao

def on_button_click(label, text):
    label.config(text=text)

root = tk.Tk()
root.title("Menu")
root.state('zoomed')  # Maximiza a janela

# Cria um frame para os botões
button_frame = tk.Frame(root, bg="black")
button_frame.pack(side=tk.LEFT, fill=tk.Y)

# Cria um label centralizado
label = tk.Label(root, text="Alocação de salas:", font=("Arial", 24))
label.pack(pady=10)

# Cria um label com planilha
gradelabel = tk.Label(root, text= gradefuncao(root), font=("Arial", 24))
gradelabel.pack(pady=50)

# Tamanho dos botões
button_width = 20
button_height = 5

# Botão 1
button1 = tk.Button(button_frame, text="Tabela de Salas", command=lambda: on_button_click(label, "Redirecionado para Tabela de Salas"), bg="green", fg="white", width=button_width, height=button_height)
button1.pack(fill=tk.X, padx=5, pady=10)

# Botão 2
button2 = tk.Button(button_frame, text="Tabela de Horários", command=lambda: on_button_click(label, "Redirecionado para Tabela de Horários"), bg="green", fg="white", width=button_width, height=button_height)
button2.pack(fill=tk.X, padx=5, pady=10)

# Botão 3 (adicional)
button3 = tk.Button(button_frame, text="Opção 3", command=lambda: on_button_click(label, "Redirecionado para Opção 3"), bg="green", fg="white", width=button_width, height=button_height)
button3.pack(fill=tk.X, padx=5, pady=10)

# Botão 4 (adicional)
button4 = tk.Button(button_frame, text="Opção 4", command=lambda: on_button_click(label, "Redirecionado para Opção 4"), bg="green", fg="white", width=button_width, height=button_height)
button4.pack(fill=tk.X, padx=5, pady=10)

# Executa a aplicação
root.mainloop()
