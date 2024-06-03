import tkinter as tk

def on_button_click():
    root = tk.Tk()
    root.title("Salas")
    label.config(text="Redirecionado")

def on_button_click2():
    root = tk.Tk()
    root.title("Horários")
    label.config(text="Redirecionando")

root = tk.Tk()
root.title("Menu")

label = tk.Label(root, text="Alocação de salas:")
label.pack(pady=10)

button1 = tk.Button(root, text="Tabela de Salas", command=on_button_click, bg="green", fg="white")
button1.pack(pady=5)

button2 = tk.Button(root, text="Tabela de Horários", command=on_button_click2, bg="green", fg="white")
button2.pack(pady=10)

root.mainloop()
