import tkinter as tk

def gradefuncao(root):
    # Cria um frame para a grade
    grid_frame = tk.Frame(root)
    grid_frame.pack(pady=20)

    # Cria a grade de 5 linhas e 12 colunas
    for i in range(6):
        for j in range(13):
            cell = tk.Label(grid_frame, text=f"", borderwidth=1, relief="solid", width=12, height=3)
            cell.grid(row=i, column=j, padx=5, pady=5)



    first_cell = grid_frame.grid_slaves(row=0, column=0)[0]
    first_cell.config(text="")

    #dias
    first_cell = grid_frame.grid_slaves(row=1, column=0)[0]
    first_cell.config(text="SEGUNDA")
    first_cell = grid_frame.grid_slaves(row=2, column=0)[0]
    first_cell.config(text="TERÇA")
    first_cell = grid_frame.grid_slaves(row=3, column=0)[0]
    first_cell.config(text="QUARTA")
    first_cell = grid_frame.grid_slaves(row=4, column=0)[0]
    first_cell.config(text="QUINTA")
    first_cell = grid_frame.grid_slaves(row=5, column=0)[0]
    first_cell.config(text="SEXTA")
    for v in range(13):
        first_cell = grid_frame.grid_slaves(row=0, column=v)[0]
        first_cell.config(text=v)  
    # Centraliza a grade na janela principal
    grid_frame.pack(expand=True)
