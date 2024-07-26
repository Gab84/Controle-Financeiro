import customtkinter as ctk
import tkinter as tk
from tkinter import ttk

# Criação da janela principal
root = ctk.CTk()
root.geometry("600x400")
root.title("Exemplo de Treeview com CustomTkinter")

# Configuração do estilo do ttk para combinar com CustomTkinter
style = ttk.Style()
style.theme_use("default")
style.configure("Treeview",
                background="#D3D3D3",
                foreground="black",
                rowheight=25,
                fieldbackground="#D3D3D3")

style.map('Treeview',
          background=[('selected', '#347083')])

# Frame para o Treeview
tree_frame = ctk.CTkFrame(root)
tree_frame.pack(pady=20)

# Scrollbar para o Treeview
tree_scroll = ctk.CTkScrollbar(tree_frame)
tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)

# Criação do Treeview
tree = ttk.Treeview(tree_frame, yscrollcommand=tree_scroll.set, selectmode="extended")
tree.pack()

# Configuração da scrollbar
tree_scroll.configure(command=tree.yview)

# Definição das colunas
tree['columns'] = ("Nome", "Idade", "Cidade")
tree.column("#0", width=0, stretch=tk.NO)
tree.column("Nome", anchor=tk.W, width=140)
tree.column("Idade", anchor=tk.CENTER, width=100)
tree.column("Cidade", anchor=tk.W, width=140)

# Definição dos cabeçalhos das colunas
tree.heading("#0", text="", anchor=tk.W)
tree.heading("Nome", text="Nome", anchor=tk.W)
tree.heading("Idade", text="Idade", anchor=tk.CENTER)
tree.heading("Cidade", text="Cidade", anchor=tk.W)

# Adição de dados ao Treeview
dados = [
    ("Alice", 30, "Nova York"),
    ("Bob", 25, "São Francisco"),
    ("Charlie", 35, "Chicago")
]

for i, (nome, idade, cidade) in enumerate(dados):
    tree.insert(parent="", index="end", iid=i, text="", values=(nome, idade, cidade))

# Inicialização da janela principal
root.mainloop()
