import pandas as pd
import customtkinter as ctk
from tkinter import ttk

class ExcelViewerApp:
    def __init__(self, root, file_path):
        self.root = root
        self.file_path = file_path
        
        self.root.title("Visualizador de Planilha Excel")
        self.root.geometry("1024x720")
        
        self.frame = ctk.CTkFrame(self.root, width=800, height=600)
        self.frame.place(relx=0.5, rely=0.5, anchor="center")
        
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

        # Criação do Treeview
        self.tree = ttk.Treeview(self.frame, show="headings", height=20)
        self.tree.pack(side="left", fill="both", expand=True)
        
        self.filter_var = ctk.StringVar(value="Todos")
        
        self.load_excel()
        self.column_widths = self.get_column_widths()
        self.initialize_table()
        
        self.update_button = ctk.CTkButton(self.root, text="Atualizar", command=self.update_table)
        self.update_button.place(relx=0.5, rely=0.95, anchor="center")
        
        self.all_button = ctk.CTkButton(self.root, text="Todos", command=lambda: self.apply_filter("Todos"))
        self.all_button.place(relx=0.3, rely=0.9, anchor="center")
        
        self.entries_button = ctk.CTkButton(self.root, text="Entradas", command=lambda: self.apply_filter("Entradas"))
        self.entries_button.place(relx=0.4, rely=0.9, anchor="center")
        
        self.exits_button = ctk.CTkButton(self.root, text="Saídas", command=lambda: self.apply_filter("Saidas"))
        self.exits_button.place(relx=0.5, rely=0.9, anchor="center")

    def load_excel(self):
        self.df = pd.read_excel(self.file_path).fillna("")
        self.df = self.df.iloc[::-1]  # Inverte a ordem das linhas
        self.df = self.df.iloc[:, :7]  # Seleciona apenas as colunas A a G (primeiras 7 colunas)
        
        print("Colunas disponíveis no DataFrame:", self.df.columns.tolist())
        
    def get_column_widths(self):
        column_widths = {}
        for i, column in enumerate(self.df.columns):
            if i == 6:  # Verifica se é a coluna 7 (índice 6)
                column_widths[column] = 100  # Define uma largura maior
            else:
                column_widths[column] = 30
        return column_widths
        
    def initialize_table(self):
        self.tree["column"] = list(self.df.columns)
        
        for column in self.tree["column"]:
            self.tree.heading(column, text=column)
            self.tree.column(column, minwidth=30, width=self.column_widths[column], stretch=True)
        
        self.update_table()

    def apply_filter(self, filter_type):
        self.filter_var.set(filter_type)
        self.update_table()

    def update_table(self):
        self.tree.delete(*self.tree.get_children())
        
        filtered_df = self.df
        filter_type = self.filter_var.get()
        
        if "Tipo (Receita/Despesa)" in self.df.columns:
            if filter_type == "Entradas":
                filtered_df = self.df[self.df["Tipo (Receita/Despesa)"] == "Entrada"]
            elif filter_type == "Saidas":
                filtered_df = self.df[self.df["Tipo (Receita/Despesa)"] == "Saida"]
        else:
            print("A coluna 'Tipo (Receita/Despesa)' não foi encontrada no DataFrame.")
        
        for row in filtered_df.to_numpy().tolist():
            self.tree.insert("", "end", values=row)
        
def main():
    file_path = "planilha_anotacoes/gabs_controle_financeiro.xlsx"  # Substitua pelo caminho correto para o seu arquivo
    
    root = ctk.CTk()
    app = ExcelViewerApp(root, file_path)
    root.mainloop()

if __name__ == "__main__":
    main()
