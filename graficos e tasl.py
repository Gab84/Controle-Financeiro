import pandas as pd
import customtkinter as ctk
from tkinter import ttk
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

class ExcelViewerApp:
    def __init__(self, root, file_path):
        self.root = root
        self.file_path = file_path
        
        self.root.title("Visualizador de Planilha Excel")
        self.root.geometry("1024x720")  # Tamanho da janela
        
        self.frame = ctk.CTkFrame(self.root, width=800, height=600)
        self.frame.place(relx=0.5, rely=0.5, anchor="center")  # Posição e tamanho da tabela
        
        # Configuração do estilo do ttk para combinar com CustomTkinter
        style = ttk.Style()
        style.theme_use("default")
        style.configure("Treeview",
                        background="#D3D3D3",  # Cor de fundo das linhas
                        foreground="black",   # Cor do texto
                        rowheight=25,         # Altura das linhas
                        fieldbackground="#D3D3D3")  # Cor de fundo do campo
        style.map('Treeview',
                  background=[('selected', '#347083')])  # Cor da linha selecionada

        # Criação do Treeview
        self.tree = ttk.Treeview(self.frame, show="headings", height=10)
        self.tree.pack(side="left", fill="both", expand=True)
        
        self.tree.bind("<Button-1>", self.on_treeview_click)
        
        self.filter_var = ctk.StringVar(value="Todos")
        self.sort_order = {"Dia": False, "Mês": False, "Ano": False}
        
        self.load_excel()
        self.column_widths = self.get_column_widths()
        self.initialize_table()
        
        # Botões de filtro e atualização
        self.update_button = ctk.CTkButton(self.root, text="Atualizar", command=self.update_table)
        self.update_button.place(relx=0.5, rely=0.95, anchor="center")  # Posição do botão

        self.all_button = ctk.CTkButton(self.root, text="Todos", command=lambda: self.apply_filter("Todos"))
        self.all_button.place(relx=0.3, rely=0.9, anchor="center")  # Posição do botão

        self.entries_button = ctk.CTkButton(self.root, text="Entradas", command=lambda: self.apply_filter("Entradas"))
        self.entries_button.place(relx=0.4, rely=0.9, anchor="center")  # Posição do botão

        self.exits_button = ctk.CTkButton(self.root, text="Saídas", command=lambda: self.apply_filter("Saidas"))
        self.exits_button.place(relx=0.5, rely=0.9, anchor="center")  # Posição do botão

        # Adicionando o frame para o gráfico de barras
        self.graph_frame = ctk.CTkFrame(self.root, width=800, height=300)
        self.graph_frame.place(relx=0.5, rely=0.25, anchor="center")  # Posição e tamanho do gráfico de barras
        
        # Adicionando o frame para o gráfico de pizza
        self.pie_chart_frame = ctk.CTkFrame(self.root, width=400, height=300)
        self.pie_chart_frame.place(relx=0.75, rely=0.75, anchor="center")  # Posição e tamanho do gráfico de pizza
        
        self.create_bar_chart()
        self.create_pie_chart()
        
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
    
    def apply_sort(self, sort_type):
        self.sort_order[sort_type] = not self.sort_order[sort_type]
        self.sort_by = sort_type
        self.update_table()
    
    def on_treeview_click(self, event):
        region = self.tree.identify("region", event.x, event.y)
        if region == "heading":
            column = self.tree.identify_column(event.x)
            column_name = self.tree.column(column, option="id")
            if column_name in self.sort_order:
                self.apply_sort(column_name)

    def update_table(self):
        self.load_excel()
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
        
        if hasattr(self, 'sort_by'):
            if self.sort_by == "Dia":
                filtered_df = filtered_df.sort_values(by="Dia", ascending=self.sort_order["Dia"])
            elif self.sort_by == "Mês":
                filtered_df = filtered_df.sort_values(by="Mês", ascending=self.sort_order["Mês"])
            elif self.sort_by == "Ano":
                filtered_df = filtered_df.sort_values(by="Ano", ascending=self.sort_order["Ano"])

        for row in filtered_df.to_numpy().tolist():
            self.tree.insert("", "end", values=row)
        
        self.create_bar_chart()
        self.create_pie_chart()
        
    def create_bar_chart(self):
        
        # Adicionando o frame para o gráfico de barras
        self.graph_frame = ctk.CTkFrame(self.root, width=800, height=300)
        self.graph_frame.place(relx=0.5, rely=0.25, anchor="center")  # Posição e tamanho do gráfico de barras
        
      

        entries_sum = self.df[self.df["Tipo (Receita/Despesa)"] == "Entrada"]["Valor"].sum()
        exits_sum = self.df[self.df["Tipo (Receita/Despesa)"] == "Saida"]["Valor"].sum()
        balance = entries_sum - exits_sum
        
        # Garantir que todos os valores sejam não negativos
        entries_sum = max(entries_sum, 0)
        exits_sum = max(exits_sum, 0)
        balance = max(balance, 0)
        
        data = {"Entradas": entries_sum, "Saídas": exits_sum, "Saldo": balance}
        names = list(data.keys())
        values = list(data.values())
        
        fig = Figure(figsize=(4, 3))  # Tamanho do gráfico de barras
        ax = fig.add_subplot(111)
        ax.bar(names, values, color=["#4CAF50", "#F44336", "#FFC107"])  # Cores dos bastões
        ax.set_title("Entradas, Saídas e Saldo")  # Título do gráfico
        
        canvas = FigureCanvasTkAgg(fig, master=self.graph_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(side="top", fill="both", expand=True)
        
    def create_pie_chart(self):

        # Adicionando o frame para o gráfico de pizza
        self.pie_chart_frame = ctk.CTkFrame(self.root, width=400, height=300)
        self.pie_chart_frame.place(relx=0.75, rely=0.75, anchor="center")  # Posição e tamanho do gráfico de pizza  

        # Categorias a serem mostradas no gráfico de pizza
        categories = ["Lazer", "Contas", "Saúde", "Comidas",]
        
        # Verifica se a coluna 'Categoria' e 'Valor' existem no DataFrame
        if "Categoria" not in self.df.columns or "Valor" not in self.df.columns:
            print("As colunas 'Categoria' e/ou 'Valor' não foram encontradas no DataFrame.")
            return
        
        # Calcula a soma dos valores para cada categoria
        data = self.df[self.df["Categoria"].isin(categories)].groupby("Categoria")["Valor"].sum()
        
        # Garantir que todos os valores sejam não negativos
        data = data.clip(lower=0)
        
        labels = data.index.tolist()
        values = data.tolist()
        colors = ["#4CAF50", "#FFC107", "#FF5722", "#03A9F4", "#E91E63"]  # Cores das fatias

        fig = Figure(figsize=(3,2))  # Tamanho do gráfico de pizza
        ax = fig.add_subplot(111)
        ax.pie(values, labels=labels, colors=colors, autopct='%1.1f%%', startangle=90)  # Configuração do gráfico de pizza
        ax.set_title("Distribuição Financeira")  # Título do gráfico
        
        canvas = FigureCanvasTkAgg(fig, master=self.pie_chart_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(side="top", fill="both", expand=True)

def main():
    file_path = "planilha_anotacoes/gabs_controle_financeiro.xlsx"  # Substitua pelo caminho correto para o seu arquivo
    
    root = ctk.CTk()
    app = ExcelViewerApp(root, file_path)
    root.mainloop()

main()
