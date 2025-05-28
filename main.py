import tkinter as tk
from tkinter import filedialog, ttk
import pandas as pd

# Função para seleção do arquivo e exibição dos dados
def selecionar_arquivo():
    file_path = filedialog.askopenfilename(title="Selecione a Planilha", filetypes=[("Excel Files", "*.xlsx;*.xls")])
    
    if file_path:
        # Carregar a quarta aba (índice começa em 0, então a quarta aba é a de índice 3)
        df = pd.read_excel(file_path, sheet_name=3)

        # Limpar tabela antes de carregar novos dados
        for item in tabela.get_children():
            tabela.delete(item)

        # Configurar colunas
        tabela["columns"] = list(df.columns)
        tabela["show"] = "headings"  # Ocultar coluna padrão do Tkinter

        for col in df.columns:
            tabela.column(col, anchor="center", width=150)
            tabela.heading(col, text=col)

        # Adicionar os dados na tabela
        for _, row in df.iterrows():
            tabela.insert("", "end", values=list(row))

        lbl_status.config(text="Planilha carregada com sucesso!")

# Criar janela principal
root = tk.Tk()
root.title("Visualizador de Planilha")

# Expandir a janela
root.geometry("800x500")  # Largura x Altura

# Criar um frame para a tabela
frame = tk.Frame(root)
frame.pack(pady=10, fill="both", expand=True)

# Criar um widget Treeview para a tabela
tabela = ttk.Treeview(frame)
tabela.pack(side="left", fill="both", expand=True)

# Adicionar barra de rolagem
scrollbar = ttk.Scrollbar(frame, orient="vertical", command=tabela.yview)
scrollbar.pack(side="right", fill="y")
tabela.config(yscrollcommand=scrollbar.set)

# Botão para selecionar arquivo
btn = tk.Button(root, text="Selecionar Arquivo", command=selecionar_arquivo)
btn.pack(pady=10)

# Rótulo de status
lbl_status = tk.Label(root, text="", font=("Arial", 12))
lbl_status.pack()

# Iniciar aplicação
root.mainloop()
