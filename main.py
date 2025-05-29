import tkinter as tk
from tkinter import filedialog, ttk, simpledialog
import pandas as pd
import locale

# Definir o padrão de localização para português do Brasil
locale.setlocale(locale.LC_TIME, "pt_BR.UTF-8")

# Função para selecionar o arquivo e processar os dados
def selecionar_arquivo():
    global df  # Tornar o DataFrame acessível para edições futuras
    file_path = filedialog.askopenfilename(title="Selecione a Planilha", filetypes=[("Excel Files", "*.xlsx;*.xls")])
    
    if file_path:
        # Carregar a terceira aba da planilha (índice 2)
        df = pd.read_excel(file_path, sheet_name=2)

        # Remover as linhas 0, 1, 2 e 3
        df = df.drop([0, 1, 2, 3]).reset_index(drop=True)

        # Definir novos nomes para as colunas
        novos_nomes = [
            "ID", "Nome", "Área", "Data", "Entrada", "Saída-Almoço", "Volta-Almoço",
            "Saída", "Horas Devidas", "Horas Extras", "Horas Normais", "Nota"
        ]
        df.columns = novos_nomes  # Aplicar renomeação das colunas

        # Converter a coluna "Data" para formato datetime corretamente (formato brasileiro)
        df["Data"] = pd.to_datetime(df["Data"], dayfirst=True, errors="coerce")

        # Criar a coluna "Semana" com os dias da semana corretamente formatados
        df["Semana"] = df["Data"].dt.strftime("%A")  # Obtém o nome do dia da semana em português

        # Preenchimento automatico da coluna Horas normais
        df["Horas Normais"] = "08:48"

        # Adicionar colunas "Salário Base" (manual) e "Valor Hora Extra" (automático)
        df["Salário Base"] = ""  # O RH preencherá manualmente
        df["Valor Hora Extra"] = ""  # Inicialmente vazia

        # Reorganizar as colunas para inserir as novas antes da "Nota"
        ordem_colunas = [
            "ID", "Nome", "Área", "Data", "Semana", "Entrada", "Saída-Almoço", "Volta-Almoço",
            "Saída", "Horas Devidas", "Horas Extras", "Horas Normais", "Salário Base",
            "Valor Hora Extra", "Nota"
        ]
        df = df[ordem_colunas]

        # Substituir valores NaN na coluna "Nota" por strings vazias
        df["Nota"] = df["Nota"].fillna("")

        # Substituir todas as ocorrências de "Omissão" em qualquer coluna por strings vazias
        df.replace("Omissão", "", inplace=True)

        # Atualizar a interface com os dados
        atualizar_tabela()

        lbl_status.config(text="Planilha carregada e ajustada com sucesso!")

# Função para fazer o calculo das Horas trabalhadas, Devidas e Horas Extras

def calcular_horas():
    for i in range(len(df)):
        try:
            # Converter horários para datetime
            entrada = pd.to_datetime(df.at[i, "Entrada"], errors="coerce")
            saida_almoco = pd.to_datetime(df.at[i, "Saída-Almoço"], errors="coerce")
            volta_almoço = pd.to_datetime(df.at[i, "Volta-Almoço"], errors="coerce")
            saida_final = pd.to_datetime(df.at[i, "Saída"], errors="coerce")

            # Verificar se todos os horários estão preenchidos
            if pd.isna(entrada) or pd.isna(saida_almoco) or pd.isna(volta_almoço) or pd.isna(saida_final):
                df.at[i, "Horas Devidas"] = ""
                df.at[i, "Horas Extras"] = ""
                continue

            #Calcular o total de horas trabalhadas
            periodo_manha = (saida_almoco - entrada).total_seconds()/3600
            periodo_tarde = (saida_final - volta_almoço).total_seconds()/3600
            total_trabalhado = periodo_manha + periodo_tarde

            # Definir Horas normais (08:48 convertido para formato decimal)
            horas_normais = 8.8

            #Atualizar "Horas Devidas" e "Horas Extras" com base nas horas normais
            if total_trabalhado < horas_normais:
                df.at[i,"Horas Devidas"] = f"{int(horas_normais - total_trabalhado):02}:{int((horas_normais - total_trabalhado)*60) % 60:02}"
                df.at[i,"Horas Extras"] = "00:00"
            else:
                df.at[i, "Horas Devidas"] = "00:00"
                df.at[i, "Horas Extras"] = f"{int(total_trabalhado - horas_normais):02}:{int((total_trabalhado - horas_normais)*60) % 60:02}"

        except ValueError:
            df.at[i, "Horas Devidas"] = ""
            df.at[i, "Horas Extras"] = ""

# Função para atualizar a tabela após edição
def atualizar_tabela():

    calcular_horas() # Atualizar os calculos antes de exibir

    # Calcular automaticamente o valor da hora extra **somente quando Horas Extras > 0**
    for i in range(len(df)):
        if df.at[i, "Salário Base"] != "" and df.at[i, "Horas Extras"] != "00:00":
            try:
                salario_base = float(df.at[i, "Salário Base"])
                valor_hora = salario_base / 220  # Considerando 220 horas mensais

                # Converter Horas Extras para formato decimal para multiplicação
                horas_extras_dec = int(df.at[i, "Horas Extras"].split(":")[0]) + (int(df.at[i, "Horas Extras"].split(":")[1]) / 60)

                df.at[i, "Valor Hora Extra"] = round(valor_hora * 1.5 * horas_extras_dec, 2)  # Aplicando acréscimo de 50% sobre as Horas Extras
            except ValueError:
                df.at[i, "Valor Hora Extra"] = ""
        else:
            df.at[i, "Valor Hora Extra"] = "0.00"

    # Limpar a tabela antes de carregar os novos dados
    for item in tabela.get_children():
        tabela.delete(item)

    # Configurar colunas na tabela
    tabela["columns"] = list(df.columns)
    tabela["show"] = "headings"

    for col in df.columns:
        tabela.column(col, anchor="center", width=150)
        tabela.heading(col, text=col)

    # Adicionar os registros corrigidos à tabela
    for index, row in df.iterrows():
        tabela.insert("", "end", iid=index, values=list(row))

# Função para permitir edição manual dos horários e do salário base
def editar_celula():
    item_selecionado = tabela.selection()
    if not item_selecionado:
        return  # Se nada estiver selecionado, sair da função

    # Obter índice da linha selecionada
    item = item_selecionado[0]
    indice = int(item)

    # Obter valores atuais da linha
    valores = tabela.item(item, "values")

    # Colunas que podem ser editadas
    colunas_editaveis = ["Entrada", "Saída-Almoço", "Volta-Almoço", "Saída", "Salário Base"]
    novos_valores = list(valores)

    for i, coluna in enumerate(df.columns):
        if coluna in colunas_editaveis and (novos_valores[i] == "" or pd.isna(novos_valores[i])):
            novo_valor = simpledialog.askstring("Editar", f"Digite o novo valor para {coluna}:")
            if novo_valor:
                novos_valores[i] = novo_valor
                df.at[indice, coluna] = novo_valor  # Atualizar diretamente no DataFrame

    # Atualizar a tabela com os novos valores e recalcular o valor da hora extra
    atualizar_tabela()

# Criar interface gráfica
root = tk.Tk()
root.title("Visualizador de Planilha Corrigida")
root.geometry("1000x600")

# Criar um frame para a tabela
frame = tk.Frame(root)
frame.pack(pady=10, fill="both", expand=True)

# Criar tabela interativa
tabela = ttk.Treeview(frame)
tabela.pack(side="left", fill="both", expand=True)

# Adicionar barra de rolagem
scrollbar = ttk.Scrollbar(frame, orient="vertical", command=tabela.yview)
scrollbar.pack(side="right", fill="y")
tabela.config(yscrollcommand=scrollbar.set)

# Botão para selecionar arquivo
btn = tk.Button(root, text="Selecionar Arquivo", command=selecionar_arquivo)
btn.pack(pady=10)

# Botão para editar manualmente os horários e salário base
btn_editar = tk.Button(root, text="Editar Dados", command=editar_celula)
btn_editar.pack(pady=10)

# Rótulo de status
lbl_status = tk.Label(root, text="", font=("Arial", 12))
lbl_status.pack()

# Iniciar aplicação
root.mainloop()
