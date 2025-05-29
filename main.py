import tkinter as tk
from tkinter import filedialog, ttk, simpledialog
import pandas as pd
import locale

# Definir o padrão de localização para português do Brasil
locale.setlocale(locale.LC_TIME, "pt_BR.UTF-8")

# Variável global para o DataFrame
df = pd.DataFrame()

# Função para selecionar o arquivo e processar os dados
def selecionar_arquivo():
    global df  # Tornar o DataFrame acessível para edições futuras
    file_path = filedialog.askopenfilename(title="Selecione a Planilha", filetypes=[("Excel Files", "*.xlsx;*.xls")])
    
    if file_path:
        try:
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

            # Garantir que a coluna "ID" seja string desde o carregamento
            df["ID"] = df["ID"].astype(str)

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
        except Exception as e:
            lbl_status.config(text=f"Erro ao carregar ou processar a planilha: {e}", fg="red")

# Função para excluir funcionários com base no ID
def excluir_funcionario_por_id():
    global df  # Declare df as global ONLY ONCE at the beginning of the function
    
    ids_para_excluir = simpledialog.askstring("Excluir Funcionário", "Digite os IDs a serem removidos separados por vírgula:")

    if ids_para_excluir:
        ids_lista = [id.strip() for id in ids_para_excluir.split(",")]

        # Não é necessário converter df["ID"] aqui novamente se já for feito em selecionar_arquivo()
        # df["ID"] = df["ID"].astype(str) 

        # Verificar se os IDs existem na tabela antes de excluir
        ids_existentes = df["ID"].tolist()
        ids_a_remover = [id for id in ids_lista if id in ids_existentes]

        if not ids_a_remover:
            lbl_status.config(text="Nenhum ID válido encontrado para remoção.")
            return

        # Remover os funcionários da tabela
        df = df[~df["ID"].isin(ids_a_remover)].reset_index(drop=True)  # Resetando índice após remoção

        # Atualizar a interface após a exclusão
        atualizar_tabela()
        lbl_status.config(text=f"Funcionários removidos: {', '.join(ids_a_remover)}")
    else:
        lbl_status.config(text="Operação de exclusão cancelada.")
        
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
                # Arredondar para duas casas decimais antes de formatar
                diff_horas = horas_normais - total_trabalhado
                horas = int(diff_horas)
                minutos = int((diff_horas - horas) * 60)
                df.at[i,"Horas Devidas"] = f"{horas:02}:{minutos:02}"
                df.at[i,"Horas Extras"] = "00:00"
            else:
                diff_horas = total_trabalhado - horas_normais
                horas = int(diff_horas)
                minutos = int((diff_horas - horas) * 60)
                df.at[i, "Horas Devidas"] = "00:00"
                df.at[i, "Horas Extras"] = f"{horas:02}:{minutos:02}"

        except Exception as e: # Captura qualquer exceção para evitar quebrar o loop
            df.at[i, "Horas Devidas"] = ""
            df.at[i, "Horas Extras"] = ""
            # Opcional: print(f"Erro ao calcular horas para o registro {i}: {e}") # Para depuração

# Função para atualizar a tabela após edição
def atualizar_tabela():
    # Verifica se o DataFrame está vazio para evitar erros
    if df.empty:
        for item in tabela.get_children():
            tabela.delete(item)
        lbl_status.config(text="Nenhum dado para exibir na tabela.")
        return

    calcular_horas() # Atualizar os calculos antes de exibir

    # Calcular automaticamente o valor da hora extra **somente quando Horas Extras > 0**
    for i in range(len(df)):
        # Certifique-se de que "Salário Base" não é uma string vazia e pode ser convertido para float
        if pd.notna(df.at[i, "Salário Base"]) and str(df.at[i, "Salário Base"]).strip() != "" and df.at[i, "Horas Extras"] != "00:00":
            try:
                salario_base = float(df.at[i, "Salário Base"])
                valor_hora = salario_base / 220  # Considerando 220 horas mensais

                # Converter Horas Extras para formato decimal para multiplicação
                horas_extras_parts = df.at[i, "Horas Extras"].split(":")
                horas_extras_dec = int(horas_extras_parts[0]) + (int(horas_extras_parts[1]) / 60)

                df.at[i, "Valor Hora Extra"] = round(valor_hora * 1.5 * horas_extras_dec, 2)  # Aplicando acréscimo de 50% sobre as Horas Extras
            except ValueError:
                df.at[i, "Valor Hora Extra"] = ""
            except IndexError: # Caso horas_extras_parts não tenha 2 elementos
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
        lbl_status.config(text="Nenhuma linha selecionada para edição.", fg="orange")
        return

    # Obter índice da linha selecionada
    item = item_selecionado[0]
    indice = int(item)

    # Colunas que podem ser editadas
    colunas_editaveis = ["Entrada", "Saída-Almoço", "Volta-Almoço", "Saída", "Salário Base"]
    
    mudancas_feitas = False
    for i, coluna in enumerate(df.columns):
        if coluna in colunas_editaveis:
            current_value = df.at[indice, coluna]
            if pd.isna(current_value):
                current_value = "" # Para exibir vazio no dialog se for NaN

            novo_valor = simpledialog.askstring("Editar", f"Digite o novo valor para {coluna} (Valor Atual: {current_value}):")
            
            if novo_valor is not None: # Verifica se o usuário não cancelou
                if novo_valor != str(current_value): # Verifica se houve mudança real
                    df.at[indice, coluna] = novo_valor
                    mudancas_feitas = True
            
    if mudancas_feitas:
        atualizar_tabela()
        lbl_status.config(text="Dados editados com sucesso!")
    else:
        lbl_status.config(text="Nenhuma alteração foi feita ou a edição foi cancelada.")

# Criar interface gráfica
root = tk.Tk()
root.title("Visualizador de Planilha de Ponto")
root.geometry("1200x700") # Aumentado o tamanho da janela para melhor visualização

# Criar um frame para a tabela
frame = tk.Frame(root)
frame.pack(pady=10, fill="both", expand=True)

# Criar tabela interativa
tabela = ttk.Treeview(frame)
tabela.pack(side="left", fill="both", expand=True)

# Adicionar barra de rolagem
scrollbar_y = ttk.Scrollbar(frame, orient="vertical", command=tabela.yview)
scrollbar_y.pack(side="right", fill="y")
tabela.config(yscrollcommand=scrollbar_y.set)

scrollbar_x = ttk.Scrollbar(root, orient="horizontal", command=tabela.xview)
scrollbar_x.pack(side="bottom", fill="x")
tabela.config(xscrollcommand=scrollbar_x.set)


# Botão para selecionar arquivo
btn = tk.Button(root, text="Selecionar Arquivo", command=selecionar_arquivo)
btn.pack(side="left", padx=5, pady=10)

# Botão para editar manualmente os horários e salário base
btn_editar = tk.Button(root, text="Editar Dados", command=editar_celula)
btn_editar.pack(side="left", padx=5, pady=10)

# Botão para excluir funcionários
btn_excluir = tk.Button(root, text="Excluir Funcionário", command=excluir_funcionario_por_id)
btn_excluir.pack(side="left", padx=5, pady=10)

# Rótulo de status
lbl_status = tk.Label(root, text="", font=("Arial", 12))
lbl_status.pack(side="left", padx=10, pady=10)

# Iniciar aplicação
root.mainloop()