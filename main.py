import tkinter as tk
from tkinter import filedialog, ttk, simpledialog, messagebox
import pandas as pd
import locale
import numpy as np
import re
import unicodedata
import json

# Definir o padrão de localização para português do Brasil
try:
    locale.setlocale(locale.LC_TIME, "pt_BR.UTF-8")
    locale.setlocale(locale.LC_MONETARY, "pt_BR.UTF-8")
except locale.Error:
    print("Locale pt_BR.UTF-8 não encontrado. Usando locale padrão.")


# --- CONSTANTES PARA NOMES DE COLUNAS ---
COL_ID = "ID"
COL_NOME = "Nome"
COL_AREA = "Área"
COL_DATA = "Data"
COL_SEMANA = "Semana"
COL_ENTRADA = "Entrada"
COL_SAIDA_ALMOCO = "Saída-Almoço"
COL_VOLTA_ALMOCO = "Volta-Almoço"
COL_SAIDA = "Saída"
COL_HORAS_DEVIDAS = "Horas Devidas"
COL_HORAS_EXTRAS = "Horas Extras"
COL_HORAS_NORMAIS = "Horas Normais"
COL_SALARIO_BASE = "Salário Base"
COL_VALOR_HORA_EXTRA = "Valor Hora Extra"
COL_NOTA = "Nota"

# Variável global para o DataFrame
df = pd.DataFrame()

# Variáveis de configuração globais
CONFIG_FILE = "config.json"
app_config = {
    "horas_normais_h": 8.8, # Equivalente a 08:48
    "multiplicador_hora_extra": 1.5
}

# Constantes para tratamento de valores nulos/especiais
OMISSAO_VALS = ["Omissão", "nan", ""]
ERRO_FORMATO = "INV_FORMATO"
ERRO_SEQUENCIA = "INV_SEQ"
HORA_ZERO = "00:00"

# Carregar configurações
def load_config():
    global app_config
    try:
        with open(CONFIG_FILE, "r") as f:
            loaded_config = json.load(f)
            for key, value in loaded_config.items():
                if key in app_config:
                    app_config[key] = value
        print("Configurações carregadas com sucesso.")
    except FileNotFoundError:
        print("Arquivo de configuração não encontrado. Usando configurações padrão.")
        save_config()
    except json.JSONDecodeError:
        print("Erro ao decodificar JSON do arquivo de configuração. Usando configurações padrão.")
        save_config()
    except Exception as e:
        print(f"Erro inesperado ao carregar configurações: {e}. Usando configurações padrão.")
        save_config()

# Salvar configurações
def save_config():
    try:
        with open(CONFIG_FILE, "w") as f:
            json.dump(app_config, f, indent=4)
        print("Configurações salvas com sucesso.")
    except Exception as e:
        messagebox.showerror("Erro ao Salvar Configurações", f"Não foi possível salvar as configurações:\n{e}")

# Função para selecionar o arquivo e processar os dados
def selecionar_arquivo():
    global df

    root.config(cursor="watch")
    root.update_idletasks()

    file_path = filedialog.askopenfilename(title="Selecione a Planilha", filetypes=[("Excel Files", "*.xlsx;*.xls")])

    if file_path:
        try:
            df_raw = pd.read_excel(file_path, sheet_name=2)
            df = df_raw.drop([0, 1, 2, 3]).reset_index(drop=True)

            # Apenas as colunas que REALMENTE vêm da planilha original
            novos_nomes_existentes = [
                COL_ID, COL_NOME, COL_AREA, COL_DATA, COL_ENTRADA, COL_SAIDA_ALMOCO,
                COL_VOLTA_ALMOCO, COL_SAIDA, COL_HORAS_DEVIDAS, COL_HORAS_EXTRAS,
                COL_HORAS_NORMAIS, COL_NOTA
            ]
            
            # Garante que o número de colunas da lista novos_nomes_existentes não excede o número de colunas no DataFrame
            if len(novos_nomes_existentes) > len(df.columns):
                # Se a lista de nomes tem mais itens que colunas reais, ajusta
                novos_nomes_existentes = novos_nomes_existentes[:len(df.columns)]
                
            df.columns = novos_nomes_existentes # Renomeia as colunas existentes

            df[COL_ID] = df[COL_ID].astype(str)
            df[COL_DATA] = pd.to_datetime(df[COL_DATA], dayfirst=True, errors="coerce")
            df[COL_SEMANA] = df[COL_DATA].dt.strftime("%A").str.capitalize()
            
            horas_normais_h_int = int(app_config["horas_normais_h"])
            minutos_normais_h = int((app_config["horas_normais_h"] * 60) % 60)
            df[COL_HORAS_NORMAIS] = f"{horas_normais_h_int:02}:{minutos_normais_h:02}"

            # --- Adicionar e inicializar COLUNAS NOVAS AQUI ---
            # Elas devem ser adicionadas após o DataFrame ser carregado e ter as colunas existentes renomeadas.
            if COL_SALARIO_BASE not in df.columns:
                df[COL_SALARIO_BASE] = np.nan # Inicializa com NaN
                df[COL_SALARIO_BASE] = df[COL_SALARIO_BASE].astype(float) # Garante o tipo float
            
            if COL_VALOR_HORA_EXTRA not in df.columns:
                df[COL_VALOR_HORA_EXTRA] = np.nan # Inicializa com NaN
                df[COL_VALOR_HORA_EXTRA] = df[COL_VALOR_HORA_EXTRA].astype(float) # Garante o tipo float
            # --- FIM DA MUDANÇA CRÍTICA ---

            # Definir a ordem final das colunas, incluindo as recém-criadas
            ordem_colunas = [
                COL_ID, COL_NOME, COL_AREA, COL_DATA, COL_SEMANA, COL_ENTRADA,
                COL_SAIDA_ALMOCO, COL_VOLTA_ALMOCO, COL_SAIDA, COL_HORAS_DEVIDAS,
                COL_HORAS_EXTRAS, COL_HORAS_NORMAIS, COL_SALARIO_BASE,
                COL_VALOR_HORA_EXTRA, COL_NOTA
            ]
            df = df[ordem_colunas] # Reorganiza as colunas

            df[COL_NOTA] = df[COL_NOTA].fillna("")
            df.replace("Omissão", "", inplace=True) # Manter para outras colunas de tempo ou nota

            calcular_todas_horas_e_extras() # Calcula as horas para todo o DataFrame carregado
            atualizar_tabela()
            lbl_status.config(text="Planilha carregada e ajustada com sucesso!", fg="green")
        except Exception as e:
            lbl_status.config(text=f"Erro ao carregar ou processar a planilha: {e}", fg="red")
            messagebox.showerror("Erro de Leitura", f"Ocorreu um erro: {e}")
        finally:
            root.config(cursor="")
    else:
        root.config(cursor="")
        lbl_status.config(text="Seleção de arquivo cancelada.", fg="orange")


def excluir_funcionario_por_id():
    global df
    if df.empty:
        messagebox.showinfo("Informação", "Nenhuma planilha carregada.")
        return

    ids_para_excluir = simpledialog.askstring("Excluir Funcionário", "Digite os IDs a serem removidos separados por vírgula:")

    if ids_para_excluir:
        ids_lista = [id_str.strip() for id_str in ids_para_excluir.split(",")]
        ids_existentes_no_df = df[COL_ID].unique()
        ids_a_remover = [id_str for id_str in ids_lista if id_str in ids_existentes_no_df]
        ids_nao_encontrados = [id_str for id_str in ids_lista if id_str not in ids_existentes_no_df]

        if not ids_a_remover:
            lbl_status.config(text="Nenhum ID válido encontrado para remoção.", fg="orange")
            if ids_nao_encontrados:
                messagebox.showwarning("Aviso", f"Os seguintes IDs não foram encontrados na planilha e não foram removidos: {', '.join(ids_nao_encontrados)}")
            return

        confirmar = messagebox.askyesno("Confirmar Exclusão", f"Tem certeza que deseja remover todos os registros dos IDs: {', '.join(ids_a_remover)}?")
        if confirmar:
            df = df[~df[COL_ID].isin(ids_a_remover)].reset_index(drop=True)
            aplicar_filtros() # Atualiza a tabela com os filtros aplicados
            lbl_status.config(text=f"Funcionários removidos: {', '.join(ids_a_remover)}", fg="green")
            if ids_nao_encontrados:
                messagebox.showwarning("Aviso", f"Os seguintes IDs não foram encontrados na planilha e, portanto, não removidos: {', '.join(ids_nao_encontrados)}")
        else:
            lbl_status.config(text="Operação de exclusão cancelada.", fg="orange")
    else:
        lbl_status.config(text="Operação de exclusão cancelada.", fg="orange")

# Helper function to calculate hours for a single row
def _calculate_single_row_hours(row):
    horas_normais_h_config = app_config["horas_normais_h"]
    multiplicador = app_config["multiplicador_hora_extra"]

    entrada_str = str(row[COL_ENTRADA]).strip()
    saida_almoco_str = str(row[COL_SAIDA_ALMOCO]).strip()
    volta_almoco_str = str(row[COL_VOLTA_ALMOCO]).strip()
    saida_final_str = str(row[COL_SAIDA]).strip()

    # Normalize "Omissão", "nan", and empty strings to empty strings for consistent parsing
    entrada_str = "" if entrada_str in OMISSAO_VALS else entrada_str
    saida_almoco_str = "" if saida_almoco_str in OMISSAO_VALS else saida_almoco_str
    volta_almoco_str = "" if volta_almoco_str in OMISSAO_VALS else volta_almoco_str
    saida_final_str = "" if saida_final_str in OMISSAO_VALS else saida_final_str

    horas_devidas_output = ""
    horas_extras_output = ""
    nota_adicional = ""
    valor_hora_extra_output = 0.0

    try:
        entrada_dt = pd.to_datetime(entrada_str, format='%H:%M', errors="raise") if entrada_str else pd.NaT
        saida_almoco_dt = pd.to_datetime(saida_almoco_str, format='%H:%M', errors="raise") if saida_almoco_str else pd.NaT
        volta_almoco_dt = pd.to_datetime(volta_almoco_str, format='%H:%M', errors="raise") if volta_almoco_str else pd.NaT
        saida_final_dt = pd.to_datetime(saida_final_str, format='%H:%M', errors="raise") if saida_final_str else pd.NaT
    except ValueError:
        horas_devidas_output = ERRO_FORMATO
        horas_extras_output = ERRO_FORMATO
        nota_adicional = "Erro: Formato de horário inválido."
        return pd.Series({
            COL_HORAS_DEVIDAS: horas_devidas_output,
            COL_HORAS_EXTRAS: horas_extras_output,
            COL_NOTA: f"{row[COL_NOTA]} {nota_adicional}".strip() if row[COL_NOTA] else nota_adicional,
            COL_VALOR_HORA_EXTRA: valor_hora_extra_output
        })

    total_trabalhado_s = 0

    # Case 1: No lunch break, or lunch times explicitly 00:00 (treated as no break)
    if pd.notna(entrada_dt) and pd.notna(saida_final_dt) and \
       ((pd.isna(saida_almoco_dt) and pd.isna(volta_almoco_dt)) or \
        (saida_almoco_str == HORA_ZERO and volta_almoco_str == HORA_ZERO)):
        
        # If all fields are 00:00, consider no record for the day
        if entrada_str == HORA_ZERO and saida_final_str == HORA_ZERO and \
           saida_almoco_str == HORA_ZERO and volta_almoco_str == HORA_ZERO:
            return pd.Series({
                COL_HORAS_DEVIDAS: "",
                COL_HORAS_EXTRAS: "",
                COL_NOTA: row[COL_NOTA],
                COL_VALOR_HORA_EXTRA: 0.0
            })

        # Adjust saida_final_dt for overnight shifts
        if saida_final_dt < entrada_dt:
            saida_final_dt += pd.Timedelta(days=1)
        
        if entrada_dt >= saida_final_dt:
            horas_devidas_output = ERRO_SEQUENCIA
            horas_extras_output = ERRO_SEQUENCIA
            nota_adicional = "Erro: Entrada >= Saída. Horários sequenciais incorretos."
            return pd.Series({
                COL_HORAS_DEVIDAS: horas_devidas_output,
                COL_HORAS_EXTRAS: horas_extras_output,
                COL_NOTA: f"{row[COL_NOTA]} {nota_adicional}".strip() if row[COL_NOTA] else nota_adicional,
                COL_VALOR_HORA_EXTRA: valor_hora_extra_output
            })
        
        total_trabalhado_s = (saida_final_dt - entrada_dt).total_seconds()
            
    # Case 2: With lunch break
    elif pd.notna(entrada_dt) and pd.notna(saida_almoco_dt) and pd.notna(volta_almoco_dt) and pd.notna(saida_final_dt):
        # Adjust for overnight shifts
        if saida_almoco_dt < entrada_dt: saida_almoco_dt += pd.Timedelta(days=1)
        if volta_almoco_dt < saida_almoco_dt: volta_almoco_dt += pd.Timedelta(days=1)
        if saida_final_dt < volta_almoco_dt: saida_final_dt += pd.Timedelta(days=1)

        if not (entrada_dt < saida_almoco_dt and saida_almoco_dt < volta_almoco_dt and volta_almoco_dt < saida_final_dt):
            horas_devidas_output = ERRO_SEQUENCIA
            horas_extras_output = ERRO_SEQUENCIA
            nota_adicional = "Erro: Sequência de horários incorreta."
            return pd.Series({
                COL_HORAS_DEVIDAS: horas_devidas_output,
                COL_HORAS_EXTRAS: horas_extras_output,
                COL_NOTA: f"{row[COL_NOTA]} {nota_adicional}".strip() if row[COL_NOTA] else nota_adicional,
                COL_VALOR_HORA_EXTRA: valor_hora_extra_output
            })

        periodo_manha_s = (saida_almoco_dt - entrada_dt).total_seconds()
        periodo_tarde_s = (saida_final_dt - volta_almoco_dt).total_seconds()
        total_trabalhado_s = periodo_manha_s + periodo_tarde_s
    else:
        # Horários incompletos ou inválidos para cálculo
        nota_adicional = "Horários incompletos para cálculo."
        return pd.Series({
            COL_HORAS_DEVIDAS: "",
            COL_HORAS_EXTRAS: "",
            COL_NOTA: f"{row[COL_NOTA]} {nota_adicional}".strip() if row[COL_NOTA] else nota_adicional,
            COL_VALOR_HORA_EXTRA: valor_hora_extra_output
        })
            
    total_trabalhado_h = total_trabalhado_s / 3600

    if total_trabalhado_h < horas_normais_h_config:
        diff_total_s = (horas_normais_h_config - total_trabalhado_h) * 3600
        horas = int(diff_total_s // 3600)
        minutos = int((diff_total_s % 3600) // 60)
        horas_devidas_output = f"{horas:02}:{minutos:02}"
        horas_extras_output = HORA_ZERO
    else:
        diff_total_s = (total_trabalhado_h - horas_normais_h_config) * 3600
        horas = int(diff_total_s // 3600)
        minutos = int((diff_total_s % 3600) // 60)
        horas_devidas_output = HORA_ZERO
        horas_extras_output = f"{horas:02}:{minutos:02}"
    
    # Calculate extra hour value for the current row
    salario_base_val = row[COL_SALARIO_BASE]
    if pd.notna(salario_base_val) and salario_base_val > 0 and horas_extras_output and \
       horas_extras_output not in [HORA_ZERO, ERRO_FORMATO, ERRO_SEQUENCIA] and ":" in horas_extras_output:
        valor_hora = salario_base_val / 220
        horas_extras_parts = horas_extras_output.split(":")
        horas_extras_dec = int(horas_extras_parts[0]) + (int(horas_extras_parts[1]) / 60)
        valor_hora_extra_output = round(valor_hora * multiplicador * horas_extras_dec, 2)
    else:
        valor_hora_extra_output = 0.0

    return pd.Series({
        COL_HORAS_DEVIDAS: horas_devidas_output,
        COL_HORAS_EXTRAS: horas_extras_output,
        COL_NOTA: row[COL_NOTA], # Keep original note unless specific error
        COL_VALOR_HORA_EXTRA: valor_hora_extra_output
    })

# Main function to calculate hours for the entire DataFrame
def calcular_todas_horas_e_extras():
    global df
    if df.empty:
        return

    # Apply the single-row calculation function to all rows
    # This will update COL_HORAS_DEVIDAS, COL_HORAS_EXTRAS, COL_VALOR_HORA_EXTRA, and potentially COL_NOTA
    df_calculated_cols = df.apply(_calculate_single_row_hours, axis=1)

    # Update the original DataFrame with the calculated values
    df[COL_HORAS_DEVIDAS] = df_calculated_cols[COL_HORAS_DEVIDAS]
    df[COL_HORAS_EXTRAS] = df_calculated_cols[COL_HORAS_EXTRAS]
    df[COL_VALOR_HORA_EXTRA] = df_calculated_cols[COL_VALOR_HORA_EXTRA]
    # Merge notes: if a new note was generated by calculation, append it
    df[COL_NOTA] = df_calculated_cols.apply(
        lambda row: f"{df.loc[row.name, COL_NOTA]} {row[COL_NOTA]}".strip()
        if (df.loc[row.name, COL_NOTA] and row[COL_NOTA] and df.loc[row.name, COL_NOTA] != row[COL_NOTA])
        else row[COL_NOTA] if row[COL_NOTA] else df.loc[row.name, COL_NOTA],
        axis=1
    )
    # The above COL_NOTA logic can be tricky to get right for all cases.
    # A simpler approach might be to just overwrite the note if the calculation generates one,
    # or ensure _calculate_single_row_hours manages the full note string correctly.
    # For now, if _calculate_single_row_hours returns a new note, it will replace the old one.
    # The previous code was: df_temp.at[i, COL_NOTA] = f"{df_temp.at[i, COL_NOTA]} {nota_adicional}".strip() if df_temp.at[i, COL_NOTA] else nota_adicional
    # This means the single row calculation is already managing the merge, so we just need to assign.
    # Let's adjust _calculate_single_row_hours to directly return the updated note.

    # Re-running the note update to avoid complex merge logic here,
    # assuming _calculate_single_row_hours handles note appending.
    df[COL_NOTA] = df_calculated_cols[COL_NOTA]


def atualizar_tabela(data_frame_exibir=None):
    current_df = data_frame_exibir if data_frame_exibir is not None else df.copy()

    if current_df.empty:
        for item_view in tabela.get_children():
            tabela.delete(item_view)
        lbl_status.config(text="Nenhum dado para exibir na tabela.", fg="orange")
        tabela["columns"] = []
        return

    # No need to recalculate all hours here if apply_filters calls calcular_todas_horas_e_extras already
    # or if the single row edit calls it. The goal is to avoid redundant full recalculations.
    # If data_frame_exibir is None, it means we're showing the full, possibly updated, df.
    # If it's a filtered df, calculations should have already happened on the original df.
    
    # We still need to calculate extra hour value here for the *displayed* DataFrame,
    # because the base df might have just been loaded or filtered, and the
    # COL_VALOR_HORA_EXTRA depends on COL_HORAS_EXTRAS and COL_SALARIO_BASE,
    # and COL_SALARIO_BASE can be edited.
    # However, since _calculate_single_row_hours already calculates COL_VALOR_HORA_EXTRA,
    # and calcular_todas_horas_e_extras uses it, this block below is redundant if
    # `calcular_todas_horas_e_extras` is called on full df and `_calculate_single_row_hours`
    # handles all dependent calculations.

    # This part was for calculating COL_VALOR_HORA_EXTRA if it wasn't done yet,
    # but now it's integrated into _calculate_single_row_hours.
    # Removing this loop to avoid recalculating already calculated values.
    # The assumption is that `df` (the global one) is always up-to-date with calculations.
    # `data_frame_exibir` is just a view of `df`.

    for item_view in tabela.get_children():
        tabela.delete(item_view)

    tabela["columns"] = list(current_df.columns)
    tabela["show"] = "headings"

    for col in current_df.columns:
        tabela.column(col, anchor="center", width=120)
        tabela.heading(col, text=col)

    for index, row in current_df.iterrows():
        formatted_values = []
        for col_name, val in row.items():
            if pd.isna(val) or val == "":
                formatted_values.append("")
            elif isinstance(val, float) and (col_name == COL_SALARIO_BASE or col_name == COL_VALOR_HORA_EXTRA):
                formatted_values.append(f"{val:.2f}".replace('.', ','))
            elif isinstance(val, pd.Timestamp):
                formatted_values.append(val.strftime('%d/%m/%Y'))
            else:
                formatted_values.append(str(val))
        tabela.insert("", "end", iid=index, values=formatted_values)


def editar_celula():
    global df
    if df.empty:
        messagebox.showinfo("Informação", "Nenhuma planilha carregada para editar.")
        return

    item_selecionado = tabela.selection()
    if not item_selecionado:
        lbl_status.config(text="Nenhuma linha selecionada para edição.", fg="orange")
        return

    # The iid of the Treeview corresponds to the actual index of the global DataFrame
    # This is crucial for editing when filters are active
    indice_df = int(item_selecionado[0]) 
    
    # Ensure the index still exists in the global DataFrame
    if indice_df not in df.index:
        messagebox.showwarning("Aviso", "A linha selecionada não existe mais no conjunto de dados original (pode ter sido filtrada ou excluída).")
        aplicar_filtros() # Reload table to show current data
        return

    colunas_treeview = tabela["columns"]

    coluna_idx_str = simpledialog.askstring("Selecionar Coluna",
                                            "Digite o número da coluna para editar (começando em 1):\n" +
                                            "\n".join([f"{i+1}. {col}" for i, col in enumerate(colunas_treeview)]))
    if not coluna_idx_str:
        lbl_status.config(text="Edição cancelada.", fg="orange")
        return

    try:
        coluna_idx = int(coluna_idx_str) - 1
        if not (0 <= coluna_idx < len(colunas_treeview)):
            messagebox.showerror("Erro", "Número da coluna inválido.")
            return
        coluna_para_editar = colunas_treeview[coluna_idx]
    except ValueError:
        messagebox.showerror("Erro", "Entrada inválida para número da coluna.")
        return

    current_value = df.at[indice_df, coluna_para_editar]
    mudancas_feitas = False

    if coluna_para_editar == COL_NOTA:
        display_value_nota = "" if pd.isna(current_value) else str(current_value)
        novo_valor = simpledialog.askstring("Editar Nota", f"Digite a nova anotação para '{COL_NOTA}' (Atual: {display_value_nota}):")
        if novo_valor is not None:
            df.at[indice_df, COL_NOTA] = novo_valor.strip()
            mudancas_feitas = True

    # Lógica para COL_SALARIO_BASE (editável se vazio, com validação)
    elif coluna_para_editar == COL_SALARIO_BASE:
        # Considera 0.0, np.nan ou None como "vazio" para edição
        is_empty_salario = pd.isna(current_value) or (isinstance(current_value, (float, int)) and current_value == 0.0) 
        if is_empty_salario:
            while True:
                novo_valor_str = simpledialog.askstring("Editar Salário Base", f"Digite o novo valor para {COL_SALARIO_BASE} (deixe em branco para limpar):")
                if novo_valor_str is None:
                    break  # Cancelou

                novo_valor_strip = novo_valor_str.strip().replace(",", ".")
                if novo_valor_strip == "":
                    # Preencher todas as células 'Salário Base' vazias para este funcionário com NaN
                    # Ensure COL_ID is present for this logic
                    if COL_ID in df.columns:
                        id_funcionario_selecionado = df.at[indice_df, COL_ID]
                        df.loc[(df[COL_ID] == id_funcionario_selecionado) & (df[COL_SALARIO_BASE].isna() | (df[COL_SALARIO_BASE] == 0.0)), COL_SALARIO_BASE] = np.nan
                    else: # Fallback if COL_ID somehow missing, just update the single cell
                         df.at[indice_df, COL_SALARIO_BASE] = np.nan
                    mudancas_feitas = True
                    break

                try:
                    float_value = float(novo_valor_strip)
                    if float_value < 0:
                        messagebox.showerror("Erro de Entrada", "Salário Base não pode ser negativo.")
                        continue
                    
                    # Preencher todas as células 'Salário Base' vazias para este funcionário
                    if COL_ID in df.columns:
                        id_funcionario_selecionado = df.at[indice_df, COL_ID]
                        # Apply to all rows for this employee that have empty/0 salary
                        df.loc[(df[COL_ID] == id_funcionario_selecionado) & (df[COL_SALARIO_BASE].isna() | (df[COL_SALARIO_BASE] == 0.0)), COL_SALARIO_BASE] = float_value
                    else: # Fallback if COL_ID somehow missing, just update the single cell
                        df.at[indice_df, COL_SALARIO_BASE] = float_value
                    
                    mudancas_feitas = True
                    break
                except ValueError:
                    messagebox.showerror("Erro de Entrada", "Valor inválido para Salário Base. Insira um número (ex: 1500.50).")

        else:
            messagebox.showinfo("Informação", f"A coluna '{COL_SALARIO_BASE}' já está preenchida e não pode ser editada por aqui. Para alterar, edite a planilha original ou use outra funcionalidade se disponível.")

    elif coluna_para_editar in [COL_ENTRADA, COL_SAIDA_ALMOCO, COL_VOLTA_ALMOCO, COL_SAIDA]:
        time_pattern = re.compile(r"^\d{1,2}:\d{2}$")
        display_value_time = str(current_value) if pd.notna(current_value) and str(current_value).strip() not in OMISSAO_VALS else ""

        while True:
            novo_valor_str = simpledialog.askstring(f"Editar {coluna_para_editar}", f"Digite o novo valor para {coluna_para_editar} (Atual: {display_value_time}, formato HH:MM, deixe em branco para limpar):")
            if novo_valor_str is None: break

            novo_valor_strip = novo_valor_str.strip()
            if novo_valor_strip == "":
                df.at[indice_df, coluna_para_editar] = ""
                mudancas_feitas = True
                break
            if time_pattern.match(novo_valor_strip):
                try:
                    h_str, m_str = novo_valor_strip.split(':')
                    h, m = int(h_str), int(m_str)
                    if not (0 <= h <= 23 and 0 <= m <= 59):
                        raise ValueError("Hora ou minuto fora do intervalo válido.")
                    df.at[indice_df, coluna_para_editar] = f"{h:02}:{m:02}"
                    mudancas_feitas = True
                    break
                except ValueError as e:
                    messagebox.showerror("Erro de Entrada", f"Valor de tempo inválido: {e}. Use HH:MM (ex: 08:00, 23:59).")
            else:
                messagebox.showerror("Erro de Formato", f"Formato inválido para {coluna_para_editar}. Use HH:MM (ex: 08:00).")
    else:
        messagebox.showinfo("Informação", f"A coluna '{coluna_para_editar}' não é editável diretamente por aqui. Considere editar a planilha original se necessário.")

    if mudancas_feitas:
        # Recalculate only the affected row after edit
        updated_row_series = _calculate_single_row_hours(df.loc[indice_df])
        for col_to_update in [COL_HORAS_DEVIDAS, COL_HORAS_EXTRAS, COL_VALOR_HORA_EXTRA, COL_NOTA]:
            df.at[indice_df, col_to_update] = updated_row_series[col_to_update]
        
        aplicar_filtros() # Updates the table respecting filters after edit
        lbl_status.config(text="Dados editados com sucesso!", fg="green")
    elif novo_valor is not None:
        lbl_status.config(text="Nenhuma alteração válida foi feita.", fg="orange")
    else:
        lbl_status.config(text="Edição cancelada.", fg="orange")


def remover_sabado_domingo_manual():
    global df
    if df.empty:
        messagebox.showinfo("Informação", "Nenhuma planilha carregada.")
        return

    selecionados_iids = tabela.selection()

    if not selecionados_iids:
        lbl_status.config(text="Selecione as linhas de Sábados/Domingos para remover.", fg="orange")
        return

    confirmar = messagebox.askyesno(
        "Confirmar Remoção",
        f"Você realmente deseja tentar remover {len(selecionados_iids)} registro(s) selecionado(s) que sejam Sábados ou Domingos?"
    )

    if confirmar:
        indices_df_para_remover_final = []
        indices_df_invalidos_dia_semana = []
        dias_para_remover = {"sabado", "domingo"}

        def normalize_text(text):
            return unicodedata.normalize('NFKD', text).encode('ASCII', 'ignore').decode('utf-8')

        for item_iid in selecionados_iids:
            indice_df_original = int(item_iid) 
            
            if indice_df_original not in df.index:
                continue 

            try:
                dia_da_semana_original = str(df.loc[indice_df_original, COL_SEMANA])
                dia_da_semana_normalizado = normalize_text(dia_da_semana_original).lower().strip()
                if dia_da_semana_normalizado in dias_para_remover:
                    indices_df_para_remover_final.append(indice_df_original)
                else:
                    indices_df_invalidos_dia_semana.append(indice_df_original)
            except KeyError:
                messagebox.showerror("Erro", f"Coluna '{COL_SEMANA}' não encontrada no DataFrame. Não é possível verificar o dia.")
                return
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao processar linha {indice_df_original}: {e}")
                continue

        if indices_df_invalidos_dia_semana:
            messagebox.showwarning("Aviso de Remoção",
                                    f"Os seguintes IDs de linha selecionados não são Sábados ou Domingos e não foram removidos: {', '.join(map(str, indices_df_invalidos_dia_semana))}. \n\n"
                                    "Apenas linhas com 'Sábado' ou 'Domingo' na coluna 'Semana' serão removidas.")

        if not indices_df_para_remover_final:
            lbl_status.config(text="Nenhum Sábado ou Domingo válido encontrado entre os selecionados para remoção.", fg="orange")
            return

        df = df.drop(indices_df_para_remover_final).reset_index(drop=True)
        aplicar_filtros() # Atualiza a tabela com os filtros aplicados
        lbl_status.config(text=f"{len(indices_df_para_remover_final)} registro(s) de Sábado/Domingo removido(s) com sucesso!", fg="green")
    else:
        lbl_status.config(text="Remoção de Sábado/Domingo cancelada.", fg="orange")

def calcular_totais_funcionario():
    global df
    if df.empty:
        messagebox.showinfo("Informação", "Nenhuma planilha carregada.")
        return

    root.config(cursor="watch")
    root.update_idletasks()

    try:
        resumo_funcionarios = {}
        # Calculate totals based on the global DataFrame (which is always up-to-date)
        df_for_totals = df.copy() 

        for nome in df_for_totals[COL_NOME].unique():
            df_funcionario = df_for_totals[df_for_totals[COL_NOME] == nome]
            
            total_horas_normais = pd.to_timedelta(df_funcionario[COL_HORAS_NORMAIS].astype(str), errors='coerce').sum()
            total_horas_extras = pd.to_timedelta(
                df_funcionario[COL_HORAS_EXTRAS].astype(str).replace(['', ERRO_FORMATO, ERRO_SEQUENCIA], HORA_ZERO), 
                errors='coerce'
            ).sum()
            total_horas_devidas = pd.to_timedelta(
                df_funcionario[COL_HORAS_DEVIDAS].astype(str).replace(['', ERRO_FORMATO, ERRO_SEQUENCIA], HORA_ZERO), 
                errors='coerce'
            ).sum()
            
            total_valor_hora_extra = df_funcionario[COL_VALOR_HORA_EXTRA].sum()

            def format_timedelta(td):
                total_seconds = int(td.total_seconds())
                sign = "-" if total_seconds < 0 else ""
                total_seconds = abs(total_seconds)
                hours = total_seconds // 3600
                minutes = (total_seconds % 3600) // 60
                return f"{sign}{hours:02}:{minutes:02}"

            resumo_funcionarios[nome] = {
                "Total Horas Normais": format_timedelta(total_horas_normais),
                "Total Horas Extras": format_timedelta(total_horas_extras),
                "Total Horas Devidas": format_timedelta(total_horas_devidas),
                "Total a Receber Horas Extras": f"{total_valor_hora_extra:.2f}"
            }

        exibir_resumo_totais(resumo_funcionarios)
        lbl_status.config(text="Cálculo de totais por funcionário realizado.", fg="green")

    except Exception as e:
        lbl_status.config(text=f"Erro ao calcular totais por funcionário: {e}", fg="red")
        messagebox.showerror("Erro no Cálculo", f"Ocorreu um erro ao calcular os totais: {e}")
    finally:
        root.config(cursor="")


def exibir_resumo_totais(resumo_data):
    total_window = tk.Toplevel(root)
    total_window.title("Resumo de Totais por Funcionário")
    total_window.geometry("700x500")
    total_window.transient(root) # Faz com que a janela de resumo fique sobre a principal
    total_window.grab_set() # Bloqueia interações com a janela principal

    frame_resumo = ttk.Frame(total_window)
    frame_resumo.pack(fill="both", expand=True, padx=10, pady=10)

    tree_resumo = ttk.Treeview(frame_resumo, columns=("Funcionário", "Horas Normais", "Horas Extras", "Horas Devidas", "Valor Hora Extra"), show="headings")
    tree_resumo.pack(side="left", fill="both", expand=True)

    scrollbar_y_resumo = ttk.Scrollbar(frame_resumo, orient="vertical", command=tree_resumo.yview)
    scrollbar_y_resumo.pack(side="right", fill="y")
    tree_resumo.config(yscrollcommand=scrollbar_y_resumo.set)

    tree_resumo.heading("Funcionário", text="Funcionário")
    tree_resumo.heading("Horas Normais", text="H. Normais")
    tree_resumo.heading("Horas Extras", text="H. Extras")
    tree_resumo.heading("Horas Devidas", text="H. Devidas")
    tree_resumo.heading("Valor Hora Extra", text="V. H. Extra (R$)")

    tree_resumo.column("Funcionário", width=150, anchor="w")
    tree_resumo.column("Horas Normais", width=100, anchor="center")
    tree_resumo.column("Horas Extras", width=100, anchor="center")
    tree_resumo.column("Horas Devidas", width=100, anchor="center")
    tree_resumo.column("Valor Hora Extra", width=120, anchor="e")

    for nome, totais in resumo_data.items():
        tree_resumo.insert("", "end", values=(
            nome,
            totais["Total Horas Normais"],
            totais["Total Horas Extras"],
            totais["Total Horas Devidas"],
            totais["Total a Receber Horas Extras"].replace('.', ',')
        ))

    btn_fechar = ttk.Button(total_window, text="Fechar", command=total_window.destroy)
    btn_fechar.pack(pady=10)

    root.wait_window(total_window)

def salvar_planilha():
    global df
    if df.empty:
        messagebox.showinfo("Salvar", "Não há dados para salvar.")
        return

    root.config(cursor="watch")
    root.update_idletasks()

    file_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        title="Salvar Planilha Modificada Como..."
    )
    if file_path:
        try:
            with pd.ExcelWriter(file_path) as writer:
                df_to_save = df.copy()
                df_to_save[COL_SALARIO_BASE] = pd.to_numeric(df_to_save[COL_SALARIO_BASE], errors='coerce')
                df_to_save[COL_VALOR_HORA_EXTRA] = pd.to_numeric(df_to_save[COL_VALOR_HORA_EXTRA], errors='coerce')
                
                # Substituir np.nan e strings de erro por vazio para exportação
                df_to_save.replace({np.nan: '', ERRO_FORMATO: "", ERRO_SEQUENCIA: ""}, inplace=True) 

                for col_monetary in [COL_SALARIO_BASE, COL_VALOR_HORA_EXTRA]:
                    df_to_save[col_monetary] = df_to_save[col_monetary].apply(
                        lambda x: f"{x:.2f}".replace('.', ',') if isinstance(x, (float, int)) else x
                    )

                df_to_save.to_excel(writer, sheet_name="Consolidado", index=False)

                for nome in df[COL_NOME].unique():
                    clean_nome = re.sub(r'[\\/*?:"<>|]', '', nome)[:30]
                    df_funcionario = df[df[COL_NOME] == nome].copy()
                    
                    df_funcionario[COL_SALARIO_BASE] = pd.to_numeric(df_funcionario[COL_SALARIO_BASE], errors='coerce')
                    df_funcionario[COL_VALOR_HORA_EXTRA] = pd.to_numeric(df_funcionario[COL_VALOR_HORA_EXTRA], errors='coerce')

                    df_funcionario.replace({np.nan: '', ERRO_FORMATO: "", ERRO_SEQUENCIA: ""}, inplace=True)
                    for col_monetary in [COL_SALARIO_BASE, COL_VALOR_HORA_EXTRA]:
                        df_funcionario[col_monetary] = df_funcionario[col_monetary].apply(
                            lambda x: f"{x:.2f}".replace('.', ',') if isinstance(x, (float, int)) else x
                        )

                    df_funcionario.to_excel(writer, sheet_name=clean_nome, index=False)

            lbl_status.config(text=f"Planilha salva com sucesso em: {file_path}", fg="green")
        except Exception as e:
            lbl_status.config(text=f"Erro ao salvar planilha: {e}", fg="red")
            messagebox.showerror("Erro ao Salvar", f"Não foi possível salvar a planilha:\n{e}")
        finally:
            root.config(cursor="")
    else:
        root.config(cursor="")
        lbl_status.config(text="Operação de salvar cancelada.", fg="orange")

def abrir_configuracoes():
    config_window = tk.Toplevel(root)
    config_window.title("Configurações do Aplicativo")
    config_window.geometry("400x250")
    config_window.resizable(False, False)
    config_window.transient(root)
    config_window.grab_set()

    lbl_horas_normais = ttk.Label(config_window, text="Horas Normais de Trabalho (decimal, ex: 8.8 para 08:48):")
    lbl_horas_normais.pack(pady=5, padx=10, anchor="w")
    entry_horas_normais = ttk.Entry(config_window)
    entry_horas_normais.insert(0, str(app_config["horas_normais_h"]).replace('.', ','))
    entry_horas_normais.pack(pady=2, padx=10, fill="x")

    lbl_multiplicador = ttk.Label(config_window, text="Multiplicador de Hora Extra (ex: 1.5 para 50% de adicional):")
    lbl_multiplicador.pack(pady=5, padx=10, anchor="w")
    entry_multiplicador = ttk.Entry(config_window)
    entry_multiplicador.insert(0, str(app_config["multiplicador_hora_extra"]).replace('.', ','))
    entry_multiplicador.pack(pady=2, padx=10, fill="x")

    def salvar_e_fechar():
        try:
            novas_horas_normais = float(entry_horas_normais.get().replace(',', '.'))
            novo_multiplicador = float(entry_multiplicador.get().replace(',', '.'))

            if novas_horas_normais <= 0 or novas_horas_normais > 24:
                messagebox.showerror("Erro de Validação", "Horas normais devem ser um valor positivo entre 0 e 24.")
                return
            if novo_multiplicador <= 0:
                messagebox.showerror("Erro de Validação", "Multiplicador de hora extra deve ser um valor positivo.")
                return

            app_config["horas_normais_h"] = novas_horas_normais
            app_config["multiplicador_hora_extra"] = novo_multiplicador
            save_config()
            messagebox.showinfo("Configurações", "Configurações salvas e aplicadas! A tabela será atualizada.")
            config_window.destroy()
            calcular_todas_horas_e_extras() # Recalcula todo o DataFrame com as novas configurações
            aplicar_filtros() # Atualiza a exibição da tabela
        except ValueError:
            messagebox.showerror("Erro de Entrada", "Por favor, insira valores numéricos válidos (use vírgula ou ponto para decimais).")
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro ao salvar as configurações: {e}")

    btn_salvar_config = ttk.Button(config_window, text="Salvar e Fechar", command=salvar_e_fechar)
    btn_salvar_config.pack(pady=15)

    root.wait_window(config_window)


# --- FUNCIONALIDADES DE FILTRO ---
def aplicar_filtros():
    if df.empty:
        atualizar_tabela()
        return

    df_filtrado = df.copy()

    id_filtro_val = entry_filtro_id.get().strip()
    if id_filtro_val:
        df_filtrado = df_filtrado[df_filtrado[COL_ID].str.contains(id_filtro_val, case=False, na=False)]

    nome_filtro_val = entry_filtro_nome.get().strip()
    if nome_filtro_val:
        df_filtrado = df_filtrado[df_filtrado[COL_NOME].apply(lambda x: unicodedata.normalize('NFKD', str(x)).encode('ASCII', 'ignore').decode('utf-8').lower().find(unicodedata.normalize('NFKD', nome_filtro_val).encode('ASCII', 'ignore').decode('utf-8').lower()) != -1)]

    area_filtro_val = entry_filtro_area.get().strip()
    if area_filtro_val:
        df_filtrado = df_filtrado[df_filtrado[COL_AREA].apply(lambda x: unicodedata.normalize('NFKD', str(x)).encode('ASCII', 'ignore').decode('utf-8').lower().find(unicodedata.normalize('NFKD', area_filtro_val).encode('ASCII', 'ignore').decode('utf-8').lower()) != -1)]

    atualizar_tabela(df_filtrado)


def limpar_filtros():
    entry_filtro_id.delete(0, tk.END)
    entry_filtro_nome.delete(0, tk.END)
    entry_filtro_area.delete(0, tk.END)
    aplicar_filtros()


# --- INTERFACE GRÁFICA ---
root = tk.Tk()
root.title("Visualizador e Editor de Planilha de Ponto")
root.geometry("1400x850")

frame_tabela = tk.Frame(root)
frame_tabela.pack(pady=10, padx=10, fill="both", expand=True)

tabela = ttk.Treeview(frame_tabela)

scrollbar_y = ttk.Scrollbar(frame_tabela, orient="vertical", command=tabela.yview)
scrollbar_y.pack(side="right", fill="y")
tabela.config(yscrollcommand=scrollbar_y.set)

scrollbar_x = ttk.Scrollbar(frame_tabela, orient="horizontal", command=tabela.xview)
scrollbar_x.pack(side="bottom", fill="x")
tabela.config(xscrollcommand=scrollbar_x.set)

tabela.pack(side="left", fill="both", expand=True)

frame_botoes = tk.Frame(root)
frame_botoes.pack(pady=5, padx=10, fill="x")

btn_selecionar = tk.Button(frame_botoes, text="Selecionar Arquivo", command=selecionar_arquivo, width=15)
btn_selecionar.pack(side="left", padx=5, pady=5)

btn_editar = tk.Button(frame_botoes, text="Editar Célula", command=editar_celula, width=15)
btn_editar.pack(side="left", padx=5, pady=5)

btn_excluir_id = tk.Button(frame_botoes, text="Excluir por ID", command=excluir_funcionario_por_id, width=15)
btn_excluir_id.pack(side="left", padx=5, pady=5)

btn_remover_fds = tk.Button(frame_botoes, text="Remover Sab/Dom", command=remover_sabado_domingo_manual, width=18)
btn_remover_fds.pack(side="left", padx=5, pady=5)

btn_calcular_totais = tk.Button(frame_botoes, text="Calcular Totais", command=calcular_totais_funcionario, width=15)
btn_calcular_totais.pack(side="left", padx=5, pady=5)

btn_salvar = tk.Button(frame_botoes, text="Salvar como Excel", command=salvar_planilha, width=18)
btn_salvar.pack(side="left", padx=5, pady=5)

btn_config = tk.Button(frame_botoes, text="Configurações", command=abrir_configuracoes, width=15)
btn_config.pack(side="left", padx=5, pady=5)


frame_filtros = tk.LabelFrame(root, text="Filtros")
frame_filtros.pack(pady=10, padx=10, fill="x")

lbl_filtro_id = ttk.Label(frame_filtros, text="ID:")
lbl_filtro_id.pack(side="left", padx=(5,2), pady=5)
entry_filtro_id = ttk.Entry(frame_filtros, width=10)
entry_filtro_id.pack(side="left", padx=2, pady=5)
entry_filtro_id.bind("<KeyRelease>", lambda event: aplicar_filtros())

lbl_filtro_nome = ttk.Label(frame_filtros, text="Nome:")
lbl_filtro_nome.pack(side="left", padx=(15,2), pady=5)
entry_filtro_nome = ttk.Entry(frame_filtros, width=20)
entry_filtro_nome.pack(side="left", padx=2, pady=5)
entry_filtro_nome.bind("<KeyRelease>", lambda event: aplicar_filtros())

lbl_filtro_area = ttk.Label(frame_filtros, text="Área:")
lbl_filtro_area.pack(side="left", padx=(15,2), pady=5)
entry_filtro_area = ttk.Entry(frame_filtros, width=15)
entry_filtro_area.pack(side="left", padx=2, pady=5)
entry_filtro_area.bind("<KeyRelease>", lambda event: aplicar_filtros())

btn_limpar_filtros = ttk.Button(frame_filtros, text="Limpar Filtros", command=limpar_filtros)
btn_limpar_filtros.pack(side="left", padx=(15,5), pady=5)


lbl_status = tk.Label(root, text="Pronto.", font=("Arial", 10), relief=tk.SUNKEN, anchor='w')
lbl_status.pack(side="bottom", fill="x", padx=10, pady=5)

load_config()
aplicar_filtros() # Chamamos aplicar_filtros para inicializar a tabela com os filtros vazios (mostrar tudo)

root.mainloop()