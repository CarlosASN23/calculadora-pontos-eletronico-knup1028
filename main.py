import tkinter as tk
from tkinter import filedialog, ttk, simpledialog, messagebox
import pandas as pd
import locale
import numpy as np # Adicionado para np.nan
import re # Adicionado para validação de tempo
import unicodedata # Adicionado para normalização de texto

# Definir o padrão de localização para português do Brasil
try:
    locale.setlocale(locale.LC_TIME, "pt_BR.UTF-8")
    locale.setlocale(locale.LC_MONETARY, "pt_BR.UTF-8") # Para formatação monetária se necessário
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

            novos_nomes = [
                COL_ID, COL_NOME, COL_AREA, COL_DATA, COL_ENTRADA, COL_SAIDA_ALMOCO,
                COL_VOLTA_ALMOCO, COL_SAIDA, COL_HORAS_DEVIDAS, COL_HORAS_EXTRAS,
                COL_HORAS_NORMAIS, COL_NOTA
            ]
            df.columns = novos_nomes

            df[COL_ID] = df[COL_ID].astype(str)
            df[COL_DATA] = pd.to_datetime(df[COL_DATA], dayfirst=True, errors="coerce")
            df[COL_SEMANA] = df[COL_DATA].dt.strftime("%A").str.capitalize()
            df[COL_HORAS_NORMAIS] = "08:48"

            # Inicializar colunas numéricas com np.nan e tipo float
            df[COL_SALARIO_BASE] = np.nan
            df[COL_SALARIO_BASE] = df[COL_SALARIO_BASE].astype(float)
            df[COL_VALOR_HORA_EXTRA] = np.nan
            df[COL_VALOR_HORA_EXTRA] = df[COL_VALOR_HORA_EXTRA].astype(float)

            ordem_colunas = [
                COL_ID, COL_NOME, COL_AREA, COL_DATA, COL_SEMANA, COL_ENTRADA,
                COL_SAIDA_ALMOCO, COL_VOLTA_ALMOCO, COL_SAIDA, COL_HORAS_DEVIDAS,
                COL_HORAS_EXTRAS, COL_HORAS_NORMAIS, COL_SALARIO_BASE,
                COL_VALOR_HORA_EXTRA, COL_NOTA
            ]
            df = df[ordem_colunas]

            df[COL_NOTA] = df[COL_NOTA].fillna("")
            df.replace("Omissão", "", inplace=True)

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
                messagebox.showwarning("Aviso", f"Os seguintes IDs não foram encontrados na planilha: {', '.join(ids_nao_encontrados)}")
            return

        confirmar = messagebox.askyesno("Confirmar Exclusão", f"Tem certeza que deseja remover todos os registros dos IDs: {', '.join(ids_a_remover)}?")
        if confirmar:
            df = df[~df[COL_ID].isin(ids_a_remover)].reset_index(drop=True)
            atualizar_tabela()
            lbl_status.config(text=f"Funcionários removidos: {', '.join(ids_a_remover)}", fg="green")
            if ids_nao_encontrados:
                messagebox.showwarning("Aviso", f"Os seguintes IDs não foram encontrados e, portanto, não removidos: {', '.join(ids_nao_encontrados)}")
        else:
            lbl_status.config(text="Operação de exclusão cancelada.", fg="orange")
    else:
        lbl_status.config(text="Operação de exclusão cancelada.", fg="orange")

def calcular_horas():
    if df.empty:
        return

    for i in range(len(df)):
        try:
            entrada_str = df.at[i, COL_ENTRADA]
            saida_almoco_str = df.at[i, COL_SAIDA_ALMOCO]
            volta_almoco_str = df.at[i, COL_VOLTA_ALMOCO]
            saida_final_str = df.at[i, COL_SAIDA]

            # Verificar se o horário de almoço foi zerado
            if saida_almoco_str == "00:00" and volta_almoco_str == "00:00":
                entrada = pd.to_datetime(entrada_str, format='%H:%M', errors="coerce")
                saida_final = pd.to_datetime(saida_final_str, format='%H:%M', errors="coerce")

                if pd.isna(entrada) or pd.isna(saida_final):
                    df.at[i, COL_HORAS_DEVIDAS] = ""
                    df.at[i, COL_HORAS_EXTRAS] = ""
                    continue

                total_trabalhado_s = (saida_final - entrada).total_seconds()

                if total_trabalhado_s < 0:
                    total_trabalhado_s += 24 * 3600

                total_trabalhado_h = total_trabalhado_s / 3600
                horas_normais_h = 8 + (48/60) # 8.8 horas

                if total_trabalhado_h < horas_normais_h:
                    diff_total_s = (horas_normais_h - total_trabalhado_h) * 3600
                    horas = int(diff_total_s / 3600)
                    minutos = int((diff_total_s % 3600) / 60)
                    df.at[i, COL_HORAS_DEVIDAS] = f"{horas:02}:{minutos:02}"
                    df.at[i, COL_HORAS_EXTRAS] = "00:00"
                else:
                    diff_total_s = (total_trabalhado_h - horas_normais_h) * 3600
                    horas = int(diff_total_s / 3600)
                    minutos = int((diff_total_s % 3600) / 60)
                    df.at[i, COL_HORAS_DEVIDAS] = "00:00"
                    df.at[i, COL_HORAS_EXTRAS] = f"{horas:02}:{minutos:02}"

            else:
                entrada = pd.to_datetime(entrada_str, format='%H:%M', errors="coerce")
                saida_almoco = pd.to_datetime(saida_almoco_str, format='%H:%M', errors="coerce")
                volta_almoco = pd.to_datetime(volta_almoco_str, format='%H:%M', errors="coerce")
                saida_final = pd.to_datetime(saida_final_str, format='%H:%M', errors="coerce")

                if pd.isna(entrada) or pd.isna(saida_almoco) or pd.isna(volta_almoco) or pd.isna(saida_final):
                    df.at[i, COL_HORAS_DEVIDAS] = ""
                    df.at[i, COL_HORAS_EXTRAS] = ""
                    continue

                periodo_manha_s = (saida_almoco - entrada).total_seconds()
                periodo_tarde_s = (saida_final - volta_almoco).total_seconds()

                # Lidar com virada de dia (horários negativos)
                if periodo_manha_s < 0: periodo_manha_s += 24 * 3600
                if periodo_tarde_s < 0: periodo_tarde_s += 24 * 3600

                total_trabalhado_h = (periodo_manha_s + periodo_tarde_s) / 3600
                horas_normais_h = 8 + (48/60) # 8.8 horas

                if total_trabalhado_h < horas_normais_h:
                    diff_total_s = (horas_normais_h - total_trabalhado_h) * 3600
                    horas = int(diff_total_s / 3600)
                    minutos = int((diff_total_s % 3600) / 60)
                    df.at[i, COL_HORAS_DEVIDAS] = f"{horas:02}:{minutos:02}"
                    df.at[i, COL_HORAS_EXTRAS] = "00:00"
                else:
                    diff_total_s = (total_trabalhado_h - horas_normais_h) * 3600
                    horas = int(diff_total_s / 3600)
                    minutos = int((diff_total_s % 3600) / 60)
                    df.at[i, COL_HORAS_DEVIDAS] = "00:00"
                    df.at[i, COL_HORAS_EXTRAS] = f"{horas:02}:{minutos:02}"

        except Exception: # Captura qualquer exceção
            df.at[i, COL_HORAS_DEVIDAS] = "ERRO"
            df.at[i, COL_HORAS_EXTRAS] = "ERRO"

def atualizar_tabela():
    if df.empty:
        for item_view in tabela.get_children():
            tabela.delete(item_view)
        lbl_status.config(text="Nenhum dado para exibir na tabela.", fg="orange")
        # Limpar colunas se o df estiver vazio
        tabela["columns"] = []
        return

    calcular_horas()

    for i in range(len(df)):
        try:
            salario_base_val = df.at[i, COL_SALARIO_BASE] # Já é float ou NaN
            horas_extras_str = str(df.at[i, COL_HORAS_EXTRAS]) # Garantir que é string

            if pd.notna(salario_base_val) and salario_base_val > 0 and horas_extras_str and horas_extras_str != "00:00" and ":" in horas_extras_str:
                valor_hora = salario_base_val / 220
                horas_extras_parts = horas_extras_str.split(":")
                horas_extras_dec = int(horas_extras_parts[0]) + (int(horas_extras_parts[1]) / 60)
                df.at[i, COL_VALOR_HORA_EXTRA] = round(valor_hora * 1.5 * horas_extras_dec, 2)
            else:
                df.at[i, COL_VALOR_HORA_EXTRA] = 0.0 # Usar float 0.0
        except Exception:
            df.at[i, COL_VALOR_HORA_EXTRA] = 0.0 # Ou np.nan se preferir indicar erro


    for item_view in tabela.get_children():
        tabela.delete(item_view)

    tabela["columns"] = list(df.columns)
    tabela["show"] = "headings"

    for col in df.columns:
        tabela.column(col, anchor="center", width=120) # Ajustar largura conforme necessidade
        tabela.heading(col, text=col)

    for index, row in df.iterrows():
        formatted_values = []
        for col_name, val in row.items():
            if pd.isna(val):
                formatted_values.append("")
            elif isinstance(val, float) and (col_name == COL_SALARIO_BASE or col_name == COL_VALOR_HORA_EXTRA):
                formatted_values.append(f"{val:.2f}") # Formata float para 2 casas decimais
            elif isinstance(val, pd.Timestamp): # Formatar data se necessário
                formatted_values.append(val.strftime('%d/%m/%Y'))
            else:
                formatted_values.append(str(val))
        tabela.insert("", "end", iid=index, values=formatted_values)


# Função para permitir edição manual
def editar_celula():
    if df.empty:
        messagebox.showinfo("Informação", "Nenhuma planilha carregada para editar.")
        return

    item_selecionado = tabela.selection()
    if not item_selecionado:
        lbl_status.config(text="Nenhuma linha selecionada para edição.", fg="orange")
        return

    item_iid = item_selecionado[0]
    indice_df = int(item_iid)
    id_funcionario_selecionado = df.at[indice_df, COL_ID]

    colunas_editaveis_se_vazio = [COL_ENTRADA, COL_SAIDA_ALMOCO, COL_VOLTA_ALMOCO, COL_SAIDA, COL_SALARIO_BASE]
    time_pattern = re.compile(r"^\d{1,2}:\d{2}$") # HH:MM ou H:MM

    mudancas_feitas = False

    # Obter os nomes das colunas da Treeview (que são os mesmos do df)
    colunas_treeview = tabela["columns"]

    # Perguntar qual coluna editar
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

    # Lógica para COL_NOTA (sempre editável)
    if coluna_para_editar == COL_NOTA:
        display_value_nota = "" if pd.isna(current_value) else str(current_value)
        novo_valor = simpledialog.askstring("Editar Nota", f"Digite a nova anotação para '{COL_NOTA}' (Atual: {display_value_nota}):")
        if novo_valor is not None:
            df.at[indice_df, COL_NOTA] = novo_valor.strip()
            mudancas_feitas = True

    # Lógica para COL_SALARIO_BASE (editável se vazio, com validação)
    elif coluna_para_editar == COL_SALARIO_BASE:
        is_empty_salario = pd.isna(current_value) or current_value == 0.0 # Considera 0.0 como "vazio" para edição
        if is_empty_salario:
            while True:
                novo_valor_str = simpledialog.askstring("Editar Salário Base", f"Digite o novo valor para {COL_SALARIO_BASE} (deixe em branco para limpar):")
                if novo_valor_str is None: break # Cancelou
                novo_valor_strip = novo_valor_str.strip().replace(",",".")
                if novo_valor_strip == "":
                    # Preencher todas as células 'Salário Base' vazias para este funcionário com NaN
                    df.loc[(df[COL_ID] == id_funcionario_selecionado) & pd.isna(df[COL_SALARIO_BASE]), COL_SALARIO_BASE] = np.nan
                    df.at[indice_df, COL_SALARIO_BASE] = np.nan # Limpa a célula específica também
                    mudancas_feitas = True
                    break
                try:
                    float_value = float(novo_valor_strip)
                    if float_value < 0:
                        messagebox.showerror("Erro de Entrada", "Salário Base não pode ser negativo.")
                        continue
                    # Preencher todas as células 'Salário Base' vazias para este funcionário
                    df.loc[(df[COL_ID] == id_funcionario_selecionado) & pd.isna(df[COL_SALARIO_BASE]), COL_SALARIO_BASE] = float_value
                    df.at[indice_df, COL_SALARIO_BASE] = float_value # Garante que a célula atual seja atualizada
                    mudancas_feitas = True
                    break
                except ValueError:
                    messagebox.showerror("Erro de Entrada", "Valor inválido para Salário Base. Insira um número (ex: 1500.50).")
        else:
            messagebox.showinfo("Informação", f"A coluna '{COL_SALARIO_BASE}' já está preenchida e não pode ser editada por aqui. Para alterar, edite a planilha original ou use outra funcionalidade se disponível.")

    # Lógica para colunas de tempo (editáveis se vazias, com validação)
    elif coluna_para_editar in [COL_ENTRADA, COL_SAIDA_ALMOCO, COL_VOLTA_ALMOCO, COL_SAIDA]:
        is_empty_time = pd.isna(current_value) or str(current_value).strip() == ""
        if is_empty_time:
            while True:
                novo_valor_str = simpledialog.askstring(f"Editar {coluna_para_editar}", f"Digite o novo valor para {coluna_para_editar} (formato HH:MM):")
                if novo_valor_str is None: break # Cancelou
                novo_valor_strip = novo_valor_str.strip()
                if novo_valor_strip == "":
                    df.at[indice_df, coluna_para_editar] = "" # Limpa o campo
                    mudancas_feitas = True
                    break
                if time_pattern.match(novo_valor_strip):
                    try:
                        h_str, m_str = novo_valor_strip.split(':')
                        h, m = int(h_str), int(m_str)
                        if not (0 <= h <= 23 and 0 <= m <= 59):
                            raise ValueError("Hora ou minuto fora do intervalo válido.")
                        df.at[indice_df, coluna_para_editar] = f"{h:02}:{m:02}" # Armazena formatado
                        mudancas_feitas = True
                        break
                    except ValueError as e:
                        messagebox.showerror("Erro de Entrada", f"Valor de tempo inválido: {e}. Use HH:MM (ex: 08:00, 23:59).")
                else:
                    messagebox.showerror("Erro de Formato", f"Formato inválido para {coluna_para_editar}. Use HH:MM (ex: 08:00).")
        else:
            messagebox.showinfo("Informação", f"A coluna '{coluna_para_editar}' já está preenchida. Para alterar, edite a planilha original.")
    else:
        messagebox.showinfo("Informação", f"A coluna '{coluna_para_editar}' não é editável ou não está vazia (se aplicável).")


    if mudancas_feitas:
        atualizar_tabela()
        lbl_status.config(text="Dados editados com sucesso!", fg="green")
    elif novo_valor is not None : # Se não houve mudança mas o usuário não cancelou a ultima caixa de dialogo
        lbl_status.config(text="Nenhuma alteração válida foi feita.", fg="orange")
    else: # Se cancelou
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
        dias_para_remover = {"sabado", "domingo"} # Usar um conjunto para comparação eficiente

        def normalize_text(text):
            return unicodedata.normalize('NFKD', text).encode('ASCII', 'ignore').decode('utf-8')

        for item_iid in selecionados_iids:
            indice_df = int(item_iid)
            try:
                # A coluna COL_SEMANA deve existir e estar populada
                dia_da_semana_original = str(df.loc[indice_df, COL_SEMANA])
                dia_da_semana_normalizado = normalize_text(dia_da_semana_original).lower().strip()
                if dia_da_semana_normalizado in dias_para_remover:
                    indices_df_para_remover_final.append(indice_df)
                else:
                    indices_df_invalidos_dia_semana.append(indice_df)
            except KeyError:
                messagebox.showerror("Erro", f"Coluna '{COL_SEMANA}' não encontrada no DataFrame. Não é possível verificar o dia.")
                return
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao processar linha {indice_df}: {e}")
                # Decide se quer pular esta linha ou parar tudo
                continue


        if indices_df_invalidos_dia_semana:
            messagebox.showwarning("Aviso de Remoção",
                                    f"{len(indices_df_invalidos_dia_semana)} linha(s) selecionada(s) não são Sábados ou Domingos e não foram removidas.")

        if not indices_df_para_remover_final:
            lbl_status.config(text="Nenhum Sábado ou Domingo válido encontrado entre os selecionados para remoção.", fg="orange")
            return

        df = df.drop(indices_df_para_remover_final).reset_index(drop=True)
        atualizar_tabela()
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
        for nome in df[COL_NOME].unique():
            df_funcionario = df[df[COL_NOME] == nome]
            total_horas_normais = pd.to_timedelta(df_funcionario[COL_HORAS_NORMAIS], errors='coerce').sum()
            total_horas_extras = pd.to_timedelta(df_funcionario[COL_HORAS_EXTRAS], errors='coerce').sum()
            total_horas_devidas = pd.to_timedelta(df_funcionario[COL_HORAS_DEVIDAS].replace('', '00:00'), errors='coerce').sum() # Tratar vazios como 0
            total_valor_hora_extra = df_funcionario[COL_VALOR_HORA_EXTRA].sum()

            def format_timedelta(td):
                total_seconds = int(td.total_seconds())
                hours = total_seconds // 3600
                minutes = (total_seconds % 3600) // 60
                return f"{hours:02}:{minutes:02}"

            resumo_funcionarios[nome] = {
                "Total Horas Normais": format_timedelta(total_horas_normais),
                "Total Horas Extras": format_timedelta(total_horas_extras),
                "Total Horas Devidas": format_timedelta(total_horas_devidas),
                "Total a Receber Horas Extras": f"{total_valor_hora_extra:.2f}"
            }

        # Exibir o resumo (pode ser melhorado com uma janela ou tabela)
        mensagem_resumo = "Resumo por Funcionário:\n\n"
        for nome, totais in resumo_funcionarios.items():
            mensagem_resumo += f"Funcionário: {nome}\n"
            for chave, valor in totais.items():
                mensagem_resumo += f"  {chave}: {valor}\n"
            mensagem_resumo += "\n"

        messagebox.showinfo("Resumo de Totais", mensagem_resumo)
        lbl_status.config(text="Cálculo de totais por funcionário realizado.", fg="green")

    except Exception as e:
        lbl_status.config(text=f"Erro ao calcular totais por funcionário: {e}", fg="red")
        messagebox.showerror("Erro no Cálculo", f"Ocorreu um erro ao calcular os totais: {e}")
    finally:
        root.config(cursor="")

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
                # Salvar a planilha principal
                df.to_excel(writer, sheet_name="Consolidado", index=False)

                # Salvar abas separadas por funcionário
                for nome in df[COL_NOME].unique():
                    df_funcionario = df[df[COL_NOME] == nome]
                    df_funcionario.to_excel(writer, sheet_name=nome, index=False)

            lbl_status.config(text=f"Planilha salva com sucesso em: {file_path}", fg="green")
        except Exception as e:
            lbl_status.config(text=f"Erro ao salvar planilha: {e}", fg="red")
            messagebox.showerror("Erro ao Salvar", f"Não foi possível salvar a planilha:\n{e}")
        finally:
            root.config(cursor="")
    else:
        root.config(cursor="")
        lbl_status.config(text="Operação de salvar cancelada.", fg="orange")

# --- INTERFACE GRÁFICA ---
root = tk.Tk()
root.title("Visualizador e Editor de Planilha de Ponto")
root.geometry("1300x800") # Aumentando um pouco a altura para o novo botão

# Frame para a tabela e scrollbars
frame_tabela = tk.Frame(root)
frame_tabela.pack(pady=10, padx=10, fill="both", expand=True)

tabela = ttk.Treeview(frame_tabela)

scrollbar_y = ttk.Scrollbar(frame_tabela, orient="vertical", command=tabela.yview)
scrollbar_y.pack(side="right", fill="y")
tabela.config(yscrollcommand=scrollbar_y.set)

scrollbar_x = ttk.Scrollbar(frame_tabela, orient="horizontal", command=tabela.xview) # Scrollbar X no frame_tabela
scrollbar_x.pack(side="bottom", fill="x")
tabela.config(xscrollcommand=scrollbar_x.set)

tabela.pack(side="left", fill="both", expand=True) # Tabela depois das scrollbars para elas ficarem ao redor

# Frame para botões
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

# Rótulo de status
lbl_status = tk.Label(root, text="Pronto.", font=("Arial", 10), relief=tk.SUNKEN, anchor='w')
lbl_status.pack(side="bottom", fill="x", padx=10, pady=5)

# Inicializar tabela vazia
atualizar_tabela()

root.mainloop()