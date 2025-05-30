# calculadora_ponto_main.py
# Copyright (c) 2025 Carlos Alberto Souza Nascimento
# Licenciado sob a Licença MIT. Veja o arquivo LICENSE para mais detalhes.

"""
Calculadora de ponto e Horas Extras

Este aplicativo permite aos usuários carregar planilhas de ponto de funcionários,
referente ao aparelho de ponto eletrônico Knup 1028, através da planilha padrão gerada pelo equipamento,
calcular horas trabalhadas, horas devidas, horas extras e o valor correspondente.
Oferece funcionalidades para edição de dados, filtros, salvamento em formato Excel
com resumos individuais por funcionário, e configurações personalizáveis para
cálculo de horas.
"""

import tkinter as tk
from tkinter import filedialog, ttk, simpledialog, messagebox
import pandas as pd
import locale
import numpy as np
import re
import unicodedata
import json
from PIL import Image, ImageTk 
import sys # Adicionado para resource_path
import os  # Adicionado para resource_path

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

# --- CONFIGURAÇÕES DO APLICATIVO ---
def resource_path(relative_path):
    """
    Obtém o caminho absoluto para o recurso, funciona para desenvolvimento e para PyInstaller.

    Args:
        relative_path (str): O caminho relativo para o arquivo de recurso a partir
                             do diretório base do script ou do diretório temporário _MEIPASS
                             do PyInstaller.
    Returns:
        str: O caminho absoluto para o recurso.
    """
    try:
        # PyInstaller cria uma pasta temporária e armazena o caminhao em _MEIPASS
        base_path = sys._MEIPASS # pylint: disable=no-member
    except Exception: # pylint: disable=broad-except # Opcional: para o aviso de exceção muito geral
        base_path = os.path.abspath(".") # Caminho base para desenvolvimento normal
    return os.path.join(base_path, relative_path)

CONFIG_FILE = resource_path("config.json")
app_config = {
    "horas_normais_h": 8.8,
    "multiplicador_hora_extra": 1.5
}

OMISSAO_VALS = ["omissão", "omissao", "nan", ""]
ERRO_FORMATO = "INV_FORMATO"
ERRO_SEQUENCIA = "INV_SEQ"
HORA_ZERO = "00:00"

# --- FUNÇÕES CORE (Lógica do Aplicativo - sem grandes alterações visuais aqui) ---

def load_config():

    """
    Carrega as configurações do aplicativo do arquivo JSON (config.json).

    Se o arquivo não for encontrado ou houver um erro de decodificação,
    usa as configurações padrão e tenta salvar um novo arquivo de configuração.

    Side Effects:
        Modifica a variável global `app_config`.
        Pode chamar `save_config()` se o arquivo de configuração não existir ou for inválido.
        Imprime mensagens no console sobre o status do carregamento.
    """
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

def save_config():
    """
    Salva as configurações atuais da variável global `app_config` em um arquivo JSON.

    Exibe uma caixa de mensagem de erro se o salvamento falhar.

    Side Effects:
        Cria ou sobrescreve o arquivo `config.json`.
        Imprime mensagens no console sobre o status do salvamento.
    """
    try:
        with open(CONFIG_FILE, "w") as f:
            json.dump(app_config, f, indent=4)
        print("Configurações salvas com sucesso.")
    except Exception as e:
        messagebox.showerror("Erro ao Salvar Configurações", f"Não foi possível salvar as configurações:\n{e}")

def update_button_states():
    """
    Atualiza o estado (habilitado/desabilitado) dos botões da interface
    com base no estado atual da aplicação (DataFrame carregado, seleção na tabela).

    Side Effects:
        Modifica o atributo 'state' de vários botões da UI (btn_salvar,
        btn_excluir_id, btn_calcular_totais, btn_editar, btn_remover_fds).
    """
    if df.empty:
        btn_salvar.config(state="disabled")
        btn_excluir_id.config(state="disabled")
        btn_calcular_totais.config(state="disabled")
        # Os botões de edição dependem da seleção na tabela, tratados em on_treeview_select
    else:
        btn_salvar.config(state="normal")
        btn_excluir_id.config(state="normal")
        btn_calcular_totais.config(state="normal")

    # Estado dos botões de edição/seleção
    if tabela.selection():
        btn_editar.config(state="normal")
        btn_remover_fds.config(state="normal")
    else:
        btn_editar.config(state="disabled")
        btn_remover_fds.config(state="disabled")


def selecionar_arquivo():
    """
    Abre um diálogo para o usuário selecionar uma planilha Excel.

    Após a seleção, lê os dados da planilha, processa as colunas,
    calcula as horas e atualiza a tabela na interface.
    Atualiza a barra de status com o resultado da operação.

    Side Effects:
        Modifica a variável global `df` com os dados da planilha.
        Chama `calcular_todas_horas_e_extras()` e `aplicar_filtros()`.
        Atualiza `lbl_status` e o estado dos botões através de `update_button_states()`.
    """
    global df
    root.config(cursor="watch")
    root.update_idletasks()
    file_path = filedialog.askopenfilename(title="Selecione a Planilha", filetypes=[("Excel Files", "*.xlsx;*.xls")])
    root.config(cursor="")

    if file_path:
        try:
            df_raw = pd.read_excel(file_path, sheet_name=2)
            df = df_raw.iloc[4:].reset_index(drop=True)

            novos_nomes_existentes = [
                COL_ID, COL_NOME, COL_AREA, COL_DATA, COL_ENTRADA, COL_SAIDA_ALMOCO,
                COL_VOLTA_ALMOCO, COL_SAIDA, COL_HORAS_DEVIDAS, COL_HORAS_EXTRAS,
                COL_HORAS_NORMAIS, COL_NOTA
            ]
            colunas_para_renomear = min(len(novos_nomes_existentes), len(df.columns))
            df = df.iloc[:, :colunas_para_renomear]
            df.columns = novos_nomes_existentes[:colunas_para_renomear]

            df[COL_ID] = df[COL_ID].astype(str)
            df[COL_DATA] = pd.to_datetime(df[COL_DATA], dayfirst=True, errors="coerce")
            df[COL_SEMANA] = df[COL_DATA].dt.strftime("%A").str.capitalize()
            
            horas_normais_h_int = int(app_config["horas_normais_h"])
            minutos_normais_h = int((app_config["horas_normais_h"] * 60) % 60)
            df[COL_HORAS_NORMAIS] = f"{horas_normais_h_int:02}:{minutos_normais_h:02}"

            ordem_colunas = [
                COL_ID, COL_NOME, COL_AREA, COL_DATA, COL_SEMANA, COL_ENTRADA,
                COL_SAIDA_ALMOCO, COL_VOLTA_ALMOCO, COL_SAIDA, COL_HORAS_DEVIDAS,
                COL_HORAS_EXTRAS, COL_HORAS_NORMAIS, COL_SALARIO_BASE,
                COL_VALOR_HORA_EXTRA, COL_NOTA
            ]
            for col in ordem_colunas:
                if col not in df.columns:
                    if col in [COL_SALARIO_BASE, COL_VALOR_HORA_EXTRA]:
                         df[col] = np.nan
                         df[col] = df[col].astype(float)
                    else:
                        df[col] = ""
            df = df[ordem_colunas]
            df[COL_NOTA] = df[COL_NOTA].fillna("")
            df.replace("Omissão", "", inplace=True, regex=True) # regex=True para case-insensitive "Omissão"

            calcular_todas_horas_e_extras()
            aplicar_filtros()
            lbl_status.config(text=f"✅ Sucesso: Planilha '{file_path.split('/')[-1]}' carregada!", foreground="green")
        except Exception as e:
            df = pd.DataFrame() # Limpa o DataFrame em caso de erro
            aplicar_filtros() # Atualiza a tabela para mostrar que está vazia
            lbl_status.config(text=f"❌ Erro ao carregar planilha: {e}", foreground="red")
            messagebox.showerror("Erro de Leitura", f"Ocorreu um erro: {e}")
        finally:
            update_button_states()
    else:
        lbl_status.config(text="ℹ️ Seleção de arquivo cancelada.", foreground="darkorange")
        update_button_states()


def _calculate_single_row_hours(row):
    """
    Calcula horas devidas, extras, valor de hora extra e notas para uma única linha de dados.

    A função processa os horários de entrada, saída e almoço para determinar o tempo
    trabalhado. Compara este tempo com as horas normais configuradas para calcular
    diferenças (devidas ou extras). Também calcula o valor monetário das horas extras
    com base no salário base e multiplicador configurados. Adiciona notas sobre
    erros de formato ou sequência de horários.

    Args:
        row (pd.Series): Uma linha do DataFrame contendo, no mínimo, as colunas:
                         COL_ENTRADA, COL_SAIDA_ALMOCO, COL_VOLTA_ALMOCO, COL_SAIDA (como strings "HH:MM" ou vazias),
                         COL_SALARIO_BASE (como float ou NaN),
                         COL_NOTA (como string).

    Returns:
        pd.Series: Uma Series contendo os resultados calculados para as colunas:
                   COL_HORAS_DEVIDAS (str "HH:MM" ou código de erro),
                   COL_HORAS_EXTRAS (str "HH:MM" ou código de erro),
                   COL_NOTA (str, potencialmente atualizada com mensagens de erro),
                   COL_VALOR_HORA_EXTRA (float).
    """

    horas_normais_h_config = app_config["horas_normais_h"]
    multiplicador = app_config["multiplicador_hora_extra"]

    entrada_str = str(row[COL_ENTRADA]).strip()
    saida_almoco_str = str(row[COL_SAIDA_ALMOCO]).strip()
    volta_almoco_str = str(row[COL_VOLTA_ALMOCO]).strip()
    saida_final_str = str(row[COL_SAIDA]).strip()

    # Normalização mais robusta para omissão e nan
    entrada_str = "" if entrada_str.lower() in OMISSAO_VALS or entrada_str.lower() == 'nan' else entrada_str
    saida_almoco_str = "" if saida_almoco_str.lower() in OMISSAO_VALS or saida_almoco_str.lower() == 'nan' else saida_almoco_str
    volta_almoco_str = "" if volta_almoco_str.lower() in OMISSAO_VALS or volta_almoco_str.lower() == 'nan' else volta_almoco_str
    saida_final_str = "" if saida_final_str.lower() in OMISSAO_VALS or saida_final_str.lower() == 'nan' else saida_final_str
    
    nota_final = str(row[COL_NOTA]) if pd.notna(row[COL_NOTA]) else ""
    horas_devidas_output = ""
    horas_extras_output = ""
    valor_hora_extra_output = 0.0

    try:
        entrada_dt = pd.to_datetime(entrada_str, format='%H:%M', errors="raise") if entrada_str else pd.NaT
        saida_almoco_dt = pd.to_datetime(saida_almoco_str, format='%H:%M', errors="raise") if saida_almoco_str else pd.NaT
        volta_almoco_dt = pd.to_datetime(volta_almoco_str, format='%H:%M', errors="raise") if volta_almoco_str else pd.NaT
        saida_final_dt = pd.to_datetime(saida_final_str, format='%H:%M', errors="raise") if saida_final_str else pd.NaT
    except ValueError:
        horas_devidas_output = ERRO_FORMATO
        horas_extras_output = ERRO_FORMATO
        nota_final = f"{nota_final} (Erro: Formato de horário inválido)".strip()
        return pd.Series({
            COL_HORAS_DEVIDAS: horas_devidas_output, COL_HORAS_EXTRAS: horas_extras_output,
            COL_NOTA: nota_final, COL_VALOR_HORA_EXTRA: valor_hora_extra_output
        })

    total_trabalhado_s = 0
    # CASO 1: Sem almoço OU almoço zerado (00:00)
    if pd.notna(entrada_dt) and pd.notna(saida_final_dt) and \
       ((pd.isna(saida_almoco_dt) and pd.isna(volta_almoco_dt)) or \
        (saida_almoco_str == HORA_ZERO and volta_almoco_str == HORA_ZERO)):
        
        if entrada_str == HORA_ZERO and saida_final_str == HORA_ZERO and \
           (saida_almoco_str == HORA_ZERO or saida_almoco_str == "") and \
           (volta_almoco_str == HORA_ZERO or volta_almoco_str == ""):
            return pd.Series({
                COL_HORAS_DEVIDAS: "", COL_HORAS_EXTRAS: "",
                COL_NOTA: nota_final, COL_VALOR_HORA_EXTRA: 0.0
            })

        if saida_final_dt < entrada_dt: saida_final_dt += pd.Timedelta(days=1)
        
        if entrada_dt >= saida_final_dt:
            horas_devidas_output = ERRO_SEQUENCIA
            horas_extras_output = ERRO_SEQUENCIA
            nota_final = f"{nota_final} (Erro Seq: E>=S s/almoço)".strip()
        else:
            total_trabalhado_s = (saida_final_dt - entrada_dt).total_seconds()
            
    # CASO 2: Com almoço
    elif pd.notna(entrada_dt) and pd.notna(saida_almoco_dt) and pd.notna(volta_almoco_dt) and pd.notna(saida_final_dt):
        if saida_almoco_dt < entrada_dt: saida_almoco_dt += pd.Timedelta(days=1)
        if volta_almoco_dt < saida_almoco_dt: volta_almoco_dt += pd.Timedelta(days=1) # Volta pode ser no dia seguinte
        if saida_final_dt < volta_almoco_dt: saida_final_dt += pd.Timedelta(days=1) # Saída pode ser no dia seguinte

        # Permitir almoço de duração zero (SaidaAlmoco == VoltaAlmoco)
        if not (entrada_dt <= saida_almoco_dt and saida_almoco_dt <= volta_almoco_dt and volta_almoco_dt <= saida_final_dt and entrada_dt < saida_final_dt):
            horas_devidas_output = ERRO_SEQUENCIA
            horas_extras_output = ERRO_SEQUENCIA
            nota_final = f"{nota_final} (Erro Seq: c/almoço)".strip()
        else:
            periodo_manha_s = (saida_almoco_dt - entrada_dt).total_seconds()
            periodo_tarde_s = (saida_final_dt - volta_almoco_dt).total_seconds()
            total_trabalhado_s = periodo_manha_s + periodo_tarde_s
    # CASO 3: Horários incompletos para cálculo
    else:
        if any(s for s in [entrada_str, saida_almoco_str, volta_almoco_str, saida_final_str]): # Se algum campo foi preenchido
             nota_final = f"{nota_final} (Horários incompletos)".strip()
        # Se todos os campos de horário estiverem vazios, considera-se ausência, sem nota adicional aqui.
        return pd.Series({
            COL_HORAS_DEVIDAS: "", COL_HORAS_EXTRAS: "",
            COL_NOTA: nota_final, COL_VALOR_HORA_EXTRA: 0.0
        })

    # Se já houve erro de sequência, retorna
    if horas_devidas_output == ERRO_SEQUENCIA:
         return pd.Series({
            COL_HORAS_DEVIDAS: horas_devidas_output, COL_HORAS_EXTRAS: horas_extras_output,
            COL_NOTA: nota_final, COL_VALOR_HORA_EXTRA: 0.0
        })

    # Cálculo de horas devidas/extras
    if total_trabalhado_s > 0: # Só calcula se houve tempo trabalhado válido
        total_trabalhado_h = total_trabalhado_s / 3600.0
        diff_total_s = total_trabalhado_s - (horas_normais_h_config * 3600.0)

        if diff_total_s < -1: # Deu horas a menos (considera uma pequena margem para arredondamento)
            segundos_devidos = abs(diff_total_s)
            horas_dev = int(segundos_devidos // 3600)
            minutos_dev = int((segundos_devidos % 3600) // 60)
            horas_devidas_output = f"{horas_dev:02}:{minutos_dev:02}"
            horas_extras_output = HORA_ZERO
        else: # Cumpriu ou fez horas extras
            segundos_extras = diff_total_s if diff_total_s > 0 else 0
            horas_ext = int(segundos_extras // 3600)
            minutos_ext = int((segundos_extras % 3600) // 60)
            horas_extras_output = f"{horas_ext:02}:{minutos_ext:02}"
            horas_devidas_output = HORA_ZERO
    # Se total_trabalhado_s == 0 e não houve erro de formatação ou sequência, não faz nada (ausência)
    elif total_trabalhado_s == 0 and not horas_devidas_output and not horas_extras_output:
        pass # Mantém horas devidas/extras como ""

    # Cálculo do valor da hora extra
    salario_base_val = row[COL_SALARIO_BASE] # Já deve ser float ou NaN
    if pd.notna(salario_base_val) and salario_base_val > 0 and \
       horas_extras_output and horas_extras_output != HORA_ZERO and \
       horas_extras_output not in [ERRO_FORMATO, ERRO_SEQUENCIA] and ":" in horas_extras_output:
        try:
            valor_hora = salario_base_val / 220.0 # Carga horária mensal padrão CLT
            h_extra, m_extra = map(int, horas_extras_output.split(':'))
            horas_extras_dec = h_extra + (m_extra / 60.0)
            valor_hora_extra_output = round(valor_hora * multiplicador * horas_extras_dec, 2)
        except ValueError:
            valor_hora_extra_output = 0.0 
            nota_final = f"{nota_final} (Erro calc. Vlr HE)".strip()
    
    return pd.Series({
        COL_HORAS_DEVIDAS: horas_devidas_output,
        COL_HORAS_EXTRAS: horas_extras_output,
        COL_NOTA: nota_final,
        COL_VALOR_HORA_EXTRA: valor_hora_extra_output
    })

def calcular_todas_horas_e_extras():
    """
    Aplica o cálculo de horas (_calculate_single_row_hours) para todas as linhas
    do DataFrame global `df`.

    Garante que as colunas necessárias para o cálculo existam e tenham tipos
    adequados antes de aplicar a função de cálculo por linha.

    Side Effects:
        Modifica as colunas COL_HORAS_DEVIDAS, COL_HORAS_EXTRAS, COL_NOTA,
        e COL_VALOR_HORA_EXTRA no DataFrame global `df`.
    """
    global df
    if df.empty: return

    cols_horarios = [COL_ENTRADA, COL_SAIDA_ALMOCO, COL_VOLTA_ALMOCO, COL_SAIDA]
    for col in cols_horarios:
        if col not in df.columns: df[col] = ""
        df[col] = df[col].astype(str).fillna("")

    if COL_NOTA not in df.columns: df[COL_NOTA] = ""
    df[COL_NOTA] = df[COL_NOTA].astype(str).fillna("")

    if COL_SALARIO_BASE not in df.columns: df[COL_SALARIO_BASE] = np.nan
    df[COL_SALARIO_BASE] = pd.to_numeric(df[COL_SALARIO_BASE], errors='coerce')
    
    # Garantir que COL_VALOR_HORA_EXTRA exista antes de ser preenchida
    if COL_VALOR_HORA_EXTRA not in df.columns: df[COL_VALOR_HORA_EXTRA] = np.nan
    df[COL_VALOR_HORA_EXTRA] = pd.to_numeric(df[COL_VALOR_HORA_EXTRA], errors='coerce')


    calculated_data = df.apply(_calculate_single_row_hours, axis=1)
    df[[COL_HORAS_DEVIDAS, COL_HORAS_EXTRAS, COL_NOTA, COL_VALOR_HORA_EXTRA]] = calculated_data


def atualizar_tabela(data_frame_exibir=None):
    """
    Atualiza o widget Treeview (tabela) da interface com os dados fornecidos.

    Se `data_frame_exibir` for None, usa o DataFrame global `df`.
    Formata valores monetários e datas para exibição.

    Args:
        data_frame_exibir (pd.DataFrame, optional): O DataFrame a ser exibido.
                                                   Padrão é None (usa o `df` global).
    Side Effects:
        Limpa e repopula o widget `tabela` da UI.
        Atualiza o estado dos botões através de `update_button_states()`.
    """
    current_df = data_frame_exibir if data_frame_exibir is not None else df.copy()
    for item_view in tabela.get_children():
        tabela.delete(item_view)

    if current_df.empty:
        tabela["columns"] = []
        # lbl_status já é atualizado por aplicar_filtros se df_filtrado for vazio
        return

    tabela["columns"] = list(current_df.columns)
    tabela["show"] = "headings"

    col_widths = {
        COL_ID: 60, COL_NOME: 220, COL_AREA: 120, COL_DATA: 90, COL_SEMANA: 100,
        COL_ENTRADA: 70, COL_SAIDA_ALMOCO: 70, COL_VOLTA_ALMOCO: 70, COL_SAIDA: 70,
        COL_HORAS_DEVIDAS: 70, COL_HORAS_EXTRAS: 70, COL_HORAS_NORMAIS: 70,
        COL_SALARIO_BASE: 100, COL_VALOR_HORA_EXTRA: 110, COL_NOTA: 250
    }
    col_anchors = {
        COL_SALARIO_BASE: "e", COL_VALOR_HORA_EXTRA: "e",
        COL_ID: "center", COL_DATA: "center", COL_ENTRADA: "center", COL_SAIDA_ALMOCO: "center",
        COL_VOLTA_ALMOCO: "center", COL_SAIDA: "center", COL_HORAS_DEVIDAS: "center",
        COL_HORAS_EXTRAS: "center", COL_HORAS_NORMAIS: "center"
    }

    for col in current_df.columns:
        width = col_widths.get(col, 100)
        anchor = col_anchors.get(col, "w") # Default anchor "w" (west/esquerda)
        tabela.column(col, anchor=anchor, width=width, minwidth=40)
        tabela.heading(col, text=col)

    for index, row in current_df.iterrows():
        formatted_values = []
        for col_name, val in row.items():
            if pd.isna(val) or str(val).strip() == "":
                formatted_values.append("")
            elif isinstance(val, float) and (col_name == COL_SALARIO_BASE or col_name == COL_VALOR_HORA_EXTRA):
                try:
                    formatted_values.append(locale.format_string("%.2f", val, grouping=True))
                except (TypeError, ValueError):
                    formatted_values.append(str(val))
            elif isinstance(val, pd.Timestamp):
                formatted_values.append(val.strftime('%d/%m/%Y'))
            else:
                formatted_values.append(str(val))
        tabela.insert("", "end", iid=index, values=formatted_values)
    update_button_states() # Atualiza botões após popular a tabela


def editar_celula():
    """
    Permite ao usuário editar o conteúdo de uma célula selecionada na tabela.

    Abre diálogos para selecionar a coluna e inserir o novo valor.
    Valida a entrada de acordo com o tipo da coluna (hora, salário, nota, etc.).
    Se a edição for em uma coluna que afeta cálculos (horários, salário),
    recalcula a linha modificada e atualiza a tabela.

    Side Effects:
        Modifica o DataFrame global `df` na linha e coluna editada.
        Pode chamar `_calculate_single_row_hours` e `aplicar_filtros()`.
        Atualiza `lbl_status`.
    """
    global df
    if df.empty or not tabela.selection():
        messagebox.showwarning("Aviso", "Nenhuma planilha carregada ou nenhuma linha selecionada para edição.")
        return

    item_selecionado = tabela.selection()[0]
    indice_df_original = int(item_selecionado)

    if indice_df_original not in df.index:
        messagebox.showerror("Erro Crítico", "Índice da linha selecionada não encontrado no DataFrame. Sincronia perdida.")
        return

    colunas_treeview = list(df.columns)
    col_options_str = "\n".join([f"{i+1}. {col}" for i, col in enumerate(colunas_treeview)])
    coluna_idx_str = simpledialog.askstring("Selecionar Coluna para Editar",
                                            f"Linha (Índice DF: {indice_df_original})\nDigite o nº da coluna para editar (1 a {len(colunas_treeview)}):\n\n{col_options_str}")

    if not coluna_idx_str:
        lbl_status.config(text="ℹ️ Edição cancelada.", foreground="blue")
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

    valor_atual_df = df.loc[indice_df_original, coluna_para_editar]
    display_valor_atual = str(valor_atual_df) if pd.notna(valor_atual_df) else ""
    
    novo_valor_str = simpledialog.askstring(f"Editar: {coluna_para_editar}", 
                                       f"Linha (Índice DF: {indice_df_original}), Coluna: '{coluna_para_editar}'\nValor Atual: {display_valor_atual}\n\nNovo valor:")

    if novo_valor_str is None:
        lbl_status.config(text="ℹ️ Edição cancelada.", foreground="blue")
        return

    novo_valor_strip = novo_valor_str.strip()
    mudancas_feitas = False

    # --- Lógica de edição por coluna ---
    if coluna_para_editar in [COL_ENTRADA, COL_SAIDA_ALMOCO, COL_VOLTA_ALMOCO, COL_SAIDA]:
        if novo_valor_strip == "":
            df.loc[indice_df_original, coluna_para_editar] = ""
            mudancas_feitas = True
        elif re.fullmatch(r"\d{1,2}:\d{2}", novo_valor_strip):
            h, m = map(int, novo_valor_strip.split(':'))
            if 0 <= h <= 23 and 0 <= m <= 59:
                df.loc[indice_df_original, coluna_para_editar] = f"{h:02}:{m:02}"
                mudancas_feitas = True
            else: messagebox.showerror("Erro", "Hora/minuto inválido.")
        else: messagebox.showerror("Erro", f"Formato para {coluna_para_editar} deve ser HH:MM ou vazio.")

    elif coluna_para_editar == COL_SALARIO_BASE:
        if novo_valor_strip == "":
            df.loc[indice_df_original, COL_SALARIO_BASE] = np.nan
            mudancas_feitas = True
        else:
            try:
                val_float = float(novo_valor_strip.replace(",", "."))
                if val_float < 0: messagebox.showerror("Erro", "Salário não pode ser negativo.")
                else:
                    id_func = df.loc[indice_df_original, COL_ID]
                    df.loc[df[COL_ID] == id_func, COL_SALARIO_BASE] = val_float
                    mudancas_feitas = True
            except ValueError: messagebox.showerror("Erro", "Salário inválido.")

    elif coluna_para_editar == COL_NOTA:
        df.loc[indice_df_original, COL_NOTA] = novo_valor_strip
        mudancas_feitas = True
    
    # Adicione outras colunas editáveis aqui (ID, Nome, Área, Data)
    elif coluna_para_editar == COL_ID:
        if novo_valor_strip: 
            df.loc[indice_df_original, COL_ID] = novo_valor_strip
            mudancas_feitas = True
        else: messagebox.showerror("Erro", "ID não pode ser vazio.")

    elif coluna_para_editar == COL_NOME:
        if novo_valor_strip:
            df.loc[indice_df_original, COL_NOME] = novo_valor_strip
            mudancas_feitas = True
        else: messagebox.showerror("Erro", "Nome não pode ser vazio.")
    
    elif coluna_para_editar == COL_AREA:
        df.loc[indice_df_original, COL_AREA] = novo_valor_strip
        mudancas_feitas = True

    elif coluna_para_editar == COL_DATA:
        if novo_valor_strip == "":
            df.loc[indice_df_original, COL_DATA] = pd.NaT
            mudancas_feitas = True
        else:
            try:
                nova_data = pd.to_datetime(novo_valor_strip, dayfirst=True, errors='raise')
                df.loc[indice_df_original, COL_DATA] = nova_data
                df.loc[indice_df_original, COL_SEMANA] = nova_data.strftime("%A").str.capitalize()
                mudancas_feitas = True
            except ValueError: messagebox.showerror("Erro", "Formato de data inválido. Use DD/MM/AAAA.")
    else:
        messagebox.showinfo("Informação", f"Coluna '{coluna_para_editar}' não é diretamente editável ou não possui lógica de edição definida.")

    if mudancas_feitas:
        if coluna_para_editar in [COL_ENTRADA, COL_SAIDA_ALMOCO, COL_VOLTA_ALMOCO, COL_SAIDA, COL_SALARIO_BASE, COL_DATA]:
            # Recalcular a linha modificada
             updated_row_series = _calculate_single_row_hours(df.loc[indice_df_original])
             df.loc[indice_df_original, [COL_HORAS_DEVIDAS, COL_HORAS_EXTRAS, COL_NOTA, COL_VALOR_HORA_EXTRA]] = updated_row_series
        
        aplicar_filtros()
        lbl_status.config(text=f"✅ Linha {indice_df_original}, Coluna '{coluna_para_editar}' atualizada.", foreground="green")
    elif novo_valor_str is not None: # Se não cancelou, mas também não houve mudança válida
        lbl_status.config(text="ℹ️ Nenhuma alteração aplicada.", foreground="blue")


def excluir_funcionario_por_id():
    """
    Remove todos os registros de funcionários com os IDs fornecidos pelo usuário.

    Pede ao usuário uma lista de IDs separados por vírgula.
    Confirma a exclusão antes de remover os dados do DataFrame global `df`.

    Side Effects:
        Modifica o DataFrame global `df`.
        Chama `aplicar_filtros()` para atualizar a UI.
        Atualiza `lbl_status` e `update_button_states()`.
    """
    global df
    if df.empty:
        messagebox.showwarning("Aviso", "Nenhuma planilha carregada.")
        return

    ids_para_excluir_str = simpledialog.askstring("Excluir Funcionário por ID", "Digite os IDs a serem removidos, separados por vírgula:")
    if not ids_para_excluir_str:
        lbl_status.config(text="ℹ️ Exclusão cancelada.", foreground="blue")
        return

    ids_lista = [s.strip() for s in ids_para_excluir_str.split(",") if s.strip()]
    if not ids_lista:
        lbl_status.config(text="ℹ️ Nenhum ID válido fornecido.", foreground="orange")
        return

    ids_encontrados_df = df[COL_ID].unique()
    ids_a_remover = [id_val for id_val in ids_lista if id_val in ids_encontrados_df]
    ids_nao_encontrados = [id_val for id_val in ids_lista if id_val not in ids_encontrados_df]

    msg_nao_encontrados = f"IDs não encontrados: {', '.join(ids_nao_encontrados)}." if ids_nao_encontrados else ""

    if not ids_a_remover:
        lbl_status.config(text=f"ℹ️ Nenhum dos IDs fornecidos foi encontrado. {msg_nao_encontrados}", foreground="orange")
        messagebox.showinfo("Exclusão", f"Nenhum dos IDs fornecidos foi encontrado na planilha.\n{msg_nao_encontrados}")
        return

    confirmar = messagebox.askyesno("Confirmar Exclusão", 
                                     f"Remover todos os registros dos IDs: {', '.join(ids_a_remover)}?\n{msg_nao_encontrados}")
    if confirmar:
        df = df[~df[COL_ID].isin(ids_a_remover)].reset_index(drop=True)
        aplicar_filtros()
        status_msg = f"✅ IDs removidos: {', '.join(ids_a_remover)}. {msg_nao_encontrados}"
        lbl_status.config(text=status_msg.strip(), fg="green")
        if df.empty: # Se todos os dados foram removidos
            update_button_states() # Desabilitar botões
    else:
        lbl_status.config(text="ℹ️ Exclusão cancelada.", foreground="blue")


def remover_sabado_domingo_manual():
    """
    Remove as linhas selecionadas na tabela que correspondem a Sábados ou Domingos.

    Verifica a coluna 'Semana' das linhas selecionadas.
    Pede confirmação ao usuário antes de remover as linhas do DataFrame global `df`.

    Side Effects:
        Modifica o DataFrame global `df`.
        Chama `aplicar_filtros()` para atualizar a UI.
        Atualiza `lbl_status` e `update_button_states()`.
    """
    global df
    if df.empty or not tabela.selection():
        messagebox.showwarning("Aviso", "Nenhuma planilha carregada ou nenhuma linha selecionada.")
        return

    selecionados_iids_treeview = tabela.selection()
    confirmar = messagebox.askyesno("Confirmar Remoção",
        f"Remover os {len(selecionados_iids_treeview)} registro(s) selecionado(s) que sejam Sábados ou Domingos?")

    if confirmar:
        indices_df_para_remover = []
        indices_nao_removidos = []
        dias_para_remover_norm = {"sabado", "domingo"}

        def normalize_text_simple(text):
            if pd.isna(text): return ""
            return unicodedata.normalize('NFKD', str(text)).encode('ASCII', 'ignore').decode('utf-8').lower().strip()

        for item_iid_str in selecionados_iids_treeview:
            indice_df = int(item_iid_str)
            if indice_df in df.index:
                dia_semana = normalize_text_simple(df.loc[indice_df, COL_SEMANA])
                if dia_semana in dias_para_remover_norm:
                    indices_df_para_remover.append(indice_df)
                else:
                    indices_nao_removidos.append(str(indice_df))
            else: # Segurança, caso o índice não exista mais no df
                print(f"Aviso: Índice {indice_df} da seleção não encontrado no DataFrame durante remoção S/D.")


        if indices_nao_removidos:
            messagebox.showwarning("Aviso Parcial", f"Algumas linhas selecionadas não eram Sábados/Domingos e não foram removidas (Índices: {', '.join(indices_nao_removidos)}).")

        if not indices_df_para_remover:
            lbl_status.config(text="ℹ️ Nenhum Sábado/Domingo válido selecionado para remoção.", foreground="orange")
            if not indices_nao_removidos: messagebox.showinfo("Informação", "Nenhuma linha válida (Sábado/Domingo) foi selecionada.")
            return
        
        df.drop(indices_df_para_remover, inplace=True)
        df.reset_index(drop=True, inplace=True)
        aplicar_filtros()
        lbl_status.config(text=f"✅ {len(indices_df_para_remover)} registro(s) de Sábado/Domingo removido(s).", fg="green")
        if df.empty: update_button_states()
    else:
        lbl_status.config(text="ℹ️ Remoção de Sábado/Domingo cancelada.", foreground="blue")


def calcular_totais_funcionario():
    """
    Calcula os totais de horas normais, extras, devidas e valor de HE por funcionário.

    Os resultados são então exibidos em uma nova janela através da função
    `exibir_resumo_totais`.

    Side Effects:
        Mostra uma janela de resumo (`exibir_resumo_totais`).
        Atualiza `lbl_status`.
        Pode modificar o DataFrame `df` para garantir tipos corretos antes da soma.
    """
    global df
    if df.empty:
        messagebox.showwarning("Aviso", "Nenhuma planilha carregada para calcular totais.")
        return

    root.config(cursor="watch"); root.update_idletasks()
    try:
        for col_hora in [COL_HORAS_NORMAIS, COL_HORAS_EXTRAS, COL_HORAS_DEVIDAS]:
            if col_hora not in df.columns: df[col_hora] = HORA_ZERO
            df[col_hora] = df[col_hora].astype(str).fillna(HORA_ZERO)
        if COL_VALOR_HORA_EXTRA not in df.columns: df[COL_VALOR_HORA_EXTRA] = 0.0
        df[COL_VALOR_HORA_EXTRA] = pd.to_numeric(df[COL_VALOR_HORA_EXTRA], errors='coerce').fillna(0.0)

        resumo_funcionarios = {}
        nomes_unicos = [n for n in df[COL_NOME].unique() if pd.notna(n) and str(n).strip() != ""]

        for nome in nomes_unicos:
            df_f = df[df[COL_NOME] == nome]
            total_hn_td = pd.to_timedelta(df_f[COL_HORAS_NORMAIS].replace(['',ERRO_FORMATO,ERRO_SEQUENCIA,'nan','NaT'], HORA_ZERO), errors='coerce').sum()
            total_he_td = pd.to_timedelta(df_f[COL_HORAS_EXTRAS].replace(['',ERRO_FORMATO,ERRO_SEQUENCIA,'nan','NaT'], HORA_ZERO), errors='coerce').sum()
            total_hd_td = pd.to_timedelta(df_f[COL_HORAS_DEVIDAS].replace(['',ERRO_FORMATO,ERRO_SEQUENCIA,'nan','NaT'], HORA_ZERO), errors='coerce').sum()
            total_valor_he = df_f[COL_VALOR_HORA_EXTRA].sum()

            def fmt_td(td):
                if pd.isna(td): return "00:00"
                s, h, m = int(td.total_seconds()), 0, 0
                sign = "-" if s < 0 else ""
                s = abs(s)
                h = s // 3600
                m = (s % 3600) // 60
                return f"{sign}{h:02d}:{m:02d}"

            resumo_funcionarios[nome] = {
                "Total Horas Normais": fmt_td(total_hn_td),
                "Total Horas Extras": fmt_td(total_he_td),
                "Total Horas Devidas": fmt_td(total_hd_td),
                "Total a Receber Horas Extras": locale.format_string("%.2f", total_valor_he, grouping=True)
            }
        
        if not resumo_funcionarios: messagebox.showinfo("Resumo", "Nenhum dado para resumir.")
        else: exibir_resumo_totais(resumo_funcionarios)
        lbl_status.config(text="✅ Cálculo de totais por funcionário realizado.", fg="green")
    except Exception as e:
        lbl_status.config(text=f"❌ Erro ao calcular totais: {e}", fg="red")
        messagebox.showerror("Erro no Cálculo", f"Ocorreu um erro: {e}")
    finally:
        root.config(cursor="")

def exibir_resumo_totais(resumo_data):
    """
    Exibe uma nova janela (Toplevel) com o resumo dos totais por funcionário.

    Args:
        resumo_data (dict): Um dicionário onde as chaves são nomes de funcionários
                            e os valores são dicionários com seus totais calculados.
    Side Effects:
        Cria e mostra uma nova janela Toplevel.
        Bloqueia interação com a janela principal até ser fechada.
    """
    total_window = tk.Toplevel(root)
    total_window.title("Resumo de Totais por Funcionário")
    total_window.geometry("800x550")
    total_window.transient(root); total_window.grab_set()
    
    style_resumo = ttk.Style(total_window)
    if 'clam' in style_resumo.theme_names(): style_resumo.theme_use('clam')
    style_resumo.configure('Resumo.Treeview.Heading', font=('Calibri', 10, 'bold'))


    frame_resumo = ttk.Frame(total_window, padding="10")
    frame_resumo.pack(fill="both", expand=True)

    cols_r = ("Funcionário", "H. Normais", "H. Extras", "H. Devidas", "Valor HE (R$)")
    tree_r = ttk.Treeview(frame_resumo, columns=cols_r, show="headings", style='Resumo.Treeview')
    tree_r.pack(side="left", fill="both", expand=True)
    scrolly_r = ttk.Scrollbar(frame_resumo, orient="vertical", command=tree_r.yview)
    scrolly_r.pack(side="right", fill="y")
    tree_r.config(yscrollcommand=scrolly_r.set)

    col_widths_r = {"Funcionário": 220, "H. Normais":100, "H. Extras":100, "H. Devidas":100, "Valor HE (R$)":130}
    col_anchors_r = {"Funcionário": "w", "Valor HE (R$)": "e"}
    for col in cols_r:
        tree_r.heading(col, text=col)
        tree_r.column(col, width=col_widths_r.get(col, 100), anchor=col_anchors_r.get(col, "center"), minwidth=60)

    for nome, totais in resumo_data.items():
        tree_r.insert("", "end", values=(
            nome, totais["Total Horas Normais"], totais["Total Horas Extras"],
            totais["Total Horas Devidas"], totais["Total a Receber Horas Extras"]
        ))
    
    ttk.Button(total_window, text="Fechar", command=total_window.destroy).pack(pady=10)
    total_window.update_idletasks()
    x = (total_window.winfo_screenwidth() // 2) - (total_window.winfo_width() // 2)
    y = (total_window.winfo_screenheight() // 2) - (total_window.winfo_height() // 2)
    total_window.geometry(f'+{x}+{y}')
    root.wait_window(total_window)


def salvar_planilha():
    """
    Salva os dados atuais do DataFrame em um arquivo Excel.

    Cria uma aba "Consolidado" com todos os dados e abas individuais para cada
    funcionário, incluindo um resumo de horas e valores no final de cada aba individual.
    Exibe notificações de sucesso ou falha.

    Side Effects:
        Cria um arquivo Excel no local especificado pelo usuário.
        Atualiza `lbl_status`.
        Exibe `messagebox` de informação ou erro.
    """
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
            messagebox.showinfo("Sucesso ao Salvar", f"Planilha salva com sucesso em:\n{file_path}")
        except Exception as e:
            lbl_status.config(text=f"Erro ao salvar planilha: {e}", fg="red")
            messagebox.showerror("Erro ao Salvar", f"Não foi possível salvar a planilha:\n{e}")
        finally:
            root.config(cursor="")
    else:
        root.config(cursor="")
        lbl_status.config(text="Operação de salvar cancelada.", fg="orange")

def abrir_configuracoes():
    """
    Abre uma janela Toplevel para o usuário editar as configurações da aplicação.

    Permite alterar horas normais de trabalho e multiplicador de hora extra.
    As alterações são salvas em `config.json` e aplicadas ao DataFrame atual.

    Side Effects:
        Cria e mostra uma nova janela Toplevel.
        Pode modificar `app_config`, `config.json`, e o DataFrame global `df`.
        Pode chamar `save_config()`, `calcular_todas_horas_e_extras()`, `aplicar_filtros()`.
    """
    config_window = tk.Toplevel(root)
    config_window.title("Configurações")
    config_window.geometry("480x280")
    config_window.resizable(False, False)
    config_window.transient(root); config_window.grab_set()

    style_cfg = ttk.Style(config_window)
    if 'clam' in style_cfg.theme_names(): style_cfg.theme_use('clam')
    style_cfg.configure('.', font=('Calibri', 10))

    frame_cfg = ttk.Frame(config_window, padding="15")
    frame_cfg.pack(expand=True, fill="both")

    ttk.Label(frame_cfg, text="Horas Normais de Trabalho por Dia:").grid(row=0, column=0, sticky="w", pady=5)
    entry_hn = ttk.Entry(frame_cfg, width=10)
    entry_hn.grid(row=0, column=1, sticky="e", pady=5, padx=(10,0))
    entry_hn.insert(0, str(app_config["horas_normais_h"]).replace('.', ','))
    ttk.Label(frame_cfg, text="(Ex: 8.8 para 08:48)").grid(row=1, column=0, columnspan=2, sticky="w", padx=5, pady=(0,10))

    ttk.Label(frame_cfg, text="Multiplicador de Hora Extra:").grid(row=2, column=0, sticky="w", pady=5)
    entry_mult = ttk.Entry(frame_cfg, width=10)
    entry_mult.grid(row=2, column=1, sticky="e", pady=5, padx=(10,0))
    entry_mult.insert(0, str(app_config["multiplicador_hora_extra"]).replace('.', ','))
    ttk.Label(frame_cfg, text="(Ex: 1.5 para 50% adicional)").grid(row=3, column=0, columnspan=2, sticky="w", padx=5, pady=(0,10))

    def salvar_cfg_local():
        try:
            hn_str = entry_hn.get().replace(',', '.')
            mult_str = entry_mult.get().replace(',', '.')
            if not hn_str or not mult_str:
                messagebox.showerror("Erro", "Campos não podem ser vazios.", parent=config_window)
                return
            
            novas_hn = float(hn_str)
            novo_mult = float(mult_str)

            if not (0 < novas_hn <= 24):
                messagebox.showerror("Erro", "Horas normais entre 0 e 24.", parent=config_window)
                return
            if novo_mult <= 0:
                messagebox.showerror("Erro", "Multiplicador deve ser positivo.", parent=config_window)
                return

            app_config["horas_normais_h"] = novas_hn
            app_config["multiplicador_hora_extra"] = novo_mult
            save_config()
            
            if not df.empty:
                h_int = int(app_config["horas_normais_h"])
                m_int = int((app_config["horas_normais_h"] * 60) % 60)
                df[COL_HORAS_NORMAIS] = f"{h_int:02}:{m_int:02}"
                calcular_todas_horas_e_extras()
                aplicar_filtros()
            
            messagebox.showinfo("Sucesso", "Configurações salvas!", parent=config_window)
            config_window.destroy()
        except ValueError:
            messagebox.showerror("Erro", "Valores numéricos inválidos.", parent=config_window)
        except Exception as e_cfg:
            messagebox.showerror("Erro", f"Erro ao salvar: {e_cfg}", parent=config_window)

    frame_botoes_cfg = ttk.Frame(frame_cfg)
    frame_botoes_cfg.grid(row=4, column=0, columnspan=2, pady=(20,0), sticky="e")
    ttk.Button(frame_botoes_cfg, text="Salvar", command=salvar_cfg_local).pack(side="left", padx=5)
    ttk.Button(frame_botoes_cfg, text="Cancelar", command=config_window.destroy).pack(side="left")
    
    config_window.update_idletasks()
    x = (config_window.winfo_screenwidth() // 2) - (config_window.winfo_width() // 2)
    y = (config_window.winfo_screenheight() // 2) - (config_window.winfo_height() // 2)
    config_window.geometry(f'+{x}+{y}')
    root.wait_window(config_window)


def aplicar_filtros(event=None):
    """
    Aplica os filtros de ID, Nome e Área ao DataFrame global `df`
    e atualiza a tabela na UI com os resultados filtrados.

    Args:
        event (tk.Event, optional): Evento do Tkinter (geralmente de um bind).
                                   Não utilizado diretamente pela função, mas permite
                                   que ela seja usada como callback de evento.
    Side Effects:
        Chama `atualizar_tabela()` com o DataFrame filtrado.
        Atualiza `lbl_status`.
    """
    if df.empty:
        atualizar_tabela()
        lbl_status.config(text="ℹ️ Nenhuma planilha carregada para filtrar.", foreground="blue")
        return

    df_filtrado = df.copy()
    id_f = entry_filtro_id.get().strip().lower()
    nome_f = unicodedata.normalize('NFKD', entry_filtro_nome.get().strip().lower()).encode('ASCII', 'ignore').decode('utf-8')
    area_f = unicodedata.normalize('NFKD', entry_filtro_area.get().strip().lower()).encode('ASCII', 'ignore').decode('utf-8')

    if id_f: df_filtrado = df_filtrado[df_filtrado[COL_ID].str.lower().str.contains(id_f, na=False)]
    if nome_f:
        df_filtrado = df_filtrado[df_filtrado[COL_NOME].astype(str).apply(
            lambda x: unicodedata.normalize('NFKD', x.lower()).encode('ASCII', 'ignore').decode('utf-8')
        ).str.contains(nome_f, na=False)]
    if area_f:
        df_filtrado = df_filtrado[df_filtrado[COL_AREA].astype(str).apply(
            lambda x: unicodedata.normalize('NFKD', x.lower()).encode('ASCII', 'ignore').decode('utf-8')
        ).str.contains(area_f, na=False)]

    atualizar_tabela(df_filtrado)
    if df_filtrado.empty and (id_f or nome_f or area_f):
        lbl_status.config(text="ℹ️ Nenhum resultado para os filtros aplicados.", foreground="orange")
    elif not df_filtrado.empty :
         lbl_status.config(text=f"ℹ️ Filtros aplicados. {len(df_filtrado)} linha(s) exibida(s).", foreground="blue")
    elif df.empty: # Se o df original já estava vazio
        lbl_status.config(text="ℹ️ Nenhuma planilha carregada.", foreground="blue")
    else: # df original tem dados, mas filtro limpou tudo ou nenhum filtro aplicado
        lbl_status.config(text=f"ℹ️ Tabela atualizada. {len(df_filtrado)} linha(s) exibida(s).", foreground="blue")



def limpar_filtros():
    """
    Limpa os campos de filtro da interface e reaplica os filtros (mostrando todos os dados).

    Side Effects:
        Modifica o texto dos widgets `entry_filtro_id`, `entry_filtro_nome`, `entry_filtro_area`.
        Chama `aplicar_filtros()`.
        Atualiza `lbl_status`.
    """
    entry_filtro_id.delete(0, tk.END)
    entry_filtro_nome.delete(0, tk.END)
    entry_filtro_area.delete(0, tk.END)
    aplicar_filtros()
    lbl_status.config(text="ℹ️ Filtros limpos. Exibindo todos os dados.", foreground="blue")


def on_treeview_select(event=None):
    """
    Callback para o evento de seleção na Treeview (tabela).
    Atualiza o estado dos botões que dependem de uma seleção de linha.

    Args:
        event (tk.Event, optional): Evento do Tkinter. Não utilizado diretamente.

    Side Effects:
        Chama `update_button_states()`.
    """
    update_button_states()


# --- INTERFACE GRÁFICA ---
root = tk.Tk()
root.title("Calculadora de Ponto e Horas Extras")
root.geometry("1450x800") # Aumentei um pouco para acomodar melhor os espaçamentos

# --- ESTILO E TEMA ---
style = ttk.Style()
available_themes = style.theme_names()
# print(f"Temas disponíveis: {available_themes}") # Para debug
if 'clam' in available_themes:
    style.theme_use('clam')
elif 'alt' in available_themes:
    style.theme_use('alt')
# Adicione outros temas de fallback se desejar

style.configure('.', font=('Calibri', 10)) # Fonte padrão para todos os widgets ttk
style.configure('Treeview.Heading', font=('Calibri', 10, 'bold')) # Cabeçalhos da tabela
style.configure('TLabelframe.Label', font=('Calibri', 10, 'bold')) # Título do LabelFrame

try:
#     # Substitua "path/to/icon.png" pelo caminho real dos seus ícones
    icon_folder_img = Image.open("icons/folder-open.png").resize((16,16))
    icon_folder = ImageTk.PhotoImage(icon_folder_img)
    icon_save_img = Image.open("icons/save.png").resize((16,16))
    icon_save_action = ImageTk.PhotoImage(icon_save_img)
#     # Carregue outros ícones conforme necessário
except Exception as e_icon: 
    print(f"Erro ao carregar ícones (Pillow): {e_icon}. Usando botões sem ícones.")
icon_folder = None # Define como None se não carregar
icon_save_action = None


# --- LAYOUT DA INTERFACE ---

# 1. Frame para Ações Principais (Topo)
frame_acoes_topo = ttk.Frame(root, padding="10 5 10 5") # E, C, D, B
frame_acoes_topo.pack(fill='x')

btn_selecionar = ttk.Button(frame_acoes_topo, text="Selecionar Arquivo", command=selecionar_arquivo, image=icon_folder, compound="left")
btn_selecionar.pack(side="left", padx=(0,5)) # (padx_esq, padx_dir)

btn_salvar = ttk.Button(frame_acoes_topo, text="Salvar como Excel", command=salvar_planilha, state="disabled", image=icon_save_action, compound="left")
btn_salvar.pack(side="left", padx=5)

btn_config = ttk.Button(frame_acoes_topo, text="Configurações", command=abrir_configuracoes)
btn_config.pack(side="right", padx=5) # Alinha à direita


# 2. Frame para Filtros
frame_filtros_ui = ttk.LabelFrame(root, text="Filtros de Exibição", padding="10 10 10 10")
frame_filtros_ui.pack(fill='x', padx=10, pady=5)

ttk.Label(frame_filtros_ui, text="ID:").pack(side="left", padx=(0,2))
entry_filtro_id = ttk.Entry(frame_filtros_ui, width=12)
entry_filtro_id.pack(side="left", padx=(0,10))
entry_filtro_id.bind("<KeyRelease>", aplicar_filtros)

ttk.Label(frame_filtros_ui, text="Nome:").pack(side="left", padx=(0,2))
entry_filtro_nome = ttk.Entry(frame_filtros_ui, width=25)
entry_filtro_nome.pack(side="left", padx=(0,10))
entry_filtro_nome.bind("<KeyRelease>", aplicar_filtros)

ttk.Label(frame_filtros_ui, text="Área:").pack(side="left", padx=(0,2))
entry_filtro_area = ttk.Entry(frame_filtros_ui, width=18)
entry_filtro_area.pack(side="left", padx=(0,15))
entry_filtro_area.bind("<KeyRelease>", aplicar_filtros)

btn_limpar_filtros = ttk.Button(frame_filtros_ui, text="Limpar Filtros", command=limpar_filtros)
btn_limpar_filtros.pack(side="left")


# 3. Frame para a Tabela (Principal)
frame_tabela_ui = ttk.Frame(root, padding=(10, 0, 10, 5)) # (E, C, D, B)
frame_tabela_ui.pack(fill='both', expand=True)

tabela = ttk.Treeview(frame_tabela_ui, selectmode='browse') # browse = seleciona uma linha
tabela.bind("<<TreeviewSelect>>", on_treeview_select) # Chama a função ao selecionar

scrollbar_y = ttk.Scrollbar(frame_tabela_ui, orient="vertical", command=tabela.yview)
scrollbar_y.pack(side="right", fill="y")
tabela.configure(yscrollcommand=scrollbar_y.set)

scrollbar_x = ttk.Scrollbar(frame_tabela_ui, orient="horizontal", command=tabela.xview)
scrollbar_x.pack(side="bottom", fill="x")
tabela.configure(xscrollcommand=scrollbar_x.set)

tabela.pack(side="left", fill='both', expand=True)


# 4. Frame para Ações de Edição e Cálculo (Abaixo da tabela)
frame_acoes_edicao_calc = ttk.Frame(root, padding="10 5 10 5")
frame_acoes_edicao_calc.pack(fill='x')

btn_editar = ttk.Button(frame_acoes_edicao_calc, text="Editar Célula Sel.", command=editar_celula, state="disabled")
btn_editar.pack(side="left", padx=(0,5))

btn_excluir_id = ttk.Button(frame_acoes_edicao_calc, text="Excluir por ID Digitado", command=excluir_funcionario_por_id, state="disabled")
btn_excluir_id.pack(side="left", padx=5)

btn_remover_fds = ttk.Button(frame_acoes_edicao_calc, text="Remover Sab/Dom Sel.", command=remover_sabado_domingo_manual, state="disabled")
btn_remover_fds.pack(side="left", padx=5)

btn_calcular_totais = ttk.Button(frame_acoes_edicao_calc, text="Calcular Totais (GUI)", command=calcular_totais_funcionario, state="disabled")
btn_calcular_totais.pack(side="right", padx=5) # À direita


# 5. Barra de Status (Inferior)
lbl_status = ttk.Label(root, text="ℹ️ Pronto. Carregue uma planilha para começar.", relief=tk.SUNKEN, anchor='w', padding=5)
lbl_status.pack(side="bottom", fill="x", padx=10, pady=(0, 5))


# --- INICIALIZAÇÃO ---
load_config()
aplicar_filtros() # Para configurar a tabela e status inicial
update_button_states() # Define o estado inicial dos botões

root.mainloop()