# tests/test_calculos.py

import pytest
import pandas as pd
import numpy as np

# Para importar do diretório pai (seu_projeto/)
import sys
import os
# Adiciona o diretório pai (onde está seu_script_principal.py) ao path
# Isso permite que o Python encontre seu script principal para importação
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

# Importe as funções e constantes do seu script principal
from seu_script_principal import (
    _calculate_single_row_hours,
    app_config, # Acessaremos e modificaremos isso para testes
    COL_ID, COL_NOME, COL_AREA, COL_DATA, COL_SEMANA,
    COL_ENTRADA, COL_SAIDA_ALMOCO, COL_VOLTA_ALMOCO, COL_SAIDA,
    COL_HORAS_DEVIDAS, COL_HORAS_EXTRAS, COL_HORAS_NORMAIS,
    COL_SALARIO_BASE, COL_VALOR_HORA_EXTRA, COL_NOTA,
    ERRO_FORMATO, ERRO_SEQUENCIA, HORA_ZERO # Suas constantes de erro e hora zero
)

# --- Helper para criar uma linha de teste (row) ---
def criar_linha_teste(entrada="08:00", saida_almoco="12:00", volta_almoco="13:00", saida="17:00",
                       salario_base=2000.0, nota_inicial="", id_val="1", nome_val="Func Teste",
                       area_val="TesteArea", data_val=pd.Timestamp("2023-10-26"),
                       semana_val="Quinta-feira", horas_normais_config_val="08:00"):
    """
    Cria uma pd.Series simulando uma linha do DataFrame para os testes.
    A coluna COL_HORAS_NORMAIS na linha é apenas para completude,
    a função _calculate_single_row_hours usa app_config['horas_normais_h'].
    """
    data = {
        COL_ID: id_val, COL_NOME: nome_val, COL_AREA: area_val, COL_DATA: data_val, COL_SEMANA: semana_val,
        COL_ENTRADA: entrada, COL_SAIDA_ALMOCO: saida_almoco, COL_VOLTA_ALMOCO: volta_almoco, COL_SAIDA: saida,
        COL_HORAS_DEVIDAS: "", # Será preenchido pela função testada
        COL_HORAS_EXTRAS: "",  # Será preenchido pela função testada
        COL_HORAS_NORMAIS: horas_normais_config_val, # Valor de referência, mas a função usa app_config
        COL_SALARIO_BASE: salario_base,
        COL_VALOR_HORA_EXTRA: np.nan, # Será preenchido pela função testada
        COL_NOTA: nota_inicial
    }
    # Garante que todas as colunas que a função pode acessar existam
    return pd.Series(data)

# --- Testes para _calculate_single_row_hours ---

def test_jornada_normal_com_almoco():
    # Configuração específica para este teste (horas normais = 08:00)
    app_config["horas_normais_h"] = 8.0
    app_config["multiplicador_hora_extra"] = 1.5 # Exemplo

    linha = criar_linha_teste(entrada="09:00", saida_almoco="12:00", volta_almoco="13:00", saida="18:00", salario_base=2200.0)
    resultado = _calculate_single_row_hours(linha)

    assert resultado[COL_HORAS_DEVIDAS] == HORA_ZERO
    assert resultado[COL_HORAS_EXTRAS] == HORA_ZERO
    assert resultado[COL_VALOR_HORA_EXTRA] == 0.0
    assert ERRO_FORMATO not in resultado[COL_NOTA]
    assert ERRO_SEQUENCIA not in resultado[COL_NOTA]

def test_horas_extras_com_almoco():
    app_config["horas_normais_h"] = 8.0
    app_config["multiplicador_hora_extra"] = 1.5

    # Trabalhou 9 horas (1 hora extra)
    linha = criar_linha_teste(entrada="09:00", saida_almoco="12:00", volta_almoco="13:00", saida="19:00", salario_base=2200.0)
    resultado = _calculate_single_row_hours(linha)

    assert resultado[COL_HORAS_DEVIDAS] == HORA_ZERO
    assert resultado[COL_HORAS_EXTRAS] == "01:00"
    # Valor/hora = 2200/220 = 10. Valor HE = 10 * 1.5 * 1h = 15.00
    assert resultado[COL_VALOR_HORA_EXTRA] == 15.00
    assert ERRO_FORMATO not in resultado[COL_NOTA]

def test_horas_devidas_sem_almoco():
    app_config["horas_normais_h"] = 8.8 # 08:48
    app_config["multiplicador_hora_extra"] = 1.5

    # Trabalhou 08:00, devendo 00:48
    linha = criar_linha_teste(entrada="08:00", saida_almoco="00:00", volta_almoco="00:00", saida="16:00", salario_base=2200.0)
    resultado = _calculate_single_row_hours(linha)

    assert resultado[COL_HORAS_DEVIDAS] == "00:48"
    assert resultado[COL_HORAS_EXTRAS] == HORA_ZERO
    assert resultado[COL_VALOR_HORA_EXTRA] == 0.0
    assert ERRO_FORMATO not in resultado[COL_NOTA]

def test_erro_formato_hora_entrada():
    app_config["horas_normais_h"] = 8.0
    linha = criar_linha_teste(entrada="INVALIDO", saida="17:00")
    resultado = _calculate_single_row_hours(linha)

    assert resultado[COL_HORAS_DEVIDAS] == ERRO_FORMATO
    assert resultado[COL_HORAS_EXTRAS] == ERRO_FORMATO
    assert ERRO_FORMATO.lower() in resultado[COL_NOTA].lower() # Verifica se a nota contém o erro

def test_erro_sequencia_saida_antes_entrada():
    app_config["horas_normais_h"] = 8.0
    linha = criar_linha_teste(entrada="18:00", saida="08:00", saida_almoco="00:00", volta_almoco="00:00") # Saída antes da entrada, sem cruzar meia-noite na lógica simples
    resultado = _calculate_single_row_hours(linha)
    # Sua lógica _calculate_single_row_hours já trata overnight, então este teste pode precisar ser ajustado
    # Se a intenção é testar a lógica de erro de sequência quando saida_final_dt < entrada_dt (antes do ajuste de +1 dia)
    # o resultado esperado seria ERRO_SEQUENCIA. Se a lógica sempre ajusta, pode dar um resultado diferente.
    # Ajustando para um erro de sequência claro sem almoço:
    linha_erro_seq = criar_linha_teste(entrada="10:00", saida_almoco="00:00", volta_almoco="00:00", saida="09:00")
    resultado_erro_seq = _calculate_single_row_hours(linha_erro_seq)

    assert resultado_erro_seq[COL_HORAS_DEVIDAS] == ERRO_SEQUENCIA
    assert resultado_erro_seq[COL_HORAS_EXTRAS] == ERRO_SEQUENCIA
    assert ERRO_SEQUENCIA.lower() in resultado_erro_seq[COL_NOTA].lower()

def test_todos_horarios_zerados_ou_vazios():
    app_config["horas_normais_h"] = 8.0
    linha = criar_linha_teste(entrada="00:00", saida_almoco="00:00", volta_almoco="00:00", saida="00:00")
    resultado = _calculate_single_row_hours(linha)
    assert resultado[COL_HORAS_DEVIDAS] == "" # Ou HORA_ZERO dependendo da sua lógica para ausência
    assert resultado[COL_HORAS_EXTRAS] == "" # Ou HORA_ZERO
    assert resultado[COL_VALOR_HORA_EXTRA] == 0.0

    linha_vazia = criar_linha_teste(entrada="", saida_almoco="", volta_almoco="", saida="")
    resultado_vazio = _calculate_single_row_hours(linha_vazia)
    assert resultado_vazio[COL_HORAS_DEVIDAS] == ""
    assert resultado_vazio[COL_HORAS_EXTRAS] == ""
    assert resultado_vazio[COL_VALOR_HORA_EXTRA] == 0.0

# Adicione mais cenários:
# - Horários noturnos que cruzam meia-noite
# - Almoço que cruza meia-noite (se aplicável)
# - Casos onde apenas alguns horários estão preenchidos
# - Sem salário base (valor hora extra deve ser 0)
# - Com anotações pré-existentes (verificar se a nota é concatenada corretamente)

# --- Exemplo de teste para outra função (se você tiver helpers) ---
# from seu_script_principal import sua_funcao_helper
# def test_sua_funcao_helper():
#     assert sua_funcao_helper(entrada_exemplo) == resultado_esperado