# Calculadora de Ponto e Horas Extras - Ponto Eletrônico Knup 1028

## Descrição Curta

A **Calculadora de Ponto e Horas Extras** é uma aplicação de desktop desenvolvida em Python com Tkinter. Ela permite aos usuários carregar planilhas de ponto de funcionários (em formato Excel), visualizar e editar os dados, e automaticamente calcular horas trabalhadas, horas devidas, horas extras e o valor monetário correspondente a essas horas extras. O software também gera um arquivo Excel consolidado com os dados processados e abas individuais para cada funcionário contendo seus registros e um resumo dos totais.

## Funcionalidades Principais

* **Carregamento de Planilhas:** Importa dados de ponto a partir de arquivos Excel (`.xlsx`, `.xls`).
* **Cálculo Automático:**
    * Horas Devidas
    * Horas Extras
    * Valor a Receber por Horas Extras (com base em salário e multiplicador configuráveis)
* **Visualização e Edição:**
    * Exibe os dados em uma tabela interativa.
    * Permite a edição direta de células (ID, Nome, Área, Data, horários, Salário Base, Notas).
    * Recálculo automático após edições que impactam as horas.
* **Filtragem de Dados:** Filtra os registros por ID, Nome ou Área do funcionário.
* **Gerenciamento de Dados:**
    * Exclui todos os registros de um funcionário por ID.
    * Remove registros selecionados que correspondam a Sábados ou Domingos.
* **Relatório de Totais:** Exibe uma janela com o resumo de horas normais, extras, devidas e valor total de HE por funcionário.
* **Exportação para Excel:**
    * Gera um arquivo Excel com uma aba "Consolidado" contendo todos os dados processados.
    * Cria abas individuais para cada funcionário com seus respectivos registros e um resumo de totais (horas normais, extras, devidas e valor de HE).
* **Configurações Personalizáveis:**
    * Permite definir as horas normais de trabalho diárias.
    * Permite definir o multiplicador para cálculo do valor da hora extra.
    * As configurações são salvas em um arquivo `config.json`.
* **Interface Gráfica Intuitiva:** Desenvolvida com Tkinter e `ttk` para uma melhor experiência do usuário.

## Capturas de Tela (Opcional)

*Sugestão: Adicione aqui algumas capturas de tela da aplicação em funcionamento.*

* `[Link ou Imagem da Tela Principal com Dados Carregados]`
* `[Link ou Imagem da Janela de Resumo de Totais]`
* `[Link ou Imagem da Janela de Configurações]`

## Como Usar (Guia Rápido)

1.  **Carregar Planilha:**
    * Clique no botão "Selecionar Arquivo".
    * Escolha a planilha Excel contendo os dados de ponto.
    * Os dados serão carregados na tabela e os cálculos iniciais realizados.
2.  **Visualizar e Filtrar:**
    * Utilize as barras de rolagem para navegar pela tabela.
    * Use os campos de "Filtros de Exibição" (ID, Nome, Área) para refinar os dados mostrados. Clique em "Limpar Filtros" para ver todos os dados novamente.
3.  **Editar Dados:**
    * Selecione uma linha na tabela.
    * Clique em "Editar Célula Sel.".
    * Siga as instruções para escolher a coluna e inserir o novo valor.
    * *Nota:* Alterações em horários ou salário base acionarão o recálculo automático para a linha.
4.  **Outras Ações:**
    * **Excluir por ID Digitado:** Remove todos os registros de um ou mais IDs especificados.
    * **Remover Sab/Dom Sel.:** Remove as linhas selecionadas que forem Sábados ou Domingos.
    * **Calcular Totais (GUI):** Abre uma janela com o resumo de horas e valores por funcionário.
5.  **Configurações:**
    * Clique em "Configurações" para ajustar as horas normais de trabalho e o multiplicador de hora extra.
    * As alterações são salvas e aplicadas imediatamente se houver dados carregados.
6.  **Salvar Resultados:**
    * Clique em "Salvar como Excel".
    * Escolha o local e nome para o novo arquivo Excel.
    * Uma notificação indicará o sucesso ou falha da operação.

## Formato da Planilha de Entrada

Para que o software funcione corretamente, a planilha Excel de entrada deve seguir um formato específico:

* **Aba de Leitura:** Os dados devem estar na **terceira aba** da planilha (índice 2).
* **Cabeçalho:** O software assume que as **primeiras 4 linhas** da planilha são cabeçalhos ou informações não relevantes e as ignora. A leitura dos dados começa a partir da quinta linha.
* **Colunas Esperadas (na ordem):**
    1.  `ID`: Identificador único do funcionário (Texto/Número).
    2.  `Nome`: Nome completo do funcionário (Texto).
    3.  `Área`: Área/Departamento do funcionário (Texto).
    4.  `Data`: Data do registro de ponto (Formato DD/MM/AAAA).
    5.  `Entrada`: Horário de entrada (Formato HH:MM, ex: "08:00").
    6.  `Saída-Almoço`: Horário de saída para o almoço (Formato HH:MM).
    7.  `Volta-Almoço`: Horário de volta do almoço (Formato HH:MM).
    8.  `Saída`: Horário de saída final (Formato HH:MM).
    9.  `Horas Devidas`: (Esta coluna pode vir da planilha original, mas será recalculada).
    10. `Horas Extras`: (Esta coluna pode vir da planilha original, mas será recalculada).
    11. `Horas Normais`: (Esta coluna pode vir da planilha original, mas será recalculada com base na configuração).
    12. `Nota`: Campo para anotações diversas (Texto).
* **Valores Especiais em Horários:**
    * Campos de horário vazios, ou contendo "Omissão" (ou variações), "nan", são tratados como ausência de marcação.
    * Para indicar que não houve intervalo de almoço, os campos `Saída-Almoço` e `Volta-Almoço` podem ser deixados em branco ou preenchidos com "00:00".

## Arquivo de Configuração (`config.json`)

O software utiliza um arquivo `config.json` para armazenar as configurações de cálculo. Este arquivo é criado automaticamente na primeira execução ou se for deletado, e fica no mesmo diretório do executável (ou script).

* **`horas_normais_h`**: Número decimal de horas que compõem uma jornada normal de trabalho diária.
    * Exemplo: `8.8` (para 08 horas e 48 minutos), `8.0` (para 08 horas).
* **`multiplicador_hora_extra`**: Fator pelo qual o valor da hora normal é multiplicado para calcular o valor da hora extra.
    * Exemplo: `1.5` (para um adicional de 50%), `2.0` (para um adicional de 100%).

## Lógica de Cálculo de Horas (Resumo)

* O tempo total trabalhado é calculado com base nos horários de entrada, saída e almoço.
* Intervalos de almoço com "00:00" ou vazios são considerados como dia trabalhado sem pausa para almoço.
* A diferença entre o tempo trabalhado e as `horas_normais_h` configuradas determina se há horas devidas ou extras.
* O valor da hora extra é calculado como: `(Salário Base / 220) * multiplicador_hora_extra * (total de horas extras em decimal)`.
* **Códigos de Erro nas colunas de horas:**
    * `INV_FORMATO`: Indica que um dos horários fornecidos está em formato inválido (diferente de HH:MM).
    * `INV_SEQ`: Indica uma inconsistência na sequência dos horários (ex: saída antes da entrada sem ser um turno noturno corretamente configurado, ou volta do almoço antes da saída para o almoço).

## Pré-requisitos (para executar o script Python)

* Python 3.8 ou superior.
* As seguintes bibliotecas Python:
    * `pandas`
    * `XlsxWriter` (para salvar em formato `.xlsx` com formatação avançada)
    * `(Opcional)` `Pillow` (se você habilitar os ícones nos botões)

Recomenda-se o uso de um ambiente virtual Python.

## Instalação (das dependências)

1.  **Crie um Ambiente Virtual (Recomendado):**
    ```bash
    python -m venv venv_ponto
    # No Windows:
    venv_ponto\Scripts\activate
    # No macOS/Linux:
    source venv_ponto/bin/activate
    ```
2.  **Crie um arquivo `requirements.txt`** na raiz do seu projeto com o seguinte conteúdo:
    ```
    pandas
    XlsxWriter
    # Pillow # Adicione se for usar ícones
    ```
3.  **Instale as Dependências:**
    ```bash
    pip install -r requirements.txt
    ```

## Como Executar o Script

Após instalar as dependências (dentro do ambiente virtual, se estiver usando um):
```bash
python calculadora_ponto_main.py
