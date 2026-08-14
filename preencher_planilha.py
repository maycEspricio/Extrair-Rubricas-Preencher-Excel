# =========================
# IMPORTAÇÕES
# =========================

# Biblioteca padrão para leitura do arquivo JSON gerado na etapa anterior.
# Esse JSON contém a rubrica extraída do Google Classroom.
import json
import re

# Biblioteca padrão usada aqui para arredondamento da mediana das capacidades.
# O uso de math.floor(mediana + 0.5) implementa o arredondamento "tradicional":
# valores com .5 sobem para o próximo inteiro.
import math

# Biblioteca para automação do Excel via COM no Windows.
# Permite controlar uma instância já aberta do Excel Desktop:
# ler células, escrever valores, acessar planilhas, etc.
import win32com.client as win32

from statistics import median

# =========================
# CONFIGURAÇÕES FIXAS DO ARQUIVO/PLANILHA
# =========================
#
# Essas constantes definem a "estrutura esperada" da planilha.
# Se o layout do Excel mudar, muito provavelmente será necessário
# ajustar uma ou mais dessas constantes.
#

# Número da coluna onde estão os textos dos critérios e também das capacidades.
# No Excel, a coluna 7 corresponde à coluna G.
COLUNA_CRITERIO = 7

# Primeira coluna da faixa de avaliação onde as marcações serão feitas.
# Como a pontuação interna vai de 1 a 20, e a faixa começa na coluna 8,
# então a coluna 8 representa a pontuação 1.
COL_INICIO_AVALIACAO = 8

# Última coluna da faixa de avaliação.
# Como a faixa vai de 8 a 27, há 20 colunas disponíveis, uma para cada pontuação.
COL_FIM_AVALIACAO = 27

# Linha em que a busca por critérios/capacidades deve começar.
# Conforme layout do professor, inicia-se na linha 12.
LINHA_INICIAL_BUSCA = 12

LINHA_NOME_ALUNO = 10
COLUNA_NOME_ALUNO = 13  # coluna M

# Ordem lógica dos níveis de desempenho.
# Essa ordem é importante porque o sistema pergunta os subníveis na sequência esperada
# e também porque a lógica de pontuação depende dessa organização conceitual.
NIVEIS_ORDENADOS = [
    "não satisfatório",
    "apoiado",
    "parcialmente autônomo",
    "autônomo",
]

# Lista de graus válidos (em minúsculas) que podem vir do Classroom.
# O Classroom agora retorna o nível do critério como um grau (G1–G5).
GRAUS_VALIDOS = ["g1", "g2", "g3", "g4", "g5"]


# =========================
# FUNÇÕES AUXILIARES DE TEXTO E LEITURA
# =========================

def extrair_numero_grau(nivel_texto: str):
    """
    Extrai o número (1 a 5) de uma string que representa um grau do Classroom.

    Aceita variações como:
        "G1", "g1", "Grau 1", "grau 3", "G 4", "5"

    Retorno:
        int entre 1 e 5 se o grau for reconhecido, ou None caso contrário.

    Exemplos:
        extrair_numero_grau("G3")     -> 3
        extrair_numero_grau("grau 2") -> 2
        extrair_numero_grau("Apoiado")-> None
    """
    import re
    if not nivel_texto:
        return None
    texto = str(nivel_texto).strip().lower()
    # Tenta casar padrões como "g1", "g 1", "grau 1", "grau1" ou apenas "1"
    match = re.search(r'(?:grau\s*|g\s*)(\d)', texto)
    if match:
        numero = int(match.group(1))
        if 1 <= numero <= 5:
            return numero
    # Tenta número isolado
    match = re.fullmatch(r'\d', texto)
    if match:
        numero = int(match.group())
        if 1 <= numero <= 5:
            return numero
    return None


def normalizar(texto):
    """
    Normaliza um texto para comparação.

    Objetivo:
    padronizar valores vindos de diferentes fontes (Excel, JSON, input do usuário),
    reduzindo as chances de erro por diferença de formatação.

    O que esta função faz:
    - converte None em string vazia
    - transforma qualquer valor em string
    - remove espaços duplicados entre palavras
    - remove espaços no início e no fim
    - converte tudo para minúsculas
    - remove "expandir critério" e "expandir criterio" de forma resiliente

    Exemplo:
    "Expandir critério Parcialmente Autônomo" -> "parcialmente autônomo"

    Isso ajuda porque, sem normalização, textos visualmente iguais
    podem falhar numa comparação simples.
    """
    txt = " ".join(str(texto or "").split()).strip().lower()
    txt = re.sub(r'expandir\s+crit[eé]rio', '', txt)
    return " ".join(txt.split()).strip()


def carregar_rubricas(nome_arquivo="rubricas.json"):
    with open(nome_arquivo, "r", encoding="utf-8") as f:
        return json.load(f)


# =========================
# LÓGICA DE PONTUAÇÃO
# =========================
#
# A planilha representa a avaliação em uma escala interna de 1 a 20.
# Cada nível principal ocupa um bloco de 5 posições.
#
# Mapeamento conceitual:
# - não satisfatório       ->  1 a  5
# - apoiado                ->  6 a 10
# - parcialmente autônomo  -> 11 a 15
# - autônomo               -> 16 a 20
#
# Dentro de cada nível, o usuário ainda informa um subnível de 1 a 5.
# Isso refina a marcação final.
#

def obter_inicio_faixa(nivel_texto: str) -> int:
    nivel = normalizar(nivel_texto)

    mapa = {
        "não satisfatório": 1,
        "apoiado": 6,
        "parcialmente autônomo": 11,
        "autônomo": 16,
    }

    if nivel not in mapa:
        raise ValueError(f"Nível desconhecido: {nivel_texto}")

    return mapa[nivel]


def obter_pontuacao_interna(nivel_texto: str, nivel_1_a_5: int) -> int:
    if nivel_1_a_5 < 1 or nivel_1_a_5 > 5:
        raise ValueError("O nível informado deve estar entre 1 e 5.")

    return obter_inicio_faixa(nivel_texto) + (nivel_1_a_5 - 1)


def obter_coluna_excel_por_pontuacao(pontuacao_interna: int) -> int:
    if pontuacao_interna < 1 or pontuacao_interna > 20:
        raise ValueError("A pontuação interna deve estar entre 1 e 20.")

    return pontuacao_interna + 7


# =========================
# CONEXÃO E CONTEXTO DO EXCEL
# =========================

def conectar_excel_aberto():
    """
    Conecta a uma instância do Excel que já esteja aberta no Windows.

    Este código NÃO cria um novo Excel.
    Ele se conecta a um Excel já em execução.

    Isso é útil porque o usuário pode:
    - abrir manualmente a planilha correta
    - conferir se está na aba certa
    - deixar o arquivo já preparado antes da automação começar

    Retorno:
        objeto COM representando Excel.Application

    Observação:
    Se o Excel não estiver aberto, a chamada poderá falhar.
    """
    return win32.GetActiveObject("Excel.Application")


def escolher_planilha(workbook, nome_aba=None):
    """
    Define qual planilha (worksheet) será usada no preenchimento.

    Regras:
    - se nome_aba for informado, usa explicitamente essa aba
    - se nome_aba for None, usa a aba ativa no momento

    Parâmetros:
        workbook: pasta de trabalho do Excel já aberta.
        nome_aba (str | None): nome da aba a ser usada.

    Retorno:
        objeto worksheet do Excel

    Observação:
    Usar a aba ativa é conveniente, mas exige que o usuário
    tenha deixado a aba correta selecionada antes de rodar.
    """
    if nome_aba:
        return workbook.Worksheets(nome_aba)
    return workbook.ActiveSheet


# =========================
# BUSCA DE LINHAS E LIMPEZA DE FAIXAS
# =========================

def limpar_linha_faixa(sheet, linha):
    """
    Limpa qualquer marcação existente nas colunas de notas (H a AA, colunas 8 a 27) para uma determinada linha.
    Evita que o aluno fique com múltiplas notas na mesma linha de critério ou capacidade.
    """
    try:
        sheet.Range(sheet.Cells(linha, 8), sheet.Cells(linha, 27)).Value = ""
    except Exception as e:
        print(f"[ERRO LIMPEZA] Falha ao limpar faixa da linha {linha}: {e}")

def encontrar_linha_por_criterio(sheet, criterio, coluna_criterio=COLUNA_CRITERIO, linha_inicial=1, linha_final=None, col_values=None):
    criterio_norm = normalizar(criterio)

    if col_values is not None:
        # Busca em memória para ser extremamente rápido e evitar erros COM de "Excel Ocupado"
        for idx in range(linha_inicial - 1, len(col_values)):
            linha_valores = col_values[idx]
            valor = linha_valores[0] if isinstance(linha_valores, (list, tuple)) else linha_valores
            texto = normalizar(valor)
            if criterio_norm and criterio_norm in texto:
                return idx + 1
        return None

    if linha_final is None:
        linha_final = obter_linha_final(sheet)

    for linha in range(linha_inicial, linha_final + 1):
        valor = sheet.Cells(linha, coluna_criterio).Value
        texto = normalizar(valor)

        if criterio_norm and criterio_norm in texto:
            return linha

    return None


def debug_coluna_nomes(sheet, coluna_nome=COLUNA_NOME_ALUNO, limite=200):
    """
    Mostra no terminal os valores não vazios da coluna do nome do aluno.
    Útil para descobrir em qual linha o nome realmente aparece.
    """
    linha_final = min(obter_linha_final(sheet), limite)

    print(f"\n[DEBUG] Planilha: {sheet.Name}")
    print(f"[DEBUG] Lendo coluna {coluna_nome} até a linha {linha_final}")

    encontrou_algo = False

    for linha in range(1, linha_final + 1):
        valor = sheet.Cells(linha, coluna_nome).Value
        texto = str(valor).strip() if valor is not None else ""

        if texto:
            encontrou_algo = True
            print(f"[DEBUG] Linha {linha} | Coluna {coluna_nome} = {texto}")

    if not encontrou_algo:
        print("[DEBUG] Nenhum valor não vazio encontrado na coluna do nome.")

def encontrar_aluno_na_planilha(sheet, nome_aluno, linha_nome=LINHA_NOME_ALUNO, coluna_nome=COLUNA_NOME_ALUNO, debug=False):
    """
    Verifica se o nome do aluno está na célula fixa M10 da aba.
    Retorna True se a aba pertence ao aluno.
    """
    valor = sheet.Cells(linha_nome, coluna_nome).Value
    texto = normalizar(valor)
    nome_aluno_norm = normalizar(nome_aluno)

    if debug:
        print(f"[DEBUG] Aba: {sheet.Name} | Célula ({linha_nome}, {coluna_nome}) = {valor}")

    if not texto:
        return False

    if texto == nome_aluno_norm:
        print(f"[ALUNO ENCONTRADO] {nome_aluno} -> aba '{sheet.Name}'")
        return True

    if nome_aluno_norm in texto or texto in nome_aluno_norm:
        print(f"[ALUNO ENCONTRADO - APROX] {nome_aluno} -> aba '{sheet.Name}' | célula='{valor}'")
        return True

    return False

def encontrar_planilha_do_aluno(workbook, nome_aluno, debug=False):
    """
    Percorre todas as abas e verifica se o nome do aluno está em M10.
    Retorna apenas a planilha correspondente.
    """
    for sheet in workbook.Worksheets:
        if encontrar_aluno_na_planilha(sheet, nome_aluno, debug=debug):
            return sheet

    return None

def perguntar_subniveis_por_aluno(rubricas_alunos, solicitar_subnivel_callback=None, should_stop_callback=None):
    """
    Pergunta o subnível (1 a 5) para cada desempenho de cada aluno.
    Retorna um dicionário no formato:
    {
        "Nome do Aluno": {
            "apoiado": 3,
            "autônomo": 5
        }
    }
    """
    respostas = {}

    for aluno in rubricas_alunos:
        if should_stop_callback is not None and should_stop_callback():
            print("Processo interrompido pelo usuário antes de concluir os subníveis.")
            break

        nome_aluno = aluno.get("aluno", "").strip()
        rubrica = aluno.get("rubrica", [])

        if not nome_aluno or not rubrica:
            continue

        respostas[nome_aluno] = {}
        niveis_presentes = []
        vistos = set()

        for item in rubrica:
            nivel = normalizar(item.get("nivel"))
            if nivel in NIVEIS_ORDENADOS and nivel not in vistos:
                vistos.add(nivel)
                niveis_presentes.append(nivel)

        if not niveis_presentes:
            continue

        print("\n" + "=" * 70)
        print(f"Aluno: {nome_aluno}")
        print("=" * 70)

        for nivel in NIVEIS_ORDENADOS:
            if nivel not in niveis_presentes:
                continue

            while True:
                if should_stop_callback is not None and should_stop_callback():
                    print("Processo interrompido pelo usuário durante a coleta de subníveis.")
                    return respostas

                try:
                    if solicitar_subnivel_callback is None:
                        valor = int(
                            input(f"Aluno: {nome_aluno} | Desempenho: {nivel.title()} -> digite um valor de 1 a 5: ").strip()
                        )
                    else:
                        # No modo interface gráfica, o valor é pedido em janela.
                        # O callback deve retornar um inteiro de 1 a 5.
                        valor = int(solicitar_subnivel_callback(nome_aluno, nivel))

                    if 1 <= valor <= 5:
                        respostas[nome_aluno][nivel] = valor
                        break

                    print("Digite um número entre 1 e 5.")

                except ValueError:
                    print("Digite um número válido.")

    return respostas

def escrever_celula_com_retry(sheet, linha, coluna, valor, retries=15, delay=0.5):
    import time
    for i in range(retries):
        try:
            sheet.Cells(linha, coluna).Value = valor
            return
        except Exception as e:
            if "800ac472" in str(e) or "-2146777998" in str(e):
                time.sleep(delay)
            else:
                raise e
    sheet.Cells(linha, coluna).Value = valor

def limpar_faixa(sheet, linha, col_inicio=COL_INICIO_AVALIACAO, col_fim=COL_FIM_AVALIACAO):
    import time
    for i in range(15):
        try:
            sheet.Range(f"H{linha}:AA{linha}").Value = ""
            return
        except Exception as e:
            if "800ac472" in str(e) or "-2146777998" in str(e):
                time.sleep(0.5)
            else:
                raise e
    sheet.Range(f"H{linha}:AA{linha}").Value = ""

def obter_linha_final(sheet):
    """
    Retorna a última linha útil da planilha com base no UsedRange.

    Fórmula:
        última_linha = quantidade_de_linhas_usadas + linha_inicial_do_UsedRange - 1

    Isso é necessário porque o UsedRange pode começar depois da linha 1.

    Retorno:
        int: número da última linha considerada usada.

    Observação:
    UsedRange nem sempre é perfeito em planilhas que já tiveram conteúdo apagado,
    então esse valor é prático, mas não absolutamente infalível.
    """
    return sheet.UsedRange.Rows.Count + sheet.UsedRange.Row - 1


# =========================
# INTERAÇÃO COM O USUÁRIO
# =========================

def perguntar_subniveis_por_desempenho(dados):
    """
    Pergunta ao usuário o subnível (1 a 5) de cada nível de desempenho presente.

    Contexto:
    A rubrica extraída informa apenas o nível principal, por exemplo:
    - apoiado
    - autônomo

    Mas, para marcar corretamente a planilha, o sistema precisa transformar isso
    em uma pontuação de 1 a 20. Para isso, cada nível principal precisa ser refinado
    por um subnível de 1 a 5, informado manualmente pelo usuário.

    Estratégia:
    1. percorre os dados da rubrica
    2. identifica quais níveis aparecem de fato
    3. evita repetir perguntas para níveis iguais
    4. pergunta apenas os níveis presentes
    5. valida a resposta para garantir valor entre 1 e 5

    Parâmetros:
        dados (list[dict]): dados da rubrica carregados do JSON.

    Retorno:
        dict: mapeamento no formato:
            {
                "apoiado": 3,
                "autônomo": 5
            }

    Observação:
    A pergunta é feita na ordem de NIVEIS_ORDENADOS, não na ordem em que aparecem no JSON.
    """

    # Lista com os níveis encontrados nos dados.
    # Será usada apenas para saber o que perguntar ao usuário.
    niveis_presentes = []

    # Conjunto para impedir repetição.
    vistos = set()

    # Percorre todos os itens da rubrica para descobrir quais níveis aparecem.
    for item in dados:
        nivel = normalizar(item["nivel"])

        # Só registra níveis reconhecidos e ainda não vistos.
        if nivel in NIVEIS_ORDENADOS and nivel not in vistos:
            vistos.add(nivel)
            niveis_presentes.append(nivel)

    # Dicionário final com as respostas do usuário.
    respostas = {}

    print("\nInforme agora o subnível (1 a 5) para cada desempenho encontrado:\n")

    # Percorre os níveis na ordem padrão da aplicação.
    for nivel in NIVEIS_ORDENADOS:
        # Pula níveis que não apareceram nos dados extraídos.
        if nivel not in niveis_presentes:
            continue

        # Loop de validação: só sai quando o usuário digitar algo válido.
        while True:
            try:
                valor = int(input(f"{nivel.title()} -> digite um valor de 1 a 5: ").strip())

                # Garante faixa válida.
                if 1 <= valor <= 5:
                    respostas[nivel] = valor
                    break

                print("Digite um número entre 1 e 5.")

            except ValueError:
                # Captura casos como letras, vazio, símbolos etc.
                print("Digite um número válido.")

    return respostas


# =========================
# PREENCHIMENTO DOS CRITÉRIOS INDIVIDUAIS
# =========================

def mapear_estrutura_planilha(sheet):
    """
    Varre a planilha a partir da linha 12 para mapear a estrutura física:
    - Lista de linhas de critérios (na ordem de aparecimento).
    - Lista de blocos de capacidades com suas respectivas linhas de critérios.
    """
    linha_final = obter_linha_final(sheet)
    if linha_final < LINHA_INICIAL_BUSCA:
        return [], []

    # Lê todas as colunas de B a G (2 a 7) usando A1-notation
    valores = sheet.Range(f"B{LINHA_INICIAL_BUSCA}:G{linha_final}").Value
    if not valores:
        return [], []

    criterios_linhas = []
    capacidades_blocos = []
    capacidade_atual = None

    for i, linha_valores in enumerate(valores):
        linha_real = i + LINHA_INICIAL_BUSCA

        val_b = linha_valores[0]
        val_c = linha_valores[1]
        val_d = linha_valores[2]
        val_e = linha_valores[3]
        val_f = linha_valores[4]
        texto = linha_valores[5] # Coluna G (7)

        tem_bcd = (val_b is not None and str(val_b).strip() != "") or \
                  (val_c is not None and str(val_c).strip() != "") or \
                  (val_d is not None and str(val_d).strip() != "")

        tem_ef = (val_e is not None and str(val_e).strip() != "") or \
                 (val_f is not None and str(val_f).strip() != "")

        if tem_bcd:
            capacidade_atual = {
                "linha_capacidade": linha_real,
                "texto_capacidade": str(texto).strip() if texto else f"Capacidade Linha {linha_real}",
                "criterios": []
            }
            capacidades_blocos.append(capacidade_atual)
        elif tem_ef:
            criterios_linhas.append(linha_real)
            if capacidade_atual:
                capacidade_atual["criterios"].append(linha_real)

    return criterios_linhas, [b for b in capacidades_blocos if b["criterios"]]


def preencher_criterios(sheet, dados, respostas_subnivel_aluno, capacidades_blocos=None):
    # Filtra os dados para remover itens de controle do Classroom
    dados_filtrados = []
    for item in dados:
        crit_text = normalizar(item.get("criterio", ""))
        if not crit_text or crit_text == "expandir critério" or "critério não encontrado" in crit_text:
            continue
        dados_filtrados.append(item)

    # Obtém todas as linhas de critérios físicas mapeadas na planilha
    criterios_linhas, _ = mapear_estrutura_planilha(sheet)

    preenchidos = 0
    nao_encontrados = []
    criterios_preenchidos = []

    for idx, item in enumerate(dados_filtrados):
        if idx >= len(criterios_linhas):
            print(f"[AVISO] Mais critérios no Classroom ({len(dados_filtrados)}) do que linhas na planilha ({len(criterios_linhas)}).")
            break

        linha = criterios_linhas[idx]
        criterio = item["criterio"]
        nivel = normalizar(item["nivel"])

        is_nao_atingiu = (nivel == "não atingiu")

        if nivel not in respostas_subnivel_aluno and not is_nao_atingiu:
            print(f"[IGNORADO] {criterio} na linha {linha} -> nível sem resposta: {item['nivel']}")
            continue

        try:
            if is_nao_atingiu:
                pontuacao_interna = 1
            else:
                pontuacao_interna = obter_pontuacao_interna(nivel, respostas_subnivel_aluno[nivel])
            coluna_destino = obter_coluna_excel_por_pontuacao(pontuacao_interna)
        except ValueError as e:
            print(f"[IGNORADO] {criterio} na linha {linha} -> {e}")
            continue

        # Limpa marcações anteriores na faixa de colunas antes de escrever a nova
        limpar_linha_faixa(sheet, linha)
        escrever_celula_com_retry(sheet, linha, coluna_destino, "✓")

        criterios_preenchidos.append({
            "linha": linha,
            "criterio": criterio,
            "nivel": nivel,
            "pontuacao": pontuacao_interna,
            "coluna_excel": coluna_destino,
        })

        preenchidos += 1
        print(f"[OK] {criterio} -> linha {linha}, coluna {coluna_destino}")

    return criterios_preenchidos, preenchidos, nao_encontrados


def arredondar_media_para_int(mediana):
    import math
    return int(math.floor(mediana + 0.5))


# O preencher_capacidades agora verifica a planilha fisicamente (colunas 8 a 27) para decidir se calcula a capacidade
def preencher_capacidades(sheet, capacidades_blocos):
    print("\nCalculando capacidades com base nas notas da planilha...\n")

    linha_final = obter_linha_final(sheet)
    if linha_final < LINHA_INICIAL_BUSCA:
        return

    # Lê todas as notas da planilha de uma vez (linhas 12 até linha_final, colunas H a AA, ou seja, 8 a 27)
    valores_notas = sheet.Range(f"H{LINHA_INICIAL_BUSCA}:AA{linha_final}").Value
    if not valores_notas:
        return

    for bloco in capacidades_blocos:
        pontuacoes = []
        todos_avaliados = True

        for c_linha in bloco["criterios"]:
            # Mapeia c_linha para o index correspondente na tupla valores_notas
            idx_linha = c_linha - LINHA_INICIAL_BUSCA
            if idx_linha < 0 or idx_linha >= len(valores_notas):
                todos_avaliados = False
                break

            linha_valores = valores_notas[idx_linha]
            
            # Procura se há um "✓" nesta linha nas colunas de nota
            nota_encontrada = None
            # Trata caso de linha única retornada como tupla plana pelo win32com
            if not isinstance(linha_valores, (list, tuple)):
                # Se for valor único
                if linha_valores == "✓":
                    # Nota única (precisamos saber a coluna do range, mas nesse caso a linha só tem 1 coluna, improvável)
                    nota_encontrada = 1
            else:
                for col_idx, val in enumerate(linha_valores):
                    if val == "✓":
                        nota_encontrada = col_idx + 1 # pontuação é de 1 a 20
                        break
            
            if nota_encontrada is not None:
                pontuacoes.append(nota_encontrada)
            else:
                todos_avaliados = False
                break

        if not todos_avaliados:
            print(f"[CAPACIDADE IGNORADA] '{bloco['texto_capacidade']}' -> nem todos os {len(bloco['criterios'])} critérios foram preenchidos na planilha.")
            continue

        mediana = median(pontuacoes)
        if mediana <= 10:
            mediana = 11

        mediana_final = arredondar_media_para_int(mediana)
        coluna = obter_coluna_excel_por_pontuacao(mediana_final)

        # Limpa marcações anteriores na faixa de colunas para a capacidade
        limpar_linha_faixa(sheet, bloco["linha_capacidade"])
        escrever_celula_com_retry(sheet, bloco["linha_capacidade"], coluna, "X")

        print(f"[CAPACIDADE] {bloco['texto_capacidade']} -> mediana={mediana_final} marcada com 'X' na coluna {coluna}")


def limpar_marcacoes_invalidas(sheet):
    """
    Limpa marcações residuais incorretas da planilha em lote (para ser 1000x mais rápido).
    """
    linha_final = obter_linha_final(sheet)
    if linha_final < LINHA_INICIAL_BUSCA:
        return

    # Lê todas as colunas de B a AA (2 a 27) de uma vez usando A1-notation
    valores = sheet.Range(f"B{LINHA_INICIAL_BUSCA}:AA{linha_final}").Value
    if not valores:
        return

    for i, linha_valores in enumerate(valores):
        linha_real = i + LINHA_INICIAL_BUSCA

        # Mapeamento do slice retornado: 
        # 0->B(2), 1->C(3), 2->D(4), 3->E(5), 4->F(6)
        val_b = linha_valores[0]
        val_c = linha_valores[1]
        val_d = linha_valores[2]
        val_e = linha_valores[3]
        val_f = linha_valores[4]

        tem_bcd = (val_b is not None and str(val_b).strip() != "") or \
                  (val_c is not None and str(val_c).strip() != "") or \
                  (val_d is not None and str(val_d).strip() != "")

        tem_ef = (val_e is not None and str(val_e).strip() != "") or \
                 (val_f is not None and str(val_f).strip() != "")

        if tem_ef:
            # É Critério: limpa 'X's das colunas de nota (colunas 8 a 27, ou seja, index 6 a 25 no slice valores)
            for col_idx in range(6, 26):
                val = linha_valores[col_idx]
                if val == "X" or val == "x":
                    col_real = col_idx + 2
                    sheet.Cells(linha_real, col_real).Value = ""
        elif tem_bcd:
            # É Capacidade: limpa '✓'s das colunas de nota (colunas 8 a 27, ou seja, index 6 a 25 no slice valores)
            for col_idx in range(6, 26):
                val = linha_valores[col_idx]
                if val == "✓":
                    col_real = col_idx + 2
                    sheet.Cells(linha_real, col_real).Value = ""


# =========================
# FUNÇÃO PRINCIPAL / ORQUESTRADORA
# =========================

def preencher_planilha_excel_aberta(debug_nomes=True, solicitar_subnivel_callback=None, should_stop_callback=None):
    """
    Agora percorre todos os alunos do rubricas.json e tenta localizar,
    em todas as abas da workbook, onde está o nome do aluno na coluna 13.
    """
    rubricas_alunos = carregar_rubricas("rubricas.json")

    excel = conectar_excel_aberto()
    workbook = excel.ActiveWorkbook

    if workbook is None:
        raise RuntimeError("Nenhuma planilha aberta.")

    respostas_subnivel = perguntar_subniveis_por_aluno(
        rubricas_alunos,
        solicitar_subnivel_callback=solicitar_subnivel_callback,
        should_stop_callback=should_stop_callback
    )

    total_alunos = 0
    alunos_localizados = 0
    total_preenchidos = 0
    total_nao_encontrados = []

    for aluno_data in rubricas_alunos:
        if should_stop_callback is not None and should_stop_callback():
            print("Processo interrompido pelo usuário durante o preenchimento.")
            break

        nome_aluno = aluno_data.get("aluno", "").strip()
        rubrica = aluno_data.get("rubrica", [])

        if not nome_aluno:
            print("\n[IGNORADO] Registro sem nome de aluno.")
            continue

        total_alunos += 1
        print("\n" + "=" * 70)
        print(f"Processando aluno: {nome_aluno}")

        if not rubrica:
            print(f"[SEM RUBRICA] {nome_aluno}")
            continue

        sheet = encontrar_planilha_do_aluno(
            workbook,
            nome_aluno,
            debug=debug_nomes
        )

        if sheet is None:
            print(f"[ALUNO NÃO LOCALIZADO] {nome_aluno} não foi encontrado em nenhuma aba.")
            continue

        alunos_localizados += 1

        print(f"[ABA SELECIONADA] {sheet.Name}")
        print(f"[NOME LOCALIZADO EM] linha {LINHA_NOME_ALUNO}, coluna {COLUNA_NOME_ALUNO}")

        # Limpa marcações incorretas de execuções passadas
        # limpar_marcacoes_invalidas(sheet)

        # Mapeia as capacidades e estruturas da planilha
        criterios_linhas, capacidades_blocos = mapear_estrutura_planilha(sheet)

        respostas_aluno = respostas_subnivel.get(nome_aluno, {})

        criterios_preenchidos, preenchidos, nao_encontrados = preencher_criterios(
            sheet,
            rubrica,
            respostas_aluno,
            capacidades_blocos
        )

        preencher_capacidades(sheet, capacidades_blocos)

        total_preenchidos += preenchidos
        total_nao_encontrados.extend([f"{nome_aluno}: {c}" for c in nao_encontrados])

    print("\n" + "=" * 70)
    print("RESUMO FINAL")
    print(f"Total de alunos no JSON: {total_alunos}")
    print(f"Alunos localizados na workbook: {alunos_localizados}")
    print(f"Total de critérios preenchidos: {total_preenchidos}")
    print(f"Total de critérios não encontrados: {len(total_nao_encontrados)}")

    if total_nao_encontrados:
        print("\nCritérios não encontrados:")
        for item in total_nao_encontrados:
            print(f" - {item}")

        total_preenchidos += preenchidos
        total_nao_encontrados.extend([f"{nome_aluno}: {c}" for c in nao_encontrados])

    print("\n" + "=" * 70)
    print("RESUMO FINAL")
    print(f"Total de alunos no JSON: {total_alunos}")
    print(f"Alunos localizados na workbook: {alunos_localizados}")
    print(f"Total de critérios preenchidos: {total_preenchidos}")
    print(f"Total de critérios não encontrados: {len(total_nao_encontrados)}")

    if total_nao_encontrados:
        print("\nCritérios não encontrados:")
        for item in total_nao_encontrados:
            print(f" - {item}")


if __name__ == "__main__":
    preencher_planilha_excel_aberta(debug_nomes=True)