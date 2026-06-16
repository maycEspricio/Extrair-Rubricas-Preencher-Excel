# =========================
# IMPORTAÇÕES
# =========================

# Biblioteca padrão para leitura do arquivo JSON gerado na etapa anterior.
# Esse JSON contém a rubrica extraída do Google Classroom.
import json

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
# Isso evita procurar em cabeçalhos, títulos, instruções ou outras áreas acima.
LINHA_INICIAL_BUSCA = 21

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

    Exemplo:
        "  Parcialmente   Autônomo  " -> "parcialmente autônomo"

    Isso ajuda porque, sem normalização, textos visualmente iguais
    podem falhar numa comparação simples.
    """
    return " ".join(str(texto or "").split()).strip().lower()


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

def encontrar_linha_por_criterio(sheet, criterio, coluna_criterio=COLUNA_CRITERIO, linha_inicial=1, linha_final=None):
    criterio_norm = normalizar(criterio)

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

def limpar_faixa(sheet, linha, col_inicio=COL_INICIO_AVALIACAO, col_fim=COL_FIM_AVALIACAO):
    for col in range(col_inicio, col_fim + 1):
        sheet.Cells(linha, col).Value = ""

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

def preencher_criterios(sheet, dados, respostas_subnivel_aluno):
    preenchidos = 0
    nao_encontrados = []
    criterios_preenchidos = []

    linha_final = obter_linha_final(sheet)

    for item in dados:
        criterio = item["criterio"]
        nivel = normalizar(item["nivel"])

        if nivel not in respostas_subnivel_aluno:
            print(f"[IGNORADO] {criterio} -> nível sem resposta: {item['nivel']}")
            continue

        try:
            pontuacao_interna = obter_pontuacao_interna(nivel, respostas_subnivel_aluno[nivel])
            coluna_destino = obter_coluna_excel_por_pontuacao(pontuacao_interna)
        except ValueError as e:
            print(f"[IGNORADO] {criterio} -> {e}")
            continue

        linha = encontrar_linha_por_criterio(
            sheet,
            criterio,
            coluna_criterio=COLUNA_CRITERIO,
            linha_inicial=LINHA_INICIAL_BUSCA,
            linha_final=linha_final
        )

        if linha is None:
            nao_encontrados.append(criterio)
            print(f"[NÃO ENCONTRADO] {criterio}")
            continue

        limpar_faixa(sheet, linha)
        sheet.Cells(linha, coluna_destino).Value = "✓"

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


# =========================
# IDENTIFICAÇÃO E PREENCHIMENTO DAS CAPACIDADES
# =========================

def identificar_blocos_capacidade(sheet, criterios_preenchidos):
    if not criterios_preenchidos:
        return []

    linhas_criterio = sorted(item["linha"] for item in criterios_preenchidos)
    conjunto_linhas_criterio = set(linhas_criterio)
    linha_final = obter_linha_final(sheet)

    blocos = []
    capacidade_atual = None

    for linha in range(LINHA_INICIAL_BUSCA, linha_final + 1):
        texto = sheet.Cells(linha, COLUNA_CRITERIO).Value
        texto_norm = normalizar(texto)

        if not texto_norm:
            continue

        if linha in conjunto_linhas_criterio:
            if capacidade_atual:
                capacidade_atual["criterios"].append(linha)
            continue

        capacidade_atual = {
            "linha_capacidade": linha,
            "texto_capacidade": str(texto).strip(),
            "criterios": []
        }
        blocos.append(capacidade_atual)

    return [b for b in blocos if b["criterios"]]


def arredondar_media_para_int(mediana):
    return int(math.floor(mediana + 0.5))


def preencher_capacidades(sheet, criterios_preenchidos):
    if not criterios_preenchidos:
        print("\nNenhum critério preenchido.")
        return

    mapa = {item["linha"]: item["pontuacao"] for item in criterios_preenchidos}
    blocos = identificar_blocos_capacidade(sheet, criterios_preenchidos)

    print("\nCalculando capacidades...\n")

    for bloco in blocos:
        pontuacoes = [mapa[l] for l in bloco["criterios"] if l in mapa]

        if not pontuacoes:
            continue

        mediana = median(pontuacoes)

        if mediana <= 10:
            mediana = 11

        mediana_final = arredondar_media_para_int(mediana)
        coluna = obter_coluna_excel_por_pontuacao(mediana_final)

        limpar_faixa(sheet, bloco["linha_capacidade"])
        sheet.Cells(bloco["linha_capacidade"], coluna).Value = "X"

        print(f"[CAPACIDADE] {bloco['texto_capacidade']} -> mediana={mediana_final}")


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

        respostas_aluno = respostas_subnivel.get(nome_aluno, {})

        criterios_preenchidos, preenchidos, nao_encontrados = preencher_criterios(
            sheet,
            rubrica,
            respostas_aluno
        )

        preencher_capacidades(sheet, criterios_preenchidos)

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