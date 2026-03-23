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


# Ordem lógica dos níveis de desempenho.
# Essa ordem é importante porque o sistema pergunta os subníveis na sequência esperada
# e também porque a lógica de pontuação depende dessa organização conceitual.
NIVEIS_ORDENADOS = [
    "não satisfatório",
    "apoiado",
    "parcialmente autônomo",
    "autônomo",
]


# =========================
# FUNÇÕES AUXILIARES DE TEXTO E LEITURA
# =========================

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


def carregar_rubrica(nome_arquivo="rubrica.json"):
    """
    Carrega do disco o arquivo JSON contendo os critérios e níveis extraídos.

    Esse arquivo deve ter sido gerado anteriormente pela etapa de extração
    da rubrica no Google Classroom.

    Parâmetros:
        nome_arquivo (str): nome do arquivo JSON a ser carregado.

    Retorno:
        list[dict]: estrutura com os dados da rubrica.

    Exemplo de item esperado:
        {
            "criterion_id": "...",
            "criterio": "Texto do critério",
            "nivel": "Apoiado"
        }

    Observação:
    Se o arquivo não existir ou estiver corrompido, esta função lançará erro,
    porque o fluxo depende totalmente desse arquivo para continuar.
    """
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
    """
    Retorna o primeiro valor da faixa correspondente a um nível principal.

    Exemplo:
        "não satisfatório"      -> 1
        "apoiado"               -> 6
        "parcialmente autônomo" -> 11
        "autônomo"              -> 16

    Etapas:
    1. normaliza o texto recebido
    2. procura no mapa de níveis
    3. retorna o início da faixa

    Parâmetros:
        nivel_texto (str): texto do nível principal.

    Retorno:
        int: início da faixa de pontuação.

    Erro:
        ValueError se o texto não corresponder a um nível conhecido.
    """

    # Normaliza para evitar erro por caixa alta/baixa ou espaços.
    nivel = normalizar(nivel_texto)

    # Mapeamento fixo entre o nome do nível e o início da sua faixa.
    mapa = {
        "não satisfatório": 1,
        "apoiado": 6,
        "parcialmente autônomo": 11,
        "autônomo": 16,
    }

    # Proteção contra entradas inesperadas.
    # Isso evita seguir adiante com uma pontuação inválida.
    if nivel not in mapa:
        raise ValueError(f"Nível desconhecido: {nivel_texto}")

    return mapa[nivel]


def obter_pontuacao_interna(nivel_texto: str, nivel_1_a_5: int) -> int:
    """
    Combina o nível principal com o subnível para gerar
    uma pontuação interna final entre 1 e 20.

    Regra:
        pontuação = início_da_faixa + (subnível - 1)

    Exemplos:
        "não satisfatório" + 1 -> 1
        "não satisfatório" + 5 -> 5
        "apoiado" + 1          -> 6
        "apoiado" + 3          -> 8
        "autônomo" + 5         -> 20

    Parâmetros:
        nivel_texto (str): nome do nível principal.
        nivel_1_a_5 (int): refinamento interno do nível, de 1 a 5.

    Retorno:
        int: pontuação final entre 1 e 20.

    Erro:
        ValueError se o subnível estiver fora da faixa permitida.
    """

    # Garante que o subnível informado esteja dentro do intervalo aceito.
    if nivel_1_a_5 < 1 or nivel_1_a_5 > 5:
        raise ValueError("O nível informado deve estar entre 1 e 5.")

    # Soma o deslocamento dentro da faixa.
    return obter_inicio_faixa(nivel_texto) + (nivel_1_a_5 - 1)


def obter_coluna_excel_por_pontuacao(pontuacao_interna: int) -> int:
    """
    Converte a pontuação interna (1 a 20) na coluna correspondente do Excel.

    Regra:
    - a pontuação 1 fica na coluna 8
    - a pontuação 2 fica na coluna 9
    - ...
    - a pontuação 20 fica na coluna 27

    Fórmula:
        coluna_excel = pontuação + 7

    Parâmetros:
        pontuacao_interna (int): valor entre 1 e 20.

    Retorno:
        int: número da coluna no Excel.

    Erro:
        ValueError se a pontuação estiver fora do intervalo válido.
    """
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
    """
    Procura na planilha a linha correspondente a um critério textual.

    Estratégia adotada:
    - normaliza o texto do critério procurado
    - percorre as linhas da planilha
    - lê o valor da coluna onde estão os critérios
    - normaliza o valor lido
    - verifica se o critério procurado está contido no texto da célula

    Importante:
    A comparação usa "contém" em vez de igualdade exata:
        if criterio_norm in texto

    Vantagem:
    - tolera pequenas diferenças de formatação

    Risco:
    - pode encontrar falso positivo se houver textos muito parecidos

    Parâmetros:
        sheet: planilha do Excel.
        criterio (str): texto do critério a localizar.
        coluna_criterio (int): coluna onde o critério será procurado.
        linha_inicial (int): primeira linha a ser considerada.
        linha_final (int | None): última linha a ser considerada.

    Retorno:
        int | None:
            - número da linha encontrada
            - None se não localizar

    Observação:
    Quando linha_final não é informada, ela é calculada com base no UsedRange.
    """

    # Padroniza o critério que veio do JSON para permitir comparação mais robusta.
    criterio_norm = normalizar(criterio)

    # Se o limite superior não for informado, usa a última linha "usada" da planilha.
    if linha_final is None:
        linha_final = sheet.UsedRange.Rows.Count + sheet.UsedRange.Row - 1

    # Percorre cada linha da faixa definida.
    for linha in range(linha_inicial, linha_final + 1):
        # Lê o valor da célula onde supostamente há texto de critério/capacidade.
        valor = sheet.Cells(linha, coluna_criterio).Value

        # Normaliza o valor lido do Excel.
        texto = normalizar(valor)

        # Só considera válido se houver critério a buscar
        # e se esse critério estiver contido no texto da célula.
        if criterio_norm and criterio_norm in texto:
            return linha

    # Se não encontrou nada, retorna None explicitamente.
    return None


def limpar_faixa(sheet, linha, col_inicio=COL_INICIO_AVALIACAO, col_fim=COL_FIM_AVALIACAO):
    """
    Limpa toda a faixa de avaliação de uma linha antes de escrever uma nova marcação.

    Isso é importante para garantir que:
    - não fiquem marcações antigas
    - haja apenas uma coluna marcada por vez
    - o resultado final não fique ambíguo

    Exemplo de uso:
    - antes de colocar "✓" em um critério
    - antes de colocar "X" em uma capacidade

    Parâmetros:
        sheet: planilha do Excel.
        linha (int): linha a ser limpa.
        col_inicio (int): primeira coluna da faixa.
        col_fim (int): última coluna da faixa.
    """
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

def preencher_criterios(sheet, dados, respostas_subnivel):
    """
    Preenche os critérios individualmente na planilha.

    Fluxo geral de cada item:
    1. lê o critério e o nível vindo do JSON
    2. verifica se existe subnível informado para aquele nível
    3. calcula a pontuação interna (1 a 20)
    4. converte essa pontuação na coluna correspondente do Excel
    5. procura a linha do critério na planilha
    6. limpa a faixa de avaliação daquela linha
    7. grava "✓" na coluna correta

    Além disso, a função gera três saídas:
    - lista dos critérios preenchidos com seus metadados
    - quantidade de critérios preenchidos com sucesso
    - lista dos critérios não encontrados

    Parâmetros:
        sheet: planilha do Excel.
        dados (list[dict]): dados carregados do JSON.
        respostas_subnivel (dict): subníveis informados pelo usuário.

    Retorno:
        tuple:
            (
                criterios_preenchidos,
                preenchidos,
                nao_encontrados
            )
    """

    # Contador de sucessos.
    preenchidos = 0

    # Lista de critérios que não foram localizados na planilha.
    nao_encontrados = []

    # Estrutura detalhada dos critérios preenchidos.
    # Será usada depois para calcular as capacidades.
    criterios_preenchidos = []

    # Calcula uma vez só o limite final da busca.
    linha_final = obter_linha_final(sheet)

    # Percorre cada item vindo do JSON.
    for item in dados:
        criterio = item["criterio"]
        nivel = normalizar(item["nivel"])

        # Se não houver subnível correspondente para o nível atual,
        # o critério não pode ser convertido em pontuação.
        if nivel not in respostas_subnivel:
            print(f"[IGNORADO] {criterio} -> nível sem resposta: {item['nivel']}")
            continue

        try:
            # Calcula pontuação de 1 a 20.
            pontuacao_interna = obter_pontuacao_interna(nivel, respostas_subnivel[nivel])

            # Converte pontuação para coluna do Excel.
            coluna_destino = obter_coluna_excel_por_pontuacao(pontuacao_interna)

        except ValueError as e:
            # Se houver problema na pontuação, ignora o critério e segue.
            print(f"[IGNORADO] {criterio} -> {e}")
            continue

        # Localiza a linha do critério na planilha.
        linha = encontrar_linha_por_criterio(
            sheet,
            criterio,
            coluna_criterio=COLUNA_CRITERIO,
            linha_inicial=LINHA_INICIAL_BUSCA,
            linha_final=linha_final
        )

        # Se não encontrou, registra para relatório final.
        if linha is None:
            nao_encontrados.append(criterio)
            print(f"[NÃO ENCONTRADO] {criterio}")
            continue

        # Antes de escrever a nova marcação, limpa a faixa da linha.
        limpar_faixa(sheet, linha)

        # Escreve o símbolo de marcação do critério.
        sheet.Cells(linha, coluna_destino).Value = "✓"

        # Guarda metadados úteis para rastrear o preenchimento
        # e para calcular capacidades depois.
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
    """
    Identifica os blocos de capacidade na planilha com base nas linhas
    dos critérios que foram preenchidos.

    Regra de negócio adotada:
    - a coluna 7 contém tanto linhas de capacidade quanto linhas de critério
    - toda linha não vazia na coluna 7 que NÃO seja um critério preenchido
      será tratada como uma linha de capacidade
    - os critérios encontrados pertencem à última capacidade válida acima deles

    Em outras palavras:
    o algoritmo percorre a planilha de cima para baixo.
    Quando encontra uma linha não vazia que não é critério, assume que começou
    uma nova capacidade. Os critérios encontrados depois dela passam a pertencer
    a essa capacidade, até surgir uma nova capacidade.

    Parâmetros:
        sheet: planilha do Excel.
        criterios_preenchidos (list[dict]): critérios que já foram localizados e marcados.

    Retorno:
        list[dict]: lista de blocos no formato:
            {
                "linha_capacidade": 30,
                "texto_capacidade": "Capacidade X",
                "criterios": [31, 32, 33]
            }

    Observação importante:
    Esta lógica depende bastante da organização visual da planilha.
    Se houver linhas intermediárias com texto que não sejam capacidades nem critérios,
    elas podem ser interpretadas incorretamente como capacidade.
    """

    # Se nenhum critério foi preenchido, não há como montar blocos.
    if not criterios_preenchidos:
        return []

    # Extrai as linhas dos critérios preenchidos e ordena.
    linhas_criterio = sorted(item["linha"] for item in criterios_preenchidos)

    # Conjunto para busca rápida "linha é critério?".
    conjunto_linhas_criterio = set(linhas_criterio)

    linha_final = obter_linha_final(sheet)

    # Lista final de blocos de capacidade.
    blocos = []

    # Referência para a capacidade "corrente" enquanto percorremos a planilha.
    capacidade_atual = None

    # Percorre toda a área relevante da planilha.
    for linha in range(LINHA_INICIAL_BUSCA, linha_final + 1):
        texto = sheet.Cells(linha, COLUNA_CRITERIO).Value
        texto_norm = normalizar(texto)

        # Ignora linhas vazias, pois elas não definem capacidade nem critério.
        if not texto_norm:
            continue

        # Se a linha atual é uma linha de critério preenchido,
        # então ela pertence à última capacidade encontrada anteriormente.
        if linha in conjunto_linhas_criterio:
            if capacidade_atual:
                capacidade_atual["criterios"].append(linha)
            continue

        # Se chegou aqui, a linha tem texto e não é critério preenchido.
        # Pela regra adotada, isso é tratado como uma nova capacidade.
        capacidade_atual = {
            "linha_capacidade": linha,
            "texto_capacidade": str(texto).strip(),
            "criterios": []
        }
        blocos.append(capacidade_atual)

    # Remove capacidades que não receberam nenhum critério associado.
    # Isso evita calcular médias vazias depois.
    return [b for b in blocos if b["criterios"]]


def arredondar_media_para_int(mediana):
    """
    Aplica arredondamento tradicional em uma mediana numérica.

    Estratégia usada:
        floor(mediana + 0.5)

    Exemplos:
        7.2 -> 7
        7.5 -> 8
        7.8 -> 8

    Isso foi escolhido para evitar o comportamento padrão de algumas abordagens
    de arredondamento que podem tratar .5 de forma diferente.

    Parâmetros:
        mediana (float): mediana calculada.

    Retorno:
        int: mediana arredondada.
    """
    return int(math.floor(mediana + 0.5))


def preencher_capacidades(sheet, criterios_preenchidos):
    """
    Calcula e preenche as capacidades com base na mediana dos critérios associados.

    Fluxo:
    1. monta um mapa {linha_do_criterio: pontuação}
    2. identifica os blocos de capacidade
    3. para cada bloco, coleta as pontuações dos critérios associados
    4. calcula a mediana
    5. arredonda a mediana
    6. converte a pontuação final em coluna do Excel
    7. limpa a faixa da linha da capacidade
    8. escreve "X" na coluna correspondente

    Diferença conceitual:
    - critérios recebem "✓"
    - capacidades recebem "X"

    Parâmetros:
        sheet: planilha do Excel.
        criterios_preenchidos (list[dict]): saída de preencher_criterios().

    Observação:
    Esta função depende diretamente da qualidade da associação feita por
    identificar_blocos_capacidade().
    """

    # Sem critérios preenchidos, não há o que calcular.
    if not criterios_preenchidos:
        print("\nNenhum critério preenchido.")
        return

    # Cria mapa para acesso rápido da pontuação pela linha.
    # Exemplo: {25: 8, 26: 10, 27: 14}
    mapa = {item["linha"]: item["pontuacao"] for item in criterios_preenchidos}

    # Descobre quais capacidades existem e quais critérios pertencem a cada uma.
    blocos = identificar_blocos_capacidade(sheet, criterios_preenchidos)

    print("\nCalculando capacidades...\n")

    # Processa cada capacidade individualmente.
    for bloco in blocos:

        # Coleta as pontuações dos critérios pertencentes a este bloco.
        pontuacoes = [mapa[l] for l in bloco["criterios"] if l in mapa]

        # Se, por algum motivo, não houver pontuações, pula o bloco.
        if not pontuacoes:
            continue

        # Calcula mediana aritmética simples.
        mediana = median(pontuacoes)

        if mediana <= 10:
            mediana = 11

        # Arredonda para inteiro.
        mediana_final = arredondar_media_para_int(mediana)

        # Descobre em qual coluna essa mediana deve ser marcada.
        coluna = obter_coluna_excel_por_pontuacao(mediana_final)

        # Limpa a faixa da linha da capacidade antes de marcar.
        limpar_faixa(sheet, bloco["linha_capacidade"])

        # Escreve a marcação da capacidade.
        sheet.Cells(bloco["linha_capacidade"], coluna).Value = "X"

        print(f"[CAPACIDADE] {bloco['texto_capacidade']} -> mediana={mediana_final}")


# =========================
# FUNÇÃO PRINCIPAL / ORQUESTRADORA
# =========================

def preencher_planilha_excel_aberta(nome_aba=None):
    """
    Função principal responsável por coordenar todo o preenchimento da planilha.

    Etapas executadas:
    1. carrega o arquivo rubrica.json
    2. conecta ao Excel que já está aberto
    3. obtém o workbook ativo
    4. escolhe a planilha correta
    5. pergunta ao usuário os subníveis de cada desempenho encontrado
    6. preenche os critérios individualmente
    7. calcula e preenche as capacidades
    8. exibe um resumo final no terminal

    Parâmetros:
        nome_aba (str | None):
            - se informado, usa a aba com esse nome
            - caso contrário, usa a aba ativa

    Erro:
        RuntimeError se não houver workbook ativo no Excel.
    """

    # Carrega os dados extraídos previamente da rubrica.
    dados = carregar_rubrica("rubrica.json")

    # Conecta ao Excel já aberto.
    excel = conectar_excel_aberto()

    # Recupera a pasta de trabalho que estiver ativa no Excel.
    workbook = excel.ActiveWorkbook

    # Se não houver workbook ativo, o processo não pode continuar.
    if workbook is None:
        raise RuntimeError("Nenhuma planilha aberta.")

    # Seleciona a aba que será usada no preenchimento.
    sheet = escolher_planilha(workbook, nome_aba)

    # Solicita ao usuário o refinamento dos níveis principais.
    respostas_subnivel = perguntar_subniveis_por_desempenho(dados)

    # Preenche os critérios e guarda informações úteis para a etapa seguinte.
    criterios_preenchidos, preenchidos, nao_encontrados = preencher_criterios(
        sheet, dados, respostas_subnivel
    )

    # Usa os critérios preenchidos para calcular a mediana por capacidade.
    preencher_capacidades(sheet, criterios_preenchidos)

    # Exibe resumo final do processamento.
    print("\nResumo:")
    print(f"Preenchidos: {preenchidos}")
    print(f"Não encontrados: {len(nao_encontrados)}")


# =========================
# PONTO DE ENTRADA DO SCRIPT
# =========================
#
# Este bloco garante que a função principal só seja executada automaticamente
# quando este arquivo for rodado diretamente.
#
# Se este arquivo for importado por outro módulo, como um main.py,
# o código abaixo NÃO será executado automaticamente.
#

if __name__ == "__main__":
    preencher_planilha_excel_aberta(nome_aba=None)