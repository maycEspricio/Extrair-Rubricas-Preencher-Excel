"""
update_planilha.py
------------------
Contém as funções injetadas no fluxo de preenchimento via API:

  - preencher_criterios_automatico: nova lógica com grau (G1–G5) como
    subnível e autonomia informada pelo professor como faixa (1–20).
  - rodar_preenchimento_com_respostas: orquestrador chamado pelo app.py.

Mapeamento de pontuação:
    Autonomia (faixa)          → pontos base
      Não satisfatório         →  1 a  5
      Apoiado                  →  6 a 10
      Parcialmente autônomo    → 11 a 15
      Autônomo                 → 16 a 20

    Grau (subnível dentro da faixa):
      G1 → +0  (posição 1)
      G2 → +1
      G3 → +2
      G4 → +3
      G5 → +4
"""

from preencher_planilha import (
    carregar_rubricas,
    conectar_excel_aberto,
    encontrar_planilha_do_aluno,
    obter_pontuacao_interna,
    obter_coluna_excel_por_pontuacao,
    encontrar_linha_por_criterio,
    limpar_faixa,
    preencher_capacidades,
    obter_linha_final,
    normalizar,
    NIVEIS_ORDENADOS,
    COLUNA_CRITERIO,
    LINHA_INICIAL_BUSCA,
    extrair_numero_grau,
)


def preencher_criterios_automatico(sheet, dados, autonomia_fornecida):
    """
    Preenche os critérios de um aluno usando:
      - autonomia_fornecida (str): nível informado pelo professor → define a FAIXA (1–20)
      - grau extraído de item["nivel"] (G1–G5) → define o SUBNÍVEL dentro da faixa

    Critérios com "Nenhum nível marcado" (ou grau inválido) são silenciosamente ignorados.

    Retorno:
        (criterios_preenchidos, total_preenchidos, nao_encontrados)
    """
    preenchidos = 0
    nao_encontrados = []
    criterios_preenchidos = []

    if not autonomia_fornecida:
        print("[AVISO] Autonomia não fornecida para este aluno — nenhum critério será preenchido.")
        return criterios_preenchidos, preenchidos, nao_encontrados

    autonomia_norm = normalizar(autonomia_fornecida)
    if autonomia_norm not in NIVEIS_ORDENADOS:
        print(f"[AVISO] Autonomia inválida: '{autonomia_fornecida}' — ignorando aluno.")
        return criterios_preenchidos, preenchidos, nao_encontrados

    linha_final = obter_linha_final(sheet)

    for item in dados:
        criterio = item["criterio"]
        nivel_texto = item.get("nivel", "")

        # Extrai o número do grau (1–5) a partir do campo nivel
        numero_grau = extrair_numero_grau(nivel_texto)

        if numero_grau is None:
            # "Nenhum nível marcado" ou grau não reconhecível → ignora critério
            print(f"[IGNORADO] {criterio} → grau inválido ou ausente: '{nivel_texto}'")
            continue

        try:
            pontuacao_interna = obter_pontuacao_interna(autonomia_norm, numero_grau)
            coluna_destino = obter_coluna_excel_por_pontuacao(pontuacao_interna)
        except ValueError as e:
            print(f"[ERRO PONTUAÇÃO] {criterio} → {e}")
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
            "nivel": autonomia_norm,
            "pontuacao": pontuacao_interna,
            "coluna_excel": coluna_destino,
        })

        preenchidos += 1
        print(f"[OK] {criterio} → autonomia={autonomia_norm}, grau=G{numero_grau}, pontuação={pontuacao_interna}")

    return criterios_preenchidos, preenchidos, nao_encontrados


def rodar_preenchimento_com_respostas(respostas_autonomia, should_stop_callback=None):
    """
    Orquestrador chamado pelo app.py quando o professor confirma as autonomias.

    Parâmetros:
        respostas_autonomia (dict): { "Nome do Aluno": "Autônomo", ... }
        should_stop_callback (callable | None): retorna True se o processo deve parar.
    """
    rubricas_alunos = carregar_rubricas("rubricas.json")
    excel = conectar_excel_aberto()
    workbook = excel.ActiveWorkbook

    if workbook is None:
        raise RuntimeError("Nenhuma planilha aberta no Excel.")

    total_alunos = 0
    alunos_localizados = 0
    total_preenchidos = 0
    total_nao_encontrados = []

    for aluno_data in rubricas_alunos:
        if should_stop_callback is not None and should_stop_callback():
            print("Processo interrompido.")
            break

        nome_aluno = aluno_data.get("aluno", "").strip()
        rubrica = aluno_data.get("rubrica", [])

        if not nome_aluno or not rubrica:
            continue

        # Alunos sem nenhum grau válido não aparecem na lista de pendentes
        # e portanto não terão autonomia fornecida — são silenciosamente ignorados.
        autonomia_fornecida = respostas_autonomia.get(nome_aluno)
        if not autonomia_fornecida:
            print(f"\n[IGNORADO] {nome_aluno} — sem autonomia definida (sem graus válidos).")
            continue

        total_alunos += 1
        print(f"\n===== Processando aluno: {nome_aluno} =====")

        sheet = encontrar_planilha_do_aluno(workbook, nome_aluno, debug=False)
        if sheet is None:
            print(f"[ALUNO NÃO LOCALIZADO] {nome_aluno}")
            continue

        alunos_localizados += 1

        criterios_preenchidos, preenchidos, nao_encontrados = preencher_criterios_automatico(
            sheet, rubrica, autonomia_fornecida
        )

        preencher_capacidades(sheet, criterios_preenchidos)

        total_preenchidos += preenchidos
        total_nao_encontrados.extend([f"{nome_aluno}: {c}" for c in nao_encontrados])

    print("\nResumo Final:")
    print(f"Total de alunos processados: {total_alunos}")
    print(f"Alunos localizados na planilha: {alunos_localizados}")
    print(f"Critérios preenchidos: {total_preenchidos}")
    if total_nao_encontrados:
        print("Critérios não encontrados na planilha:")
        for item in total_nao_encontrados:
            print(f"  - {item}")
