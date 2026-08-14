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
    limpar_marcacoes_invalidas,
    mapear_estrutura_planilha,
    escrever_celula_com_retry,
    limpar_linha_faixa,
)



def preencher_criterios_automatico(sheet, dados, autonomia_fornecida, capacidades_blocos=None):
    """
    Localiza cada critério no Excel e preenche com "✓" na coluna correspondente
    à pontuação calculada.

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

    # Filtra os dados para remover itens de controle do Classroom
    dados_filtrados = []
    for item in dados:
        crit_text = normalizar(item.get("criterio", ""))
        if not crit_text or crit_text == "expandir critério" or "critério não encontrado" in crit_text:
            continue
        dados_filtrados.append(item)

    # Obtém todas as linhas de critérios físicas mapeadas na planilha
    criterios_linhas, _ = mapear_estrutura_planilha(sheet)
    linha_final = obter_linha_final(sheet)

    # Lê toda a coluna G em lote
    valores_col = sheet.Range(f"G1:G{linha_final}").Value
    if valores_col and not isinstance(valores_col, (list, tuple)):
        valores_col = (valores_col,)

    # Passo 1: Identificar candidatas e travar critérios únicos
    itens_resolvidos = []
    for item in dados_filtrados:
        criterio_json = normalizar(item["criterio"])
        candidatas = []
        for r in criterios_linhas:
            if r - 1 < len(valores_col):
                val_planilha = valores_col[r - 1]
                txt_planilha = val_planilha[0] if isinstance(val_planilha, (list, tuple)) else val_planilha
                txt_planilha_norm = normalizar(txt_planilha)
                if criterio_json and (criterio_json in txt_planilha_norm or txt_planilha_norm in criterio_json):
                    candidatas.append(r)
        
        itens_resolvidos.append({
            "item": item,
            "candidatas": candidatas,
            "linha_final": candidatas[0] if len(candidatas) == 1 else None
        })

    # Passo 2: Resolver duplicados usando o vizinho resolvido mais próximo
    for idx, info in enumerate(itens_resolvidos):
        if info["linha_final"] is not None or not info["candidatas"]:
            continue
        
        vizinho_linha = None
        distancia_minima = float('inf')
        
        for i_viz, viz_info in enumerate(itens_resolvidos):
            if i_viz == idx:
                continue
            if viz_info["linha_final"] is not None:
                dist = abs(i_viz - idx)
                if dist < distancia_minima:
                    distancia_minima = dist
                    vizinho_linha = viz_info["linha_final"]
        
        if vizinho_linha is not None:
            info["linha_final"] = min(info["candidatas"], key=lambda r: abs(r - vizinho_linha))
        else:
            info["linha_final"] = info["candidatas"][0]

    # Passo 3: Preencher no Excel
    for info in itens_resolvidos:
        linha = info["linha_final"]
        item = info["item"]
        criterio = item["criterio"]
        nivel_texto = item.get("nivel", "")

        if not linha:
            nao_encontrados.append(criterio)
            print(f"[NÃO ENCONTRADO] {criterio}")
            continue

        # Extrai o número do grau (1–5) a partir do campo nivel
        numero_grau = extrair_numero_grau(nivel_texto)
        low_nivel = normalizar(nivel_texto)
        is_nao_atingiu = "não atingiu" in low_nivel or "nao atingiu" in low_nivel

        if numero_grau is None and not is_nao_atingiu:
            print(f"[IGNORADO] {criterio} na linha {linha} → grau inválido ou ausente: '{nivel_texto}'")
            continue

        try:
            if is_nao_atingiu:
                pontuacao_interna = 1
            else:
                pontuacao_interna = obter_pontuacao_interna(autonomia_norm, numero_grau)
            coluna_destino = obter_coluna_excel_por_pontuacao(pontuacao_interna)
        except ValueError as e:
            print(f"[ERRO PONTUAÇÃO] {criterio} na linha {linha} → {e}")
            continue

        # Limpa qualquer marcação anterior nas colunas de nota do critério
        limpar_linha_faixa(sheet, linha)
        escrever_celula_com_retry(sheet, linha, coluna_destino, "✓")

        criterios_preenchidos.append({
            "linha": linha,
            "criterio": criterio,
            "nivel": autonomia_norm,
            "pontuacao": pontuacao_interna,
            "coluna_excel": coluna_destino,
        })

        preenchidos += 1
        print(f"[OK] {criterio} → autonomia={autonomia_norm}, grau=G{numero_grau}, pontuação={pontuacao_interna} na linha {linha}")

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

        # Limpa marcações incorretas de execuções passadas
        # limpar_marcacoes_invalidas(sheet)

        # Obtém estrutura e capacidades mapeadas
        _, capacidades_blocos = mapear_estrutura_planilha(sheet)

        criterios_preenchidos, preenchidos, nao_encontrados = preencher_criterios_automatico(
            sheet, rubrica, autonomia_fornecida, capacidades_blocos
        )

        preencher_capacidades(sheet, capacidades_blocos)

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
