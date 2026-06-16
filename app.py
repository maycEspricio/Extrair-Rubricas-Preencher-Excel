import queue
import threading
import json
import logging
from contextlib import redirect_stderr, redirect_stdout
from flask import Flask, jsonify, request
from flask_cors import CORS
import pythoncom

from extrair_rubrica import extrair_rubricas_todos_alunos
from preencher_planilha import (
    carregar_rubricas,
    preencher_planilha_excel_aberta,
    obter_linha_final,
    limpar_faixa,
    COLUNA_CRITERIO,
    LINHA_INICIAL_BUSCA,
    COL_INICIO_AVALIACAO,
    COL_FIM_AVALIACAO,
    encontrar_linha_por_criterio,
    identificar_blocos_capacidade,
    arredondar_media_para_int,
    median,
    encontrar_planilha_do_aluno,
    conectar_excel_aberto,
    NIVEIS_ORDENADOS,
    GRAUS_VALIDOS,
    extrair_numero_grau,
)

app = Flask(__name__)
CORS(app)
log_queue = queue.Queue()

# Variáveis globais de controle
is_processing = False
stop_requested = False
confirm_event = threading.Event()

class LogRedirector:
    def write(self, texto):
        if texto:
            log_queue.put(texto)

    def flush(self):
        pass

def should_stop():
    return stop_requested

@app.route('/api/status', methods=['GET'])
def status():
    return jsonify({"processando": is_processing})

@app.route('/api/logs', methods=['GET'])
def get_logs():
    logs = []
    while not log_queue.empty():
        logs.append(log_queue.get_nowait())
    return jsonify({"logs": "".join(logs)})

@app.route('/api/stop', methods=['POST'])
def stop():
    global stop_requested
    if is_processing:
        stop_requested = True
        log_queue.put("\n[AVISO] Parada solicitada...\n")
    return jsonify({"status": "ok"})

@app.route('/api/extrair', methods=['POST'])
def start_extract():
    global is_processing, stop_requested
    if is_processing:
        return jsonify({"error": "Já existe um processo em execução"}), 400

    is_processing = True
    stop_requested = False
    confirm_event.clear()

    def confirmar():
        log_queue.put("\\n[SISTEMA] Navegador aberto. Acesse a atividade com a tabela de alunos e clique em 'Confirmar Página' na interface...\\n")
        confirm_event.wait()
        return not stop_requested

    def runner():
        global is_processing
        redirecionador = LogRedirector()
        try:
            with redirect_stdout(redirecionador), redirect_stderr(redirecionador):
                extrair_rubricas_todos_alunos(
                    confirmar_inicio_callback=confirmar,
                    should_stop_callback=should_stop
                )
        except Exception as e:
            log_queue.put(f"\\n[ERRO] {str(e)}\\n")
        finally:
            is_processing = False
            confirm_event.set() # Unblock if stuck

    threading.Thread(target=runner, daemon=True).start()
    return jsonify({"status": "started"})

@app.route('/api/extrair/confirmar', methods=['POST'])
def confirm_extract():
    confirm_event.set()
    return jsonify({"status": "ok"})

@app.route('/api/preencher/check', methods=['GET'])
def check_preencher():
    """
    Retorna os alunos que possuem pelo menos um critério com grau válido (G1–G5).
    Estes são os alunos para os quais o professor precisa informar a autonomia.
    Alunos com todos os critérios como 'Nenhum nível marcado' são silenciosamente ignorados.
    """
    try:
        rubricas = carregar_rubricas("rubricas.json")
    except Exception as e:
        return jsonify({"error": "Arquivo rubricas.json não encontrado ou inválido"}), 400

    alunos_pendentes = []
    for aluno_data in rubricas:
        nome_aluno = aluno_data.get("aluno", "").strip()
        rubrica = aluno_data.get("rubrica", [])

        if not nome_aluno or not rubrica:
            continue

        # Verifica se ao menos um critério possui grau válido (G1–G5)
        tem_grau_valido = any(
            extrair_numero_grau(item.get("nivel", "")) is not None
            for item in rubrica
        )

        if tem_grau_valido:
            alunos_pendentes.append({"nome": nome_aluno})

    return jsonify({"pendentes": alunos_pendentes})

@app.route('/api/preencher/execute', methods=['POST'])
def execute_preencher():
    global is_processing, stop_requested
    if is_processing:
        return jsonify({"error": "Já existe um processo em execução"}), 400

    payload = request.json
    respostas_autonomia = payload.get("autonomias", {}) # dict: nome -> nivel (Autônomo, Apoiado, etc)

    is_processing = True
    stop_requested = False

    def runner():
        global is_processing
        pythoncom.CoInitialize() # Necessário para rodar win32com em uma thread paralela
        redirecionador = LogRedirector()
        try:
            with redirect_stdout(redirecionador), redirect_stderr(redirecionador):
                # O script original preencher_planilha_excel_aberta() usa input().
                # Iremos chamar a nossa versão modificada (ou injetar) que usa as respostas.
                from update_planilha import rodar_preenchimento_com_respostas
                rodar_preenchimento_com_respostas(respostas_autonomia, should_stop)
        except Exception as e:
            log_queue.put(f"\\n[ERRO] {str(e)}\\n")
        finally:
            pythoncom.CoUninitialize()
            is_processing = False

    threading.Thread(target=runner, daemon=True).start()
    return jsonify({"status": "started"})

if __name__ == '__main__':
    logging.getLogger('werkzeug').disabled = True
    app.run(port=5000)
