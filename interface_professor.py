"""
Interface gráfica simples para professores (modo leigo).

O objetivo desta tela é reduzir uso de terminal e guiar o processo em passos:
1) Extrair rubricas do Google Classroom;
2) Preencher planilha Excel com os dados extraídos.

Como usar:
    python interface_professor.py
"""

import queue
import threading
import tkinter as tk
from contextlib import redirect_stderr, redirect_stdout
from tkinter import messagebox

from extrair_rubrica import extrair_rubricas_todos_alunos
from preencher_planilha import preencher_planilha_excel_aberta


class _FilaTexto:
    """
    "Arquivo" de saída para capturar prints de funções existentes.

    A ideia é redirecionar stdout/stderr para uma fila, e a interface
    vai consumindo essa fila para mostrar os logs em tempo real.
    """

    def __init__(self, fila_saida):
        self.fila_saida = fila_saida

    def write(self, texto):
        if texto:
            self.fila_saida.put(texto)

    def flush(self):
        # Método exigido por file-like objects.
        return None


class AppProfessor:
    def __init__(self, root):
        self.root = root
        self.root.title("Preenchimento de Rubricas - Interface Professor")
        self.root.geometry("980x650")

        self.fila_logs = queue.Queue()
        self.processando = False
        self.stop_requested = False

        self._montar_layout()
        self._agendar_consumo_logs()

    def _montar_layout(self):
        frame_topo = tk.Frame(self.root, padx=12, pady=12)
        frame_topo.pack(fill="x")

        titulo = tk.Label(
            frame_topo,
            text="Preenchimento de Rubricas",
            font=("Segoe UI", 16, "bold"),
        )
        titulo.pack(anchor="w")

        subtitulo = tk.Label(
            frame_topo,
            text=(
                "Fluxo guiado: extraia do Classroom e depois preencha o Excel.\n"
                "Importante: mantenha o Classroom em primeiro plano na extração e o Excel aberto no preenchimento."
            ),
            justify="left",
            font=("Segoe UI", 10),
        )
        subtitulo.pack(anchor="w", pady=(4, 0))

        frame_botoes = tk.Frame(self.root, padx=12, pady=8)
        frame_botoes.pack(fill="x")

        self.btn_extrair = tk.Button(
            frame_botoes,
            text="1) Extrair Rubricas do Classroom",
            width=34,
            height=2,
            command=self.rodar_extracao,
        )
        self.btn_extrair.pack(side="left", padx=(0, 8))

        self.btn_preencher = tk.Button(
            frame_botoes,
            text="2) Preencher Planilha no Excel",
            width=34,
            height=2,
            command=self.rodar_preenchimento,
        )
        self.btn_preencher.pack(side="left", padx=(0, 8))

        self.btn_fluxo = tk.Button(
            frame_botoes,
            text="Fluxo Completo (1 + 2)",
            width=28,
            height=2,
            command=self.rodar_fluxo_completo,
        )
        self.btn_fluxo.pack(side="left")

        self.btn_parar = tk.Button(
            frame_botoes,
            text="Parar processo",
            width=20,
            height=2,
            state="disabled",
            command=self.solicitar_parada,
        )
        self.btn_parar.pack(side="left", padx=(8, 0))

        frame_logs = tk.Frame(self.root, padx=12, pady=8)
        frame_logs.pack(fill="both", expand=True)

        lbl_logs = tk.Label(frame_logs, text="Logs do processo", font=("Segoe UI", 11, "bold"))
        lbl_logs.pack(anchor="w", pady=(0, 6))

        self.txt_logs = tk.Text(frame_logs, wrap="word", font=("Consolas", 10))
        self.txt_logs.pack(fill="both", expand=True)

        frame_rodape = tk.Frame(self.root, padx=12, pady=8)
        frame_rodape.pack(fill="x")

        self.lbl_status = tk.Label(frame_rodape, text="Status: pronto", anchor="w")
        self.lbl_status.pack(fill="x")

    def _agendar_consumo_logs(self):
        """
        Atualiza a caixa de logs sem travar a interface.
        """
        while not self.fila_logs.empty():
            texto = self.fila_logs.get_nowait()
            self.txt_logs.insert("end", texto)
            self.txt_logs.see("end")

        self.root.after(120, self._agendar_consumo_logs)

    def _set_processando(self, em_execucao: bool, status: str):
        self.processando = em_execucao
        self.btn_extrair.config(state="disabled" if em_execucao else "normal")
        self.btn_preencher.config(state="disabled" if em_execucao else "normal")
        self.btn_fluxo.config(state="disabled" if em_execucao else "normal")
        self.btn_parar.config(state="normal" if em_execucao else "disabled")
        self.lbl_status.config(text=f"Status: {status}")

    def _confirmar_inicio_extracao(self):
        messagebox.showinfo(
            "Etapa de Extração",
            "Abra a atividade do Google Classroom com a tabela de estudantes e deixe essa janela em primeiro plano.\n"
            "Se o Classroom ficar em segundo plano, a extração pode falhar.\n\n"
            "Quando estiver pronta, clique em OK para continuar.",
        )
        return True

    def _should_stop(self):
        return self.stop_requested

    def solicitar_parada(self):
        if not self.processando:
            return
        self.stop_requested = True
        self.fila_logs.put("\n[AVISO] Parada solicitada. Encerrando no próximo ponto seguro...\n")
        self.lbl_status.config(text="Status: encerrando...")

    def _solicitar_subnivel(self, nome_aluno, nivel):
        """
        Pergunta o subnível (1 a 5) via janela com botões de opção (radio).
        """
        while True:
            valor = self._abrir_dialogo_subnivel_radio(nome_aluno, nivel)

            # Se cancelar, interrompe o fluxo com exceção controlada.
            if valor is None:
                raise RuntimeError("Processo cancelado pelo usuário ao informar subníveis.")

            if 1 <= valor <= 5:
                return valor

            messagebox.showwarning("Valor inválido", "Digite um número entre 1 e 5.")

    def _abrir_dialogo_subnivel_radio(self, nome_aluno, nivel):
        """
        Abre um diálogo modal com opções de 1 a 5 em formato radio.

        Retorno:
        - int (1..5) quando usuário confirma;
        - None quando usuário cancela.
        """
        dialog = tk.Toplevel(self.root)
        dialog.title("Subnível da Rubrica")
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.resizable(False, False)

        frame = tk.Frame(dialog, padx=14, pady=12)
        frame.pack(fill="both", expand=True)

        tk.Label(
            frame,
            text=f"Aluno: {nome_aluno}",
            anchor="w",
            justify="left",
            font=("Segoe UI", 10, "bold"),
        ).pack(anchor="w")

        tk.Label(
            frame,
            text=f"Desempenho: {nivel.title()}",
            anchor="w",
            justify="left",
            font=("Segoe UI", 10),
        ).pack(anchor="w", pady=(2, 10))

        tk.Label(
            frame,
            text="Selecione o subnível (1 a 5):",
            anchor="w",
            justify="left",
            font=("Segoe UI", 10),
        ).pack(anchor="w")

        valor_var = tk.IntVar(value=3)
        resultado = {"valor": None}

        frame_radios = tk.Frame(frame, pady=8)
        frame_radios.pack(anchor="w")

        for numero in range(1, 6):
            tk.Radiobutton(
                frame_radios,
                text=str(numero),
                variable=valor_var,
                value=numero,
                font=("Segoe UI", 10),
            ).pack(side="left", padx=(0, 10))

        frame_botoes = tk.Frame(frame, pady=6)
        frame_botoes.pack(fill="x")

        def confirmar():
            resultado["valor"] = int(valor_var.get())
            dialog.destroy()

        def cancelar():
            resultado["valor"] = None
            dialog.destroy()

        tk.Button(frame_botoes, text="Confirmar", width=12, command=confirmar).pack(side="left")
        tk.Button(frame_botoes, text="Cancelar", width=12, command=cancelar).pack(side="left", padx=(8, 0))

        dialog.protocol("WM_DELETE_WINDOW", cancelar)
        self.root.wait_window(dialog)
        return resultado["valor"]

    def _executar_em_thread(self, alvo, status_inicio, status_fim):
        if self.processando:
            messagebox.showwarning("Aguarde", "Já existe um processo em execução.")
            return

        self.stop_requested = False
        self._set_processando(True, status_inicio)

        def _runner():
            redirecionador = _FilaTexto(self.fila_logs)
            try:
                with redirect_stdout(redirecionador), redirect_stderr(redirecionador):
                    alvo()

                if self.stop_requested:
                    self.fila_logs.put("\n[PARADO] Processo interrompido pelo usuário.\n")
                    self.root.after(0, lambda: self._set_processando(False, "interrompido"))
                else:
                    self.fila_logs.put("\n[OK] Processo finalizado com sucesso.\n")
                    self.root.after(0, lambda: self._set_processando(False, status_fim))
            except Exception as exc:
                self.fila_logs.put(f"\n[ERRO] {exc}\n")
                self.root.after(0, lambda: self._set_processando(False, "erro"))
                self.root.after(
                    0,
                    lambda: messagebox.showerror("Erro", f"O processo falhou:\n{exc}")
                )

        threading.Thread(target=_runner, daemon=True).start()

    def rodar_extracao(self):
        def tarefa():
            extrair_rubricas_todos_alunos(
                confirmar_inicio_callback=self._confirmar_inicio_extracao,
                should_stop_callback=self._should_stop,
            )

        self._executar_em_thread(
            tarefa,
            status_inicio="extraindo rubricas...",
            status_fim="extração concluída",
        )

    def rodar_preenchimento(self):
        def tarefa():
            messagebox.showinfo(
                "Etapa de Preenchimento",
                "Abra a planilha no Excel Desktop antes de continuar.\n"
                "A interface vai pedir os subníveis em janelas.",
            )
            preencher_planilha_excel_aberta(
                debug_nomes=False,
                solicitar_subnivel_callback=self._solicitar_subnivel,
                should_stop_callback=self._should_stop,
            )

        self._executar_em_thread(
            tarefa,
            status_inicio="preenchendo planilha...",
            status_fim="preenchimento concluído",
        )

    def rodar_fluxo_completo(self):
        def tarefa():
            extrair_rubricas_todos_alunos(
                confirmar_inicio_callback=self._confirmar_inicio_extracao,
                should_stop_callback=self._should_stop,
            )

            if self._should_stop():
                return

            messagebox.showinfo(
                "Próxima Etapa",
                "Extração concluída.\nAgora abra a planilha no Excel e clique em OK.",
            )

            preencher_planilha_excel_aberta(
                debug_nomes=False,
                solicitar_subnivel_callback=self._solicitar_subnivel,
                should_stop_callback=self._should_stop,
            )

        self._executar_em_thread(
            tarefa,
            status_inicio="executando fluxo completo...",
            status_fim="fluxo completo concluído",
        )


def main():
    root = tk.Tk()
    app = AppProfessor(root)
    root.mainloop()


if __name__ == "__main__":
    main()
