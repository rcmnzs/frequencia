import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext
import os
import threading

try:
    from modulos.logica import processar_dados_diarios, gerar_relatorio_faltas, gerar_relatorio_simples
except ImportError:
    print("ERRO: Certifique-se de que a pasta 'modulos' e o arquivo 'logica.py' existem.")
    exit()

class App:
    def __init__(self, root, db_alunos_status, db_alunos_count, db_horarios_status, db_horarios_count):
        self.root = root
        self.root.title("Sistema de Apuração de Faltas")
        self.root.geometry("800x600")

        self.ausentes_pdf_path = ""
        self.frequencia_pdf_path = ""
        
        # --- CORREÇÃO DEFINITIVA AQUI ---
        # A linha que inicializa o dicionário para acumular dados da sessão.
        # Esta linha estava faltando.
        self.dados_processados_da_sessao = {}
        # --- FIM DA CORREÇÃO ---

        self.pdf_dir = os.path.join(os.path.dirname(__file__), 'pdf')
        if not os.path.exists(self.pdf_dir):
            os.makedirs(self.pdf_dir)

        self._create_widgets()
        self._update_status_bar(db_alunos_status, db_alunos_count, db_horarios_status, db_horarios_count)

    def _create_widgets(self):
        attach_frame = tk.Frame(self.root, padx=10, pady=10)
        attach_frame.pack(fill=tk.X, side=tk.TOP)
        btn_ausentes = tk.Button(attach_frame, text="Anexar PDF de Ausentes", command=self._select_ausentes_pdf)
        btn_ausentes.pack(side=tk.LEFT, padx=5)
        self.lbl_ausentes = tk.Label(attach_frame, text="Nenhum arquivo", fg="grey")
        self.lbl_ausentes.pack(side=tk.LEFT, fill=tk.X, expand=True)
        btn_frequencia = tk.Button(attach_frame, text="Anexar PDF de Frequência", command=self._select_frequencia_pdf)
        btn_frequencia.pack(side=tk.LEFT, padx=(20, 5))
        self.lbl_frequencia = tk.Label(attach_frame, text="Nenhum arquivo", fg="grey")
        self.lbl_frequencia.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        action_frame = tk.Frame(self.root)
        action_frame.pack(pady=10)
        
        self.btn_processar = tk.Button(action_frame, text="1. Processar Dados do Dia", font=('Helvetica', 10, 'bold'), bg="#008CBA", fg="white", command=self._iniciar_processamento)
        self.btn_processar.pack(side=tk.LEFT, padx=10, ipadx=10, ipady=5)
        self.btn_gerar_detalhado = tk.Button(action_frame, text="2. Gerar Relatório Detalhado", state="disabled", command=self._gerar_relatorio_faltas)
        self.btn_gerar_detalhado.pack(side=tk.LEFT, padx=10)
        self.btn_gerar_simples = tk.Button(action_frame, text="3. Gerar Relatório Simples", state="disabled", command=self._gerar_relatorio_simples)
        self.btn_gerar_simples.pack(side=tk.LEFT, padx=10)
        
        status_frame = tk.Frame(self.root, bd=1, relief=tk.SUNKEN)
        status_frame.pack(side=tk.BOTTOM, fill=tk.X)
        self.lbl_db_alunos_status = tk.Label(status_frame, text="BD Alunos: -", anchor='w')
        self.lbl_db_alunos_status.pack(side=tk.LEFT, padx=5)
        self.lbl_db_horarios_status = tk.Label(status_frame, text="BD Horários: -", anchor='w')
        self.lbl_db_horarios_status.pack(side=tk.LEFT, padx=5)
        
        notebook = ttk.Notebook(self.root)
        console_frame = tk.Frame(notebook)
        notebook.add(console_frame, text='Console de Processamento')
        notebook.pack(expand=True, fill='both', padx=10, pady=(0, 10))
        self.console = scrolledtext.ScrolledText(console_frame, wrap=tk.WORD, state='disabled', font=("Courier New", 9))
        self.console.pack(expand=True, fill='both')

    def _iniciar_processamento(self):
        self.console.config(state='normal')
        self.console.delete('1.0', tk.END)
        self.console.config(state='disabled')
        self.btn_processar.config(state="disabled", text="Processando...")
        self.btn_gerar_detalhado.config(state="disabled")
        self.btn_gerar_simples.config(state="disabled")
        thread = threading.Thread(target=self._executar_processamento_em_thread)
        thread.start()

    def _executar_processamento_em_thread(self):
        if not self.ausentes_pdf_path or not self.frequencia_pdf_path:
            self._write_to_console("ERRO: Anexe os dois arquivos PDF antes de processar.")
            self.root.after(0, self._finalizar_processamento, False)
            return
        try:
            self._write_to_console("--- INICIANDO PROCESSAMENTO DOS DADOS (em background) ---")
            dados_do_dia = processar_dados_diarios(
                ausentes_path=self.ausentes_pdf_path,
                frequencia_path=self.frequencia_pdf_path,
                logger=self._write_to_console
            )
            sucesso = dados_do_dia and dados_do_dia[0] is not None
            if sucesso:
                report_date, _, _ = dados_do_dia
                sheet_name = report_date.strftime('%d-%m-%Y')
                self.dados_processados_da_sessao[sheet_name] = dados_do_dia
                self._write_to_console(f"\n--- PROCESSAMENTO DO DIA {sheet_name} CONCLUÍDO ---")
                self._write_to_console("Dados processados com sucesso. Agora você pode gerar os relatórios.")
            else:
                self._write_to_console("\n--- FALHA NO PROCESSAMENTO ---")
                self._write_to_console("Verifique as mensagens de erro acima.")
            self.root.after(0, self._finalizar_processamento, sucesso)
        except Exception as e:
            self._write_to_console(f"\nOcorreu um erro crítico durante o processamento:\n{e}")
            self.root.after(0, self._finalizar_processamento, False)
            
    def _finalizar_processamento(self, sucesso):
        self.btn_processar.config(state="normal", text="1. Processar Dados do Dia")
        if sucesso:
            self.btn_gerar_detalhado.config(state="normal")
            self.btn_gerar_simples.config(state="normal")

    def _update_status_bar(self, alunos_status, alunos_count, horarios_status, horarios_count):
        if alunos_status == "success":
            self.lbl_db_alunos_status.config(text=f"BD Alunos: Conectado ({alunos_count} registros)", fg="green")
        else:
            self.lbl_db_alunos_status.config(text="BD Alunos: Erro de Conexão", fg="red")
        if horarios_status == "success":
            self.lbl_db_horarios_status.config(text=f"BD Horários: Conectado ({horarios_count} registros)", fg="green")
        else:
            self.lbl_db_horarios_status.config(text="BD Horários: Erro de Conexão", fg="red")

    def _select_ausentes_pdf(self):
        filepath = filedialog.askopenfilename(title="Selecione o PDF de Ausentes", initialdir=self.pdf_dir, filetypes=[("Arquivos PDF", "*.pdf")])
        if filepath:
            self.ausentes_pdf_path = filepath
            self.lbl_ausentes.config(text=os.path.basename(filepath), fg="black")

    def _select_frequencia_pdf(self):
        filepath = filedialog.askopenfilename(title="Selecione o PDF de Frequência", initialdir=self.pdf_dir, filetypes=[("Arquivos PDF", "*.pdf")])
        if filepath:
            self.frequencia_pdf_path = filepath
            self.lbl_frequencia.config(text=os.path.basename(filepath), fg="black")
            
    def _write_to_console(self, message):
        self.console.config(state='normal')
        self.console.insert(tk.END, message + "\n")
        self.console.see(tk.END)
        self.console.config(state='disabled')
        self.root.update_idletasks()
        
    def _gerar_relatorio_faltas(self):
        if not self.dados_processados_da_sessao:
            self._write_to_console("ERRO: Nenhum dado foi processado ainda (Botão 1).")
            return
        gerar_relatorio_faltas(self.dados_processados_da_sessao, self._write_to_console)

    def _gerar_relatorio_simples(self):
        if not self.dados_processados_da_sessao:
            self._write_to_console("ERRO: Nenhum dado foi processado ainda (Botão 1).")
            return
        ultimo_dia_processado = list(self.dados_processados_da_sessao.values())[-1]
        report_date, _, df_problemas = ultimo_dia_processado
        gerar_relatorio_simples(df_problemas, report_date, self._write_to_console)