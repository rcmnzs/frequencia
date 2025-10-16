import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext
import os

# Importa as novas funções do módulo de lógica
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
        
        # Variáveis para armazenar os dados processados em memória
        self.processed_data = None

        self.pdf_dir = os.path.join(os.path.dirname(__file__), 'pdf')
        if not os.path.exists(self.pdf_dir):
            os.makedirs(self.pdf_dir)

        self._create_widgets()
        self._update_status_bar(db_alunos_status, db_alunos_count, db_horarios_status, db_horarios_count)

    def _create_widgets(self):
        # Frame para anexar PDFs
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
        
        # Frame para os botões de ação
        action_frame = tk.Frame(self.root)
        action_frame.pack(pady=10)
        
        # Botão 1: Processar Dados
        btn_processar = tk.Button(action_frame, text="1. Processar Dados", font=('Helvetica', 10, 'bold'), bg="#008CBA", fg="white", command=self._processar_dados)
        btn_processar.pack(side=tk.LEFT, padx=10, ipadx=10, ipady=5)
        
        # Botão 2: Gerar Relatório Detalhado
        self.btn_gerar_detalhado = tk.Button(action_frame, text="2. Gerar Relatório Detalhado", state="disabled", command=self._gerar_relatorio_faltas)
        self.btn_gerar_detalhado.pack(side=tk.LEFT, padx=10)

        # Botão 3: Gerar Relatório Simples
        self.btn_gerar_simples = tk.Button(action_frame, text="3. Gerar Relatório Simples", state="disabled", command=self._gerar_relatorio_simples)
        self.btn_gerar_simples.pack(side=tk.LEFT, padx=10)
        
        # (O resto da UI, como status bar e console, permanece o mesmo)
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

    def _update_status_bar(self, alunos_status, alunos_count, horarios_status, horarios_count):
        # ... (código sem alterações) ...
        if alunos_status == "success":
            self.lbl_db_alunos_status.config(text=f"BD Alunos: Conectado ({alunos_count} registros)", fg="green")
        else:
            self.lbl_db_alunos_status.config(text="BD Alunos: Erro de Conexão", fg="red")
        if horarios_status == "success":
            self.lbl_db_horarios_status.config(text=f"BD Horários: Conectado ({horarios_count} registros)", fg="green")
        else:
            self.lbl_db_horarios_status.config(text="BD Horários: Erro de Conexão", fg="red")

    def _select_ausentes_pdf(self):
        # ... (código sem alterações) ...
        filepath = filedialog.askopenfilename(title="Selecione o PDF de Ausentes", initialdir=self.pdf_dir, filetypes=[("Arquivos PDF", "*.pdf")])
        if filepath:
            self.ausentes_pdf_path = filepath
            self.lbl_ausentes.config(text=os.path.basename(filepath), fg="black")

    def _select_frequencia_pdf(self):
        # ... (código sem alterações) ...
        filepath = filedialog.askopenfilename(title="Selecione o PDF de Frequência", initialdir=self.pdf_dir, filetypes=[("Arquivos PDF", "*.pdf")])
        if filepath:
            self.frequencia_pdf_path = filepath
            self.lbl_frequencia.config(text=os.path.basename(filepath), fg="black")
            
    def _write_to_console(self, message):
        # ... (código sem alterações) ...
        self.console.config(state='normal')
        self.console.insert(tk.END, message + "\n")
        self.console.see(tk.END)
        self.console.config(state='disabled')
        self.root.update_idletasks()

    # --- NOVAS FUNÇÕES DE AÇÃO DOS BOTÕES ---
    
    def _processar_dados(self):
        self.console.config(state='normal')
        self.console.delete('1.0', tk.END)
        self.console.config(state='disabled')
        
        self.btn_gerar_detalhado.config(state="disabled")
        self.btn_gerar_simples.config(state="disabled")
        self.processed_data = None
        
        if not self.ausentes_pdf_path or not self.frequencia_pdf_path:
            self._write_to_console("ERRO: Anexe os dois arquivos PDF antes de processar.")
            return
            
        try:
            self._write_to_console("--- INICIANDO PROCESSAMENTO DOS DADOS ---")
            
            # Chama a função de processamento e armazena os resultados
            self.processed_data = processar_dados_diarios(
                ausentes_path=self.ausentes_pdf_path,
                frequencia_path=self.frequencia_pdf_path,
                logger=self._write_to_console
            )
            
            if self.processed_data and self.processed_data[0] is not None:
                self._write_to_console("\n--- PROCESSAMENTO CONCLUÍDO ---")
                self._write_to_console("Dados processados com sucesso. Agora você pode gerar os relatórios.")
                # Habilita os botões de gerar relatório
                self.btn_gerar_detalhado.config(state="normal")
                self.btn_gerar_simples.config(state="normal")
            else:
                self._write_to_console("\n--- FALHA NO PROCESSAMENTO ---")
                self._write_to_console("Verifique as mensagens de erro acima.")

        except Exception as e:
            self._write_to_console(f"\nOcorreu um erro crítico durante o processamento:\n{e}")

    def _gerar_relatorio_faltas(self):
        if not self.processed_data:
            self._write_to_console("ERRO: Processe os dados primeiro (Botão 1).")
            return
        
        report_date, faltas_registradas, _ = self.processed_data
        gerar_relatorio_faltas(faltas_registradas, report_date, self._write_to_console)

    def _gerar_relatorio_simples(self):
        if not self.processed_data:
            self._write_to_console("ERRO: Processe os dados primeiro (Botão 1).")
            return
            
        report_date, _, df_problemas = self.processed_data
        gerar_relatorio_simples(df_problemas, report_date, self._write_to_console)