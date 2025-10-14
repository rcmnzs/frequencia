import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext
import os

# 1. Ajusta a importação para buscar o módulo 'logica' dentro da pasta 'modulos'
try:
    from modulos.logica import executar_logica_completa
except ImportError:
    print("ERRO: Certifique-se de que a pasta 'modulos' e o arquivo 'logica.py' existem no mesmo diretório que 'main.py'.")
    exit()

class App:
    def __init__(self, root, db_alunos_status, db_alunos_count, db_horarios_status, db_horarios_count):
        self.root = root
        self.root.title("Sistema de Apuração de Faltas")
        self.root.geometry("600x400")

        self.ausentes_pdf_path = ""
        self.frequencia_pdf_path = ""

        # 2. Define o caminho para a pasta 'pdf' e a cria se não existir
        # os.path.dirname(__file__) pega o diretório onde o script está rodando.
        self.pdf_dir = os.path.join(os.path.dirname(__file__), 'pdf')
        if not os.path.exists(self.pdf_dir):
            os.makedirs(self.pdf_dir)

        self._create_widgets()
        self._update_status_bar(db_alunos_status, db_alunos_count, db_horarios_status, db_horarios_count)

    def _create_widgets(self):
        # Frame para os botões de anexar
        top_frame = tk.Frame(self.root, padx=10, pady=10)
        top_frame.pack(fill=tk.X, side=tk.TOP)

        btn_ausentes = tk.Button(top_frame, text="Anexar PDF de Ausentes", command=self._select_ausentes_pdf)
        btn_ausentes.pack(side=tk.LEFT, padx=5)
        self.lbl_ausentes = tk.Label(top_frame, text="Nenhum arquivo selecionado", fg="grey")
        self.lbl_ausentes.pack(side=tk.LEFT, fill=tk.X, expand=True)

        btn_frequencia = tk.Button(top_frame, text="Anexar PDF de Frequência", command=self._select_frequencia_pdf)
        btn_frequencia.pack(side=tk.LEFT, padx=(20, 5))
        self.lbl_frequencia = tk.Label(top_frame, text="Nenhum arquivo selecionado", fg="grey")
        self.lbl_frequencia.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        btn_processar = tk.Button(self.root, text="Gerar Relatório de Faltas", font=('Helvetica', 10, 'bold'), bg="#4CAF50", fg="white", command=self._process_files)
        btn_processar.pack(pady=10, ipadx=10, ipady=5)

        # Barra de Status no Rodapé
        status_frame = tk.Frame(self.root, bd=1, relief=tk.SUNKEN)
        status_frame.pack(side=tk.BOTTOM, fill=tk.X)

        self.lbl_db_alunos_status = tk.Label(status_frame, text="BD Alunos: -", anchor='w')
        self.lbl_db_alunos_status.pack(side=tk.LEFT, padx=5)
        
        self.lbl_db_horarios_status = tk.Label(status_frame, text="BD Horários: -", anchor='w')
        self.lbl_db_horarios_status.pack(side=tk.LEFT, padx=5)

        # Abas e Console
        notebook = ttk.Notebook(self.root)
        console_frame = tk.Frame(notebook)
        notebook.add(console_frame, text='Console de Processamento')
        notebook.pack(expand=True, fill='both', padx=10, pady=(0, 10))

        self.console = scrolledtext.ScrolledText(console_frame, wrap=tk.WORD, state='disabled', font=("Courier New", 9))
        self.console.pack(expand=True, fill='both')

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
        """Abre a caixa de diálogo diretamente na pasta 'pdf'."""
        filepath = filedialog.askopenfilename(
            title="Selecione o PDF de Ausentes",
            initialdir=self.pdf_dir, # Define o diretório inicial
            filetypes=[("Arquivos PDF", "*.pdf")]
        )
        if filepath:
            self.ausentes_pdf_path = filepath
            self.lbl_ausentes.config(text=os.path.basename(filepath), fg="black")

    def _select_frequencia_pdf(self):
        """Abre a caixa de diálogo diretamente na pasta 'pdf'."""
        filepath = filedialog.askopenfilename(
            title="Selecione o PDF de Frequência",
            initialdir=self.pdf_dir, # Define o diretório inicial
            filetypes=[("Arquivos PDF", "*.pdf")]
        )
        if filepath:
            self.frequencia_pdf_path = filepath
            self.lbl_frequencia.config(text=os.path.basename(filepath), fg="black")
            
    def _write_to_console(self, message):
        self.console.config(state='normal')
        self.console.insert(tk.END, message + "\n")
        self.console.see(tk.END)
        self.console.config(state='disabled')
        self.root.update_idletasks()

    def _process_files(self):
        try:
            self._write_to_console("--- INICIANDO PROCESSAMENTO COMPLETO ---")
            
            executar_logica_completa(
                ausentes_path=self.ausentes_pdf_path,
                frequencia_path=self.frequencia_pdf_path,
                logger=self._write_to_console
            )
            
            self._write_to_console("\n--- PROCESSAMENTO CONCLUÍDO ---")
            # --- MENSAGEM ATUALIZADA AQUI ---
            self._write_to_console("O arquivo 'relatorios/relatorio_faltas.xlsx' foi criado ou atualizado.")

        except Exception as e:
            self._write_to_console(f"\nOcorreu um erro crítico durante o processamento:\n{e}")
        
        self.console.config(state='normal')
        self.console.delete('1.0', tk.END)
        self.console.config(state='disabled')
        
        if not self.ausentes_pdf_path or not self.frequencia_pdf_path:
            self._write_to_console("ERRO: Por favor, anexe os dois arquivos PDF antes de processar.")
            return
            
        try:
            self._write_to_console("--- INICIANDO PROCESSAMENTO COMPLETO ---")
            
            executar_logica_completa(
                ausentes_path=self.ausentes_pdf_path,
                frequencia_path=self.frequencia_pdf_path,
                logger=self._write_to_console
            )
            
            self._write_to_console("\n--- PROCESSAMENTO CONCLUÍDO ---")
            self._write_to_console("O arquivo 'relatorio_faltas.xlsx' foi gerado na pasta principal do programa.")

        except Exception as e:
            self._write_to_console(f"\nOcorreu um erro crítico durante o processamento:\n{e}")

# O if __name__ == "__main__" foi removido daqui, pois este script
# agora é importado e controlado pelo main.py