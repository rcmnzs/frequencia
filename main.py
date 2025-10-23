import tkinter as tk
import sys
import os

# --- CORREÇÃO DE IMPORTAÇÃO ---
# Adiciona o diretório raiz do projeto (onde 'main.py' está) ao path do sistema.
# Isso garante que 'import config' e 'from modulos import ...' funcionem corretamente.
sys.path.append(os.path.dirname(os.path.abspath(__file__)))
# --- FIM DA CORREÇÃO ---

from interface import App 
from modulos.consulta_alunos import obter_contagem_alunos
from modulos.consulta_horarios import obter_contagem_horarios

def main():
    print("Verificando conexão com os bancos de dados...")
    
    alunos_count = obter_contagem_alunos()
    alunos_status = "success" if alunos_count is not None else "error"
    print(f"BD Alunos: {alunos_status} ({alunos_count} registros)")
    
    horarios_count = obter_contagem_horarios()
    horarios_status = "success" if horarios_count is not None else "error"
    print(f"BD Horários: {horarios_status} ({horarios_count} registros)")

    root = tk.Tk()
    app = App(root, 
              db_alunos_status=alunos_status, db_alunos_count=alunos_count,
              db_horarios_status=horarios_status, db_horarios_count=horarios_count)
    root.mainloop()

if __name__ == "__main__":
    main()