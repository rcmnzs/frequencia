import tkinter as tk
from interface import App 

# Ajusta a forma de importar para refletir a nova pasta 'modulos'
from modulos.consulta_alunos import obter_contagem_alunos
from modulos.consulta_horarios import obter_contagem_horarios

def main():
    # ... (o resto do código permanece o mesmo) ...
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