import tkinter as tk
from interface import App 
from consulta_alunos import obter_contagem_alunos
# Importa a nova função de contagem de horários
from consulta_horarios import obter_contagem_horarios

def main():
    """
    Função principal que prepara os dados dos bancos de dados e inicializa a interface.
    """
    print("Verificando conexão com os bancos de dados...")
    
    # 1. Verifica o banco de dados de alunos
    alunos_count = obter_contagem_alunos()
    alunos_status = "success" if alunos_count is not None else "error"
    print(f"BD Alunos: {alunos_status} ({alunos_count} registros)")
    
    # 2. Verifica o banco de dados de horários
    horarios_count = obter_contagem_horarios()
    horarios_status = "success" if horarios_count is not None else "error"
    print(f"BD Horários: {horarios_status} ({horarios_count} registros)")

    # 3. Inicializa a interface gráfica
    root = tk.Tk()
    
    # 4. Cria a instância da aplicação, passando todos os dados de status
    app = App(root, 
              db_alunos_status=alunos_status, db_alunos_count=alunos_count,
              db_horarios_status=horarios_status, db_horarios_count=horarios_count)
    
    root.mainloop()

if __name__ == "__main__":
    main()