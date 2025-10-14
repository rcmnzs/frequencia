import sqlite3
import os

# --- Bloco de Código Corrigido ---
# Pega o diretório onde este script (consulta_horarios.py) está localizado.
script_dir = os.path.dirname(__file__)

# Constrói o caminho para a pasta principal do projeto.
project_root = os.path.dirname(script_dir)

# Constrói o caminho completo e correto para o banco de dados.
nome_banco_de_dados = os.path.join(project_root, 'db', 'horarios.db')
# --- Fim do Bloco de Código Corrigido ---

def obter_contagem_horarios():
    """
    Conecta ao banco de dados e retorna o número total de horários.
    """
    conn = None
    try:
        conn = sqlite3.connect(nome_banco_de_dados)
        cursor = conn.cursor()
        query = "SELECT COUNT(*) FROM horarios"
        cursor.execute(query)
        contagem = cursor.fetchone()[0]
        return contagem
    except sqlite3.Error:
        return None
    finally:
        if conn:
            conn.close()

# ... (O resto das funções acessar_dados_horarios, inserir_horario, etc. permanecem as mesmas) ...

def acessar_dados_horarios():
    conn = None
    try:
        conn = sqlite3.connect(nome_banco_de_dados)
        cursor = conn.cursor()
        query = "SELECT id, turma, dia_semana, hora_inicio, hora_fim, disciplina FROM horarios"
        cursor.execute(query)
        horarios = cursor.fetchall()
        if not horarios:
            print("\nNenhum horário encontrado.")
            return
        print("\n--- Lista de Horários ---")
        print(f"{'ID':<5} | {'Turma':<10} | {'Dia da Semana':<15} | {'Início':<10} | {'Fim':<10} | {'Disciplina':<25}")
        print("-" * 90)
        for horario in horarios:
            print(f"{horario[0]:<5} | {horario[1]:<10} | {horario[2]:<15} | {horario[3]:<10} | {horario[4]:<10} | {horario[5]:<25}")
        print("-" * 90)
    except sqlite3.Error as e:
        print(f"Ocorreu um erro: {e}")
    finally:
        if conn:
            conn.close()

def inserir_horario(turma, dia_semana, hora_inicio, hora_fim, disciplina):
    if not all([turma, dia_semana, hora_inicio, hora_fim, disciplina]):
        print("Erro: Todos os campos são obrigatórios.")
        return
    conn = None
    try:
        conn = sqlite3.connect(nome_banco_de_dados)
        cursor = conn.cursor()
        query = "INSERT INTO horarios (turma, dia_semana, hora_inicio, hora_fim, disciplina) VALUES (?, ?, ?, ?, ?)"
        cursor.execute(query, (turma, dia_semana, hora_inicio, hora_fim, disciplina))
        conn.commit()
        print(f"Horário para a turma '{turma}' inserido.")
    except sqlite3.Error as e:
        print(f"Ocorreu um erro: {e}")
    finally:
        if conn:
            conn.close()

def atualizar_horario(id_horario, nova_turma, novo_dia, nova_hora_inicio, nova_hora_fim, nova_disciplina):
    if not all([id_horario, nova_turma, novo_dia, nova_hora_inicio, nova_hora_fim, nova_disciplina]):
        print("Erro: Todos os campos são obrigatórios.")
        return
    conn = None
    try:
        conn = sqlite3.connect(nome_banco_de_dados)
        cursor = conn.cursor()
        query = "UPDATE horarios SET turma = ?, dia_semana = ?, hora_inicio = ?, hora_fim = ?, disciplina = ? WHERE id = ?"
        cursor.execute(query, (nova_turma, novo_dia, nova_hora_inicio, nova_hora_fim, nova_disciplina, id_horario))
        conn.commit()
        if cursor.rowcount == 0:
            print(f"Nenhum horário com ID '{id_horario}' encontrado.")
        else:
            print(f"Horário com ID '{id_horario}' atualizado.")
    except sqlite3.Error as e:
        print(f"Ocorreu um erro: {e}")
    finally:
        if conn:
            conn.close()

def excluir_horario(id_horario):
    if not id_horario:
        print("Erro: ID é obrigatório.")
        return
    conn = None
    try:
        conn = sqlite3.connect(nome_banco_de_dados)
        cursor = conn.cursor()
        cursor.execute("SELECT turma, disciplina FROM horarios WHERE id = ?", (id_horario,))
        horario_existente = cursor.fetchone()
        if not horario_existente:
            print(f"Horário com ID '{id_horario}' não encontrado.")
            return
        turma, disciplina = horario_existente
        confirmacao = input(f"Tem certeza que deseja excluir o horário de '{disciplina}' da turma '{turma}' (ID: {id_horario})? [s/n]: ").lower()
        if confirmacao == 's':
            cursor.execute("DELETE FROM horarios WHERE id = ?", (id_horario,))
            conn.commit()
            if cursor.rowcount > 0:
                print(f"Horário com ID '{id_horario}' excluído.")
            else:
                print("Exclusão falhou.")
        else:
            print("Exclusão cancelada.")
    except sqlite3.Error as e:
        print(f"Ocorreu um erro: {e}")
    finally:
        if conn:
            conn.close()

if __name__ == "__main__":
    print("--- Testando Módulo de Consulta de Horários ---")
    acessar_dados_horarios()