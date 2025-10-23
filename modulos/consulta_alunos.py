import sqlite3
import os

script_dir = os.path.dirname(__file__)
project_root = os.path.dirname(script_dir)
nome_banco_de_dados = os.path.join(project_root, 'db', 'unico.db')

def obter_contagem_alunos():
    """
    Conecta ao banco de dados e retorna o número total de alunos.
    """
    conn = None
    try:
        conn = sqlite3.connect(nome_banco_de_dados)
        cursor = conn.cursor()
        query = "SELECT COUNT(*) FROM alunos"
        cursor.execute(query)
        contagem = cursor.fetchone()[0]
        return contagem
    except sqlite3.Error as e:
        print(f"Erro ao acessar o banco de dados de alunos: {e}")
        return None
    finally:
        if conn:
            conn.close()

def acessar_dados_alunos():
    """
    Função para conectar ao banco de dados, ler e exibir todos os dados da tabela 'alunos'.
    """
    conn = None
    try:
        conn = sqlite3.connect(nome_banco_de_dados)
        cursor = conn.cursor()
        query = "SELECT matricula, nome, turma FROM alunos"
        cursor.execute(query)
        alunos = cursor.fetchall()
        
        if not alunos:
            print("\nNenhum aluno encontrado na tabela.")
            return

        print("\n--- Lista de Alunos ---")
        print(f"{'Matrícula':<15} | {'Nome':<40} | {'Turma':<10}")
        print("-" * 70)
        
        for aluno in alunos:
            matricula, nome, turma = aluno
            print(f"{matricula:<15} | {nome:<40} | {turma:<10}")
        print("-" * 70)
    except sqlite3.Error as e:
        print(f"Ocorreu um erro ao acessar o banco de dados: {e}")
    finally:
        if conn:
            conn.close()

# ... (O resto das funções inserir_aluno, atualizar_aluno, excluir_aluno e o bloco __main__ permanecem os mesmos) ...

def inserir_aluno(matricula, nome, turma):
    if not all([matricula, nome, turma]):
        print("Erro: Todos os campos são obrigatórios.")
        return
    conn = None
    try:
        conn = sqlite3.connect(nome_banco_de_dados)
        cursor = conn.cursor()
        query = "INSERT INTO alunos (matricula, nome, turma) VALUES (?, ?, ?)"
        cursor.execute(query, (matricula, nome, turma))
        conn.commit()
        print(f"Aluno '{nome}' inserido com sucesso!")
    except sqlite3.IntegrityError:
        print(f"Erro: A matrícula '{matricula}' já existe.")
    except sqlite3.Error as e:
        print(f"Ocorreu um erro: {e}")
    finally:
        if conn:
            conn.close()

def atualizar_aluno(matricula, novo_nome, nova_turma):
    if not all([matricula, novo_nome, nova_turma]):
        print("Erro: Todos os campos são obrigatórios.")
        return
    conn = None
    try:
        conn = sqlite3.connect(nome_banco_de_dados)
        cursor = conn.cursor()
        query = "UPDATE alunos SET nome = ?, turma = ? WHERE matricula = ?"
        cursor.execute(query, (novo_nome, nova_turma, matricula))
        conn.commit()
        if cursor.rowcount == 0:
            print(f"Nenhum aluno com matrícula '{matricula}' encontrado.")
        else:
            print(f"Dados do aluno com matrícula '{matricula}' atualizados.")
    except sqlite3.Error as e:
        print(f"Ocorreu um erro: {e}")
    finally:
        if conn:
            conn.close()

def excluir_aluno(matricula):
    if not matricula:
        print("Erro: Matrícula é obrigatória.")
        return
    conn = None
    try:
        conn = sqlite3.connect(nome_banco_de_dados)
        cursor = conn.cursor()
        cursor.execute("SELECT nome FROM alunos WHERE matricula = ?", (matricula,))
        aluno_existente = cursor.fetchone()
        if not aluno_existente:
            print(f"Aluno com matrícula '{matricula}' não encontrado.")
            return
        nome_aluno = aluno_existente[0]
        confirmacao = input(f"Tem certeza que deseja excluir '{nome_aluno}' (matrícula: {matricula})? [s/n]: ").lower()
        if confirmacao == 's':
            cursor.execute("DELETE FROM alunos WHERE matricula = ?", (matricula,))
            conn.commit()
            if cursor.rowcount > 0:
                print(f"Aluno '{nome_aluno}' excluído com sucesso.")
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
    print("--- Testando Módulo de Consulta de Alunos ---")
    acessar_dados_alunos()