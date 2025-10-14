import sqlite3

# Define o nome do arquivo do banco de dados
nome_banco_de_dados = 'alunos.db'

def obter_contagem_alunos():
    """
    Conecta ao banco de dados e retorna o número total de alunos.
    Retorna um número (int) em caso de sucesso ou None em caso de erro.
    """
    conn = None
    try:
        conn = sqlite3.connect(nome_banco_de_dados)
        cursor = conn.cursor()
        # Query otimizada para apenas contar as linhas
        query = "SELECT COUNT(*) FROM alunos"
        cursor.execute(query)
        # fetchone() retornará uma tupla como (150,), então pegamos o primeiro item [0]
        contagem = cursor.fetchone()[0]
        return contagem
    except sqlite3.Error as e:
        print(f"Erro ao acessar o banco de dados: {e}")
        return None # Retorna None para indicar que a conexão falhou
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

def inserir_aluno(matricula, nome, turma):
    """
    Função para inserir um novo aluno na tabela 'alunos'.
    Todos os campos são obrigatórios.
    """
    if not all([matricula, nome, turma]):
        print("Erro: Todos os campos (matrícula, nome e turma) são obrigatórios.")
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
        print(f"Erro: A matrícula '{matricula}' já existe no banco de dados.")
    except sqlite3.Error as e:
        print(f"Ocorreu um erro ao inserir dados no banco de dados: {e}")
    finally:
        if conn:
            conn.close()

def atualizar_aluno(matricula, novo_nome, nova_turma):
    """
    Função para atualizar o nome e a turma de um aluno existente,
    identificado pela matrícula.
    """
    if not all([matricula, novo_nome, nova_turma]):
        print("Erro: A matrícula, o novo nome e a nova turma são obrigatórios.")
        return

    conn = None
    try:
        conn = sqlite3.connect(nome_banco_de_dados)
        cursor = conn.cursor()
        query = "UPDATE alunos SET nome = ?, turma = ? WHERE matricula = ?"
        cursor.execute(query, (novo_nome, nova_turma, matricula))
        conn.commit()
        if cursor.rowcount == 0:
            print(f"Nenhum aluno encontrado com a matrícula '{matricula}'. Nenhuma alteração foi feita.")
        else:
            print(f"Dados do aluno com matrícula '{matricula}' atualizados com sucesso!")
    except sqlite3.Error as e:
        print(f"Ocorreu um erro ao atualizar os dados: {e}")
    finally:
        if conn:
            conn.close()

def excluir_aluno(matricula):
    """
    Função para excluir um aluno da tabela, com confirmação.
    """
    if not matricula:
        print("Erro: A matrícula é obrigatória para excluir um aluno.")
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
        confirmacao = input(f"Você tem certeza que deseja excluir o aluno '{nome_aluno}' (matrícula: {matricula})? [s/n]: ").lower()

        if confirmacao == 's':
            cursor.execute("DELETE FROM alunos WHERE matricula = ?", (matricula,))
            conn.commit()
            if cursor.rowcount > 0:
                print(f"Aluno '{nome_aluno}' foi excluído com sucesso.")
            else:
                print("Não foi possível excluir o aluno. Registro não encontrado.")
        else:
            print("Operação de exclusão cancelada pelo usuário.")
    except sqlite3.Error as e:
        print(f"Ocorreu um erro ao excluir o aluno: {e}")
    finally:
        if conn:
            conn.close()

# --- Bloco de Exemplo de Uso (quando o script é executado diretamente) ---
if __name__ == "__main__":
    print("--- Módulo de Consulta de Alunos (Teste) ---")
    
    # 1. Lista os alunos para sabermos quem existe
    print("\nBuscando dados existentes...")
    acessar_dados_alunos()
    
    # 2. Insere um aluno apenas para o teste
    print("\nInserindo aluno temporário para teste...")
    inserir_aluno("9999999999", "Aluno a ser Excluido", "EX9999")
    acessar_dados_alunos()

    # 3. Tenta excluir o aluno recém-criado
    print("\n--- Teste de Exclusão ---")
    matricula_para_excluir = "9999999999"
    # Para testes automáticos, podemos simular a entrada 's'
    # Em uso real, a linha abaixo seria: excluir_aluno(matricula_para_excluir)
    print(f"Simulando exclusão do aluno com matrícula {matricula_para_excluir}. (Em um teste real, o input seria solicitado)")

    # 4. Mostra a lista final para confirmar
    print("\nBuscando dados finais...")
    acessar_dados_alunos()