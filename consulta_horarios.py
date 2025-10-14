import sqlite3

# Define o nome do arquivo do banco de dados de horários
nome_banco_de_dados = 'horarios.db'

# ADICIONE ESTA FUNÇÃO NO topo do arquivo consulta_horarios.py

def obter_contagem_horarios():
    """
    Conecta ao banco de dados e retorna o número total de horários.
    """
    conn = None
    try:
        conn = sqlite3.connect('horarios.db')
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

def acessar_dados_horarios():
    """
    Função para conectar ao banco de dados, ler e exibir todos os horários.
    """
    conn = None
    try:
        conn = sqlite3.connect(nome_banco_de_dados)
        cursor = conn.cursor()
        query = "SELECT id, turma, dia_semana, hora_inicio, hora_fim, disciplina FROM horarios"
        cursor.execute(query)
        horarios = cursor.fetchall()
        
        if not horarios:
            print("\nNenhum horário encontrado na tabela.")
            return

        print("\n--- Lista de Horários ---")
        print(f"{'ID':<5} | {'Turma':<10} | {'Dia da Semana':<15} | {'Início':<10} | {'Fim':<10} | {'Disciplina':<25}")
        print("-" * 90)
        
        for horario in horarios:
            id, turma, dia_semana, hora_inicio, hora_fim, disciplina = horario
            print(f"{id:<5} | {turma:<10} | {dia_semana:<15} | {hora_inicio:<10} | {hora_fim:<10} | {disciplina:<25}")
        print("-" * 90)

    except sqlite3.Error as e:
        print(f"Ocorreu um erro ao acessar o banco de dados de horários: {e}")
    finally:
        if conn:
            conn.close()

def inserir_horario(turma, dia_semana, hora_inicio, hora_fim, disciplina):
    """
    Função para inserir um novo horário na tabela.
    """
    if not all([turma, dia_semana, hora_inicio, hora_fim, disciplina]):
        print("Erro: Todos os campos são obrigatórios para inserir um horário.")
        return

    conn = None
    try:
        conn = sqlite3.connect(nome_banco_de_dados)
        cursor = conn.cursor()
        query = "INSERT INTO horarios (turma, dia_semana, hora_inicio, hora_fim, disciplina) VALUES (?, ?, ?, ?, ?)"
        cursor.execute(query, (turma, dia_semana, hora_inicio, hora_fim, disciplina))
        conn.commit()
        print(f"Horário para a turma '{turma}' inserido com sucesso!")
    except sqlite3.Error as e:
        print(f"Ocorreu um erro ao inserir o horário: {e}")
    finally:
        if conn:
            conn.close()

def atualizar_horario(id_horario, nova_turma, novo_dia, nova_hora_inicio, nova_hora_fim, nova_disciplina):
    """
    Função para atualizar um horário existente, identificado pelo seu ID.
    """
    if not all([id_horario, nova_turma, novo_dia, nova_hora_inicio, nova_hora_fim, nova_disciplina]):
        print("Erro: Todos os campos, incluindo o ID, são obrigatórios para atualizar.")
        return

    conn = None
    try:
        conn = sqlite3.connect(nome_banco_de_dados)
        cursor = conn.cursor()
        query = """UPDATE horarios 
                   SET turma = ?, dia_semana = ?, hora_inicio = ?, hora_fim = ?, disciplina = ?
                   WHERE id = ?"""
        cursor.execute(query, (nova_turma, novo_dia, nova_hora_inicio, nova_hora_fim, nova_disciplina, id_horario))
        conn.commit()
        
        if cursor.rowcount == 0:
            print(f"Nenhum horário encontrado com o ID '{id_horario}'. Nenhuma alteração foi feita.")
        else:
            print(f"Horário com ID '{id_horario}' atualizado com sucesso!")
    except sqlite3.Error as e:
        print(f"Ocorreu um erro ao atualizar o horário: {e}")
    finally:
        if conn:
            conn.close()

def excluir_horario(id_horario):
    """
    Função para excluir um horário da tabela, com confirmação.
    """
    if not id_horario:
        print("Erro: O ID é obrigatório para excluir um horário.")
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
        confirmacao = input(f"Você tem certeza que deseja excluir o horário de '{disciplina}' da turma '{turma}' (ID: {id_horario})? [s/n]: ").lower()

        if confirmacao == 's':
            cursor.execute("DELETE FROM horarios WHERE id = ?", (id_horario,))
            conn.commit()
            if cursor.rowcount > 0:
                print(f"Horário com ID '{id_horario}' foi excluído com sucesso.")
            else:
                print("Não foi possível excluir o horário.")
        else:
            print("Operação de exclusão cancelada.")

    except sqlite3.Error as e:
        print(f"Ocorreu um erro ao excluir o horário: {e}")
    finally:
        if conn:
            conn.close()

# --- Exemplo de Uso ---
if __name__ == "__main__":
    # Para usar este script, você precisa primeiro criar o banco de dados 'horarios.db'
    # e a tabela 'horarios' com as colunas corretas.
    
    # Exemplo de como criar a tabela (execute uma vez):
    # conn = sqlite3.connect(nome_banco_de_dados)
    # cursor = conn.cursor()
    # cursor.execute("""
    # CREATE TABLE IF NOT EXISTS horarios (
    #     id INTEGER PRIMARY KEY AUTOINCREMENT,
    #     turma TEXT NOT NULL,
    #     dia_semana TEXT NOT NULL,
    #     hora_inicio TEXT NOT NULL,
    #     hora_fim TEXT NOT NULL,
    #     disciplina TEXT NOT NULL
    # );
    # """)
    # conn.commit()
    # conn.close()
    
    # 1. Lista os horários existentes
    print("Buscando horários existentes...")
    acessar_dados_horarios()

    # 2. Insere um novo horário
    print("\nInserindo um novo horário de teste...")
    inserir_horario("EF9999", "SEXTA-FEIRA", "10:40", "11:30", "PROJETO DE VIDA - EQUIPE")
    
    # 3. Atualiza um horário (substitua '1' pelo ID que deseja atualizar)
    print("\nAtualizando um horário...")
    atualizar_horario(1, "EF1601", "SEGUNDA-FEIRA", "07:00", "07:50", "CIÊNCIAS DA NATUREZA - LUCAS")
    
    # 4. Exclui um horário (substitua '500' ou outro ID que deseje excluir)
    print("\nExcluindo um horário...")
    excluir_horario(500) # Este ID provavelmente não existe, então mostrará a mensagem de 'não encontrado'
    
    # 5. Lista os horários novamente para ver as alterações
    acessar_dados_horarios()