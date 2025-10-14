import sqlite3
import pandas as pd
import os
import re
import fitz  # PyMuPDF
from datetime import datetime

# A função de busca permanece a mesma
def buscar_aluno(df_alunos, matricula_pdf=None, nome_pdf=None):
    if matricula_pdf:
        resultado = df_alunos[df_alunos['matricula'] == str(matricula_pdf)]
        if not resultado.empty: return resultado.iloc[0]
    if nome_pdf:
        nome_pdf_normalizado = ' '.join(nome_pdf.strip().upper().split())
        resultado = df_alunos[df_alunos['nome'].str.strip().str.upper() == nome_pdf_normalizado]
        if not resultado.empty: return resultado.iloc[0]
        palavras_pdf = nome_pdf_normalizado.split()
        if len(palavras_pdf) < 2: return None
        for index, row in df_alunos.iterrows():
            nome_db_normalizado = ' '.join(row['nome'].strip().upper().split())
            palavras_db = nome_db_normalizado.split()
            if len(palavras_pdf) > len(palavras_db): continue
            match = True
            for i in range(len(palavras_pdf)):
                palavra_pdf, palavra_db = palavras_pdf[i], palavras_db[i]
                if i == len(palavras_pdf) - 1:
                    if not palavra_db.startswith(palavra_pdf): match = False; break
                else:
                    if palavra_pdf != palavra_db: match = False; break
            if match: return row
    return None

def carregar_dados_base(logger):
    logger("Conectando aos bancos de dados...")
    try:
        script_dir = os.path.dirname(__file__)
        project_root = os.path.dirname(script_dir)
        path_alunos_db = os.path.join(project_root, 'db', 'alunos.db')
        path_horarios_db = os.path.join(project_root, 'db', 'horarios.db')
        
        conn_alunos = sqlite3.connect(path_alunos_db)
        df_alunos = pd.read_sql_query("SELECT * FROM alunos", conn_alunos)
        conn_alunos.close()
        
        conn_horarios = sqlite3.connect(path_horarios_db)
        df_horarios = pd.read_sql_query("SELECT * FROM horarios", conn_horarios)
        conn_horarios.close()

        df_alunos['matricula'] = df_alunos['matricula'].astype(str)
        df_horarios['hora_inicio'] = pd.to_datetime(df_horarios['hora_inicio'], format='%H:%M').dt.time
        df_horarios['hora_fim'] = pd.to_datetime(df_horarios['hora_fim'], format='%H:%M').dt.time
        
        logger("Bancos de dados carregados com sucesso.")
        return df_alunos, df_horarios
    except Exception as e:
        logger(f"ERRO ao carregar os bancos de dados: {e}")
        return None, None

def processar_arquivos_ausentes(df_alunos, df_horarios, faltas_registradas, ausentes_path, logger):
    logger("\n--- Processando Relatório de Ausentes ---")
    dias_semana_map = {'Monday': 'SEGUNDA-FEIRA', 'Tuesday': 'TERÇA-FEIRA', 'Wednesday': 'QUARTA-FEIRA',
                       'Thursday': 'QUINTA-FEIRA', 'Friday': 'SEXTA-FEIRA', 'Saturday': 'SÁBADO', 'Sunday': 'DOMINGO'}

    logger(f"Analisando arquivo: {os.path.basename(ausentes_path)}...")
    doc = fitz.open(ausentes_path)
    content = "".join(page.get_text() for page in doc)
    doc.close()

    match_data = re.search(r"Período: de (\d{2}/\d{2}/\d{4})", content)
    if not match_data:
        logger(f"  ERRO: Não foi possível encontrar a data no PDF de ausentes. O processo não pode continuar.")
        return None # Retorna None se a data não for encontrada
    
    report_date = datetime.strptime(match_data.group(1), '%d/%m/%Y')
    dia_semana = dias_semana_map.get(report_date.strftime('%A'))
    logger(f"  Data do relatório identificada: {report_date.date()}")

    ausentes = re.findall(r'(\d{9,11})\s+([A-ZÀ-Ú\s.-]+?)\s+ALUNO', content)
    logger(f"  Encontrados {len(ausentes)} alunos ausentes.")

    for matricula_pdf, nome_pdf in ausentes:
        info_aluno = buscar_aluno(df_alunos, matricula_pdf=matricula_pdf, nome_pdf=nome_pdf)
        if info_aluno is not None:
            turma, nome_db, matricula_db = info_aluno['turma'], info_aluno['nome'], info_aluno['matricula']
            aulas_do_dia = df_horarios[(df_horarios['turma'] == turma) & (df_horarios['dia_semana'] == dia_semana)]
            for _, aula in aulas_do_dia.iterrows():
                chave_falta = (matricula_db, nome_db, turma, aula['disciplina'])
                faltas_registradas[chave_falta] = faltas_registradas.get(chave_falta, 0) + 1
        else:
            logger(f"  Aviso: Aluno '{nome_pdf.strip()}' (Matrícula: {matricula_pdf}) do PDF não encontrado no BD.")
            
    return report_date # Retorna o objeto de data para ser usado por outras funções

def processar_arquivos_frequencia(df_alunos, df_horarios, faltas_registradas, frequencia_path, report_date, logger):
    logger("\n--- Processando Relatório de Frequência (Atrasos/Saídas) ---")
    dias_semana_map = {'Monday': 'SEGUNDA-FEIRA', 'Tuesday': 'TERÇA-FEIRA', 'Wednesday': 'QUARTA-FEIRA',
                       'Thursday': 'QUINTA-FEIRA', 'Friday': 'SEXTA-FEIRA', 'Saturday': 'SÁBADO', 'Sunday': 'DOMINGO'}
    dia_semana = dias_semana_map.get(report_date.strftime('%A'))
    
    logger(f"Analisando arquivo: {os.path.basename(frequencia_path)} para o dia {report_date.date()}...")
    from modulos.extrator_frequencias import extrair_dados_frequencia
    df_acessos = extrair_dados_frequencia(frequencia_path)

    if df_acessos is None or df_acessos.empty:
        logger(f"  Nenhum dado de acesso extraído do arquivo.")
        return
    
    df_acessos['Hora'] = pd.to_datetime(df_acessos['Hora'], format='%H:%M:%S').dt.time
    
    for grupo_keys, acesso_aluno_df in df_acessos.groupby(['Crachá', 'Nome']):
        matricula_pdf, nome_pdf = grupo_keys
        info_aluno = buscar_aluno(df_alunos, matricula_pdf=matricula_pdf, nome_pdf=nome_pdf)
        
        if info_aluno is None:
            logger(f"  Aviso: Aluno '{nome_pdf.strip()}' (Crachá: {matricula_pdf}) do PDF não encontrado no BD.")
            continue

        turma, nome_db, matricula_db = info_aluno['turma'], info_aluno['nome'], info_aluno['matricula']
        aulas_do_dia = df_horarios[(df_horarios['turma'] == turma) & (df_horarios['dia_semana'] == dia_semana)]

        primeira_entrada = acesso_aluno_df[acesso_aluno_df['Sentido'] == 'Entrada']['Hora'].min() if not acesso_aluno_df[acesso_aluno_df['Sentido'] == 'Entrada'].empty else None
        ultima_saida = acesso_aluno_df[acesso_aluno_df['Sentido'] == 'Saída']['Hora'].max() if not acesso_aluno_df[acesso_aluno_df['Sentido'] == 'Saída'].empty else None

        if primeira_entrada:
            for _, aula in aulas_do_dia[aulas_do_dia['hora_fim'] < primeira_entrada].iterrows():
                chave_falta = (matricula_db, nome_db, turma, aula['disciplina'])
                faltas_registradas[chave_falta] = faltas_registradas.get(chave_falta, 0) + 1
                logger(f"  Falta por CHEGADA TARDIA para {nome_db} em {aula['disciplina']}")
        if ultima_saida:
            for _, aula in aulas_do_dia[aulas_do_dia['hora_inicio'] > ultima_saida].iterrows():
                chave_falta = (matricula_db, nome_db, turma, aula['disciplina'])
                faltas_registradas[chave_falta] = faltas_registradas.get(chave_falta, 0) + 1
                logger(f"  Falta por SAÍDA ANTECIPADA para {nome_db} em {aula['disciplina']}")

# --- FUNÇÃO DE GERAÇÃO DE RELATÓRIO TOTALMENTE REESCRITA ---
def gerar_relatorio_excel(faltas_registradas, report_date, logger):
    logger("\n--- Gerando Relatório Final em Excel ---")
    if not faltas_registradas:
        logger("Nenhuma falta foi registrada. Nenhum relatório será gerado.")
        return False

    # Define o nome da aba com base na data (formato amigável para abas de Excel)
    sheet_name = report_date.strftime('%d-%m-%Y')
    
    # Define o caminho para a pasta e o arquivo de saída
    script_dir = os.path.dirname(__file__)
    project_root = os.path.dirname(script_dir)
    relatorios_dir = os.path.join(project_root, 'relatorios')
    os.makedirs(relatorios_dir, exist_ok=True) # Cria a pasta se não existir
    output_path = os.path.join(relatorios_dir, 'relatorio_faltas.xlsx')

    logger(f"Preparando para escrever na aba '{sheet_name}' do arquivo '{output_path}'...")
    
    lista_faltas = [list(chave) + [valor] for chave, valor in faltas_registradas.items()]
    df_final = pd.DataFrame(lista_faltas, columns=['Matricula', 'Nome', 'Turma', 'Disciplina', 'Total de Faltas'])
    df_final.sort_values(by=['Turma', 'Nome', 'Disciplina'], inplace=True)
    
    try:
        # Usa ExcelWriter para adicionar ou substituir uma aba sem apagar as outras
        with pd.ExcelWriter(output_path, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
            df_final.to_excel(writer, sheet_name=sheet_name, index=False)
        
        logger(f"Relatório de faltas foi salvo/atualizado com sucesso!")
        return True
    except FileNotFoundError:
        # Se o arquivo não existe, o modo 'a' do ExcelWriter não funciona em algumas versões.
        # Criamos o arquivo pela primeira vez.
        df_final.to_excel(output_path, sheet_name=sheet_name, index=False)
        logger(f"Novo relatório de faltas criado e salvo com sucesso!")
        return True
    except Exception as e:
        logger(f"ERRO ao salvar o arquivo Excel: {e}")
        return False

def executar_logica_completa(ausentes_path, frequencia_path, logger):
    df_alunos, df_horarios = carregar_dados_base(logger)
    if df_alunos is None or df_horarios is None:
        return

    faltas_registradas = {}
    
    # Processa ausentes e captura a data do relatório
    data_do_relatorio = processar_arquivos_ausentes(df_alunos, df_horarios, faltas_registradas, ausentes_path, logger)
    
    # Se não houver data, o processo para
    if data_do_relatorio is None:
        logger("Processo interrompido por falta de data no relatório.")
        return
        
    # Passa a data para o processamento de frequência
    processar_arquivos_frequencia(df_alunos, df_horarios, faltas_registradas, frequencia_path, data_do_relatorio, logger)
    
    # Passa a data para a geração do relatório
    gerar_relatorio_excel(faltas_registradas, data_do_relatorio, logger)