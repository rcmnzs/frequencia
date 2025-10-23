import pandas as pd
import os
import re
from datetime import datetime
import locale
import sqlite3
import fitz
import sys

# Adiciona o diretório raiz ao path para encontrar o 'config'
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import config

from .extrator_ausentes import extrair_dados_ausentes
from .extrator_frequencias import extrair_dados_frequencia

try:
    locale.setlocale(locale.LC_TIME, 'pt_BR.UTF-8')
except locale.Error:
    print("Aviso: Locale pt_BR.UTF-8 não encontrado.")

def buscar_aluno(df_alunos, matricula_pdf=None, nome_pdf=None, logger=print):
    # ... (código inalterado) ...
    if matricula_pdf:
        resultado = df_alunos[df_alunos['matricula'] == str(matricula_pdf)]
        if not resultado.empty: return resultado.iloc[0]
    if nome_pdf:
        nome_pdf_normalizado = ' '.join(nome_pdf.strip().upper().split())
        resultado = df_alunos[df_alunos['nome'].str.strip().str.upper() == nome_pdf_normalizado]
        if not resultado.empty: return resultado.iloc[0]
        palavras_pdf = nome_pdf_normalizado.split()
        if len(palavras_pdf) < 2: return None
        correspondencias_parciais = []
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
            if match: correspondencias_parciais.append(row)
        if len(correspondencias_parciais) == 1:
            return correspondencias_parciais[0]
        elif len(correspondencias_parciais) > 1:
            logger(f"  AVISO: Múltiplos alunos para '{nome_pdf}'.")
    return None

def carregar_dados_base(logger):
    logger("Conectando ao banco de dados unificado...")
    try:
        conn = sqlite3.connect(config.DB_PATH)
        df_alunos = pd.read_sql_query("SELECT * FROM alunos", conn)
        df_horarios = pd.read_sql_query("SELECT * FROM horarios", conn)
        conn.close()
        df_alunos['matricula'] = df_alunos['matricula'].astype(str)
        df_horarios['hora_inicio'] = pd.to_datetime(df_horarios['hora_inicio'], format='%H:%M').dt.time
        df_horarios['hora_fim'] = pd.to_datetime(df_horarios['hora_fim'], format='%H:%M').dt.time
        logger("Bancos de dados carregados com sucesso.")
        return df_alunos, df_horarios
    except Exception as e:
        logger(f"ERRO ao carregar banco de dados: {e}")
        return None, None

def processar_dados_diarios(ausentes_path, frequencia_path, logger):
    df_alunos, df_horarios = carregar_dados_base(logger)
    if df_alunos is None or df_horarios is None:
        return None, None, None

    faltas_registradas = {}
    problemas_alunos = []
    
    _, date_ausentes = extrair_dados_ausentes(ausentes_path)
    _, date_frequencia = extrair_dados_frequencia(frequencia_path)

    if date_ausentes is None or date_frequencia is None or date_ausentes.date() != date_frequencia.date():
        logger("ERRO CRÍTICO: Datas dos PDFs não coincidem ou não foram encontradas.")
        return None, None, None
        
    report_date = date_ausentes
    dias_semana_map = {0: 'SEGUNDA-FEIRA', 1: 'TERÇA-FEIRA', 2: 'QUARTA-FEIRA', 3: 'QUINTA-FEIRA', 4: 'SEXTA-FEIRA', 5: 'SÁBADO', 6: 'DOMINGO'}
    dia_semana = dias_semana_map.get(report_date.weekday())
    
    # Resto da lógica de processamento de ausentes e frequência, exatamente como estava em `logica.py`
    # ... (código omitido para brevidade, mas é o mesmo) ...
    df_ausentes, _ = extrair_dados_ausentes(ausentes_path)
    if df_ausentes is not None:
        for _, row in df_ausentes.iterrows():
            info_aluno = buscar_aluno(df_alunos, matricula_pdf=row['Matrícula'], nome_pdf=row['Nome'], logger=logger)
            if info_aluno is not None:
                turma, nome_db, matricula_db = info_aluno['turma'], info_aluno['nome'], info_aluno['matricula']
                problemas_alunos.append({'Matricula': matricula_db, 'Nome do Aluno': nome_db, 'Turma': turma, 'Problema': 'FALTOU', 'Acesso': 'Sem registro'})
                aulas_do_dia = df_horarios[(df_horarios['turma'] == turma) & (df_horarios['dia_semana'] == dia_semana)]
                for _, aula in aulas_do_dia.iterrows():
                    chave_falta = (matricula_db, nome_db, turma, aula['disciplina'])
                    faltas_registradas[chave_falta] = faltas_registradas.get(chave_falta, 0) + 1
    
    df_acessos, _ = extrair_dados_frequencia(frequencia_path)
    if df_acessos is not None:
        df_acessos['Hora'] = pd.to_datetime(df_acessos['Hora'], format='%H:%M:%S').dt.time
        for grupo_keys, acesso_aluno_df in df_acessos.groupby(['Crachá', 'Nome']):
            info_aluno = buscar_aluno(df_alunos, matricula_pdf=grupo_keys[0], nome_pdf=grupo_keys[1], logger=logger)
            if info_aluno is not None:
                turma, nome_db, matricula_db = info_aluno['turma'], info_aluno['nome'], info_aluno['matricula']
                aulas_do_dia = df_horarios[(df_horarios['turma'] == turma) & (df_horarios['dia_semana'] == dia_semana)]
                if not aulas_do_dia.empty:
                    primeira_aula_do_dia, ultima_aula_do_dia = aulas_do_dia['hora_inicio'].min(), aulas_do_dia['hora_fim'].max()
                    primeira_entrada, ultima_saida = acesso_aluno_df[acesso_aluno_df['Sentido'] == 'Entrada']['Hora'].min(), acesso_aluno_df[acesso_aluno_df['Sentido'] == 'Saída']['Hora'].max()
                    if pd.notna(primeira_entrada) and primeira_entrada > primeira_aula_do_dia:
                        problemas_alunos.append({'Matricula': matricula_db, 'Nome do Aluno': nome_db, 'Turma': turma, 'Problema': 'CHEGOU ATRASADO', 'Acesso': f"Entrada: {primeira_entrada.strftime('%H:%M:%S')}"})
                    if pd.notna(ultima_saida) and ultima_saida < ultima_aula_do_dia:
                        problemas_alunos.append({'Matricula': matricula_db, 'Nome do Aluno': nome_db, 'Turma': turma, 'Problema': 'SAIU CEDO', 'Acesso': f"Saída: {ultima_saida.strftime('%H:%M:%S')}"})
                    if pd.notna(primeira_entrada):
                        for _, aula in aulas_do_dia[aulas_do_dia['hora_fim'] < primeira_entrada].iterrows():
                            chave_falta = (matricula_db, nome_db, turma, aula['disciplina'])
                            faltas_registradas[chave_falta] = faltas_registradas.get(chave_falta, 0) + 1
                    if pd.notna(ultima_saida):
                        for _, aula in aulas_do_dia[aulas_do_dia['hora_inicio'] > ultima_saida].iterrows():
                            chave_falta = (matricula_db, nome_db, turma, aula['disciplina'])
                            faltas_registradas[chave_falta] = faltas_registradas.get(chave_falta, 0) + 1
                            
    df_problemas = pd.DataFrame(problemas_alunos)
    return report_date, faltas_registradas, df_problemas