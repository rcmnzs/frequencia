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

def processar_dados_diarios(ausentes_path, frequencia_path, logger, 
                            filtro_ativo=False, hora_inicio="00:00", hora_fim="23:59"):
    from modulos.extrator_ausentes import extrair_dados_ausentes
    from modulos.extrator_frequencias import extrair_dados_frequencia
    from datetime import time

    # Validação dos horários de entrada
    try:
        if filtro_ativo:
            hora_inicio_obj = datetime.strptime(hora_inicio, '%H:%M').time()
            hora_fim_obj = datetime.strptime(hora_fim, '%H:%M').time()
            logger(f"Filtro de horário ativado: das {hora_inicio} às {hora_fim}.")
    except ValueError:
        logger("ERRO: Formato de hora inválido no filtro. Use HH:MM. Processamento abortado.")
        return None, None, None

    df_alunos, df_horarios = carregar_dados_base(logger)
    if df_alunos is None or df_horarios is None:
        return None, None, None

    faltas_registradas = {}
    problemas_alunos = []

    # --- Processa Ausentes ---
    logger("\n--- Processando Relatório de Ausentes ---")
    df_ausentes, report_date = extrair_dados_ausentes(ausentes_path)
    # ... (lógica de extração de data e dia da semana inalterada) ...
    if report_date is None:
        raise ErroExtracaoDados("A data não foi encontrada no PDF de ausentes.")
    dia_numero = report_date.weekday()
    dias_semana_map = {0: 'SEGUNDA-FEIRA', 1: 'TERÇA-FEIRA', 2: 'QUARTA-FEIRA', 3: 'QUINTA-FEIRA', 4: 'SEXTA-FEIRA', 5: 'SÁBADO', 6: 'DOMINGO'}
    dia_semana = dias_semana_map.get(dia_numero)
    
    if df_ausentes is not None and not df_ausentes.empty:
        for _, row in df_ausentes.iterrows():
            info_aluno = buscar_aluno(df_alunos, matricula_pdf=row['Matrícula'], nome_pdf=row['Nome'], logger=logger)
            if info_aluno is not None:
                turma, nome_db, matricula_db = info_aluno['turma'], info_aluno['nome'], info_aluno['matricula']
                problemas_alunos.append({'Matricula': matricula_db, 'Nome do Aluno': nome_db, 'Turma': turma, 'Problema': 'FALTOU', 'Acesso': 'Sem registro'})
                
                aulas_do_dia_bruto = df_horarios[(df_horarios['turma'] == turma) & (df_horarios['dia_semana'] == dia_semana)]
                
                # --- APLICAÇÃO DO FILTRO ---
                aulas_do_dia = aulas_do_dia_bruto
                if filtro_ativo:
                    aulas_do_dia = aulas_do_dia_bruto[
                        (aulas_do_dia_bruto['hora_inicio'] >= hora_inicio_obj) & 
                        (aulas_do_dia_bruto['hora_inicio'] < hora_fim_obj)
                    ]
                
                for _, aula in aulas_do_dia.iterrows():
                    chave_falta = (matricula_db, nome_db, turma, aula['disciplina'])
                    faltas_registradas[chave_falta] = faltas_registradas.get(chave_falta, 0) + 1

    # --- Processa Frequência ---
    logger("\n--- Processando Relatório de Frequência ---")
    df_frequencia, date_frequencia = extrair_dados_frequencia(frequencia_path)
    # ... (validação de data inalterada) ...
    if date_frequencia is None or report_date.date() != date_frequencia.date():
        raise ErroValidacao(f"As datas dos PDFs não coincidem.")

    if df_frequencia is not None and not df_frequencia.empty:
        df_frequencia['Hora'] = pd.to_datetime(df_frequencia['Hora'], format='%H:%M:%S').dt.time
        for grupo_keys, acesso_aluno_df in df_frequencia.groupby(['Crachá', 'Nome']):
            info_aluno = buscar_aluno(df_alunos, matricula_pdf=grupo_keys[0], nome_pdf=grupo_keys[1], logger=logger)
            if info_aluno is not None:
                turma, nome_db, matricula_db = info_aluno['turma'], info_aluno['nome'], info_aluno['matricula']
                
                aulas_do_dia_bruto = df_horarios[(df_horarios['turma'] == turma) & (df_horarios['dia_semana'] == dia_semana)]
                
                # --- APLICAÇÃO DO FILTRO ---
                aulas_do_dia = aulas_do_dia_bruto
                if filtro_ativo:
                    aulas_do_dia = aulas_do_dia_bruto[
                        (aulas_do_dia_bruto['hora_inicio'] >= hora_inicio_obj) & 
                        (aulas_do_dia_bruto['hora_inicio'] < hora_fim_obj)
                    ]

                # ... (resto da lógica de cálculo de problemas e faltas inalterada) ...
    
    df_problemas = pd.DataFrame(problemas_alunos)
    return report_date, faltas_registradas, df_problemas