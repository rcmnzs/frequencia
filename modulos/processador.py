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
    
    if report_date is None:
        logger("ERRO: A data não foi encontrada no PDF de ausentes.")
        return None, None, None
        
    dia_numero = report_date.weekday()
    dias_semana_map = {0: 'SEGUNDA-FEIRA', 1: 'TERÇA-FEIRA', 2: 'QUARTA-FEIRA', 
                       3: 'QUINTA-FEIRA', 4: 'SEXTA-FEIRA', 5: 'SÁBADO', 6: 'DOMINGO'}
    dia_semana = dias_semana_map.get(dia_numero)
    
    logger(f"Data do relatório: {report_date.strftime('%d/%m/%Y')} ({dia_semana})")
    
    if df_ausentes is not None and not df_ausentes.empty:
        logger(f"Total de alunos ausentes encontrados: {len(df_ausentes)}")
        for _, row in df_ausentes.iterrows():
            info_aluno = buscar_aluno(df_alunos, matricula_pdf=row['Matrícula'], nome_pdf=row['Nome'], logger=logger)
            if info_aluno is not None:
                turma, nome_db, matricula_db = info_aluno['turma'], info_aluno['nome'], info_aluno['matricula']
                problemas_alunos.append({
                    'Matricula': matricula_db, 
                    'Nome do Aluno': nome_db, 
                    'Turma': turma, 
                    'Problema': 'FALTOU', 
                    'Acesso': 'Sem registro'
                })
                
                aulas_do_dia_bruto = df_horarios[(df_horarios['turma'] == turma) & (df_horarios['dia_semana'] == dia_semana)]
                
                # Aplicação do filtro
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
    
    if date_frequencia is None:
        logger("ERRO: A data não foi encontrada no PDF de frequência.")
        return None, None, None
        
    if report_date.date() != date_frequencia.date():
        logger(f"ERRO: As datas dos PDFs não coincidem. Ausentes: {report_date.date()}, Frequência: {date_frequencia.date()}")
        return None, None, None

    if df_frequencia is not None and not df_frequencia.empty:
        logger(f"Total de registros de frequência encontrados: {len(df_frequencia)}")
        df_frequencia['Hora'] = pd.to_datetime(df_frequencia['Hora'], format='%H:%M:%S').dt.time
        
        for grupo_keys, acesso_aluno_df in df_frequencia.groupby(['Crachá', 'Nome']):
            info_aluno = buscar_aluno(df_alunos, matricula_pdf=grupo_keys[0], nome_pdf=grupo_keys[1], logger=logger)
            if info_aluno is not None:
                turma, nome_db, matricula_db = info_aluno['turma'], info_aluno['nome'], info_aluno['matricula']
                
                aulas_do_dia_bruto = df_horarios[(df_horarios['turma'] == turma) & (df_horarios['dia_semana'] == dia_semana)]
                
                # Aplicação do filtro
                aulas_do_dia = aulas_do_dia_bruto
                if filtro_ativo:
                    aulas_do_dia = aulas_do_dia_bruto[
                        (aulas_do_dia_bruto['hora_inicio'] >= hora_inicio_obj) & 
                        (aulas_do_dia_bruto['hora_inicio'] < hora_fim_obj)
                    ]

                if aulas_do_dia.empty:
                    continue
                
                # Ordena os acessos por hora
                acesso_aluno_df = acesso_aluno_df.sort_values('Hora')
                
                # Pega primeira aula do dia
                primeira_aula = aulas_do_dia.iloc[0]
                hora_inicio_aulas = primeira_aula['hora_inicio']
                
                # Pega última aula do dia
                ultima_aula = aulas_do_dia.iloc[-1]
                hora_fim_aulas = ultima_aula['hora_fim']
                
                # Encontra primeira entrada e última saída
                entradas = acesso_aluno_df[acesso_aluno_df['Sentido'] == 'Entrada']
                saidas = acesso_aluno_df[acesso_aluno_df['Sentido'] == 'Saída']
                
                primeira_entrada = entradas.iloc[0] if not entradas.empty else None
                ultima_saida = saidas.iloc[-1] if not saidas.empty else None
                
                # Verifica atraso (tolerância de 15 minutos)
                if primeira_entrada is not None:
                    hora_entrada = primeira_entrada['Hora']
                    # Converte time para minutos para facilitar comparação
                    minutos_entrada = hora_entrada.hour * 60 + hora_entrada.minute
                    minutos_inicio = hora_inicio_aulas.hour * 60 + hora_inicio_aulas.minute
                    tolerancia_minutos = 15
                    
                    if minutos_entrada > (minutos_inicio + tolerancia_minutos):
                        problemas_alunos.append({
                            'Matricula': matricula_db,
                            'Nome do Aluno': nome_db,
                            'Turma': turma,
                            'Problema': 'CHEGOU ATRASADO',
                            'Acesso': f"Entrada: {hora_entrada.strftime('%H:%M:%S')}"
                        })
                        logger(f"  - {nome_db} chegou atrasado às {hora_entrada}")
                
                # Verifica saída antecipada (tolerância de 15 minutos antes do fim)
                if ultima_saida is not None:
                    hora_saida = ultima_saida['Hora']
                    minutos_saida = hora_saida.hour * 60 + hora_saida.minute
                    minutos_fim = hora_fim_aulas.hour * 60 + hora_fim_aulas.minute
                    tolerancia_minutos = 15
                    
                    if minutos_saida < (minutos_fim - tolerancia_minutos):
                        problemas_alunos.append({
                            'Matricula': matricula_db,
                            'Nome do Aluno': nome_db,
                            'Turma': turma,
                            'Problema': 'SAIU CEDO',
                            'Acesso': f"Saída: {hora_saida.strftime('%H:%M:%S')}"
                        })
                        logger(f"  - {nome_db} saiu cedo às {hora_saida}")
                
                # Calcula faltas em aulas específicas (lógica existente de presença)
                for _, aula in aulas_do_dia.iterrows():
                    hora_ini_aula = aula['hora_inicio']
                    hora_fim_aula = aula['hora_fim']
                    
                    # Verifica se há registro de presença durante a aula
                    presenca_na_aula = False
                    for _, acesso in acesso_aluno_df.iterrows():
                        hora_acesso = acesso['Hora']
                        if acesso['Sentido'] == 'Entrada' and hora_acesso <= hora_fim_aula:
                            # Verifica se há saída após o fim da aula ou se é a última entrada do dia
                            saidas_posteriores = acesso_aluno_df[
                                (acesso_aluno_df['Sentido'] == 'Saída') & 
                                (acesso_aluno_df['Hora'] > hora_acesso)
                            ]
                            if saidas_posteriores.empty or saidas_posteriores.iloc[0]['Hora'] >= hora_ini_aula:
                                presenca_na_aula = True
                                break
                    
                    if not presenca_na_aula:
                        chave_falta = (matricula_db, nome_db, turma, aula['disciplina'])
                        faltas_registradas[chave_falta] = faltas_registradas.get(chave_falta, 0) + 1
    
    df_problemas = pd.DataFrame(problemas_alunos)
    logger(f"\n--- Processamento Concluído ---")
    logger(f"Total de problemas detectados: {len(problemas_alunos)}")
    
    return report_date, faltas_registradas, df_problemas