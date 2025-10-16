import sqlite3
import pandas as pd
import os
import re
import fitz  # PyMuPDF
from datetime import datetime
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# ... (as funções buscar_aluno e carregar_dados_base permanecem as mesmas) ...
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

# Em modulos/logica.py, substitua apenas esta função:

def processar_dados_diarios(ausentes_path, frequencia_path, logger):
    """
    Função central que processa os PDFs e retorna os dados de faltas em memória.
    Retorna: (data_relatorio, dict_faltas_detalhado, df_problemas_simples)
    """
    df_alunos, df_horarios = carregar_dados_base(logger)
    if df_alunos is None or df_horarios is None:
        return None, None, None

    # Dicionário para o relatório detalhado (lógica inalterada)
    faltas_registradas = {}
    # Lista de dicionários para o relatório simples (lógica será diferente)
    problemas_alunos = []

    # --- Processa Ausentes ---
    logger("\n--- Processando Relatório de Ausentes ---")
    dias_semana_map = {'Monday': 'SEGUNDA-FEIRA', 'Tuesday': 'TERÇA-FEIRA', 'Wednesday': 'QUARTA-FEIRA',
                       'Thursday': 'QUINTA-FEIRA', 'Friday': 'SEXTA-FEIRA', 'Saturday': 'SÁBADO', 'Sunday': 'DOMINGO'}
    
    doc = fitz.open(ausentes_path)
    content = "".join(page.get_text() for page in doc)
    doc.close()

    match_data = re.search(r"Período: de (\d{2}/\d{2}/\d{4})", content)
    if not match_data:
        logger("ERRO: Não foi possível encontrar a data no PDF de ausentes.")
        return None, None, None
        
    report_date = datetime.strptime(match_data.group(1), '%d/%m/%Y')
    dia_semana = dias_semana_map.get(report_date.strftime('%A'))
    logger(f"Data do relatório identificada: {report_date.date()}")

    ausentes = re.findall(r'(\d{9,11})\s+([A-ZÀ-Ú\s.-]+?)\s+ALUNO', content)
    logger(f"Encontrados {len(ausentes)} alunos ausentes.")

    # Adiciona todos os ausentes a ambas as listas
    for matricula_pdf, nome_pdf in ausentes:
        info_aluno = buscar_aluno(df_alunos, matricula_pdf=matricula_pdf, nome_pdf=nome_pdf)
        if info_aluno is not None:
            turma, nome_db, matricula_db = info_aluno['turma'], info_aluno['nome'], info_aluno['matricula']
            # Adiciona à lista do relatório simples
            problemas_alunos.append({
                'Matricula': matricula_db, 'Nome do Aluno': nome_db, 'Turma': turma,
                'Problema': 'FALTOU', 'Acesso': 'Sem registro'
            })
            # Calcula faltas para o relatório detalhado
            aulas_do_dia = df_horarios[(df_horarios['turma'] == turma) & (df_horarios['dia_semana'] == dia_semana)]
            for _, aula in aulas_do_dia.iterrows():
                chave_falta = (matricula_db, nome_db, turma, aula['disciplina'])
                faltas_registradas[chave_falta] = faltas_registradas.get(chave_falta, 0) + 1
        else:
            logger(f"  Aviso: Aluno ausente '{nome_pdf.strip()}' não encontrado no BD.")

    # --- Processa Frequência (com lógica separada para cada relatório) ---
    logger("\n--- Processando Relatório de Frequência ---")
    from modulos.extrator_frequencias import extrair_dados_frequencia
    df_acessos = extrair_dados_frequencia(frequencia_path)

    if df_acessos is not None and not df_acessos.empty:
        df_acessos['Hora'] = pd.to_datetime(df_acessos['Hora'], format='%H:%M:%S').dt.time
        for grupo_keys, acesso_aluno_df in df_acessos.groupby(['Crachá', 'Nome']):
            matricula_pdf, nome_pdf = grupo_keys
            info_aluno = buscar_aluno(df_alunos, matricula_pdf=matricula_pdf, nome_pdf=nome_pdf)
            if info_aluno is None:
                logger(f"  Aviso: Aluno presente '{nome_pdf.strip()}' não encontrado no BD.")
                continue

            turma, nome_db, matricula_db = info_aluno['turma'], info_aluno['nome'], info_aluno['matricula']
            aulas_do_dia = df_horarios[(df_horarios['turma'] == turma) & (df_horarios['dia_semana'] == dia_semana)]
            primeira_entrada = acesso_aluno_df[acesso_aluno_df['Sentido'] == 'Entrada']['Hora'].min()
            ultima_saida = acesso_aluno_df[acesso_aluno_df['Sentido'] == 'Saída']['Hora'].max()
            
            # --- LÓGICA PARA RELATÓRIO SIMPLES ---
            if not aulas_do_dia.empty:
                primeira_aula_do_dia = aulas_do_dia['hora_inicio'].min()
                ultima_aula_do_dia = aulas_do_dia['hora_fim'].max()

                if pd.notna(primeira_entrada) and primeira_entrada > primeira_aula_do_dia:
                    problemas_alunos.append({
                        'Matricula': matricula_db, 'Nome do Aluno': nome_db, 'Turma': turma,
                        'Problema': 'CHEGOU ATRASADO', 'Acesso': f"Entrada: {primeira_entrada.strftime('%H:%M:%S')}"
                    })
                
                if pd.notna(ultima_saida) and ultima_saida < ultima_aula_do_dia:
                    problemas_alunos.append({
                        'Matricula': matricula_db, 'Nome do Aluno': nome_db, 'Turma': turma,
                        'Problema': 'SAIU CEDO', 'Acesso': f"Saída: {ultima_saida.strftime('%H:%M:%S')}"
                    })

            # --- LÓGICA PARA RELATÓRIO DETALHADO (inalterada) ---
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


# Em modulos/logica.py, substitua apenas esta função:

# Em modulos/logica.py, substitua apenas esta função:

def gerar_relatorio_faltas(faltas_registradas, report_date, logger):
    """
    Atualiza o arquivo de relatório detalhado, adicionando/substituindo a aba do dia
    e criando/atualizando uma aba de resumo semanal com a coluna de STATUS.
    """
    # Importações necessárias para formatação
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.formatting.rule import FormulaRule
    from openpyxl.worksheet.datavalidation import DataValidation
    from openpyxl.utils import get_column_letter
    from openpyxl import load_workbook

    logger("\n--- Gerando Relatório Detalhado de Faltas com Resumo Semanal ---")
    if not faltas_registradas:
        logger("Nenhuma falta detalhada foi registrada para o dia.")
        return False

    sheet_name_dia = report_date.strftime('%d-%m-%Y')
    sheet_name_resumo = 'Quantitativo Total da Semana'
    script_dir = os.path.dirname(__file__)
    project_root = os.path.dirname(script_dir)
    relatorios_dir = os.path.join(project_root, 'relatorios')
    os.makedirs(relatorios_dir, exist_ok=True)
    output_path = os.path.join(relatorios_dir, 'relatorio_faltas_detalhado.xlsx')

    # Prepara o DataFrame do dia atual, SEM a coluna de STATUS
    lista_faltas_dia = [list(chave) + [valor] for chave, valor in faltas_registradas.items()]
    df_dia_atual = pd.DataFrame(lista_faltas_dia, columns=['Matricula', 'Nome', 'Turma', 'Disciplina', 'Total de Faltas'])
    # --- MUDANÇA 1: Linha removida daqui ---
    # df_dia_atual['STATUS'] = 'PENDENTE' 

    sheets_data = {}

    if os.path.exists(output_path):
        try:
            existing_sheets = pd.read_excel(output_path, sheet_name=None)
            for name, df in existing_sheets.items():
                if name != sheet_name_resumo:
                    sheets_data[name] = df
        except Exception as e:
            logger(f"Aviso: Não foi possível ler o arquivo Excel existente. Ele será sobrescrito. Erro: {e}")

    sheets_data[sheet_name_dia] = df_dia_atual
    
    if sheets_data:
        # Precisamos garantir que a coluna 'STATUS' não interfira no cálculo de concatenação
        # Criamos uma lista de dataframes sem a coluna 'STATUS' para o cálculo do resumo
        dfs_para_resumo = []
        for df in sheets_data.values():
            # Remove a coluna 'STATUS' se ela existir (de leituras anteriores)
            if 'STATUS' in df.columns:
                dfs_para_resumo.append(df.drop(columns=['STATUS']))
            else:
                dfs_para_resumo.append(df)

        combined_df = pd.concat(dfs_para_resumo, ignore_index=True)
        df_resumo = combined_df.groupby(['Matricula', 'Nome', 'Turma', 'Disciplina'])['Total de Faltas'].sum().reset_index()
        df_resumo.rename(columns={'Total de Faltas': 'Total na Semana'}, inplace=True)
        df_resumo.sort_values(by=['Turma', 'Nome', 'Disciplina'], inplace=True)
        
        # --- MUDANÇA 2: Adiciona a coluna 'STATUS' APENAS no DataFrame de resumo ---
        df_resumo['STATUS'] = 'PENDENTE'

    try:
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            for sheet_name, df in sorted(sheets_data.items()):
                df.sort_values(by=['Turma', 'Nome', 'Disciplina'], inplace=True)
                df.to_excel(writer, sheet_name=sheet_name, index=False)
            
            if 'df_resumo' in locals() and not df_resumo.empty:
                df_resumo.to_excel(writer, sheet_name=sheet_name_resumo, index=False)

        workbook = load_workbook(output_path)

        def formatar_aba(worksheet, has_status_col=False):
            # ... (código interno da função de formatação permanece o mesmo) ...
            header_font = Font(bold=True)
            thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
            
            for row in worksheet.iter_rows(min_row=1, max_row=worksheet.max_row, min_col=1, max_col=worksheet.max_column):
                for cell in row:
                    cell.border = thin_border
                    if cell.row == 1:
                        cell.font = header_font
            
            for col_idx, col in enumerate(worksheet.columns, 1):
                try:
                    max_length = max(len(str(cell.value)) for cell in col if cell.value)
                    worksheet.column_dimensions[get_column_letter(col_idx)].width = max_length + 2
                except ValueError:
                    pass
            
            worksheet.auto_filter.ref = worksheet.dimensions

            if has_status_col:
                status_col_letter = get_column_letter(worksheet.max_column)
                dv = DataValidation(type="list", formula1='"PENDENTE,LANÇADA"', allow_blank=True)
                worksheet.add_data_validation(dv)
                validation_range = f'{status_col_letter}2:{status_col_letter}{worksheet.max_row}'
                dv.add(validation_range)

                red_fill = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
                green_fill = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')
                
                worksheet.conditional_formatting.add(validation_range,
                    FormulaRule(formula=[f'LEFT({status_col_letter}2, 8)="PENDENTE"'], fill=red_fill))
                worksheet.conditional_formatting.add(validation_range,
                    FormulaRule(formula=[f'LEFT({status_col_letter}2, 7)="LANÇADA"'], fill=green_fill))
        
        # --- MUDANÇA 3: Inverte a lógica de chamada da formatação ---
        for sheet_name in workbook.sheetnames:
            if sheet_name == sheet_name_resumo:
                # Aplica a formatação completa na aba de resumo
                formatar_aba(workbook[sheet_name], has_status_col=True)
            else:
                # Aplica a formatação simples (sem status) nas abas diárias
                formatar_aba(workbook[sheet_name], has_status_col=False)

        workbook.save(output_path)
        logger(f"Relatório detalhado e resumo semanal salvos/atualizados com sucesso!")
        return True
    except Exception as e:
        logger(f"ERRO ao salvar e formatar o arquivo Excel: {e}")
        return False

# --- FUNÇÃO 2: GERAR RELATÓRIO SIMPLES (O NOVO) ---

def gerar_relatorio_simples(df_problemas, report_date, logger):
    logger("\n--- Gerando Relatório Simples de Frequência ---")
    if df_problemas.empty:
        logger("Nenhum problema de frequência encontrado para gerar o relatório simples.")
        return

    script_dir = os.path.dirname(__file__)
    project_root = os.path.dirname(script_dir)
    relatorios_dir = os.path.join(project_root, 'relatorios')
    os.makedirs(relatorios_dir, exist_ok=True)
    output_path = os.path.join(relatorios_dir, f"relatorio_frequencia_{report_date.strftime('%d%m%y')}.xlsx")

    dias_semana_pt = {
        'Monday': 'Segunda-feira', 'Tuesday': 'Terça-feira', 'Wednesday': 'Quarta-feira',
        'Thursday': 'Quinta-feira', 'Friday': 'Sexta-feira', 'Saturday': 'Sábado', 'Sunday': 'Domingo'
    }
    dia_da_semana_en = report_date.strftime('%A')
    dia_da_semana_pt = dias_semana_pt.get(dia_da_semana_en, dia_da_semana_en)

    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl import Workbook
    from openpyxl.utils import get_column_letter
    
    title_font = Font(name='Calibri', size=14, bold=True, color="FFFFFF")
    title_fill = PatternFill(start_color="008000", end_color="008000", fill_type="solid")
    turma_font = Font(name='Calibri', size=12, bold=True)
    turma_fill = PatternFill(start_color="E7E6E6", end_color="E7E6E6", fill_type="solid")
    header_font = Font(name='Calibri', size=11, bold=True)
    
    fill_faltou = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    fill_atrasado = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
    fill_saiu_cedo = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")

    thin_border = Border(left=Side(style='thin'), 
                         right=Side(style='thin'), 
                         top=Side(style='thin'), 
                         bottom=Side(style='thin'))

    wb = Workbook()
    ws = wb.active
    ws.title = report_date.strftime('%d-%m-%Y')

    ws['A1'] = f"Relatório de Frequência - {dia_da_semana_pt}, {report_date.strftime('%d/%m/%Y')}"
    ws.merge_cells('A1:D1')
    ws['A1'].font = title_font
    ws['A1'].fill = title_fill
    ws['A1'].alignment = Alignment(horizontal='center')
    
    current_row = 3
    
    df_problemas.sort_values(by=['Turma', 'Nome do Aluno'], inplace=True)
    for turma, data in df_problemas.groupby('Turma'):
        # --- LÓGICA DE BORDAS PARA HEADERS ---
        # Aplica borda ao header da turma
        for col_num in range(1, 5):
            ws.cell(row=current_row, column=col_num).border = thin_border
        
        cell_turma = ws[f'A{current_row}']
        cell_turma.value = f"TURMA: {turma}"
        ws.merge_cells(f'A{current_row}:D{current_row}')
        cell_turma.font = turma_font
        cell_turma.fill = turma_fill
        current_row += 1

        headers = ['Matrícula', 'Nome do Aluno', 'Problema', 'Acesso']
        for col_num, header_text in enumerate(headers, 1):
            cell = ws.cell(row=current_row, column=col_num, value=header_text)
            cell.font = header_font
            cell.border = thin_border # Aplica borda ao header das colunas
        current_row += 1

        for _, aluno in data.iterrows():
            # Aplica borda e preenchimento a todas as células da linha
            for col_num in range(1, 5):
                cell = ws.cell(row=current_row, column=col_num)
                if col_num == 1: cell.value = aluno['Matricula']
                elif col_num == 2: cell.value = aluno['Nome do Aluno']
                elif col_num == 3: cell.value = aluno['Problema']
                elif col_num == 4: cell.value = aluno['Acesso']
                
                cell.border = thin_border # Aplica a borda
                
                # Determina a cor de preenchimento
                if aluno['Problema'] == 'FALTOU': cell.fill = fill_faltou
                elif aluno['Problema'] == 'CHEGOU ATRASADO': cell.fill = fill_atrasado
                elif aluno['Problema'] == 'SAIU CEDO': cell.fill = fill_saiu_cedo
            
            current_row += 1
        
        current_row += 1

    # --- LÓGICA FINAL E CORRIGIDA DE AUTO-AJUSTE ---
    for col_idx, col in enumerate(ws.columns, 1):
        max_length = 0
        column_letter = get_column_letter(col_idx)
        for cell in col:
            # Ignora células mescladas na verificação de tamanho
            if cell.coordinate in ws.merged_cells:
                continue
            if cell.value:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
        # Adiciona um pouco de espaço extra, especialmente para a coluna B (Nome)
        adjusted_width = (max_length + 4) if col_idx == 2 else (max_length + 2)
        ws.column_dimensions[column_letter].width = adjusted_width

    try:
        wb.save(output_path)
        logger(f"Relatório simples formatado salvo com sucesso em '{output_path}'")
    except Exception as e:
        logger(f"ERRO ao salvar relatório simples formatado: {e}")