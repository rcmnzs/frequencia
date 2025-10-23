import pandas as pd
import os
from openpyxl import load_workbook, Workbook
import sys

# Adiciona o diretório raiz ao path para encontrar o 'config'
project_root_path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.append(project_root_path)

import config
# --- CORREÇÃO AQUI: REMOVIDA A IMPORTAÇÃO DO 'formatador_excel' ---
# A função 'formatar_relatorio_detalhado' será definida aqui mesmo.

# --- FUNÇÃO DE FORMATAÇÃO DO RELATÓRIO DETALHADO ---
def formatar_relatorio_detalhado(worksheet, has_status_col=False):
    """Aplica a formatação ao relatório detalhado e de resumo."""
    from openpyxl.styles import Font, Border, Side, PatternFill
    from openpyxl.formatting.rule import FormulaRule
    from openpyxl.worksheet.datavalidation import DataValidation
    from openpyxl.utils import get_column_letter

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
        except (ValueError, TypeError):
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
        
        worksheet.conditional_formatting.add(validation_range, FormulaRule(formula=[f'LEFT({status_col_letter}2, 8)="PENDENTE"'], fill=red_fill))
        worksheet.conditional_formatting.add(validation_range, FormulaRule(formula=[f'LEFT({status_col_letter}2, 7)="LANÇADA"'], fill=green_fill))

# --- FUNÇÃO PRINCIPAL DE GERAÇÃO DO RELATÓRIO DETALHADO ---
def gerar_relatorio_faltas(dados_da_sessao, logger):
    logger("\n--- Gerando Relatório Detalhado de Faltas com Resumo Semanal ---")
    if not dados_da_sessao:
        logger("Nenhum dado novo foi processado para gerar relatório.")
        return

    sheet_name_resumo = 'Quantitativo Total da Semana'
    output_path = os.path.join(config.REPORTS_DIR, config.DETAILED_REPORT_FILENAME)
    os.makedirs(config.REPORTS_DIR, exist_ok=True)
    
    dados_existentes = {}
    if os.path.exists(output_path):
        try:
            dados_existentes = pd.read_excel(output_path, sheet_name=None)
            if sheet_name_resumo in dados_existentes:
                del dados_existentes[sheet_name_resumo]
        except Exception as e:
            logger(f"Aviso: Não foi possível ler arquivo existente. Erro: {e}")

    for sheet_name, (report_date, faltas_dict, _) in dados_da_sessao.items():
        if faltas_dict:
            lista_faltas = [list(chave) + [valor] for chave, valor in faltas_dict.items()]
            dados_existentes[sheet_name] = pd.DataFrame(lista_faltas, columns=['Matricula', 'Nome', 'Turma', 'Disciplina', 'Total de Faltas'])
    
    if dados_existentes:
        dfs_para_resumo = [df for df in dados_existentes.values() if 'Total de Faltas' in df.columns and not df.empty]
        if dfs_para_resumo:
            combined_df = pd.concat(dfs_para_resumo, ignore_index=True)
            df_resumo = combined_df.groupby(['Matricula', 'Nome', 'Turma', 'Disciplina'])['Total de Faltas'].sum().reset_index()
            df_resumo.rename(columns={'Total de Faltas': 'Total na Semana'}, inplace=True)
            df_resumo['STATUS'] = 'PENDENTE'

    try:
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            for sheet_name in sorted(dados_existentes.keys()):
                df = dados_existentes[sheet_name]
                df.sort_values(by=['Turma', 'Nome', 'Disciplina'], inplace=True)
                df.to_excel(writer, sheet_name=sheet_name, index=False)
            if 'df_resumo' in locals() and not df_resumo.empty:
                df_resumo.sort_values(by=['Turma', 'Nome', 'Disciplina'], inplace=True)
                df_resumo.to_excel(writer, sheet_name=sheet_name_resumo, index=False)

        workbook = load_workbook(output_path)
        for sheet_name in workbook.sheetnames:
            formatar_relatorio_detalhado(workbook[sheet_name], has_status_col=(sheet_name == sheet_name_resumo))
        workbook.save(output_path)
        logger(f"Relatório detalhado salvo/atualizado com sucesso!")
    except Exception as e:
        logger(f"ERRO ao salvar relatório detalhado: {e}")

# Em modulos/gerador_relatorios.py, substitua apenas esta função:

def gerar_relatorio_simples(df_problemas, report_date, logger):
    logger("\n--- Gerando Relatório Simples de Frequência ---")
    if df_problemas.empty:
        logger("Nenhum problema de frequência encontrado.")
        return

    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    from openpyxl import Workbook
    
    file_name = f"{config.SIMPLE_REPORT_PREFIX}{report_date.strftime('%d%m%y')}.xlsx"
    output_path = os.path.join(config.REPORTS_DIR, file_name)
    os.makedirs(config.REPORTS_DIR, exist_ok=True)
    
    dias_semana_pt = {0: 'Segunda-feira', 1: 'Terça-feira', 2: 'Quarta-feira', 3: 'Quinta-feira', 4: 'Sexta-feira', 5: 'Sábado', 6: 'Domingo'}
    dia_da_semana_pt = dias_semana_pt.get(report_date.weekday(), '')

    wb = Workbook()
    ws = wb.active
    ws.title = report_date.strftime('%d-%m-%Y')
    
    # Estilos
    title_font = Font(name='Calibri', size=14, bold=True, color="FFFFFF")
    title_fill = PatternFill(start_color="008000", end_color="008000", fill_type="solid")
    turma_font = Font(name='Calibri', size=12, bold=True)
    turma_fill = PatternFill(start_color="E7E6E6", end_color="E7E6E6", fill_type="solid")
    header_font = Font(name='Calibri', size=11, bold=True)
    fill_faltou = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    fill_atrasado = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
    fill_saiu_cedo = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    
    # Título Principal
    cell_titulo = ws['A1']
    cell_titulo.value = f"Relatório de Frequência - {dia_da_semana_pt}, {report_date.strftime('%d/%m/%Y')}"
    ws.merge_cells('A1:D1')
    cell_titulo.font = title_font
    cell_titulo.fill = title_fill
    cell_titulo.alignment = Alignment(horizontal='center')
    # Aplica borda em todas as células da linha do título após mesclar
    for i in range(1, 5): ws.cell(row=1, column=i).border = thin_border
    
    current_row = 3
    df_problemas.sort_values(by=['Turma', 'Nome do Aluno'], inplace=True)
    
    todas_as_turmas_hardcoded = ['EF1601', 'EF1602', 'EF1701', 'EF1702', 'EF1801', 'EF1802', 'EF1803', 'EF1901', 'EF1902', 'EM2101', 'EM2102', 'EM2201', 'EM2202', 'EM2301', 'EM2302']
    
    for turma in todas_as_turmas_hardcoded:
        # Header da Turma
        cell_turma = ws.cell(row=current_row, column=1, value=f"TURMA: {turma}")
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=4)
        cell_turma.font = turma_font
        cell_turma.fill = turma_fill
        # Aplica borda em todas as células da linha da turma
        for i in range(1, 5): ws.cell(row=current_row, column=i).border = thin_border
        current_row += 1
        
        data = df_problemas[df_problemas['Turma'] == turma]
        
        if not data.empty:
            headers = ['Matrícula', 'Nome do Aluno', 'Problema', 'Acesso']
            for col_num, header_text in enumerate(headers, 1):
                cell = ws.cell(row=current_row, column=col_num, value=header_text)
                cell.font = header_font
                cell.border = thin_border
            current_row += 1
            
            for _, aluno in data.iterrows():
                row_data = [aluno['Matricula'], aluno['Nome do Aluno'], aluno['Problema'], aluno['Acesso']]
                fill_color = None
                if aluno['Problema'] == 'FALTOU': fill_color = fill_faltou
                elif aluno['Problema'] == 'CHEGOU ATRASADO': fill_color = fill_atrasado
                elif aluno['Problema'] == 'SAIU CEDO': fill_color = fill_saiu_cedo

                for col_num, cell_value in enumerate(row_data, 1):
                    cell = ws.cell(row=current_row, column=col_num, value=cell_value)
                    cell.border = thin_border
                    if fill_color:
                        cell.fill = fill_color
                current_row += 1
        
        # Deixa uma linha em branco entre os blocos
        current_row += 1

    # Ajuste final de colunas
    for col_idx, col in enumerate(ws.columns, 1):
        max_length = 0
        column_letter = get_column_letter(col_idx)
        for cell in col:
            if cell.coordinate in ws.merged_cells: continue
            if cell.value:
                try: max_length = max(max_length, len(str(cell.value)))
                except: pass
        adjusted_width = (max_length + 4) if col_idx == 2 else (max_length + 2)
        ws.column_dimensions[column_letter].width = adjusted_width
    
    try:
        wb.save(output_path)
        logger(f"Relatório simples salvo com sucesso!")
    except Exception as e:
        logger(f"ERRO ao salvar relatório simples: {e}")