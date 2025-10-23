import pandas as pd
import os
from openpyxl import load_workbook, Workbook
from .formatador_excel import formatar_relatorio_detalhado, formatar_relatorio_simples

def gerar_relatorio_faltas(dados_da_sessao, logger):
    logger("\n--- Gerando Relatório Detalhado de Faltas com Resumo Semanal ---")
    if not dados_da_sessao:
        logger("Nenhum dado novo foi processado para gerar relatório.")
        return

    sheet_name_resumo = 'Quantitativo Total da Semana'
    project_root = os.path.dirname(os.path.dirname(__file__))
    output_path = os.path.join(project_root, 'relatorios', 'relatorio_faltas_detalhado.xlsx')
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    
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
        dfs_para_resumo = [df for df in dados_existentes.values() if 'Total de Faltas' in df.columns]
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
                df_resumo.to_excel(writer, sheet_name=sheet_name_resumo, index=False)

        workbook = load_workbook(output_path)
        for sheet_name in workbook.sheetnames:
            formatar_relatorio_detalhado(workbook[sheet_name], has_status_col=(sheet_name == sheet_name_resumo))
        workbook.save(output_path)
        logger(f"Relatório detalhado salvo/atualizado com sucesso!")
    except Exception as e:
        logger(f"ERRO ao salvar relatório detalhado: {e}")

# Em modulos/gerador_relatorios.py, substitua esta função:

# Em modulos/gerador_relatorios.py, substitua apenas esta função:

def gerar_relatorio_simples(df_problemas, report_date, logger):
    logger("\n--- Gerando Relatório Simples de Frequência ---")
    if df_problemas.empty:
        logger("Nenhum problema de frequência encontrado.")
        return

    project_root = os.path.dirname(os.path.dirname(__file__))
    output_path = os.path.join(project_root, 'relatorios', f"relatorio_frequencia_{report_date.strftime('%d%m%y')}.xlsx")
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    
    dias_semana_pt = {0: 'Segunda-feira', 1: 'Terça-feira', 2: 'Quarta-feira', 3: 'Quinta-feira', 4: 'Sexta-feira', 5: 'Sábado', 6: 'Domingo'}
    dia_da_semana_pt = dias_semana_pt.get(report_date.weekday(), '')

    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    ws.title = report_date.strftime('%d-%m-%Y')

    # --- ESCRITA DOS DADOS (SIMPLIFICADA) ---
    
    # Título Principal
    ws.append([f"Relatório de Frequência - {dia_da_semana_pt}, {report_date.strftime('%d/%m/%Y')}"])
    ws.merge_cells('A1:D1')
    ws.append([]) # Linha em branco

    df_problemas.sort_values(by=['Turma', 'Nome do Aluno'], inplace=True)
    for turma, data in df_problemas.groupby('Turma'):
        # Header da Turma
        ws.append([f"TURMA: {turma}"])
        ws.merge_cells(f'A{ws.max_row}:D{ws.max_row}')
        
        # Headers das Colunas
        headers = ['Matrícula', 'Nome do Aluno', 'Problema', 'Acesso']
        ws.append(headers)
        
        # Dados dos Alunos
        for _, aluno in data.iterrows():
            ws.append([aluno['Matricula'], aluno['Nome do Aluno'], aluno['Problema'], aluno['Acesso']])
        
        ws.append([]) # Linha em branco entre as turmas

    # --- APLICAÇÃO DA FORMATAÇÃO (SEPARADA) ---
    
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
    
    # Formata Título
    ws['A1'].font = title_font
    ws['A1'].fill = title_fill
    ws['A1'].alignment = Alignment(horizontal='center')

    # Itera sobre todas as linhas para aplicar estilos e bordas
    for row_idx, row in enumerate(ws.iter_rows(), 1):
        cell_A = ws[f'A{row_idx}']
        
        # Formata header de turma
        if cell_A.value and str(cell_A.value).startswith("TURMA:"):
            cell_A.font = turma_font
            cell_A.fill = turma_fill
            for cell in row: cell.border = thin_border
            continue

        # Formata header de colunas
        if cell_A.value == "Matrícula":
            for cell in row:
                cell.font = header_font
                cell.border = thin_border
            continue

        # Formata linhas de dados
        if isinstance(cell_A.value, (int, str)) and len(str(cell_A.value)) > 5: # Assumindo que matrícula tem mais de 5 dígitos
            problema_cell_value = ws[f'C{row_idx}'].value
            fill_color = None
            if problema_cell_value == 'FALTOU': fill_color = fill_faltou
            elif problema_cell_value == 'CHEGOU ATRASADO': fill_color = fill_atrasado
            elif problema_cell_value == 'SAIU CEDO': fill_color = fill_saiu_cedo
            
            for cell in row:
                cell.border = thin_border
                if fill_color:
                    cell.fill = fill_color
    
    # Ajuste de colunas
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