import fitz  # O nome da biblioteca PyMuPDF é 'fitz'
import pandas as pd
import os
import re

# Em modulos/extrator_ausentes.py, substitua esta função:

def extrair_dados_ausentes(caminho_pdf):
    """
    Extrai os dados de Matrícula e Nome de um PDF de ausentes e também a data do relatório,
    com verificação explícita de falha na extração.
    """
    from datetime import datetime
    import re
    import os
    import fitz
    import pandas as pd

    nome_arquivo = os.path.basename(caminho_pdf)

    if not os.path.exists(caminho_pdf):
        print(f"Erro: O arquivo '{nome_arquivo}' não foi encontrado.")
        return None, None

    try:
        doc = fitz.open(caminho_pdf)
        full_text = "".join(page.get_text() for page in doc)
        doc.close()

        # Extrai a data do relatório
        report_date = None
        match_data = re.search(r"Período: de (\d{2}/\d{2}/\d{4})", full_text)
        
        # --- MELHORIA DE ROBUSTEZ AQUI ---
        if not match_data:
            print(f"AVISO: Não foi possível encontrar o padrão de DATA no arquivo '{nome_arquivo}'. O formato do relatório pode ter mudado.")
            # Continuamos a extração de alunos, mas a data será None

        else:
            report_date = datetime.strptime(match_data.group(1), '%d/%m/%Y')

        # Extrai os dados dos alunos
        pattern = re.compile(r'(\d+)\s+(.*?)\s+ALUNO')
        matches = pattern.findall(full_text)
        
        # --- MELHORIA DE ROBUSTEZ AQUI ---
        if not matches:
            print(f"AVISO: Nenhum ALUNO encontrado no arquivo '{nome_arquivo}'. O formato do relatório pode ter mudado.")
            return None, report_date # Retorna a data (se encontrada) e um DataFrame vazio

        df_ausentes = pd.DataFrame(matches, columns=['Matrícula', 'Nome'])
        df_ausentes['Nome'] = df_ausentes['Nome'].str.strip()

        return df_ausentes, report_date

    except Exception as e:
        print(f"Ocorreu um erro inesperado durante a extração de ausentes: {e}")
        return None, None