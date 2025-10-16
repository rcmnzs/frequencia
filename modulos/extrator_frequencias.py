import fitz  # PyMuPDF
import pandas as pd
import os
import re

# Em modulos/extrator_frequencias.py, substitua esta função:

def extrair_dados_frequencia(caminho_pdf):
    """
    Extrai registros de frequência de um PDF e também a data do relatório,
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
        full_text = "".join(page.get_text("text") for page in doc)
        doc.close()

        # Extrai a data do relatório
        report_date = None
        match_data = re.search(r"Período: de (\d{2}/\d{2}/\d{4})", full_text)
        
        # --- MELHORIA DE ROBUSTEZ AQUI ---
        if not match_data:
            print(f"AVISO: Não foi possível encontrar o padrão de DATA no arquivo '{nome_arquivo}'. O formato do relatório pode ter mudado.")

        else:
            report_date = datetime.strptime(match_data.group(1), '%d/%m/%Y')

        # Extrai os dados de acesso
        blocos_alunos = full_text.split('Total de Acessos do Pedestre:')[0:-1]
        
        # --- MELHORIA DE ROBUSTEZ AQUI ---
        if not blocos_alunos:
            print(f"AVISO: Nenhum BLOCO DE ALUNO ('Total de Acessos do Pedestre:') encontrado em '{nome_arquivo}'. O formato pode ter mudado.")
            return None, report_date

        todos_acessos = []
        for bloco in blocos_alunos:
            padrao_aluno = re.compile(r'Nome:\n(\d+)\n(.*?)\n', re.DOTALL)
            match_aluno = padrao_aluno.search(bloco)
            if not match_aluno:
                continue
                
            cracha, nome = match_aluno.group(1).strip(), match_aluno.group(2).strip()
            padrao_acesso = re.compile(r'\d{2}/\d{2}/\d{4}\s+(\d{2}:\d{2}:\d{2})\s+(Entrada|Saída)')
            acessos = padrao_acesso.findall(bloco)
            # Não adicionamos um aviso aqui, pois um aluno pode não ter acessos no período filtrado.
            # A falta de blocos_alunos é o erro principal a ser pego.
            
            for hora, sentido in acessos:
                todos_acessos.append({'Crachá': cracha, 'Nome': nome, 'Hora': hora, 'Sentido': sentido})
        
        if not todos_acessos:
            # Este aviso é útil se houver blocos, mas nenhum acesso extraído
            print(f"AVISO: Foram encontrados blocos de alunos em '{nome_arquivo}', mas nenhum registro de acesso individual (Entrada/Saída) foi extraído.")
            return None, report_date

        df_frequencia = pd.DataFrame(todos_acessos)
        return df_frequencia, report_date

    except Exception as e:
        print(f"Ocorreu um erro inesperado durante a extração de frequência: {e}")
        return None, None