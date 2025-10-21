import fitz  # PyMuPDF
import pandas as pd
import os
import re
from datetime import datetime

def extrair_dados_frequencia(caminho_pdf):
    """
    Extrai registros de frequência de um PDF, com lógica aprimorada para lidar
    com múltiplos formatos de layout de texto.
    """
    nome_arquivo = os.path.basename(caminho_pdf)

    if not os.path.exists(caminho_pdf):
        print(f"Erro: O arquivo '{nome_arquivo}' não foi encontrado.")
        return None, None

    try:
        doc = fitz.open(caminho_pdf)
        full_text = "".join(page.get_text("text") for page in doc)
        doc.close()

        report_date = None
        match_data = re.search(r"Período: de (\d{2}/\d{2}/\d{4})", full_text)
        if not match_data:
            print(f"AVISO: Não foi possível encontrar o padrão de DATA no arquivo '{nome_arquivo}'.")
        else:
            report_date = datetime.strptime(match_data.group(1), '%d/%m/%Y')

        blocos_alunos = full_text.split('Total de Acessos do Pedestre:')[0:-1]
        if not blocos_alunos:
            print(f"AVISO: Nenhum BLOCO DE ALUNO encontrado em '{nome_arquivo}'. O formato pode ter mudado.")
            return None, report_date

        todos_acessos = []

        # --- INÍCIO DA CORREÇÃO DE LÓGICA ---
        # Define os dois padrões de regex que conhecemos
        padrao_novo = re.compile(r"Crachá:\s*(\d+)\s+Nome:\s*(.*?)\n", re.DOTALL)
        padrao_antigo = re.compile(r'Nome:\n(\d+)\n(.*?)\n', re.DOTALL)

        for bloco in blocos_alunos:
            cracha, nome = None, None
            
            # Tenta encontrar o padrão novo primeiro
            match = padrao_novo.search(bloco)
            if match:
                cracha, nome = match.groups()
            else:
                # Se falhar, tenta encontrar o padrão antigo
                match = padrao_antigo.search(bloco)
                if match:
                    cracha, nome = match.groups()

            # Se nenhum dos padrões encontrou uma correspondência, pula este bloco
            if not cracha or not nome:
                continue

            cracha = cracha.strip()
            nome = nome.strip()

            padrao_acesso = re.compile(r'\d{2}/\d{2}/\d{4}\s+(\d{2}:\d{2}:\d{2})\s+(Entrada|Saída)')
            acessos = padrao_acesso.findall(bloco)
            
            for hora, sentido in acessos:
                todos_acessos.append({'Crachá': cracha, 'Nome': nome, 'Hora': hora, 'Sentido': sentido})
        # --- FIM DA CORREÇÃO DE LÓGICA ---
        
        if not todos_acessos:
            print(f"AVISO: Foram encontrados blocos de alunos em '{nome_arquivo}', mas nenhum registro de acesso individual foi extraído.")
            return None, report_date

        df_frequencia = pd.DataFrame(todos_acessos)
        return df_frequencia, report_date

    except Exception as e:
        print(f"Ocorreu um erro inesperado durante a extração de frequência: {e}")
        return None, None