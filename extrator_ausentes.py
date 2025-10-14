import fitz  # O nome da biblioteca PyMuPDF é 'fitz'
import pandas as pd
import os
import re

def extrair_dados_ausentes(caminho_pdf):
    """
    Extrai os dados de Matrícula e Nome de um PDF de relatório de ausentes
    usando PyMuPDF para uma extração de texto robusta.

    Args:
        caminho_pdf (str): O caminho para o arquivo PDF.

    Returns:
        pandas.DataFrame: Um DataFrame com as colunas 'Matrícula' e 'Nome',
                          ou None se o arquivo não for encontrado ou ocorrer um erro.
    """
    if not os.path.exists(caminho_pdf):
        print(f"Erro: O arquivo '{caminho_pdf}' não foi encontrado.")
        return None

    try:
        # 1. Abre o arquivo PDF com PyMuPDF
        doc = fitz.open(caminho_pdf)
        
        full_text = ""
        # 2. Itera por todas as páginas do documento
        for page in doc:
            # Extrai todo o texto da página e adiciona à nossa string principal
            full_text += page.get_text()
            
        doc.close()

        # 3. Define o padrão (Expressão Regular) para encontrar as linhas dos alunos
        #    - Grupo 1 (\d+): Captura a Matrícula (uma ou mais sequências de números).
        #    - Grupo 2 (.*?): Captura o Nome (qualquer caractere, sem ser ganancioso).
        #    - O padrão termina quando encontra a palavra "ALUNO".
        pattern = re.compile(r'(\d+)\s+(.*?)\s+ALUNO')

        # 4. Usa re.findall para encontrar TODAS as correspondências no texto completo.
        #    Isso retorna uma lista de tuplas, onde cada tupla é (matricula, nome).
        matches = pattern.findall(full_text)

        if not matches:
            print("Não foi possível encontrar nenhum dado de aluno correspondente ao padrão no PDF.")
            return None

        # 5. Cria o DataFrame do Pandas diretamente a partir das correspondências encontradas.
        df_ausentes = pd.DataFrame(matches, columns=['Matrícula', 'Nome'])
        
        # 6. Limpa e formata o nome, removendo espaços extras no início/fim.
        df_ausentes['Nome'] = df_ausentes['Nome'].str.strip()

        return df_ausentes

    except Exception as e:
        print(f"Ocorreu um erro inesperado durante a extração: {e}")
        return None

# --- Bloco de Execução Principal ---
if __name__ == "__main__":
    nome_arquivo_pdf = "ausentes_131025.pdf"

    print(f"Iniciando a extração de dados do arquivo '{nome_arquivo_pdf}'...")
    dados_extraidos = extrair_dados_ausentes(nome_arquivo_pdf)

    if dados_extraidos is not None and not dados_extraidos.empty:
        print("\n--- Dados dos Alunos Ausentes Extraídos com Sucesso ---")
        print(dados_extraidos.to_string()) 
        print(f"\nTotal de alunos ausentes extraídos: {len(dados_extraidos)}")

        try:
            nome_arquivo_csv = 'alunos_ausentes.csv'
            dados_extraidos.to_csv(nome_arquivo_csv, index=False, encoding='utf-8')
            print(f"\nOs dados também foram salvos no arquivo '{nome_arquivo_csv}'.")
        except Exception as e:
            print(f"\nNão foi possível salvar o arquivo CSV: {e}")