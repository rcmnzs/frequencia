import fitz  # PyMuPDF
import pandas as pd
import os
import re

def extrair_dados_frequencia(caminho_pdf):
    """
    Extrai registros de frequência de um PDF, com base na estrutura de texto
    real identificada durante a depuração.
    """
    if not os.path.exists(caminho_pdf):
        print(f"Erro: O arquivo '{caminho_pdf}' não foi encontrado.")
        return None

    try:
        doc = fitz.open(caminho_pdf)
        full_text = ""
        for page in doc:
            # Extrai o texto preservando um layout básico para ajudar na análise
            full_text += page.get_text("text") 
        doc.close()

        # Usamos "Total de Acessos do Pedestre:" como um separador mais confiável entre os blocos de alunos
        blocos_alunos = full_text.split('Total de Acessos do Pedestre:')[0:-1] # Ignora o último split que é o final do doc
        
        if not blocos_alunos:
            print("Não foi possível encontrar nenhum bloco de dados de aluno no PDF.")
            return None

        todos_acessos = []

        for bloco in blocos_alunos:
            # --- EXPRESSÃO REGULAR FINAL E CORRIGIDA ---
            # Procura por: "Nome:", quebra de linha, um grupo de números (Crachá), quebra de linha, e o resto da linha (Nome)
            padrao_aluno = re.compile(r'Nome:\n(\d+)\n(.*?)\n', re.DOTALL)
            match_aluno = padrao_aluno.search(bloco)

            if not match_aluno:
                continue

            cracha = match_aluno.group(1).strip()
            nome = match_aluno.group(2).strip()

            # Padrão para encontrar TODAS as linhas de acesso (Hora e Sentido)
            # Ele procura por uma data, depois uma hora, depois Entrada ou Saída
            padrao_acesso = re.compile(r'\d{2}/\d{2}/\d{4}\s+(\d{2}:\d{2}:\d{2})\s+(Entrada|Saída)')
            acessos = padrao_acesso.findall(bloco)

            for hora, sentido in acessos:
                registro = {
                    'Crachá': cracha,
                    'Nome': nome,
                    'Hora': hora,
                    'Sentido': sentido
                }
                todos_acessos.append(registro)
        
        if not todos_acessos:
            print("Nenhum registro de acesso foi extraído.")
            return None

        df_frequencia = pd.DataFrame(todos_acessos)
        
        return df_frequencia

    except Exception as e:
        print(f"Ocorreu um erro inesperado durante a extração: {e}")
        return None

# --- Bloco de Execução Principal ---
if __name__ == "__main__":
    nome_arquivo_pdf = "frequencia_131025.pdf" 

    print(f"Iniciando a extração de dados do arquivo '{nome_arquivo_pdf}'...")
    dados_extraidos = extrair_dados_frequencia(nome_arquivo_pdf)

    if dados_extraidos is not None and not dados_extraidos.empty:
        print("\n--- Dados de Frequência Extraídos com Sucesso ---")
        print(dados_extraidos.to_string()) 
        print(f"\nTotal de registros de acesso extraídos: {len(dados_extraidos)}")

        try:
            nome_arquivo_csv = 'frequencia_alunos.csv'
            dados_extraidos.to_csv(nome_arquivo_csv, index=False, encoding='utf-8')
            print(f"\nOs dados também foram salvos no arquivo '{nome_arquivo_csv}'.")
        except Exception as e:
            print(f"\nNão foi possível salvar o arquivo CSV: {e}")