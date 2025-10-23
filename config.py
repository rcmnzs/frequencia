import os

# --- CAMINHOS DE DIRETÓRIOS ---
# Define o diretório raiz do projeto (a pasta onde este config.py está)
PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))

# Constrói os caminhos para as outras pastas a partir da raiz
DB_DIR = os.path.join(PROJECT_ROOT, 'db')
PDF_DIR = os.path.join(PROJECT_ROOT, 'pdf')
REPORTS_DIR = os.path.join(PROJECT_ROOT, 'relatorios')

# --- CONFIGURAÇÕES DO BANCO DE DADOS ---
DB_NAME = 'unico.db'
DB_PATH = os.path.join(DB_DIR, DB_NAME)

# --- CONFIGURAÇÕES DOS NOMES DE ARQUIVOS DE RELATÓRIO ---
DETAILED_REPORT_FILENAME = 'relatorio_faltas_detalhado.xlsx'
SIMPLE_REPORT_PREFIX = 'relatorio_frequencia_'

# --- PALAVRAS-CHAVE PARA VALIDAÇÃO E EXTRAÇÃO ---
KEYWORD_AUSENTES = "Matrícula"
KEYWORD_FREQUENCIA = "Crachá:"