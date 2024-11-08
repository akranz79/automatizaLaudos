import os
import json
from docx import Document

# Carregar configurações do arquivo JSON
with open('config.json', 'r') as config_file:
    config = json.load(config_file)

# Definir diretório dos laudos a partir do arquivo de configuração
diretorio_laudos = config["diretorio_laudos"]

def remover_tabela(arquivo_docx):
    # Abrir o arquivo docx
    doc = Document(arquivo_docx)

    # Remover todas as tabelas do documento
    for table in doc.tables:
        table._element.clear()

    # Salvar o arquivo sem a tabela
    doc.save(arquivo_docx)

def varrer_diretorio_laudos():
    # Listar todos os arquivos na pasta de laudos
    for arquivo in os.listdir(diretorio_laudos):
        # Verificar se é um arquivo .docx e não é o modelo.docx
        if arquivo.endswith('.docx') and arquivo != 'modelo.docx':
            caminho_arquivo_docx = os.path.join(diretorio_laudos, arquivo)
            remover_tabela(caminho_arquivo_docx)
            print(f"Tabela removida do arquivo: {caminho_arquivo_docx}.[5]")

if __name__ == "__main__":
    varrer_diretorio_laudos()
