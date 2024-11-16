import logging
import os
from docx import Document
import logging
from log_config import setup_logging


def remover_tabela(arquivo_docx):
    logging.info("Esteira de processo[6.1]: chamando remover_tabela : removeTabelaLaudo.py")


    # Abrir o arquivo docx
    doc = Document(arquivo_docx)

    # Remover todas as tabelas do documento
    for table in doc.tables:
        table._element.clear()

    # Salvar o arquivo sem a tabela
    doc.save(arquivo_docx)

def varrer_diretorio_laudos():
    # Diretório dos laudos
    diretorio_laudos = 'D:/Laudos'

    # Listar todos os arquivos na pasta de laudos
    for arquivo in os.listdir(diretorio_laudos):
        # Verificar se é um arquivo .docx e não é o modelo.docx
        if arquivo.endswith('.docx') and arquivo != 'modelo.docx':
            caminho_arquivo_docx = os.path.join(diretorio_laudos, arquivo)
            remover_tabela(caminho_arquivo_docx)
            logging.info(f"Esteira de processo[6.2]: Tabela removida do arquivo: {caminho_arquivo_docx}.")

if __name__ == "__main__":
    varrer_diretorio_laudos()
