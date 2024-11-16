# log_config.py
import logging
import os

def setup_logging(log_directory="D:\\Laudos\\Logs"):
    """Configura o sistema de logs com o diretório especificado."""
    if not os.path.exists(log_directory):
        os.makedirs(log_directory)
    log_path = os.path.join(log_directory, "automacao_laudos.log")

    logging.basicConfig(
        filename=log_path,
        level=logging.INFO,
        format="%(asctime)s - %(levelname)s - %(message)s"
    )

# Executa a configuração do log automaticamente ao importar o módulo
setup_logging()
