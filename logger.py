import os
import datetime
import json
import inspect


def gerar_log(mensagem, config):
    # Obtenha o nome do arquivo chamador
    nome_arquivo_chamador = inspect.stack()[1].filename
    nome_arquivo_chamador = os.path.basename(nome_arquivo_chamador).split('.')[0]  # Remove a extensão do arquivo

    # Obtenha a data atual
    data_atual = datetime.datetime.now().strftime("%Y-%m-%d")
    # Crie o nome do arquivo de log com a data atual e o nome do arquivo chamador
    nome_arquivo_log = f"log_{nome_arquivo_chamador}_{data_atual}.txt"
    # Caminho completo para o arquivo de log
    caminho_arquivo = os.path.join(config["LOG_DIRECTORY"], nome_arquivo_log)

    # Abra o arquivo de log em modo de anexo (append)
    with open(caminho_arquivo, "a") as arquivo_log:
        # Obtenha a data e hora atual para o log
        data_hora_log = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        # Escreva a mensagem de log no arquivo
        arquivo_log.write(f"[{data_hora_log}] {mensagem}\n")


# Exemplo de uso da função gerar_log
if __name__ == "__main__":
    with open('config.json') as f:
        config = json.load(f)

    # Exemplo de chamada da função gerar_log
    gerar_log("Exemplo de mensagem de log", config)
