import time
import os
from gerarLaudo import gerar_laudos_pendentes
from enviaEmail import enviar_email
from datetime import datetime
from logger import gerar_log
import json


def enviar_arquivos_email(config):
    os.system('python enviaEmail.py')


def gerar_laudos_com_retry():
    """
    Chama a função 'gerar_laudos_pendentes' com tentativas de reconexão em caso de falha,
    com intervalos progressivos de espera entre as tentativas.
    """
    attempt = 1  # Contador de tentativas

    while True:
        try:
            gerar_laudos_pendentes()
            break  # Sai do loop se a chamada for bem-sucedida
        except Exception as e:
            # Define o intervalo de retry conforme a tentativa
            if attempt == 1:
                retry_delay = 10
            elif attempt == 2:
                retry_delay = 20
            elif attempt == 3:
                retry_delay = 30
            else:
                retry_delay = 50

            print(f"Erro ao acessar Google Sheets (tentativa {attempt}): {e}")
            print(f"Tentando novamente em {retry_delay} segundos...")
            time.sleep(retry_delay)
            attempt += 1


def automatizar(config):
    while True:
        # Obtém a data e hora atual
        now = datetime.now()
        current_time = now.strftime("%d/%m/%Y %H:%M:%S")

        # Printar mensagem
        print("                                                  ")
        log4 = f"Automatizando geração de laudos - ({current_time})"
        print("                                                  ")
        gerar_log(log4, config)

        # Chamar função para gerar laudos pendentes com lógica de retry
        gerar_laudos_com_retry()

        # Aguardar 5 segundos
        time.sleep(2)

        # Chamar função do arquivo criaGraficoGeraImagem.py
        os.system('python criaGraficoGeraImagem.py')
        time.sleep(1)

        # Chamar a função para inserir a imagem criada no parágrafo 42 do laudo
        os.system('python inserirImgPar42.py')
        time.sleep(1)

        # Aguardar 5 segundos
        time.sleep(1)

        # Chamar função do arquivo moveImagemParaPastaImgLaudo.py
        os.system('python moveImagemParaPastaImgLaudo.py')

        # Aguardar 5 segundos
        time.sleep(2)

        # Chamar função do arquivo removeTabelaLaudo.py
        os.system('python removeTabelaLaudo.py')

        # Printar mensagem
        print(f"Aguardando {config['tempo_loop'] / 60} minutos para próxima verificação.")

        # Chamar função para enviar arquivos por email, passando o argumento config
        enviar_arquivos_email(config)

        # Aguardar o tempo configurado no arquivo de configuração para execução da próxima iteração do while
        time.sleep(config['tempo_loop'])


if __name__ == "__main__":
    with open('config.json') as f:
        config = json.load(f)
    automatizar(config)
