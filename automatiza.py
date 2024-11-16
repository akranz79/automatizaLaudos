import time
import os
import logging
from log_config import setup_logging
from gerarLaudo import gerar_laudos_pendentes
from log_config import setup_logging
from enviaEmail import enviar_email
from datetime import datetime
import json

# Obtém a data e hora atual
now = datetime.now()
current_time = now.strftime("%d/%m/%Y %H:%M:%S")

logging.info("Logging configurado com sucesso.")


def enviar_arquivos_email(config):
    os.system('python enviaEmail.py')
    logging.info("Chamada enviar_arquivos_email() : automatiza.py")


def gerar_laudos_com_retry():
    """
    Chama a função 'gerar_laudos_pendentes' com tentativas de reconexão em caso de falha,
    com intervalos progressivos de espera entre as tentativas.
    """
    attempt = 1  # Contador de tentativas

    while True:
        try:
            logging.info("Iniciando a geração de laudos com tentativas de reconexão.")
            gerar_laudos_pendentes()
            logging.info("Laudos gerados com sucesso.")
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

            logging.warning(f"Erro ao acessar Google Sheets (tentativa {attempt}): {e}")
            logging.info(f"Tentando novamente em {retry_delay} segundos...")
            time.sleep(retry_delay)
            attempt += 1


def automatizar(config):
    logging.info("Chamada automatizar() : automatiza.py")
    while True:
        # Obtém a data e hora atual
        now = datetime.now()
        logging.info("Esteira de processo[1]: obtendo data e hora atual : automatiza.py")
        current_time = now.strftime("%d/%m/%Y %H:%M:%S")


        # Chamar função para gerar laudos pendentes com lógica de retry
        logging.info("Esteira de processo[2]: chamando gerar_laudos_com_retry : automatiza.py")
        gerar_laudos_com_retry()


        # Executar criaGraficoGeraImagem.py
        logging.info("Esteira de processo[3]: chamando criaGraficoGeraImagem : automatiza.py")
        os.system('python criaGraficoGeraImagem.py')
        time.sleep(1)

        # Executar inserirImgPar42.py
        logging.info("Esteira de processo[4]: chamando inserirImgPar42 : automatiza.py")
        os.system('python inserirImgPar42.py')
        time.sleep(1)

        # Executar moveImagemParaPastaImgLaudo.py
        logging.info("Esteira de processo[5]: chamando moveImagemParaPastaImgLaudo : automatiza.py")
        os.system('python moveImagemParaPastaImgLaudo.py')
        time.sleep(2)

        # Executar removeTabelaLaudo.py
        logging.info("Esteira de processo[6]: chamando removeTabelaLaudo : automatiza.py")
        os.system('python removeTabelaLaudo.py')

        # Enviar arquivos por email
        logging.info("Esteira de processo[7]: chamando enviar_arquivos_email : automatiza.py")
        enviar_arquivos_email(config)

        # Aguardar o tempo configurado no arquivo de configuração para execução da próxima iteração do while
        logging.info("Esteira de processo[8]: aguardando tempo de loop  : automatiza.py")
        time.sleep(config['tempo_loop'])
        print("end of shelf")




if __name__ == "__main__":
    with open('config.json') as f:
        config = json.load(f)

    # Configura o logging usando o diretório especificado no arquivo de configuração
    setup_logging(config["LOG_DIRECTORY"])

    print(" ")
    print(f"Ini: {current_time}")
    # Registrar log de início de automação
    logging.info(f"{current_time} Iniciando a execução do script de automação.")
    automatizar(config)
