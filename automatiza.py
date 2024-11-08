import time
import os
from gerarLaudo import gerar_laudos_pendentes
from enviaEmail import enviar_email
from datetime import datetime
from logger import gerar_log
import json

def enviar_arquivos_email(config):
    os.system('python enviaEmail.py')

def automatizar(config):
    while True:
        # Obtém a data e hora atual
        now = datetime.now()
        current_time = now.strftime("%d/%m/%Y %H:%M:%S")

        # Printar mensagem
        print(f"                                                  ")
        log4 = (f"Automatizando geração de laudos - ({current_time})")
        print(f"                                                  ")
        gerar_log(log4, config)

        # Chamar função para gerar laudos pendentes
        gerar_laudos_pendentes()

        # Aguardar 5 segundos
        time.sleep(2)

        # Chamar função do arquivo criaGraficoGeraImagem.py
        os.system('python criaGraficoGeraImagem.py')
        time.sleep(1)

        # Chamar a função para inserir a imagem criada no paragrafo 42 do laudo
        os.system('inserirImgPar42.py')
        time.sleep(1)

        # Aguardar 5 segundos
        time.sleep(1)

        # Chamar função do arquivo moveImagemParaPastaImgLaudo.py
        #print("Movendo imagens para pasta imgLaudos [load...]")
        os.system('python moveImagemParaPastaImgLaudo.py')

        # Aguardar 5 segundos
        time.sleep(2)

        # Chamar função do arquivo removeTabelaLaudo.py
        os.system('python removeTabelaLaudo.py')

        # Printar mensagem
        print(f"Aguardando {config['tempo_loop']/60} minutos para próxima verificação.")

        # Chamar função para enviar arquivos por email, passando o argumento config
        enviar_arquivos_email(config)

        # Aguardar o tempo configurado no arquivo de configuração para execução da próxima iteração do while
        time.sleep(config['tempo_loop'])

if __name__ == "__main__":
    with open('config.json') as f:
        config = json.load(f)
    automatizar(config)
