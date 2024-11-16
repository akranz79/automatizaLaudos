import os
import json
import logging
import logging
from log_config import setup_logging

def gerar_laudo():
    logging.info("Chamada gerar_laudo() : mayn.py")
    os.system('python gerarLaudo.py')

def verificar_laudos_nao_gerados():
    logging.info("Chamada verificar_laudos_nao_gerados() : mayn.py")
    os.system('python dadosNaoGerados.py')

def enviar_arquivos_email():
    logging.info("Chamada enviar_arquivos_email() : mayn.py")
    os.system('python enviaEmail.py')

def automatizar():
    logging.info("Chamada automatizar() : mayn.py")
    os.system('python automatiza.py')  # Substitua 'automatizacao.py' pelo nome do seu script de automação

def configurar():
    logging.info("Chamada configurar() : mayn.py")
    logging.info("Iniciando configuração do sistema...")
    diretorio_laudos = input("Digite o diretório onde os laudos serão salvos: ")
    tempo_loop = int(input("Digite o tempo de execução do loop no script automatiza.py (em segundos): "))
    email_destino = input("Digite o endereço de email a ser configurado para envio dos laudos: ")

    config = {
        "diretorio_laudos": diretorio_laudos,
        "tempo_loop": tempo_loop,
        "email_destino": email_destino
    }

    with open(CONFIG_FILE, 'w') as f:
        json.dump(config, f)
    logging.info(f"Configuração salva: {config}")

def sair():
    logging.info("Encerrando o programa.")
    exit()

def main():
    options = {
        "1": gerar_laudo,
        "2": verificar_laudos_nao_gerados,
        "3": enviar_arquivos_email,
        "4": automatizar,
        "5": configurar,
        "0": sair
    }

    try:
        while True:
            print("\nSelecione uma opção:")
            print("1. Gerar laudo")
            print("2. Verificar laudos não gerados")
            print("3. Enviar arquivos por email")
            print("4. Automatizar geração de laudos e envio por email")
            print("5. Configurar")
            print("0. Sair")

            opcao = input("Opção: ")

            if opcao in options:
                logging.info(f"Opção selecionada: {opcao}")
                options[opcao]()
            else:
                logging.warning("Opção inválida selecionada.")
                print("Opção inválida. Tente novamente.")
    except KeyboardInterrupt:
        logging.warning("Execução interrompida pelo usuário (CTRL + C).")
        print("\nExecução interrompida pelo usuário.")

if __name__ == "__main__":
    logging.info("Chamada Main : Iniciando o programa.")
    main()
