import os
import json

CONFIG_FILE = "config.json"

def gerar_laudo():
    print("Gerando laudo...")
    os.system('python gerarLaudo.py')

def verificar_laudos_nao_gerados():
    print("Verificando laudos não gerados...")
    os.system('python dadosNaoGerados.py')

def enviar_arquivos_email():
    print("Enviando arquivos por email...")
    os.system('python enviaEmail.py')

def automatizar():
    print("Automatizando geração de laudos e envio por email...")
    os.system('python automatiza.py')  # Substitua 'automatizacao.py' pelo nome do seu script de automação

def configurar():
    print("\nConfiguração:")
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

def sair():
    print("Encerrando o programa.")
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
            options[opcao]()
        else:
            print("Opção inválida. Tente novamente.")

if __name__ == "__main__":
    main()
