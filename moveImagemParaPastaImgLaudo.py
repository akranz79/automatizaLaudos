import os
import shutil
import json

# Carregar configurações do arquivo JSON
with open('config.json', 'r') as config_file:
    config = json.load(config_file)

# Definir diretórios de origem e destino
diretorio_origem = config["diretorio_laudos"]
diretorio_destino = os.path.join(diretorio_origem, 'imgLaudos')

def mover_arquivos_png():
    # Criar o diretório de destino se não existir
    if not os.path.exists(diretorio_destino):
        os.makedirs(diretorio_destino)

    # Listar todos os arquivos na pasta de origem
    for arquivo in os.listdir(diretorio_origem):
        if arquivo.endswith('.png'):
            caminho_arquivo_origem = os.path.join(diretorio_origem, arquivo)
            caminho_arquivo_destino = os.path.join(diretorio_destino, arquivo)

            # Mover o arquivo para o diretório de destino
            shutil.move(caminho_arquivo_origem, caminho_arquivo_destino)
            print(f"Arquivo {arquivo} movido para {diretorio_destino}.[4]")

if __name__ == "__main__":
    mover_arquivos_png()
