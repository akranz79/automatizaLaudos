import logging
import os
import shutil

def mover_arquivos_png():
    logging.info("Esteira de processo[5.1]: chamando mover_arquivos_png : moveImagemParaPastaImgLaudo.py")
    diretorio_origem = 'D:/Laudos'
    diretorio_destino = 'D:/Laudos/imgLaudos'

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
            logging.info(f"Arquivo {arquivo} movido para {diretorio_destino}.")

if __name__ == "__main__":
    mover_arquivos_png()
