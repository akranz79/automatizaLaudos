import os
import json
from docx import Document
import matplotlib.pyplot as plt

# Carregar configurações do arquivo JSON
with open('config.json', 'r') as config_file:
    config = json.load(config_file)

# Acessar o diretório dos laudos
diretorio_laudos = config["diretorio_laudos"]

def gerar_grafico_e_salvar_imagem(aguia, gato, tubarao, lobo, nome_laudo):
    # Convertendo os valores para números
    aguia = float(aguia.strip('%'))
    gato = float(gato.strip('%'))
    tubarao = float(tubarao.strip('%'))
    lobo = float(lobo.strip('%'))

    # Dados para o gráfico
    animais = ['Águia', 'Gato', 'Tubarão', 'Lobo']
    resultados = [aguia, gato, tubarao, lobo]

    # Calcular a altura do gráfico adicionando 5% à altura da barra mais alta
    altura_grafico = max(resultados) * 1.1

    # Criar o gráfico
    fig, ax = plt.subplots()
    ax.bar(animais, resultados, color=['blue', 'orange', 'green', 'red'])  # Ajuste as cores conforme necessário
    ax.set_xlabel('Animais')
    ax.set_ylabel('Resultados (%)')
    ax.set_title('AVALIAÇÃO DE PREFERÊNCIA CEREBRAL')

    # Adicionar o valor da porcentagem sobre cada barra
    for i, v in enumerate(resultados):
        ax.text(i, v + 1, f"{v:.0f}%", ha='center')

    # Definir o limite superior do eixo y com a altura do gráfico calculada
    ax.set_ylim(top=altura_grafico)

    # Salvar o gráfico como uma imagem com o mesmo nome do laudo
    nome_arquivo_imagem = f"{nome_laudo.replace('.docx', '.png')}"
    caminho_imagem = os.path.join(diretorio_laudos, nome_arquivo_imagem)
    plt.savefig(caminho_imagem)

    # Fechar a figura para liberar recursos
    plt.close()

    return caminho_imagem

def inserir_imagem_laudo():
    # Listar todos os arquivos na pasta de laudos
    for arquivo in os.listdir(diretorio_laudos):
        if arquivo.endswith('.docx') and arquivo != 'modelo.docx':
            caminho_arquivo_docx = os.path.join(diretorio_laudos, arquivo)

            # Abrir o documento docx
            doc = Document(caminho_arquivo_docx)

            # Gerar o gráfico e salvar a imagem
            table = doc.tables[0]
            aguia = table.cell(0, 1).text
            gato = table.cell(1, 1).text
            tubarao = table.cell(2, 1).text
            lobo = table.cell(3, 1).text
            caminho_imagem = gerar_grafico_e_salvar_imagem(aguia, gato, tubarao, lobo, arquivo)

            print(f"Imagem do gráfico salva em: {caminho_imagem}.[2]")

if __name__ == "__main__":
    inserir_imagem_laudo()
