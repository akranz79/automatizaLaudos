import gspread
from oauth2client.service_account import ServiceAccountCredentials
import os
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from datetime import datetime
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive

# Defina a variável gc globalmente
gc = gspread.service_account(filename='clean-beaker-414415-8887a7f9338c.json')
def obter_dados_planilha(linha):
    planilha = gc.open_by_url(
        "https://docs.google.com/spreadsheets/d/1U9ZFaFmNf4aJ13gbzxuF8AObk-tDWFOaE9If5iOplNc/edit?resourcekey#gid=307731029"). \
        worksheet("Respostas ao formulário 1")
    dados = planilha.row_values(linha)
    # Remover espaços em branco no início e no final de cada elemento
    dados = [item.strip() if isinstance(item, str) else item for item in dados]
    return dados

def ler_estado_laudo():
    try:
        with open('estado_laudo.txt', 'r') as arquivo:
            estado = arquivo.read().splitlines()
            return estado
    except FileNotFoundError:
        return []

def salvar_estado_laudo(estado):
    with open('estado_laudo.txt', 'w') as arquivo:
        for linha_estado in estado:
            arquivo.write(linha_estado + '\n')

def verificar_laudos_nao_gerados():
    planilha = gc.open_by_url(
        "https://docs.google.com/spreadsheets/d/1U9ZFaFmNf4aJ13gbzxuF8AObk-tDWFOaE9If5iOplNc/edit?resourcekey#gid=307731029"). \
        worksheet("Respostas ao formulário 1")
    total_linhas_preenchidas = len(planilha.get_all_values())

    # Obter o estado dos laudos
    laudo_gerado = ler_estado_laudo()

    # Verificar quais linhas ainda não tiveram laudos gerados
    laudos_nao_gerados = []
    for linha in range(2, total_linhas_preenchidas + 1):  # Começa da linha 2, presumindo que a primeira linha seja cabeçalho
        if str(linha) not in laudo_gerado:
            laudos_nao_gerados.append(linha)

    # Imprimir as linhas que ainda não tiveram laudos gerados
    if laudos_nao_gerados:
        print("Os seguintes laudos ainda não foram gerados:")
        for linha in laudos_nao_gerados:
            dados = obter_dados_planilha(linha)
            carimbo, email, usuario, data_nascimento, candidato_vaga, pretensao_salarial = dados[:6]
            print(f"Line: [{linha}].[{usuario} . {candidato_vaga}]")
    else:
        print("Todos os laudos foram gerados.")



def main():
       verificar_laudos_nao_gerados()

if __name__ == "__main__":
    main()
