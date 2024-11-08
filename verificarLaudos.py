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


def laudo_ja_gerado(linha):
    laudo_gerado = ler_estado_laudo()
    return str(linha) in laudo_gerado
