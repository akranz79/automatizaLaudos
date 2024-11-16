import os
import win32com.client as win32
from shutil import move
import json
import time
import logging
from log_config import setup_logging

def obter_numero_laudo():
    try:
        with open('estado_laudo.txt', 'r') as file:
            # Lê a última linha e converte para inteiro
            numero_laudo = int(file.readlines()[-1].strip())
            return numero_laudo
    except FileNotFoundError:
        # Se o arquivo não existir, retorna 0 como padrão
        return 0
    except ValueError:
        # Se houver um erro de leitura, retorna 0 como padrão
        return 0

def enviar_email(destinatario, copia_destinatario, assunto, corpo, anexo):
    try:
        # Criar a integração com o Outlook
        outlook = win32.Dispatch('Outlook.Application')

        # Criar um email
        email = outlook.CreateItem(0)

        # Configurar as informações do seu email
        email.To = destinatario
        email.CC = copia_destinatario  # Adiciona o e-mail em cópia
        email.Subject = assunto
        email.Body = corpo

        # Adicionar anexo
        email.Attachments.Add(anexo)

        # Enviar email
        email.Send()
        print(f"Email enviado para {destinatario} , {copia_destinatario} com assunto: {assunto}.")
        logging.info(f"Email enviado para {destinatario} , {copia_destinatario} com assunto: {assunto}.")

        return True
    except Exception as e:
        print(f"Falha ao enviar o email para {destinatario} com assunto: {assunto}. Erro: {e}")
        return False


def enviar_email_candidato():
    try:
        # Verificar se o arquivo de informações do candidato existe
        if not os.path.exists("candidato_info.json"):
            print("Verificar se o arquivo de informações do candidato existe...")
            return

        # Ler as informações do candidato
        with open("candidato_info.json") as f:
            info_candidato = json.load(f)

        # Extrair informações do candidato
        destinatario = info_candidato.get("email")
        #destinatario = 'ahkranz79@gmail.com'
        nome = info_candidato.get("nome")
        caminho_imagem = info_candidato.get("caminho_imagem")
        candidato_vaga = info_candidato.get("candidato_vaga")
        resultado_teste = info_candidato.get("resultado_teste")

        # Configurar o assunto e corpo do email
        assunto = f"Confirmação de Recebimento - {nome}"
        corpo = f"""
                Olá {nome} , 
                Seu formulário para a vaga de {candidato_vaga} foi recebido com sucesso.

                Resultado do seu teste:
                {resultado_teste}

                Em anexo, está a imagem do resultado do teste.
                Atenciosamente,
                Equipe de Recrutamento


                Desenvolvido por Alexandre Kranz
                """

        # Enviar email para o candidato
        enviar_email(destinatario, "", assunto, corpo, caminho_imagem)

        # Remover o arquivo após o envio para evitar duplicidade
        os.remove("candidato_info.json")

    except Exception as e:
        print(f"Erro ao enviar email para o candidato: {e}")


# Chamada para enviar email para o candidato
enviar_email_candidato()
logging.info("Esteira de processo[7.1]: chamando enviar_email_candidato : enviaEmail.py")


def enviar_emails_laudos(diretorio, destinatario, copia_destinatario):
    numero_laudo = obter_numero_laudo()
    try:
        # Listar arquivos no diretório
        arquivos = [arq for arq in os.listdir(diretorio) if
                    arq.endswith(".docx") and not arq.startswith("~$") and arq != "modelo.docx"]

        # Verificar se há arquivos para enviar
        if not arquivos:
            return

        for arquivo in arquivos:
            # Configurar as informações do email
            nome_arquivo = os.path.splitext(arquivo)[0]
            assunto = f"Laudo nº {numero_laudo} - {nome_arquivo}"
            corpo = f"""
                    Olá,
                    Segue em anexo o laudo nº {numero_laudo} .
                    Para suporte, enviar email: ahkranz79@gmail.com
                    Atenciosamente,
                    Alexandre Kranz

                    """

            anexo = os.path.join(diretorio, arquivo)

            # Enviar email
            if enviar_email(destinatario, copia_destinatario, assunto, corpo, anexo):
                # Esperar 5 segundos antes de mover o arquivo para 'Enviados'
                time.sleep(3)

                # Mover o arquivo enviado para a pasta 'Enviados'
                move(anexo, os.path.join(diretorio, "Enviados", arquivo))
                logging.info("Esteira de processo[7.2]: chamando enviar_emails_laudos : enviaEmail.py")

    except FileNotFoundError:
        print(f"O diretório {diretorio} não foi encontrado.")
    except PermissionError:
        print(f"Sem permissão para acessar o diretório {diretorio}.")
    except Exception as e:
        print(f"Ocorreu um erro: {e}")


# Diretório a ser verificado
diretorio = r"D:\Laudos"
imagem = r"D:\Laudos\imgLaudos"

# Ler o destinatário e o e-mail em cópia do arquivo config.json
with open('config.json') as f:
    config = json.load(f)
destinatario = config.get('destinatario', '')  # Obtém o destinatário do arquivo config.json
copia_destinatario = config.get('copia_destinatario', '')  # Obtém o e-mail em cópia do arquivo config.json

# Criar a pasta 'Enviados' se ela não existir
pasta_enviados = os.path.join(diretorio, "Enviados")
if not os.path.exists(pasta_enviados):
    os.makedirs(pasta_enviados)

# Chamada da função para enviar emails
enviar_emails_laudos(diretorio, destinatario, copia_destinatario)
