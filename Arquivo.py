from ftplib import FTP
import re
import os
import sys
import time
# import PyPDF2
import win32print
import win32api
import datetime


# Configurações do servidor FTP
ftp_server = 'IP'
ftp_user = 'user'
ftp_password = 'senha'

# Ler as configurações do arquivo config.txt
config = {}
with open('config.txt', 'r') as arquivo_config:
    for linha in arquivo_config:
        chave, valor = linha.strip().split(':')
        config[chave.strip()] = valor.strip()

ftp_root_folder = config.get('ftp_root_folder', '')  # Obtenha a pasta raiz FTP do arquivo de configuração

if not ftp_root_folder:
    print("Erro: A chave 'ftp_root_folder' não encontrada no arquivo de configuração.")
    sys.exit(1)  # Saia do programa se não houver configuração válida

# Solicitar ao usuário a lista de números de NFs
numeros_nf = input("Digite a lista de números de NFs que deseja imprimir (separados por quebras de linha):\n")
# Dividir a sequência em uma lista
lista_de_numeros = numeros_nf.split(";")

# Remover strings vazias da lista
numero_procurado = [numero for numero in lista_de_numeros if numero]

# Função para verificar se um arquivo existe no servidor FTP
def verifica_arquivo_ftp(ftp, pasta_ftp, NF):
    try:
        arquivos_ftp = ftp.nlst(pasta_ftp)
        for nome_arquivo in arquivos_ftp:
            if re.search(NF, nome_arquivo):
                return nome_arquivo
        return None
    except Exception as e:
        print(f"Erro ao verificar arquivo no FTP: {e}")
        salvarLog(NF, f"Erro ao verificar arquivo no FTP: {e}")
        return None
    
# Função para salvar log
def salvarLog(text, status):
    # Abrir ou criar o arquivo de log
    log_file = open('Arquivo de log.log', 'a')  # 'a' significa modo de apêndice, para adicionar ao arquivo existente

    # Obter a data e hora atual
    data_hora_atual = datetime.datetime.now()

    # Escrever informações no arquivo de log
    log_file.write(f'{data_hora_atual}: NF: {text} {status}. \n')

    # Fechar o arquivo de log
    log_file.close()

print ('\nLISTA DE IMPRESSORA DISPONIVEL')
contador = 0
for printer in win32print.EnumPrinters(2):
    print(f'{contador} - {printer[2]}')
    contador += 1

num_impressora = input ('\nDigite o numero da impressora que deseja utilizar: ')
lista_impressoras = win32print.EnumPrinters(2)
impressora = lista_impressoras[int(num_impressora)]

win32print.SetDefaultPrinter(impressora[2])

# Conectar-se ao servidor FTP
with FTP(ftp_server) as ftp:
    ftp.login(ftp_user, ftp_password)
    ftp.cwd(ftp_root_folder)  # Usar a pasta raiz FTP do arquivo de configuração

    for NF in numero_procurado:
        #time.sleep(2)
        arquivo_encontrado = verifica_arquivo_ftp(ftp, ftp_root_folder, NF)
        
        if arquivo_encontrado:
            print(f"O arquivo {NF} foi encontrado no servidor FTP.")

            # Obtenha o nome do arquivo do caminho completo
            nome_arquivo = os.path.basename(arquivo_encontrado)

            # Caminho completo do arquivo local (na mesma pasta do arquivo Python)
            local_file_path = os.path.join(os.path.dirname(sys.argv[0]), 'imprimir', nome_arquivo)

            # Baixar o arquivo encontrado para a pasta onde o arquivo Python está
            try:
                with open(local_file_path, 'wb') as local_file:
                    ftp.retrbinary('RETR ' + nome_arquivo, local_file.write)
            except:
                salvarLog(NF, "Não salvou na pasta")

            # Imprimir o arquivo PDF baixado na impressora padrão
            win32api.ShellExecute(0, "print", local_file_path, None, "", 0)
            
            # Excluir o arquivo após a impressão
            salvarLog(NF, "Encontrada e impressa na impressora com sucesso")
        else:
            print(f"Nenhum arquivo com a NF: {NF} foi encontrado no servidor FTP")
            salvarLog(NF, "Não encontrado no FTP.")

input('\nInformações salvas com sucesso no LOG\nAperte qualquer tecla para fechar.')
