from corpo.funcoes import *
import time
#import configparser
#import os

#thisfolder = os.path.dirname(os.path.abspath(__file__))
#configfile = os.path.join(thisfolder, f'config.ini')
#config = configparser.RawConfigParser()
#res = config.read(configfile)

#ler email
download_user = f'eventlog@i9brgroup.com.br' #config.get('download_email_loguin', 'user')
download_password = f'EventLog@2021' #config.get('download_email_loguin', 'password')
download_server = f'mail.i9brgroup.com.br' #config.get('download_email_loguin', 'server')
download_pasta = f'emails' #config.get('download_email_loguin', 'pasta')

#config
enviar_tempo_email = 6 #config.getint('enviar_email_tempo', 'tempo_hrs')
tempo_prox_email = 24 #config.get('tempo_proximo_email', 'tempo_prox_email_hrs')

#enviar email
enviar_email_loguin = f'backoffice@i9brgroup.com.br' #config.get('enviar_email_loguin', 'user')
enviar_email_password= f'i91234B@'#config.get('enviar_email_loguin', 'password')
enviar_server = f'mail.i9brgroup.com.br' #config.get('enviar_email_loguin', 'server')
enviar_porta = 587 #config.getint('enviar_email_loguin', 'port')

while True:
    print('Baixando relatorio')
    download_atr_status(download_user, download_password, download_server, download_pasta)
    time.sleep(2)

    print('Processando Relatorio')
    processar(enviar_tempo_email,enviar_email_loguin, enviar_email_password, tempo_prox_email, enviar_server, enviar_porta)
    time.sleep(1800)





