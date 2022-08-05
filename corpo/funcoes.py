# -*- coding: utf-8 -*-
import csv
import datetime
from datetime import timedelta
import smtplib    #biblioteca para enviar email
from email.mime.text import MIMEText
import imaplib   #biblioteca do download de emails
import email
import pandas as pd
import os
import  sys
caminho_status = r"C:\\cribmonitor\\arquivos\\atr_status.csv"
caminho_contatos = r"C:\\cribmonitor\\arquivos\\contatos.csv"


ROOT_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(ROOT_DIR, 'config.ini')  # requires `import os`


print(ROOT_DIR)
#config = configparser.RawConfigParser()
#res = config.read(configfile)





def gravar_log(text):
    with open('log.txt', 'a') as arquivo:
        arquivo.write(f'{datetime.datetime.today()} {text}\n')
        arquivo.close()

def download_atr_status(download_user, download_password, download_server, download_pasta):
    '''user = config.get('download_email_loguin', 'user')
    password = config.get('download_email_loguin', 'password')
    server = config.get('download_email_loguin', 'server')
    pasta = config.get('download_email_loguin', 'pasta')'''
    user = download_user
    password = download_password
    server = download_server
    pasta = download_pasta



    def connect(server, user, password):
        m = imaplib.IMAP4_SSL(server)
        m.login(user, password)
        m.select()
        return m

    def downloaAttachmentsInEmail(m, emailid, outputdir):
        resp, data = m.fetch(emailid, "(BODY.PEEK[])")
        email_body = data[0][1]
        mail = email.message_from_bytes(email_body)
        dados2 = mail.values()
        if dados2[16] == 'CribMonitor':
            for part in mail.walk():
                if part.get_content_maintype() != 'multipart' and part.get_all('Content-Disposition') is not None:

                    open(outputdir + '/' + 'atr_status' + '.csv', 'wb').write(part.get_payload(decode=True))

    def downloadAllAttachmentsInInbox(server, user, password, outputdir):
        m = connect(server, user, password)
        resp, items = m.search(None, "(ALL)")
        items = items[0].split()
        for emailid in items:
            idmail = f"b'{len(items)}'"
            if str(emailid) == idmail:
                downloaAttachmentsInEmail(m, emailid, outputdir)


    try:
        downloadAllAttachmentsInInbox(server, user, password, pasta)
        gravar_log('Download do relatorio concluido com sucesso.')
    except:
        gravar_log('Não foi possivel baixar o relatorio, utilizaremos o relatorio antigo.')
        print('Não foi possivel baixar o relatorio, utilizaremos o relatorio antigo.')




def processar(enviar_tempo_email,enviar_email_loguin, enviar_email_password, tempo_prox_email, enviar_server, enviar_porta):

    def ajustar_tempo(hora, sigla):
        if sigla == 'pm':
            hora = hora + timedelta(hours=12)
        else:
            hora = hora
        return hora


    def ler_atr_status():


        with open(caminho_status, 'r') as arquivo_csv:
            leitor = csv.reader(arquivo_csv, delimiter=',')
            next(leitor)
            for linha in leitor:
                '''Inicio Tempo'''
                gravar_log(f' ---- lendo crib')
                crib = linha[0]
                gravar_log(f' ---- lendo nome do ATR')
                nomeatr = linha[1]
                gravar_log(f' ---- lendo ultima hora lida')
                last_sinc = linha[2].lower()#
                time_sig = last_sinc[-2] + last_sinc[-1]
                last_sinc = last_sinc.replace('am', '').replace('pm','')
                gravar_log([last_sinc, crib, nomeatr])
                last_sinc = datetime.datetime.strptime(last_sinc, '%m/%d/%Y %I:%M:%S')
                last_sinc = ajustar_tempo(last_sinc, time_sig)
                status = linha[3]
                hoje = datetime.datetime.today()
                hoje = datetime.datetime.strftime(hoje, '%m/%d/%Y %I:%M:%S')
                hoje = datetime.datetime.strptime(hoje, '%m/%d/%Y %I:%M:%S')
                tempo_estipulado = datetime.timedelta(hours=enviar_tempo_email)
                tempo_off = hoje - last_sinc
                if status == '':
                    if int(crib) < 998:
                        gravar_log(f' ---- comparando tempo')
                        if tempo_off > tempo_estipulado:
                            gravar_log(f' ---- tempo comparado --- crib:{crib} --- tempo de atraso:{tempo_off} --- tempo configurado:{tempo_estipulado} ')
                            gravar_log(f' ---- enviando email')
                            enviar_email(crib, last_sinc, nomeatr)
                            gravar_log(f' ---- email enviado')
    def alterar_data_csv(crib):
        contato = pd.read_csv(caminho_contatos, sep=';')
        for (i, row) in contato.iterrows():
            if row['crib'] == int(crib):
                contato.loc[i, 'last_mail_contat'] = datetime.datetime.today()
                gravar_log(f' ---- data alterada crib{crib}')
        gravar_log(f' ---- salvando arquivo')
        contato.to_csv(caminho_contatos, index=False, sep=';')


    def enviar_email(crib, data, nome):
        gravar_log(f' ---- requisitando dados de loguin')
        user = enviar_email_loguin
        password = enviar_email_password
        time_next_email = tempo_prox_email
        server = enviar_server
        port = enviar_porta
        gravar_log(f' ---- dados ok')
        gravar_log(f' ---- abrindo arquivo de contatos')
        with open(caminho_contatos, 'r') as arquivo_csv:
            leitor = csv.reader(arquivo_csv, delimiter=';')
            gravar_log(f' ---- contatos abertos')
            next(leitor)
            for linha in leitor:
                crib_contat = linha[0]
                last_mail_contat = datetime.datetime.strptime(linha[1], '%Y-%m-%d %H:%M:%S.%f')
                emails = linha[2],linha[3]
                if crib == crib_contat:
                    tempo_email_contat = datetime.datetime.today() - last_mail_contat
                    if tempo_email_contat > datetime.timedelta(hours=int(time_next_email)): #maualmente novo ate aqui
                        email = emails[0]  # retorna o valor 0 e exclui ele da lista
                        #print(email)
                        sender = user
                        recipients = email
                        copia = str()
                        if len(emails) > 0:
                            copia = emails[1]

                        msg = MIMEText(f'PREZADO CLIENTE,\n'
                                       f'POR FAVOR, VERIFICAR A CONEXAO DO DISPENSÁRIO {nome} CRIB {crib}.\n'
                                       f'QUE INDICA ESTAR DESCONECTADO. DESDE O DIA {data}.\n\n\n'
                                       f'CASO SEU MODELO DE MÁQUINA SEJA COM TELA TOUCH (DO TAMANHO DE UM CELULAR), VERIFICAR:\n\n'
                                       f'1- A DATA GRAVADA NO SISTEMA DA MÁQUINA. PASSE SEU CRACHÁ (DEVERÁ TER PERFIL DE ADMININSTRADOR), VÁ EM DIAGNÓSTICO E VÁ EM ALTERAR A DATA DO SISTEMA. CASO NÃO ESTEJA O DIA DE HOJE, ALTERE CLICANDO NO ANO E DEPOIS NO DIA E HORA\n'
                                       f'2- A API, TESTANDO-A ATRAVÉS DA PÁGINA DIAGNÓSTICO, TESTAR API E AGUARDAR A MENSAGEM DE BEM SUCEDIDO. CASO APAREÇA UMA MENSAGEM DE ERRO, VERIFIQUE SE APARECE, NESTA MESMA TELA, LOGO ABAIXO DA OPCAO TESTAR API, O NÚMERO DE IP. CASO NÃO TENHA NÚMERO DE IP, POR FAVOR, CONTATE A I9.\n'
                                       f'EM CASO DE DÚVIDA OU INSUCESSO AO RECONECTAR, ENTRAR EM CONTATO COM:\n\n\n'
                                       f'- BACKOFFICE@I9BRGROUP.COM.BR\n\n'
                                       f'- 11 3141 0749\n\n'
                                       f'- 11 98439 9685\n\n')
                        msg['Subject'] = f'URGENTE MAQUINA DESCONECTADA CRIB {crib}'
                        msg['From'] = sender
                        msg['To'] = email
                        msg['CC'] = copia
                        finalto = [recipients] + [copia]  # corrigir metodo de envio, precisa passar uma string para cada email, nao da suporte a lista ou uma string para varias emails.
                        #finalto = ['pedrohique@hotmail.com'] + ["backoffice@i9brgroup.com.br"]
                        s = smtplib.SMTP(server, port)
                        s.set_debuglevel(0)
                        s.login(user, password)
                        s.sendmail(sender, finalto, msg.as_string())
                        s.quit()

                        print(f'Email enviado Crib {crib} - {finalto}')
                        gravar_log(f' ---- email enviado')
                        gravar_log(f' ---- lalterando data de envio de email')
                        alterar_data_csv(crib)
                    else:
                        gravar_log(f' ---- email enviado a menos de {time_next_email}hrs --- crib:{crib} --- ultimo email:{last_mail_contat}')
                        print(f'Crib: {crib} atrasado a mais de 4 horas, email enviado anteriormente')

    ler_atr_status()



