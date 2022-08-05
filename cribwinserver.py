import os
import win32serviceutil
import win32service
import win32event
import servicemanager
import socket
import logging
import sys

from corpo.funcoes import *
import time



caminho_log = f'C:/cribmonitor/Logs/log2.txt'
logging.basicConfig(
    filename=caminho_log,
    level=logging.DEBUG,
    format='[helloworld-service] %(levelname)-7.7s %(message)s'
)


class HelloWorldSvc(win32serviceutil.ServiceFramework):
    logging.info('entrou na classe ...')
    _svc_name_ = "cribmonitor"
    _svc_display_name_ = "CribMonitor"


    def __init__(self, args):
        logging.info('Entrou no init...')
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.stop_event = win32event.CreateEvent(None, 0, 0, None)
        socket.setdefaulttimeout(60)
        self.stop_requested = False

    def SvcStop(self):
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        win32event.SetEvent(self.stop_event)
        logging.info('Stopping service ...')
        self.stop_requested = True

    def SvcDoRun(self):
        logging.info('entrou no run ...')
        servicemanager.LogMsg(
            servicemanager.EVENTLOG_INFORMATION_TYPE,
            servicemanager.PYS_SERVICE_STARTED,
            (self._svc_name_, '')
        )
        self.main()

    def main(self):
        '''import configparser
        import os
        logging.info(' ** Hello PyWin32 World ** ')

        thisfolder = os.path.dirname(os.path.abspath(__file__))
        configfile = os.path.join(thisfolder, f'config.ini')
        config = configparser.RawConfigParser()
        res = config.read(configfile)
        # ler email
        download_user = config.get('download_email_loguin', 'user')
        download_password = config.get('download_email_loguin', 'password')
        download_server = config.get('download_email_loguin', 'server')
        download_pasta = config.get('download_email_loguin', 'pasta')

        # config
        enviar_tempo_email = config.getint('enviar_email_tempo', 'tempo_hrs')
        tempo_prox_email = config.get('tempo_proximo_email', 'tempo_prox_email_hrs')

        # enviar email
        enviar_email_loguin = config.get('enviar_email_loguin', 'user')
        enviar_email_password = config.get('enviar_email_loguin', 'password')
        enviar_server = config.get('enviar_email_loguin', 'server')
        enviar_porta = config.getint('enviar_email_loguin', 'port')'''
        # ler email
        download_user = f'eventlog@i9brgroup.com.br'  # config.get('download_email_loguin', 'user')
        download_password = f'EventLog@2021'  # config.get('download_email_loguin', 'password')
        download_server = f'mail.i9brgroup.com.br'  # config.get('download_email_loguin', 'server')
        download_pasta = f'emails'  # config.get('download_email_loguin', 'pasta')

        # config
        enviar_tempo_email = 6  # config.getint('enviar_email_tempo', 'tempo_hrs')
        tempo_prox_email = 24  # config.get('tempo_proximo_email', 'tempo_prox_email_hrs')

        # enviar email
        enviar_email_loguin = f'backoffice@i9brgroup.com.br'  # config.get('enviar_email_loguin', 'user')
        enviar_email_password = f'i91234B@'  # config.get('enviar_email_loguin', 'password')
        enviar_server = f'mail.i9brgroup.com.br'  # config.get('enviar_email_loguin', 'server')
        enviar_porta = 587  # config.getint('enviar_email_loguin', 'port')

        while True:

            print('Baixando relatorio')
            download_atr_status(download_user, download_password, download_server, download_pasta)
            time.sleep(2)

            print('Processando Relatorio')
            processar(enviar_tempo_email, enviar_email_loguin, enviar_email_password, tempo_prox_email, enviar_server,
                      enviar_porta)
            time.sleep(1800)




if __name__ == '__main__':
    logging.info('entrou no if1 ...')

    if len(sys.argv) == 1:
        logging.info('entrou no if 2...')
        servicemanager.Initialize()
        servicemanager.PrepareToHostSingle(HelloWorldSvc)
        servicemanager.StartServiceCtrlDispatcher()
    else:
        logging.info('entrou no else ...')
        win32serviceutil.HandleCommandLine(HelloWorldSvc)