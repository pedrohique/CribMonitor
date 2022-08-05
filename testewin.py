import os
import win32serviceutil
import win32service
import win32event
import servicemanager
import socket
import logging
import sys

import time



caminho_log = f'C:/cribmonitor/Logs/log2.txt'
logging.basicConfig(
    filename=caminho_log,
    level=logging.DEBUG,
    format='[helloworld-service] %(levelname)-7.7s %(message)s'
)


class HelloWorldSvc(win32serviceutil.ServiceFramework):
    logging.info('entrou na classe ...')
    _svc_name_ = "aaateste"
    _svc_display_name_ = "aaateste"


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

        arquivo = open(f"c:\\cribmonitor\\euestouaqui.txt", "a")
        arquivo.write('Ola eu estou aqui')




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