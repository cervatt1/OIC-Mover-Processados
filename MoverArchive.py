import win32serviceutil
import win32service
import win32event
import os
import shutil


class MyService(win32serviceutil.ServiceFramework):
    _svc_name_ = "MoverArquive"
    _svc_display_name_ = "Mover de Processados para P/R"
   
    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)
        self.processados_path = {
            'APagar': 'C:/Inout/contabil/Bsacontabilidade/Unidasrepresentacoese_#286/APagar/Processado',
            'AReceber': 'C:/Inout/contabil/Bsacontabilidade/Unidasrepresentacoese_#286/AReceber/Processado',
            'Tareffa': 'C:/Tareffa/processado'
        }
        self.destino_path = {
            'APagar': 'C:/Inout/contabil/Bsacontabilidade/Unidasrepresentacoese_#286/APagar',
            'AReceber': 'C:/Inout/contabil/Bsacontabilidade/Unidasrepresentacoese_#286/AReceber',
            'Tareffa': 'C:/Tareffa'
        }
        self.arquivo_especial = 'ARQUIVO_APOIO_PAGAR.csv'
        self.arquivo_especial2 = 'ARQUIVO_APOIO_RECEBER.csv'        


    def SvcDoRun(self):
        while True:
            for pasta, path in self.processados_path.items():
                arquivos_processados = os.listdir(path)
                for arquivo in arquivos_processados:
                    caminho_arquivo = os.path.join(path, arquivo)
                    if os.path.isfile(caminho_arquivo) and arquivo != self.arquivo_especial and arquivo != self.arquivo_especial2:
                        win32event.WaitForSingleObject(self.hWaitStop, 5000)
                        shutil.move(caminho_arquivo, self.destino_path[pasta])
            win32event.WaitForSingleObject(self.hWaitStop, 5000)


    def SvcStop(self):
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        win32event.SetEvent(self.hWaitStop)


if __name__ == '__main__':
    win32serviceutil.HandleCommandLine(MyService)
