import os
from datetime import datetime
from Enumeracoes import EPath
from time import sleep
import shutil


class PathController:
    def __init__(self, diretory):
        self._epath = EPath.PathNames
        self._dataatual = datetime.now().strftime('%d-%m-%Y')
        self._diretory = diretory
        self._pathend = f'{self._diretory}\\Relatórios - {self._dataatual}'
        self._paths = \
            ['1 - CREDITO VISA', '2 - DEBITO VISA', '3 - CREDITO MASTERCARD',
             '4 - DEBITO MASTERCARD', '5 - CREDITO ELO', '6 - DEBITO ELO',
             '7 - CREDITO HIPERCARD']
        self._createpaths()
    def _createpaths(self):
        try:
            os.makedirs(self._pathend)
            for name in self._paths:
                os.makedirs(f'{self._pathend}\\{name}')
        except:
            pass

    def moveAndRename(self, file, newname, path:EPath.PathNames):
        fvector = file.split('.')[-1]
        print(fvector)
        newname = f'{newname}.{fvector}'
        original = os.path.join(self._diretory, file)
        newnamefile = os.path.join(self._diretory, newname)
        os.rename(original, newnamefile)
        newdiretory = os.path.join(f'{self._pathend}\\{path}', newname)
        shutil.move(newnamefile, newdiretory)

    def detectNewFile(self, listaoriginal):
        ligado = True
        efile = ''
        while ligado == True:
            detect = os.listdir(self._diretory)
            for file in detect:
                dfile = os.path.join(self._diretory, file)
                if file not in listaoriginal and os.path.isfile(dfile):
                    efile = file
                    ligado = False
                    break
                else:
                    listaoriginal = detect
            sleep(2)
        return efile


comandos = PathController(r'C:\Users\Fábio Nabor\Desktop\Teste')
scan = os.listdir(r'C:\Users\Fábio Nabor\Desktop\Teste')
file = comandos.detectNewFile(scan)
comandos.moveAndRename(file, 'encontrado', EPath.PathNames.DEBITOELO.value)

