import os

def usersLoading():
    informacoes = {'Scpc:':'',
                   'ScpcUser': None,
                   'ScpcPassword': None,
                   'VAZIO':'',

                   'Serasa:':'',
                   'SerasaUser': None,
                   'SerasaPassword': None,

                   'ENTER': '',
                   'Spc-Brasil:': '',
                   'SpcUser': None,
                   'SpcChave': None,
                   'SpcPassword': None}
    try:
        with open('Load.txt', 'r') as load:
            for rows in load:
                if rows != '':
                    rowVector = rows.split('=')
                    find = rowVector[0].strip()
                    try:
                        inf = rowVector[1].strip()
                        if find in informacoes:
                            informacoes[find] = inf
                    except:
                        pass
    except:
        with open('Load.txt', 'wt+') as load:
            for d in informacoes:
                if d == 'VAZIO' or d == 'ENTER':
                    load.write('\n')
                elif ':' in d:
                    load.write(f'{d}\n')
                else:
                    load.write(f'{d} = *******\n')
        raise Exception('Preencha as informações de Login')
    return informacoes




usersLoading()

