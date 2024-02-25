from IAdcancend import initSeWeb as isw
from selenium.webdriver.common.by import By
import Excentions as et
from pathlib import Path
from time import sleep
import pandas as pd
import SubJudice as sj
import FileLoad

duser = FileLoad.usersLoading()

def scpcBaixada(filename):
    scpc = isw('https://www.scpc.inf.br/cgi-bin/spcnweb?HTML_PROGRAMA=md000001.int#')
    #FAZENDO LOGIN NO SITE
    scpc.writeText('//*[@id="HTML_COD"]', duser['ScpcUser'].strip())
    scpc.writeText('//*[@id="HTML_SEN"]', duser['ScpcPassword'].strip())
    scpc.clickElement('//*[@id="HTML_BOTAO"]')

    #INDO ATÉ A AREA DE MANUTENÇOES
    scpc.clickElement('//*[@id="menu_principal_spcn"]')
    scpc.clickElement('//*[@id="sm02_SPCN"]/a')

    #ACESSANDO O IFRAME
    scpc.entryIframe('//*[@id="menu_vertical"]')

    #ABRINDO EXCEL COM OS CARTÕES
    listclients = pd.read_excel(filename)
    cards = listclients['Cartões']

    #INTERANDO SOBRE CADA CARTÃO
    for card in cards:
        cpf = str(listclients.loc[listclients['Cartões'] == card, 'Cpf'].values[0]).strip('.').lstrip('0')
        manual = str(listclients.loc[listclients['Cartões'] == card, 'SCPC'].values[0]).strip()

        #CASO O manual RETORNE OUTRO NOME SEM SER MANUAL VAMOS PULAR ESSE CARTÃO
        if manual != 'MANUAL':
            continue

        #CORRIGINDO O CPF
        t = 11 - len(cpf)
        cpf = f'{'0' * t}{cpf}'

        #INDO ATÉ PESSOA FISICA PARA REALIZAR CONSULTA
        scpc.exitIframe()
        scpc.entryIframe('//*[@id="menu_vertical"]')
        scpc.clickElement('//*[@id="form_001"]/a')
        scpc.exitIframe()

        #AS VAZES OCORRE ALGUM ERRO QUANDO CLICAMOS EM PESSOA FISICA FIZ UM LOOP PARA REPETIR O PROCESSO ATÉ CONSEGUIR FAZER A CONSULTA
        while True:
            try:
                #PREENCHENDO O CPF
                scpc.entryIframe('//*[@id="TelaNovo"]')
                scpc.loadingElement('//*[@id="cpf"]', 5)
                scpc.web.find_element(By.XPATH, '//*[@id="cpf"]').send_keys(cpf)
                break
            except:
                scpc.exitIframe()
                scpc.entryIframe('//*[@id="menu_vertical"]')
                scpc.clickElement('//*[@id="form_001"]/a')
                scpc.exitIframe()
                sleep(2)

        #CLICANDO PARA REALIZAR A CONSULTA
        scpc.clickElement('//*[@id="btn_pesquisar"]')

        #PEGANDO TODAS RESTRIÇOES QUE O CLIENTE POSSUI
        incluso = scpc.getListElement(By.XPATH, '//*[@id="tbl_fis"]/tbody/tr')

        for i in incluso:

            #CASO ELE NÃO TENHA NENHUMA RESTRIÇÃO VAMOS PARA O PROXIMO CARTÃO
            if str(i.text) == 'Nenhum registro encontrado':
                print(f"Cliente {cpf}, não possui restrição")
                continue

            #PEGANDO O NUMERO DO CONTRATO INCLUSO
            td = i.find_element(By.XPATH, './td[5]').text

            if len(td.strip()) == 14:
                cardinc = td.strip()
            else:
                cardinc = td.strip().lstrip('0')
                card = card.lstrip('0')

            #CASO FOR O CONTRATO QUE QUEREMOS EXCLUIR VAMOS EXCLUIR
            if cardinc == card:
                i.find_element(By.XPATH, './td[9]').click()
                scpc.clickElement('//*[@id="btn_excluir"]')
                scpc.clickElement('/html/body/div[3]/div/div[3]/button[1]')
                sleep(4)

                try:
                    scpc.loadingElement('/html/body/div[3]/div/div[2]', 10)
                    regs = str(scpc.web.find_element(By.XPATH, '/html/body/div[3]/div/div[2]').text).strip()
                    print(regs)

                    # VERICICANDO SE NÃO É UM CASO DE RESGISTRO SUSPENSO
                    # CASO FOR VAMOS GERAR O REGISTRO SUSPENSO PARA SER ENVIADO PARA O ÓRGÃO
                    if 'EXCLUSAO NAO PERMITIDA, REGISTRO SUSPENSO' in regs:
                        scpc.clickElement('/html/body/div[3]/div/div[1]/button')
                        element_value = scpc.web.find_element(By.ID, 'valor')
                        element_name = scpc.web.find_element(By.ID, 'nome')
                        valor = element_value.get_attribute('value')
                        nome = element_name.get_attribute('value')
                        sj.createRegistro(cpf, nome, card, valor)

                except:
                    sleep(2)

def serasaBaixa(filename):
    serasa = isw('https://empresas.serasaexperian.com.br/meus-produtos/login')
    #Login
    serasa.writeText('//*[@id="loginUser"]', duser['SerasaUser'].strip())
    serasa.writeText('//*[@id="loginPassword"]', duser['SerasaPassword'].strip())
    serasa.clickElement('//*[@id="loginFormSubmit"]')

    serasa.clickElement('//*[@id="mat-mdc-dialog-0"]/div/div/div/mat-dialog-content/div/div[2]/div/a')

    serasa.clickElement('//*[@id="prod-64adce65a625ed4687bde841"]/div/div[3]/button')

    clientinexcel = pd.read_excel(filename)

    cards = clientinexcel['Cartões']

    for card in cards:
        cpf = str(clientinexcel.loc[clientinexcel['Cartões'] == card, 'Cpf'].values[0]).strip('.').lstrip('0')
        manual = str(clientinexcel.loc[clientinexcel['Cartões'] == card, 'SERASA'].values[0]).strip()

        if manual != 'MANUAL':
            continue

        card = str(card).strip()

        t = 11-len(cpf)
        cpf = f'{'0'*t}{cpf}'

        serasa.clearBoxTheText('//*[@id="debtorDocument"]')
        serasa.writeText('//*[@id="debtorDocument"]', cpf)
        serasa.clickElement('//*[@id="__next"]/main/div/form/div[1]/div[6]/button')

        try:
            serasa.loadingElement('//*[@id="__next"]/main/div/article/div/span[2]', 5)
            inf = serasa.web.find_element(By.XPATH, '//*[@id="__next"]/main/div/article/div/span[2]').text
            if inf.strip() == 'Nenhuma dívida encontrada':
                print('Nenhuma dívida encontrada')
                continue
        except:
            print('passado')
            pass

        inclusoes = serasa.getListElement(By.XPATH, '//*[@id="__next"]/main/div/article/div/table/tbody/tr')

        for incluso in inclusoes:
            td = incluso.find_element(By.XPATH, './td[3]').text
            if et.configCard(td) == card:
                incluso.find_element(By.XPATH, '//*[@id="__next"]/main/div/article/div/table/tbody/tr/td[7]/div/button[1]').click()
                serasa.clickElement('//*[@id="modal"]/div[1]/div/div/div[2]/div/div/div[2]/button')
                serasa.clickElement('/html/body/div[2]/div[1]/div/div/div[2]/form/div[1]/label[13]/input')
                serasa.clickElement('//*[@id="modal"]/div[1]/div/div/div[2]/form/div[2]/div[2]')
                print(et.configCard(td))
                sleep(30)

def spcBaixa(filename):
    spc = isw('https://sistema.spc.org.br/spc/controleacesso/autenticacao/entry.action;jsessionid=5ed367a9-15f7-4720-bfca-eff5e8f1875a_node186')

    #login
    spc.writeText('//*[@id="j_username"]', duser['SpcUser'].strip())
    spc.writeText('//*[@id="j_password"]', duser['SpcChave'].strip())
    spc.clickElement('//*[@id="submitButton"]/span')
    spc.writeText('//*[@id="passphrase"]', duser['SpcPassword'].strip())
    spc.clickElement('//*[@id="submitButton"]/span')

    spc.clickElement('/html/body/section[2]/div/table/tbody/tr[2]/td[2]/div/a/div/figure/img')
    spc.clickElement('//*[@id="accordion2"]/li[8]/a')

    listclient = pd.read_excel(filename)
    cardlist = listclient['Cartões']
    for cardc in cardlist:
        cpf = str(listclient.loc[listclient['Cartões'] == cardc, 'Cpf'].values[0]).strip('.').lstrip('0')
        manual = str(listclient.loc[listclient['Cartões'] == cardc, 'SPC'].values[0]).strip()

        if manual != 'MANUAL':
            continue

        spc.clickElement('//*[@id="accordion2"]/li[8]/ul/li/a')
        spc.clickElement('//*[@id="m50"]/div/a')

        t = 11 - len(cpf)
        cpf = f'{'0' * t}{cpf}'

        spc.writeText('//*[@id="numeroDocumento"]', cpf)
        spc.clickElement('//*[@id="conteudoInsumo"]/tbody/tr[2]/td/table/tbody/tr/td/input[1]')

        spc.loadingElement('//*[@id="dataGrid"]/tbody', 100)
        list = spc.getListElement(By.XPATH, '//*[@id="dataGrid"]/tbody')
        for l in list:
            td = l.find_elements(By.XPATH, './tr')
            for tds in td:
                cardTD = tds.find_element(By.XPATH, './td[7]')
                card = cardTD.text
                card = et.configCard(card)
                if card != None and cardc == card:
                    cardTD.click()
                    spc.clickElement('//*[@id="motivoExclusaoRegistro.id"]/option[14]')
                    spc.clickElement('//*[@id="formSPC"]/table[3]/tbody/tr/td/input[2]')
                    spc.acceptAlert()
                    sleep(1.5)







# scpcBaixada(r'C:\Users\fabio.prado\Desktop\Área de Trabalho\cpf.xlsx')





