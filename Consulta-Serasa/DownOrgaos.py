from IAdcancend import initSeWeb as isw
from selenium.webdriver.common.by import By
import Excentions as et
from pathlib import Path
from time import sleep
import pandas as pd
import os

def scpcBaixada(filename):
    scpc = isw('https://www.scpc.inf.br/cgi-bin/spcnweb?HTML_PROGRAMA=md000001.int#')
    #Login
    scpc.writeText('//*[@id="HTML_COD"]', '36541')
    scpc.writeText('//*[@id="HTML_SEN"]', '42684258')
    scpc.clickElement('//*[@id="HTML_BOTAO"]')

    #Indo até manutenções
    scpc.clickElement('//*[@id="menu_principal_spcn"]')
    scpc.clickElement('//*[@id="sm02_SPCN"]/a')
    scpc.entryIframe('//*[@id="menu_vertical"]')
    scpc.clickElement('//*[@id="esquerda_menu"]/ul[7]/li[1]/form/a')
    scpc.exitIframe()

    scpc.entryIframe('//*[@id="Tela"]')
    scpc.clickElement('//*[@id="todoform"]/tbody/tr[2]/td[2]/input[2]')
    scpc.entryWindow()
    scpc.entryIframe('//*[@id="pform3"]')

    list = scpc.getListElement(By.XPATH, '/html/body/table/tbody')

    arquivo_ret = filename
    patch_home = Path.home()
    download = patch_home / "Downloads"
    arquivodiretory = download / arquivo_ret

    for txt in list:
        files = txt.find_elements(By.XPATH, f'/html/body/table/tbody/tr')
        for file in files:
            tds = str(file.find_element(By.XPATH, './td[1]').text).strip()
            if tds.upper() == arquivo_ret.upper():
                file.find_element(By.XPATH, './td[4]').click()
                print('Encontrado')
                while True:
                    if et.verificyExistFile(arquivo_ret):
                        break
                    sleep(5)
                break

    scpc.closeWindow()
    scpc.exitWindow()

    with open(arquivodiretory, 'r') as file:
        for row in file:
            if row[371:379].strip() != 'EXCLUSAO' and row[371:379].strip() != '':
                cpf = row[59:70].strip()
                card = row[316:338].strip()
                scpc.exitIframe()
                scpc.entryIframe('//*[@id="menu_vertical"]')
                scpc.clickElement('//*[@id="form_001"]/a')
                scpc.exitIframe()

                scpc.entryIframe('//*[@id="TelaNovo"]')
                scpc.writeText('//*[@id="cpf"]', cpf)
                scpc.clickElement('//*[@id="btn_pesquisar"]')

                incluso = scpc.getListElement(By.XPATH, '//*[@id="tbl_fis"]/tbody/tr')

                for i in incluso:
                    if str(i.text) == 'Nenhum registro encontrado':
                        print(f"Cliente {cpf}, não possui restrição")
                        continue
                    td = i.find_element(By.XPATH, './td[5]').text
                    if len(td.strip()) == 14:
                        cardinc = td.strip()
                    else:
                        cardinc = td.strip().lstrip('0')
                        card = card.lstrip('0')
                    if cardinc == card:
                        i.find_element(By.XPATH, './td[9]').click()
                        scpc.clickElement('//*[@id="btn_excluir"]')
                        scpc.clickElement('/html/body/div[3]/div/div[1]/button')

def serasaBaixa(filename):
    serasa = isw('https://empresas.serasaexperian.com.br/meus-produtos/login')
    #Login
    serasa.writeText('//*[@id="loginUser"]', '57729563')
    serasa.writeText('//*[@id="loginPassword"]', 'M@ya6842')
    serasa.clickElement('//*[@id="loginFormSubmit"]')

    serasa.clickElement('//*[@id="mat-mdc-dialog-0"]/div/div/div/mat-dialog-content/div/div[2]/div/a')

    serasa.clickElement('//*[@id="prod-64adce65a625ed4687bde841"]/div/div[3]/button')

    clientinexcel = pd.read_excel(filename)

    cards = clientinexcel['Cartões']

    for card in cards:
        cpf = str(clientinexcel.loc[clientinexcel['Cartões'] == card, 'Cpf'].values[0]).strip('.').lstrip('0')
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
    spc.writeText('//*[@id="j_username"]', '103584514')
    spc.writeText('//*[@id="j_password"]', '@Bolinhabranca1012')
    spc.clickElement('//*[@id="submitButton"]/span')
    spc.writeText('//*[@id="passphrase"]', '@Bola1012')
    spc.clickElement('//*[@id="submitButton"]/span')

    spc.clickElement('/html/body/section[2]/div/table/tbody/tr[2]/td[2]/div/a/div/figure/img')
    spc.clickElement('//*[@id="accordion2"]/li[8]/a')

    listclient = pd.read_excel(filename)
    cardlist = listclient['Cartões']
    for cardc in cardlist:
        cpf = str(clientinexcel.loc[clientinexcel['Cartões'] == cardc, 'Cpf'].values[0]).strip('.').lstrip('0')
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








    sleep(15)
spcBaixa('dwadwda')





