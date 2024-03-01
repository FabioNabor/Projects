from IAdcancend import initSeWeb as isw
from selenium.webdriver.common.by import By
import Excentions as et
from time import sleep
import SubJudice as sj


class Scpc:
    def __init__(self, user, password):
        self.user = user
        self.password = password

    def login(self):
        self.scpc = isw('https://www.scpc.inf.br/cgi-bin/spcnweb?HTML_PROGRAMA=md000001.int#', True)
        self.scpc.writeText('//*[@id="HTML_COD"]', self.user)
        self.scpc.writeText('//*[@id="HTML_SEN"]', self.password)
        self.scpc.clickElement('//*[@id="HTML_BOTAO"]')

    def __manutencoes(self):
        self.scpc.writeText('//*[@id="menu_principal_spcn"]')
        self.scpc.clickElement('//*[@id="sm02_SPCN"]/a')

    def __iframe_menu_vertical(self):
        self.scpc.exitIframe()
        self.scpc.entryIframe('//*[@id="menu_vertical"]')

    def consulta_cpf(self, cpf):
        self.__manutencoes()
        self.__iframe_menu_vertical()
        while True:
            try:
                self.scpc.entryIframe('//*[@id="TelaNovo"]')
                self.scpc.loadingElement('//*[@id="cpf"]', 5)
                self.scpc.web.find_element(By.XPATH, '//*[@id="cpf"]').send_keys(cpf)
                self.scpc.clickElement('//*[@id="btn_pesquisar"]')
                break
            except:
                self.scpc.exitIframe()
                self.scpc.entryIframe('//*[@id="menu_vertical"]')
                self.scpc.clickElement('//*[@id="form_001"]/a')
                self.scpc.exitIframe()
                sleep(5)

    def baixar_contrato(self, contrato):
        contratoativos = self.scpc.getListElement(By.XPATH, '//*[@id="tbl_fis"]/tbody/tr')
        quantidadeativos = len(contratoativos)

        if quantidadeativos > 1:
            return 'VERIFICAR'

        for contratoincluso in contratoativos:

            if 'Nenhum registro encontrado' in contratoincluso.text:
                return 'EXCLUSO'

            while True:
                try:
                    tagtds = contratoincluso.find_element(By.XPATH, './td[5]').text
                    break
                except:
                    sleep(2)

            if et.configCard(contrato) == et.configCard(tagtds):
                sleep(0.5)
                contratoincluso.find_element(By.XPATH, './td[9]').click()
                self.scpc.clickElement('//*[@id="btn_excluir"]')
                self.scpc.clickElement('/html/body/div[3]/div/div[3]/button[1]')
                sleep(4)
            try:
                info = self.scpc.getText('/html/body/div[3]/div/div[2]')

                if 'REGISTRO SUSPENSO' in info:
                    self.scpc.clickElement('/html/body/div[3]/div/div[1]/button')
                    element_cpf = self.scpc.web.find_element(By.ID, 'cpf')
                    element_value = self.scpc.web.find_element(By.ID, 'valor')
                    element_name = self.scpc.web.find_element(By.ID, 'nome')
                    cpfid = element_cpf.get_attribute('value')
                    valor = element_value.get_attribute('value')
                    nome = element_name.get_attribute('value')
                    sj.createRegistro(cpfid, nome, et.configCard(contrato), valor)
                    return 'REGISTRO SUSPENSO'
            except:
                return 'VERIFICAR'

    def montagem_registro_suspenso(self, contrato):
        contratoativos = self.scpc.getListElement(By.XPATH, '//*[@id="tbl_fis"]/tbody/tr')
        for contratoincluso in contratoativos:

            if 'Nenhum registro encontrado' in contratoincluso.text:
                return 'NENHUM REGISTRO ATIVO'

            while True:
                try:
                    tagtds = contratoincluso.find_element(By.XPATH, './td[5]').text
                    break
                except:
                    sleep(2)

            if et.configCard(contrato) == et.configCard(tagtds):
                sleep(0.5)
                contratoincluso.find_element(By.XPATH, './td[9]').click()
                self.scpc.loadingElement(By.ID, 'cpf')
                element_cpf = self.scpc.web.find_element(By.ID, 'cpf')
                element_value = self.scpc.web.find_element(By.ID, 'valor')
                element_name = self.scpc.web.find_element(By.ID, 'nome')
                cpfid = element_cpf.get_attribute('value')
                valor = element_value.get_attribute('value')
                nome = element_name.get_attribute('value')
                sj.createRegistro(cpfid, nome, et.configCard(contrato), valor)
                return 'REGISTRO MONTADO'

    def baixando_registros(self, contrato, valor):
        contratoativos = self.scpc.getListElement(By.XPATH, '//*[@id="tbl_fis"]/tbody/tr')

        for contratoincluso in contratoativos:

            if 'Nenhum registro encontrado' in contratoincluso.text:
                return 'NENHUM REGISTRO ENCONTRATO'

            while True:
                try:
                    card = contratoincluso.find_element(By.XPATH, './td[5]').text
                    value = contratoincluso.find_element(By.XPATH, './td[7]').text
                    break
                except:
                    sleep(2)

            cardpadrao = et.configCard(card)
            if et.configCard(contrato) == cardpadrao and value == valor:
                sleep(0.5)
                contratoincluso.find_element(By.XPATH, './td[9]').click()
                self.scpc.clickElement('//*[@id="btn_excluir"]')
                self.scpc.clickElement('/html/body/div[3]/div/div[3]/button[1]')
                sleep(4)
            try:
                info = self.scpc.getText('/html/body/div[3]/div/div[2]')
                if 'REGISTRO SUSPENSO' in info:
                    return 'REGISTRO ATIVO'
            except:
                return 'BAIXADO'

class Serasa:
    def __init__(self, user, password):
        self.user = user
        self.password = password

    def login(self):
        self.serasa = isw('https://empresas.serasaexperian.com.br/meus-produtos/login', True)
        self.serasa.writeText('//*[@id="loginUser"]', self.user)
        self.serasa.writeText('//*[@id="loginPassword"]', self.password)
        self.serasa.clickElement('//*[@id="loginFormSubmit"]')

        self.serasa.clickElement('//*[@id="mat-mdc-dialog-0"]/div/div/div/mat-dialog-content/div/div[2]/div/a')

    def consulta_cpf(self, cpf):
        self.serasa.clickElement('//*[@id="prod-64adce65a625ed4687bde841"]/div/div[3]/button')
        self.serasa.clearBoxTheText('//*[@id="debtorDocument"]')
        self.serasa.writeText('//*[@id="debtorDocument"]', cpf)
        self.serasa.clickElement('//*[@id="__next"]/main/div/form/div[1]/div[6]/button')

    def contrato_ativos(self):
        contratos_ativos = []
        try:
            self.serasa.loadingElement('//*[@id="__next"]/main/div/article/div/span[2]', 5)
            inf = self.serasa.web.find_element(By.XPATH, '//*[@id="__next"]/main/div/article/div/span[2]').text
            if inf.strip() == 'Nenhuma dívida encontrada':
                return 'Nenhuma dívida encontrada'
        except:
            pass
        contrato_inclusos = self.serasa.getListElement(By.XPATH, '//*[@id="__next"]/main/div/article/div/table/tbody/tr')
        for contrato in contrato_inclusos:
            contrato_incluso = contrato.find_element(By.XPATH, './td[3]').text
            vencimento_contrato = contrato.find_element(By.XPATH, './td[2]').text #VALIDAR
            value_contrato = contrato.find_element(By.XPATH, './td[4]').text #VALIDAR
            dionario_ativo = {
                'CONTRATO':et.configCard(contrato_incluso),
                'VENCIMENTO': vencimento_contrato,
                'VALOR': value_contrato
            }
            contratos_ativos.append(dionario_ativo)
        return contratos_ativos

    def baixa_contrato(self, contrato):
        try:
            self.serasa.loadingElement('//*[@id="__next"]/main/div/article/div/span[2]', 5)
            inf = self.serasa.web.find_element(By.XPATH, '//*[@id="__next"]/main/div/article/div/span[2]').text
            if inf.strip() == 'Nenhuma dívida encontrada':
                return 'Nenhuma dívida encontrada'
        except:
            pass
        contrato_inclusos = self.serasa.getListElement(By.XPATH,
                                                       '//*[@id="__next"]/main/div/article/div/table/tbody/tr')
        if len(contrato_inclusos) > 1:
            return 'VALIDAR'

        for contrato in contrato_inclusos:
            contrato_incluso = contrato.find_element(By.XPATH, './td[3]').text
            if et.configCard(contrato_incluso) == et.configCard(contrato):
                contrato.find_element(By.XPATH, './td[7]/div/button[1]').click()
                self.serasa.clickElement('//*[@id="modal"]/div[1]/div/div/div[2]/div/div/div[2]/button')
                self.serasa.clickElement('/html/body/div[2]/div[1]/div/div/div[2]/form/div[1]/label[13]/input')
                self.serasa.clickElement('//*[@id="modal"]/div[1]/div/div/div[2]/form/div[2]/div[2]')
                return 'EXCLUSO'
            else:
                return 'VERIFICAR'

    def baixa_contrato_valor(self, contrato, valor):
        try:
            self.serasa.loadingElement('//*[@id="__next"]/main/div/article/div/span[2]', 5)
            inf = self.serasa.web.find_element(By.XPATH, '//*[@id="__next"]/main/div/article/div/span[2]').text
            if inf.strip() == 'Nenhuma dívida encontrada':
                return 'Nenhuma dívida encontrada'
        except:
            pass
        contrato_inclusos = self.serasa.getListElement(By.XPATH,'//*[@id="__next"]/main/div/article/div/table/tbody/tr')


        for contrato in contrato_inclusos:
            contrato_incluso = contrato.find_element(By.XPATH, './td[3]').text #VALIDAR
            value_contrato = contrato.find_element(By.XPATH, './td[4]').text #VALIDAR
            if et.configCard(contrato_incluso) == et.configCard(contrato) and value_contrato == valor:
                contrato.find_element(By.XPATH, './td[7]/div/button[1]').click()
                self.serasa.clickElement('//*[@id="modal"]/div[1]/div/div/div[2]/div/div/div[2]/button')
                self.serasa.clickElement('/html/body/div[2]/div[1]/div/div/div[2]/form/div[1]/label[13]/input')
                self.serasa.clickElement('//*[@id="modal"]/div[1]/div/div/div[2]/form/div[2]/div[2]')
                return 'EXCLUSO'
            else:
                return 'VERIFICAR'


class Spc:
    def __init__(self, user, chave, password):
        self.user = user
        self.chave = chave
        self.password = password

    def login(self):
        self.spc = isw('https://sistema.spc.org.br/spc/controleacesso/autenticacao/entry.action;jsessionid=5ed367a9-15f7-4720-bfca-eff5e8f1875a_node186', True)
        self.spc.writeText('//*[@id="j_username"]', self.user)
        self.spc.writeText('//*[@id="j_password"]', self.chave)
        self.spc.clickElement('//*[@id="submitButton"]/span')
        self.spc.writeText('//*[@id="passphrase"]', self.password)
        self.spc.clickElement('//*[@id="submitButton"]/span')

    def __tela_consulta(self):
        self.spc.clickElement('/html/body/section[2]/div/table/tbody/tr[2]/td[2]/div/a/div/figure/img')
        self.spc.clickElement('//*[@id="accordion2"]/li[8]/a')

    def consulta_cpf(self, cpf):
        self.spc.clickElement('//*[@id="accordion2"]/li[8]/ul/li/a')
        self.spc.clickElement('//*[@id="m50"]/div/a')

        self.spc.writeText('//*[@id="numeroDocumento"]', cpf)
        self.spc.clickElement('//*[@id="conteudoInsumo"]/tbody/tr[2]/td/table/tbody/tr/td/input[1]')

    def baixa_contrato(self, contrato):
        try:
            self.spc.loadingElement('/html/body/table/tbody/tr/td[2]/div[2]/form/table/tbody/tr[1]/td/div/i', 10)
            notinc = self.spc.web.find_element(By.XPATH, '/html/body/table/tbody/tr/td[2]/div[2]/form/table/tbody/tr[1]/td/div/i').text
            if 'Nenhum registro de SPC' in notinc:
                return 'EXCLUSO'
        except:
            pass

        self.spc.loadingElement('//*[@id="dataGrid"]/tbody', 100)
        contratos_inclusos = self.spc.getListElement(By.XPATH, '//*[@id="dataGrid"]/tbody')

        if len(list) > 2:
            return 'VERIFICAR'

        for elementos in contratos_inclusos:
            contratosativo = elementos.find_elements(By.XPATH, './tr')
            for contratoativo in contratosativo:
                card = et.configCard(contratoativo.find_element(By.XPATH, './td[7]').text)
                if card == et.configCard(contrato):
                    contratoativo.find_element(By.XPATH, './td[7]').click()
                    self.spc.clickElement('//*[@id="motivoExclusaoRegistro.id"]/option[14]')
                    self.spc.clickElement('//*[@id="formSPC"]/table[3]/tbody/tr/td/input[2]')
                    self.spc.acceptAlert()
                    return 'EXCLUSO'


