from time import sleep
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.common.keys import Keys

class initSeWeb():
    def __init__(self, url:str, view:bool=True):
        options = Options()
        options.add_argument('--headless=new')
        if view == True:
            self.web = webdriver.Chrome()
            self.web.get(url)
        else:
            self.web = webdriver.Chrome(options)
            self.web.get(url)

    def _loadingWeb(self, xpatch:str):
        WebDriverWait(self.web, 1000).until(EC.visibility_of_element_located((By.XPATH, xpatch)))

    def loadingElement(self, xpatch:str, time:int):
        WebDriverWait(self.web, time).until(EC.visibility_of_element_located((By.XPATH, xpatch)))

    def clickElement(self, xpatch:str):
        while True:
            try:
                initSeWeb._loadingWeb(self, xpatch)
                self.web.find_element(By.XPATH, xpatch).click()
                break
            except :
                sleep(3)

    def writeText(self, xpatch:str, text):
        while True:
            try:
                initSeWeb._loadingWeb(self, xpatch)
                self.web.find_element(By.XPATH, xpatch).send_keys(text)
                break
            except Exception as e:
                print('Deu não ',e)
                sleep(3)

    def readText(self, xpatch:str):
        while True:
            try:
                initSeWeb._loadingWeb(self, xpatch)
                text = self.web.find_element(By.XPATH, xpatch).text
                break
            except:
                sleep(3)
        return text

    def clearBoxTheText(self, xpatch:str):
        initSeWeb._loadingWeb(self, xpatch)
        self.web.find_element(By.XPATH, xpatch).clear()

    def getListElement(self, by, xpatch):
        initSeWeb._loadingWeb(self, xpatch)
        return self.web.find_elements(by, xpatch)

    def entryIframe(self, xpatch:str):
        self.onIframe = False
        while True:
            try:
                initSeWeb._loadingWeb(self, xpatch)
                iframe = self.web.find_element(By.XPATH, xpatch)
                self.web.switch_to.frame(iframe)
                break
            except:
                print('erro')
                sleep(3)

    def exitIframe(self):
        self.web.switch_to.default_content()

    def entryWindow(self, window=None):
        if window != None:
            self.web.switch_to.window(window)
        else:
            self.windows = self.web.window_handles
            self.web.switch_to.window(self.windows[1])

    def exitWindow(self):
        self.web.switch_to.window(self.windows[0])

    def sendKeyEnter(self, xpatch):
        element = self.web.find_element(By.XPATH, xpatch).send_keys(Keys.ENTER)

    def acceptAlert(self):
        WebDriverWait(self.web, 60).until(EC.alert_is_present())
        alerta = Alert(self.web)
        alerta.accept()

    def closeWindow(self):
        self.web.close()