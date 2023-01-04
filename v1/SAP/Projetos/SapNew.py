from core import Navigator
from selenium.webdriver.common.keys import Keys
from datetime import date, timedelta
import shutil
import time
import os

yesterday = date.today()# - timedelta(days=1)
YESTERDAY_DAY = yesterday.strftime('%d.%m.%y')
HOUR_DAY = time.strftime('_%H.%M')

NAME = YESTERDAY_DAY + HOUR_DAY

SITE = "https://saperpprod.XXXX"
SAP = "C:\\Users\\XXXXXXXX"
ACCOUNT_1 = "C:\\Users\\XXXXXXXXXXX"
TIPO1 = "_AC"
BASIC_1 = "C:\\Users\\XXXXXXXXXXX"
TIPO2 = "_BA"

USER = "sap-user"
PASS = "sap-password"
LOGON = "b1"

BARCODE = 'M0:D:10::okcd'

PURCHASE = "M0:U:::1:34"
DATE1 = "M0:U:::11:34"
DATE2 = "M0:U:::11:59"
EXECUTE = "M0:D:13::btn[8]"

SPREEDSHEET = "M0:D:13::btn[43]"
OPENER = "RCua2OldToolbar-hiddenBtn" 
ACCOUNT = "M0:D:13::btn[24]-BtnMenu"
BASIC = "cua2sapmenu_btn"
#"M0:D:13::btn[7]-BtnMenu"

MAINFRAME = '//*[@id="ITSFRAME1"]'
#'ITSFRAME1'

IFRAME1 = 'URLSPW-0'
IFRAME2 = 'URLSPW-1'

CONFIRM = "M1:D:13::btn[0]"
POPUP = "popupDialogInputField"
CHOOSE = "UpDownDialogChoose"
OK = "PromptDialogOk"

class Sap(Navigator): # * Classe herdando atributos da Class-Mãe Navigator vindo do arquivo Core

    def __init__(self):
        self.n = Navigator()
        self.n.get_site(SITE)

    def login(self, username, password): # * Função para logar no SAP.
        self.n.find_fill_id(USER, username)
        self.a = self.n.find_fill_id(PASS, password)
        self.a.send_keys(Keys.RETURN)
        time.sleep(30)
        
    def move_download(self, path, type): # * Função para checar se a pasta tem algo e mover o arquivo baixado.
        os.chdir(SAP)
        data = os.listdir()
        if len(data) > 0:
            shutil.move(SAP+data[0], path + NAME + type + ".XLS")

    def transaction(self, transacao): # * Função para selecionar a barra de pesquisa e digitar a transação a ser usada.
        self.n.iframe_handle(MAINFRAME)

        self.a = self.n.find_element(BARCODE)
        self.n.find_fill_id(BARCODE, transacao)
        self.a.send_keys(Keys.RETURN)
        time.sleep(20)

    def purchase(self): # * A compra requer alguns parametros, essa função cria esse ambiente ideal
        self.compra = self.n.find_element(PURCHASE)
        i = 0
        while i <= 2:
            self.compra.send_keys(Keys.BACKSPACE)
            i += 1

        self.n.find_fill_id(DATE1, YESTERDAY_DAY)
        self.n.find_fill_id(DATE2, YESTERDAY_DAY)
        self.n.find_bt_id(EXECUTE, 30)

    def spreadsheet(self): # * A partir daqui é feito o primeiro download e são a mistura de algumas funções previamentes vistas

        #self.n.iframe_handle(MAINFRAME) #teste
        #self.a = self.n.find_element(BARCODE) #teste

        self.n.find_bt_id(OPENER, 5) 
        self.n.find_bt_id(BASIC, 10)

        self.n.find_bt_id(SPREEDSHEET, 30)
        self.n.iframe_switch(IFRAME1)

        self.n.find_bt_id(CONFIRM, 30)
        self.n.iframe_switch(IFRAME2)

        self.popup = self.n.find_element(POPUP)
        i = 0
        while i <= 9:
            self.popup.send_keys(Keys.BACKSPACE)
            i += 1
        self.popup.send_keys(YESTERDAY_DAY+"01.XLS")

        self.n.find_bt_id(CHOOSE, 30)

        self.n.find_bt_id(OK, 20)

        self.n.handle_window()

        self.move_download(BASIC_1, TIPO2)

        time.sleep(10)
    

    def spreadsheet_account(self):
        
        self.n.find_bt_id(OPENER, 10)
        self.n.find_bt_id(ACCOUNT, 10)

        self.n.find_bt_id(SPREEDSHEET, 30)
        self.n.iframe_switch(IFRAME1)


        self.n.find_bt_id(CONFIRM, 30)
        self.n.iframe_switch(IFRAME2)

        self.popup = self.n.find_element(POPUP)
        i = 0
        while i <= 9:
            self.popup.send_keys(Keys.BACKSPACE)
            i += 1
        self.popup.send_keys(YESTERDAY_DAY+".XLS")

        self.n.find_bt_id(CHOOSE, 20)

        self.n.find_bt_id(OK, 20)

        self.n.handle_window()
        #self.n.iframe_switch(IFRAME1) #pra tentar voltar pro IFRAME1
        #self.n.driver.switch_to.frame(IFRAME1) #mesmo que o de cima

        self.move_download(ACCOUNT_1, TIPO1) 

    def quit(self): # * Fecha o Driver, o que faz fechar o chrome
        self.n.quit()

    def return_ACCOUNT_1(self): # * Retorna o nome do arquivo que foi baixado, junto com o horario
        nomeAC = ACCOUNT_1 + NAME + TIPO1 + ".XLS"
        return nomeAC

    def return_BASIC_1(self): # * Retorna o nome do arquivo que foi baixado, junto com o horario 
        nomeBA = BASIC_1 + NAME + TIPO2 + ".XLS"
        return nomeBA