from argparse import Action
from lib2to3.pgen2 import driver
from lib2to3.pgen2.driver import Driver
from multiprocessing.dummy import active_children
from webbrowser import Chrome
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import time

class Navigator():

    def __init__(self):
        chrome_options = webdriver.ChromeOptions()
        prefs = {'download.default_directory' : 'C:\\Users\\administrator.PLANBRAZIL\\Music\\SAP'}
        chrome_options.add_experimental_option('prefs', prefs)
        driver = webdriver.Chrome(executable_path= "C:\\Users\\administrator.PLANBRAZIL\\AppData\\Roaming\\Microsoft\\Windows\\Start Menu\\Programs\\Python 3.8\\chromedriver.exe")#self.
        self.driver = webdriver.Chrome(chrome_options=chrome_options)
        

    def get_site(self, site):
        self.driver.get(site)
        self.driver.maximize_window()

    def close(self):
        self.driver.close()
    
    def quit(self):
        self.driver.quit()

    def find_element(self, field):
        print('find element parte do codigo')#teste
        return self.driver.find_element(By.ID, field)
        
    def find_fill_id(self, field, text):
        self.driver.find_element(By.ID, field).send_keys(text)
        return self.driver.find_element(By.ID, field)

    def find_bt_id(self, field, tempo):
        self.driver.find_element(By.ID, field).click()
        time.sleep(tempo)

    def command(self, tecla, tempo):
        self.action = ActionChains(self.driver)
        self.action.send_keys(Keys.SHIFT, tecla).perform()
        time.sleep(tempo)

    def handle_window(self):
        self.driver.switch_to.window(self.driver.window_handles[1])
        self.close()
        self.driver.switch_to.window(self.driver.window_handles[0])
        time.sleep(5)

        # * A funçao Iframe_handle é um pouco mais complicada, ele procura uma camada nomeada <iframe> especifica dentro do HTML do SAP...
        # * para ir para ele, então ai ser possivel "tocar" o objeto
    def iframe_handle(self, iframe):
        self.main_window = self.driver.current_window_handle
        #Seleciona iFrame aonde fica o campo para digitar a transação
        try:
            self.iframeSAP = self.driver.find_element(By.XPATH, iframe)
        except NoSuchElementException:
            print("Elemento não encontrado")
        self.iframeSAP = self.driver.find_element(By.XPATH, iframe)
        self.driver.switch_to.frame(self.iframeSAP)
        time.sleep(15)

    def iframe_switch(self, iframe):
        self.transition_window = self.driver.current_window_handle
        self.driver.switch_to.window(self.transition_window)
        #Seleciona iFrame aonde fica o campo para digitar a transação
        try:
            self.iframeSAP = self.driver.find_element(By.ID, iframe)
        except NoSuchElementException:
            print("Elemento não encontrado")
        self.iframeSAP = self.driver.find_element(By.ID, iframe)
        self.driver.switch_to.frame(self.iframeSAP)
        time.sleep(15)
