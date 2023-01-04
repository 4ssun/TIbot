#%%
from sqlite3 import Row
from SapNew import Sap
import win32com.client as win32
from datetime import datetime
import pandas as pd
import ExcelNew
import time
import sys
import os
from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import shutil
from importlib.resources import path
import glob
import openpyxl
# * Primeira parte do codigo Acessando o SAP e baixando os arquivos
#%%

print("---------------------FASE 1-----------------------------")
print("Começando Fase 1: " + str(datetime.now()))

sap = Sap()

sap.login("XXXXXXX", "XXXXXXXXX")

sap.transaction("me5a")

sap.purchase()

sap.spreadsheet()

sap.transaction("XXXXXXXXXXXX")

sap.purchase()

sap.spreadsheet_account()

print('finalizando sap' + str(datetime.now()))


print("Terminando Fase 1: " + str(datetime.now()))

# * Segunda parte, começando a fusão dos excels

print("\n")
print("---------------------FASE 2-----------------------------")
print("Começando Fase 2: " + str(datetime.now()))

MainRBS = "C:\\Users\\XXXXX"
ArquivoAccount = sap.return_ACCOUNT_1()
ArquivoBasic = sap.return_BASIC_1()

sap.quit()

#If erro:  Remove-Item -path $env:LOCALAPPDATA\Temp\gen_py -recurse -> executar codigo com powershell
xl = win32.gencache.EnsureDispatch("Excel.Application")
xl.Visible = False
xl.WindowState = win32.constants.xlMaximized
XlDirectionDown = win32.constants.xlDown

ac = xl.Workbooks.Open(ArquivoAccount)
bc = xl.Workbooks.Open(ArquivoBasic)
rbs = xl.Workbooks.Open(MainRBS)

ac = xl.Workbooks.Open(ArquivoAccount)
CostCenter = xl.Range("D:D").Select()
CC = xl.Selection.Copy(Destination= bc.Worksheets("Sheet1").Range("X1"))

ac = xl.Workbooks.Open(ArquivoAccount)
Wbs = xl.Range("E:E").Select()
WW = xl.Selection.Copy(Destination= bc.Worksheets("Sheet1").Range("Y1"))

bc.Save()
xl.Workbooks.Close()
xl.Workbooks.Close()

print("Terminando Fase 2: " + str(datetime.now()))


# * Terceira parte, jogando os dados do Excel merged no Sharepoint

print("\n")
print("---------------------FASE 3-----------------------------")
print("Começando Fase 3: " + str(datetime.now()))

##################### MEU CODIGO 
# VARIÁVEIS GLOBAIS COM LINKS E BOTÕES #
#%%
SHAREURL = "https://XXXXXXXXXXXXXXXX"
SHAREURL2 = "https://XXXXXXXXXX"
#BOTÕES DE LOGIN MICROSOFT#
LOGIN_ID = 'i0116'
LOGIN_PSW = 'i0118'
LOGIN_OK = 'idSIButton9'
LOGIN_CONTINUE = 'idBtn_Back'
#BOTÕES DE OKTA LOGIN#
OKTA_EMAIL = 'okta-signin-username'
OKTA_PSWRD = 'okta-signin-password'
OKTA_ACCESS = 'okta-signin-submit'
#BOTÕES DE NOVO FORMS E SALVAR DO SHAREPOINT1#
BUTTONN_NEW = 'root-81' 
BUTTONN_SAVE = '/html/body/div/div/div[2]/div[2]/div/div[4]/div/div/div[4]/div/div/div[2]/div/div/div/div/div[5]/div/button'
# PATHS DA BIBLIOTECA #
PATHLIB = 'C:\\Users\\XXXXX'
ATUAL = glob.glob(PATHLIB + '*.xls')
PROCESSADO = 'C:\\Users\\XXXXX'
# Campos que são trabalhados no lançamento das listas do excel -> sharepoint
campos_lista_sharepoint = [1,2,3,4,5,6,7,8,9,11,12,13,14,15]
#campos_lista_sharepoint = [2,3,4,5,6,7,8,9,10,12,13,14,15,16]
#colunas_arquivo_excel = (4, 5, 6, 7, 9, 10, 11, 12, 13, 15, 22, 23, 24, 25)
colunas_arquivo_excel = (3, 4, 5, 6, 8, 9, 10, 11, 12, 14, 21, 22, 23, 24)

''' Explicando os arquivos em Excel e suas variáveis declaradas'''
#Arquivo de USUÁRIO/EMAILS do sap
sap_emails = pd.read_excel('C:\\Users\\XXXXXXXX')
#%% Essa parte agora vai pegar e definir todos os arquivos excel que serão usados e refinados antes de começar seu devido uso.
for sap_arquivo in ATUAL:
    RBS = pd.read_html(sap_arquivo)[1]
    RBS.to_excel('C:\\Users\\XXXXXX')
    sap_excel = RBS
    #sap_excel = sap_excel.drop(sap_excel.columns[0:1], axis= 'columns')    
    sap_excel.columns = ['Purch.  Organization', 'Purcha= sing Group', 'Release  indicator','Purchase', 'Requis= ition Date', 'Created','Changed=  on','Releas= e Strategy', 'Item of  Requisition','Short=  Text', 'Quanti= ty  Requested', 'Unit of=  Measure','Materia= l Group', 'Deleti= on  Indicator', 'Valuati= on Price','Goods R= eceipt', 'Invoice=  Receipt', 'Purchas= e Order','Purcha= se Order  Item', 'Purcha= se Order  Date','Quanti= ty Ordered', 'Total V= alue', 'Currenc= y', 'Cost Ce= nter','WBS E= lement']
    sap_excel = sap_excel.drop(sap_excel.index[0])
    
processadas = pd.read_excel('C:\\Users\\administrator.PLANBRAZIL\\Downloads\\Brazil IT - Tibot\\v1\\SAP\\Projetos\\processadas.xlsx')
processadas = processadas.drop(processadas.columns[0:1], axis= 'columns')
processadas.columns = ['Purch.  Organization', 'Purcha= sing Group', 'Release  indicator',
        'Purchase', 'Requis= ition Date', 'Created',
        'Changed=  on', 'Releas= e Strategy', 'Item of  Requisition',
        'Short=  Text', 'Quanti= ty  Requested', 'Unit of=  Measure',
        'Materia= l Group', 'Deleti= on  Indicator', 'Valuati= on Price',
        'Goods R= eceipt', 'Invoice=  Receipt', 'Purchas= e Order',
        'Purcha= se Order  Item', 'Purcha= se Order  Date',
        'Quanti= ty Ordered', 'Total V= alue', 'Currenc= y', 'Cost Ce= nter',
        'WBS E= lement']
processadas = processadas.drop(processadas.index[0])
lista_processadas = processadas['Purchase'].tolist()
lista_str_processadas = map(str, lista_processadas)
df_share1 = sap_excel.loc[~sap_excel['Purchase'].isin(lista_str_processadas)]

print(df_share1)  #é o que vai ser lançado =)'''
#%% Funções que realizam todo webscrapping
def Navigator(DRIVER,url):
    DRIVER.get(url)
    DRIVER.maximize_window()
    time.sleep(4)
def Login_Out(DRIVER,user):
    log = DRIVER.find_element(By.ID, LOGIN_ID)
    time.sleep(4)
    log.send_keys(user)
    time.sleep(4)
    prosseg = DRIVER.find_element(By.ID, LOGIN_OK)
    prosseg.click()
    time.sleep(4)
def Login_Okta(DRIVER,user,pswrd):    
    log_okta = DRIVER.find_element(By.ID,OKTA_EMAIL)
    log_okta.send_keys(user)
    pswrd_okta = DRIVER.find_element(By.ID,OKTA_PSWRD)
    pswrd_okta.send_keys(pswrd)
    acesso_okta = DRIVER.find_element(By.ID,OKTA_ACCESS)
    acesso_okta.click()
    time.sleep(25)
    prosseg_2 = DRIVER.find_element(By.ID,LOGIN_CONTINUE) 
    prosseg_2.click()
    time.sleep(25)
    #CLICK DO PRIMEIRO SHARE
def click_novo(DRIVER, BUTTONN_NEW):
    DRIVER.find_element(By.CLASS_NAME, BUTTONN_NEW).click()
    time.sleep(4)
def click_salva(DRIVER, BUTTONN_SAVE):
    DRIVER.find_element(By.XPATH, BUTTONN_SAVE).click()
    time.sleep(4)
def fecha(DRIVER):
    DRIVER.close()


#%%  Função que faz o tratamento do nome do usuário sap pelo seu email institucional #
sap_emails = pd.read_excel('C:\\Users\\XXXXXXX')
def tratamento_email(usuario):
      sap_emails = pd.read_excel('C:\\Users\\XXXXXX')
      for eml in range(len(sap_emails)):
         if usuario == sap_emails.iat[eml,0]:
          return(sap_emails.iat[eml,3])
#%% Função que faz o primeiro lançamento das RBS no sharepoint, usando o arquivo BRUTO do sap
def lancamento_share_linhas(DRIVER,df_share1):
    click_novo(DRIVER, BUTTONN_NEW)
    time.sleep(3)
    for linha in range(0,len(df_share1)):
        for (coluna,i) in zip(colunas_arquivo_excel,campos_lista_sharepoint):
            if(i==1):
                registro = DRIVER.find_elements(By.TAG_NAME, "input")
                registro[i].send_keys(df_share1.iat[linha,coluna]) 
            elif(i==3):
                registro[i].send_keys(tratamento_email(df_share1.iat[linha,coluna]))
            elif(i==15):
                if type(df_share1.iat[linha,coluna]) == float:
                    registro[i].send_keys('0000000')
                else:
                    registro[i].send_keys(df_share1.iat[linha,coluna])
            else:
                registro[i].send_keys(df_share1.iat[linha,coluna])   
        #click_salva(DRIVER, BUTTONN_SAVE)
        salva = WebDriverWait(DRIVER, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div/div/div[2]/div[2]/div/div[4]/div/div/div[4]/div/div/div[2]/div/div/div/div/div[5]/div/button')))
        time.sleep(4)
        actions = ActionChains(DRIVER)
        actions.click(salva).perform()
        time.sleep(4)
        if(linha<len(df_share1)-1):
            click_novo(DRIVER, BUTTONN_NEW)
            time.sleep(5)

#%% Função que faz o segundo lançamento das RBS no sharepoint já sem duplicatas de RBS's
def lancamento_share_linhas_2(driver,df_share1):
    campos_lista_sharepoint = [2,4]
    colunas_arquivo_excel = (3,5)
    registros = df_share1
    NOVO2 = driver.find_elements(By.XPATH, '//*[@id="appRoot"]/div[1]/div[2]/div[3]/div/div[2]/div[2]/div[2]/div[1]/div/div/div/div/div/div/div[1]/div[1]/button')
    NOVO2[0].click() #len é 1,  0 é o botão novo.
     #tava dentro de i==3
    time.sleep(4)
    for row in range(0,len(registros)):
        for (coluna,i) in zip(colunas_arquivo_excel,campos_lista_sharepoint):
            if (i==2):
                registrar = driver.find_elements(By.TAG_NAME,'input')
                registrar[i].send_keys(registros.iat[row,coluna])
            if (i==4):
                registrar[i].send_keys(tratamento_email(registros.iat[row,coluna]))
                #sugest = driver.find_element(By.XPATH, "//*[@id='sug-0']/span")
                # "//*[@type='button']//following::button[27]") #//*[@id="sug-0"]/span #/html/body/div[5]/div/div/div/div/div/div[2]/div/div/button/span
                sugest = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='sug-0']/span")))
                time.sleep(4)
                actions = ActionChains(driver)
                actions.click(sugest).perform()
            
        time.sleep(10)
        save2 = driver.find_elements(By.XPATH, '//*[@id="appRoot"]/div[2]/div/div[4]/div/div/div[4]/div/div/div[2]/div/div/div/div/div[5]/div/button')
        #print(len(save2))
        save2[0].click()  #resolver botao de save, problema tá no email que ta vindo
        time.sleep(3)
        if(row<len(registros)-1):
            novo2 = driver.find_elements(By.XPATH, '//*[@id="appRoot"]/div[1]/div[2]/div[3]/div/div[2]/div[2]/div[2]/div[1]/div/div/div/div/div/div/div[1]/div[1]/button')
            novo2[0].click()
        time.sleep(4)
    fecha(DRIVER)
    #"//*[@type='button']//following::button[27]"
#%%     PRIMEIRO WEBSCRAPING
DRIVER = webdriver.Chrome(executable_path="C:\\Users\\XXXXXX\\chromedriver.exe")
Navigator(DRIVER, SHAREURL)
Login_Out(DRIVER,'XXXXXXXXXXX')
Login_Okta(DRIVER,'XXXXXXXXXXX','XXXXXXXX')
time.sleep(15)
#%%     PRIMEIRO LANÇAMENTO SHAREPOINT
lancamento_share_linhas(DRIVER,df_share1)
print('----- ACABANDO O PRIMEIRO LANÇAMENTO -----')
time.sleep(5)
fecha(DRIVER)
#%%     SEGUNDO WEBSCRAPING
DRIVER = webdriver.Chrome(executable_path="C:\\Users\\XXXXXXXXXXX\\chromedriver.exe")
Navigator(DRIVER, SHAREURL2)
Login_Out(DRIVER,'XXXXXXXXXXX')
time.sleep(5)
Login_Okta(DRIVER,'XXXXXXXXXXXX','XXXXXXXXX')
#time.sleep(3)
wait = WebDriverWait(DRIVER, 5)
element = wait.until(EC.element_to_be_clickable((By.TAG_NAME,'input')))

#%%     SEGUNDO LANÇAMENTO SHAREPOINT
df2 = df_share1.drop_duplicates(subset=['Purchase'],keep='first')
registro=df2
#display(registro)
lancamento_share_linhas_2(DRIVER,registro)

# %%    ATUALIZA BASE PROCESSADAS
df_registro = registro
df_processadas = processadas
df_finalizadas = df_processadas.append(registro)
df_finalizadas.to_excel('processadas.xlsx')
df_finalizadas = df_finalizadas.drop(df_finalizadas.index[0])
#display(df_finalizadas)

# %%    MOVE O ARQUIVO SAP PRO DIRETÓRIO DE PROCESSADAS
def mover_diretorio(PATHLIB, PROCESSADO):
    arquivo = os.listdir(PATHLIB)[0]
    origem = PATHLIB + arquivo
    destino = PROCESSADO + arquivo
    shutil.move(origem, PROCESSADO)
    
mover_diretorio(PATHLIB, PROCESSADO)


# %%
#display(df_finalizadas)
print(df_finalizadas)
# %%

