import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import numpy as np
from datetime import datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import os
import glob
import chromedriver_autoinstaller
import datetime

scope = ['https://www.googleapis.com/auth/spreadsheets',
            "https://www.googleapis.com/auth/drive"]

credentials = ServiceAccountCredentials.from_json_keyfile_name("service_account_cemag.json", scope)
client = gspread.authorize(credentials)
filename = 'service_account_cemag.json'
sa = gspread.service_account(filename)

sheet = 'RQ EP-005-000 (Romaneios)'
#worksheet= input('Nome da aba:')
worksheet = 'FT10500 SS T BB M23' # Local para fazer alteração do nome da planilha

sh1 = sa.open(sheet)
wks1 = sh1.worksheet(worksheet)

df = wks1.get()

tabela = pd.DataFrame(df)

def iframes(nav):

    iframe_list = nav.find_elements(By.CLASS_NAME,'tab-frame')

    for iframe in range(len(iframe_list)):
        time.sleep(1)
        try: 
            nav.switch_to.default_content()
            nav.switch_to.frame(iframe_list[iframe])
            print(iframe)
        except:
            pass

def saida_iframe(nav):
    nav.switch_to.default_content()

def listar(nav, classe):
    
    lista_menu = nav.find_elements(By.CLASS_NAME, classe)
    
    elementos_menu = []

    for x in range (len(lista_menu)):
        a = lista_menu[x].text
        elementos_menu.append(a)

    test_lista = pd.DataFrame(elementos_menu)
    test_lista = test_lista.loc[test_lista[0] != ""].reset_index()

    return(lista_menu, test_lista)

nav = webdriver.Chrome(r'C:\\Users\\TI\\anaconda3\\lib\\site-packages\\chromedriver_autoinstaller\\114\\chromedriver.exe')
time.sleep(1)
nav.maximize_window()
time.sleep(1)
nav.get('http://192.168.3.141/sistema')

nav.find_element(By.ID, 'username').send_keys('ti.dev') #ti.dev
time.sleep(2)
nav.find_element(By.ID, 'password').send_keys('cem@#161010')
time.sleep(1)
nav.find_element(By.ID, 'submit-login').click() 
WebDriverWait(nav,20).until(EC.presence_of_element_located((By.ID, 'bt_1892603865')))
time.sleep(1)
nav.find_element(By.ID, 'bt_1892603865').click()
time.sleep(3)
lista_menu, test_list = listar(nav, 'webguiTreeNodeLabel')
time.sleep(1)

click_producao = test_list.loc[test_list[0] == 'Projeto'].reset_index(drop=True)['index'][0]
lista_menu[click_producao].click()
time.sleep(2)

lista_menu, test_list = listar(nav, 'webguiTreeNodeLabel')
time.sleep(1)
click_producao = test_list.loc[test_list[0] == 'Materiais e Produtos'].reset_index(drop=True)['index'][0]
lista_menu[click_producao].click()
time.sleep(6)

iframes(nav)
WebDriverWait(nav,20).until(EC.presence_of_element_located((By.XPATH, '/html/body/table/tbody/tr[2]/td/div/form/table/thead/tr[2]/td[1]/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div')))
time.sleep(1)
nav.find_element(By.XPATH,'/html/body/table/tbody/tr[2]/td/div/form/table/thead/tr[2]/td[1]/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div').click()
time.sleep(3)
input_localizar = nav.find_element(By.ID, 'grInputSearch_explorer')
time.sleep(1.5)
input_localizar.send_keys('000000')
time.sleep(1.5)
input_localizar.send_keys(Keys.ENTER)
time.sleep(1.5)
nav.find_element(By.XPATH,'/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[4]/td[1]/input').click()
time.sleep(1.5)
nav.find_element(By.XPATH, '/html/body/table/tbody/tr[2]/td/div/form/table/thead/tr[2]/td[1]/table/tbody/tr/td[1]/table/tbody/tr/td[7]/div').click()
time.sleep(13)
nav.find_element(By.XPATH,'/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[4]/td[1]/input').click()
time.sleep(1.5)
nav.find_element(By.XPATH, '/html/body/table/tbody/tr[2]/td/div/form/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[1]/div').click()
time.sleep(2)

#-------------------------------------------------------------------------------------------
codigo_tabela = tabela.iloc[22,14]
codigo_tabela = codigo_tabela[0:6]
nome_generico = tabela.iloc[7,16]

codigo_inov = nav.find_element(By.XPATH, '//*[@id="explorer"]/tbody/tr[1]/td[1]/table/tbody/tr[2]/td/table/tbody/tr[3]/td[2]/table/tbody/tr/td[1]/input')
time.sleep(1.5)
codigo_inov.send_keys(Keys.CONTROL + 'a')
time.sleep(1.5)
codigo_inov.send_keys(Keys.BACKSPACE)
time.sleep(1.5)
codigo_inov.send_keys(codigo_tabela)
time.sleep(1.5)
nav.find_element(By.XPATH,'//*[@id="explorer"]/tbody/tr[1]/td[1]/table/tbody/tr[2]/td/table/tbody/tr[5]/td[2]/table/tbody/tr/td[1]/input').send_keys(Keys.CONTROL + 'a')
time.sleep(1.5)
nav.find_element(By.XPATH,'//*[@id="explorer"]/tbody/tr[1]/td[1]/table/tbody/tr[2]/td/table/tbody/tr[5]/td[2]/table/tbody/tr/td[1]/input').send_keys(Keys.BACKSPACE)
time.sleep(1.5)
nav.find_element(By.XPATH,'//*[@id="explorer"]/tbody/tr[1]/td[1]/table/tbody/tr[2]/td/table/tbody/tr[5]/td[2]/table/tbody/tr/td[1]/input').send_keys(codigo_tabela)
time.sleep(1.5)
nome_inov = nav.find_element(By.XPATH,'//*[@id="explorer"]/tbody/tr[1]/td[1]/table/tbody/tr[2]/td/table/tbody/tr[3]/td[4]/table/tbody/tr/td[1]/input')
nome_inov.send_keys(Keys.CONTROL + 'a')
time.sleep(1.5)
nome_inov.send_keys(Keys.BACKSPACE)
time.sleep(1.5)
nome_inov.send_keys(nome_generico)
time.sleep(1.5)
nav.find_element(By.XPATH,'//*[@id="explorer"]/tbody/tr[1]/td[1]/table/tbody/tr[2]/td/table/tbody/tr[5]/td[4]/table/tbody/tr/td[1]/input').send_keys(Keys.CONTROL + 'a')
time.sleep(1.5)
nav.find_element(By.XPATH,'//*[@id="explorer"]/tbody/tr[1]/td[1]/table/tbody/tr[2]/td/table/tbody/tr[5]/td[4]/table/tbody/tr/td[1]/input').send_keys(Keys.BACKSPACE)
time.sleep(1.5)
nav.find_element(By.XPATH,'//*[@id="explorer"]/tbody/tr[1]/td[1]/table/tbody/tr[2]/td/table/tbody/tr[5]/td[4]/table/tbody/tr/td[1]/input').send_keys(nome_generico)
time.sleep(2)
nav.find_element(By.XPATH,'//*[@id="gridTitle_explorer_RECURSOETAPA"]').click()
time.sleep(1.5)
nav.find_element(By.XPATH,'/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[15]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[1]/div').click()
time.sleep(1.5)
nav.find_element(By.XPATH,'/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[15]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[15]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[8]').click()
time.sleep(1.5)
#---------------------------------------------------------------------------------------------------------------------------------------------------------------

planilha = tabela.iloc[22:, 14:21]
cabecalho = wks1.row_values(9)
cabecalho = cabecalho[14:21]
planilha = planilha.set_axis(cabecalho,axis=1)

indice_linha_descricao = planilha[planilha['Descrição'] == 'Descrição'].index[0]

planilha = planilha.loc[:indice_linha_descricao - 1]

planilha.fillna('',inplace=True)

planilha = planilha[(planilha['Código'] != '')]

planilha.reset_index(drop=True,inplace=True)
# ----------------------------------------------------------------------------------------------------------------------
nav.find_element(By.XPATH,'/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[15]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[15]/td/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[2]/div').click()
time.sleep(1.5)
linha = 3
var = 0

for i in range(len(planilha)):

    if var == 10 and linha == 23:
        var -= 1
        linha -= 2

    time.sleep(2)
    nav.find_element(By.XPATH,'//*[@id="explorer_RECURSOETAPA_ETAPARECURSOS"]/tbody/tr[1]/td[1]/table/tbody/tr['+ str(linha) +']/td[2]').click()
    time.sleep(1.5)
    nav.find_element(By.XPATH,'//*[@id='+ str(var) +']/td[2]/div/input').send_keys(planilha['Ordem'][i])
    time.sleep(1.5)

    nav.find_element(By.XPATH,'//*[@id="explorer_RECURSOETAPA_ETAPARECURSOS"]/tbody/tr[1]/td[1]/table/tbody/tr['+ str(linha) +']/td[3]').click()
    time.sleep(1.5)
    nav.find_element(By.XPATH,'//*[@id='+ str(var) +']/td[3]/div/input').send_keys(planilha['Código'][i])
    time.sleep(1.5)
    
    nav.find_element(By.XPATH,'//*[@id="explorer_RECURSOETAPA_ETAPARECURSOS"]/tbody/tr[1]/td[1]/table/tbody/tr['+ str(linha) +']/td[5]').click()
    time.sleep(1.5)
    nav.find_element(By.XPATH,'//*[@id='+ str(var) +']/td[5]/div/input').send_keys(planilha['SUM de Qtd'][i])
    time.sleep(1.5)
    
    nav.find_element(By.XPATH,'//*[@id="explorer_RECURSOETAPA_ETAPARECURSOS"]/tbody/tr[1]/td[1]/table/tbody/tr['+ str(linha) +']/td[7]').click()
    time.sleep(1.5)
    nav.find_element(By.XPATH,'//*[@id='+ str(var) +']/td[7]/div/input').send_keys(planilha['Depósito'][i])
    time.sleep(3)

    nav.find_element(By.XPATH,'//*[@id='+ str(var) +']/td[7]/div/input').send_keys(Keys.TAB)
    time.sleep(1)
    nav.find_element(By.XPATH,'//*[@id='+ str(var) +']/td[9]/div/input').send_keys(Keys.TAB)
    time.sleep(1)
    nav.find_element(By.XPATH,'//*[@id='+ str(var) +']/td[10]/div/input').send_keys(Keys.TAB)
    time.sleep(1)
    nav.find_element(By.XPATH,'//*[@id='+ str(var) +']/td[11]/div/input').send_keys(Keys.TAB)

    time.sleep(1.5)
    nav.find_element(By.XPATH,'//*[@id='+ str(var) +']/td[13]/div/textarea').send_keys(planilha['Observação'][i])
    time.sleep(1.5)
    nav.find_element(By.XPATH,'//*[@id="explorer_RECURSOETAPA_ETAPARECURSOS"]/tbody/tr[1]/td[1]/table/tbody/tr['+ str(linha) +']/td[2]').click()
    nav.find_element(By.XPATH,'//*[@id="explorer_RECURSOETAPA_ETAPARECURSOS"]/tbody/tr[1]/td[1]/table/tbody/tr['+ str(linha) +']/td[2]').click()
    time.sleep(1.5)
    nav.find_element(By.XPATH,'//*[@id='+ str(var) +']/td[2]/div/input').send_keys(Keys.CONTROL + 'm')
    time.sleep(6)

    if i+1 == planilha.shape[0]:
        pass
    else:
        nav.find_element(By.XPATH,'//*[@id='+ str(var) +']/td[2]/div/input').send_keys(Keys.INSERT)
        nav.find_element(By.XPATH,'//*[@id='+ str(var) +']/td[2]/div/input').send_keys(Keys.INSERT)

    # nav.find_element(By.XPATH,'//*[@id='+ str(var) +']/td[13]/div/div').click()
    # time.sleep(1.5)
    # nav.find_element(By.XPATH,'//*[@id='+ str(var) +']/td[13]/div/textarea').send_keys(Keys.CONTROL + 'm')
    # time.sleep(4)
    # nav.find_element(By.XPATH,'//*[@id='+ str(var) +']/td[13]/div/div').click()
    # time.sleep(1.5)
    # nav.find_element(By.XPATH,'//*[@id='+ str(var) +']/td[13]/div/div').send_keys(Keys.INSERT)
    
    var += 1
    linha += 2
    
    time.sleep(5)

nav.close()