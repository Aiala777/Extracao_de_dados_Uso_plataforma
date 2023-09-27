from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
import time, subprocess
from datetime import datetime, timedelta
from enviooutlook import envio, emailInsusesso
import os, shutil
from planilha import dados




def acessar():
    caminho_destino = r'C:/Users/arthuraiala/Desktop/uso_gup/'
    data_atual = datetime.now()
    today = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    print('Iniciando: ', today)
    options = Options()
    options.add_experimental_option("excludeSwitches", ["enable-logging"])
    options.add_experimental_option("detach", True)

    global navegador
    navegador = webdriver.Chrome(service=Service(executable_path='chromedriver.exe'), options=options)

    #ENTRAR NO SITE 
    navegador.get("https://chat.sonax.net.br/#/app/chat")
    time.sleep(5)
    try:
        user = navegador.find_element(By.XPATH, '/html/body/app-root/app-login/div/div[1]/div/div[1]/input') #Login
        user.send_keys("")
        senha = navegador.find_element(By.XPATH, '/html/body/app-root/app-login/div/div[1]/div/div[2]/input') #Senha
        senha.send_keys("")
        navegador.find_element(By.XPATH, '/html/body/app-root/app-login/div/div[1]/div/button').click() #Entrar
        time.sleep(5)
        navegador.get("https://chat.sonax.net.br/#/app/chatbox")
        time.sleep(5)
    except:
        navegador.get("https://chat.sonax.net.br/#/app/chatbox")
    
    #seleciona Whatsapp bussines
    navegador.find_element(By.XPATH, '/html/body/app-root/app-layout/ng-sidebar-container/div/div/app-chatbox/app-content/main/div/app-portlet/div/div[2]/div/div[1]/div/div/div/div[3]/div[3]/div/select/option[4]').click()
    # #clica na data
    # navegador.find_element(By.XPATH,"/html/body/app-root/app-layout/ng-sidebar-container/div/div/app-chatbox/app-content/main/div/app-portlet/div/div[2]/div/div[1]/div/div/div/div[6]/div/ngb-datepicker-input/input").click()
    # time.sleep(3)
    # # #seleciona sempre o dia 1
    # navegador.find_element(By.XPATH,'/html/body/app-root/app-layout/ng-sidebar-container/div/div/app-chatbox/app-content/main/div/app-portlet/div/div[2]/div/div[1]/div/div/div/div[6]/div/ngb-datepicker-input/ngb-datepicker/div[2]/div[1]/ngb-datepicker-month/div[2]/div[5]/span').click()
    # time.sleep(2)
    # #seleciona o dia anterior para data final
    # navegador.find_element(By.XPATH,f"//div[contains(@class, 'ngb-dp-day') and not(contains(@class, 'disabled')) and contains(@aria-label, '{data_formatada}')]").click
    # time.sleep(2)
    # #volta pra tela 
    # navegador.find_element(By.XPATH,'/html/body/app-root/app-layout/ng-sidebar-container/div/div/app-chatbox/app-content/main/div/app-portlet/div/div[2]/div/div[1]/div/div/div').click()
    # time.sleep(2)
    #clica para buscar
    navegador.find_element(By.XPATH,'/html/body/app-root/app-layout/ng-sidebar-container/div/div/app-chatbox/app-content/main/div/app-portlet/div/div[2]/div/div[1]/div/div/div/div[7]/div/div/button[1]').click()
    time.sleep(2)
    #exporta o relatório de chat 
    navegador.find_element(By.XPATH,'/html/body/app-root/app-layout/ng-sidebar-container/div/div/app-chatbox/app-content/main/div/app-portlet/div/div[2]/div/div[1]/div/div/div/div[7]/div/div/button[3]').click()
    time.sleep(5)
    print(f"Planilha baixada do dia {data_atual}")
    time.sleep(3)
    shutil.move('C:/Users/arthuraiala/Downloads/Chats Box - Export.xlsx', os.path.join(caminho_destino, 'Chats Box - Export.xlsx'))
    navegador.quit()


try:
    acessar()
    dados()
    envio()
except:
    emailInsusesso()