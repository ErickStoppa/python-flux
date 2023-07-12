import os
import shutil
import pyautogui
import time
import pandas as pd
import win32gui
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

import win32con
import datetime




def baixa_arquivo(nav):
    try:
        nav.get("http://prwvrs01ma/Reports_RS1P/Pages/Folder.aspx?ItemPath=%2fIntelig%c3%aancia+Comercial")
        nav.find_element(By.XPATH, '//*[@id="ui_a1"]/tbody/tr/td[2]/a').click()
        nav.find_element(By.XPATH, '//*[@id="ui_a6"]/tbody/tr/td[2]/a').click()
        time.sleep(14)
        pyautogui.click(x = 688,y = 224)
        nav.find_element(By.ID, 'ctl31_ctl05_ctl04_ctl00_ButtonLink').click()
        nav.find_element(By.XPATH, '//*[@id="ctl31_ctl05_ctl04_ctl00_Menu"]/div[5]/a').click()
        time.sleep(20)
        #vencimentos
        nav.get("http://prwvrs01ma/Reports_RS1P/Pages/Report.aspx?ItemPath=%2fIntelig%c3%aancia+Comercial%2fRelat%c3%b3rios+Di%c3%a1rios%2fVencimentos+-+Parcelas")
        nav.find_element(By.XPATH, '//*[@id="ctl31_ctl04_ctl03_txtValue"]').send_keys('01/01/2021')
        nav.find_element(By.XPATH, '//*[@id="ctl31_ctl04_ctl05_txtValue"]').send_keys('31/12/2022')
        time.sleep(2)
        nav.find_element(By.XPATH, '//*[@id="ctl31_ctl04_ctl00"]').click()
        time.sleep(30)
        pyautogui.click(x = 707,y = 329)
        nav.find_element(By.XPATH, '//*[@id="ctl31_ctl05_ctl04_ctl00_Menu"]/div[2]/a').click()
        time.sleep(7)
    except Exception as e:
        print(e)
        nav.close()

def modifica_arquivo(hwnd):
    os.startfile(r"C:\Users\esmartins\Downloads\Limite Risco - % Utilização.xls")
    time.sleep(5)
    win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
    time.sleep(5)

    for i in range(6):
        pyautogui.press("tab")

    listax = [1221, 1192, 1353, 1192, 1580, 1192, 364]
    listay = [401, 180, 401, 180, 399, 180, 29]

    for i in range(7):
        pyautogui.click(x=listax[i-1], y=listay[i-1])

    pyautogui.hotkey('ctrl', 'b')
    time.sleep(5)
    pyautogui.hotkey('alt', 'F4')

def transfere():

    #mover os arquivos
    source = "C:/Users/esmartins/Downloads/Limite Risco - % Utilização.xls"
    source2 = "C:/Users/esmartins/Downloads/Vencimentos - Parcelas.csv"
    destination = "C:/Users/esmartins/OneDrive - Banco Sofisa (1)/Inteligência Comercial - Power BI/01. Relatórios/00. Bases Salesforce"
    shutil.move(source, destination)
    shutil.move(source2, destination)

    #abrir arquivos sales
def sales(hwnd):
    base_risco_atualizada = r"C:\Users\esmartins\OneDrive - Banco Sofisa (1)\Inteligência Comercial - Power BI\01. Relatórios\00. Bases Salesforce\Base_Risco_Atualizada.xlsx"
    bases_vencimentos_atualizada = r"C:\Users\esmartins\OneDrive - Banco Sofisa (1)\Inteligência Comercial - Power BI\01. Relatórios\00. Bases Salesforce\Bases_Vencimentos_Atualizada.xlsx"
    base_risco_regularizacao = r"C:\Users\esmartins\OneDrive - Banco Sofisa (1)\Inteligência Comercial - Power BI\01. Relatórios\00. Bases Salesforce\Base_Risco_Regularizacao.xlsx"

    listatempo = [15, 20, 60]
    lista_arquivos = [base_risco_regularizacao,bases_vencimentos_atualizada,base_risco_atualizada]
    for i in range(len(lista_arquivos)):
        os.startfile(lista_arquivos[i])
        time.sleep(6)
        win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
        time.sleep(5)
        pyautogui.hotkey('ctrl', 'alt', 'F5')
        time.sleep(listatempo[i])
        pyautogui.hotkey('ctrl', 'b')
        time.sleep(5)
        pyautogui.hotkey('alt', 'F4')
        time.sleep(5)


def hist_sales(hwnd):
    os.startfile(
        r"C:\Users\esmartins\OneDrive - Banco Sofisa (1)\Inteligência Comercial - Power BI\01. Relatórios\00. Bases Salesforce\Base_Risco_Historico.xlsx")
    time.sleep(8)
    pyautogui.click(x=909, y=694)
    pyautogui.hotkey('ctrl', 'left')
    pyautogui.hotkey('ctrl', 'up')
    pyautogui.hotkey('down')
    pyautogui.keyDown('ctrl')
    pyautogui.keyDown('shiftright')
    pyautogui.keyDown('shiftleft')
    pyautogui.press('right')
    pyautogui.press('down')
    pyautogui.keyUp('ctrl')
    pyautogui.keyUp('shiftright')
    pyautogui.keyUp('shiftleft')
    pyautogui.press('delete')
    time.sleep(3)

    # abrir o regularizacao
    os.startfile(r"C:\Users\esmartins\OneDrive - Banco Sofisa (1)\Inteligência Comercial - Power BI\01. Relatórios\00. Bases Salesforce\Base_Risco_Regularizacao.xlsx")
    time.sleep(4)
    win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
    pyautogui.click(x=909, y=694)
    pyautogui.hotkey('ctrl', 'left')
    pyautogui.hotkey('ctrl', 'up')
    pyautogui.press('down')
    pyautogui.keyDown('ctrl')
    pyautogui.keyDown('shiftright')
    pyautogui.keyDown('shiftleft')
    pyautogui.press('right')
    pyautogui.press('down')
    pyautogui.keyUp('ctrl')
    pyautogui.keyUp('shiftright')
    pyautogui.keyUp('shiftleft')
    pyautogui.hotkey('ctrl', 'c')
    pyautogui.hotkey('alt', 'TAB')
    time.sleep(3)
    pyautogui.press('enter')
    # colar no hist
    pyautogui.click(x=909, y=694)
    pyautogui.hotkey('ctrl', 'left')
    pyautogui.hotkey('ctrl', 'up')
    pyautogui.press('down')
    pyautogui.click(button='right', x=909, y=694)
    pyautogui.press('v')
    time.sleep(1)
    pyautogui.hotkey('ctrl', 'b')
    time.sleep(5)
    pyautogui.hotkey('alt', 'F4')


    #pandas\
def bambu():
    df_hist = pd.read_excel(r"C:\Users\esmartins\OneDrive - Banco Sofisa (1)\Inteligência Comercial - Power BI\01. Relatórios\00. Bases Salesforce\Base_Risco_Historico.xlsx")
    hoje = datetime.datetime.today()
    pd.to_datetime(df_hist['DATA_BASE'])
    df_hist = df_hist[df_hist['DATA_BASE'] != hoje]

    df_reg = pd.read_excel(r"C:\Users\esmartins\OneDrive - Banco Sofisa (1)\Inteligência Comercial - Power BI\01. Relatórios\00. Bases Salesforce\Base_Risco_Regularizacao.xlsx")
    df_reg = df_reg[df_reg['DATA_BASE'] == hoje]

    df_concat = pd.concat([df_hist,df_reg], ignore_index=True)
    df_concat = df_concat[["CNPJ 14", "Cliente", "Limite", "Saldo Contábil", "Total Geral", "Superint Exec", "Superintendente", "Gerente", "Segmento","DATA_BASE", "FIDC", "FINAME"]]
    df_concat.to_excel(r"C:\Users\esmartins\OneDrive - Banco Sofisa (1)\Inteligência Comercial - Power BI\01. Relatórios\00. Bases Salesforce\Base_Risco_Historico.xlsx", index=False)

def pbi():
    os.startfile(r"C:\Program Files\Google\Chrome\Application\chrome.exe")
    win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
    time.sleep(3)
    pyautogui.press("F11")
    time.sleep(2)
    pyautogui.click(x = 309,y = 20)
    time.sleep(60)
    pyautogui.moveTo(x=515,y=937)
    time.sleep(5)
    pyautogui.click(x=515,y=937)
    time.sleep(100)
    pyautogui.moveTo(x = 515,y = 545)
    time.sleep(5)
    pyautogui.click(x=515, y=545)


if __name__ == '__main__':

    hwnd = win32gui.GetForegroundWindow()
    win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)

    pyautogui.PAUSE = 1

    options = webdriver.ChromeOptions()
    options.add_argument("--log-level=3")
    s = Service(r"C:\chrome_driver\chromedriver.exe")
    driver = webdriver.Chrome(service=s, options=options)
    driver.maximize_window()

    baixa_arquivo(driver)
    modifica_arquivo(hwnd)
    transfere()
    sales(hwnd)
    hist_sales(hwnd)
    bambu()
    pbi()
#pegar posição
import pyautogui
import time
import pandas as pd
import zipfile
import os
import shutil

time.sleep(3)
x, y = pyautogui.position()
print ("Posicao atual do mouse:")
print ("x = "+str(x)+" y = "+str(y))
print('finalizou')