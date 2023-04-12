import time
from tracemalloc import stop
import pandas as pd
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium import webdriver
import pyautogui
import pyperclip
from soupsieve import select
from sympy import prime
from webdriver_manager.chrome import ChromeDriverManager
import warnings
import pandas as pd
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
import subprocess
from pathlib import Path


try:
    import pyautogui
except ImportError:
    subprocess.call("pip install pyautogui")

caminho = input("Digite o caminho do excel: ")

df = pd.read_excel(caminho)

alias = Path.home().parts[2]

occur = df[df.columns[0]].count()
ciclos = occur//20
resto = occur%20
if occur < 20:
    resto = occur


erros = []
warnings.filterwarnings("ignore")

options = webdriver.ChromeOptions()
options.add_argument('--log-level=3')
options.add_argument('--height=900')  
options.add_argument('--width=1600')
options.add_experimental_option('excludeSwitches', ['enable-logging'])
#options.add_argument(rf'user-data-dir=C:\Users\mcampinh\AppData\Local\Google\Chrome\New')
driver = webdriver.Chrome(executable_path=ChromeDriverManager().install(), options = options)

driver.get('https://sclens.corp.amazon.com/')

df_merchant = df['Merchant_ID']
cnpj_cpf = []
merchants = []
phone = []
email = []


rows = df.shape[0]
actions = ActionChains(driver)      

def phone_email():
    driver.get("https://www.sellercentral.amazon.dev/gp/account-manager/home.html/ref=xx_userperms_dnav_xx")

    try:
        WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH,'/html/body/div[1]/div[2]/div/div[1]/table/tbody[1]/tr[1]/td/h1'))).click()
        
        email_a = driver.find_element("xpath","/html/body/div[1]/div[2]/div/div[1]/table/tbody[1]/tr[5]/td[3]/span").text
        email.append(email_a)
    except:
        email.append("nulo")


    driver.get("https://www.sellercentral.amazon.dev/easyship/panjeekaran/accountSettings")
    try:
        WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH,'/html/body/div[1]/div[2]/div/div[3]/span[2]'))).click()
        phone_number = driver.find_element("xpath","/html/body/div[1]/div[2]/div/div[4]/div/div/div[7]/div/div[1]/p[2]/span[7]").text
        phone.append(phone_number)
    except:
        phone.append("nulo")


def seller_central(a):
    merchants.append(df_merchant[a])
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH,'/html/body/div[1]/div[2]/div/form/div/div/div[2]/div/label'))).click()
    try:
        WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.XPATH,'/html/body/div[1]/div[2]/div/form/div/div/div[2]/div/div[1]/div[1]/div[1]/label'))).click()
        cnpj_cpf.append("CPF")
        phone_email()

    except:
        cnpj_cpf.append("CNPJ")
        phone_email()

    time.sleep(0.5)
    #driver.switch_to.window(original_window)


if ciclos > 0:
    print(ciclos)
    for i in range(0, ciclos):
        WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH,'/html/body/div/div/div/kat-tabs/a[2]/kat-tab/span'))).click() #user-seller-mapping
        WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH,'/html/body/div/div/div/div/div/div/div[2]/div[1]/div/span/kat-button'))).click() #request
        WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH,'/html/body/div/div/div/div/div/div/div[2]/div[2]/div/span'))).click() #add-new-seller
        for j in range(1, 21):
            linha = 20*i + j - 1
            WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH,'/html/body/div/div/div/div/div/div/div[2]/div[2]/div/span'))).click() #add-new-seller
            row = j - 1

            WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH,'/html/body/div/div/div/div/div/div/div[2]/div[1]/kat-table/kat-table-body/kat-table-row[{}]/kat-table-cell[1]'.format(j)))).click()
            merchant = df_merchant[linha]
            merchant = str(merchant)
            pyperclip.copy(merchant)
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(0.5)

            css_tag = '.scl-self-request-marketplace-column'
            els = driver.find_elements('css selector', css_tag)
            el = els[row]
            actions.click(on_element=el)
            actions.send_keys(Keys.DOWN)
            actions.send_keys('b')
            actions.send_keys(Keys.RETURN)
            actions.perform()
            
            time.sleep(.1)

            if j == 19:
                WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH,'/html/body/div/div/div/div/div/div/div[2]/div[1]/kat-table/kat-table-body/kat-table-row[20]/kat-table-cell[8]/span'))).click() #remove a ultima linha vazia
        WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH,'/html/body/div/div/div/div/div/div/div[2]/div[3]/div/span/kat-button[1]'))).click() #clica em submit
        
    
if resto > 0:
    WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH,'/html/body/div/div/div/kat-tabs/a[2]/kat-tab/span'))).click() #user-seller-mapping
    WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH,'/html/body/div/div/div/div/div/div/div[2]/div[1]/div/span/kat-button'))).click() #request
    WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH,'/html/body/div/div/div/div/div/div/div[2]/div[2]/div/span'))).click() #add-new-seller
    total = ciclos*20
    try:
        for r in range(1,resto+1):
            linha = r - 1 + total
            row = r - 1
            WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH,'/html/body/div/div/div/div/div/div/div[2]/div[2]/div/span'))).click() #add-new-seller     
            WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH,'/html/body/div/div/div/div/div/div/div[2]/div[1]/kat-table/kat-table-body/kat-table-row[{}]/kat-table-cell[1]'.format(r)))).click()
            merchant = df_merchant[linha]
            merchant = str(merchant)
            pyperclip.copy(merchant)
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(0.5)

            css_tag = '.scl-self-request-marketplace-column'
            els = driver.find_elements('css selector', css_tag)
            el = els[row]
            actions.click(on_element=el)
            actions.send_keys(Keys.DOWN)
            actions.send_keys('b')
            actions.send_keys(Keys.RETURN)
            actions.perform()
            time.sleep(.1)

            if  r == (resto-1):
                WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH,'/html/body/div/div/div/div/div/div/div[2]/div[1]/kat-table/kat-table-body/kat-table-row[{}]/kat-table-cell[8]/span'.format(resto)))).click() #remove a ultima linha vazia
        WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH,'/html/body/div/div/div/div/div/div/div[2]/div[3]/div/span/kat-button[1]'))).click() #clica em submit
    except:
        print("Erro.")

else:
    print("Ok.")

WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH,'/html/body/div/div/div/kat-tabs/a[3]/kat-tab'))).click() #aba lens-merchant-picker
count = 1
original_window = driver.current_window_handle

for a in range(0,rows):
    try:
        WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH,'/html/body/div/div/div/div/div/div/div[2]/div[1]'))).click()
        merchant = df_merchant[a]
        merchant = str(merchant)
        time.sleep(0.5)
        #elm = driver.find_elements("css selector",'input[value="katal-id-4"]')
        #driver.execute_script("arguments[0].value = 'arguments[1]';", elm, merchant)

        pyperclip.copy(merchant)
        pyautogui.hotkey('ctrl', 'a')       
        pyautogui.hotkey('delete') 
        pyautogui.hotkey('ctrl', 'v') 
        WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH,'/html/body/div/div/div/div/div/div/div[2]/div[2]/kat-button'))).click()
        time.sleep(.2)

        tag = '.scl-grant-page-merchant-search-button kat-button'
        driver.find_element('css selector', tag).click()
        time.sleep(1)
        driver.execute_script('''document.querySelector('input[value="A2Q3Y263D00KWC"]').click()''')
        time.sleep(.5)
        driver.execute_script('''document.querySelector('kat-button[label="Access"]').click()''')
        time.sleep(3)
        driver.switch_to.window(driver.window_handles[1])
        WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH,'/html/body/div[1]/div[1]/div[3]/div[1]/div[1]/div[2]/div/form/input'))).click()
        driver.get('https://www.sellercentral.amazon.dev/tax/br/taxinformation?ref_=macs_xxcpfinf_cont_acinfohm')
        time.sleep(0.8)
    except:
        erros.append(df_merchant[a])
        merchants.append(df_merchant[a])
        cnpj_cpf.append("erro")
        break

    try:
        seller_central(a)                                                      
        count += 1
    except:
        driver.switch_to.window(original_window)
        erros.append(df_merchant[a])
        merchants.append(df_merchant[a])
        cnpj_cpf.append("nula")

    handles = driver.window_handles 
    for i in handles: 
        driver.switch_to.window(i)     
        if driver.title == "Amazon": 
            driver.close()
    
    driver.switch_to.window(original_window)


erros_final = list()
for c in erros:
    if c not in erros_final:
        erros_final.append(c)

print(erros_final)
print(merchants)
print(cnpj_cpf)
print(phone)
print(email)
df = pd.DataFrame((zip(merchants,cnpj_cpf,phone,email)), columns = ['Merchant_ID', 'CNPJ/CPF','Phone','E-mail'])
df.to_excel(caminho)
print("\n")
input("Exit.")