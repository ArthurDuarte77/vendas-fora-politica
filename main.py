import schedule
import tkinter as tk
from tkinter import ttk
from tkcalendar import DateEntry
from tkinter import messagebox
import subprocess
import argparse
from datetime import datetime, timedelta
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
from selenium.common.exceptions import *
import time
import json

def chamar_script(dia_inicial, dia_final, cookie):
    arquivos = [
        "modelos_jfa.xlsx",
        "modelos_stetson.xlsx",
        "modelos_taramps.xlsx",
        "modelos_usina.xlsx",
        "produtos.xlsx"
    ]
    
    for arquivo in arquivos:
        if os.path.exists(arquivo):
            os.remove(arquivo)
    
    scripts = ['jfa.py']
    
    # dia_inicial = dia_inicial.get_date().strftime('%Y-%m-%d')
    # dia_final = dia_final.get_date().strftime('%Y-%m-%d')
    
    for script in scripts:
        comando = [
            'python',
            script,
            '--dia_inicial', dia_inicial,
            '--dia_final', dia_final,
            '--cookie', cookie
        ]
        subprocess.run(comando)

parser = argparse.ArgumentParser(description='Executar scripts com datas específicas.')
parser.add_argument('--dia_inicial', type=str, help='Data inicial no formato YYYY-MM-DD')
parser.add_argument('--dia_final', type=str, help='Data final no formato YYYY-MM-DD')
parser.add_argument('--cookie', type=str, help='Cookie')

args = parser.parse_args()

if not args.dia_inicial or not args.dia_final:
    data_anterior = datetime.now() - timedelta(days=1)
    dia_padrao = data_anterior.strftime('%Y-%m-%d')
    dia_inicial = args.dia_inicial or dia_padrao
    dia_final = args.dia_final or dia_padrao
    cookie = args.cookie

service = Service()
options = webdriver.ChromeOptions()
options.add_argument("--disable-gpu")
options.add_argument("--disable-extensions")
prefs = {"profile.managed_default_content_settings.images": 2}
options.add_experimental_option("prefs", prefs)

driver = webdriver.Chrome(service=service, options=options)
driver.get("https://www.google.com.br/?hl=pt-BR")
time.sleep(3)
try:
    driver.get("https://corp.shoppingdeprecos.com.br/login")
    counter = 0
    while True:
        test = driver.find_elements(By.XPATH, '//*[@id="email"]')
        if test:
            break
        else:
            counter += 1
            if counter > 20:
                break
            time.sleep(0.5)
    driver.find_element(By.XPATH, '//*[@id="email"]').send_keys("loja@jfaeletronicos.com")
    driver.find_element(By.XPATH, '//*[@id="password"]').send_keys("922982PC")
    driver.find_element(By.XPATH, '//*[@id="btnLogin"]').click()
except TimeoutException as e:
    print(f"Timeout ao tentar carregar a página ou encontrar um elemento: {e}")
except NoSuchElementException as e:
    print(f"Elemento não encontrado na página: {e}")
except WebDriverException as e:
    print(f"Erro no WebDriver: {e}")

time.sleep(3)
driver.get("https://corp.shoppingdeprecos.com.br/vendedores/vendasMarca")

cookies_list = []

cookies = driver.get_cookies()
for cookie in cookies:
    objeto = cookie['name']
    value = cookie['value']
    cookies_list.append(f"{objeto}={value};")

cookie = "".join(cookies_list)

data_atual = datetime.now()

data_inicial = data_atual - timedelta(days=1)


# Obter a data de ontem
ontem = datetime.now() - timedelta(days=1)
data_ontem = ontem.strftime("%Y-%m-%d")



schedule.every().day.at("07:50").do(chamar_script(data_ontem, data_ontem, cookie=cookie))
schedule.every().day.at("10:50").do(chamar_script(data_ontem, data_ontem, cookie=cookie))
schedule.every().day.at("13:50").do(chamar_script(data_ontem, data_ontem, cookie=cookie))
schedule.every().day.at("15:50").do(chamar_script(data_ontem, data_ontem, cookie=cookie))
schedule.every().day.at("19:50").do(chamar_script(data_ontem, data_ontem, cookie=cookie))
schedule.every().day.at("23:50").do(chamar_script(data_ontem, data_ontem, cookie=cookie))

while True:
    schedule.run_pending()
    time.sleep(60) 