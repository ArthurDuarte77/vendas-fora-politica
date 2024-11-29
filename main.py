import sys
import customtkinter as ctk
from tqdm import tqdm
from unidecode import unidecode
import pandas as pd
import numpy as np
from sklearn.model_selection import train_test_split, cross_val_score, StratifiedKFold
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.preprocessing import StandardScaler, LabelEncoder
from sklearn.ensemble import RandomForestClassifier
from sklearn.metrics import classification_report, accuracy_score
from sklearn.pipeline import Pipeline
from sklearn.compose import ColumnTransformer
from sklearn.utils.class_weight import compute_class_weight
import joblib  # Para salvar e carregar o modelo
import requests
import os
import time
import json
import pandas as pd
import tkinter as tk
from tkinter import messagebox
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, WebDriverException, JavascriptException
import tkinter as tk
from tkinter import ttk, messagebox
from tkcalendar import DateEntry
import datetime

dataInicial = ""
dataFinal = ""

start_row = 20  
end_row = 37
num_rows = end_row - start_row

df = pd.read_excel("GESTÃO DE AÇÕES E-COMMERCE.xlsx", usecols='C:O', skiprows=start_row, nrows=num_rows, engine='openpyxl', sheet_name="POLÍTICA COMERCIAL Nov24")

df.columns = ['PRODUTO', 'inutil1', 'SITE', 'COLUNA3','inutil2', 'CLÁSSICO ML', 'COLUNA5','inutil3', 'PREMIUM ML', 'COLUNA7','inutil4', 'MARKETPLACES', 'COLUNA9']

for index, i in df.iterrows():
    if i['PRODUTO'] == "FONTE 40A":
        fonte40Marketplace = round(i['COLUNA3']- 0.03 , 2) ;
        fonte40Classico = round(i['COLUNA5']- 0.03 , 2) ;
        fonte40Premium = round(i['COLUNA7']- 0.03 , 2) ;
        fonte40PremiumPrice = round(i['PREMIUM ML']- 0.03 , 2) ;
        fonte40ClassicoPrice = round(i['CLÁSSICO ML']- 0.03 , 2) ;
        fonte40Marketplaceprice = round(i['SITE']- 0.03 , 2) ;
    elif i['PRODUTO'] == "FONTE 60A":
        fonte60Marketplace = round(i['COLUNA3']- 0.03 , 2) ;
        fonte60Classico = round(i['COLUNA5']- 0.03 , 2) ;
        fonte60Premium = round(i['COLUNA7']- 0.03 , 2) ;
        fonte60PremiumPrice = round(i['PREMIUM ML']- 0.03 , 2) ;
        fonte60ClassicoPrice = round(i['CLÁSSICO ML']- 0.03 , 2) ;
        fonte60Marketplaceprice = round(i['SITE']- 0.03 , 2) ;
    elif i['PRODUTO'] == "FONTE 60A LITE":
        fonte60liteMarketplace = round(i['COLUNA3']- 0.03 , 2) ;
        fonte60liteClassico = round(i['COLUNA5']- 0.03 , 2) ;
        fonte60litePremium = round(i['COLUNA7']- 0.03 , 2) ;
        fonte60litePremiumPrice = round(i['PREMIUM ML']- 0.03 , 2) ;
        fonte60liteClassicoPrice = round(i['CLÁSSICO ML']- 0.03 , 2) ;
        fonte60liteMarketplaceprice = round(i['SITE']- 0.03 , 2) ;
    elif i['PRODUTO'] == "FONTE 70A":
        fonte70Marketplace = round(i['COLUNA3']- 0.03 , 2) ;
        fonte70Classico = round(i['COLUNA5']- 0.03 , 2) ;
        fonte70Premium = round(i['COLUNA7']- 0.03 , 2) ;
        fonte70PremiumPrice = round(i['PREMIUM ML']- 0.03 , 2) ;
        fonte70ClassicoPrice = round(i['CLÁSSICO ML']- 0.03 , 2) ;
        fonte70Marketplaceprice = round(i['SITE']- 0.03 , 2) ;
    elif i['PRODUTO'] == "FONTE 70A LITE":
        fonte70liteMarketplace = round(i['COLUNA3']- 0.03 , 2) ;
        fonte70liteClassico = round(i['COLUNA5']- 0.03 , 2) ;
        fonte70litePremium = round(i['COLUNA7']- 0.03 , 2) ;
        fonte70litePremiumPrice = round(i['PREMIUM ML']- 0.03 , 2) ;
        fonte70liteClassicoPrice = round(i['CLÁSSICO ML']- 0.03 , 2) ;
        fonte70liteMarketplaceprice = round(i['SITE']- 0.03 , 2) ;
    elif i['PRODUTO'] == "FONTE 90 BOB":
        fonte90bobMarketplace = round(i['COLUNA3']- 0.03 , 2) ;
        fonte90bobClassico = round(i['COLUNA5']- 0.03 , 2) ;
        fonte90bobPremium = round(i['COLUNA7']- 0.03 , 2) ;
        fonte90bobPremiumPrice = round(i['PREMIUM ML']- 0.03 , 2) ;
        fonte90bobClassicoPrice = round(i['CLÁSSICO ML']- 0.03 , 2) ;
        fonte90bobMarketplaceprice = round(i['SITE']- 0.03 , 2) ;
    elif i['PRODUTO'] == "FONTE 120 BOB":
        fonte120bobMarketplace = round(i['COLUNA3']- 0.03 , 2) ;
        fonte120bobClassico = round(i['COLUNA5']- 0.03 , 2) ;
        fonte120bobPremium = round(i['COLUNA7']- 0.03 , 2) ;
        fonte120bobPremiumPrice = round(i['PREMIUM ML']- 0.03 , 2) ;
        fonte120bobClassicoPrice = round(i['CLÁSSICO ML']- 0.03 , 2) ;
        fonte120bobMarketplaceprice = round(i['SITE']- 0.03 , 2) ;
    elif i['PRODUTO'] == "FONTE 120A LITE":
        fonte120liteMarketplace = round(i['COLUNA3']- 0.03 , 2) ;
        fonte120liteClassico = round(i['COLUNA5']- 0.03 , 2) ;
        fonte120litePremium = round(i['COLUNA7']- 0.03 , 2) ;
        fonte120litePremiumPrice = round(i['PREMIUM ML']- 0.03 , 2) ;
        fonte120liteClassicoPrice = round(i['CLÁSSICO ML']- 0.03 , 2) ;
        fonte120liteMarketplaceprice = round(i['SITE']- 0.03 , 2) ;
    elif i['PRODUTO'] == "FONTE 120A":
        fonte120Marketplace = round(i['COLUNA3']- 0.03 , 2) ;
        fonte120Classico = round(i['COLUNA5']- 0.03 , 2) ;
        fonte120Premium = round(i['COLUNA7']- 0.03 , 2) ;
        fonte120PremiumPrice = round(i['PREMIUM ML']- 0.03 , 2) ;
        fonte120ClassicoPrice = round(i['CLÁSSICO ML']- 0.03 , 2) ;
        fonte120Marketplaceprice = round(i['SITE']- 0.03 , 2) ;
    elif i['PRODUTO'] == "FONTE 200 BOB":
        fonte200bobMarketplace = round(i['COLUNA3']- 0.03 , 2) ;
        fonte200bobClassico = round(i['COLUNA5']- 0.03 , 2) ;
        fonte200bobPremium = round(i['COLUNA7']- 0.03 , 2) ;
        fonte200bobPremiumPrice = round(i['PREMIUM ML']- 0.03 , 2) ;
        fonte200bobClassicoPrice = round(i['CLÁSSICO ML']- 0.03 , 2) ;
        fonte200bobMarketplaceprice = round(i['SITE']- 0.03 , 2) ;
    elif i['PRODUTO'] == "FONTE 200A LITE":
        fonte200liteMarketplace = round(i['COLUNA3']- 0.03 , 2) ;
        fonte200liteClassico = round(i['COLUNA5']- 0.03 , 2) ;
        fonte200litePremium = round(i['COLUNA7']- 0.03 , 2) ;
        fonte200litePremiumPrice = round(i['PREMIUM ML']- 0.03 , 2) ;
        fonte200liteClassicoPrice = round(i['CLÁSSICO ML']- 0.03 , 2) ;
        fonte200liteMarketplaceprice = round(i['SITE']- 0.03 , 2) ;
    elif i['PRODUTO'] == "FONTE 200 MONO":
        fonte200monoMarketplace = round(i['COLUNA3']- 0.03 , 2) ;
        fonte200monoClassico = round(i['COLUNA5']- 0.03 , 2) ;
        fonte200monoPremium = round(i['COLUNA7']- 0.03 , 2) ;
        fonte200monoPremiumPrice = round(i['PREMIUM ML']- 0.03 , 2) ;
        fonte200monoClassicoPrice = round(i['CLÁSSICO ML']- 0.03 , 2) ;
        fonte200monoMarketplaceprice = round(i['SITE']- 0.03 , 2) ;
    elif i['PRODUTO'] == "FONTE 200A":
        fonte200Marketplace = round(i['COLUNA3']- 0.03 , 2) ;
        fonte200Classico = round(i['COLUNA5']- 0.03 , 2) ;
        fonte200Premium = round(i['COLUNA7']- 0.03 , 2) ;
        fonte200PremiumPrice = round(i['PREMIUM ML']- 0.03 , 2) ;
        fonte200ClassicoPrice = round(i['CLÁSSICO ML']- 0.03 , 2) ;
        fonte200Marketplaceprice = round(i['SITE']- 0.03 , 2) ;
    elif i['PRODUTO'] == "K1200":
        controleK1200Marketplace = round(i['COLUNA3']- 0.03 , 2) ;
        controleK1200Classico = round(i['COLUNA5']- 0.03 , 2) ;
        controleK1200Premium = round(i['COLUNA7']- 0.03 , 2) ;
        controleK1200PremiumPrice = round(i['PREMIUM ML']- 0.03 , 2) ;
        controleK1200ClassicoPrice = round(i['CLÁSSICO ML']- 0.03 , 2) ;
        controleK1200Marketplaceprice = round(i['SITE']- 0.03 , 2) ;
    elif i['PRODUTO'] == "K600":
        controleK600Marketplace = round(i['COLUNA3']- 0.03 , 2) ;
        controleK600Classico = round(i['COLUNA5']- 0.03 , 2) ;
        controleK600Premium = round(i['COLUNA7']- 0.03 , 2) ;
        controleK600PremiumPrice = round(i['PREMIUM ML']- 0.03 , 2) ;
        controleK600ClassicoPrice = round(i['CLÁSSICO ML']- 0.03 , 2) ;
        controleK600Marketplaceprice = round(i['SITE']- 0.03 , 2) ;
    elif i['PRODUTO'] == "CONTROLE WR":
        controleRedlineMarketplace = round(i['COLUNA3']- 0.03 , 2) ;
        controleRedlineClassico = round(i['COLUNA5']- 0.03 , 2) ;
        controleRedlinePremium = round(i['COLUNA7']- 0.03 , 2) ;
        controleRedlinePremiumPrice = round(i['PREMIUM ML']- 0.03 , 2) ;
        controleRedlineClassicoPrice = round(i['CLÁSSICO ML']- 0.03 , 2) ;
        controleRedlineMarketplaceprice = round(i['SITE']- 0.03 , 2) ;
    elif i['PRODUTO'] == "ACQUA":
        controleAcquaMarketplace = round(i['COLUNA3']- 0.03 , 2) ;
        controleAcquaClassico = round(i['COLUNA5']- 0.03 , 2) ;
        controleAcquaPremium = round(i['COLUNA7']- 0.03 , 2) ;
        controleAcquaPremiumPrice = round(i['PREMIUM ML']- 0.03 , 2) ;
        controleAcquaClassicoPrice = round(i['CLÁSSICO ML']- 0.03 , 2) ;
        controleAcquaMarketplaceprice = round(i['SITE']- 0.03 , 2) ;


def login_to_website(driver, email, password):
    try:
        driver.get("https://corp.shoppingdeprecos.com.br/login")
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "email")))

        driver.find_element(By.ID, "email").send_keys(email)
        driver.find_element(By.ID, "password").send_keys(password)
        driver.find_element(By.ID, "btnLogin").click()
        time.sleep(5)
        return
    except (TimeoutException, NoSuchElementException) as e:
        print(f"Error during login: {e}")

def extract_items(driver):
    items = []
    try:
        # Define a marca e as datas desejadas
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "cmbMarca"))).click()
        driver.find_element(By.XPATH, '//*[@id="cmbMarca"]').click()
        time.sleep(1)
        driver.find_element(By.XPATH, '//*[@id="cmbMarca"]/option[25]').click()
        time.sleep(1)
        driver.find_element(By.ID, "txtIni").send_keys(dataInicial)
        time.sleep(1)
        driver.find_element(By.ID, "txtFim").send_keys(dataFinal)
        time.sleep(1)
        driver.find_element(By.ID, "btnBuscar").click()

        time.sleep(5) 
        
        ids = [i for i in driver.find_elements(By.XPATH, '/html/body/div[2]/div[2]/div[2]/div/div/div[2]/div/div/div[1]/div/table/tbody/tr') if i.get_attribute("id")]

        for element in ids:
            element_id = element.get_attribute("id")
            try:
                WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, element_id)))
                driver.execute_script(f"tabelaItens('{element_id}', 0)")
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="tr_concorrente_"]')))
                count = 0;
                while True:
                    time.sleep(2)
                    for item in driver.find_elements(By.XPATH, '//*[@id="tr_concorrente_"]'):
                        imagem = item.find_element(By.XPATH, './td[1]/img').get_attribute("src")
                        nome = item.find_element(By.XPATH, './td[2]').text.split("\n")[0]
                        tipo = item.find_element(By.XPATH, './td[2]').text.split("\n")[1].lower().replace("anúncio", "").strip()
                        vendedor = item.find_element(By.XPATH, './td[3]').text
                        quantidade = item.find_element(By.XPATH, './td[5]').text
                        valor_unitario = item.find_element(By.XPATH, './td[6]').text
                        total = item.find_element(By.XPATH, './td[7]').text
                        items.append({
                            "data": datetime.datetime.strptime(dataInicial, "%d%m%Y").strftime("%d-%m-%Y"),
                            "imagem": imagem,
                            "Produto": nome,
                            "Tipo de Anúncio": tipo,
                            "vendedor": vendedor,
                            "quantidade": quantidade,
                            "Preço Unitário": valor_unitario,
                            "total": total,
                            "Produto2": "OUTROS"
                        })
                    try:
                        if driver.find_element(By.XPATH, '//li[@class="next page"]/a'):
                            count += 30
                            driver.execute_script(f"tabela({count});")
                        else:
                            break
                    except: 
                        break
            except (TimeoutException, JavascriptException) as e:
                print(f"Error processing element with ID {element_id}: {e}")
            
    except Exception as e:
        print(f"An error occurred during data extraction: {e}")

    return items

def SelecionarFonte(item):
    price = item["Preço Unitário"]
    tipo = unidecode(item["Tipo de Anúncio"].strip().lower())
    if item['Produto2'] == "FONTE 40A":
        if tipo == "classico" and price < fonte40Classico:
            return f"FORA,{fonte40Classico + 0.03}"
        elif tipo == "premium" and price < fonte40Premium:
            return f"FORA,{fonte40Premium + 0.03}"

    if item['Produto2'] == "FONTE 60A":
        if tipo == "classico" and price < fonte60Classico:
            return f"FORA,{fonte60Classico + 0.03}"
        elif tipo == "premium" and price < fonte60Premium:
            return f"FORA,{fonte60Premium + 0.03}"

    if item['Produto2'] == "FONTE LITE 60A":
        if tipo == "classico" and price < fonte60liteClassico:
            return f"FORA,{fonte60liteClassico + 0.03}"
        elif tipo == "premium" and price < fonte60litePremium:
            return f"FORA,{fonte60litePremium + 0.03}"

    if item['Produto2'] == "FONTE 70A":
        if tipo == "classico" and price < fonte70Classico:
            return f"FORA,{fonte70Classico + 0.03}"
        elif tipo == "premium" and price < fonte70Premium:
            return f"FORA,{fonte70Premium + 0.03}"

    if item['Produto2'] == "FONTE LITE 70A":
        if tipo == "classico" and price < fonte70liteClassico:
            return f"FORA,{fonte70liteClassico + 0.03}"
        elif tipo == "premium" and price < fonte70litePremium:
            return f"FORA,{fonte70litePremium + 0.03}"

    if item['Produto2'] == "FONTE BOB 90A":
        if tipo == "classico" and price < fonte90bobClassico:
            return f"FORA,{fonte90bobClassico + 0.03}"
        elif tipo == "premium" and price < fonte90bobPremium:
            return f"FORA,{fonte90bobPremium + 0.03}"

    if item['Produto2'] == "FONTE 120A":
        if tipo == "classico" and price < fonte120Classico:
            return f"FORA,{fonte120Classico + 0.03}"
        elif tipo == "premium" and price < fonte120Premium:
            return f"FORA,{fonte120Premium + 0.03}"

    if item['Produto2'] == "FONTE LITE 120A":
        if tipo == "classico" and price < fonte120liteClassico:
            return f"FORA,{fonte120liteClassico + 0.03}"
        elif tipo == "premium" and price < fonte120litePremium:
            return f"FORA,{fonte120litePremium + 0.03}"

    if item['Produto2'] == "FONTE BOB 120A":
        if tipo == "classico" and price < fonte120bobClassico:
            return f"FORA,{fonte120bobClassico + 0.03}"
        elif tipo == "premium" and price < fonte120bobPremium:
            return f"FORA,{fonte120bobPremium + 0.03}"

    if item['Produto2'] == "FONTE 200A":
        if tipo == "classico" and price < fonte200Classico:
            return f"FORA,{fonte200Classico + 0.03}"
        elif tipo == "premium" and price < fonte200Premium:
            return f"FORA,{fonte200Premium + 0.03}"

    if item['Produto2'] == "FONTE MONO 200A":
        if tipo == "classico" and price < fonte200monoClassico:
            return f"FORA,{fonte200monoClassico + 0.03}"
        elif tipo == "premium" and price < fonte200monoPremium:
            return f"FORA,{fonte200monoPremium + 0.03}"

    if item['Produto2'] == "FONTE LITE 200A":
        if tipo == "classico" and price < fonte200liteClassico:
            return f"FORA,{fonte200liteClassico + 0.03}"
        elif tipo == "premium" and price < fonte200litePremium:
            return f"FORA,{fonte200litePremium + 0.03}"

    if item['Produto2'] == "FONTE BOB 200A":
        if tipo == "classico" and price < fonte200bobClassico:
            return f"FORA,{fonte200bobClassico + 0.03}"
        elif tipo == "premium" and price < fonte200bobPremium:
            return f"FORA,{fonte200bobPremium + 0.03}"
        
    if item['Produto2'] == "FONTE 200A MONO":
        if tipo == "classico" and price < fonte200monoClassico:
            return f"FORA,{fonte200bobClassico + 0.03}"
        elif tipo == "premium" and price < fonte200monoPremium:
            return f"FORA,{fonte200bobPremium + 0.03}"
        
    return "DENTRO,0"


def main():
    service = Service()
    options = webdriver.ChromeOptions()
    options.add_argument("--disable-gpu")
    options.add_argument("--disable-extensions")
    prefs = {"profile.managed_default_content_settings.images": 2}
    options.add_experimental_option("prefs", prefs)
    driver = webdriver.Chrome(service=service, options=options)
    driver.get("https://www.google.com.br/?hl=pt-BR")
    time.sleep(3)

    email = "loja@jfaeletronicos.com"  # Substitua com o seu e-mail
    password = "922982PC"  # Substitua com a sua senha
    login_to_website(driver, email, password)

    driver.get("https://corp.shoppingdeprecos.com.br/vendedores/vendasMarca")
    time.sleep(4)
    try:
        # Verifica se o elemento existe e o texto corresponde
        error_element = driver.find_element(By.XPATH, '//*[@id="container"]/h1')
        if error_element.text.strip() == "A Database Error Occurred":
            driver.refresh()
    except NoSuchElementException:
        # Caso o elemento não seja encontrado, continua sem erro
        pass

    # Extrai os itens após o refresh ou caso não haja erro
    items = extract_items(driver)
        

    driver.quit()

    for i in tqdm(items):
        response = requests.get(f"https://api.mercadolibre.com/sites/MLB/search?q={i['Produto']}")
        results = response.json()['results']
        for product in results:
            if product["thumbnail_id"] == i['imagem'].split("D_")[1].split("-I")[0] and i['Produto'] in product["title"]:
                i['Produto'] = product['title']
            
        

    pd.DataFrame(items).to_excel("items.xlsx", index=False)

    
    novos_dados = pd.read_excel("items.xlsx", engine='openpyxl', decimal=",", thousands=".")
    
    for index, item in novos_dados.iterrows():
        price = item['Preço Unitário']
        title = unidecode(item['Produto'].lower())
        item.loc['Produto2'] = "OUTROS"
        if "bob" not in title and "lite" not in title and "light" not in title  and "controle" not in title and "usina" not in title and ("jfa" in title or "fonte carregador" in title or "fonte automotiva" in title or "fonte e carregador" in title or "carregador de baterias" in title):
            if "40a" in title or "40 " in title or "40 amperes" in title or "40amperes" in title or "36a" in title or "36" in title or "36 amperes" in title or "36amperes" in title:
                item['Produto2'] = "FONTE 40A"

        if "bob" not in title and "lite" not in title and "light" not in title  and "controle" not in title and "usina" not in title and ("jfa" in title or "fonte carregador" in title or "fonte automotiva" in title or "fonte e carregador" in title or "carregador de baterias" in title):
            if "60a" in title or "60 " in title or "60 amperes" in title or "60amperes" in title or "60 a" in title:
                item['Produto2'] = "FONTE 60A"

        if "bob" not in title and ("lite" in title or "light" in title) and "controle" not in title and "usina" not in title and ("jfa" in title or "fonte carregador" in title or "fonte automotiva" in title or "fonte e carregador" in title or "carregador de baterias" in title):
            if "60a" in title or "60 " in title or "60 amperes" in title or "60amperes" in title or "60 a" in title: 
                item['Produto2'] = "FONTE LITE 60A"

        
        if "bob" not in title and "lite" not in title and "light" not in title  and "controle" not in title and "usina" not in title and ("jfa" in title or "fonte carregador" in title or "fonte automotiva" in title or "fonte e carregador" in title or "carregador de baterias" in title):
            if "70a" in title or "70 " in title or "70 amperes" in title or "70amperes" in title or "70 a" in title:
                item['Produto2'] = "FONTE 70A"

        if "bob" not in title and  ("lite" in title or "light" in title) and "controle" not in title and "usina" not in title and ("jfa" in title or "fonte carregador" in title or "fonte automotiva" in title or "fonte e carregador" in title or "carregador de baterias" in title):
            if "70a" in title or "70 " in title or "70 amperes" in title or "70amperes" in title or "70 a" in title:
                item['Produto2'] = "FONTE LITE 70A"
                
        if "bob" not in title and "lite" not in title and "light" not in title  and "controle" not in title and "usina" not in title and ("jfa" in title or "fonte carregador" in title or "fonte automotiva" in title or "fonte e carregador" in title or "carregador de baterias" in title):
            if "120a" in title or "120 " in title or "120 amperes" in title or "120amperes" in title or "120 a" in title: 
                item['Produto2'] = "FONTE 120A"

        if "bob" not in title and  ("lite" in title or "light" in title) and "controle" not in title and "usina" not in title and ("jfa" in title or "fonte carregador" in title or "fonte automotiva" in title or "fonte e carregador" in title or "carregador de baterias" in title):
            if "120a" in title or "120 " in title or "120 amperes" in title or "120amperes" in title or "120 a" in title:
                item['Produto2'] = "FONTE LITE 120A"

        if "bob" not in title and "lite" not in title and "light" not in title and "controle" not in title and 'mono' not in title and 'monovolt' not in title and "220v" not in title:
            if "200a" in title or "200 " in title or "200 amperes" in title or "200amperes" in title or "200 a" in title:
                item['Produto2'] = "FONTE 200A"

        if "bob" not in title and  ("lite" in title or "light" in title) and "controle" not in title and 'mono' not in title and 'monovolt' not in title:
            if "200a" in title or "200 " in title or "200 amperes" in title or "200amperes" in title or "200 a" in title:
                item['Produto2'] = "FONTE LITE 200A"


        if "bob" in title and "lite" not in title and "light" not in title  and "controle" not in title and "usina" not in title and ("jfa" in title or "fonte carregador" in title or "fonte automotiva" in title or "fonte e carregador" in title or "carregador de baterias" in title):
            if "120a" in title or "120 " in title or "120 amperes" in title or "120amperes" in title or "120 a" in title:
                item['Produto2'] = "FONTE BOB 120A"
                
        if "bob" in title and "lite" not in title and "light" not in title  and "controle" not in title and 'mono' not in title and 'mono' not in title and 'monovolt' not in title and "usina" not in title and ("jfa" in title or "fonte carregador" in title or "fonte automotiva" in title or "fonte e carregador" in title or "carregador de baterias" in title):
            if "200a" in title or "200 " in title or "200 amperes" in title or "200amperes" in title or "200 a" in title:
                item['Produto2'] = "FONTE BOB 200A"


        if "bob" not in title and "lite" not in title and "light" not in title  and "controle" not in title and ("mono" in title or "220v" in title or "monovolt" in title):
            if "200a" in title or "200 " in title or "200 amperes" in title or "200amperes" in title or "200 a" in title:
                item['Produto2'] = "FONTE 200A MONO"
                
        
        fonte = SelecionarFonte(item).split(",")
        novos_dados.loc[index, 'politica'] = fonte[0]
        novos_dados.loc[index, 'Produto2'] = item['Produto2']
        novos_dados.loc[index, 'preço_previsto'] = round(float(fonte[1]), 2)
    

    # Juntar os novos dados ao DataFrame original
    all_dados = novos_dados
    
    if os.path.exists('resultado.xlsx'):  
        existing_data = pd.read_excel("resultado.xlsx")
        all_dados = pd.concat([existing_data, novos_dados], ignore_index=True)
        all_dados.drop_duplicates(inplace=True)
        all_dados.to_excel("resultado.xlsx", index=False)
    else:  
        all_dados.to_excel("resultado.xlsx", index=False)

if __name__ == "__main__":
    # Verifica se os argumentos foram passados
    if len(sys.argv) != 3:
        print("Erro: As datas não foram passadas corretamente.")
        print("Uso: python main.py <dataInicial> <dataFinal>")
        sys.exit(1)
    
    # Recebe os argumentos
    dataInicial = sys.argv[1]
    dataFinal = sys.argv[2]
    
    main()