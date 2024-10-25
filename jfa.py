from datetime import datetime
from collections import defaultdict
import argparse
from unidecode import unidecode
from selenium.webdriver.support.ui import Select
import threading
import subprocess
import os
import time
from tqdm import tqdm
import shutil
import json
from tqdm import tqdm
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
from selenium.common.exceptions import *
import re
import sys
import numpy as np
import cv2
import requests
from typing import Dict, List
import pandas

items = []

titulo_arquivo = ""
# options.add_argument("--headless=new")

def get_greeting():
    current_hour = datetime.now().hour
    if 5 <= current_hour < 12:
        return "Bom dia!"
    elif 12 <= current_hour < 18:
        return "Boa tarde!"
    else:
        return "Boa noite!"

def get_loja(loja):
    response = requests.get(f"https://api.mercadolibre.com/sites/MLB/search?nickname={loja}")
    user_id = response.json()['results'][0]['seller']['id']
    user_response = requests.get(f"https://api.mercadolibre.com/users/{user_id}")
    address = user_response.json()['address']['city']
    state = user_response.json()['address']['state']
    return address + " - " + state

def enviar(grouped_by_seller):
    print(grouped_by_seller)
    return
    if len(grouped_by_seller) > 0:
        
        requests.post("http://134.122.29.170:3000/api/sendText", {
            "chatId": "120363330531801612@g.us",
            "text": f"{get_greeting()} \n Segue vendas fora da política",
            "session": "default"
        })
        try:
            # print(grouped_by_seller)
            for seller, products in grouped_by_seller.items():
                dados = f"*{seller}* \n"
                time.sleep(1)
                for item in products:
                    if item['listing_type'] == "gold_special":
                        item['listing_type'] = "Clássico"
                    else:
                        item['listing_type'] = "Premium"
                    
                    # loja_info = get_loja(item['seller'])
                    dados =  dados + f"{item['model']} - {item['seller']} - Preço Anúncio: {item['price']} - Preço Política: {round(item['predicted_price'], 2)} ({item['listing_type']}) \n titulo: {item['title']} \n quantidade: {item['qtd']} \n total: {item['total']} \n"
                requests.post("http://134.122.29.170:3000/api/sendText", {
                "chatId": "120363330531801612@g.us",
                "text": dados,
                "session": "default"
                })
        except Exception as e:
            print(f"Erro ao enviar mensagens: {e}")   
    else:
        requests.post("http://134.122.29.170:3000/api/sendText", {
            "chatId": "120363330531801612@g.us",
            "text": f"{get_greeting()} \n Nada fora da política",
            "session": "default"
        })


start_row = 20  
end_row = 37
num_rows = end_row - start_row

df = pandas.read_excel("GESTÃO DE AÇÕES E-COMMERCE.xlsx", usecols='C:O', skiprows=start_row, nrows=num_rows, engine='openpyxl', sheet_name="POLÍTICA COMERCIAL Out24 II")

df.columns = ['PRODUTO', 'inutil1', 'SITE', 'COLUNA3','inutil2', 'CLÁSSICO ML', 'COLUNA5','inutil3', 'PREMIUM ML', 'COLUNA7','inutil4', 'MARKETPLACES', 'COLUNA9']

for index, i in df.iterrows():
    if i['PRODUTO'] == "FONTE 40A":
        fonte40Marketplace = round(i['COLUNA3']- 0.01 , 2) ;
        fonte40Classico = round(i['COLUNA5']- 0.01 , 2) ;
        fonte40Premium = round(i['COLUNA7']- 0.01 , 2) ;
        fonte40PremiumPrice = round(i['PREMIUM ML']- 0.01 , 2) ;
        fonte40ClassicoPrice = round(i['CLÁSSICO ML']- 0.01 , 2) ;
        fonte40Marketplaceprice = round(i['SITE']- 0.01 , 2) ;
    elif i['PRODUTO'] == "FONTE 60A":
        fonte60Marketplace = round(i['COLUNA3']- 0.01 , 2) ;
        fonte60Classico = round(i['COLUNA5']- 0.01 , 2) ;
        fonte60Premium = round(i['COLUNA7']- 0.01 , 2) ;
        fonte60PremiumPrice = round(i['PREMIUM ML']- 0.01 , 2) ;
        fonte60ClassicoPrice = round(i['CLÁSSICO ML']- 0.01 , 2) ;
        fonte60Marketplaceprice = round(i['SITE']- 0.01 , 2) ;
    elif i['PRODUTO'] == "FONTE 60A LITE":
        fonte60liteMarketplace = round(i['COLUNA3']- 0.01 , 2) ;
        fonte60liteClassico = round(i['COLUNA5']- 0.01 , 2) ;
        fonte60litePremium = round(i['COLUNA7']- 0.01 , 2) ;
        fonte60litePremiumPrice = round(i['PREMIUM ML']- 0.01 , 2) ;
        fonte60liteClassicoPrice = round(i['CLÁSSICO ML']- 0.01 , 2) ;
        fonte60liteMarketplaceprice = round(i['SITE']- 0.01 , 2) ;
    elif i['PRODUTO'] == "FONTE 70A":
        fonte70Marketplace = round(i['COLUNA3']- 0.01 , 2) ;
        fonte70Classico = round(i['COLUNA5']- 0.01 , 2) ;
        fonte70Premium = round(i['COLUNA7']- 0.01 , 2) ;
        fonte70PremiumPrice = round(i['PREMIUM ML']- 0.01 , 2) ;
        fonte70ClassicoPrice = round(i['CLÁSSICO ML']- 0.01 , 2) ;
        fonte70Marketplaceprice = round(i['SITE']- 0.01 , 2) ;
    elif i['PRODUTO'] == "FONTE 70A LITE":
        fonte70liteMarketplace = round(i['COLUNA3']- 0.01 , 2) ;
        fonte70liteClassico = round(i['COLUNA5']- 0.01 , 2) ;
        fonte70litePremium = round(i['COLUNA7']- 0.01 , 2) ;
        fonte70litePremiumPrice = round(i['PREMIUM ML']- 0.01 , 2) ;
        fonte70liteClassicoPrice = round(i['CLÁSSICO ML']- 0.01 , 2) ;
        fonte70liteMarketplaceprice = round(i['SITE']- 0.01 , 2) ;
    elif i['PRODUTO'] == "FONTE 90 BOB":
        fonte90bobMarketplace = round(i['COLUNA3']- 0.01 , 2) ;
        fonte90bobClassico = round(i['COLUNA5']- 0.01 , 2) ;
        fonte90bobPremium = round(i['COLUNA7']- 0.01 , 2) ;
        fonte90bobPremiumPrice = round(i['PREMIUM ML']- 0.01 , 2) ;
        fonte90bobClassicoPrice = round(i['CLÁSSICO ML']- 0.01 , 2) ;
        fonte90bobMarketplaceprice = round(i['SITE']- 0.01 , 2) ;
    elif i['PRODUTO'] == "FONTE 120 BOB":
        fonte120bobMarketplace = round(i['COLUNA3']- 0.01 , 2) ;
        fonte120bobClassico = round(i['COLUNA5']- 0.01 , 2) ;
        fonte120bobPremium = round(i['COLUNA7']- 0.01 , 2) ;
        fonte120bobPremiumPrice = round(i['PREMIUM ML']- 0.01 , 2) ;
        fonte120bobClassicoPrice = round(i['CLÁSSICO ML']- 0.01 , 2) ;
        fonte120bobMarketplaceprice = round(i['SITE']- 0.01 , 2) ;
    elif i['PRODUTO'] == "FONTE 120A LITE":
        fonte120liteMarketplace = round(i['COLUNA3']- 0.01 , 2) ;
        fonte120liteClassico = round(i['COLUNA5']- 0.01 , 2) ;
        fonte120litePremium = round(i['COLUNA7']- 0.01 , 2) ;
        fonte120litePremiumPrice = round(i['PREMIUM ML']- 0.01 , 2) ;
        fonte120liteClassicoPrice = round(i['CLÁSSICO ML']- 0.01 , 2) ;
        fonte120liteMarketplaceprice = round(i['SITE']- 0.01 , 2) ;
    elif i['PRODUTO'] == "FONTE 120A":
        fonte120Marketplace = round(i['COLUNA3']- 0.01 , 2) ;
        fonte120Classico = round(i['COLUNA5']- 0.01 , 2) ;
        fonte120Premium = round(i['COLUNA7']- 0.01 , 2) ;
        fonte120PremiumPrice = round(i['PREMIUM ML']- 0.01 , 2) ;
        fonte120ClassicoPrice = round(i['CLÁSSICO ML']- 0.01 , 2) ;
        fonte120Marketplaceprice = round(i['SITE']- 0.01 , 2) ;
    elif i['PRODUTO'] == "FONTE 200 BOB":
        fonte200bobMarketplace = round(i['COLUNA3']- 0.01 , 2) ;
        fonte200bobClassico = round(i['COLUNA5']- 0.01 , 2) ;
        fonte200bobPremium = round(i['COLUNA7']- 0.01 , 2) ;
        fonte200bobPremiumPrice = round(i['PREMIUM ML']- 0.01 , 2) ;
        fonte200bobClassicoPrice = round(i['CLÁSSICO ML']- 0.01 , 2) ;
        fonte200bobMarketplaceprice = round(i['SITE']- 0.01 , 2) ;
    elif i['PRODUTO'] == "FONTE 200A LITE":
        fonte200liteMarketplace = round(i['COLUNA3']- 0.01 , 2) ;
        fonte200liteClassico = round(i['COLUNA5']- 0.01 , 2) ;
        fonte200litePremium = round(i['COLUNA7']- 0.01 , 2) ;
        fonte200litePremiumPrice = round(i['PREMIUM ML']- 0.01 , 2) ;
        fonte200liteClassicoPrice = round(i['CLÁSSICO ML']- 0.01 , 2) ;
        fonte200liteMarketplaceprice = round(i['SITE']- 0.01 , 2) ;
    elif i['PRODUTO'] == "FONTE 200 MONO":
        fonte200monoMarketplace = round(i['COLUNA3']- 0.01 , 2) ;
        fonte200monoClassico = round(i['COLUNA5']- 0.01 , 2) ;
        fonte200monoPremium = round(i['COLUNA7']- 0.01 , 2) ;
        fonte200monoPremiumPrice = round(i['PREMIUM ML']- 0.01 , 2) ;
        fonte200monoClassicoPrice = round(i['CLÁSSICO ML']- 0.01 , 2) ;
        fonte200monoMarketplaceprice = round(i['SITE']- 0.01 , 2) ;
    elif i['PRODUTO'] == "FONTE 200A":
        fonte200Marketplace = round(i['COLUNA3']- 0.01 , 2) ;
        fonte200Classico = round(i['COLUNA5']- 0.01 , 2) ;
        fonte200Premium = round(i['COLUNA7']- 0.01 , 2) ;
        fonte200PremiumPrice = round(i['PREMIUM ML']- 0.01 , 2) ;
        fonte200ClassicoPrice = round(i['CLÁSSICO ML']- 0.01 , 2) ;
        fonte200Marketplaceprice = round(i['SITE']- 0.01 , 2) ;
    elif i['PRODUTO'] == "K1200":
        controleK1200Marketplace = round(i['COLUNA3']- 0.01 , 2) ;
        controleK1200Classico = round(i['COLUNA5']- 0.01 , 2) ;
        controleK1200Premium = round(i['COLUNA7']- 0.01 , 2) ;
        controleK1200PremiumPrice = round(i['PREMIUM ML']- 0.01 , 2) ;
        controleK1200ClassicoPrice = round(i['CLÁSSICO ML']- 0.01 , 2) ;
        controleK1200Marketplaceprice = round(i['SITE']- 0.01 , 2) ;
    elif i['PRODUTO'] == "K600":
        controleK600Marketplace = round(i['COLUNA3']- 0.01 , 2) ;
        controleK600Classico = round(i['COLUNA5']- 0.01 , 2) ;
        controleK600Premium = round(i['COLUNA7']- 0.01 , 2) ;
        controleK600PremiumPrice = round(i['PREMIUM ML']- 0.01 , 2) ;
        controleK600ClassicoPrice = round(i['CLÁSSICO ML']- 0.01 , 2) ;
        controleK600Marketplaceprice = round(i['SITE']- 0.01 , 2) ;
    elif i['PRODUTO'] == "CONTROLE WR":
        controleRedlineMarketplace = round(i['COLUNA3']- 0.01 , 2) ;
        controleRedlineClassico = round(i['COLUNA5']- 0.01 , 2) ;
        controleRedlinePremium = round(i['COLUNA7']- 0.01 , 2) ;
        controleRedlinePremiumPrice = round(i['PREMIUM ML']- 0.01 , 2) ;
        controleRedlineClassicoPrice = round(i['CLÁSSICO ML']- 0.01 , 2) ;
        controleRedlineMarketplaceprice = round(i['SITE']- 0.01 , 2) ;
    elif i['PRODUTO'] == "ACQUA":
        controleAcquaMarketplace = round(i['COLUNA3']- 0.01 , 2) ;
        controleAcquaClassico = round(i['COLUNA5']- 0.01 , 2) ;
        controleAcquaPremium = round(i['COLUNA7']- 0.01 , 2) ;
        controleAcquaPremiumPrice = round(i['PREMIUM ML']- 0.01 , 2) ;
        controleAcquaClassicoPrice = round(i['CLÁSSICO ML']- 0.01 , 2) ;
        controleAcquaMarketplaceprice = round(i['SITE']- 0.01 , 2) ;


# Dicionário com produtos e seus preços para cada categoria
produtos = {
    "FONTE 40A": {"classico": fonte40Classico, "premium": fonte40Premium},
    "FONTE 60A": {"classico": fonte60Classico, "premium": fonte60Premium},
    "FONTE LITE 60A": {"classico": fonte60liteClassico, "premium": fonte60litePremium},
    "FONTE 70A": {"classico": fonte70Classico, "premium": fonte70Premium},
    "FONTE LITE 70A": {"classico": fonte70liteClassico, "premium": fonte70litePremium},
    "FONTE 120A": {"classico": fonte120Classico, "premium": fonte120Premium},
    "FONTE LITE 120A": {"classico": fonte120liteClassico, "premium": fonte120litePremium},
    "FONTE 200A": {"classico": fonte200Classico, "premium": fonte200Premium},
    "FONTE LITE 200A": {"classico": fonte200liteClassico, "premium": fonte200litePremium},
    "FONTE BOB 90A": {"classico": fonte90bobClassico, "premium": fonte90bobPremium},
    "FONTE BOB 120A": {"classico": fonte120bobClassico, "premium": fonte120bobPremium},
    "FONTE BOB 200A": {"classico": fonte200bobClassico, "premium": fonte200bobPremium},
    "FONTE 200A MONO": {"classico": fonte200monoClassico, "premium": fonte200monoPremium},
    "CONTROLE K1200": {"classico": controleK1200Classico, "premium": controleK1200Premium},
    "CONTROLE K600": {"classico": controleK600Classico, "premium": controleK600Premium},
    "CONTROLE REDLINE": {"classico": controleRedlineClassico, "premium": controleRedlinePremium},
    "CONTROLE ACQUA": {"classico": controleAcquaClassico, "premium": controleAcquaPremium}
}

def identificar_produto(tipo, preco):
    tolerancia = 0.05  # Tolerância de 1%
    for produto, precos in produtos.items():
        if tipo.lower() == "classico":
            preco_base = precos["classico"]
        elif tipo.lower() == "premium":
            preco_base = precos["premium"]
        else:
            return "Tipo inválido. Use 'classico' ou 'premium'."
        
        if preco_base * (1 - tolerancia) <= preco <= preco_base * (1 + tolerancia):
            return produto
    return "OUTROS"

if os.path.exists(r"produtos.xlsx"):
    os.remove(r"produtos.xlsx")
if os.path.exists(r"modelos_jfa.xlsx"):
    os.remove(r"modelos_jfa.xlsx")

    


def SelecionarFonte(item):
    nome = unidecode(item["Produto"].strip().lower())
    price = float(item["Preço Unitário"].replace(".", "").replace(",", "."))
    tipo = unidecode(item["Tipo de Anúncio"].strip().lower())
    total = float(item["Total"].replace(".", "").replace(",", "."))
    if "inversor" in nome or "amplificador" in nome or "processador" in nome or "capa" in nome or "nobreak" in nome or "retificadora" in nome or "multimidia" in nome or "gerenciador" in nome or "suspensao" in nome or "stetsom" in nome or "central" in nome or 'k600' in nome or 'k1200' in nome or "fonte" not in nome:
        return
    
    if "controle" not in nome and "lite" not in nome and "light" not in nome:
        if "40" in nome or "40a" in nome or "40 amperes" in nome or "40amperes" in nome or "36a" in nome or "36" in nome or "36 amperes" in nome or "36amperes" in nome:
            if tipo == "classico" and price < fonte40Classico:
                items.append({"seller": item["Vendedor"], "title": nome,"listing_type": tipo, "price": price, "qtd": item["Qtde"], "total": total, "model": "FONTE 40A", "predicted_price": fonte40Classico})
                return
            elif tipo == "premium" and price < fonte40Premium:
                items.append({"seller": item["Vendedor"], "title": nome,"listing_type": tipo, "price": price, "qtd": item["Qtde"], "total": total, "model": "FONTE 40A", "predicted_price": fonte40Premium})
                return
            
            
    if "bob" not in nome and "lite" not in nome and "light" not in nome  and "controle" not in nome:
        if "60" in nome or "60a" in nome or "60 amperes" in nome or "60amperes" in nome or "60 a" in nome or "-60" in nome:
            if tipo == "classico" and price < fonte60Classico:
                items.append({"seller": item["Vendedor"], "title": nome,"listing_type": tipo, "price": price, "qtd": item["Qtde"], "total": total, "model": "FONTE 60A", "predicted_price": fonte60Classico})
                return
            elif tipo == "premium" and price < fonte60Premium:
                items.append({"seller": item["Vendedor"], "title": nome,"listing_type": tipo, "price": price, "qtd": item["Qtde"], "total": total, "model": "FONTE 60A", "predicted_price": fonte60Premium})
                return
            
    if "bob" not in nome and ("lite" in nome or "light" in nome or "lit" in nome) and "controle" not in nome:
        if "60" in nome or "60a" in nome or "60 amperes" in nome or "60amperes" in nome or "60 a" in nome:
            if tipo == "classico" and price < fonte60liteClassico:
                items.append({"seller": item["Vendedor"], "title": nome,"listing_type": tipo, "price": price, "qtd": item["Qtde"], "total": total, "model": "FONTE LITE 60A", "predicted_price": fonte60liteClassico})
                return
            elif tipo == "premium" and price < fonte60litePremium:
                items.append({"seller": item["Vendedor"], "title": nome,"listing_type": tipo, "price": price, "qtd": item["Qtde"], "total": total, "model": "FONTE LITE 60A", "predicted_price": fonte60litePremium})
                return
            
    if "bob" not in nome and "lite" not in nome and "light" not in nome  and "controle" not in nome:
        if "70" in nome or "70a" in nome or "70 amperes" in nome or "70amperes" in nome or "70 a" in nome:
            if tipo == "classico" and price < fonte70Classico:
                items.append({"seller": item["Vendedor"], "title": nome,"listing_type": tipo, "price": price, "qtd": item["Qtde"], "total": total, "model": "FONTE 70A", "predicted_price": fonte70Classico})
                return
            elif tipo == "premium" and price < fonte70Premium:
                items.append({"seller": item["Vendedor"], "title": nome,"listing_type": tipo, "price": price, "qtd": item["Qtde"], "total": total, "model": "FONTE 70A", "predicted_price": fonte70Premium})
                return

    if "bob" not in nome and  ("lite" in nome or "light" in nome or "lit" in nome) and "controle" not in nome:
        if "70" in nome or "70a" in nome or "70 amperes" in nome or "70amperes" in nome or "70 a" in nome:
            if tipo == "classico" and price < fonte70liteClassico:
                items.append({"seller": item["Vendedor"], "title": nome,"listing_type": tipo, "price": price, "qtd": item["Qtde"], "total": total, "model": "FONTE LITE 70A", "predicted_price": fonte70liteClassico})
                return
            elif tipo == "premium" and price < fonte70litePremium:
                items.append({"seller": item["Vendedor"], "title": nome,"listing_type": tipo, "price": price, "qtd": item["Qtde"], "total": total, "model": "FONTE LITE 70A", "predicted_price": fonte70litePremium})
                return

            
    if "lite" not in nome and "light" not in nome  and "controle" not in nome:
        if "90" in nome or "90a" in nome or "90 amperes" in nome or "90amperes" in nome or "90 a" in nome:
            if tipo == "classico" and price < fonte90bobClassico:
                items.append({"seller": item["Vendedor"], "title": nome,"listing_type": tipo, "price": price, "qtd": item["Qtde"], "total": total, "model": "FONTE BOB 90A", "predicted_price": fonte90bobClassico})
                return
            elif tipo == "premium" and price < fonte90bobPremium:
                items.append({"seller": item["Vendedor"], "title": nome,"listing_type": tipo, "price": price, "qtd": item["Qtde"], "total": total, "model": "FONTE BOB 90A", "predicted_price": fonte90bobPremium})
                return
            
    if "bob" not in nome and "lite" not in nome and "light" not in nome  and "controle" not in nome and "lit" not in nome:
        if "120" in nome or "120a" in nome or "120 amperes" in nome or "120amperes" in nome or "120 a" in nome:
            if tipo == "classico" and price < fonte120Classico:
                items.append({"seller": item["Vendedor"], "title": nome,"listing_type": tipo, "price": price, "qtd": item["Qtde"], "total": total, "model": "FONTE 120A", "predicted_price": fonte120Classico})
                return
            elif tipo == "premium" and price < fonte120Premium:
                items.append({"seller": item["Vendedor"], "title": nome,"listing_type": tipo, "price": price, "qtd": item["Qtde"], "total": total, "model": "FONTE 120A", "predicted_price": fonte120Premium})
                return


             
    if "bob" not in nome and  ("lite" in nome or "light" in nome or "lit" in nome) and "controle" not in nome:
        if "120" in nome or "120a" in nome or "120 amperes" in nome or "120amperes" in nome or "120 a" in nome:
            if tipo == "classico" and price < fonte120liteClassico:
                items.append({"seller": item["Vendedor"], "title": nome,"listing_type": tipo, "price": price, "qtd": item["Qtde"], "total": total, "model": "FONTE LITE 120A", "predicted_price": fonte120liteClassico})
                return
            elif tipo == "premium" and price < fonte120litePremium:
                items.append({"seller": item["Vendedor"], "title": nome,"listing_type": tipo, "price": price, "qtd": item["Qtde"], "total": total, "model": "FONTE LITE 120A", "predicted_price": fonte120litePremium})
                return
                
    if "bob" in nome and "lite" not in nome and "light" not in nome  and "controle" not in nome and "lit" not in nome:
        if "120" in nome or "120a" in nome or "120 amperes" in nome or "120amperes" in nome or "120 a" in nome:
            if tipo == "classico" and price < fonte120bobClassico:
                items.append({"seller": item["Vendedor"], "title": nome,"listing_type": tipo, "price": price, "qtd": item["Qtde"], "total": total, "model": "FONTE BOB 120A", "predicted_price": fonte120bobClassico})
                return
            elif tipo == "premium" and price < fonte120bobPremium:
                items.append({"seller": item["Vendedor"], "title": nome,"listing_type": tipo, "price": price, "qtd": item["Qtde"], "total": total, "model": "FONTE BOB 120A", "predicted_price": fonte120bobPremium})
                return
                
    if "bob" not in nome and "lite" not in nome and "light" not in nome  and "controle" not in nome and 'mono' not in nome and 'mono' not in nome and 'monovolt' not in nome and '220v' not in nome and "lit" not in nome:
        if "200" in nome or "200a" in nome or "200 amperes" in nome or "200amperes" in nome or "200 a" in nome:
            if tipo == "classico" and price < fonte200Classico:
                items.append({"seller": item["Vendedor"], "title": nome,"listing_type": tipo, "price": price, "qtd": item["Qtde"], "total": total, "model": "FONTE 200A", "predicted_price": fonte200Classico})
                return
            elif tipo == "premium" and price < fonte200Premium:
                items.append({"seller": item["Vendedor"], "title": nome,"listing_type": tipo, "price": price, "qtd": item["Qtde"], "total": total, "model": "FONTE 200A", "predicted_price": fonte200Premium})
                return

    if "bob" not in nome and "lite" not in nome and "light" not in nome  and "controle" not in nome and ("mono" in nome or "220v" in nome or "monovolt" in nome):
        if "200" in nome or "200a" in nome or "200 amperes" in nome or "200amperes" in nome or "200 a" in nome:
            if tipo == "classico" and price < fonte200monoClassico:
                items.append({"seller": item["Vendedor"], "title": nome,"listing_type": tipo, "price": price, "qtd": item["Qtde"], "total": total, "model": "FONTE 200A MONO", "predicted_price": fonte200monoClassico})
                return
            elif tipo == "premium" and price < fonte200monoPremium:
                items.append({"seller": item["Vendedor"], "title": nome,"listing_type": tipo, "price": price, "qtd": item["Qtde"], "total": total, "model": "FONTE 200A MONO", "predicted_price": fonte200monoPremium})
                return
                
    if "bob" not in nome and  ("lite" in nome or "light" in nome or "lit" in nome) and "controle" not in nome and 'mono' not in nome:
        if "200" in nome or "200a" in nome or "200 amperes" in nome or "200amperes" in nome or "200 a" in nome:
            if tipo == "classico" and price < fonte200liteClassico:
                items.append({"seller": item["Vendedor"], "title": nome,"listing_type": tipo, "price": price, "qtd": item["Qtde"], "total": total, "model": "FONTE LITE 200A", "predicted_price": fonte200liteClassico})
                return
            elif tipo == "premium" and price < fonte200litePremium:
                items.append({"seller": item["Vendedor"], "title": nome,"listing_type": tipo, "price": price, "qtd": item["Qtde"], "total": total, "model": "FONTE LITE 200A", "predicted_price": fonte200litePremium})
                return
                
    if "bob" in nome and "lite" not in nome and "light" not in nome  and "controle" not in nome and 'mono' not in nome and 'mono' not in nome and 'monovolt' not in nome and '220v' not in nome:
        if "200" in nome or "200a" in nome or "200 amperes" in nome or "200amperes" in nome or "200 a" in nome:
            if tipo == "classico" and price < fonte200bobClassico:
                items.append({"seller": item["Vendedor"], "title": nome,"listing_type": tipo, "price": price, "qtd": item["Qtde"], "total": total, "model": "FONTE BOB 200A", "predicted_price": fonte200bobClassico})
                return
            elif tipo == "premium" and price < fonte200bobPremium:
                items.append({"seller": item["Vendedor"], "title": nome,"listing_type": tipo, "price": price, "qtd": item["Qtde"], "total": total, "model": "FONTE BOB 200A", "predicted_price": fonte200bobPremium})
                return
                
    
       

parser = argparse.ArgumentParser(description='Processar datas de início e fim.')
parser.add_argument('--dia_inicial', type=str, required=True, help='Data inicial no formato AAAA-MM-DD')
parser.add_argument('--dia_final', type=str, required=True, help='Data final no formato AAAA-MM-DD')
parser.add_argument('--cookie', type=str, required=True, help='Cookies')

args = parser.parse_args()

dia_inicial = args.dia_inicial
dia_final = args.dia_final
cookie = args.cookie

headers = {
    "Cookie": cookie
}

urls = ["JFA", "JFA%20ELETRONICOS"]             
for i in urls:
    response = requests.get(f"https://corp.shoppingdeprecos.com.br/vendedores/exportar_vendas_marca?id={i}&ini={dia_inicial}&fim={dia_final}", headers=headers)

    if response.status_code == 200:  
        print("resposta ok")
        time.sleep(20)
        with open("produtos.xlsx", 'wb') as file:

            file.write(response.content)

    time.sleep(5)



    db = pd.read_excel("produtos.xlsx", engine='openpyxl')
                    
    for index, item in db.iterrows():
        SelecionarFonte(item)

grouped_by_seller = defaultdict(list)

for item in items:
    seller = item['seller']
    grouped_by_seller[seller].append(item)
    
grouped_by_seller = dict(grouped_by_seller)
enviar(grouped_by_seller)
    
    