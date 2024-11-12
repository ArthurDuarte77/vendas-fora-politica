import joblib  # Para salvar e carregar o modelo
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

all_dados = pd.DataFrame()

titulo_arquivo = ""


start_row = 20  
end_row = 37
num_rows = end_row - start_row

df = pandas.read_excel("GESTÃO DE AÇÕES E-COMMERCE.xlsx", usecols='C:O', skiprows=start_row, nrows=num_rows, engine='openpyxl', sheet_name="POLÍTICA COMERCIAL Nov24")

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



def SelecionarFonte(item):
    price = item["Preço Unitário"]
    tipo = unidecode(item["Tipo de Anúncio"].strip().lower())
    if item['Produto2'] == "FONTE 40A":
        if tipo == "classico" and price < fonte40Classico:
            return f"FORA,{fonte40Classico + 0.01}"
        elif tipo == "premium" and price < fonte40Premium:
            return f"FORA,{fonte40Premium + 0.01}"

    if item['Produto2'] == "FONTE 60A":
        if tipo == "classico" and price < fonte60Classico:
            return f"FORA,{fonte60Classico + 0.01}"
        elif tipo == "premium" and price < fonte60Premium:
            return f"FORA,{fonte60Premium + 0.01}"

    if item['Produto2'] == "FONTE LITE 60A":
        if tipo == "classico" and price < fonte60liteClassico:
            return f"FORA,{fonte60liteClassico + 0.01}"
        elif tipo == "premium" and price < fonte60litePremium:
            return f"FORA,{fonte60litePremium + 0.01}"

    if item['Produto2'] == "FONTE 70A":
        if tipo == "classico" and price < fonte70Classico:
            return f"FORA,{fonte70Classico + 0.01}"
        elif tipo == "premium" and price < fonte70Premium:
            return f"FORA,{fonte70Premium + 0.01}"

    if item['Produto2'] == "FONTE LITE 70A":
        if tipo == "classico" and price < fonte70liteClassico:
            return f"FORA,{fonte70liteClassico + 0.01}"
        elif tipo == "premium" and price < fonte70litePremium:
            return f"FORA,{fonte70litePremium + 0.01}"

    if item['Produto2'] == "FONTE BOB 90A":
        if tipo == "classico" and price < fonte90bobClassico:
            return f"FORA,{fonte90bobClassico + 0.01}"
        elif tipo == "premium" and price < fonte90bobPremium:
            return f"FORA,{fonte90bobPremium + 0.01}"

    if item['Produto2'] == "FONTE 120A":
        if tipo == "classico" and price < fonte120Classico:
            return f"FORA,{fonte120Classico + 0.01}"
        elif tipo == "premium" and price < fonte120Premium:
            return f"FORA,{fonte120Premium + 0.01}"

    if item['Produto2'] == "FONTE LITE 120A":
        if tipo == "classico" and price < fonte120liteClassico:
            return f"FORA,{fonte120liteClassico + 0.01}"
        elif tipo == "premium" and price < fonte120litePremium:
            return f"FORA,{fonte120litePremium + 0.01}"

    if item['Produto2'] == "FONTE BOB 120A":
        if tipo == "classico" and price < fonte120bobClassico:
            return f"FORA,{fonte120bobClassico + 0.01}"
        elif tipo == "premium" and price < fonte120bobPremium:
            return f"FORA,{fonte120bobPremium + 0.01}"

    if item['Produto2'] == "FONTE 200A":
        if tipo == "classico" and price < fonte200Classico:
            return f"FORA,{fonte200Classico + 0.01}"
        elif tipo == "premium" and price < fonte200Premium:
            return f"FORA,{fonte200Premium + 0.01}"

    if item['Produto2'] == "FONTE MONO 200A":
        if tipo == "classico" and price < fonte200monoClassico:
            return f"FORA,{fonte200monoClassico + 0.01}"
        elif tipo == "premium" and price < fonte200monoPremium:
            return f"FORA,{fonte200monoPremium + 0.01}"

    if item['Produto2'] == "FONTE LITE 200A":
        if tipo == "classico" and price < fonte200liteClassico:
            return f"FORA,{fonte200liteClassico + 0.01}"
        elif tipo == "premium" and price < fonte200litePremium:
            return f"FORA,{fonte200litePremium + 0.01}"

    if item['Produto2'] == "FONTE BOB 200A":
        if tipo == "classico" and price < fonte200bobClassico:
            return f"FORA,{fonte200bobClassico + 0.01}"
        elif tipo == "premium" and price < fonte200bobPremium:
            return f"FORA,{fonte200bobPremium + 0.01}"
        
    return "DENTRO,0"
                
    
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

all_dados = pd.DataFrame()  # Inicializar o DataFrame all_dados

for i in urls:
    response = requests.get(f"https://corp.shoppingdeprecos.com.br/vendedores/exportar_vendas_marca?id={i}&ini={dia_inicial}&fim={dia_final}", headers=headers)

    if response.status_code == 200:  
        with open("produtos.xlsx", 'wb') as file:
            file.write(response.content)

    time.sleep(5)

    novos_dados = pd.read_excel("produtos.xlsx", engine='openpyxl')
    novos_dados['Preço Unitário'] = novos_dados['Preço Unitário'].str.replace('.', '', regex=False).str.replace(',', '.', regex=False).astype(float)
    # Carregar o pipeline treinado
    pipeline_carregado = joblib.load('modelo_treinado.pkl')

    # Fazer previsões nos novos dados
    previsoes = pipeline_carregado.predict(novos_dados)

    # Carregar o label encoder
    label_encoder_carregado = joblib.load('label_encoder.pkl')

    # Decodificar as previsões para obter os nomes das classes
    nomes_classes = label_encoder_carregado.inverse_transform(previsoes)

    # Adicionar as previsões ao DataFrame original
    novos_dados['Produto2'] = nomes_classes
                    
    for index, item in novos_dados.iterrows():
        novos_dados.loc[index, 'politica'] = SelecionarFonte(item).split(",")[0]
        novos_dados.loc[index, 'preço_previsto'] = round(float(SelecionarFonte(item).split(",")[1]),2)

        
    all_dados = pd.concat([all_dados, novos_dados])
all_dados.to_excel("market-share-historico.xlsx", index=False)


# grouped_by_seller = defaultdict(list)

# for item in items:
#     seller = item['seller']
#     grouped_by_seller[seller].append(item)
    
# grouped_by_seller = dict(grouped_by_seller)
# enviar(grouped_by_seller)
    
    