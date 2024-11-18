from unidecode import unidecode
import tqdm
import joblib  # Para salvar e carregar o modelo
import requests
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
all_dados = pd.DataFrame()

start_row = 20  
end_row = 37
num_rows = end_row - start_row

df = pd.read_excel("GESTÃO DE AÇÕES E-COMMERCE.xlsx", usecols='C:O', skiprows=start_row, nrows=num_rows, engine='openpyxl', sheet_name="POLÍTICA COMERCIAL Nov24")

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
    tipo = unidecode(item["tipo"].strip().lower())
    if item['Produto2'] == "FONTE 40A":
        if tipo == "gold_special" and price < fonte40Classico:
            return f"FORA,{fonte40Classico + 0.01}"
        elif tipo == "gold_pro" and price < fonte40Premium:
            return f"FORA,{fonte40Premium + 0.01}"

    if item['Produto2'] == "FONTE 60A":
        if tipo == "gold_special" and price < fonte60Classico:
            return f"FORA,{fonte60Classico + 0.01}"
        elif tipo == "gold_pro" and price < fonte60Premium:
            return f"FORA,{fonte60Premium + 0.01}"

    if item['Produto2'] == "FONTE LITE 60A":
        if tipo == "gold_special" and price < fonte60liteClassico:
            return f"FORA,{fonte60liteClassico + 0.01}"
        elif tipo == "gold_pro" and price < fonte60litePremium:
            return f"FORA,{fonte60litePremium + 0.01}"

    if item['Produto2'] == "FONTE 70A":
        if tipo == "gold_special" and price < fonte70Classico:
            return f"FORA,{fonte70Classico + 0.01}"
        elif tipo == "gold_pro" and price < fonte70Premium:
            return f"FORA,{fonte70Premium + 0.01}"

    if item['Produto2'] == "FONTE LITE 70A":
        if tipo == "gold_special" and price < fonte70liteClassico:
            return f"FORA,{fonte70liteClassico + 0.01}"
        elif tipo == "gold_pro" and price < fonte70litePremium:
            return f"FORA,{fonte70litePremium + 0.01}"

    if item['Produto2'] == "FONTE BOB 90A":
        if tipo == "gold_special" and price < fonte90bobClassico:
            return f"FORA,{fonte90bobClassico + 0.01}"
        elif tipo == "gold_pro" and price < fonte90bobPremium:
            return f"FORA,{fonte90bobPremium + 0.01}"

    if item['Produto2'] == "FONTE 120A":
        if tipo == "gold_special" and price < fonte120Classico:
            return f"FORA,{fonte120Classico + 0.01}"
        elif tipo == "gold_pro" and price < fonte120Premium:
            return f"FORA,{fonte120Premium + 0.01}"

    if item['Produto2'] == "FONTE LITE 120A":
        if tipo == "gold_special" and price < fonte120liteClassico:
            return f"FORA,{fonte120liteClassico + 0.01}"
        elif tipo == "gold_pro" and price < fonte120litePremium:
            return f"FORA,{fonte120litePremium + 0.01}"

    if item['Produto2'] == "FONTE BOB 120A":
        if tipo == "gold_special" and price < fonte120bobClassico:
            return f"FORA,{fonte120bobClassico + 0.01}"
        elif tipo == "gold_pro" and price < fonte120bobPremium:
            return f"FORA,{fonte120bobPremium + 0.01}"

    if item['Produto2'] == "FONTE 200A":
        if tipo == "gold_special" and price < fonte200Classico:
            return f"FORA,{fonte200Classico + 0.01}"
        elif tipo == "gold_pro" and price < fonte200Premium:
            return f"FORA,{fonte200Premium + 0.01}"

    if item['Produto2'] == "FONTE MONO 200A":
        if tipo == "gold_special" and price < fonte200monoClassico:
            return f"FORA,{fonte200monoClassico + 0.01}"
        elif tipo == "gold_pro" and price < fonte200monoPremium:
            return f"FORA,{fonte200monoPremium + 0.01}"

    if item['Produto2'] == "FONTE LITE 200A":
        if tipo == "gold_special" and price < fonte200liteClassico:
            return f"FORA,{fonte200liteClassico + 0.01}"
        elif tipo == "gold_pro" and price < fonte200litePremium:
            return f"FORA,{fonte200litePremium + 0.01}"

    if item['Produto2'] == "FONTE BOB 200A":
        if tipo == "gold_special" and price < fonte200bobClassico:
            return f"FORA,{fonte200bobClassico + 0.01}"
        elif tipo == "gold_pro" and price < fonte200bobPremium:
            return f"FORA,{fonte200bobPremium + 0.01}"
        
    return "DENTRO,0"


# service = Service()
# options = webdriver.ChromeOptions()
# options.add_argument("--disable-gpu")
# options.add_argument("--disable-extensions")
# prefs = {"profile.managed_default_content_settings.images": 2}
# options.add_experimental_option("prefs", prefs)

# driver = webdriver.Chrome(service=service, options=options)
# driver.get("https://www.google.com.br/?hl=pt-BR")
# time.sleep(3)
# try:
#     driver.get("https://corp.shoppingdeprecos.com.br/login")
#     counter = 0
#     while True:
#         test = driver.find_elements(By.XPATH, '//*[@id="email"]')
#         if test:
#             break
#         else:
#             counter += 1
#             if counter > 20:
#                 break
#             time.sleep(0.5)
#     driver.find_element(By.XPATH, '//*[@id="email"]').send_keys("loja@jfaeletronicos.com")
#     driver.find_element(By.XPATH, '//*[@id="password"]').send_keys("922982PC")
#     driver.find_element(By.XPATH, '//*[@id="btnLogin"]').click()
# except TimeoutException as e:
#     print(f"Timeout ao tentar carregar a página ou encontrar um elemento: {e}")
# except NoSuchElementException as e:
#     print(f"Elemento não encontrado na página: {e}")
# except WebDriverException as e:
#     print(f"Erro no WebDriver: {e}")

# time.sleep(3)
# driver.get("https://corp.shoppingdeprecos.com.br/vendedores/vendasMarca")

# time.sleep(3)

# driver.find_element(By.XPATH,'//*[@id="cmbMarca"]').click()
# driver.find_element(By.XPATH, '//*[@id="cmbMarca"]/option[24]').click()

# time.sleep(1)
# driver.find_element(By.XPATH, '//*[@id="txtIni"]').send_keys("17112024")
# time.sleep(1)
# driver.find_element(By.XPATH, '//*[@id="txtFim"]').send_keys("17112024")
# time.sleep(1)

# driver.find_element(By.XPATH, '//*[@id="btnBuscar"]').click()

# time.sleep(5)

# passou = False

# items = []

# ids = []

# for i in driver.find_elements(By.XPATH, '/html/body/div[2]/div[2]/div[2]/div/div/div[2]/div/div/div[1]/div/table/tbody/tr'):
#     ids.append(i)
  
  
# for i in ids:
#     element_id = i.get_attribute("id")
#     if element_id:
#         wait = WebDriverWait(driver, 10) 
#         element = wait.until(EC.element_to_be_clickable((By.ID, element_id)))
#         driver.execute_script(f"tabelaItens('{element_id}', 0)")
#         while True:
#             try:

#                 list = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="tr_concorrente_"]')))

#                 for i in driver.find_elements(By.XPATH, '//*[@id="tr_concorrente_"]'):
#                     imagem = i.find_element(By.XPATH, './td[1]/img').get_attribute("src");
#                     nome = i.find_element(By.XPATH, './td[2]').text;
#                     quantidade = i.find_element(By.XPATH, './td[5]').text;
#                     valor_unitario = i.find_element(By.XPATH, './td[6]').text;
#                     total = i.find_element(By.XPATH, './td[7]').text;
#                     items.append({
#                         "imagem": imagem,
#                         "nome": nome,
#                         "quantidade": quantidade,
#                         "valor_unitario": valor_unitario,
#                         "total": total
#                     })
                    
#                 wait = WebDriverWait(driver, 10) 

#                 try:
#                     next_page_button = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="tblItens"]/ul/li[@class="next page"]/a')))
#                     if next_page_button.is_enabled():
#                         next_page_button.click()
#                         list = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="tr_concorrente_"]')))
#                     else:
#                         break
#                 except:
#                     break
#             except TimeoutException:
#                 print(f"Timeout waiting for element with ID: {element_id}")
#             except JavascriptException as e:
#                 print(f"Javascript error: {e}")
#             except Exception as e:
#                 print(f"An error occurred: {e}")
#     else:
#         print(f"Row with no ID found."  )


# driver.close()


# print(len(items))
# if items:
#     # Process and save the 'items' here
#     try:
#         with open("output.json", "w", encoding="utf-8") as json_file:
#             json.dump(items, json_file, ensure_ascii=False, indent=4)  #Use ensure_ascii=False to handle non-ASCII characters
#         print("Data saved to output.json")
#     except Exception as e:
#         print(f"Error saving JSON data: {e}")
# else:
#     print("No data to save.")

# new_items = []

# with open("output.json", "r", encoding="utf-8") as f:
#     items = json.load(f)

#     for i in tqdm.tqdm(items):
#         response = requests.get(f"https://api.mercadolibre.com/sites/MLB/search?q={i['nome']}")
#         if response.status_code == 200:
#             response = response.json()
#             results = response["results"]
            
#             for result in results:
#                 if result['thumbnail_id'] == i['imagem'].split("D_")[1].split("-I")[0]:
#                     nome = result['title']
#                     listing_type_id = result['listing_type_id']
#                     link = result['permalink']
#                     seller_id = result['seller']['id']
#                     price = result['price']
#                     quantidade = i['quantidade']
#                     valor_unitario = i['valor_unitario']
#                     total = i['total']
#                     new_items.append({"Produto": nome, "tipo": listing_type_id, "link": link, "vendedor_id": seller_id, "Preço Unitário": price, "quantidade": quantidade, "valor_unitario": valor_unitario, "total": total})
                    
                    
# teste = pd.DataFrame(new_items).to_excel("output.xlsx", index=False)


novos_dados = pd.read_excel("output.xlsx", engine='openpyxl')
# novos_dados['Preço Unitário'] = novos_dados['Preço Unitário'].str.replace('.', '', regex=False).str.replace(',', '.', regex=False).astype(float)
# print(novos_dados["Preço Unitário"])
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
    fonte = SelecionarFonte(item).split(",")
    novos_dados.loc[index, 'politica'] = fonte[0]
    novos_dados.loc[index, 'preço_previsto'] = round(float(fonte[1]), 2)
    try:
        print(item['vendedor_id'])
        vendedor_info = requests.get(f"https://api.mercadolibre.com/users/{item['vendedor_id']}").json()
        novos_dados.loc[index, 'lugar'] = vendedor_info.get("address").get("city")
    except:
        novos_dados.loc[index, 'lugar'] = "Não achou"
        print(f"error for: {index}")
all_dados = pd.concat([all_dados, novos_dados], ignore_index=True)
all_dados.to_excel("market-share-historico.xlsx", index=False)
       

