from time import sleep

import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

navegador = webdriver.Chrome()

# 1 Pega cotação do dólar
navegador.get('https://www.google.com.br')
navegador.find_element(
    By.XPATH, '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys('Cotação do Dólar')
navegador.find_element(
    By.XPATH, '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(Keys.ENTER)
COTACAO_DOLAR = navegador.find_element(
    By.XPATH, '//*[@id="knowledge-currency__updatable-data-column"]/div[3]/div/div[2]/input').get_attribute('value')

# 2 Pegar cotação do euro
navegador.get('https://www.google.com.br')
navegador.find_element(
    By.XPATH, '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys('Cotação do Euro')
navegador.find_element(
    By.XPATH, '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(Keys.ENTER)
COTACAO_EURO = navegador.find_element(
    By.XPATH, '//*[@id="knowledge-currency__updatable-data-column"]/div[3]/div/div[2]/input').get_attribute('value')

# 3 Pegar cotação do ouro
navegador.get(r'https://www.melhorcambio.com/ouro-hoje#:~:text=O%20valor%20do%20grama%20do,%C3%A9%20de%20car%C3%A1ter%20exclusivamente%20informativo.')
COTACAO_OURO = navegador.find_element(
    By.ID, 'comercial').get_attribute('value')

navegador.quit()

COTACAO_DOLAR = float(COTACAO_DOLAR.replace(',', '.'))
COTACAO_EURO = float(COTACAO_EURO.replace(',', '.'))
COTACAO_OURO = float(COTACAO_OURO.replace(',', '.'))

# 4 Importar a base de dados e atualizar a base

tabela = pd.read_excel("Produtos.xlsx")

# 5 Recalcular preço
tabela.loc[tabela['Moeda'] == 'Dólar', 'Cotação'] = COTACAO_DOLAR
tabela.loc[tabela['Moeda'] == 'Euro', 'Cotação'] = COTACAO_EURO
tabela.loc[tabela['Moeda'] == 'Ouro', 'Cotação'] = COTACAO_OURO

# Atualizando preço de compra
tabela['Preço de Compra'] = tabela['Cotação'] * tabela['Preço Original']

# Atualizando preço de venda
tabela['Preço de Venda'] = tabela['Preço de Compra'] * tabela['Margem']

# 6 Exportar a base atualizada
tabela.to_excel("Produtos Novo.xlsx", index=False)
print(tabela)
