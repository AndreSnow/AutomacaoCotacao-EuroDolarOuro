from selenium import webdriver
from selenium.webdriver.chrome.webdriver import WebDriver
from selenium.webdriver.common.keys import Keys
import pandas as pd


#1 - entrar na net
#chrome
navegador = webdriver.Chrome()


#2 - pegar cotação do euro
navegador.get(
    'https://www.google.com/search?q=cota%C3%A7%C3%A3o+euro&ei=uw4cYZXeK5zN1sQPx9SoqAw&oq=cota%C3%A7%C3%A3o+euro&gs_lcp=Cgdnd3Mtd2l6EAMyEAgAEIAEELEDEIMBEEYQggIyCwgAEIAEELEDEIMBMgUIABCABDILCAAQgAQQsQMQgwEyCwgAEIAEELEDEIMBMgUIABCABDIFCAAQgAQyBQgAEIAEMgUIABCABDIFCAAQgAQ6BwgAEEcQsAM6BwgAELADEEM6DQgAEIAEELEDEIMBEApKBAhBGABQiApYngxgiBRoAnACeACAAZMBiAGoBJIBAzIuM5gBAKABAcgBCsABAQ&sclient=gws-wiz&ved=0ahUKEwjVp7_-5bjyAhWcppUCHUcqCsUQ4dUDCA4&uact=5')
cotacao_euro = navegador.find_element_by_xpath(
    '//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute('data-value')
print(cotacao_euro)

#3 - pegar cotação do dolar
navegador.get(
    'https://www.google.com/search?q=cota%C3%A7%C3%A3o+dolar&oq=cotacao&aqs=chrome.1.69i57j0i512l2j0i10i433i512j0i512j0i3j0i512l2j0i10i131i433j0i512.2305j1j7&sourceid=chrome&ie=UTF-8')
cotacao_dolar = navegador.find_element_by_xpath(
    '//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute('data-value')
print(cotacao_dolar)

#4 - pegar cotação do ouro
navegador.get(
    'https://www.melhorcambio.com/ouro-hoje')
cotacao_ouro = navegador.find_element_by_xpath(
    '//*[@id="comercial"]').get_attribute('value')
cotacao_ouro = cotacao_ouro.replace(',','.')
print(cotacao_ouro)


navegador.quit()
#atualizar e importar a base de dados

tabela = pd.read_excel('../Automação web(cotações euro, dolar, ouro)/Produtos.xlsx')
#atualizar cotação
#tabela.loc[linha, coluna] = float(cotacao_dolar)
tabela.loc[tabela['Moeda'] == 'Dólar', 'Cotação'] = float(cotacao_dolar)
tabela.loc[tabela['Moeda'] == 'Euro', 'Cotação'] = float(cotacao_euro)
tabela.loc[tabela['Moeda'] == 'Ouro', 'Cotação'] = float(cotacao_ouro)

#atualizar preço de compra
tabela['Preço Base Reais'] = tabela['Preço Base Original'] * tabela['Cotação']

#atualizar preco de venda
tabela['Preço Final'] = tabela['Preço Base Original'] * tabela['Margem']

#tratamento de dados
tabela['Preço Final'] = tabela['Preço Final'].map('{:.2f}'.format)
tabela['Preço Base Original'] = tabela['Preço Base Original'].map('{:.2f}'.format)
tabela['Cotação'] = tabela['Cotação'].map('{:.2f}'.format)

#6 - exportar a base atualizada
tabela.to_excel('Produtos Novo.xlsx', index = False)