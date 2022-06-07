from selenium import webdriver  # cria navegador
from selenium.webdriver.chrome.service import Service

s = Service('c:\webdrivers\chromedriver.exe')
from selenium.webdriver.common.by import By  # localizar elementos (itens de um site)
from selenium.webdriver.common.keys import Keys  # permite clicar teclas no teclado

navegador = webdriver.Chrome(service=s)

#  passo 1: abrir o google
navegador.get("https://www.google.com.br/")

# passo 2: pesquisar a cotação do dolar
navegador.find_element(By.XPATH,
                       '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys("cotação dólar")
navegador.find_element(By.XPATH,
                       '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(Keys.ENTER)

# passo 3: pegar a cotação do dolar
cotacao_dolar = navegador.find_element(By.XPATH,
                                       '//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute('data-value')
print(cotacao_dolar)

# passo 4: pegar cotação do euro
navegador.get("https://www.google.com.br/")
navegador.find_element(By.XPATH,
                       '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys("cotação euro")
navegador.find_element(By.XPATH,
                       '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(Keys.ENTER)
cotacao_euro = navegador.find_element(By.XPATH,
                                      '//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute('data-value')
print(cotacao_euro)

# passo 5: pegar cotação do ouro
navegador.get("https://www.melhorcambio.com/ouro-hoje")
cotacao_ouro = navegador.find_element(By.XPATH, '//*[@id="comercial"]').get_attribute('value')
cotacao_ouro = cotacao_ouro.replace(",", ".")
print(cotacao_ouro)

# passo 6: importar e atualizar a base de dados
import pandas as pd

tabela = pd.read_excel(r"BaseDeDados\Produtos.xlsx")
print(tabela)

# atualizar a cotação de acordo com a moeda correspondente
# tabela.loc[linha, coluna] localizar
# dolar
tabela.loc[tabela["Moeda"] == "Dólar", "Cotação"] = float(cotacao_dolar)

# euro
tabela.loc[tabela["Moeda"] == "Euro", "Cotação"] = float(cotacao_euro)

# ouro
tabela.loc[tabela["Moeda"] == "Ouro", "Cotação"] = float(cotacao_ouro)

# atualizar o preço de compra = preço original * cotação
tabela["Preço de Compra"] = tabela["Preço Original"] * tabela["Cotação"]

# atualizar o preço de venda = preço de compra * margem
tabela["Preço de Venda"] = tabela["Preço de Compra"] * tabela["Margem"]

print(tabela)

# passo 7: exportar tabela
tabela.to_excel("Produtos_Atualizado.xlsx", index=False)
navegador.quit()
