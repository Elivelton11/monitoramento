from selenium import webdriver
from selenium.webdriver.common.by import By
import openpyxl 
from datetime import datetime
from time import sleep
#Obtendo data e hora atual
data_atual= datetime.now()

# acessar o site: https://www.casasbahia.com.br/iphone-14-apple/b?origem=autocomplete/b
driver = webdriver.Chrome(executable_path='C:\Users\usuario\Downloads\chromedriver-win64\chromedriver.exe')
driver.get('https://www.casasbahia.com.br/iphone-14-apple/b?origem=autocomplete/b')

sleep(15)

 
# extrair todos os preços 

precos=driver.find_elements(By.XPATH,"//span[@class='product-card__highlight-price']")

# extrair todos os nomes 

produtos=driver.find_elements(By.XPATH,"//h3[@class='product-card__title']")

#extrair o link

links=driver.find_elements(By.XPATH,"//a[@data-testid='product-card-link-overlay']")

sleep(15)

#criar a planilha
workbook=openpyxl.Workbook()
#criando a pagina produtos
workbook.create_sheet('produtos')
#seleciono a pagina produtos
sheet_produtos=workbook['produtos']
#inserir os titulos e preços na planilha
sheet_produtos['A1'].value='Produto'
sheet_produtos['B1'].value='Data atual'
sheet_produtos['C1'].value='Valor'
sheet_produtos['D1'].value='Link'

# converter os dados para a planilha em excel 

for produto, preco,data_atual, link in zip(produtos,precos,links):
    sheet_produtos.append([produto.text,data_atual,preco.text,link.get_attribute('href')])
    
#Salvar a planilha
workbook.save('produtos.xlsx')

#Fechar o navegador

driver.quit()

