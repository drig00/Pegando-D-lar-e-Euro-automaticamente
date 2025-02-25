#importando as options para o selenium para fazer com que a página não feche instantaneamente
from selenium.webdriver.chrome.options import Options
chrome_options = Options()
chrome_options.add_experimental_option("detach", True)

#importando o selenium
from selenium import webdriver as opcoes_selenium

#importa o objeto Keys
from selenium.webdriver.common.keys import Keys

#controle do mouse e teclado
import pyautogui as tempoPausaComputador
import pyautogui as teclasAtalho

#para trabalhar com atualizações mais recenter
from selenium.webdriver.common.by import By

meuNavegador = opcoes_selenium.Chrome(options=chrome_options)
meuNavegador.get("https://www.google.com/")
tempoPausaComputador.sleep(4)
meuNavegador.find_element(By.NAME, "q").send_keys("dolar hoje")

#tem o mesmo efeito de clicar na tecla enter no elemento HTML de propriedadde name = 'q'
tempoPausaComputador.sleep(2)
meuNavegador.find_element(By.NAME, "q").send_keys(Keys.RETURN)

tempoPausaComputador.sleep(4)
valorDolar = meuNavegador.find_elements(By.XPATH, '//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]')[0].text
tempoPausaComputador.sleep(4)
meuNavegador.find_element(By.NAME,'q').send_keys("")
tempoPausaComputador.sleep(3)
teclasAtalho.press("tab")
tempoPausaComputador.sleep(3)
teclasAtalho.press("enter")
tempoPausaComputador.sleep(2)
meuNavegador.find_element(By.NAME,'q').send_keys('euro hoje')
tempoPausaComputador.sleep(2)
meuNavegador.find_element(By.NAME,"q").send_keys(Keys.RETURN)
tempoPausaComputador.sleep(2)
valorEuro = meuNavegador.find_element(By.XPATH,"//*[@id='knowledge-currency__updatable-data-column']/div[1]/div[2]/span[1]").text
tempoPausaComputador.sleep(3)

#-------------------------------------------------------------------------------------------------------------------
#Criando uma planilha no excel para salvar os dados que extraimos da internet
import xlsxwriter
import os
#definindo onde o arquivo estará salvo
nomeCaminhoArquivo = r"C:\Users\Vania\PycharmProjects\pythonProject2\dollar-e-euro-google.xlsx"
arquivo = xlsxwriter.Workbook(nomeCaminhoArquivo)

#criando a sheet
planilha = arquivo.add_worksheet()

#escrevendo na sheet
planilha.write("A1", "Dólar")
planilha.write("B1", "Euro")
planilha.write("A2", f"{valorDolar}")
planilha.write("B2", f"{valorEuro}")

while(1):
 print("oi")

#fechando o arquivo
arquivo.close()

#abrindo o arquivo que criamos e colocamos os dados do dólar e euro
os.startfile(nomeCaminhoArquivo)
