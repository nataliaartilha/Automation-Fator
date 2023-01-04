import pandas as pd
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
import time as time
import os
from selenium.webdriver.support import expected_conditions as EC


chromedriver = "/Users/nartilha/Downloads/chromedriver_win32.zip/chromedriver"
os.environ["webdriver.chrome.driver"] = chromedriver
driver = webdriver.Chrome(chromedriver)

#s = Service('/home/halovivek/Documents/Automation/selenium_driver/chromedriver.exe')
#driver = webdriver.Chrome(service = s)
#driver.maximize_window()

username = ""
password = ""
#WebDriverWait(browser, 30).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="accountStandalone"]/div/div/div[2]/div/div/div[1]/button'))).click()
#abre a guia do login
#driver = webdriver.Chrome("chromedriver")
driver.get("https://boletador.fator.com.br/Home/Index")
#põe senha e login
driver.find_element(By.XPATH,'/html/body/div[4]/div[2]/div[2]/div/div[2]/form/div/div[1]/input').send_keys(username)
driver.find_element(By.XPATH,'/html/body/div[4]/div[2]/div[2]/div/div[2]/form/div/div[2]/input').send_keys(password)
driver.find_element(By.XPATH,'/html/body/div[4]/div[2]/div[2]/div/div[2]/form/div/div[3]/button').click()
filename="excel_python.xlsx"
os.chdir('G:/depto/RENDA/Natalia Artilha/')
df= pd.read_excel('excel_python.xlsx')
#df = pd.read_excel(filename)
df.set_index('Operação',inplace=True)
df['Qtde']=df['Qtde']*1000
#df
operacao=0
linhas=len(df.index)
for i in df.index:
        while operacao<linhas:
                if df.loc[df.index[operacao],'Indexador']=="CDI":
                        valor=str(df.loc[df.index[operacao],'Qtde'])
                        corretora=df.loc[df.index[operacao], 'CNPJ']
                        prazo1=str(df.loc[df.index[operacao],'DC'])
                        pu=str(df.loc[df.index[operacao],'PU'])
                        desagio=str(round(df.loc[df.index[operacao],'Deságio'],2))
                       
                        #desagio="0.5"
                        #taxa_em=str(df.loc[df.index[operacao],'Taxa Cliente']+0.00001)
                        #taxa_em="110.500"
                        #
                        #print(corretora)
                        driver.get("https://boletador.fator.com.br/BoletaRendaFixaEM/InserirAplicacao/Aplicacao")
                        WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="IDTipoIda_2"]'))).click()
                        #driver.find_element(By.XPATH,'//*[@id="IDTipoIda_2"]').click()
                        time.sleep(5)
                        driver.find_element(By.ID,'s2id_autogen1_search').send_keys(corretora)
                        
                        WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.CLASS_NAME, 'select2-result-label'))).click()
                        #time.sleep(5)
                        #driver.find_element(By.CLASS_NAME,'select2-result-label').click()
                        time.sleep(5)
                        #WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.ID, 'select2-chosen-3'))).click()
                        driver.find_element(By.ID,'select2-chosen-3').click()
                        #Selecione um produto
                        indexador="cdb-desagio"
                        driver.find_element(By.ID,'s2id_autogen3_search').send_keys(indexador)
                        WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.CLASS_NAME, 'select2-result-label'))).click()
                        #driver.find_element(By.CLASS_NAME,'select2-result-label').click()
                        time.sleep(1)
                        indexador2="cdb di - cdb di"
                        WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.ID, 'select2-chosen-60'))).click()
                        #driver.find_element(By.ID,'select2-chosen-60').click()
                        driver.find_element(By.ID,'s2id_autogen60_search').send_keys(indexador2)
                        WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.CLASS_NAME, 'select2-result-label'))).click()
                        #driver.find_element(By.CLASS_NAME,'select2-result-label').click()
                        time.sleep(1)
                        WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.ID, 'VLAplicacao'))).click()
                        #driver.find_element(By.ID,'VLAplicacao').click()
                        driver.find_element(By.ID,'VLAplicacao').send_keys(valor)
                        driver.find_element(By.ID,'VLPrecoUnitDesagio').click()
                        WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="popup-blockUI"]/div[1]/button'))).click()
                        
                        time.sleep(1)
                        #time.sleep(25)
                        #driver.find_element(By.XPATH,'//*[@id="popup-blockUI"]/div[1]/button').click()
                        time.sleep(1)
                        WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.ID, 'INPrazoMatriz'))).click()
                        #driver.find_element(By.ID,'INPrazoMatriz').click()
                        driver.find_element(By.ID,'INPrazoMatriz').clear()
                        driver.find_element(By.ID,'INPrazoMatriz').send_keys(prazo1)
                        WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.ID, 'VLPrecoUnitDesagio'))).click()
                        #driver.find_element(By.ID,'VLPrecoUnitDesagio').click()
                        WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="popup-blockUI"]/div[1]/button'))).click()
                        time.sleep(1)
                        #time.sleep(15)
                        #driver.find_element(By.XPATH,'//*[@id="popup-blockUI"]/div[1]/button').click()
                        #taxa emissão
                        
                        #taxa emissão
                        taxa_em=str(df.loc[df.index[operacao],'Taxa Cliente'])
                        WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.ID, 'DLPercEmissao'))).click()
                        #driver.find_element(By.ID,'DLPercEmissao').click()
                        driver.find_element(By.ID,'DLPercEmissao').send_keys(taxa_em)
                        WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.ID, 'VLPrecoUnitDesagio'))).click()
                        #WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.ID, 'DLPercEmissao'))).click()
                        WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="popup-blockUI"]/div[1]/button'))).click()
                        time.sleep(1)
                        #driver.find_element(By.ID,'DLPercEmissao').send_keys(taxa_em)
                        #WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.ID, 'VLPrecoUnitDesagio'))).click()
                        #driver.find_element(By.ID,'VLPrecoUnitDesagio').click()
                        #WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="popup-blockUI"]/div[1]/button'))).click()
                        #time.sleep(1)
                        #time.sleep(10)
                        #driver.find_element(By.XPATH,'//*[@id="popup-blockUI"]/div[1]/button').click()
                        #desagio
                        #WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.ID, 'VLPrecoUnitDesagio'))).click()
                        #driver.find_element(By.ID,'VLPrecoUnitDesagio').click()
                        #time.sleep(1)
                        WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.ID, 'VLBPDesagio'))).click()
                        driver.find_element(By.ID,'VLBPDesagio').send_keys(desagio)
                        time.sleep(1)
                        #WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.ID, 'VLPrecoUnitDesagio'))).click()
                        driver.find_element(By.ID,'VLPrecoUnitDesagio').click()
                        WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="popup-blockUI"]/div[1]/button'))).click()
                        time.sleep(1)
                        #time.sleep(10)
                        #driver.find_element(By.XPATH,'//*[@id="popup-blockUI"]/div[1]/button').click()
                        #PU
                        WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.ID, 'VLPrecoUnitDesagio'))).click()
                        #driver.find_element(By.ID,'VLPrecoUnitDesagio').click()
                        driver.find_element(By.ID,'VLPrecoUnitDesagio').send_keys(pu)
                        time.sleep(1)
                        #botao salvar 1
                        #
                        driver.execute_script("window.scrollTo(0,document.body.scrollHeight)")
                        time.sleep(1)
                        WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.NAME, 'salvar'))).click()
                        #driver.find_element(By.NAME,'salvar').click()
                        #confirmar
                        driver.execute_script("window.scrollTo(0,document.body.scrollHeight)")
                        time.sleep(1)
                        #driver.find_element(By.NAME,'confirmar').click()
                        WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.NAME, 'confirmar'))).click()
                        #time.sleep(5)
                        #WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="popup-alerta"]/div/div/div[1]/button'))).click()
                        #driver.find_element(By.XPATH,'//*[@id="popup-alerta"]/div/div/div[1]/button').click()
                        #quando passar do horário
                        WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="popup-blockUI"]/div[1]/button'))).click()
                        #operacao=operacao+1

				
                if df.loc[df.index[operacao],'Indexador']=="IPCA":
                        valor=str(df.loc[df.index[operacao],'Qtde'])
                        corretora=df.loc[df.index[operacao], 'CNPJ']
                        prazo1=str(df.loc[df.index[operacao],'DC'])
                        taxa_pre=str(round(df.loc[df.index[operacao],'Taxa Cliente'],2))
                        desagio=str(round(df.loc[df.index[operacao],'Deságio'],2))
                        taxa_em_ipca="100"
                        #valor=str(df.loc[df.index[operacao],'Qtde'])
                        #valor="10000"
                        #corretora="33.775.974/0001-04"
                        #prazo1="365"
                        #taxa_pre="16"
                        #desagio="2"
                        #taxa_em_ipca="100"
                        #print(corretora)
                        driver.get("https://boletador.fator.com.br/BoletaRendaFixaEM/InserirAplicacao/Aplicacao")
                        WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="IDTipoIda_2"]'))).click()
                        #driver.find_element(By.XPATH,'//*[@id="IDTipoIda_2"]').click()
                        time.sleep(5)
                        driver.find_element(By.ID,'s2id_autogen1_search').send_keys(corretora)
                        WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.CLASS_NAME, 'select2-result-label'))).click()
                        #time.sleep(5)
                        #driver.find_element(By.CLASS_NAME,'select2-result-label').click()
                        time.sleep(5)
                        #WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.ID, 'select2-chosen-3'))).click()
                        driver.find_element(By.ID,'select2-chosen-3').click()
                        indexador="cdb-desagio"
                        driver.find_element(By.ID,'s2id_autogen3_search').send_keys(indexador)
                        time.sleep(5)
                        WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.CLASS_NAME, 'select2-result-label'))).click()
                        #driver.find_element(By.CLASS_NAME,'select2-result-label').click()
                        time.sleep(1)
                        #indexador
                        indexador3="ipca_ind"
                        WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.ID, 'select2-chosen-5'))).click()
                        #driver.find_element(By.ID,'select2-chosen-5').click()
                        driver.find_element(By.ID,'s2id_autogen5_search').send_keys(indexador3)
                        WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.CLASS_NAME, 'select2-result-label'))).click()
                        #driver.find_element(By.CLASS_NAME,'select2-result-label').click()
                        #time.sleep(10)

                        #papel
                        WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="select2-chosen-141"]'))).click()
                        #driver.find_element(By.XPATH,'//*[@id="select2-chosen-141"]').click()
                        WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.CLASS_NAME, 'select2-result-label'))).click()
                        #driver.find_element(By.CLASS_NAME,'select2-result-label').click()
                        time.sleep(1)
                        #valor
                        #driver.find_element(By.ID,'VLAplicacao').click()
                        WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.ID, 'VLAplicacao'))).click()
                        driver.find_element(By.ID,'VLAplicacao').send_keys(valor)
                        WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.ID, 'VLPrecoUnitDesagio'))).click()
                        #driver.find_element(By.ID,'INPrazoMatriz').click()
                        time.sleep(1)
                        #time.sleep(25)
                        WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="popup-blockUI"]/div[1]/button'))).click()
                        #driver.find_element(By.XPATH,'//*[@id="popup-blockUI"]/div[1]/button').click()
                        time.sleep(1)
                        WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.ID, 'INPrazoMatriz'))).click()
                        #driver.find_element(By.ID,'INPrazoMatriz').click()
                        driver.find_element(By.ID,'INPrazoMatriz').clear()
                        driver.find_element(By.ID,'INPrazoMatriz').send_keys(prazo1)
                        WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.ID, 'VLPrecoUnitDesagio'))).click()
                        #driver.find_element(By.ID,'DLPercEmissao').click()
                        WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="popup-blockUI"]/div[1]/button'))).click()
                        time.sleep(1)
                        #time.sleep(15)
                        #driver.find_element(By.XPATH,'//*[@id="popup-blockUI"]/div[1]/button').click()
                        #taxa emissão
                        time.sleep(5)
                        taxa_em_ipca="100"
                        WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.ID, 'DLPercEmissao'))).click()
                        #driver.find_element(By.ID,'DLPercEmissao').click()
                        driver.find_element(By.ID,'DLPercEmissao').send_keys(taxa_em_ipca)
                        WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.ID, 'VLPrecoUnitDesagio'))).click()
                        #driver.find_element(By.ID,'VLBPDesagio').click()
                        WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="popup-blockUI"]/div[1]/button'))).click()
                        time.sleep(1)
                        #time.sleep(10)
                        #driver.find_element(By.XPATH,'//*[@id="popup-blockUI"]/div[1]/button').click()
                        #taxa pré
                        WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.ID, 'DLTaxaPre'))).click()
                        #driver.find_element(By.ID,'DLTaxaPre').click()
                        driver.find_element(By.ID,'DLTaxaPre').send_keys(taxa_pre)
                        WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.ID, 'VLPrecoUnitDesagio'))).click()
                        #driver.find_element(By.ID,'VLBPDesagio').click()
                        time.sleep(1)
                        #time.sleep(10)
                        WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="popup-blockUI"]/div[1]/button'))).click()
                        #driver.find_element(By.XPATH,'//*[@id="popup-blockUI"]/div[1]/button').click()
                        #desagio
                        WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.ID, 'VLBPDesagio'))).click()
                        #driver.find_element(By.ID,'VLBPDesagio').click()
                        driver.find_element(By.ID,'VLBPDesagio').send_keys(desagio)
                        WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.ID, 'VLPrecoUnitDesagio'))).click()
                        #driver.find_element(By.ID,'VLPrecoUnitDesagio').click()
                        time.sleep(1)
                        #time.sleep(10)
                        #driver.find_element(By.XPATH,'//*[@id="popup-blockUI"]/div[1]/button').click()
                        WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="popup-blockUI"]/div[1]/button'))).click()
                        #botao salvar 1
                        driver.execute_script("window.scrollTo(0,document.body.scrollHeight)")
                        time.sleep(1)
                        WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.NAME, 'salvar'))).click()
                        #driver.find_element(By.NAME,'salvar').click()
                        #confirmar
                        driver.execute_script("window.scrollTo(0,document.body.scrollHeight)")
                        time.sleep(1)
                        WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.NAME, 'confirmar'))).click()
                        #driver.find_element(By.NAME,'confirmar').click()
                        time.sleep(1)
                        WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="popup-alerta"]/div/div/div[1]/button'))).click()
                        #driver.find_element(By.XPATH,'//*[@id="popup-alerta"]/div/div/div[1]/button').click()
                        #quando passar do horário
                        #WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="popup-blockUI"]/div[1]/button'))).click()
                operacao=operacao+1
                #print(operacao)
else:
        print("Automação executada com sucesso!")
