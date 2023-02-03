# Scrapy
import scrapy 
from scrapy import Request
from scrapy.http import HtmlResponse

#Selenium Scrapy
from scrapy_selenium import SeleniumRequest
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from time import sleep


#Fechas, dataframes y tiempos
import re
import numpy as np
import pandas as pd
from datetime import datetime
from time import sleep


from cmc.items import CmcItem

# Localidad
import locale
locale.setlocale(locale.LC_TIME, "es_GT") 


# convertidor especifico de este texto de numero romanos
def convert_rome(text):
    text  = text.replace('I','3').replace('V','9').split('-')
    t = str(sum([ int(i)  for i in text[-1]]))
    if len(t)<2:
        return(text[0]+'-'+'0'+t)
    else:
        return(text[0]+'-'+t)


class CMCSpider(scrapy.Spider):
    name = 'cmc'   
    def start_requests(self):
        url = 'https://www.secmca.org/secmcadatos/'
        yield SeleniumRequest(url=url, callback=self.parse, wait_time=2)

    def parse(self, response):
        
        # driver
        driver = response.meta['driver']
        driver.implicitly_wait(5)
        driver.maximize_window()

        #listas
        var = ['Índice de precios al consumidor','Índice subyacente de inflación','Expectativas de inflación','Producto Interno Bruto trimestral','Índice Mensual de la Actividad Económica','RIN del banco central','Tipo de cambio de mercado','Índice tipo de cambio efectivo real','Remesas familiares: ingresos, egresos y neto','Exportaciones FOB', 'Importaciones CIF','Tasas de interés en moneda nacional','Tasas de interés en moneda extranjera', 'Tasa de política monetaria', 'Ingresos totales: corrientes y de capital','Gastos totales: corrientes y de capital','Deuda pública mensual interna y externa y su relación con el PIB']
        opciones = ['Importaciones CIF', 'Exportaciones FOB', 'Remesas familiares: ingresos, egresos y neto', 'Ingresos totales: corrientes y de capital','Gastos totales: corrientes y de capital', 'IMAE general', 'Índice subyacente de inflación', 'PIB en constantes (volumenes encadenados) trimestral', 'IPC general']
        

        # select urls
        urls = [url.get_attribute('href') for url in driver.find_elements(By.XPATH, '//div[@class="col-9 col-md-10"]/ul/li/a')]
        names = [name.text for name in driver.find_elements(By.XPATH, '//div[@class="col-9 col-md-10"]/ul/li/a')]

        for name, posts in zip(names, range(len(urls))):
            if any(v == name for v in var):
                driver.get(urls[posts])    
                if(posts!=len(urls)-1):
                    driver.execute_script("window.open('');")
                    try:                 
                        WebDriverWait(driver,2).until(EC.presence_of_element_located((By.XPATH, "//div[@class='button-box']/button"))).click() # select countries
                        sleep(1)
                        if any(word == WebDriverWait(driver,2).until(EC.presence_of_element_located((By.XPATH, '//div[@id="parameters"]/h2'))).text for word in opciones):
                                WebDriverWait(driver,2).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="params-form"]/div/div[1]/div[2]/div/div[2]/div[1]/label'))).click() #select first variable
                        
                        elif WebDriverWait(driver,2).until(EC.presence_of_element_located((By.XPATH, '//div[@id="parameters"]/h2'))).text =='Tipo de cambio de mercado':
                            WebDriverWait(driver,2).until(EC.presence_of_element_located((By.XPATH, '//*[@id="params-form"]/div/div[1]/div[2]/div/div[2]/div[4]/label'))).click()
                                                                                    
                        elif WebDriverWait(driver,2).until(EC.presence_of_element_located((By.XPATH, '//div[@id="parameters"]/h2'))).text == 'Índice tipo de cambio efectivo real':
                            WebDriverWait(driver,2).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="params-form"]/div/div[1]/div[2]/div/div[2]/div[1]/label'))).click() #select first variable
                            WebDriverWait(driver,2).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="params-form"]/div/div[1]/div[2]/div/div[2]/div[2]/label'))).click() #select second variable

                        elif WebDriverWait(driver,2).until(EC.presence_of_element_located((By.XPATH, '//div[@id="parameters"]/h2'))).text == 'Deuda pública mensual interna y externa y su relación con el PIB':
                            WebDriverWait(driver,2).until(EC.presence_of_element_located((By.XPATH, '//*[@id="params-form"]/div/div[1]/div[2]/div/div[2]/div[5]/label'))).click()

                        elif all(word != WebDriverWait(driver,2).until(EC.presence_of_element_located((By.XPATH, '//div[@id="parameters"]/h2'))).text for word in opciones):
                            WebDriverWait(driver,2).until(EC.presence_of_element_located((By.XPATH, '//*[@id="params-form"]/div/div[1]/div[2]/div/div[3]/button[1]'))).click() #select all variables
                        sleep(2)
                        WebDriverWait(driver,2).until(EC.presence_of_element_located((By.XPATH, "//select[@id='extra-year-first']/option[@value='2017']"))).click() 
                    except:
                        try:
                            driver.refresh()
                            sleep(1)
                            WebDriverWait(driver,2).until(EC.presence_of_element_located((By.XPATH, "//div[@class='button-box']/button"))).click() # select countries
                            sleep(1)
                            if any(word == WebDriverWait(driver,2).until(EC.presence_of_element_located((By.XPATH, '//div[@id="parameters"]/h2'))).text for word in opciones):
                                WebDriverWait(driver,2).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="params-form"]/div/div[1]/div[2]/div/div[2]/div[1]/label'))).click() #select first variable
                                                                                            
                            elif WebDriverWait(driver,2).until(EC.presence_of_element_located((By.XPATH, '//div[@id="parameters"]/h2'))).text == 'Índice tipo de cambio efectivo real':
                                WebDriverWait(driver,2).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="params-form"]/div/div[1]/div[2]/div/div[2]/div[1]/label'))).click() #select first variable
                                WebDriverWait(driver,2).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="params-form"]/div/div[1]/div[2]/div/div[2]/div[2]/label'))).click() #select second variable

                            elif WebDriverWait(driver,2).until(EC.presence_of_element_located((By.XPATH, '//div[@id="parameters"]/h2'))).text == 'Deuda pública mensual interna y externa y su relación con el PIB':
                                WebDriverWait(driver,2).until(EC.presence_of_element_located((By.XPATH, '//*[@id="params-form"]/div/div[1]/div[2]/div/div[2]/div[5]/label'))).click()
                            
                            elif all(word != WebDriverWait(driver,2).until(EC.presence_of_element_located((By.XPATH, '//div[@id="parameters"]/h2'))).text for word in opciones):
                                WebDriverWait(driver,2).until(EC.presence_of_element_located((By.XPATH, '//*[@id="params-form"]/div/div[1]/div[2]/div/div[3]/button[1]'))).click() #select all variables
                            sleep(2)
                            WebDriverWait(driver,2).until(EC.presence_of_element_located((By.XPATH, "//select[@id='extra-year-first']/option[@value='2017']"))).click() 
                        except:
                            try:
                                driver.refresh()
                                sleep(1)
                                WebDriverWait(driver,2).until(EC.presence_of_element_located((By.XPATH, "//div[@class='button-box']/button"))).click() # select countries
                                sleep(1)
                                WebDriverWait(driver,2).until(EC.presence_of_element_located((By.XPATH, "//select[@id='extra-year-first']/option[@value='2017']"))).click() 
                            except:
                                pass
                    sleep(1)
                    # selecting last month or last period
                    try: 
                        i = str(len([i for i in driver.find_elements(By.XPATH,'//select[@id="extra-mouth-last"]/option')])) # last month id
                        WebDriverWait(driver,2).until(EC.presence_of_element_located((By.XPATH, f'//select[@id="extra-mouth-last"]/option[{i}]'))).click() # select last month
                        
                    except:
                        try:
                            i = str(len([i for i in driver.find_elements(By.XPATH,'//select[@id="extra-per-last"]/option')])) # last period id
                            WebDriverWait(driver,2).until(EC.presence_of_element_located((By.XPATH, f'//select[@id="extra-per-last"]/option[{i}]'))).click()  # select last month
                            
                        except:
                            pass
                    sleep(1)
                    # imae variables (only tendency cicle)
                    try:                       
                        WebDriverWait(driver,2).until(EC.presence_of_element_located((By.XPATH, '//*[@id="params-form"]/div/div[1]/div[7]/div/div[2]/div[2]'))).click()
                    except:
                        pass
                    sleep(1)
                    # sending table
                    WebDriverWait(driver,2).until(EC.presence_of_element_located((By.XPATH, "//button[@id='send']"))).click()
                    sleep(2)  
                    
                    # convert to scrapy fields
                    self.body = driver.page_source
                    response = HtmlResponse(url=driver.current_url, body=self.body, encoding='utf-8')
                    
                    ##Extracting data##
                    # titles
                    variables = []
                    for row in response.css('td::text'):
                        if row.extract().strip() and row.extract().strip() not in variables and row.extract().strip() !='': 
                            variables.append(row.extract())
                    paises = [pais.extract() for pais in response.xpath('//th[@class="text-center p-2 test"]/text()')] 
                    titles = [( pais , var) for pais in paises for var in variables]               
                    
                
                    new_titles = []

                    for tup in titles:
                        if tup not in new_titles:
                            new_titles.append(tup)

                    new_titles =  pd.MultiIndex.from_tuples(new_titles)
                    
                    #content
                    c = ' '.join([row.extract().strip() for row in response.xpath('//td/p/text()')])
            
                    contenido = []
                    fechas = []
                    for row in  re.split('([0-9]{4}\-[a-zA-Z]+)\s', c):
                        lst=[]
                        if row:
                            for item in row.split():
                                if item.strip() !='\n':
                                    if re.match('[0-9]{4}\-\D+', item.strip()):
                                        fechas.append(item.strip())
                                    else:
                                        lst.append(item.strip())
                            contenido.append(lst)
                
                    contenido = list(filter(lambda x: x, contenido))        
                    fechas = [datetime.strptime(convert_rome(item.strip()),'%Y-%m').date() if re.match('[0-9]{4}\-[I]{1,3}|[IV]\s+',item) else  datetime.strptime(item.strip().lower().replace('setiembre', 'septiembre'), '%Y-%B').date() for item in fechas]    
                    
                    # to dataframe
                    df = pd.DataFrame(contenido, index = fechas, columns = new_titles).unstack().reset_index()
                    df.fillna(value="--", inplace=True)
                    
                    item = CmcItem()
                    for i in range(df.shape[0]):

                        item['pais'] = df.iloc[i,0]
                        item['variable'] = df.iloc[i,1]
                        item['fecha'] = df.iloc[i,2]
                        if df.iloc[i,3]=='--':
                            valor = np.nan
                        else:
                            valor = float(df.iloc[i,3])
                       
                        item['valor'] = valor 
                        yield item
                    print('exito')
                    
                    chwd = driver.window_handles
                    driver.switch_to.window(chwd[-1])
                    

        # if i like to move to specific handle    
        # chwd = driver.window_handles
        # print(chwd)


            