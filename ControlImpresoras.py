import time
import datetime
import pandas as pd
import os
import platform
import subprocess
import wmi
import requests
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.keys import Keys


from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By


import re
import lxml.html as html

#Eligo rutina de búsqueda del contador de acuerdo al modelo
def chequeo_contador(host, modelo):
    if modelo == '3260':
        contador = 326000
    elif modelo == '3320':
        
        website ='http://'+ host +'/sws/index.html'
        #print(website)

        driver = webdriver.Chrome('./chromedriver')
        driver.get(website)
        boton = driver.find_element(By.NAME,'Estado')
        boton.click
        detalle = driver.find_element(By.NAME, 'Contadores de uso')
        detalle.click
        numero_serie = driver.find_element(By.XPATH, '//*[@id="ext-gen339"]')
        contador = driver.find_element(By.XPATH, '//*[@id="ext-gen344"]')
        
        print( host, ' -- es una 3320 = #',numero_serie, ' ->',contador) 


    elif modelo == '3610':
        contador = 361000
    elif modelo == '3615':        
        website ='http://'+ host +'/status/statgeneral.htm'        
        result = requests.get(website)
        contenido = result.text
        soup = BeautifulSoup(contenido,'html.parser')
        contador = soup.find('td',class_= 'std_2').get_text()
        print( host, ' -- es una 3615 = ',contador)      
        
    elif modelo == '6500':
        contador = 65000
    elif modelo == '6505':
        contador = 65050
    elif modelo == '6655':
        contador = 66550
    elif modelo == '7225':
        contador = 722250
    elif modelo == '7845':
        contador = 78450
    elif modelo == '8880':
        XPATH = '/html/body/div/div[3]/div[1]/table[4]/tbody/tr[5]/td/table/tbody/tr[2]/td/table/tbody/tr/td[3]/text()'
        website ='http://'+ host +'/status.htm' 
        r = requests.get(website)
        home = r.content.decode('utf-8')
        parser = html.fromstring(home)
        contador = parser.xpath(XPATH)
    
    elif modelo == '8900':
        contador = 890000
    elif modelo == 'B400':
        contador = 400000
    elif modelo == 'B405':
        contador = 405000
    elif modelo == 'B615':
        contador = 615000
    elif modelo == 'c500':
        contador = 500000
    else:
        contador = 1   

    return contador




print ("Arrancamos con el Proceso")
print('plataforma de ejecuión: ', platform.system().lower())  # Verificar plataforma ejecución

# Carga del archivo de impresoras en una matriz
os.chdir('C:/Users/eettlin.ENERGIAER/Downloads')
os.getcwd()
filename = 'Impresoras.csv'

data = pd.read_csv(filename, sep=';', header=0)
print('Matriz Orignal' , data.shape)
#print (data.head(10))

#Agrego Nuevas Columnas con valores por defecto
salida=data.assign(PING=0)
salida=salida.assign(TTL=0)
salida=salida.assign(FECHA=0)
salida=salida.assign(CONTADOR=0)
print (salida.head(10))

#Rutina de Búseda de Datos x Impresora
print()
print ("Arrancamos con la Iteración")

# Prueba si hay conexión en todas las impresoras
procesado = 0
print('{0} {1:%H:%M:%S} {2}'.format(procesado, datetime.datetime.now(),
                                    "________________________________________"))
# Creo nuevo objeto WMI
c = wmi.WMI()

# Bucle para recorrer el archivo de entrada y verificar ON-LINE
for i in range(len(salida)):
        Impresora = salida.iloc[i]['OBLEA']
        servidorIP = salida.iloc[i]['IP']
        modelo = salida.iloc[i]['MODELO']
        #print(i,'Input Host ',Impresora, ' - ',servidorIP)

        if servidorIP != '0' :
            host = servidorIP                       
        else:
            host = Impresora 
        # Con este comando se dispara el comando PING
        #print(i,'\tLlamamos al ping con parámetro: ',host, 'para la impresora', Impresora)
        x = c.Win32_PingStatus(Address= host)

        for d in x:            
            salida.loc[i,'TTL']=d.ResponseTime
            salida.loc[i,'PING']=d.StatusCode 
            salida.loc[i,'FECHA']=datetime.datetime.now() 
            #salida.loc[i,'FECHA']=chequeo_contador(servidorIP, modelo)        

        procesado += 1
        #print(i,'Host ',host, ' - ',d.ProtocolAddress,' - ', d.StatusCode)    

# Grabo salida de archivo intermedio
salida.to_csv('PING procesado.csv', sep=';', decimal=',')

# Bucle para RECUPERAR LOS CONTADORES
procesado = 0
for i in range(len(salida)):
        Impresora = salida.iloc[i]['OBLEA']
        servidorIP = salida.iloc[i]['IP']
        modelo = salida.iloc[i]['MODELO']
        enlinea = salida.loc[i,'PING'] 
        #print(i,'Input Host ',Impresora, ' - ',servidorIP, ' MODELO ', modelo, 'enlinea', enlinea)

        if enlinea == 0:
            salida.loc[i,'CONTADOR']=chequeo_contador(servidorIP, str(modelo)) 
            salida.loc[i,'FECHA']=datetime.datetime.now() 
            print (Impresora, ' en linea - MODELO:',modelo,' contador =', salida.loc[i,'CONTADOR'])
                   
        procesado += 1
        #print(i,'Host ',host, ' - ',d.ProtocolAddress,' - ', d.StatusCode)    

print ("Fin de Iteración")
print('{0} {1:%H:%M:%S} {2}'.format(procesado, datetime.datetime.now(),"________________________________________"))
print (salida.head(10))

# Grabo salida de archivo
salida.to_csv('CONTADOR procesado.csv', sep=';', decimal=',')
print ("PROGRAMA FINALIZADO !!")