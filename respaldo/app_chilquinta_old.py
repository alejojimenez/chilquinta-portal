import os
import time
import shutil
import rpa as r
import requests

#from domain.chrome_node import ChromeNode

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.alert import Alert
from openpyxl import load_workbook
from datetime import datetime,timedelta
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options

class Scraper_Chilquinta():

    def __init__(self,url, email, password, driver_path):
        print(url, email, password, driver_path)
        self.url = url
        self.email = email
        self.password = password
        self.driver_path = driver_path

    def wait(self, seconds):
        return WebDriverWait(self.driver, seconds)

    def close(self):
        self.driver.close()
        self.driver = None

    def quit(self):
        self.driver.quit()
        self.driver = None
        
    def diccionario(self):
        dic_mes = {'01':'Enero','02':'Febrero',
            '03':'Marzo','04':'Abril',
            '05':'Mayo','06':'Junio',
            '07':'Julio','08':'Agosto',
            '09':'Septiembre','10':'Octubre',
            '11':'Noviembre','12':'Diciembre',
            }
        return dic_mes
        

    def login(self):
        
        #driver_exe = 'C:\\roda\\saesa_portal\\domain\\chromedriver.exe'
        credencials = 'C:\\roda\\chilquinta-portal\\config\\credenciales.xlsx'
        driver_exe = './chromedriver.exe'

        print('Entrando en la funcion login...')
        print('----------------------------------------------------------------------')
        
        #Seteo variables
        email = self.email
        url = self.url
        driver_path = self.driver_path
        password  = self.password

        
        #Version 4.9.1 de selenium
        options = webdriver.ChromeOptions()    
        self.driver = webdriver.Chrome(self.driver_path,options=options)

        self.driver.get(self.url)
        self.driver.maximize_window()
        self.driver.implicitly_wait(40)

        # Controlar evento alert() Notificacion
        try:
            alert = WebDriverWait(self.driver, 35).until(EC.alert_is_present())
            alert = Alert(self.driver)
            alert.accept()
                    
        except:
            print("No se encontró ninguna alerta.")  

        #Boton donde se encuentran todos los servicios
        intentos = 0
        servicios = True
        while (servicios):
                try:
                    print('Try en la funcion servicios...', intentos)
                    print('----------------------------------------------------------------------')
                    intentos += 1
                    boton_servicios = WebDriverWait(self.driver, 20).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, 'button.btn-services span.span-services')))
                    boton_servicios.click()
                    servicios = False
                except:    
                    print('Exception en la funcion servicios...')
                    print('----------------------------------------------------------------------')
                    servicios = intentos <= 3  
            
        #Boton de historial de boletas
        intentos = 0
        historial = True
        while (historial):
                try:
                    print('Try en la funcion historial ...', intentos)
                    print('----------------------------------------------------------------------')
                    intentos += 1
                    boton_menu= self.driver.find_element(By.XPATH,'/html/body/div[1]/div[2]/div/div[2]/div[3]/div/div/div/div/div/div[2]/div[2]/div/div[1]/div/a')
                    time.sleep(2)
                    boton_menu.click()
                    historial = False
                except:    
                    print('Exception en la funcion historial ...')
                    print('----------------------------------------------------------------------')
                    historial = intentos <= 3
            

        #Botones ID y XPATH que utilizaremos para logear
        selector_username_input = 'input-numcli'
        selector_password_input = '/html/body/div[1]/div[2]/div/div[2]/div/div[2]/div/div/div/div/div/fieldset[1]/div/input'
        selector_ingreso_button = '/html/body/div[1]/div[2]/div/div[2]/div/div[2]/div/div/div/div/div/button'
        
        #Primero seteamos el usuario
        intentos = 0
        usuario = True
        while (usuario):
                try:
                    print('Try en el ingreso del usuario ...', intentos)
                    print('----------------------------------------------------------------------')
                    intentos += 1
                    element_username = WebDriverWait(self.driver, 20).until(
                        EC.presence_of_element_located((By.ID, selector_username_input)))
                    time.sleep(5)
                    element_username.clear()
                    element_username.click()
                    element_username.send_keys(email)
                    usuario = False
                except:    
                    print('Exception en el ingreso del usuario  ...')
                    print('----------------------------------------------------------------------')
                    usuario = intentos <= 3
            
        #Luego seteamos la password para ingreso
        intentos = 0
        clave = True
        while (clave):
                try:
                    print('Try en el ingreso de la clave ...', intentos)
                    print('----------------------------------------------------------------------')
                    intentos += 1
                    element_username = WebDriverWait(self.driver, 20).until(
                        EC.presence_of_element_located((By.XPATH, selector_password_input)))
                    time.sleep(5)
                    element_username.clear()
                    element_username.click()
                    element_username.send_keys(password)
                    clave = False
                except:    
                    print('Exception en el ingreso de la clave ...')
                    print('----------------------------------------------------------------------')
                    clave = intentos <= 3
        
        intentos = 0
        ingreso = True
        while (ingreso):
                try:
                    print('Try en el ingreso de la clave ...', intentos)
                    print('----------------------------------------------------------------------')
                    intentos += 1
                    boton_ingreso = WebDriverWait(self.driver, 20).until(
                    EC.element_to_be_clickable((By.XPATH, selector_ingreso_button)))
                    boton_ingreso.click()
                    ingreso = False
                except:    
                    print('Exception en el ingreso de la clave ...')
                    print('----------------------------------------------------------------------')
                    ingreso = intentos <= 3        
        
        print('Ya logramos ingresar, ahora vamos a buscar las factruras')
        time.sleep(7)
        
    def scrapping_chilquita(self):

        #Buscamos la tabla que contiene las boletas
        intentos = 0
        tabla_total= True
        while (tabla_total):
            try:
                print('Try localizando la tabla con los datos ...', intentos)
                print('----------------------------------------------------------------------')
                intentos += 1
                WebDriverWait(self.driver, 30).until(
                EC.presence_of_element_located((By.XPATH,'//*[@id="table-last-invoice-desktop"]/tbody')))
                tabla_total = False
            except:    
                print('Exception en el ingreso de la clave ...')
                print('----------------------------------------------------------------------')
                tabla_total = intentos <= 3    
                    
        #Determinamos el largo de la tabla, que limitara la cantidad de descargas
        intentos = 0
        largo = True
        while (largo):
            try:
                print('Try localizando la tabla con los datos ...', intentos)
                print('----------------------------------------------------------------------')
                intentos += 1
                tabla= self.driver.find_element(By.XPATH,'//*[@id="table-last-invoice-desktop"]/tbody')
                filas = tabla.find_elements(By.TAG_NAME, "tr")
                cantidad_facturas = len(filas)
                largo = False
            except:    
                print('Exception en el ingreso de la clave ...')
                print('----------------------------------------------------------------------')
                largo = intentos <= 3    
             
        i = 1            
        while i < cantidad_facturas+1:
            
            #Obtenemos el numero de factura desde el texto
            intentos = 0
            n_factura = True
            while (n_factura) and intentos <= 5:
                try:
                    print('Try localizando el numero de factura en la tabla ...', intentos)
                    print('----------------------------------------------------------------------')
                    intentos += 1
                    boleta= self.driver.find_element(By.XPATH,f'/html/body/div[1]/div[2]/div/div[2]/div/div[2]/div[2]/div/table/tbody/tr[{i}]/td[1]')
                    valor_boleta = boleta.text
                    print(f'el valor de boleta es {valor_boleta}')
                    n_factura = False
                except:    
                    print('Exception localizando el numero de factura en la tabla...')
                    print('----------------------------------------------------------------------')
                    intentos += 1
                        
            #Obtenemos la fecha y dividimos dicho valor para tener mes y año            
            reintentos = 0            
            fecha_oficial = True
            while (fecha_oficial) and reintentos <= 5:
                try:
                    print('Try localizando fecha en la tabla ...', reintentos)
                    print('----------------------------------------------------------------------')
                    reintentos += 1
                    fecha_completa = self.driver.find_element(By.XPATH,f'/html/body/div[1]/div[2]/div/div[2]/div/div[2]/div[2]/div/table/tbody/tr[{i}]/td[3]')
                    fecha_texto = fecha_completa.text
                    partes = fecha_texto.split("/")
                    mes = str(partes[1])
                    año = str(partes[2])
                    fecha_oficial = False
                except:    
                    print('Exception localizando fecha en la tabla...')
                    print('----------------------------------------------------------------------')
                    reintentos += 1

            #Obtenemos el boton de descarga que haremos click           
            reintentos = 0            
            descarga_oficial = True
            while (descarga_oficial) and reintentos <= 5:
                try:
                    print('Try localizando boton descarga ...', reintentos)
                    print('----------------------------------------------------------------------')
                    reintentos += 1
                    descarga = self.driver.find_element(By.XPATH,f'/html/body/div[1]/div[2]/div/div[2]/div/div[2]/div[2]/div/table/tbody/tr[{i}]/td[6]')
                    descarga.click()
                    boton_confirmar_descarga = WebDriverWait(self.driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, '/html/body/div[4]/div[1]/div/div/div/div/button')))
                    boton_confirmar_descarga.click()
                    time.sleep(7)
                    descarga_oficial = False
                except:    
                    print('Exception en el boton descarga...')
                    print('----------------------------------------------------------------------')
                    reintentos += 1
                    
            # Esperar a que se abra la ventana emergente
            intentos = 0
            reintentar_ventana = True
            while (reintentar_ventana):
                try:
                    # Obtiene el identificador de la ventana actual
                    current_window = self.driver.current_window_handle
                    print('Ventana principal: ', current_window)
                    print('----------------------------------------------------------------------')
                    print('Try en la funcion manejo de ventanas abiertas...', intentos)
                    print('----------------------------------------------------------------------')
                    intentos += 1
                    window_handles_all = self.driver.window_handles
                    reintentar_ventana = False
                            
                except:    
                    print('Exception en la funcion manejo de ventanas abiertas')
                    print('----------------------------------------------------------------------')
                    print('----------------------------------------------------------------------')
                    time.sleep(60) #espera para que cargue ventana emergente          
                    reintentar_ventana = intentos <= 3
                                    
            # Obtiene los identificadores de las ventanas abiertas
            self.driver.implicitly_wait(35)
            print('Ventanas abiertas: ', window_handles_all, len(window_handles_all))
            print('----------------------------------------------------------------------')            
                                    
            # Cambiar al manejo de la ventana emergente
            while True:
                for window_handle in window_handles_all:
                    if window_handle != current_window:
                        self.driver.switch_to.window(window_handle)
                        print('Ventana emergente: ', window_handle)
                        print('----------------------------------------------------------------------')
                        break
                break                

            # Esperar hasta que el elemento esté presente en la página
            self.driver.implicitly_wait(35)
            while True:                
                try:
                    # Obtener la URL de la ventana emergente
                    ventana_emergente_url = self.driver.current_url
                    print("URL de la ventana emergente:", ventana_emergente_url)
                    print('----------------------------------------------------------------------')

                    # Realizar una solicitud GET para obtener la data binaria del documento
                    response = requests.get(ventana_emergente_url, stream=True)

                    #Obtengo resultado del diccionario
                    resultado_diccionario = self.diccionario()
                    
                    # Obtener el nombre del archivo a partir de los datos del proceso de descarga
                    folder_path = './input/'
                    file_name = folder_path + str(valor_boleta)+"_"+ resultado_diccionario[mes]+"_"+año+".pdf" 

                    # Guardar la data binaria en un archivo PDF
                    with open(file_name, 'wb') as file:
                        response.raw.decode_content = True
                        shutil.copyfileobj(response.raw, file)
                        print("Guardando archivo:", file_name)
                        print('----------------------------------------------------------------------')
                    break
                                
                except:    
                    print("No se encontró el elemento con el id especificado...")
                    print('----------------------------------------------------------------------')

                    print('Conteo de documentos: ', i)
                    print('----------------------------------------------------------------------')                

            # Cerrar la ventana emergente
            time.sleep(5)
            self.driver.close()
            time.sleep(15)                            

            # Cambiar de nuevo al manejo de ventana principal
            self.driver.switch_to.window(current_window)
            print('Cual ventana es: ', current_window)
            print('----------------------------------------------------------------------')
                        
            time.sleep(10)
                
            print('pasamos al siguiente archivo')
            i +=1 