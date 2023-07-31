from codigo.app_chilquinta import Scraper_Chilquinta
#import smtplib
from openpyxl import load_workbook

def send_notification():
    # Código para enviar correo electrónico de notificación
    print('')

if __name__ == '__main__':
    
    print('Obteniendo credenciales...')
    print('----------------------------------------------------------------------')
        
    credencials = 'C:\\roda\\chilquinta-portal\\config\\credenciales.xlsx'
    libro_accesos = load_workbook(credencials)
    hoja_credenciales = libro_accesos['Hoja1']
        
    for j in hoja_credenciales.iter_rows(2):
        try:
            rut = j[0].value
            passw = j[1].value
            web = j[2].value
            break
        except:
            ('no hay credenciales')
            
    email = rut
    password = passw
    url = web
    driver_path = 'chromedriver.exe'
    
    scraper = Scraper_Chilquinta(url, email, password, driver_path)

    #Primer ingreso, año actual
    print('hacemos login')
    scraper.login()
    print('hacemos scrapping')
    scraper.scrapping_chilquita()
    #scraper.upload()
    print('cerramos')
    scraper.close()
