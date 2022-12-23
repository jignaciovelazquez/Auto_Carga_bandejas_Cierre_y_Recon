from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
from selenium.webdriver import ActionChains

import gspread
from oauth2client.service_account import ServiceAccountCredentials

import pandas as pd

import re

import time


path = "/chromedriver.exe"
Service = Service(executable_path=path)
driver = webdriver.Chrome(service=Service)
# driver.maximize_window()
driver.get("http://jvelazquez:Nacho123-@crm.telecentro.local/MembersLogin.aspx")
time.sleep(1)

driver.find_element(
    by="xpath", value='//*[@id="txtPassword"]').send_keys("Nacho123-")

driver.find_element(
    by="xpath", value='//*[@id="btnAceptar"]').send_keys(Keys.RETURN)

time.sleep(1)


driver.get(
    "http://crm.telecentro.local/Edificio/Gt_Edificio/BandejaEntradaDeRelevamiento.aspx?TituloPantalla=Descarga%20De%20Relevamiento&EstadoGestionId=5&TipoGestionId=3&TipoGestion=OPERACIONES%20DE%20RED%20-%20CIERRE%20DE%20RELEVAMIENTO")

time.sleep(1)

driver.find_element(
    by="xpath", value='//*[@id="btnBuscar"]').send_keys(Keys.RETURN)

time.sleep(3)

filas = len(driver.find_elements(
    by="xpath", value='//*[@id="ctl00_ContentBody_grGestionTecEdificio"]/div/table/tbody/tr'))

columnas = len(driver.find_elements(
    by="xpath", value='//*[@id="ctl00_ContentBody_grGestionTecEdificio"]/div/table/tbody/tr[1]/td'))

filasTotal = filas

datos = []
for x in range(1, filas+1):
    for y in range(1, columnas+1):
        dato = driver.find_element(
            by="xpath", value='//*[@id="ctl00_ContentBody_grGestionTecEdificio"]/div/table/tbody/tr['+str(x)+"]/td["+str(y)+"]").text

        # ------------ separar altura y localidad ---------------------------
        if y == 7:
            direc = driver.find_element(
                by="xpath", value='/html/body/form/div[4]/div[4]/table[1]/tbody/tr[10]/td/table/tbody/tr/td/div/div/div/table/tbody/tr['+str(x)+']/td[7]/a')
            datoAltura = direc.get_attribute('onmouseover')[13:]
            comilla = datoAltura.index("'")
            datoAltura = datoAltura[:comilla]

            guion = datoAltura.index("-")
            Altura = datoAltura[:guion]
            datos.append(Altura)
            newdatoAltura = datoAltura.replace("-", "", 1)
            guion2 = newdatoAltura.index("-")
            corte = len(newdatoAltura)-guion2-2
            localidad = newdatoAltura[-corte:]
            datos.append(localidad)
        else:
            datos.append(dato)

        # ------------ obtener observacion ---------------------------
        if y == 14:
            obs = driver.find_element(
                by="xpath", value='/html/body/form/div[4]/div[4]/table[1]/tbody/tr[10]/td/table/tbody/tr/td/div/div/div/table/tbody/tr['+str(x)+']/td[14]/a')
            cadena = obs.get_attribute('onmouseover')[13:]
            comilla = cadena.index("'")
            cadena = cadena[:comilla]
            datos.append(cadena)

        # ------------ Cambiar el Nodo GPON ---------------------------
        if y == 15:
            nodoGpon = driver.find_element(
                by="xpath", value='/html/body/form/div[4]/div[4]/table[1]/tbody/tr[10]/td/table/tbody/tr/td/div/div/div/table/tbody/tr['+str(x)+']/td[15]').text
            if nodoGpon != " ":
                # si el nodo GPON esta y el nodo HFC esta vacio
                if datos[len(datos) - 13] == " ":
                    datos[len(datos) - 13] = nodoGpon
                # si el nodo GPON esta y el nodo HFC son solo numeros
                if not re.search(r'[a-zA-Z]', datos[len(datos) - 13]):
                    datos[len(datos) - 13] = nodoGpon


driver.get(
    "http://crm.telecentro.local/Edificio/Gt_Edificio/BandejaEntradaDeRelevamiento.aspx?TituloPantalla=CIERRE%20DE%20RELEVAMIENTO&EstadoGestionId=303&TipoGestionId=6&TipoGestion=RECONVERSION%20TECNOLOGICA%20-%20CIERRE%20DE%20RELEVAMIENTO")

time.sleep(5)

driver.find_element(
    by="xpath", value='//*[@id="btnBuscar"]').send_keys(Keys.RETURN)

time.sleep(3)

filas = len(driver.find_elements(
    by="xpath", value='//*[@id="ctl00_ContentBody_grGestionTecEdificio"]/div/table/tbody/tr'))

columnas = len(driver.find_elements(
    by="xpath", value='//*[@id="ctl00_ContentBody_grGestionTecEdificio"]/div/table/tbody/tr[1]/td'))

filasTotal += filas

for x in range(1, filas+1):
    for y in range(1, columnas+1):
        dato = driver.find_element(
            by="xpath", value='//*[@id="ctl00_ContentBody_grGestionTecEdificio"]/div/table/tbody/tr['+str(x)+"]/td["+str(y)+"]").text
        # ------------ separar altura y localidad ---------------------------
        if y == 7:
            direc = driver.find_element(
                by="xpath", value='/html/body/form/div[4]/div[4]/table[1]/tbody/tr[10]/td/table/tbody/tr/td/div/div/div/table/tbody/tr['+str(x)+']/td[7]/a')
            datoAltura = direc.get_attribute('onmouseover')[13:]
            comilla = datoAltura.index("'")
            datoAltura = datoAltura[:comilla]

            guion = datoAltura.index("-")
            Altura = datoAltura[:guion]
            datos.append(Altura)
            newdatoAltura = datoAltura.replace("-", "", 1)
            guion2 = newdatoAltura.index("-")
            corte = len(newdatoAltura)-guion2-2
            localidad = newdatoAltura[-corte:]
            datos.append(localidad)
        else:
            datos.append(dato)

        # ------------ nulo ---------------------------
        if y == 12:
            datos.append(" ")

        # ------------ nulo ---------------------------
        if y == 14:
            obs = driver.find_element(
                by="xpath", value='/html/body/form/div[4]/div[4]/table[1]/tbody/tr[10]/td/table/tbody/tr/td/div/div/div/table/tbody/tr['+str(x)+']/td[14]')
            cadena = obs.get_attribute('title')
            datos[len(datos)-1] = cadena

        # ------------ Cambiar el Nodo GPON ---------------------------
        if y == 15:
            nodoGpon = driver.find_element(
                by="xpath", value='/html/body/form/div[4]/div[4]/table[1]/tbody/tr[10]/td/table/tbody/tr/td/div/div/div/table/tbody/tr['+str(x)+']/td[15]').text
            if nodoGpon != " ":
                datos[len(datos) - 13] = nodoGpon


serie = pd.Series(datos)
df = pd.DataFrame(serie.values.reshape(filasTotal, columnas+2))
df.columns = ["N", "Gestion", "ID", "Nodo", "Zona", "Prioridad", "Direccion", "Localidad", "Subtipo", "Ult Visita",
              "Estado Edificio", "Cant Gestiones", "Usuario", "Contratista", "Bandeja Previa", "Observacion", "Nodo Gpon"]


# print(df)

# ------------------------- Subir a Google Sheet -----------------------------

scope = ['https://www.googleapis.com/auth/spreadsheets',
         "https://www.googleapis.com/auth/drive"]

credenciales = ServiceAccountCredentials.from_json_keyfile_name(
    "carga-de-bandeja-de-entrada-01ec277da545.json", scope)

cliente = gspread.authorize(credenciales)
# ------------- Crea y comparte la Google Sheet  -------------------------
#libro = cliente.create("AutocargaGestiones")
#libro.share("ignaciogproce3@gmail.com", perm_type="user", role="writer")
# ------------------------------------------------------------------------

hoja = cliente.open("AutocargaGestiones").sheet1

#hoja.update_cell(1, 2, 'Bingo!')
#hoja.update("A4", Titulo)

hoja.update([df.columns.values.tolist()] + df.values.tolist())


# ---------------------------------------------------------------------------------


driver.get("https://telecentro.fs.ocs.oraclecloud.com/mobility/")

time.sleep(2)

driver.find_element(
    by="xpath", value='//*[@id="username"]').send_keys("Jvelazquez-ext")

driver.find_element(
    by="xpath", value='//*[@id="password"]').send_keys("123456")

driver.find_element(
    by="xpath", value='/html/body/div/div/div/div/form/div[2]/div[5]/button').send_keys(Keys.RETURN)

time.sleep(15)


driver.find_element(
    by="xpath", value='//*[@id="manage-content"]/div/div[2]/div[2]/div/div[2]/div[3]/div[2]/div[2]/div[1]/div[1]/button[1]').click()

time.sleep(1)

driver.find_element(
    by="xpath", value='//*[@id="manage-content"]/div/div[2]/div[2]/div/div[2]/div[3]/div[2]/div[2]/div[1]/div[2]/div[1]/div[1]').click()


# ----------------------------------- Cuadro Emergente -------------------------------------------------------

time.sleep(1)

driver.find_element(
    by="xpath", value='//button[@title["Vista"]]/span[3]').click()

time.sleep(1)

Select(driver.find_element(by="xpath",
       value='/html/body/div[25]/div/div/div/div[1]/form/div/select[1]')).select_by_value("11")


time.sleep(1)

Select(driver.find_element(by="xpath",
       value='/html/body/div[25]/div/div/div/div[1]/form/div/select[2]')).select_by_value("complete")


time.sleep(1)

driver.find_element(
    by="xpath", value='/html/body/div[25]/div/div/div/div[1]/form/div/label[15]/input').click()

time.sleep(1)

driver.find_element(
    by="xpath", value='/html/body/div[25]/div/div/div/div[2]//button[@title["Aplicar"]]/span[2]').click()


#driver.find_element(by="xpath", value='//*[@id="elId231"]/div[1]/div[3]/controls:toolbar-items:toolbar-switcher/controls:app-menu-button/button').click()

#driver.find_element(by="xpath", value='//*[@id="elId303"]').click()


# -------------------------------------------------------


input("Esperando que no se cierre webdriver: ")


"""
time.sleep(1)

driver.find_element(
    by="xpath", value='//button[@title["Vista de lista"]]/span[1]').click()


"""
