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
driver.minimize_window()
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

tabla = driver.find_element(
    by="xpath", value='/html/body/form/div[4]/div[4]/table[1]/tbody/tr[10]/td/table/tbody/tr/td/div/div/div/table/tbody').text


filas = len(driver.find_elements(
    by="xpath", value='//*[@id="ctl00_ContentBody_grGestionTecEdificio"]/div/table/tbody/tr'))

columnas = len(driver.find_elements(
    by="xpath", value='//*[@id="ctl00_ContentBody_grGestionTecEdificio"]/div/table/tbody/tr[1]/td'))

filasTotal = filas

datos = []
for x in range(1, filas+1):
    print("|")
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

            bandeja = driver.find_element(
                by="xpath", value='/html/body/form/div[4]/div[4]/table[1]/tbody/tr[10]/td/table/tbody/tr/td/div/div/div/table/tbody/tr['+str(x)+']/td[14]').text
            print(bandeja)
            if bandeja == "PENDIENTE DE RE...":
                datos[len(datos)-1] = "PENDIENTE DE RELEVAMIENTO"
                print(datos[len(datos)-1])
            if bandeja == "PENDIENTE DE DI...":
                datos[len(datos)-1] = "PENDIENTE DE DISEÑO DE RED"
                print(datos[len(datos)-1])
            if bandeja == "PLANIFICACION D...":
                datos[len(datos)-1] = "PLANIFICACION DE TAREAS"
                print(datos[len(datos)-1])
            if bandeja == "EN CERTIFICACIO...":
                datos[len(datos)-1] = "EN CERTIFICACION"
                print(datos[len(datos)-1])
            if bandeja == "ANALISIS DE FAC...":
                datos[len(datos)-1] = "ANALISIS DE FACTIBILIDAD"
                print(datos[len(datos)-1])

            obs = driver.find_element(
                by="xpath", value='/html/body/form/div[4]/div[4]/table[1]/tbody/tr[10]/td/table/tbody/tr/td/div/div/div/table/tbody/tr['+str(x)+']/td[14]/a')
            cadena = obs.get_attribute('onmouseover')[13:]
            comilla = cadena.index("'")
            cadena = cadena[:comilla]
            datos.append(cadena)
            print("-")

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

# time.sleep(5)
#driver.find_element(by="xpath", value='//*[@id="btnBuscar"]').send_keys(Keys.RETURN)

time.sleep(2)

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

        # ------------ Observacion ---------------------------
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

hoja.clear()

#hoja.update_cell(1, 2, 'Bingo!')

hoja.update([df.columns.values.tolist()] + df.values.tolist())

#hoja.update("A100", tabla)

# ---------------------------------------------------------------------------------


input("Esperando que no se cierre webdriver: ")


"""



"""
