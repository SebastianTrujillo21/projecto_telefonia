import time
import tkinter
from encodings import utf_8
from tkinter import filedialog

import pandas as pd
from openpyxl import load_workbook
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.by import By

encoding: utf_8
#################################### interface #####################################
ventana1 = tkinter.Tk()
ventana1.title("Web Scrapping")
miFrame = tkinter.Frame(ventana1)
miFrame.grid()


def abrirArchivo():
    archivo = filedialog.askopenfilename(initialdir="/", title="Seleccionar el archivo",
                                         filetypes=(("xlsx files", "*.xlsx"), ("all files", "*.*")))
    dirExcel.set(archivo)
    ventana1.destroy()


dirExcel = tkinter.StringVar()
dirExcel.set("Direccion del excel")
direccionLbl = tkinter.Label(miFrame, textvariable=dirExcel, font=("Poppins", 12), fg="green")
direccionLbl.grid(row=0, column=2, padx=10, pady=10)
btnEscoger = tkinter.Button(miFrame, text="Escoger archivo", bg="pink", fg="black", font=("Poppins", 12),
                            command=abrirArchivo)
btnEscoger.grid(row=2, column=2)
ventana1.mainloop()
############################################## EXCEL #################################################################
filename = str(dirExcel.get())
filesheet = pd.read_excel(filename, usecols='A')
wb = load_workbook(filename)
hojas = wb.get_sheet_names()
print(hojas)
tam = list(filesheet.loc[:, 'telefonos'])
print(tam)
numeros = wb.get_sheet_by_name('Hoja1')
wb.close()
##############
contador = 0
cantidad = 0
################ codigo ###################
texto = {"números": [],
         "compania": []}

######################### direccion web ##############################
url = 'https://proxyium.com/'
link = 'https://www.digimobil.es/combina-telefonia-internet?movil=1363'

################# ultimo click para empezar el ingreso de data ######################
iterador = 0

while cantidad < (len(tam)):
    driver = webdriver.Chrome()
    driver.get(url)

    ############# acciones ###############
    driver.find_element(By.CSS_SELECTOR, "body > main > div > div > div:nth-child(2) > div > form > input").send_keys(
        link)

    time.sleep(3)
    driver.find_element(By.CSS_SELECTOR, "body > main > div > div > div:nth-child(2) > div > form > button").click()
    time.sleep(8)
    driver.find_element(By.CSS_SELECTOR,
                        "#infocookies2 > div > div > div.modal-body.text-center > p:nth-child(2) > a").click()

    time.sleep(3)
    try:
        driver.find_element(By.XPATH, "//*[@id='cta_configurador_contratar']").click()
    except NoSuchElementException:
        driver.find_element(By.XPATH,
                            "//*[@id='root']/div[1]/div[3]/div[3]/div[2]/div/div/div[5]/div/div/button").click()
    time.sleep(22)
    contador = 0
    print(contador, "whw1")
    print(cantidad, "wh2")
    for i in range(len(tam)):
        contador = contador + 1
        cantidad = 1 + cantidad
        print(contador, "whw3")
        print(cantidad, "wh4")
        if contador >= 28:
            contador = contador - 1
            driver.back()
            driver.refresh()
            break
        else:
            time.sleep(15)
            try:
                driver.find_element(By.CSS_SELECTOR, "#phoneNumber-0").send_keys(str(tam[iterador]))
            except NoSuchElementException:
                break
            texto["números"].append(tam[iterador])
            time.sleep(5)
            try:
                a = driver.find_element(By.XPATH, "//*[@id='phoneNumber-0-input-group']/div").get_attribute("value")
                texi = driver.find_element(By.XPATH, "//*[@id='phoneNumber-0-input-group']/div").get_attribute("value")

                if a == texi:
                    time.sleep(2)
                    texto["compania"].append('Digi Mobil')
                    time.sleep(4)
                    driver.find_element(By.CSS_SELECTOR, "#phoneNumber-0").clear()

            except:
                time.sleep(4)
                f = "#operator-0"
                tex = driver.find_element(By.XPATH, "//*[@id='operator-0']").get_attribute("value")
                if f == tex:
                    time.sleep(100)
                else:
                    texto["compania"].append(tex)
                    time.sleep(4)
                    driver.find_element(By.CSS_SELECTOR, "#phoneNumber-0").clear()
                    time.sleep(4)
        iterador += 1
        df = pd.DataFrame(texto)
        df.to_excel(filename, sheet_name="Hoja1")
driver.quit()
