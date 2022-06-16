from glob import glob
from lib2to3.pgen2 import driver
from numpy import size
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
import time
from openpyxl import load_workbook
from datetime import date
from selenium.webdriver.support.ui import Select
import os
#from sympy import matrix_multiply_elementwise
import speech_recognition as sr
from selenium import webdriver
import ffmpy
import requests
import urllib
import pydub
import win32com.client as win32
import random
from array import *
from sklearn.feature_extraction.text import CountVectorizer
from selenium.webdriver import ActionChains
import json
import pyodbc
D_CHR = None

ChromeDriver = r"C:\drivers_selenium\chromedriver_102.exe"
EdgeDriver = r"C:\drivers_selenium\msedgedriver.exe"
URL_Destino = "https://www.linkedin.com/uas/login?session_redirect=https%3A%2F%2Fwww%2Elinkedin%2Ecom%2Fsearch%2Fresults%2Fpeople%2F%3Fkeywords%3DTechnology%2F&fromSignIn=true&trk=cold_join_sign_in"
URL_Gmail = "https://accounts.google.com/AccountChooser/signinchooser?service=mail&continue=https%3A%2F%2Fmail.google.com%2Fmail%2F&flowName=GlifWebSignIn&flowEntry=AccountChooser"
URL_Elemplo = "https://www.elempleo.com/co/empresas/Home"
Excel_Ruta = r"C:\Users\Joseph Rojas\Exsis Software y Soluciones\Automation - Documentos\Solicitudes1.xlsx"
ArchivoExcel = r"C:\Users\Joseph Rojas\Exsis Software y Soluciones\Automation - Documentos\Solicitudes1.xlsx"

boton_login_empleo = "html/body/div[7]/div[1]/div[1]/fieldset/form/div[3]/button"
campocargo = "html/body/div[2]/header/nav/div/div[2]/form/div/div[1]/div/span[1]/input[2]"
campociudad = "html/body/div[2]/header/nav/div/div[2]/form/div/div[2]/span/div/span/input[2]"
campoIdioma = "html/body/div[2]/div[2]/section[2]/div/div[3]/div[3]/div/div/form/div/div[9]/div/div/div/span/input[2]"
botonIdioma = "html/body/div[2]/div[2]/section[2]/div/div[3]/div[3]/div/div/form/div/div[9]/div/div/button/div/i"
buscar_boton = "html/body/div[2]/header/nav/div/div[2]/form/div/div[3]/button"
# Formación básica
campoformbas = "html/body/div[2]/div[2]/section[2]/div/div[3]/div[3]/div/div/form/div/div[3]/div/div/div/div/span/input[2]"

popup_boton = "html/body/div[10]/div/div/div[3]/div/div[2]/a"
nombre_popup = "html/body/div[2]/div[2]/section[2]/div/div[3]/div[3]/div[2]/div/div/div[1]/div/div[2]/div[2]/div/div[2]/div/h3/a"
nombre_popup2 = "html/body/div[2]/div[2]/section[2]/div/div[3]/div[3]/div[2]/div/div/div[2]/div/div[2]/div[2]/div/div[2]/div/h3"
siguiente_i_boton = "html/body/div[19]/div/div/div[4]/a/i"
siguiente_boton = "html/body/div[19]/div/div/div[4]/a[2]/i"
volver_boton = "html/body/div[19]/div/div/div[5]/div/div[2]/a"

Cand_nombre = "html/body/div[19]/div/div/div[5]/div/div/div/div/div[2]/div[3]/div[2]/strong"
Cand_prof = "html/body/div[19]/div/div/div[5]/div/div[2]/div[2]/div/div/div[2]/div/div/strong"
Cand_exp = "html/body/div[19]/div/div/div[5]/div/div[2]/div[2]/div/div/div[2]/div/div[2]/strong"
Cand_asp = "html/body/div[19]/div/div/div[5]/div/div[2]/div[2]/div/div/div[2]/div/div[3]/strong"
Cand_aux = "html/body/div[19]/div/div/div[5]/div/div/div/div/div[2]/div[3]/div[3]/strong"
Cand_ubi = "html/body/div[19]/div/div/div[5]/div/div[2]/div[2]/div/div/div[2]/div[3]/div[2]/small"
Cand_idio = "html/body/div[19]/div/div/div[5]/div/div[2]/div[2]/div/div/div[2]/div[4]/div/div/span[2]"
Cand_idigen = "html/body/div[19]/div/div/div[5]/div/div[2]/div[2]/div/div/div[2]/div[4]/div/div"

contactar_boton = "html/body/div[19]/div[1]/div/div[5]/div/div[1]/div/div/div[3]/ul/li[3]/a"
Cand_wassa = "html/body/div[19]/div[1]/div/div[5]/div/div[1]/div/div/div[3]/div[2]/div[3]/div[3]/a"
Cand_telefono = "html/body/div[19]/div[1]/div/div[5]/div/div[1]/div/div/div[3]/div[2]/div[3]/div[4]/div/div[2]/div"
Aceptar_boton = "html/body/div[10]/div/div/div[3]/div[1]/div[2]/a"
# Cand_correo1 = ["html/body/div[19]/div[1]/div/div[5]/div/div[1]/div/div/div[3]/div[2]/div[3]/div[4]/div[4]/div[2]/span/a",
#               "html/body/div[19]/div[1]/div/div[5]/div/div[1]/div/div/div[3]/div[2]/div[3]/div[3]/div[3]/div[2]/span/a",
#              "html/body/div[19]/div[1]/div/div[5]/div/div[1]/div/div/div[3]/div[2]/div[3]/div[4]/div[3]/div[2]/span/a",
#               "html/body/div[19]/div[1]/div/div[5]/div/div[1]/div/div/div[3]/div[2]/div[3]/div[3]/div[3]/div[2]/span/a"
#               ]
Cand_correo1 = "html/body/div[19]/div[1]/div/div[5]/div/div[1]/div/div/div[3]/div[2]/div[3]/div[4]/div[4]/div[2]/span/a"

cantidad_hv = 'html/body/div[2]/div[2]/section[2]/div[1]/div[3]/div[3]/div[2]/div/div[2]/div/div/select'
Cand_desc = "html/body/div[19]/div/div/div[5]/div/div[2]/div[2]/div/div/div[2]/div[2]/div/div[2]/p"
cantidad_busquedas = "/html/body/div[2]/div[2]/section[2]/div[1]/div[3]/div[1]/div[1]/h3/span[1]"
Filtro_Fecha = "html/body/div[2]/div[2]/section[2]/div[1]/div[3]/div[1]/div[2]/div/span[1]/span[1]/span"
cantidad_link = "/html/body/div[6]/div[3]/div[2]/div/div[1]/main/div/div/h2"
reciente = "/html/body/div[2]/div[2]/section[2]/div[1]/div[3]/div[1]/div[2]/div/span[1]/span[1]/span"
exp_min = "html/body/div[2]/div[2]/section[2]/div[1]/div[3]/div[3]/div[1]/div/form/div[1]/div[6]/div/div/div/div/span[1]"
exp_max = "/html/body/div[2]/div[2]/section[2]/div[1]/div[3]/div[3]/div[1]/div/form/div[1]/div[6]/div/div/div/div/span[2]"


def validacionVaca():
    val = False
    global D_CHR
    t = "html/body/div[2]/div[2]/section[2]/div[1]/div[3]/div[3]/div[2]/div/div[1]/div/div/a"
    try:
        msg = D_CHR.find_element_by_xpath(t).text
        if (msg != "No hemos encontrado resultados"):
            val = True
    except:
        val = False
    return val


def Proc_NPL(detye, conjuntoDatos):
    roles = conjuntoDatos
    entrada = [detye]
    rolesE = CountVectorizer(binary=True, ngram_range=(1, 1), analyzer='word')
    vector_rolesE = rolesE.fit_transform(roles)
    a = rolesE.get_feature_names()
    entradaE = CountVectorizer(
        binary=True, ngram_range=(1, 1), analyzer='word')
    vector_entradaE = entradaE.fit_transform(entrada)
    b = entradaE.get_feature_names()
    aa = list(a)
    bb = list(b)
    resultante = set(aa).intersection(set(bb))
    resultanteE = CountVectorizer(
        binary=True, ngram_range=(1, 1), analyzer='word')
    vector_normalizado = resultanteE.fit_transform(resultante)
    trak = (" ".join(map(str, resultante)))
    print("resultante:  ", resultante)
    lista = list(resultante)
    print(trak)
    return trak, lista


def write_json(new_data, filename='Data.json'):
    with open(filename, 'w') as file:
        json.dump(new_data, file, indent=4)


def JsonAddMas(alte):

    df = pd.read_json('Data.json')
    tam = size(df)
    d = []
    for i in range(tam):
        a = df["roles"][i]["cargo"]
        d.append(a)
    # print(d)
    rta = Proc_NPL(detye=alte, conjuntoDatos=d)
    print(rta)

    with open("Data.json") as json_file:
        dre = json.load(json_file)
        temp = dre["roles"]
        y = {"inp": alte,
             "cargo": rta[0]
             }
        temp.append(y)

    write_json(dre)
    return rta[1]


def EnviarEmail(sub, destino, archivo="", msg='Prueba de funcionamiento correcta'):
    olApp = win32.Dispatch('Outlook.Application')
    olNS = olApp.GetNameSpace('MAPI')
    mailItem = olApp.CreateItem(0)
    mailItem.Subject = sub
    mailItem.BodyFormat = 1
    mailItem.Body = msg
    mailItem.To = destino
    #mailItem.CC = 'nicolas.mendoza@exsis.com.co ; joseph.rojas@exsis.com.co'
    mailItem._oleobj_.Invoke(
        *(64209, 0, 8, 0, olNS.Accounts.Item("reclutamiento@exsis.com.co")))
    if archivo != "":
        mailItem.Attachments.Add(os.path.join(os.getcwd(), archivo))

    mailItem.Display()
    mailItem.Save()
    mailItem.Send()


def launchBrowser():
    chr_options = Options()
    # El explorador queda abierto al finalizar la preuba
    chr_options.add_experimental_option("useAutomationExtension", False)
    
    chr_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    chr_options.add_experimental_option("detach", True)
    driverC = webdriver.Chrome(
        options=chr_options, executable_path=ChromeDriver)
    driverC.maximize_window()
    
    driverC = webdriver.Chrome(options=chr_options, executable_path=ChromeDriver)
    driverC.maximize_window()
    return driverC


def ExcelData():
    vectorsalida = []
    requer = pd.ExcelFile(ArchivoExcel)
    genData = requer.parse("Hoja1")
    c = genData['Cargo'][0]  # 0
    n = genData['Nivel formación basica'][0]  # 1
    fb = genData['Formación básica'][0]  # 2
    vfc = genData['Val_Form_Comp'][0]  # 3
    fc1 = genData['Formación Complementaria 1'][0]  # 4
    afc1 = genData['Area FormComp1'][0]  # 5
    vfc2 = genData['Val_Form_Comp_2'][0]  # 6
    fc2 = genData['Formación Complementaria 2'][0]  # 7
    afc2 = genData['Area FormComp2'][0]  # 8
    bil = genData['Bilingüe'][0]  # 9
    lvlI = genData['NivelIngles'][0]  # 10
    lvlP = genData['NivelPort'][0]  # 11
    lvlF = genData['NivelFran'][0]  # 12
    yr = genData['Años_EXP'][0]  # 13
    val = genData['Valor'][0]  # 14
    iss = genData['Task'][0]  # 15
    ci_aux = genData['Ciudad'][0]
    ciu = ci_aux.split(",")[0]  # 16
    tch1 = genData['Tecno'][0]  # 17
    tch2 = genData['Tecno2'][0]  # 18

    vectorsalida = [c, n, fb, vfc, fc1, afc1, vfc2, fc2,
                    afc2, bil, lvlI, lvlP, lvlF, yr, val, iss, ciu, tch1]
    return vectorsalida


def format_tbl(writer, sheet_name, df):
    outcols = df.columns
    if len(outcols) > 25:
        raise ValueError('table width out of range for current logic')
    tbl_hdr = [{'header': c} for c in outcols]
    bottom_num = len(df)+1
    right_letter = chr(65-1+len(outcols))
    tbl_corner = right_letter + str(bottom_num)

    worksheet = writer.sheets[sheet_name]
    worksheet.add_table('A1:' + tbl_corner,  {'columns': tbl_hdr})


def FiltroIdioma(vali, nI, nF, nP):
    lvls = ["A1 - Básico", "A2 - Pre intermedio", "B1 - Intermedio",
            "B2 - Intermedio alto", "C1 - Avanzado", "C2 - Excelente"]
    Idio = []

    if (vali == "Si"):

        for i in range(len(lvls)):

            if (nI == lvls[i]):
                Idio.append("Inglés")

            if nP == lvls[i]:
                Idio.append("Portugués")

            if nF == lvls[i]:
                Idio.append("Francés")

    else:
        Idio = []
    return Idio


def LoginLink(corr, contra):
    global D_CHR
    D_CHR.get("https://www.linkedin.com/uas/login?session_redirect=https%3A%2F%2Fwww%2Elinkedin%2Ecom%2Fsearch%2Fresults%2Fpeople%2F%3FgeoUrn%3D%255B%2522100876405%2522%255D%26keywords%3Dtecnologia%26origin%3DGLOBAL_SEARCH_HEADER%26sid%3Dm_f&fromSignIn=true&trk=cold_join_sign_in")
    usLink = D_CHR.find_element_by_id("username")
    usLink.send_keys(corr)
    psLink = D_CHR.find_element_by_id("password")
    psLink.send_keys(contra)
    psLink.send_keys(Keys.ENTER)


def busqLink(cargo):
    D_CHR.find_element_by_xpath(
        "//div[@id='global-nav-typeahead']//input").clear()
    busc = D_CHR.find_element_by_xpath(
        "//div[@id='global-nav-typeahead']//input")
    busc.send_keys(cargo)
    busc.send_keys(Keys.ENTER)

    infocand = []
    datex = []
    time.sleep(10)
    ser = D_CHR.find_element_by_xpath(cantidad_link).text
    ser = ser.split()[1].replace(".", "")
    for i in range(1, 10):
        dir_nam = "html/body/div[6]/div[3]/div[2]/div/div/main/div/div/div/ul/li["+str(
            i)+"]/div/div/div[2]/div/div/div/span/span/a/span/span"
        dir_ocu = "html/body/div[6]/div[3]/div[2]/div/div/main/div/div/div/ul/li["+str(
            i)+"]/div/div/div[2]/div/div[2]/div/div"
        dir_ubi = "html/body/div[6]/div[3]/div[2]/div/div/main/div/div/div/ul/li["+str(
            i)+"]/div/div/div[2]/div/div[2]/div/div[2]"
        dir_link = "html/body/div[6]/div[3]/div[2]/div/div/main/div/div/div/ul/li["+str(
            i)+"]/div/div/div[2]/div/div/div/span/span/a"

        try:
            nomb = D_CHR.find_element_by_xpath(dir_nam).text
        except:
            nomb = ""
        try:
            occp = D_CHR.find_element_by_xpath(dir_ocu).text
        except:
            occp = ""
        try:
            ubic = D_CHR.find_element_by_xpath(dir_ubi).text
        except:
            ubic = ""
        try:
            zled = D_CHR.find_element_by_xpath(dir_link).get_attribute("href")
        except:
            zled = ""

        infocand = [nomb, occp, ubic, zled]

        datex.append(infocand)
    return [datex, ser]


def delay():
    time.sleep(random.randint(2, 3))


def LoginELEM(corr, contra):
    global D_CHR
    D_CHR = launchBrowser()
    D_CHR.get(URL_Elemplo)  # Registrarse en cuenta de google
    user_empleo = D_CHR.find_element_by_name("EmailField")
    user_empleo.send_keys(corr)
    time.sleep(2)
    pass_empleo = D_CHR.find_element_by_name("PasswordField")
    pass_empleo.send_keys(contra)
    pass_empleo.send_keys(Keys.ENTER)
    time.sleep(10)
    print("OK")
    try:
        D_CHR.switch_to.default_content()
        frames = D_CHR.find_element_by_xpath("html/body/div[13]/div[2]/iframe")
        # print(frames)
        D_CHR.switch_to.frame(frames)
        delay()
        D_CHR.find_element_by_id("recaptcha-audio-button").click()
        # D_CHR.find_element_by_xpath("/html/body/div/div/div[3]/div[2]/div/div/div[2]/button").click()
        print("OK 2")
        time.sleep(3)
        try:
            D_CHR.find_element_by_xpath(
                "/html/body/div/div/div[3]/div/button").click()
        except:
            D_CHR.find_element_by_xpath(
                "/html/body/div/div/div[3]/div/div/div//div[2]/button").click()

        src = D_CHR.find_element_by_id("audio-source").get_attribute("src")
        print("[INFO] Audio src: %s" % src)

        urllib.request.urlretrieve(src, os.getcwd()+"\\sample.mp3")
        time.sleep(2)
        sound = pydub.AudioSegment.from_mp3(os.getcwd()+"\\sample.mp3")
        time.sleep(2)
        sound.export(os.getcwd()+"/sample.wav", format="wav")
        time.sleep(2)
        sample_audio = sr.AudioFile(os.getcwd()+"\\sample.wav")
        time.sleep(2)
        r = sr.Recognizer()
        time.sleep(2)
        with sample_audio as source:
            time.sleep(2)
            audio = r.record(source)
        time.sleep(2)
        karte = r.recognize_google(audio)
        print("[INFO] Recaptcha code:%s" % karte)
        time.sleep(5)
        tr = D_CHR.find_element_by_id("audio-response")
        tr.send_keys(karte)
        time.sleep(3)
        tr.send_keys(Keys.ENTER)
        time.sleep(6)
        try:
            time.sleep(4)
            D_CHR.find_element_by_xpath(Aceptar_boton).click()
        except:
            print("OK")
    except:
        print("No hay captcha")


def filtrosElem(f1, f2, f3, f4, f5, f6, f7, f8, f9):
    global D_CHR
    print("ASE")
    #pal_cla = Proc_NPL(f1)
    cargo_field = D_CHR.find_element_by_xpath(campocargo)
    cargo_field.send_keys(f1)  # Features[0]
    ciudad_field = D_CHR.find_element_by_xpath(campociudad)
    ciudad_field.send_keys(f2)  # Features[16]
    D_CHR.find_element_by_xpath(buscar_boton).click()
    time.sleep(30)

    time.sleep(5)
    ased = D_CHR.find_element_by_xpath(reciente)
    ased.click()
    for i in range(2):
        time.sleep(5)
        ased.send_keys(Keys.ARROW_DOWN)
    ased.send_keys(Keys.ENTER)

    a = validacionVaca()
    el1 = D_CHR.find_element_by_xpath(exp_min)
    el2 = D_CHR.find_element_by_xpath(exp_max)

    print("AJUA", a)
    if (a != True):
        print("RRAA")
        time.sleep(5)
        f3.append(f8)  # Busqueda por tecnología
        formbas_field = D_CHR.find_element_by_xpath(campoformbas)
        for i in range(size(f3)):
            time.sleep(5)
            formbas_field.send_keys(f3[i])  # Features[2]
            time.sleep(3)
            formbas_field.send_keys(Keys.ENTER)

        #formbas_field = D_CHR.fin+d_element_by_xpath(campoformbas)
        # formbas_field.send_keys(pal_cla)  # Features[2]
        time.sleep(3)
        # formbas_field.send_keys(Keys.ENTER)

        time.sleep(5)
        # Features[9] Features[10] Features[11] Features[12]
        ter = FiltroIdioma(f4, f5, f6, f7)

        if (len(ter) >= 1):
            for i in range(len(ter)):
                idioma_field = D_CHR.find_element_by_xpath(campoIdioma)
                time.sleep(2)
                idioma_field.send_keys(ter[i])
                time.sleep(4)
                idioma_field.send_keys(Keys.ARROW_DOWN)
                time.sleep(2)
                idioma_field.send_keys(Keys.ENTER)
                time.sleep(2)
                D_CHR.find_element_by_xpath(botonIdioma).click()
                time.sleep(10)

    yrs = int(f9)
    if(yrs >= 0 and yrs < 1):
        ActionChains(D_CHR).drag_and_drop_by_offset(el1, 0, 0).perform()
        time.sleep(10)
        ActionChains(D_CHR).drag_and_drop_by_offset(el2, -240, 0).perform()
        time.sleep(5)
    elif(yrs >= 1 and yrs < 3):
        ActionChains(D_CHR).drag_and_drop_by_offset(el1, 60, 0).perform()
        time.sleep(10)
        ActionChains(D_CHR).drag_and_drop_by_offset(el2, -180, 0).perform()
        time.sleep(5)
    elif(yrs >= 3 and yrs < 5):
        ActionChains(D_CHR).drag_and_drop_by_offset(el1, 120, 0).perform()
        time.sleep(10)
        ActionChains(D_CHR).drag_and_drop_by_offset(el2, -120, 0).perform()
        time.sleep(5)
    elif(yrs >= 5 and yrs < 10):
        ActionChains(D_CHR).drag_and_drop_by_offset(el1, 180, 0).perform()
        time.sleep(10)
        ActionChains(D_CHR).drag_and_drop_by_offset(el2, -60, 0).perform()
        time.sleep(5)
    elif(yrs >= 10):
        ActionChains(D_CHR).drag_and_drop_by_offset(el1, 240, 0).perform()
        time.sleep(10)
        ActionChains(D_CHR).drag_and_drop_by_offset(el2, 0, 0).perform()
        time.sleep(5)


def DatosElem():
    global D_CHR
    time.sleep(5)
    try:
        nombre = D_CHR.find_element_by_xpath(Cand_nombre).text
    except:
        time.sleep(30)
        nombre = D_CHR.find_element_by_xpath(Cand_nombre).text
    prof = D_CHR.find_element_by_xpath(Cand_prof).text
    exp = D_CHR.find_element_by_xpath(Cand_exp).text.split()[0]
    asp = D_CHR.find_element_by_xpath(Cand_asp).text
    aux = D_CHR.find_element_by_xpath(Cand_aux).text.split("-")
    lastpos = aux[0]
    #eda = (aux[1].split("|")[1]).split()[0]
    ubi = D_CHR.find_element_by_xpath(Cand_ubi).text
    try:
        idioma1 = D_CHR.find_element_by_xpath(Cand_idio).text.split()[1]
    except:
        idioma1 = ""
    idigen = D_CHR.find_element_by_xpath(Cand_idigen).text
    time.sleep(2)

    try:
        D_CHR.find_element_by_xpath(contactar_boton).click()
    except:
        time.sleep(30)
        D_CHR.find_element_by_xpath(contactar_boton).click()
    time.sleep(10)

    try:
        cor = D_CHR.find_element_by_xpath(Cand_correo1).text
    except:
        cor = ""

    try:
        tel = D_CHR.find_element_by_xpath(Cand_telefono).text
        tel = tel.split(":")[1]
    except:
        tel = ""

    try:
        wts = D_CHR.find_element_by_xpath(Cand_wassa).get_attribute("href")
    except:
        wts = ""

    desc = D_CHR.find_element_by_xpath(Cand_desc).text
    prom = asp.replace("$","")
    prom = prom.replace(",",".")
    prom = (float(prom.split()[0])+float(prom.split()[2]))/2

    r = [nombre, prof, exp, asp, prom,lastpos, ubi,
         idioma1, idigen, cor, tel, wts, desc]

    return r

# 123,456


def ExtraccionWebElem():
    bandera = False
    global D_CHR
    time.sleep(15)
    cant = D_CHR.find_element_by_xpath(cantidad_busquedas).text
    pri = cant.split(",")
    intcant = int("".join(pri))
    print(type(intcant), intcant)
    da = []
    r = []
    time.sleep(5)
    try:
        time.sleep(10)
        D_CHR.find_element_by_xpath(nombre_popup).click()

    except:
        time.sleep(60)
        D_CHR.find_element_by_xpath(nombre_popup).click()
        bandera = True
    time.sleep(15)

    lim = 10

    if (intcant == 1):
        lim = 1
    elif (intcant < 10):
        lim = intcant

    for i in range(0, lim):
        time.sleep(3)
        r = DatosElem()
        print("ITERCION", i)
        if (lim != 1):
            if (i >= 1):
                time.sleep(5)

                try:
                    D_CHR.find_element_by_xpath(siguiente_boton).click()
                except:
                    try:
                        time.sleep(40)
                        D_CHR.find_element_by_xpath(siguiente_boton).click()
                    except:
                        i = 10
            else:
                D_CHR.find_element_by_xpath(siguiente_i_boton).click()
                if bandera:
                    D_CHR.find_element_by_xpath(siguiente_i_boton).click()
                    time.sleep(10)
                    # D_CHR.find_element_by_xpath(nombre_popup2).click()
                    bandera = False
            print("Estamos en la iteración ", i)
            time.sleep(3)
        da.append(r)

    return [da, intcant]


def cerrar():
    global D_CHR
    D_CHR.close()


def connection():
    try:
        connection = pyodbc.connect(
            'Driver={SQL Server};SERVER=JOSEPHROJAS;DATABASE=BotCantidad;UID=sa;PWD=_Cl4ve2021J3RP_')
        return connection
    except:
        print("Fallo Conexion")


def insertar(issue, cargo, elem, link):
    conn = connection()
    cursor = conn.cursor()
    sql = "INSERT INTO Cantidad_Vacantes (issue, solicitud, CantidadElempleo, cantidadLinkedIn) VALUES (?,?,?,?);"

    cursor.execute(sql, issue, cargo, elem, link)
    cursor.commit()
    cursor.close()
