from msilib.schema import Feature
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
import time
from openpyxl import load_workbook
from datetime import date
from selenium.webdriver.support.ui import Select
from sympy import false
import Functions_Bot as FB

Features = FB.ExcelData()
Asunto = "Pool Issue "+Features[0]+" - "+str(int(Features[15]))
kta=FB.JsonAddMas(Features[0])

def run_away():
    FB.LoginELEM("joseph.rojas@exsis.com.co","Recru_2022_Auto")
    time.sleep(30)
    print("""

    Hecho login

    """)

    time.sleep(10)

    FB.filtrosElem(Features[2],Features[16],kta,Features[9],Features[10],Features[11],Features[12],Features[17],Features[13])

    a=FB.validacionVaca()
    print(a)

    print("""

    Hecho 30

    """)


    
    if (a!=True):
        excel_header = ["Nombre", "Profesión", "Experiencia laboral",
                    "Aspiración salarial", "Promedio aspiración salarial","Ultima Posición", "Ubicación",
                    "Porcentaje de dominio de Ingles", "Idiogen", "Correo", "Teléfono", "Whatsapp", "Descripción"]
        [Dat,cantElempleo] = FB.ExtraccionWebElem()
        df = pd.DataFrame(Dat, columns=excel_header)

        print("""



        Hecho extracción Elempleo



        """)
    else:
        df = pd.DataFrame(["No hay vacantes"])
    time.sleep(20)
    acumu = 0
    
    promedioGen = df['Promedio aspiración salarial']
    
    #print(promedioGen)
    pGn = promedioGen.mean()
    #print(pGn)
    pGn = pd.DataFrame([pGn])

    FB.LoginLink("soyexbot@gmail.com","Marketing2022.")
    time.sleep(10)
    [dpru,cantLink]=FB.busqLink(Features[0])
    enca = ["Nombre","ocupacion","ubicación","Link"]



    data_prueba = pd.DataFrame(dpru,columns=enca)

    FileName_Export = r"C:\\Users\\Joseph Rojas\\Exsis Software y Soluciones\Automation - Documentos\\Scraping_perfiles\\Historial\\Pool -" + \
        str(int(Features[15]))+"-"+str(Features[0])+".xlsx"
    with pd.ExcelWriter(FileName_Export, mode='w', engine='xlsxwriter') as writer:
        sheet_name = ['Elempleo','LinkedIn']
        for i in range(len(sheet_name)):
            if i == 0:
                df.to_excel(writer, sheet_name=sheet_name[i], index=False)
                FB.format_tbl(writer,sheet_name[i],df)
                pGn.to_excel(writer,sheet_name=sheet_name[i],startcol=4,startrow=11,index=false,header=false)            
            else:
                data_prueba.to_excel(writer, sheet_name=sheet_name[i], index=False)
                FB.format_tbl(writer,sheet_name[i],data_prueba)

    print("""



    Hecho Excel



    """)
    print(cantElempleo,cantLink)
 #   FB.insertar(int(Features[15]),Features[0],int(cantElempleo),int(cantLink))

   
    time.sleep(5)


    FB.cerrar()
