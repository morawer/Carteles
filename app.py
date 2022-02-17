from reportlab.pdfgen import canvas
import pandas as pd
import os
import openpyxl

#Tamaño de hojas
A4 = [841, 595]
A4Carpeta = [595, 841]

weight = 841
height = 595

weight2 = 595
height2 = 841

pathDestino = "U:/OPERACIONES/08 FÁBRICA/1 AUTOMATIZACIÓN CARTELES TRABAJO JEFES TURNO/"

if not os.path.exists(pathDestino):
    os.makedirs(pathDestino)

#Input a traves de un archivo de Excel

excel = "SEGUIMIENTO_PEDIDOS.xlsm"

df = pd.read_excel(excel, sheet_name= "AHU")

excel_observaciones = openpyxl.load_workbook("OBSERVACIONES PEDIDOS.xlsx")
sheet1_observaciones = excel_observaciones.active

excel_protocolo = openpyxl.load_workbook("CL - Autocontrol Producción DV - Ed. 02.xlsx")
sheet1_protocolo = excel_protocolo.active

#Títulos de carteles
title = "PANELES Y PUERTAS"
title2 = "PERFILES"
title3 = "CARPETA"

co = df["Unit"].values
mo = df["MO no"].values
modelo = df["CO Item no"].values
status = df["MO sts"].values
fecha = df["MO Start"].values
pais = df["Country"].values

language = {
    "CZ" : "REP. CHECA",
    "ES" : "ESPAÑA",
    "FR" : "FRANCIA",
    "GB" : "REINO UNIDO",
    "GR" : "GRECIA",
    "MA" : "MARRUECOS",
    "PT" : "PORTUGAL",
    "US" : "ESTADOS UNIDOS"
}

for line in range(len(mo)):
    if status[line] != "90-90":
        moFloat = f"MO: {mo[line]:.0f}"
        #Creación de cartel para puertas y paneles
        pdf_Puertas = canvas.Canvas(pathDestino + co[line] + "_" + "PUERTAS" + ".pdf", pagesize=A4)

        pdf_Puertas.setFontSize(66)
        pdf_Puertas.drawCentredString(weight/2, height - 100, title)
        pdf_Puertas.drawCentredString(
            weight/2, height - 220, "PEDIDO: " + co[line])
        pdf_Puertas.drawCentredString(weight/2, height - 340, moFloat)
        pdf_Puertas.drawCentredString(weight/2, height - 460, modelo[line])
        pdf_Puertas.setFontSize(50)
        pdf_Puertas.drawCentredString(weight/2, height - 580, language[pais[line]])
        pdf_Puertas.save()

        #Creación de cartel para perfiles
        pdf_Perfiles = canvas.Canvas(
            pathDestino + co[line] + "_" + title2 + ".pdf", pagesize=A4)

        pdf_Perfiles.setFontSize(66)
        pdf_Perfiles.drawCentredString(weight/2, height - 100, title2)
        pdf_Perfiles.drawCentredString(
            weight/2, height - 220, "PEDIDO: " + co[line])
        pdf_Perfiles.drawCentredString(weight/2, height - 340, moFloat)
        pdf_Perfiles.drawCentredString(weight/2, height - 460, modelo[line])
        pdf_Perfiles.setFontSize(50)
        pdf_Perfiles.drawCentredString(weight/2, height - 580, language[pais[line]])
        pdf_Perfiles.save()

        #Creación de cartel para arpeta
        pdf_Carpeta = canvas.Canvas(
            pathDestino + co[line] + "_" + title3 + ".pdf", pagesize=A4Carpeta)

        pdf_Carpeta.setFontSize(56)
        pdf_Carpeta.drawCentredString(weight2/2, height2-60, "CO: " + co[line])
        pdf_Carpeta.drawCentredString(weight2/2, height2-150, modelo[line])
        pdf_Carpeta.drawCentredString(weight2/2, height2-270, moFloat)
        pdf_Carpeta.drawCentredString(weight2/2, height2-390, language[pais[line]])
        pdf_Carpeta.save()

        celdaA4_observaciones = sheet1_observaciones.cell(row=4, column=1)
        celdaA4_observaciones.value = co[line]
        celdaC4_observaciones = sheet1_observaciones.cell(row=4, column=3)
        celdaC4_observaciones.value = modelo[line]
        celdaD4_observaciones = sheet1_observaciones.cell(row=4, column=4)
        celdaD4_observaciones.value = mo[line]
        celdaE4_observaciones = sheet1_observaciones.cell(row= 4, column=5)
        celdaE4_observaciones.value = language[pais[line]]
        excel_observaciones.save(pathDestino + co[line] + "_OBSERVACIONES.xlsx")

        celdaB4_protocolo = sheet1_protocolo.cell(row=4, column=2)
        celdaB4_protocolo.value = co[line]
        celdaC4_protocolo = sheet1_protocolo.cell(row= 4, column=4)
        celdaC4_protocolo.value = moFloat
        celdaB5_protocolo = sheet1_protocolo.cell(row=5, column=2)
        celdaB5_protocolo.value = modelo[line]
        
        excel_protocolo.save(pathDestino + co[line] + "_PROTOCOLO.xlsx")

        print(co[line] + " >> " + moFloat + " >> " + modelo [line])