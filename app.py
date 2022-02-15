from reportlab.pdfgen import canvas
import pandas as pd

#Tamaño de hojas
A4 = [841, 595]
A5 = [595, 420]

weight = 841
height = 595

#Input a traves de un archivo de Excel

excel = "SEGUIMIENTO_PEDIDOS.xlsm"

df = pd.read_excel(io= "SEGUIMIENTO_PEDIDOS.xlsm", sheet_name= "DV")

#Títulos de carteles
title = "PANELES Y PUERTAS"
title2 = "PERFILES"
title3 = "CARPETA"

co = df["Unit"].values
mo = df["MO no"].values
modelo = df["CO Item no"].values
status = df["MO sts"].values

for line in range(len(mo)):
    if status[line] != "90-90":
        moFloat = f"MO: {mo[line]}"
        #Creación de cartel para puertas y paneles
        pdf_Puertas = canvas.Canvas(co[line] + "_" + "PUERTAS" + ".pdf", pagesize=A4)

        pdf_Puertas.setFontSize(66)
        pdf_Puertas.drawCentredString(weight/2, height - 100, title)
        pdf_Puertas.drawCentredString(
            weight/2, height - 220, "PEDIDO: " + co[line])
        pdf_Puertas.drawCentredString(weight/2, height - 340, moFloat)
        pdf_Puertas.drawCentredString(weight/2, height - 460, modelo[line])
        pdf_Puertas.setFontSize(30)
        #pdf_Puertas.drawCentredString(weight/2, height - 580, pais)
        pdf_Puertas.save()

        #Creación de cartel para perfiles
        pdf_Perfiles = canvas.Canvas(
            co[line] + "_" + title2 + ".pdf", pagesize=A4)

        pdf_Perfiles.setFontSize(66)
        pdf_Perfiles.drawCentredString(weight/2, height - 100, title2)
        pdf_Perfiles.drawCentredString(
            weight/2, height - 220, "PEDIDO: " + co[line])
        pdf_Perfiles.drawCentredString(weight/2, height - 340, moFloat)
        pdf_Perfiles.drawCentredString(weight/2, height - 460, modelo[line])
        pdf_Perfiles.setFontSize(30)
        #pdf_Perfiles.drawCentredString(weight/2, height - 580, pais)
        pdf_Perfiles.save()

        #Creación de cartel para arpeta
        pdf_Carpeta = canvas.Canvas(
            co[line] + "_" + title3 + ".pdf", pagesize=A5)

        pdf_Carpeta.setFontSize(48)
        pdf_Carpeta.drawCentredString(595/2, 420-60, "CO: " + co[line])
        pdf_Carpeta.drawCentredString(595/2, 420-150, modelo[line])
        pdf_Carpeta.drawCentredString(595/2, 420-270, moFloat)
        #pdf_Carpeta.drawCentredString(595/2, 420-390, pais)
        pdf_Carpeta.save()

        print(co[line] + " >> " + moFloat + " >> " + modelo [line])