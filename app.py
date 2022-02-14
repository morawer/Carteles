from nbformat import write
from reportlab.pdfgen import canvas

A4 = [841, 595]
A5 = [595, 420]

weight = 841
height = 595

co = input ("Número de pedido: ")
mo = input ("Número de MO: ")
linea = input ("Linea: ")
modelo = input ("Modelo: ")
pais = input ("País: ")

title = "PANELES Y PUERTAS"
title2 = "PERFILES"
title3 = "CARPETA"

pdf_Puertas = canvas.Canvas(co + "_" + linea + "_" + "PUERTAS" + ".pdf", pagesize=A4)

pdf_Puertas.setFontSize(66)
pdf_Puertas.drawCentredString(weight/2, height - 100, title)
pdf_Puertas.drawCentredString(
    weight/2, height - 220, "PEDIDO: " + co + " - " + linea)
pdf_Puertas.drawCentredString(weight/2, height - 340, "MO: " + mo)
pdf_Puertas.drawCentredString(weight/2, height - 460, modelo)
pdf_Puertas.setFontSize(30)
pdf_Puertas.drawCentredString(weight/2, height - 580, pais)
pdf_Puertas.save()


pdf_Perfiles = canvas.Canvas(
    co + "_" + linea + "_" + title2 + ".pdf", pagesize=A4)

pdf_Perfiles.setFontSize(66)
pdf_Perfiles.drawCentredString(weight/2, height - 100, title2)
pdf_Perfiles.drawCentredString(
    weight/2, height - 220, "PEDIDO: " + co + " - " + linea)
pdf_Perfiles.drawCentredString(weight/2, height - 340, "MO: " + mo)
pdf_Perfiles.drawCentredString(weight/2, height - 460, modelo)
pdf_Perfiles.setFontSize(30)
pdf_Perfiles.drawCentredString(weight/2, height - 580, pais)
pdf_Perfiles.save()

pdf_Carpeta = canvas.Canvas(
    co + "_" + linea + "_" + title3 + ".pdf", pagesize=A5)

pdf_Carpeta.setFontSize(48)
pdf_Carpeta.drawCentredString(595/2, 420-60, "CO: " + co + " - " + linea)
pdf_Carpeta.drawCentredString(595/2, 420-150, modelo)
pdf_Carpeta.drawCentredString(595/2, 420-270, "MO: " + mo)
pdf_Carpeta.drawCentredString(595/2, 420-390, pais)
pdf_Carpeta.save()




