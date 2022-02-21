from reportlab.pdfgen import canvas
import pandas as pd
import os
import openpyxl

#Tamaño de hojas
A4 = [841, 595]
A4Carpeta = [595, 841]

#Medidas de ancho y alto de las hojas para posicionar los textos en hojas apaisadas o en vertical.
ancho = 841
alto = 595

ancho2 = 595
alto2 = 841


#Destino carpeta de los PDF y Excel creado.
pathDestino = "U:/OPERACIONES/08 FÁBRICA/1 AUTOMATIZACIÓN CARTELES TRABAJO JEFES TURNO/"

#Si no existe la capreta de destino, se crea automaticamente.
if not os.path.exists(pathDestino):
    os.makedirs(pathDestino)

#Input a traves de un archivo de Excel
#excel = "SEGUIMIENTO_PEDIDOS.xlsm"
excel = "U:/OPERACIONES/05 PLANIFICACIÓN/SEGUIMIENTO_PEDIDOS_V04.xlsm"
df = pd.read_excel(excel, sheet_name= "AHU", skiprows=1)

print(df.head())

#Plantilla Excel de hoja de observaciones.
excel_observaciones = openpyxl.load_workbook("OBSERVACIONES PEDIDOS.xlsx")
sheet1_observaciones = excel_observaciones.active

#Plantilla Excel del protocolo de inspección
excel_protocolo = openpyxl.load_workbook("CL - Autocontrol Producción DV - Ed. 02.xlsx")
sheet1_protocolo = excel_protocolo.active

#Títulos de carteles
titulo = "PANELES Y PUERTAS"
titulo2 = "PERFILES"
titulo3 = "CARPETA"

#En el Excel "SEGUIMIENTO_PEDIDOS.xlsm" recojemos en variable los valores de cada columna que necesitamos.
co = df["Unit"].values
mo = df["MO no"].values
modelo = df["CO Item no"].values
status = df["MO sts"].values
fecha = df["MO Start"].values
pais = df["Country"].values

#Creamos un diccionario "ISO2 : Nombre de país".
isoPais = {
    "AF":"Afganistán",
	"AL":"Albania",
	"DE":"Alemania",
	"DZ":"Algeria",
	"AD":"Andorra",
	"AO":"Angola",
	"AI":"Anguila",
	"AQ":"Antártida",
	"AG":"Antigua y Barbuda",
	"AN":"Antillas Neerlandesas",
	"SA":"Arabia Saudita",
	"AR":"Argentina",
	"AM":"Armenia",
	"AW":"Aruba",
	"AU":"Australia",
	"AT":"Austria",
	"AZ":"Azerbayán",
	"BE":"Bélgica",
	"BS":"Bahamas",
	"BH":"Bahrein",
	"BD":"Bangladesh",
	"BB":"Barbados",
	"BZ":"Belice",
	"BJ":"Benín",
	"BT":"Bhután",
	"BY":"Bielorrusia",
	"MM":"Birmania",
	"BO":"Bolivia",
	"BA":"Bosnia y Herzegovina",
	"BW":"Botsuana",
	"BR":"Brasil",
	"BN":"Brunéi",
	"BG":"Bulgaria",
	"BF":"Burkina Faso",
	"BI":"Burundi",
	"CV":"Cabo Verde",
	"KH":"Camboya",
	"CM":"Camerún",
	"CA":"Canadá",
	"TD":"Chad",
	"CL":"Chile",
	"CN":"China",
	"CY":"Chipre",
	"VA":"Ciudad del Vaticano",
	"CO":"Colombia",
	"KM":"Comoras",
	"CG":"Congo",
	"CD":"Congo",
	"KP":"Corea del Norte",
	"KR":"Corea del Sur",
	"CI":"Costa de Marfil",
	"CR":"Costa Rica",
	"HR":"Croacia",
	"CU":"Cuba",
	"DK":"Dinamarca",
	"DM":"Dominica",
	"EC":"Ecuador",
	"EG":"Egipto",
	"SV":"El Salvador",
	"AE":"Emiratos Árabes Unidos",
	"ER":"Eritrea",
	"SK":"Eslovaquia",
	"SI":"Eslovenia",
	"ES":"España",
	"US":"Estados Unidos de América",
	"EE":"Estonia",
	"ET":"Etiopía",
	"PH":"Filipinas",
	"FI":"Finlandia",
	"FJ":"Fiyi",
	"FR":"Francia",
	"GA":"Gabón",
	"GM":"Gambia",
	"GE":"Georgia",
	"GH":"Ghana",
	"GI":"Gibraltar",
	"GD":"Granada",
	"GR":"Grecia",
	"GL":"Groenlandia",
	"GP":"Guadalupe",
	"GU":"Guam",
	"GT":"Guatemala",
	"GF":"Guayana Francesa",
	"GG":"Guernsey",
	"GN":"Guinea",
	"GQ":"Guinea Ecuatorial",
	"GW":"Guinea-Bissau",
	"GY":"Guyana",
	"HT":"Haití",
	"HN":"Honduras",
	"HK":"Hong kong",
	"HU":"Hungría",
	"IN":"India",
	"ID":"Indonesia",
	"IR":"Irán",
	"IQ":"Irak",
	"IE":"Irlanda",
	"BV":"Isla Bouvet",
	"IM":"Isla de Man",
	"CX":"Isla de Navidad",
	"NF":"Isla Norfolk",
	"IS":"Islandia",
	"BM":"Islas Bermudas",
	"KY":"Islas Caimán",
	"CC":"Islas Cocos (Keeling)",
	"CK":"Islas Cook",
	"AX":"Islas de Åland",
	"FO":"Islas Feroe",
	"GS":"Islas Georgias del Sur y Sandwich del Sur",
	"HM":"Islas Heard y McDonald",
	"MV":"Islas Maldivas",
	"FK":"Islas Malvinas",
	"MP":"Islas Marianas del Norte",
	"MH":"Islas Marshall",
	"PN":"Islas Pitcairn",
	"SB":"Islas Salomón",
	"TC":"Islas Turcas y Caicos",
	"UM":"Islas Ultramarinas Menores de Estados Unidos",
	"VG":"Islas Vírgenes Británicas",
	"VI":"Islas Vírgenes de los Estados Unidos",
	"IL":"Israel",
	"IT":"Italia",
	"JM":"Jamaica",
	"JP":"Japón",
	"JE":"Jersey",
	"JO":"Jordania",
	"KZ":"Kazajistán",
	"KE":"Kenia",
	"KG":"Kirgizstán",
	"KI":"Kiribati",
	"KW":"Kuwait",
	"LB":"Líbano",
	"LA":"Laos",
	"LS":"Lesoto",
	"LV":"Letonia",
	"LR":"Liberia",
	"LY":"Libia",
	"LI":"Liechtenstein",
	"LT":"Lituania",
	"LU":"Luxemburgo",
	"MX":"México",
	"MC":"Mónaco",
	"MO":"Macao",
	"MK":"Macedônia",
	"MG":"Madagascar",
	"MY":"Malasia",
	"MW":"Malawi",
	"ML":"Mali",
	"MT":"Malta",
	"MA":"Marruecos",
	"MQ":"Martinica",
	"MU":"Mauricio",
	"MR":"Mauritania",
	"YT":"Mayotte",
	"FM":"Micronesia",
	"MD":"Moldavia",
	"MN":"Mongolia",
	"ME":"Montenegro",
	"MS":"Montserrat",
	"MZ":"Mozambique",
	"NA":"Namibia",
	"NR":"Nauru",
	"NP":"Nepal",
	"NI":"Nicaragua",
	"NE":"Niger",
	"NG":"Nigeria",
	"NU":"Niue",
	"NO":"Noruega",
	"NC":"Nueva Caledonia",
	"NZ":"Nueva Zelanda",
	"OM":"Omán",
	"NL":"Países Bajos",
	"PK":"Pakistán",
	"PW":"Palau",
	"PS":"Palestina",
	"PA":"Panamá",
	"PG":"Papúa Nueva Guinea",
	"PY":"Paraguay",
	"PE":"Perú",
	"PF":"Polinesia Francesa",
	"PL":"Polonia",
	"PT":"Portugal",
	"PR":"Puerto Rico",
	"QA":"Qatar",
	"GB":"Reino Unido",
	"CF":"República Centroafricana",
	"CZ":"República Checa",
	"DO":"República Dominicana",
	"RE":"Reunión",
	"RW":"Ruanda",
	"RO":"Rumanía",
	"RU":"Rusia",
	"EH":"Sahara Occidental",
	"WS":"Samoa",
	"AS":"Samoa Americana",
	"BL":"San Bartolomé",
	"KN":"San Cristóbal y Nieves",
	"SM":"San Marino",
	"MF":"San Martín (Francia)",
	"PM":"San Pedro y Miquelón",
	"VC":"San Vicente y las Granadinas",
	"SH":"Santa Elena",
	"LC":"Santa Lucía",
	"ST":"Santo Tomé y Príncipe",
	"SN":"Senegal",
	"RS":"Serbia",
	"SC":"Seychelles",
	"SL":"Sierra Leona",
	"SG":"Singapur",
	"SY":"Siria",
	"SO":"Somalia",
	"LK":"Sri lanka",
	"ZA":"Sudáfrica",
	"SD":"Sudán",
	"SE":"Suecia",
	"CH":"Suiza",
	"SR":"Surinám",
	"SJ":"Svalbard y Jan Mayen",
	"SZ":"Swazilandia",
	"TJ":"Tadjikistán",
	"TH":"Tailandia",
	"TW":"Taiwán",
	"TZ":"Tanzania",
	"IO":"Territorio Británico del Océano Índico",
	"TF":"Territorios Australes y Antárticas Franceses",
	"TL":"Timor Oriental",
	"TG":"Togo",
	"TK":"Tokelau",
	"TO":"Tonga",
	"TT":"Trinidad y Tobago",
	"TN":"Tunez",
	"TM":"Turkmenistán",
	"TR":"Turquía",
	"TV":"Tuvalu",
	"UA":"Ucrania",
	"UG":"Uganda",
	"UY":"Uruguay",
	"UZ":"Uzbekistán",
	"VU":"Vanuatu",
	"VE":"Venezuela",
	"VN":"Vietnam",
	"WF":"Wallis y Futuna",
	"YE":"Yemen",
	"DJ":"Yibuti",
	"ZM":"Zambia",
}

#Leemos todo el Excel por cada linea/fila de columna de MO.
for line in range(len(mo)):
    
    #Creamos la condición de solo trabajar con lineas de status diferente a "90-90"
    if status[line] != "90-90":
        moFloat = f"MO: {mo[line]:.0f}"
        #Creación de cartel para puertas y paneles
        pdf_Puertas = canvas.Canvas(pathDestino + co[line] + "_" + "PUERTAS" + ".pdf", pagesize=A4)

        pdf_Puertas.setFontSize(66)
        pdf_Puertas.drawCentredString(ancho/2, alto - 100, titulo)
        pdf_Puertas.drawCentredString(
            ancho/2, alto - 220, "PEDIDO: " + co[line])
        pdf_Puertas.drawCentredString(ancho/2, alto - 340, moFloat)
        pdf_Puertas.drawCentredString(ancho/2, alto - 460, modelo[line])
        pdf_Puertas.setFontSize(50)
        pdf_Puertas.drawCentredString(ancho/2, alto - 580, isoPais[pais[line]])
        pdf_Puertas.save()
        #Creación de cartel para perfiles
        pdf_Perfiles = canvas.Canvas(
            pathDestino + co[line] + "_" + titulo2 + ".pdf", pagesize=A4)

        pdf_Perfiles.setFontSize(66)
        pdf_Perfiles.drawCentredString(ancho/2, alto - 100, titulo2)
        pdf_Perfiles.drawCentredString(
            ancho/2, alto - 220, "PEDIDO: " + co[line])
        pdf_Perfiles.drawCentredString(ancho/2, alto - 340, moFloat)
        pdf_Perfiles.drawCentredString(ancho/2, alto - 460, modelo[line])
        pdf_Perfiles.setFontSize(50)
        pdf_Perfiles.drawCentredString(ancho/2, alto - 580, isoPais[pais[line]])
        pdf_Perfiles.save()

        #Creación de cartel para arpeta
        pdf_Carpeta = canvas.Canvas(
            pathDestino + co[line] + "_" + titulo3 + ".pdf", pagesize=A4Carpeta)

        pdf_Carpeta.setFontSize(56)
        pdf_Carpeta.drawCentredString(ancho2/2, alto2-60, "CO: " + co[line])
        pdf_Carpeta.drawCentredString(ancho2/2, alto2-150, modelo[line])
        pdf_Carpeta.drawCentredString(ancho2/2, alto2-270, moFloat)
        pdf_Carpeta.drawCentredString(ancho2/2, alto2-390, isoPais[pais[line]])
        pdf_Carpeta.save()
        
        #Creacion de Excel "Observaciones" y relleno de los datos en celdas.
        celdaA4_observaciones = sheet1_observaciones.cell(row=4, column=1)
        celdaA4_observaciones.value = co[line]
        celdaC4_observaciones = sheet1_observaciones.cell(row=4, column=3)
        celdaC4_observaciones.value = modelo[line]
        celdaD4_observaciones = sheet1_observaciones.cell(row=4, column=4)
        celdaD4_observaciones.value = mo[line]
        celdaE4_observaciones = sheet1_observaciones.cell(row= 4, column=5)
        celdaE4_observaciones.value = isoPais[pais[line]]
        
        excel_observaciones.save(pathDestino + co[line] + "_OBSERVACIONES.xlsx")

        #Creacion de Excel "Protocolo de Inspección" y relleno de los datos en celdas.
        celdaB4_protocolo = sheet1_protocolo.cell(row=4, column=2)
        celdaB4_protocolo.value = co[line]
        celdaC4_protocolo = sheet1_protocolo.cell(row= 4, column=4)
        celdaC4_protocolo.value = moFloat
        celdaB5_protocolo = sheet1_protocolo.cell(row=5, column=2)
        celdaB5_protocolo.value = modelo[line]
        
        excel_protocolo.save(pathDestino + co[line] + "_PROTOCOLO.xlsx")
        
        #Imprimimos por consola la C0, MO y modelo de AHU de cada linea de pedido.
        print(co[line] + " >> " + moFloat + " >> " + modelo[line] + " >> " + isoPais[pais[line]])
