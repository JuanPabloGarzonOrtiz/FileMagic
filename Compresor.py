#Compresor de Archivos
# -*- coding: latin-1 -*-

import os; import platform; import sys 
import time; from threading import Thread 
from zipfile import ZipFile 
from py7zr import SevenZipFile 
import tarfile 
import bz2; from bz2 import BZ2File 
import pdf2docx 
import pandas; import tabula 
import pdf2image
from spire.xls import Workbook, FileFormat as FileFormatXls 
from spire.doc import Document, FileFormat as FileFormatDoc 
from spire.presentation import Presentation, FileFormat as FileFormatPres
import PyPDF2 
from reportlab.pdfgen import canvas 

class metodos:
    def limpiar_Consola(self):
        if platform.system() == "Windows":
            os.system("cls")
        else:
            os.system("clear")
    
    def monitor_Tamano(ancho_Pantalla, largo_Pantalla):
        while True:
            if ancho_Pantalla != os.get_terminal_size()[0] or largo_Pantalla != os.get_terminal_size()[1]:
                metodos().limpiar_Consola()
                print("El tamaño de la pantalla se cambio")
                ancho_Pantalla = os.get_terminal_size()[0]; largo_Pantalla = os.get_terminal_size()[1]
                time.sleep(2); metodos().limpiar_Consola()
                print("Presione cualquier tecla para continuar")
            time.sleep(1)
    
    def compresor_Archivador(self,modulo, extencion_de_Fichero, operacion, binario = False):
        archivo_a_Comprimir_Archivar = input(f"Ingrese la dirección del archivo que desea comprimir en {extencion_de_Fichero}: ")
        if os.path.isfile(archivo_a_Comprimir_Archivar) or os.path.isdir(archivo_a_Comprimir_Archivar):
            ruta, nombre_Archivo = os.path.split(archivo_a_Comprimir_Archivar)
            os.chdir(ruta)
            #Archivo
            if "." in nombre_Archivo:
                nombre_Archivo_Formato = os.path.splitext(nombre_Archivo)[0]
                try:
                    if binario:
                        with open(archivo_a_Comprimir_Archivar, "rb") as archivo_original:
                            datos_Comprimidos = bz2.compress(archivo_original.read())
                    with modulo(f"{nombre_Archivo_Formato}{extencion_de_Fichero}", "w") as objeto: 
                        if operacion == "write" and binario:
                            objeto.write(datos_Comprimidos)
                        elif operacion == "write":
                            objeto.write(archivo_a_Comprimir_Archivar, arcname = nombre_Archivo)
                        elif operacion == "add":
                            objeto.add(archivo_a_Comprimir_Archivar, arcname = nombre_Archivo)
                        input("Se realizó la compresion con exito")
                except Exception as e:
                    input(f"Error al comprimir: {e}")
            #Carpeta
            else:
                try:
                    with modulo(f"{nombre_Archivo}{extencion_de_Fichero}","w") as comprimir:
                        print(f"Los archivos a comprimir son:")
                        for carpeta, _, archivos in os.walk(archivo_a_Comprimir_Archivar):
                            for archivo in archivos:
                                print(archivo)
                                ruta_Archivo = os.path.join(carpeta,archivo)
                                ruta_en_Carpeta = os.path.relpath(ruta_Archivo, ruta)
                                if operacion == "write":
                                    comprimir.write(ruta_Archivo, arcname = ruta_en_Carpeta)
                                elif operacion == "add":
                                    comprimir.add(ruta_Archivo, arcname = ruta_en_Carpeta)
                    input("Se realizo la compresion de manera exitosa")
                except Exception as e:
                    input(f"Error al copilar: {e}")
        else:
            input("La ruta esta mal especificada o el archivo no existe")

    def descompresor(self, ubicacion, modulo, ruta, binario = False, nombre_Archivo = ""):
        if os.path.isfile(ubicacion):
            try:
                if binario:
                    with modulo(ubicacion, "rb") as archivo_Comprimido:
                        with open( f"{nombre_Archivo}", "wb") as archivo_Original:
                            archivo_Original.write(bz2.decompress(archivo_Comprimido.read()))
                else:
                    with modulo(ubicacion, "r") as descomprimido:
                        descomprimido.extractall(ruta)
                input("Se realizo la extraccion de manera exitosa")
            except Exception as e: 
                input(f"Error al copilar: {e}")
        else:
            input("La ruta esta mal espesificada")

    def seleccion_Conversion(self, nombre_Archivo, direccion_Archivo, Formatos_Disponibles, formato_Entrada):
        print(f"El Archivo a convertir es {formato_Entrada}, Indique a que tipo de archivo lo desea convertir:")
        contador = 1; bucle = True
        for formato in Formatos_Disponibles:
            print(f"{contador}.{formato}"); contador +=1
        while bucle:
            formato_a_Convertir = input("Respuesta:").lower().strip()
            for formato in Formatos_Disponibles:
                if formato_a_Convertir == formato:
                    archivo_nuevo = f"{nombre_Archivo}.{formato}"
                    bucle = False; break
            else:
                input(f"No es uno de los formatos indicados para poder convertir desde un {formato_a_Convertir}")
        if formato_Entrada == ".pdf":
            metodos().conversor_from_PDF(direccion_Archivo, archivo_nuevo,formato_a_Convertir)
        elif formato_Entrada == ".docx":
            metodos().conversor_from_DOCX(direccion_Archivo,archivo_nuevo, formato_a_Convertir)
        elif formato_Entrada == ".pptx":
            metodos().conversor_from_PPTX(direccion_Archivo,archivo_nuevo, formato_a_Convertir)
        elif formato_Entrada == ".xlsx":
            metodos().conversor_from_XLSX(direccion_Archivo,archivo_nuevo, formato_a_Convertir)
        elif formato_Entrada == ".txt":
            metodos().conversor_from_TXT(direccion_Archivo,archivo_nuevo, formato_a_Convertir)

    def conversor_from_PDF(self, doc_Entry, doc_Exit, formato_a_Convertir, conversion_intermedia = False):
        try:
            if formato_a_Convertir == "docx":
                todocx = pdf2docx.Converter(doc_Entry); todocx.convert(doc_Exit)
                todocx.close()
                if conversion_intermedia == True:
                        os.remove(doc_Entry)
                input(f"Se convirtio a {formato_a_Convertir}")
            elif formato_a_Convertir == "xlsx":
                dataframe_cobinado = pandas.DataFrame(); titulos = []
                for titulo in tabula.read_pdf(doc_Entry, pages = 1) [0][:1]:
                    titulos.append(titulo)
                for tabla in tabula.read_pdf(doc_Entry, pages = "all", pandas_options = {"header": None}): 
                    if isinstance(tabla, pandas.DataFrame):
                        df = tabla.copy()
                        dataframe_cobinado = pandas.concat([dataframe_cobinado,df])
                dataframe_cobinado = dataframe_cobinado[1:]
                dataframe_cobinado.columns = titulos
                dataframe_cobinado.to_excel(doc_Exit, index = False)
                if conversion_intermedia == True:
                        os.remove(doc_Entry)
                input(f"Se convirtio a {formato_a_Convertir}")
            elif formato_a_Convertir == "png" or formato_a_Convertir == "jpg":
                i = 1; ruta,archivo = os.path.split(doc_Entry)
                if platform.system() == "Windows":
                    paginas = pdf2image.convert_from_path(pdf_path = doc_Entry, poppler_path = f"C:\\Users\\{os.getlogin()}\\scoop\\apps\\poppler\\24.02.0-0\\bin") 
                    nombre = os.path.splitext(archivo)[0]
                    for pagina in paginas:
                        if formato_a_Convertir == "png":
                            pagina.save(os.path.join(ruta, f"{nombre}-img-{i}.png"), "PNG")
                        elif formato_a_Convertir == "jpg":
                            pagina.save(os.path.join(ruta, f"{nombre}-img-{i}.jpg"), "JPEG")
                        i +=1 
                elif platform.system() == "Linux":
                    if formato_a_Convertir == "jpg":
                        formato_a_Convertir = "jpeg"
                    os.system(f"pdftoppm -{formato_a_Convertir} {doc_Entry} onepage")
                if conversion_intermedia == True:
                    os.remove(doc_Entry)
                input(f"Se convirtio en {formato_a_Convertir}")
            elif formato_a_Convertir == "pptx":
                os.system(f"pdf2pptx {doc_Entry}")
                if conversion_intermedia == True:
                    os.remove(doc_Entry)
                input(f"Se convirtio en {formato_a_Convertir}")
            elif formato_a_Convertir == "txt": 
                pdf_Object = open(doc_Entry,"rb")
                reader = PyPDF2.PdfReader(pdf_Object, strict = True); texto = ""
                for page in range (len(reader.pages)):
                    texto += reader.pages[page].extract_text()
                pdf_Object.close()
                with open(doc_Exit,"w", encoding = "utf-8") as text_file:
                    text_file.write(texto)
                if conversion_intermedia == True:
                    os.remove(doc_Entry)
                input(f"Se convirtio en {formato_a_Convertir}")
        except Exception as e:
            input(f"Se presento un error: {e}")

    def conversor_from_DOCX(self, doc_Entry, doc_Exit, formato_a_Convertir):
        try:
            doc_Contenedor = doc_Exit if formato_a_Convertir == "pdf" else f"{doc_Exit[0]}.pdf"
            document = Document()
            document.LoadFromFile(doc_Entry)
            document.SaveToFile(doc_Contenedor, FileFormatDoc.PDF)
            document.Close()
            if formato_a_Convertir == "pdf":
                input(f"Se convirtio en {formato_a_Convertir}")
            elif formato_a_Convertir  == "png" or formato_a_Convertir == "jpg" or formato_a_Convertir == "pptx" or formato_a_Convertir == "xlsx" or formato_a_Convertir == "txt":
                doc_Entry = f"{os.path.split(doc_Entry)[0]}\{doc_Contenedor}"
                conversion_intermedia = True
                metodos().conversor_from_PDF(doc_Entry,doc_Exit,formato_a_Convertir, conversion_intermedia)
        except Exception as e:
            input(f"Se presento un error: {e}")
    
    def conversor_from_PPTX(self, doc_Entry, doc_Exit, formato_a_Convertir):
        try:
            doc_Contenedor = doc_Exit if formato_a_Convertir == "pdf" else f"{doc_Exit[0]}.pdf"
            presentacion = Presentation()
            presentacion.LoadFromFile(doc_Entry)
            presentacion.SaveToFile(doc_Contenedor, FileFormatPres.PDF)
            presentacion.Dispose()
            if formato_a_Convertir == "pdf":
                input(f"Se convirtio en {formato_a_Convertir}")
            elif formato_a_Convertir == "docx" or formato_a_Convertir == "jpg" or formato_a_Convertir == "png":
                doc_Entry = f"{os.path.split(doc_Entry)[0]}\{doc_Contenedor[0]}.pdf"
                conversion_intermedia = True
                metodos().conversor_from_PDF(doc_Entry,doc_Exit,formato_a_Convertir, conversion_intermedia)
        except Exception as e:
            input(f"Se presento un error: {e}")

    def conversor_from_XLSX(self, doc_Entry, doc_Exit, formato_a_Convertir):
        try:
            doc_Contenedor = doc_Exit if formato_a_Convertir == "pdf" else f"{doc_Exit[0]}.pdf"
            workbook = Workbook()
            workbook.LoadFromFile(doc_Entry)
            workbook.ConverterSetting.SheetFitToPage = True
            workbook.SaveToFile(doc_Contenedor, FileFormatXls.PDF)
            workbook.Dispose()
            if formato_a_Convertir == "pdf":
                input(f"Se convirtio en {formato_a_Convertir}")
            elif formato_a_Convertir == "docx" or formato_a_Convertir == "jpg" or formato_a_Convertir == "png":
                doc_Entry = f"{os.path.split(doc_Entry)[0]}\{doc_Exit[0]}.pdf"
                conversion_intermedia = True
                metodos().conversor_from_PDF(doc_Entry,doc_Exit,formato_a_Convertir, conversion_intermedia)
        except Exception as e:
            input(f"Se presento un error: {e}")  

    def conversor_from_TXT(self, doc_Entry, doc_Exit, formato_a_Convertir):
        try:
            doc_Contenedor = doc_Exit if formato_a_Convertir == "pdf" else f"{doc_Exit[0]}.pdf"
            text = open(doc_Entry,"r", encoding = "utf-8").read()
            pdf_Canvas = canvas.Canvas(doc_Contenedor)
            y_psoscion= 770; num_line = 1
            while num_line < len(text.splitlines()):
                pdf_Canvas.drawString(45, y_psoscion, text.splitlines()[num_line]) 
                y_psoscion -=15; num_line +=1
                if y_psoscion <= 50:
                    pdf_Canvas.showPage(); y_psoscion = 770
            pdf_Canvas.save()
            if formato_a_Convertir == "pdf":
                input(f"Se convirtio en {formato_a_Convertir}")
            elif formato_a_Convertir == "docx":
                doc_Entry = f"{os.path.split(doc_Entry)[0]}\{doc_Exit[0]}.pdf"
                conversion_intermedia = True
                metodos().conversor_from_PDF(doc_Entry,doc_Exit,formato_a_Convertir, conversion_intermedia)
        except Exception as e:
            input(f"Se presento un error: {e}")  


class programa:
    while True:
        titulo = "Bienbenido al Compresor de Archivos"
        ancho_Pantalla = os.get_terminal_size()[0] 
        largo_Pantalla = os.get_terminal_size()[1]

        #Hilo
        thread = Thread(target=metodos.monitor_Tamano, args=(ancho_Pantalla,largo_Pantalla))
        thread.daemon = True
        thread.start()
        #Menu
        metodos().limpiar_Consola()
        print(titulo.rjust(round((ancho_Pantalla + len(titulo))/2)))
        print("1.Compresor \n2.Descompresor \n3.Agrupador \n4.Conversor \n5.Salir"); 
        match input("Respuesta:").lower().strip():
            case "compresor":
                metodos().limpiar_Consola()
                print("Compresor".rjust(round((ancho_Pantalla + len("Compresor"))/2)))
                print("Indique que tipo de compresion desea hacer: \n1.ZIP \n2.7Z \n3.BZ2")
                match input("Respuesta:").lower().strip():
                    case "zip":
                        modulo = ZipFile
                        extencion_de_Fichero = ".zip"
                        operacion = "write"
                        metodos().compresor_Archivador(modulo, extencion_de_Fichero, operacion)
                    case "7z":
                        modulo = SevenZipFile
                        extencion_de_Fichero = ".7z" 
                        operacion = "write"
                        metodos().compresor_Archivador(modulo, extencion_de_Fichero, operacion) 
                    case "bz2": #Al ser bianrio solo se comprimen archivos y no carpetas
                        modulo = BZ2File
                        extencion_de_Fichero = ".bz2"
                        operacion = "write"
                        binario = True
                        metodos().compresor_Archivador(modulo, extencion_de_Fichero, operacion ,binario)
                    case _:
                        input("Esa opcion no esta dentro de las espesificadas")        

            case "descompresor":
                metodos().limpiar_Consola()
                print("Compresor".rjust(round((ancho_Pantalla + len("Compresor"))/2)))
                print("Indique el archivo que dese descomprimir")
                ubicacion = input("Respuesta:")
                ruta, nombre_Archivo = os.path.split(ubicacion)
                nombre_Archivo, extencion_Fichero_Descompresar = os.path.splitext(nombre_Archivo)
                match (extencion_Fichero_Descompresar).lower().strip():
                    case ".zip":
                        modulo = ZipFile
                        metodos().descompresor(ubicacion, modulo, ruta)  
                    case ".7z":
                        modulo = SevenZipFile
                        metodos().descompresor(ubicacion, modulo, ruta)
                    case ".bz2":
                        modulo = BZ2File
                        binario = True
                        metodos().descompresor(ubicacion, modulo, ruta, binario, nombre_Archivo)
                    case ".tar":
                        modulo = tarfile.open
                        metodos().descompresor(ubicacion, modulo, ruta)
                    case _:
                        input("Por el momento solo se tiene descompresion de: \n1.ZIP \n2.7Z \n3.BZ2 \n4.GZIP \n5.TAR")

            case "agrupador":
                metodos().limpiar_Consola()
                print("Agrupador".rjust(round((ancho_Pantalla + len("Agrupador"))/2)))
                print("Indique que tipo de agrupacion desea hacer: \n1.TAR")
                match input("Respuesta:").lower().strip():
                    case "tar":
                        modulo = tarfile.open
                        extencion_de_Fichero = ".tar" 
                        operacion = "add"
                        metodos().compresor_Archivador(modulo, extencion_de_Fichero, operacion)

            case "conversor":
                metodos().limpiar_Consola()
                print("Conversor".rjust(round((ancho_Pantalla + len("Conversor"))/2)))
                print("Ingresea la direccion del archivo que desa convertir:")
                direccion_Archivo = input("Respuesta:").lower().strip()
                Formatos_Disponibles = []
                if os.path.isfile(direccion_Archivo):
                    archivo = os.path.split(direccion_Archivo)[1]
                    nombre_Archivo, formato_Entrada  = os.path.splitext(archivo)
                    Formatos_Disponibles.clear()
                    match formato_Entrada:
                        case ".pdf":
                            Formatos_Disponibles.extend(["docx","pptx","xlsx","jpg","png","txt"])
                        case ".docx":
                            Formatos_Disponibles.extend(["pdf","pptx","xlsx","jpg","png","txt"]) # [pdf], pptx, xlsx, jpg, png, txt
                        case ".pptx":
                            Formatos_Disponibles.extend(["pdf","docx","jpg","png"]) # (pdf), docx, jpg ,png
                        case ".xlsx":
                            Formatos_Disponibles.extend(["pdf","docx","jpg","png"]) # [pdf], docx, jpg, png "pdf"
                        case ".txt":
                            Formatos_Disponibles.extend(["pdf","docx"]) # pdf, docx
                    metodos().seleccion_Conversion(nombre_Archivo, direccion_Archivo, Formatos_Disponibles, formato_Entrada)        
                else:
                    input("El Archivo indicado no existe o no esta en la ubicacion indicada")
            
            case "salir":
                metodos().limpiar_Consola(); sys.exit(0)

            case _:
                input("Esa opcion no esta dentro de las espesificadas") 
