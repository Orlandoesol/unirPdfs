
#Autor Orlando Espinosa Olivero
import os
import PyPDF2
import openpyxl
import tkinter as tk

def unir_pdfs():

    # Directorios de entrada y salida
    carpeta_entrada1 = "C:/Users/Orlando/Desktop/c1"
    carpeta_entrada2 = "C:/Users/Orlando/Desktop/c2"
    carpeta_salida = "C:/Users/Orlando/Desktop/c3"
    archivo_excel = "C:/Users/Orlando/Desktop/lista.xlsx"

    # Leer los nombres de salida desde el archivo Excel
    nombres_de_salida = []
    wb = openpyxl.load_workbook(archivo_excel)
    hoja = wb.active

    for fila in hoja.iter_rows(min_row=1, max_col=1, values_only=True):
        nombres_de_salida.append(fila[0])

    # Obtener la lista de archivos en las carpetas de entrada
    archivos_carpeta1 = os.listdir(carpeta_entrada1)
    archivos_carpeta2 = os.listdir(carpeta_entrada2)

    # Ordenar los archivos alfabéticamente para asegurar el orden correcto
    archivos_carpeta1.sort()
    archivos_carpeta2.sort()

    # Iterar sobre los archivos en ambas carpetas
    for archivo1, archivo2 in zip(archivos_carpeta1, archivos_carpeta2):
        # Abre los archivos PDF de ambas carpetas
        pdf1 = open(os.path.join(carpeta_entrada1, archivo1), 'rb')
        pdf2 = open(os.path.join(carpeta_entrada2, archivo2), 'rb')

        # Crea objetos PDFReader para cada archivo
        pdf_reader1 = PyPDF2.PdfReader(pdf1)
        pdf_reader2 = PyPDF2.PdfReader(pdf2)

        # Crear un objeto PDFWriter para el archivo de salida
        pdf_writer = PyPDF2.PdfWriter()

        # Agrega páginas del primer archivo
        for page_num in range(len(pdf_reader1.pages)):
            page = pdf_reader1.pages[page_num]
            pdf_writer.add_page(page)

        # Agrega páginas del segundo archivo
        for page_num in range(len(pdf_reader2.pages)):
            page = pdf_reader2.pages[page_num]
            pdf_writer.add_page(page)

        # Crea un nuevo archivo PDF para la salida
        nombre_salida_pdf = os.path.splitext(archivo1)[0] +'.pdf'
        pdf_salida = open(os.path.join(carpeta_salida, nombre_salida_pdf), 'wb')

        # Escribe el contenido del PDFWriter en el archivo de salida
        pdf_writer.write(pdf_salida)

        # Cierra el archivo de salida
        pdf_salida.close()

        # Cierra los archivos PDF de ambas carpetas
        pdf1.close()
        pdf2.close()
# Crear la ventana principal
ventana = tk.Tk()
ventana.title("Unión de PDFs")

# Crear un botón para ejecutar la función de unión de PDFs
boton = tk.Button(ventana, text="Unir PDFs", command=unir_pdfs)
boton.pack()

# Ejecutar el bucle principal de Tkinter
ventana.mainloop()