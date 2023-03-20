from io import BytesIO
from fastapi import FastAPI, File, UploadFile
from fastapi.responses import StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
import os, zipfile

app = FastAPI()

origins = [
    "http://localhost",
    "http://localhost:5173",
]

app.add_middleware(
    CORSMiddleware,
    allow_origins=origins,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Inicio conversor de PDF

@app.post("/api/v1/convert-pdf-to-csv")
async def convert_pdf_csv(file: UploadFile = File(...)):
    import tabula
    # Función para eliminar los archivos y carpetas creadas
    def deleteFiles():
        # Si la carpeta pdf_files existe, elimina todos los archivos PDF y luego elimina la carpeta pdf_files
        if os.path.exists("pdf_to_csv_files"):
            for pdf_file_ in os.listdir("pdf_to_csv_files"):
                os.remove(os.path.join("pdf_to_csv_files", pdf_file_))
            os.rmdir("pdf_to_csv_files")
        # Si la carpeta csv_files existe, elimina todos los archivos CSV y luego elimina la carpeta csv_files
        if os.path.exists("pdf_csv_files"):
            for pdf_file_ in os.listdir("pdf_csv_files"):
                a = open(os.path.join("pdf_csv_files", pdf_file_), "rb")
                a = a.close()

                os.remove(os.path.join("pdf_csv_files", pdf_file_))
            os.rmdir("pdf_csv_files")
    # Elimina los archivos y carpetas creadas en la última ejecución
    deleteFiles()
    # Crea una carpeta para almacenar el PDF recibido
    os.mkdir("pdf_to_csv_files")
    # Guardar el PDF recibido en el directorio pdf_files
    file_name = os.path.join("pdf_to_csv_files", file.filename)
    # Abre el archivo PDF en modo binario y lo guarda en el directorio pdf_files
    with open(file_name, "wb") as buffer:
        # Lee el archivo PDF en trozos de 1024 bytes
        while chunk := await file.read(1024):
            # Escribe el trozo de bytes en el archivo PDF
            buffer.write(chunk)
    # Especifica la ruta de tu archivo PDF
    file = file_name

    # Extrae todas las tablas del PDF
    all_tables = tabula.read_pdf(file, pages="all", multiple_tables=True)

    # Crea una carpeta para almacenar los archivos CSV extraídos
    os.makedirs("pdf_csv_files", exist_ok=True)

    # Utiliza un bucle for para ajustar el formato de cada tabla extraída y guardarla como un archivo CSV
    for i, table in enumerate(all_tables):
        # Elimina las filas y columnas vacías
        table.dropna(how="all", inplace=True)
        table.dropna(axis=1, how="all", inplace=True)
        # Reemplaza los nombres de las filas y columnas no deseadas con valores en blanco
        table.rename(columns=lambda x: "" if "Unnamed" in str(x) else x, inplace=True)
        table.rename(index=lambda x: "" if "Unnamed" in str(x) else x, inplace=True)
        # Si hay más de un archivo CSV, crea un archivo .zip y agrega todos los archivos CSV a él
        if len(all_tables) > 1:
            csv_file = os.path.join("pdf_csv_files", f"table_{i}.csv")
        else:
            # Si solo hay un archivo CSV, renombra el archivo con el nombre del PDF
            csv_file = os.path.join("pdf_csv_files", os.path.splitext(os.path.basename(file))[0] + ".csv")
        table.to_csv(csv_file, index=False)
    # Si hay más de un archivo CSV, crea un archivo .zip y agrega todos los archivos CSV a él
    zip_file_name = "csv_files.zip"
    with zipfile.ZipFile(zip_file_name, "w") as zip_file:
        # Agrega todos los archivos CSV a la carpeta csv_files
        for pdf_file_name in os.listdir("pdf_csv_files"):
            csv_file = os.path.join("pdf_csv_files", pdf_file_name)
            zip_file.write(csv_file, pdf_file_name)

    # Elimina la carpeta csv_files y sus contenidos
    for pdf_file_name in os.listdir("pdf_csv_files"):
        csv_file = os.path.join("pdf_csv_files", pdf_file_name)
        os.remove(csv_file)

    # Mueve el archivo zip a la carpeta csv_files
    os.rename("csv_files.zip", "pdf_csv_files/csv_files.zip")
    # Si el archivo zip existe, devuelve el archivo zip como respuesta y elimina los archivos y carpetas creadas
    if os.path.exists("pdf_csv_files/csv_files.zip"):
        data = open("pdf_csv_files/csv_files.zip", "rb")
        file_contents = data.read()
        data.close()
        response = StreamingResponse(BytesIO(file_contents), media_type="application/zip", headers={"Content-Disposition": "attachment; filename=csv_files.zip"})
        deleteFiles()
        return response

@app.post("/api/v1/convert-pdf-to-xls")
async def convert_pdf_xls(file: UploadFile = File(...)):
    import tabula
    import pandas as pd
    # Función para eliminar los archivos y carpetas creadas
    def deleteFiles():
        # Si la carpeta pdf_files existe, elimina todos los archivos PDF y luego elimina la carpeta pdf_files
        if os.path.exists("pdf_to_xls_files"):
            for pdf_file_ in os.listdir("pdf_to_xls_files"):
                os.remove(os.path.join("pdf_to_xls_files", pdf_file_))
            os.rmdir("pdf_to_xls_files")
        # Si la carpeta xls_files existe, elimina todos los archivos XLS y luego elimina la carpeta xls_files
        if os.path.exists("pdf_xls_files"):
            for xls_file_ in os.listdir("pdf_xls_files"):
                a = open(os.path.join("pdf_xls_files", xls_file_), "rb")
                a = a.close()

                os.remove(os.path.join("pdf_xls_files", xls_file_))
            os.rmdir("pdf_xls_files")
    # Elimina los archivos y carpetas creadas en la última ejecución
    deleteFiles()
    # Crea una carpeta para almacenar el PDF recibido
    os.mkdir("pdf_to_xls_files")
    # Guardar el PDF recibido en el directorio pdf_files
    file_name = os.path.join("pdf_to_xls_files", file.filename)
    # Abre el archivo PDF en modo binario y lo guarda en el directorio pdf_files
    with open(file_name, "wb") as buffer:
        # Lee el archivo PDF en trozos de 1024 bytes
        while chunk := await file.read(1024):
            # Escribe el trozo de bytes en el archivo PDF
            buffer.write(chunk)
    # Especifica la ruta de tu archivo PDF
    file = file_name

    # Extrae todas las tablas del PDF
    all_tables = tabula.read_pdf(file, pages="all", multiple_tables=True)

    # Crea una carpeta para almacenar los archivos XLS extraídos
    os.makedirs("pdf_xls_files", exist_ok=True)

    # Utiliza un bucle for para ajustar el formato de cada tabla extraída y guardarla como un archivo XLS
    for i, table in enumerate(all_tables):
        # Elimina las filas y columnas vacías
        table.dropna(how="all", inplace=True)
        table.dropna(axis=1, how="all", inplace=True)
        # Reemplaza los nombres de las filas y columnas no deseadas con valores en blanco
        table.rename(columns=lambda x: "" if "Unnamed" in str(x) else x, inplace=True)
        table.rename(index=lambda x: "" if "Unnamed" in str(x) else x, inplace=True)
        # Si hay más de un archivo XLS, crea un archivo .zip y agrega todos los archivos XLS a él
        if len(all_tables) > 1:
            xls_file = os.path.join("pdf_xls_files", f"table_{i}.xls")
        else:
            # Si solo hay un archivo XLS, renombra el archivo con el nombre del PDF
            xls_file = os.path.join("pdf_xls_files", os.path.splitext(os.path.basename(file))[0] + ".xls")
        table.to_excel(xls_file, index=False)
    # Si hay más de un archivo XLS, crea un archivo .zip y agrega todos los archivos XLS a él
    zip_file_name = "xls_files.zip"
    with zipfile.ZipFile(zip_file_name, "w") as zip_file:
        # Agrega todos los archivos XLS a la carpeta xls_files
        for xls_file_name in os.listdir("pdf_xls_files"):
            xls_file = os.path.join("pdf_xls_files", xls_file_name)
            zip_file.write(xls_file, xls_file_name)

    # Elimina la carpeta xls_files y sus contenidos
    for xls_file_name in os.listdir("pdf_xls_files"):
        xls_file = os.path.join("pdf_xls_files", xls_file_name)
        os.remove(xls_file)

    # Mueve el archivo zip a la carpeta xls_files
    os.rename("xls_files.zip", "pdf_xls_files/xls_files.zip")
    # Si el archivo zip existe, devuelve el archivo zip como respuesta y elimina los archivos y carpetas creadas
    if os.path.exists("pdf_xls_files/xls_files.zip"):
        data = open("pdf_xls_files/xls_files.zip", "rb")
        file_contents = data.read()
        data.close()
        response = StreamingResponse(BytesIO(file_contents), media_type="application/zip", headers={"Content-Disposition": "attachment; filename=xls_files.zip"})
        deleteFiles()
        return response
    
@app.post("/api/v1/convert-pdf-to-xlsx")
async def convert_pdf_xlsx(file: UploadFile = File(...)):
    import tabula
    import pandas as pd
    # Función para eliminar los archivos y carpetas creadas
    def deleteFiles():
        # Si la carpeta pdf_files existe, elimina todos los archivos PDF y luego elimina la carpeta pdf_files
        if os.path.exists("pdf_to_xlsx_files"):
            for pdf_file_ in os.listdir("pdf_to_xlsx_files"):
                os.remove(os.path.join("pdf_to_xlsx_files", pdf_file_))
            os.rmdir("pdf_to_xlsx_files")
        # Si la carpeta xlsx_files existe, elimina todos los archivos XLSX y luego elimina la carpeta xlsx_files
        if os.path.exists("pdf_xlsx_files"):
            for xlsx_file_ in os.listdir("pdf_xlsx_files"):
                a = open(os.path.join("pdf_xlsx_files", xlsx_file_), "rb")
                a = a.close()

                os.remove(os.path.join("pdf_xlsx_files", xlsx_file_))
            os.rmdir("pdf_xlsx_files")
    # Elimina los archivos y carpetas creadas en la última ejecución
    deleteFiles()
    # Crea una carpeta para almacenar el PDF recibido
    os.mkdir("pdf_to_xlsx_files")
    # Guardar el PDF recibido en el directorio pdf_files
    file_name = os.path.join("pdf_to_xlsx_files", file.filename)
    # Abre el archivo PDF en modo binario y lo guarda en el directorio pdf_files
    with open(file_name, "wb") as buffer:
        # Lee el archivo PDF en trozos de 1024 bytes
        while chunk := await file.read(1024):
            # Escribe el trozo de bytes en el archivo PDF
            buffer.write(chunk)
    # Especifica la ruta de tu archivo PDF
    file = file_name

    # Extrae todas las tablas del PDF
    all_tables = tabula.read_pdf(file, pages="all", multiple_tables=True)

    # Crea una carpeta para almacenar los archivos XLSX extraídos
    os.makedirs("pdf_xlsx_files", exist_ok=True)

    # Utiliza un bucle for para ajustar el formato de cada tabla extraída y guardarla como un archivo XLSX
    for i, table in enumerate(all_tables):
        # Elimina las filas y columnas vacías
        table.dropna(how="all", inplace=True)
        table.dropna(axis=1, how="all", inplace=True)
        # Reemplaza los nombres de las filas y columnas no deseadas con valores en blanco
        table.rename(columns=lambda x: "" if "Unnamed" in str(x) else x, inplace=True)
        table.rename(index=lambda x: "" if "Unnamed" in str(x) else x, inplace=True)
        # Si hay más de un archivo XLSX, crea un archivo .zip y agrega todos los archivos XLSX a él
        if len(all_tables) > 1:
            xlsx_file = os.path.join("pdf_xlsx_files", f"table_{i}.xlsx")
        else:
            # Si solo hay un archivo XLSX, renombra el archivo con el nombre del PDF
            xlsx_file = os.path.join("pdf_xlsx_files", os.path.splitext(os.path.basename(file))[0] + ".xlsx")
        table.to_excel(xlsx_file, index=False)
    # Si hay más de un archivo XLSX, crea un archivo .zip y agrega todos los archivos XLSX a él
    zip_file_name = "xlsx_files.zip"
    with zipfile.ZipFile(zip_file_name, "w") as zip_file:
        # Agrega todos los archivos XLSX a la carpeta xlsx_files
        for xlsx_file_name in os.listdir("pdf_xlsx_files"):
            xlsx_file = os.path.join("pdf_xlsx_files", xlsx_file_name)
            zip_file.write(xlsx_file, xlsx_file_name)

    # Elimina la carpeta xlsx_files y sus contenidos
    for xlsx_file_name in os.listdir("pdf_xlsx_files"):
        xlsx_file = os.path.join("pdf_xlsx_files", xlsx_file_name)
        os.remove(xlsx_file)

    # Mueve el archivo zip a la carpeta xlsx_files
    os.rename("xlsx_files.zip", "pdf_xlsx_files/xlsx_files.zip")
    # Si el archivo zip existe, devuelve el archivo zip como respuesta y elimina los archivos y carpetas creadas
    if os.path.exists("pdf_xlsx_files/xlsx_files.zip"):
        data = open("pdf_xlsx_files/xlsx_files.zip", "rb")
        file_contents = data.read()
        data.close()
        response = StreamingResponse(BytesIO(file_contents), media_type="application/zip", headers={"Content-Disposition": "attachment; filename=xlsx_files.zip"})
        deleteFiles()
        return response

@app.post("/api/v1/convert-pdf-to-docx")
async def convert_pdf_docx(file: UploadFile = File(...)):
    from pdf2docx import Converter
    def deleteFiles():
        # Si la carpeta pdf_files existe, elimina todos los archivos PDF y luego elimina la carpeta pdf_files
        if os.path.exists("pdf_to_docx_files"):
            for pdf_file_ in os.listdir("pdf_to_docx_files"):
                os.remove(os.path.join("pdf_to_docx_files", pdf_file_))
            os.rmdir("pdf_to_docx_files")
        # Si la carpeta csv_files existe, elimina todos los archivos CSV y luego elimina la carpeta csv_files
        if os.path.exists("pdf_docx_files"):
            for pdf_file_ in os.listdir("pdf_docx_files"):
                a = open(os.path.join("pdf_docx_files", pdf_file_), "rb")
                a = a.close()

                os.remove(os.path.join("pdf_docx_files", pdf_file_))
            os.rmdir("pdf_docx_files")
    deleteFiles()
    # Crea una carpeta para almacenar el PDF recibido
    os.mkdir("pdf_to_docx_files")
    # Guardar el PDF recibido en el directorio pdf_files
    file_name = os.path.join("pdf_to_docx_files", file.filename)
    # Abre el archivo PDF en modo binario y lo guarda en el directorio pdf_files
    with open(file_name, "wb") as buffer:
        # Lee el archivo PDF en trozos de 1024 bytes
        while chunk := await file.read(1024):
            # Escribe el trozo de bytes en el archivo PDF
            buffer.write(chunk)
    os.makedirs("pdf_docx_files", exist_ok=True)
    cv = Converter(file_name)
    cv.convert("pdf_docx_files/converted.docx", start=0, end=None)
    cv.close()
    zip_file_name = "docx_files.zip"
    with zipfile.ZipFile(zip_file_name, "w") as zip_file:
        # Agrega todos los archivos CSV a la carpeta csv_files
        for pdf_file_name in os.listdir("pdf_docx_files"):
            csv_file = os.path.join("pdf_docx_files", pdf_file_name)
            zip_file.write(csv_file, pdf_file_name)
    # Mueve el archivo zip a la carpeta csv_files
    os.rename("docx_files.zip", "pdf_docx_files/docx_files.zip")
    data = open("pdf_docx_files/docx_files.zip", "rb")
    file_contents = data.read()
    data.close()
    response = StreamingResponse(BytesIO(file_contents), media_type="application/docx", headers={"Content-Disposition": "attachment; filename=csv_files.zip"})
    deleteFiles()
    return response

@app.post("/api/v1/convert-pdf-to-doc")
async def convert_pdf_doc(file: UploadFile = File(...)):
    from pdf2docx import Converter
    def deleteFiles():
        # Si la carpeta pdf_files existe, elimina todos los archivos PDF y luego elimina la carpeta pdf_files
        if os.path.exists("csv_to_pdf_files"):
            for pdf_file_ in os.listdir("csv_to_pdf_files"):
                os.remove(os.path.join("csv_to_pdf_files", pdf_file_))
            os.rmdir("csv_to_pdf_files")
        # Si la carpeta csv_files existe, elimina todos los archivos CSV y luego elimina la carpeta csv_files
        if os.path.exists("pdf_docx_files"):
            for pdf_file_ in os.listdir("csv_pdf_files"):
                a = open(os.path.join("csv_pdf_files", pdf_file_), "rb")
                a = a.close()

                os.remove(os.path.join("csv_pdf_files", pdf_file_))
            os.rmdir("csv_pdf_files")
    deleteFiles()
    # Crea una carpeta para almacenar el PDF recibido
    os.mkdir("csv_to_pdf_files")
    # Guardar el PDF recibido en el directorio pdf_files
    file_name = os.path.join("csv_to_pdf_files", file.filename)
    # Abre el archivo PDF en modo binario y lo guarda en el directorio pdf_files
    with open(file_name, "wb") as buffer:
        # Lee el archivo PDF en trozos de 1024 bytes
        while chunk := await file.read(1024):
            # Escribe el trozo de bytes en el archivo PDF
            buffer.write(chunk)
    os.makedirs("csv_pdf_files", exist_ok=True)
    cv = Converter(file_name)
    cv.convert("csv_pdf_files/converted.doc", start=0, end=None)
    cv.close()
    zip_file_name = "pdf_files.zip"
    with zipfile.ZipFile(zip_file_name, "w") as zip_file:
        # Agrega todos los archivos DOC a la carpeta csv_pdf_files
        for pdf_file_name in os.listdir("csv_pdf_files"):
            csv_file = os.path.join("csv_pdf_files", pdf_file_name)
            zip_file.write(csv_file, pdf_file_name)
    # Mueve el archivo zip a la carpeta csv_files
    os.rename("pdf_files.zip", "csv_pdf_files/pdf_files.zip")
    data = open("csv_pdf_files/pdf_files.zip", "rb")
    file_contents = data.read()
    data.close()
    response = StreamingResponse(BytesIO(file_contents), media_type="application/doc", headers={"Content-Disposition": "attachment; filename=csv_files.zip"})
    deleteFiles()
    return response


# Final conversor de PDF


@app.post("/api/v1/convert-csv-to-pdf")
async def convert_pdf_doc(file: UploadFile = File(...)):
    # imports
    import csv
    from fpdf import FPDF
    def deleteFiles():
        # Si la carpeta pdf_files existe, elimina todos los archivos PDF y luego elimina la carpeta pdf_files
        if os.path.exists("csv_to_pdf_files"):
            for pdf_file_ in os.listdir("csv_to_pdf_files"):
                os.remove(os.path.join("csv_to_pdf_files", pdf_file_))
            os.rmdir("csv_to_pdf_files")
        # Si la carpeta csv_files existe, elimina todos los archivos CSV y luego elimina la carpeta csv_files
        if os.path.exists("csv_pdf_files"):
            for pdf_file_ in os.listdir("csv_pdf_files"):
                a = open(os.path.join("csv_pdf_files", pdf_file_), "rb")
                a = a.close()

                os.remove(os.path.join("csv_pdf_files", pdf_file_))
            os.rmdir("csv_pdf_files")
    deleteFiles()
    # Crea una carpeta para almacenar el PDF recibido
    os.mkdir("csv_to_pdf_files")
    # Guardar el PDF recibido en el directorio pdf_files
    file_name = os.path.join("csv_to_pdf_files", file.filename)
    # Abre el archivo PDF en modo binario y lo guarda en el directorio pdf_files
    with open(file_name, "wb") as buffer:
        # Lee el archivo PDF en trozos de 1024 bytes
        while chunk := await file.read(1024):
            # Escribe el trozo de bytes en el archivo PDF
            buffer.write(chunk)
    os.makedirs("csv_pdf_files", exist_ok=True)
    
    # Crea un objeto PDF
    pdf = FPDF()

    # Añade una página
    pdf.add_page()

    # Define el tamaño y la fuente del texto
    pdf.set_font("Arial", size=12)

    # Abre el archivo CSV y lee los datos
    with open(file_name, "r") as csv_file:
        csv_reader = csv.reader(csv_file)

        # Itera por cada fila y coloca los datos en el PDF
        for row in csv_reader:
            for item in row:
                pdf.cell(40, 10, str(item))
            pdf.ln()

    # Guarda el archivo PDF
    pdf.output(os.path.join("csv_pdf_files", "archivo.pdf"))
    
    zip_file_name = "pdf_files.zip"
    with zipfile.ZipFile(zip_file_name, "w") as zip_file:
        # Agrega todos los archivos DOC a la carpeta csv_pdf_files
        for pdf_file_name in os.listdir("csv_pdf_files"):
            csv_file = os.path.join("csv_pdf_files", pdf_file_name)
            zip_file.write(csv_file, pdf_file_name)
    # Mueve el archivo zip a la carpeta csv_files
    os.rename("pdf_files.zip", "csv_pdf_files/pdf_files.zip")
    data = open("csv_pdf_files/pdf_files.zip", "rb")
    file_contents = data.read()
    data.close()
    response = StreamingResponse(BytesIO(file_contents), media_type="application/pdf", headers={"Content-Disposition": "attachment; filename=csv_files.zip"})
    deleteFiles()
    return response