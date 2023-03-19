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
            for csv_file_ in os.listdir("pdf_csv_files"):
                a = open(os.path.join("pdf_csv_files", csv_file_), "rb")
                a = a.close()

                os.remove(os.path.join("pdf_csv_files", csv_file_))
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
        for csv_file_name in os.listdir("pdf_csv_files"):
            csv_file = os.path.join("pdf_csv_files", csv_file_name)
            zip_file.write(csv_file, csv_file_name)

    # Elimina la carpeta csv_files y sus contenidos
    for csv_file_name in os.listdir("pdf_csv_files"):
        csv_file = os.path.join("pdf_csv_files", csv_file_name)
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
            for csv_file_ in os.listdir("pdf_docx_files"):
                a = open(os.path.join("pdf_docx_files", csv_file_), "rb")
                a = a.close()

                os.remove(os.path.join("pdf_docx_files", csv_file_))
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
        for csv_file_name in os.listdir("pdf_docx_files"):
            csv_file = os.path.join("pdf_docx_files", csv_file_name)
            zip_file.write(csv_file, csv_file_name)
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
        if os.path.exists("pdf_to_doc_files"):
            for pdf_file_ in os.listdir("pdf_to_doc_files"):
                os.remove(os.path.join("pdf_to_doc_files", pdf_file_))
            os.rmdir("pdf_to_doc_files")
        # Si la carpeta csv_files existe, elimina todos los archivos CSV y luego elimina la carpeta csv_files
        if os.path.exists("pdf_docx_files"):
            for csv_file_ in os.listdir("pdf_doc_files"):
                a = open(os.path.join("pdf_doc_files", csv_file_), "rb")
                a = a.close()

                os.remove(os.path.join("pdf_doc_files", csv_file_))
            os.rmdir("pdf_doc_files")
    deleteFiles()
    # Crea una carpeta para almacenar el PDF recibido
    os.mkdir("pdf_to_doc_files")
    # Guardar el PDF recibido en el directorio pdf_files
    file_name = os.path.join("pdf_to_doc_files", file.filename)
    # Abre el archivo PDF en modo binario y lo guarda en el directorio pdf_files
    with open(file_name, "wb") as buffer:
        # Lee el archivo PDF en trozos de 1024 bytes
        while chunk := await file.read(1024):
            # Escribe el trozo de bytes en el archivo PDF
            buffer.write(chunk)
    os.makedirs("pdf_doc_files", exist_ok=True)
    cv = Converter(file_name)
    cv.convert("pdf_doc_files/converted.doc", start=0, end=None)
    cv.close()
    zip_file_name = "doc_files.zip"
    with zipfile.ZipFile(zip_file_name, "w") as zip_file:
        # Agrega todos los archivos DOC a la carpeta pdf_doc_files
        for csv_file_name in os.listdir("pdf_doc_files"):
            csv_file = os.path.join("pdf_doc_files", csv_file_name)
            zip_file.write(csv_file, csv_file_name)
    # Mueve el archivo zip a la carpeta csv_files
    os.rename("doc_files.zip", "pdf_doc_files/doc_files.zip")
    data = open("pdf_doc_files/doc_files.zip", "rb")
    file_contents = data.read()
    data.close()
    response = StreamingResponse(BytesIO(file_contents), media_type="application/doc", headers={"Content-Disposition": "attachment; filename=csv_files.zip"})
    deleteFiles()
    return response