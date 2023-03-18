from io import BytesIO
from fastapi import FastAPI, File, UploadFile
from fastapi.responses import StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
import os, zipfile, tabula

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
async def convert(file: UploadFile = File(...)):
    def deleteFiles():
        if os.path.exists("pdf_files"):
            for pdf_file_ in os.listdir("pdf_files"):
                os.remove(os.path.join("pdf_files", pdf_file_))
            os.rmdir("pdf_files")
            
        if os.path.exists("csv_files"):
            for csv_file_ in os.listdir("csv_files"):
                a = open(os.path.join("csv_files", csv_file_), "rb")
                a = a.close()

                os.remove(os.path.join("csv_files", csv_file_))
            os.rmdir("csv_files")
    deleteFiles()
    os.mkdir("pdf_files")
    # Guardar el PDF recibido en el directorio pdf_files
    file_name = os.path.join("pdf_files", file.filename)
    with open(file_name, "wb") as buffer:
        while chunk := await file.read(1024):
            buffer.write(chunk)
    # Especifica la ruta de tu archivo PDF
    file = file_name

    # Extrae todas las tablas del PDF
    all_tables = tabula.read_pdf(file, pages="all", multiple_tables=True)

    # Crea una carpeta para almacenar los archivos CSV extraídos
    os.makedirs("csv_files", exist_ok=True)

    # Utiliza un bucle for para ajustar el formato de cada tabla extraída y guardarla como un archivo CSV
    for i, table in enumerate(all_tables):
        # Elimina las filas y columnas vacías
        table.dropna(how="all", inplace=True)
        table.dropna(axis=1, how="all", inplace=True)
        # Si hay más de un archivo CSV, crea un archivo .zip y agrega todos los archivos CSV a él
        if len(all_tables) > 1:
            csv_file = os.path.join("csv_files", f"table_{i}.csv")
        else:
            # Si solo hay un archivo CSV, renombra el archivo con el nombre del PDF
            csv_file = os.path.join("csv_files", os.path.splitext(os.path.basename(file))[0] + ".csv")
        table.to_csv(csv_file, index=False)

    zip_file_name = "csv_files.zip"
    with zipfile.ZipFile(zip_file_name, "w") as zip_file:
        for csv_file_name in os.listdir("csv_files"):
            csv_file = os.path.join("csv_files", csv_file_name)
            zip_file.write(csv_file, csv_file_name)

    # Elimina la carpeta csv_files y sus contenidos
    for csv_file_name in os.listdir("csv_files"):
        csv_file = os.path.join("csv_files", csv_file_name)
        os.remove(csv_file)

    # Mueve el archivo zip a la carpeta csv_files
    os.rename("csv_files.zip", "csv_files/csv_files.zip")
    if os.path.exists("csv_files/csv_files.zip"):
        print("complete")
        data = open("csv_files/csv_files.zip", "rb")
        file_contents = data.read()
        data.close()
        response = StreamingResponse(BytesIO(file_contents), media_type="application/zip", headers={"Content-Disposition": "attachment; filename=csv_files.zip"})
        deleteFiles()
        return response