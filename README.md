# **README**

## **Clonar el repositorio**

```bash
git clone https://github.com/xOverlap/proyecto-visor.git
```

```bash
cd proyecto-visor
```

## **Instalacion BackEnd**

### Crea un entorno virtual

```bash
python -m venv .venv
```

- **Ahora abra una nueva terminal cmd en la carpeta del proyecto como administrador**

### Activa el entorno virtual

```bash
.\.venv\Scripts\activate.bat
```

### Instalacion de requerimientos
- **Asegurece de tener java 8+ instalado en su computadora y en el path de windows**

```bash
pip install -r requirements.txt
```

### Inicia el servidor de backend

```bash
uvicorn main:app --reload
```

## **Instalación FrontEnd**

- **Abra una nueva terminal cmd en la carpeta del proyecto**

### Entrar a la carpeta del proyecto

```bash
cd FrontEnd
```

### Instalar las dependencias

```bash
npm install
```

### Inicia el servidor de frontend

```bash
npm run dev
```

### **Abrir en el navegador el siguiente sitio web**

- [Sitio Web (http://localhost:5173)](http://localhost:5173)

##### *Ya que en el codigo de python se eliminan archivos y se crean archivos es probable que su antivirus lo reconosca como uno ya que puede ser una conducta similar cree una exepcion sobre el proyecto o deshabilitelo muentras el sitio esté activo*
