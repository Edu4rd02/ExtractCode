# ExtractCode

Este proyecto permite extraer la estructura y el código fuente de un directorio y generar un documento de Word (DOCX) formateado.
## Requisitos previos
- **Python 3.x** instalado en tu sistema.
- Entorno de Windows (requerido para la interacción con `win32com.client` al actualizar la tabla de contenidos interactiva).

## Instrucciones de instalación y uso

### 1. Clonar el repositorio
Clona el repositorio en tu máquina local y navega hasta el directorio del proyecto:
```bash
git clone https://github.com/Edu4rd02/ExtractCode.git
cd ExtractCode
```

### 2. Crear e iniciar un entorno virtual
Es una buena práctica utilizar un entorno virtual para aislar las dependencias del proyecto. Para crear un entorno virtual (puedes nombrarlo `.env` o `venv`) y activarlo, ejecuta los siguientes comandos:

**En Windows (CMD / PowerShell):**
```bash
python -m venv venv
venv\Scripts\activate
```

### 3. Instalar las dependencias
Una vez activado el entorno virtual, instala las bibliotecas requeridas desde el archivo `requirements.txt`:
```bash
pip install -r requirements.txt
```

### 4. Ejecutar el proyecto
Con las dependencias instaladas, puedes iniciar la aplicación. El archivo principal se encuentra dentro de la carpeta `src`.

Navega a la carpeta `src` y ejecuta el script principal:
```bash
cd src
python main.py
```
*(También puedes ejecutarlo directamente desde la raíz del proyecto usando `python src/main.py`)*

Al ejecutarse, se abrirá un cuadro de diálogo (interfaz gráfica) solicitando que selecciones la carpeta del proyecto del que deseas extraer el código. Una vez finalizado el proceso, se generará el archivo `.docx` con el código formateado en el directorio desde donde ejecutaste el script.