# Generador de Ofertas

Generador de Ofertas es una aplicación desarrollada en Streamlit que permite generar documentos de ofertas (Word y PDF) a partir de datos extraídos de un archivo Excel (la hoja "Plantilla POST"). La aplicación permite editar la información extraída, configurar parámetros y generar los documentos utilizando una plantilla.

## Contenido del Proyecto

- **app.py**: Script principal de la aplicación.
- **config.py**: Archivo de configuración con variables de ejemplo:
  - `correo_proveedor = "proveedor@gmail.com"`
  - `modo_guardado = "Mediante descarga"`
  - `default_template = r".\plantilla.docx"`
  - `output_folder = r"C:\Users\TuUsuario\Desktop"`
  - `nombre = "TuNombre"`
- **requirements.txt**: Lista de dependencias necesarias.
- **ins.ps1**: Script de PowerShell que crea el entorno virtual (.venv), instala Python (si es necesario), instala las dependencias y, si falta la variable `nombre` en config.py, solicita su valor. Al finalizar, muestra un mensaje rodeado de asteriscos indicando que la aplicación está instalada.
- **installer.bat**: Archivo por lotes para ejecutar installer.ps1, mostrando la terminal y los outputs.
- **run.ps1**: Script de PowerShell que activa el entorno virtual y ejecuta la aplicación con `python -m streamlit run app.py`.
- **run.bat**: Archivo por lotes para ejecutar run.ps1, mostrando la terminal y los outputs.
- **run_app.vbs**: Script VBScript para ejecutar run_app.ps1 sin mostrar la ventana de la terminal.

## Instalación y Configuración

1. **Distribución**:  
   Empaqueta todo el proyecto (app.py, config.py, requirements.txt, installer.ps1, run_app.ps1, installer.bat, run_app.vbs, etc.) en un archivo ZIP.

2. **Instalación**:
   - Extrae el ZIP en una carpeta en el equipo destino.
   - Haz doble clic en **installer.bat** para iniciar el proceso de instalación. Este script:
     - Verifica si existe el entorno virtual (.venv) y lo crea si es necesario.
     - Activa el entorno virtual y actualiza pip.
     - Instala las dependencias desde requirements.txt.
     - Si la variable `nombre` no está definida en config.py, solicita al usuario que ingrese el valor y lo añade al archivo.
     - Al finalizar, muestra en la consola un mensaje rodeado de asteriscos que dice:  

       ```bash
       **** APLICACION INSTALADA ****
       ```

     - La consola permanecerá abierta para que el usuario pueda ver el progreso.

3. **Ejecución**:

- Una vez instalado, el usuario tiene una forma sencilla de ejecutar la aplicación:
  - Hacer doble clic en **run_app.vbs** para ejecutar la aplicación sin mostrar la terminal.  
     *Importante:* Si se detecta que el entorno virtual no existe, el script indicará que primero se debe ejecutar installer.bat.

## Uso de la Aplicación

- **Carga de Archivos**:  
  Sube el archivo Excel que contenga la hoja "Plantilla POST" y, opcionalmente, una plantilla Word. Si no se sube una plantilla, se usará la plantilla por defecto configurada en config.py.

- **Edición de Datos**:  
  La aplicación está dividida en dos secciones:
  - **Datos Generales**:  
    Permite editar campos como Oferta de Referencia, Nombre del Proyecto, Fechas, Correo Cliente, Correo Proveedor, Today, Descripción y un nuevo campo SDA.  
    El campo SDA se auto-rellena con el valor de la celda B8 del Excel y es editable. El nombre del proyecto se actualiza para incluir el SDA entre paréntesis si se proporciona.
  - **Posts y Totales**:  
    Permite agregar y editar hasta 5 posts (cada uno con sus horas y costo) y modificar los totales.

- **Generación del Documento**:  
  Al hacer clic en "Generar Documento", la aplicación:
  - Reemplaza los placeholders en la plantilla Word (por ejemplo, `<<oferta_referencia>>`, `<<nombre_proyecto>>`, `<<post1>>`, etc.) con los datos ingresados.
  - Realiza algunos ajustes de formato (p.ej., pone en negrita la oferta de referencia, subraya parte del párrafo).
  - Convierte el documento a PDF y lo guarda o lo ofrece para descarga, según el modo de guardado seleccionado en la configuración:
    - **Mediante descarga**: Se genera un ZIP con el documento Word, el PDF y el archivo Excel original.
    - **Mediante ubicación excel**: Se guardan los documentos automáticamente en el directorio actual.

## Requisitos

- Windows (para los scripts .bat y .vbs).
- Conexión a Internet (para instalar dependencias si es la primera vez).
- Python 3.11.9 (o una versión compatible) instalado globalmente para poder crear el entorno virtual.
- Streamlit y demás dependencias definidas en requirements.txt (se instalarán en el entorno virtual).

## Ejecución

- **Instalación**:  
  Haz doble clic en **installer.bat** y sigue las instrucciones en la consola. Al finalizar, verás un mensaje de confirmación.
  
- **Ejecución de la Aplicación**:  
  Para iniciar la aplicación:
  - Haz doble clic en **run_app.bat** para ejecutarla mostrando la terminal.
  - O haz doble clic en **run_app.vbs** para ejecutarla sin mostrar la terminal (si el entorno virtual está creado).

¡Listo! Con estos pasos, la aplicación se instalará y ejecutará de manera sencilla para usuarios no técnicos.
