# Generador de Ofertas

## Resumen

**Generador de Ofertas** es una aplicación web local desarrollada en **Streamlit** que automatiza la creación de documentos de oferta en formato **Word** y **PDF**. La herramienta extrae y procesa datos de un archivo Excel (que debe contener la hoja **"Plantilla POST"**) y permite la edición, personalización y configuración de la información mediante una interfaz intuitiva, garantizando la generación de documentos profesionales y consistentes para la presentación de ofertas comerciales.

## Características Principales

- **Extracción Automatizada de Datos**: Obtiene información clave del archivo Excel para minimizar la intervención manual.
- **Interfaz Intuitiva y Personalizable**: Permite editar datos generales, perfiles (posts), totales y agregar campos personalizados según se requiera.
- **Generación Dual de Documentos**: Produce documentos en **Word** y los convierte a **PDF**, con opción a empaquetar ambos formatos junto al Excel original en un archivo ZIP.
- **Configuración Flexible**: Ajusta parámetros como el correo del proveedor, modo de guardado, rutas de plantillas y ubicación de salida, entre otros.
- **Instalación y Ejecución Simplificadas**: Con tan solo ejecutar el archivo **run_app.vbs**, la aplicación verifica automáticamente la existencia de `config.py` y del entorno virtual `.venv`, iniciándose sin necesidad de pasos adicionales.

## Estructura del Proyecto

- **app.py**: Script principal que ejecuta la aplicación.
- **config.py**: Archivo de configuración con parámetros clave, tales como:
  - `correo_proveedor`
  - `modo_guardado` (opciones: "Mediante descarga" o "Mediante ubicación")
  - `default_template` (ruta de la plantilla de Word)
  - `output_folder` (directorio de salida)
  - `nombre` (nombre del usuario/proveedor)
  - `selected_docs` (por ejemplo: ["Word", "PDF"])
  - `enable_advanced_date_fields` y `enable_custom_fields`
- **requirements.txt**: Lista de dependencias necesarias.
- **ins.ps1**: Script de PowerShell para crear el entorno virtual (.venv) e instalar las dependencias, configurando `config.py` si es necesario.
- **run.ps1**: Script de PowerShell que activa el entorno virtual y ejecuta la aplicación con el comando `python -m streamlit run app.py`.
- **run_app.vbs**: Script VBScript que verifica la existencia de `config.py` y del entorno virtual `.venv`, y lanza la aplicación de forma silenciosa (sin mostrar la terminal).

## Instalación y Configuración

1. **Distribución del Proyecto**:  
   Empaqueta todos los archivos del proyecto (app.py, config.py, requirements.txt, ins.ps1, run.ps1, run_app.vbs, etc.) en un archivo ZIP.

2. **Proceso de Instalación y Ejecución**:
   - Extrae el contenido del ZIP en una carpeta en el equipo destino.
   - **Ejecución Simplificada**:  
     Ya no es necesario utilizar archivos Batch (.bat). Simplemente haz doble clic en el archivo **run_app.vbs**. Este script se encargará de:
     - Verificar la existencia de **config.py** y del entorno virtual **.venv**.
     - Ejecutar el proceso de instalación (si falta alguno de estos elementos) creando el entorno virtual e inicializando **config.py** según corresponda.
     - Iniciar la aplicación en **Streamlit** de forma automática y silenciosa.
   - Durante el proceso, se mostrará un mensaje de confirmación similar a:

     ```bash
     **** APLICACION INSTALADA ****
     ```

## Uso de la Aplicación

1. **Carga de Archivos**:
   - Sube el archivo Excel que contenga la hoja **"Plantilla POST"**.
   - Opcionalmente, carga una plantilla de Word. Si no se proporciona, se utilizará la plantilla por defecto definida en **config.py**.

2. **Edición y Personalización de Datos**:
   - **Datos Generales**:  
     Edita campos como la oferta de referencia, nombre del proyecto (con la opción de incluir el valor SDA extraído de la celda B8), fechas, correos y descripciones.
   - **Posts y Totales**:  
     Permite agregar y modificar hasta **10 perfiles (posts)**, cada uno con sus horas y costos, además de actualizar los totales correspondientes.
   - **Campos Personalizados** (opcional):  
     Si está habilitado, añade campos de texto corto o párrafos para personalizaciones adicionales.

3. **Generación de Documentos**:
   - Al pulsar el botón **"Generar Documento"**, la aplicación:
     - Sustituye los *placeholders* de la plantilla de Word (por ejemplo, `<<oferta_referencia>>`, `<<nombre_proyecto>>`, `<<post1>>`, etc.) por los datos ingresados.
     - Aplica ajustes de formato (como negrita y subrayado en secciones clave).
     - Convierte el documento a PDF (según la configuración) y ofrece dos opciones:
       - **Mediante descarga**: Genera un archivo ZIP que incluye el documento Word, el PDF y el archivo Excel original.
       - **Mediante ubicación**: Guarda automáticamente los documentos en la carpeta especificada en **config.py**.

## Requisitos del Sistema

- **Sistema Operativo**: Windows (compatible con VBScript y PowerShell).
- **Conectividad**: Se requiere conexión a Internet para la instalación inicial de dependencias.
- **Python**: Versión 3.11.9 o compatible, instalada globalmente para facilitar la creación del entorno virtual.
- **Dependencias**: Se instalarán automáticamente mediante **requirements.txt** (incluyendo librerías como Streamlit, pandas, python-docx, docx2pdf, entre otras).

## Ejecución de la Aplicación

Para iniciar la aplicación, simplemente haz doble clic en **run_app.vbs**. Este script se encargará de:

- Verificar la existencia de **config.py** y del entorno virtual **.venv**.
- Ejecutar la aplicación en tu navegador a través de **Streamlit**, sin necesidad de mostrar la terminal.

## Notas y Recomendaciones

- Asegúrate de que el archivo Excel incluya la hoja **"Plantilla POST"** con el formato correcto.
- Verifica que **config.py** contenga todas las variables necesarias para el correcto funcionamiento de la aplicación.
- En caso de problemas con la conversión a PDF, revisa las dependencias instaladas y la versión de Python.
- La aplicación es altamente configurable, lo que permite adaptarla a diferentes flujos de trabajo y requerimientos en la generación de ofertas comerciales.

## Licencia

Este proyecto se distribuye bajo la licencia [especificar licencia aquí]. Consulta el archivo LICENSE para más detalles.

## Contacto y Soporte

Para consultas, sugerencias o reportar incidencias, puedes contactar a <palvaroh@minsait.com>.

---

**Generador de Ofertas** es una solución integral que optimiza y profesionaliza el proceso de creación y gestión de ofertas comerciales.
