# Generador de Ofertas

Generador de Ofertas es una aplicación desarrollada en Streamlit que permite extraer datos de un archivo Excel y generar documentos de ofertas (Word y PDF) a partir de una plantilla. La app ofrece formularios para editar la información extraída y opciones de configuración personalizables para facilitar su uso por usuarios no técnicos.

## Características

- **Extracción de datos:** Lee la hoja "Plantilla POST" de un archivo Excel y extrae campos como Oferta de Referencia, Nombre del Proyecto, Fechas, Posts y Totales.
- **Edición de datos:** Permite modificar la información extraída a través de formularios intuitivos.
- **Generación de documentos:** Reemplaza *placeholders* en una plantilla Word (personalizable o predeterminada) y genera un documento final, que además se convierte a PDF.
- **Modos de guardado:** Opción para guardar los documentos automáticamente en una ubicación específica o para descargarlos en un ZIP.
- **Configuración sencilla:** Panel lateral con opciones para modificar el correo del proveedor, seleccionar el modo de guardado (Mediante descarga o Mediante ubicación) y configurar la plantilla por defecto.
- **Interfaz amigable:** Incluye un logo clicable en el sidebar que redirige a [minsait.com](https://minsait.com), y muestra la versión y autor de la app.

## Requisitos

- Python 3.7 o superior.
- Librerías de Python:
  - streamlit
  - pandas
  - python-docx
  - docx2pdf
  - requests
  - (Otras dependencias que se pueden instalar mediante el archivo `requirements.txt`)

## Instalación

1. Asegúrate de tener Python 3.7 o superior instalado.
2. Instala las dependencias necesarias. Si cuentas con un archivo `requirements.txt`, ejecútalo con:
```bash
pip install -r requirements.txt
```
3. Verifica que en la misma carpeta se encuentren los archivos `app.py` y `config.py`. El archivo `config.py` contiene las configuraciones por defecto (correo del proveedor, modo de guardado, plantilla predeterminada, etc.) que puedes modificar desde la app.

## Ejecución de la Aplicación

Para ejecutar la aplicación, abre una terminal en la carpeta del proyecto y ejecuta:

```bash
streamlit run app.py
```

La aplicación se abrirá en tu navegador predeterminado.

## Configuración

En el panel lateral de la aplicación encontrarás un apartado de **Configuración** donde podrás:
- Modificar el **Correo Proveedor**.
- Elegir el **Modo de guardado** mediante un desplegable:
  - **Mediante descarga:** Los documentos se ofrecen para descarga en un ZIP.
  - **Mediante ubicación:** Los documentos se guardan automáticamente en la carpeta especificada.
- Establecer la **Plantilla por Defecto**.
- (Opcional) Ingresar una **Ubicación de salida** (si se usa el modo "Mediante ubicación").

Después de ajustar la configuración, haz clic en "Guardar Configuración" y reinicia la app para que los cambios tengan efecto.

## Uso de la Aplicación

1. **Carga de archivos:**  
   Sube el archivo Excel (que contenga la hoja "Plantilla POST") y, opcionalmente, una plantilla Word. Si no subes una plantilla, se usará la plantilla por defecto configurada.

2. **Edición de datos:**  
   La aplicación se divide en dos secciones:
   - **Datos Generales:** Aquí se muestran y permiten editar campos como Oferta de Referencia, Nombre del Proyecto, Fechas, Correo Cliente, Correo Proveedor, Today y Descripción.
   - **Posts y Totales:** En esta sección podrás agregar o editar hasta 5 posts (cada uno con sus horas y costo) y modificar los totales.
   
3. **Generación del documento:**  
   Una vez editada la información, haz clic en "Generar Documento". La app reemplazará los *placeholders* en la plantilla Word (por ejemplo, `<<oferta_referencia>>`, `<<post1>>`, etc.) y generará el documento final. Según el modo de guardado seleccionado, el documento se guardará en una ubicación o se ofrecerá para descarga en un ZIP junto con el archivo Excel original.

## Notas

- La aplicación utiliza un sistema de *placeholders* en la plantilla Word para ubicar la información. Asegúrate de que la plantilla contenga dichos *placeholders*.
- Se incluye una barra de progreso durante la generación del documento para informar al usuario sobre el avance del proceso.
- Para facilitar el acceso a usuarios no técnicos, puedes empaquetar la aplicación como un ejecutable (por ejemplo, con PyInstaller) o alojarla en un servidor (como Streamlit Cloud).

## Licencia

Este proyecto se distribuye bajo la licencia MIT.
