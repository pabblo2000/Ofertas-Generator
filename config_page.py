import streamlit as st
import time

# Mostrar el logo
st.image("https://pbs.twimg.com/profile_images/1859630278114684929/7BumEThB_200x200.jpg", width=200)
st.title("Bienvenido al Generador de Ofertas")

# Pedir el nombre al usuario
user_name = st.text_input("Dinos su nombre:")

if st.button("Enviar"):
    config_contents = f'''correo_proveedor = ""
nombre = "{user_name}"
default_template = r"default.docx"
selected_docs = ["Word", "PDF"]
enable_advanced_date_fields = True
enable_custom_fields = False
enable_description = True
enable_alcance = True
'''
    try:
        with open("config.py", "w", encoding="utf-8") as file:
            file.write(config_contents)
        st.success("Muchas gracias. Redirigiendo a la aplicación...")
        # Redirigir a la main page
        time.sleep(2)  # Dar tiempo para mostrar el mensaje de éxito

        try:
            st.switch_page("app.py")  # Redirecciona a la página principal
        except AttributeError:
            st.error("No se pudo redirigir a la página principal. Por favor, recargue la página.")
    except Exception as e:
        st.error(f"Error al guardar el archivo de configuración: {e}")