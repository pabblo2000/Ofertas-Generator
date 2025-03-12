import streamlit as st
import pandas as pd
import time
from io import BytesIO
from docx import Document

def extraer_campos_excel(excel_file):
    # Leemos el Excel sin headers y de la hoja 'Plantilla POST'
    df = pd.read_excel(excel_file, header=None, sheet_name='Plantilla POST')
    
    # Datos generales extraídos del Excel
    oferta_referencia = excel_file.name.split('.')[0]
    nombre_proyecto = df.loc[6, 1]
    fecha_inicio = pd.to_datetime(df.loc[3, 6], format='%d.%m.%Y').strftime('%d/%m/%Y')
    fecha_fin = pd.to_datetime(df.loc[4, 6], format='%d.%m.%Y').strftime('%d/%m/%Y')
    sda = df.loc[7, 1]
    
    if pd.notna(sda):
        nombre_proyecto = f"{nombre_proyecto} ({sda})"
    
    # Extraemos posts
    post_name_list = []
    post_hours_list = []
    post_price_list = []
    for i in range(10, 78):
        if df.loc[i, 6] != 0:
            precio = '{:,.2f}'.format(float(df.loc[i, 6])).replace(',', 'X').replace('.', ',').replace('X', '.')
            post_price_list.append(precio)
            post_hours_list.append(df.loc[i, 3])
            post_name_list.append(df.loc[i, 0])
            
    # Totales
    horas_totales = df.iloc[78, 3]
    importe_total_sin_iva = '{:,.2f}'.format(float(df.loc[6, 6])).replace(',', 'X').replace('.', ',').replace('X', '.')
    importe_total_con_iva = '{:,.2f}'.format(float(df.loc[7, 6])).replace(',', 'X').replace('.', ',').replace('X', '.')
    today = time.strftime("%d/%m/%Y")
    
    # Construimos el diccionario base con los campos generales
    cambios = {
        'oferta_referencia': oferta_referencia,
        'nombre_proyecto': nombre_proyecto,
        'fecha_inicio': fecha_inicio,
        'fecha_fin': fecha_fin,
        'descripcion': '-',  # Valor por defecto, que se podrá modificar
    }
    
    # Agregamos los posts (5 máximo)
    for i in range(5):
        cambios[f"post{i+1}"] = post_name_list[i] if len(post_name_list) > i else '-'
        cambios[f"posth{i+1}"] = post_hours_list[i] if len(post_hours_list) > i else '-'
        cambios[f"postc{i+1}"] = post_price_list[i] if len(post_price_list) > i else '-'
    
    # Totales (se ubicarán debajo de los posts en la app)
    cambios['totalh'] = horas_totales
    cambios['totalsiva'] = importe_total_sin_iva
    cambios['totalciva'] = importe_total_con_iva
    cambios['today'] = today
    return cambios

def replace_text_in_doc(doc, placeholder, replacement):
    # Reemplaza el placeholder en párrafos
    for p in doc.paragraphs:
        if placeholder in p.text:
            inline = p.runs
            for i in range(len(inline)):
                if placeholder in inline[i].text:
                    inline[i].text = inline[i].text.replace(placeholder, str(replacement))
    # Reemplaza en tablas
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if placeholder in cell.text:
                    cell.text = cell.text.replace(placeholder, str(replacement))

# --- Interfaz de la app ---
st.title("Generador de Documentos a partir de Excel y Plantilla Word")

# Selección de archivos
excel_file = st.file_uploader("Selecciona el archivo Excel (.xlsx)", type=["xlsx"])
template_file = st.file_uploader("Selecciona la plantilla Word (.docx)", type=["docx"])

if excel_file is not None:
    # Extraemos los datos del Excel
    extracted = extraer_campos_excel(excel_file)
    
    st.subheader("Campos extraídos (antes de edición)")
    # Mostramos los campos generales (sin posts y totales)
    generales = {k: v for k, v in extracted.items() if not k.startswith("post") and k not in ["totalh", "totalsiva", "totalciva"]}
    df_generales = pd.DataFrame(list(generales.items()), columns=["Campo", "Valor"])
    st.table(df_generales)
    
    # Posts: sólo se muestran aquellos cuyo nombre no es "-"
    posts_data = []
    for i in range(1, 6):
        if extracted.get(f"post{i}") != "-":
            posts_data.append({
                "Post": extracted.get(f"post{i}"),
                "Horas": extracted.get(f"posth{i}"),
                "Costo": extracted.get(f"postc{i}")
            })
    if posts_data:
        st.subheader("Posts extraídos (se muestran solo los válidos)")
        df_posts = pd.DataFrame(posts_data)
        st.table(df_posts)
    else:
        st.info("No se han encontrado posts válidos.")
    
    # Totales
    st.subheader("Totales")
    df_totales = pd.DataFrame([
        {"Campo": "Total Horas", "Valor": extracted.get("totalh")},
        {"Campo": "Total sin IVA", "Valor": extracted.get("totalsiva")},
        {"Campo": "Total con IVA", "Valor": extracted.get("totalciva")}
    ])
    st.table(df_totales)
    
    st.markdown("---")
    st.subheader("Editar campos (si es necesario)")
    with st.form("editar_campos"):
        # Campos generales (permitiendo su modificación)
        oferta_referencia = st.text_input("Oferta de Referencia", value=extracted.get("oferta_referencia", ""))
        nombre_proyecto = st.text_input("Nombre del Proyecto", value=extracted.get("nombre_proyecto", ""))
        fecha_inicio = st.text_input("Fecha de Inicio", value=extracted.get("fecha_inicio", ""))
        fecha_fin = st.text_input("Fecha de Fin", value=extracted.get("fecha_fin", ""))
        descripcion = st.text_area("Descripción", value=extracted.get("descripcion", ""))
        
        # Campos nuevos: correo_cliente y correo_proveedor (se ingresan manualmente)
        correo_cliente = st.text_input("Correo Cliente", value="")
        correo_proveedor = st.text_input("Correo Proveedor", value="")
        
        st.markdown("### Posts")
        # Solo se permiten editar los posts que se hayan extraído (es decir, cuyo nombre no sea "-")
        posts = []
        for i in range(1, 6):
            post_val = extracted.get(f"post{i}", "-")
            if post_val != "-":
                post_edit = st.text_input(f"Post {i}", value=post_val)
                posth_edit = st.text_input(f"Horas Post {i}", value=str(extracted.get(f"posth{i}", "-")))
                postc_edit = st.text_input(f"Costo Post {i}", value=str(extracted.get(f"postc{i}", "-")))
                posts.append((post_edit, posth_edit, postc_edit))
        
        st.markdown("### Totales")
        totalh = st.text_input("Total Horas", value=str(extracted.get("totalh", "")))
        totalsiva = st.text_input("Total sin IVA", value=str(extracted.get("totalsiva", "")))
        totalciva = st.text_input("Total con IVA", value=str(extracted.get("totalciva", "")))
        
        submit_button = st.form_submit_button("Actualizar y Generar Documento")
    
    if submit_button:
        # Construimos el diccionario actualizado con los datos editados
        updated = {}
        updated["oferta_referencia"] = oferta_referencia
        updated["nombre_proyecto"] = nombre_proyecto
        updated["fecha_inicio"] = fecha_inicio
        updated["fecha_fin"] = fecha_fin
        updated["descripcion"] = descripcion
        updated["correo_cliente"] = correo_cliente
        updated["correo_proveedor"] = correo_proveedor
        
        # Actualizamos los posts editados
        for i, post_data in enumerate(posts, start=1):
            updated[f"post{i}"] = post_data[0]
            updated[f"posth{i}"] = post_data[1]
            updated[f"postc{i}"] = post_data[2]
        # Para los posts que no se hayan editado, se rellenan con "-"
        for i in range(len(posts)+1, 6):
            updated[f"post{i}"] = "-"
            updated[f"posth{i}"] = "-"
            updated[f"postc{i}"] = "-"
        
        # Totales (ubicados debajo de los posts)
        updated["totalh"] = totalh
        updated["totalsiva"] = totalsiva
        updated["totalciva"] = totalciva
        
        # La fecha de hoy (no editable)
        updated["today"] = extracted.get("today", "")
        
        st.subheader("Campos actualizados")
        df_updated = pd.DataFrame(list(updated.items()), columns=["Campo", "Valor"])
        st.table(df_updated)
        
        if template_file is not None:
            template_file.seek(0)
            doc = Document(template_file)
            # Se reemplazan los placeholders en el documento
            for key, value in updated.items():
                placeholder = f"<<{key}>>"
                replace_text_in_doc(doc, placeholder, value)
            output = BytesIO()
            doc.save(output)
            output.seek(0)
            st.success("Documento generado correctamente.")
            st.download_button(
                label="Descargar documento Word",
                data=output,
                file_name=f"{updated['oferta_referencia']}_generado.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        else:
            st.error("Por favor, selecciona la plantilla Word.")
