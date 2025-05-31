import streamlit as st
import pandas as pd
import os
import tempfile
import shutil
import requests
from pptx import Presentation
from pptx.util import Inches, Pt
from datetime import datetime
from PIL import Image
from io import BytesIO

st.set_page_config(layout="wide")
st.title("Generador de PowerPoint desde archivo Excel")

# Tamaños máximos para imágenes
def calcular_dimensiones(img_path, fotos_por_slide, encabezados_count):
    try:
        with Image.open(img_path) as img:
            width_px, height_px = img.size
            width_in, height_in = width_px / 96, height_px / 96
            max_width = 13 / fotos_por_slide - 0.5
            encabezado_space = 0.3 + 0.25 * min(encabezados_count, 6)
            max_height = 6.8 - encabezado_space
            ratio = min(max_width / width_in, max_height / height_in)
            return Inches(width_in * ratio), Inches(height_in * ratio)
    except Exception as e:
        print(f"Error al procesar {img_path}: {e}")
        return None, None

# Cargar plantilla opcional
st.sidebar.subheader("Plantilla PowerPoint")
template_file = st.sidebar.file_uploader("Cargar archivo .pptx de plantilla (opcional)", type=["pptx"])
template_bytes = template_file.read() if template_file else None

# Cargar archivo Excel
st.subheader("Carga de datos desde Excel")
archivo_excel = st.file_uploader("Cargar archivo Excel (.xlsx)", type=["xlsx"])

if archivo_excel:
    try:
        df = pd.read_excel(archivo_excel)
        df.to_json("resultado_consulta.json", orient="records")  # se guarda igual para mantener compatibilidad
        st.success("Archivo Excel cargado exitosamente.")
        st.dataframe(df.head(100))
    except Exception as e:
        st.error(f"Error al leer el archivo: {e}")

if os.path.exists("resultado_consulta.json"):
    df = pd.read_json("resultado_consulta.json")
    st.subheader("Filtrar datos para presentación")

    columnas = df.columns.tolist()
    filtros = {}

    if 'fecha' in df.columns:
        try:
            df['fecha'] = pd.to_datetime(df['fecha'].astype(str), errors='coerce')
            df = df.dropna(subset=['fecha'])
            fecha_min = df['fecha'].min().date()
            fecha_max = df['fecha'].max().date()
            fecha_inicio, fecha_fin = st.date_input("Selecciona el rango de fechas", value=(fecha_min, fecha_max))
            df = df[(df['fecha'] >= pd.to_datetime(fecha_inicio)) & (df['fecha'] <= pd.to_datetime(fecha_fin))]
        except: pass

    for col in columnas:
        dtype = df[col].dtype
        opciones = df[col].dropna().unique().tolist()
        if dtype == object or dtype.name == 'category':
            seleccion = st.multiselect(f"Filtrar por {col}", sorted(map(str, opciones)))
            if seleccion:
                filtros[col] = seleccion
        elif pd.api.types.is_numeric_dtype(dtype):
            valores_input = st.text_input(f"Valores para {col} (separados por comas)")
            if valores_input:
                try:
                    valores = [float(v.strip()) for v in valores_input.split(',')]
                    filtros[col] = valores
                except: pass
        elif pd.api.types.is_bool_dtype(dtype):
            seleccion = st.multiselect(f"Filtrar por {col}", [True, False])
            if seleccion:
                filtros[col] = seleccion

    for col, valores in filtros.items():
        df = df[df[col].isin(valores)]

    st.write(f"Filas después del filtrado: {len(df)}")

    encabezados_seleccionados = st.multiselect("Selecciona columnas como encabezados de imagen (máx 8)", columnas, max_selections=8)
    fotos_por_slide = st.number_input("¿Cuántas fotos por diapositiva? (máx 4)", min_value=1, max_value=4, value=1, step=1)
    orden_columna = st.selectbox("Ordenar por columna (opcional)", [""] + columnas)
    subdiv_col = st.selectbox("Subdividir por columna (opcional, genera un archivo por valor distinto)", [""] + columnas)

    if st.button("Generar PowerPoint"):
        try:
            if orden_columna:
                df = df.sort_values(by=orden_columna)

            temp_dir = tempfile.mkdtemp()
            imagen_col = None
            for col in df.columns:
                if df[col].astype(str).str.contains("photogram-livetrade-prod.s3.amazonaws.com").any():
                    imagen_col = col
                    break

            if not imagen_col:
                st.warning("No se encontró ninguna columna con URLs de imagen válidas.")
            else:
                df['img_path'] = None
                filas_validas = []
                for i, row in df.iterrows():
                    url = row[imagen_col]
                    try:
                        response = requests.get(url, stream=True, timeout=5)
                        if response.status_code == 200:
                            img_path = os.path.join(temp_dir, f"img_{i}.jpg")
                            with open(img_path, 'wb') as f:
                                f.write(response.content)
                            if os.path.getsize(img_path) > 0:
                                row['img_path'] = img_path
                                filas_validas.append(row)
                    except: continue

                df_filtrado = pd.DataFrame(filas_validas)

                def generar_presentacion(df_slice, nombre):
                    prs = Presentation(BytesIO(template_bytes)) if template_bytes else Presentation()
                    prs.slide_width = Inches(13.33)
                    prs.slide_height = Inches(7.5)
                    layout = prs.slide_layouts[6]

                    for i in range(0, len(df_slice), fotos_por_slide):
                        slide = prs.slides.add_slide(layout)
                        subset = df_slice.iloc[i:i+fotos_por_slide]
                        spacing_x = Inches(13.33 / fotos_por_slide)
                        for j, (_, row) in enumerate(subset.iterrows()):
                            img_path = row['img_path']
                            img_width, img_height = calcular_dimensiones(img_path, fotos_por_slide, len(encabezados_seleccionados))
                            img_left = spacing_x * j + Inches(0.25)
                            encabezado_height = Inches(0.25 * min(len(encabezados_seleccionados), 6))
                            img_top = encabezado_height + Inches(0.4)

                            if img_width and img_height:
                                slide.shapes.add_picture(img_path, img_left, img_top, width=img_width, height=img_height)
                                encabezado = " | ".join(f"{col}: {row.get(col, '')}" for col in encabezados_seleccionados)
                                if encabezado:
                                    text_box = slide.shapes.add_textbox(img_left, Inches(0.2), img_width, encabezado_height)
                                    text_frame = text_box.text_frame
                                    text_frame.clear()
                                    for line in encabezado.split(" | "):
                                        p = text_frame.add_paragraph()
                                        p.text = line
                                        p.font.size = Pt(8)
                                        p.font.bold = True

                    path = os.path.join(temp_dir, f"{nombre}.pptx")
                    prs.save(path)
                    return path

                archivos_generados = []

                if subdiv_col:
                    for val in df_filtrado[subdiv_col].dropna().unique():
                        df_part = df_filtrado[df_filtrado[subdiv_col] == val]
                        pptx_path = generar_presentacion(df_part, str(val))
                        archivos_generados.append((str(val), pptx_path))
                else:
                    pptx_path = generar_presentacion(df_filtrado, "presentacion")
                    archivos_generados.append(("presentacion", pptx_path))

                for nombre, path in archivos_generados:
                    with open(path, "rb") as f:
                        st.download_button(
                            label=f"📥 Descargar {nombre}.pptx",
                            data=f,
                            file_name=f"{nombre}.pptx",
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                        )

                shutil.rmtree(temp_dir)

        except Exception as e:
            st.error(f"Error al generar presentación: {e}")
