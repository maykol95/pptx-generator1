import streamlit as st
import pandas as pd
import json
import os
import pyodbc
import tempfile
import shutil
import requests
from pptx import Presentation
from pptx.util import Inches, Pt
from datetime import datetime
from PIL import Image

from dotenv import load_dotenv


load_dotenv()  # Cargar variables del archivo .env

# Usar las variables
uid = os.getenv("AWS_UID")
pwd = os.getenv("AWS_PWD")


st.set_page_config(layout="wide")

st.title("Generador de PowerPoint desde SQL Athena")

# Constantes de tamaño máximo para las imágenes
max_width = 4.5  # en pulgadas
max_height = 5.5  # en pulgadas

# Función para calcular dimensiones sin deformar la imagen
def calcular_dimensiones(img_path):
    try:
        with Image.open(img_path) as img:
            width_px, height_px = img.size
            # Convertir los píxeles a pulgadas (asumiendo una resolución de 96 píxeles por pulgada)
            width_in, height_in = width_px / 96, height_px / 96
            # Calculamos el ratio de escala para que la imagen no se deforme
            ratio = min(max_width / width_in, max_height / height_in)
            return Inches(width_in * ratio), Inches(height_in * ratio)
    except Exception as e:
        print(f"Error al procesar {img_path}: {e}")
        return None, None

# Paso 1: Input de la consulta SQL
consulta_usuario = st.text_area("Escribe tu consulta SQL para Athena")

if st.button("Ejecutar consulta y guardar JSON"):
    try:
        conn = pyodbc.connect(f'DSN=Athena-livetradeBI;UID={uid};PWD={pwd}')
        df = pd.read_sql(consulta_usuario, conn)
        st.dataframe(df)
        json_path = "resultado_consulta.json"
        df.to_json(json_path, orient="records")
        st.success(f"Consulta exitosa. Datos guardados en {json_path}")
    except Exception as e:
        st.error(f"Error en la consulta: {e}")

# Paso 2: Cargar JSON y configurar filtros
if os.path.exists("resultado_consulta.json"):
    df = pd.read_json("resultado_consulta.json")
    st.subheader("Filtrar datos para presentación")

    columnas = df.columns.tolist()
    filtros = {}

    if 'fecha' in df.columns:
        try:
            df['fecha'] = pd.to_datetime(df['fecha'].astype(str), format='%Y-%m-%d', errors='coerce')
            df = df.dropna(subset=['fecha'])
            if not df['fecha'].isna().all():
                fecha_min = df['fecha'].min().date()
                fecha_max = df['fecha'].max().date()
                fecha_inicio, fecha_fin = st.date_input("Selecciona el rango de fechas", value=(fecha_min, fecha_max), min_value=fecha_min, max_value=fecha_max)
                df = df[(df['fecha'] >= pd.to_datetime(fecha_inicio)) & (df['fecha'] <= pd.to_datetime(fecha_fin))]
            else:
                st.warning("No hay fechas válidas después de la conversión.")
        except Exception as e:
            st.warning(f"No se pudo interpretar la columna 'fecha': {e}")

    # Filtros dinámicos según tipo
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
                except ValueError:
                    st.warning(f"Valores no válidos en {col}")
        elif pd.api.types.is_bool_dtype(dtype):
            seleccion = st.multiselect(f"Filtrar por {col}", [True, False])
            if seleccion:
                filtros[col] = seleccion

    for col, valores in filtros.items():
        df = df[df[col].isin(valores)]

    st.write(f"Filas después del filtrado: {len(df)}")

    st.subheader("Configuración de diapositivas")
    encabezados_seleccionados = st.multiselect("Selecciona columnas como encabezados de imagen", columnas)
    fotos_por_slide = st.number_input("¿Cuántas fotos por diapositiva?", min_value=1, max_value=4, value=1, step=1)
    orden_columna = st.selectbox("Selecciona columna para ordenar las imágenes (opcional)", [""] + columnas)

    if st.button("Generar PowerPoint"):
        try:
            if orden_columna:
                df = df.sort_values(by=orden_columna)

            prs = Presentation()
            prs.slide_width = Inches(13.33)
            prs.slide_height = Inches(7.5)
            blank_slide_layout = prs.slide_layouts[6]

            temp_dir = tempfile.mkdtemp()
            imagen_col = None

            for col in df.columns:
                if df[col].astype(str).str.contains("photogram-livetrade-prod.s3.amazonaws.com").any():
                    imagen_col = col
                    break

            if not imagen_col:
                st.warning("No se encontró ninguna columna con URLs de imagen válidas.")
            else:
                filas_validas = []
                for i, row in df.iterrows():
                    url = row[imagen_col]
                    try:
                        response = requests.get(url, stream=True, timeout=5)
                        if response.status_code == 200:
                            img_path = os.path.join(temp_dir, f"img_{i}.jpg")
                            with open(img_path, 'wb') as f:
                                f.write(response.content)
                            if os.path.getsize(img_path) == 0:
                                continue
                            row['img_path'] = img_path
                            filas_validas.append(row)
                    except Exception:
                        continue

                for i in range(0, len(filas_validas), fotos_por_slide):
                    slide = prs.slides.add_slide(blank_slide_layout)
                    subset = filas_validas[i:i+fotos_por_slide]
                    spacing_x = Inches(13.33 / fotos_por_slide)
                    for j, row in enumerate(subset):
                        img_path = row['img_path']
                        img_width, img_height = calcular_dimensiones(img_path)
                        img_left = spacing_x * j + Inches(0.25)
                        img_top = Inches(1.2)

                        if img_width and img_height:
                            slide.shapes.add_picture(img_path, img_left, img_top, width=img_width, height=img_height)

                            encabezado = " | ".join(f"{col}: {row.get(col, '')}" for col in encabezados_seleccionados)
                            text_box = slide.shapes.add_textbox(img_left, Inches(0.2), img_width, Inches(1))
                            text_frame = text_box.text_frame

                            # Añadir encabezados con múltiples párrafos y tamaño de fuente 9
                            lines = encabezado.split(" | ")
                            for line in lines:
                                p = text_frame.add_paragraph()
                                p.text = line
                                p.font.size = Pt(9)  # Ajuste del tamaño de la fuente

                pptx_path = os.path.join(temp_dir, "presentacion_generada.pptx")
                prs.save(pptx_path)

                with open(pptx_path, "rb") as f:
                    st.download_button("Descargar PowerPoint", f, file_name="presentacion_generada.pptx")

                shutil.rmtree(temp_dir)

        except Exception as e:
            st.error(f"Error al generar presentación: {e}")
