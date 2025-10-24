import streamlit as st
import json
import re
import os
import io
import math
import matplotlib.pyplot as plt
import geopandas as gpd
import pandas as pd
import numpy as np
from shapely.geometry import Point
import contextily as cx
from matplotlib.ticker import FuncFormatter
from datetime import datetime
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Cm
from staticmap import StaticMap, CircleMarker
from typing import List, Union

def obtener_template_path(cantidadObjetos: int) -> str:
    """
    Retorna el path del template a usar basado en la cantidad de objetos.
    Ejemplo: 1 → 'templateTermoN1.docx'
             2 → 'templateTermoN2.docx'
             3 → 'templateTermoN3.docx'
             4 → 'templateTermoN4.docx'
    """
    
    nombre_template = f"templateTermoN{cantidadObjetos}.docx"
    return os.path.join('templates', nombre_template)

def safe_float_convert(key):
    value = st.session_state.data.get(key)
    if value is None:
        # Si es None (campo vacío), puedes elegir lanzar una excepción
        # o devolver 0.0, pero dado que ya validaste con 'is None', 
        # en teoría no debería llegar aquí si falta un campo crítico.
        raise ValueError(f"Campo crítico vacío: {key}") 
                
    # Si el valor ya es un número (float o int), simplemente lo devuelve
    if isinstance(value, (int, float)):
        return float(value)
            
    # Si es una cadena, reemplaza la coma por punto y convierte a float
    return float(str(value).replace(',', '.'))

    
def clasificar_delta(valor_Delta: float, prom_Temperatura: float) -> str:
    """
    Clasifica el valor del delta según los rangos especificados.
    Los primeros dos casos priorizan valor_Delta independientemente de la temperatura.
    """
    # PRIORIDAD 1: Rangos basados solo en valor_Delta
    if valor_Delta > 0 and valor_Delta < 4:
        return ["Posible deficiencia", "Se requiere más información"]
    
    if valor_Delta >= 4 and valor_Delta <= 15:
        return ["Probable deficiencia", "Reparar en la próxima parada disponible"]
    
    # PRIORIDAD 2: Rangos que consideran tanto valor_Delta como prom_Temperatura
    if valor_Delta > 15 and prom_Temperatura >= 21 and prom_Temperatura <= 40:
        return ["Deficiencia", "Reparar tan pronto como sea posible"]
    
    if valor_Delta > 15 and prom_Temperatura > 40:
        return ["Deficiencia mayor", "Reparar inmediatamente"]
    
    # Caso por defecto (valores fuera de los rangos esperados)
    return ["Sin clasificación", "Verificar datos ingresados"]

def get_map_png_bytes(lon, lat, buffer_m=300, width_px=900, height_px=700, zoom=17):
    """
    Genera un PNG (bytes) de un mapa satelital con marcador en (lon, lat).
    - buffer_m: radio en metros alrededor del punto (controla "zoom").
    - zoom: nivel de teselas (18-19 suele ser bueno).
    """
    # Crear punto y reproyectar a Web Mercator
    gdf = gpd.GeoDataFrame(geometry=[Point(lon, lat)], crs="EPSG:4326").to_crs(epsg=3857)
    pt = gdf.geometry.iloc[0]
    
    # Calcular bounding box
    bbox = (pt.x - buffer_m, pt.y - buffer_m, pt.x + buffer_m, pt.y + buffer_m)

    # Crear figura
    fig, ax = plt.subplots(figsize=(width_px/100, height_px/100), dpi=100)
    ax.set_xlim(bbox[0], bbox[2])
    ax.set_ylim(bbox[1], bbox[3])

    # Añadir basemap (Esri World Imagery)
    cx.add_basemap(ax, source=cx.providers.Esri.WorldImagery, crs="EPSG:3857", zoom=zoom)

    # Dibujar marcador
    gdf.plot(ax=ax, markersize=40, color="red")

    ax.set_axis_off()
    plt.tight_layout(pad=0)

    # Guardar a buffer en memoria
    buf = io.BytesIO()
    plt.savefig(buf, format="png", bbox_inches="tight", pad_inches=0)
    plt.close(fig)
    buf.seek(0)
    return buf.getvalue()


def convertir_a_mayusculas(data):
    if isinstance(data, str):
        return data.upper()
    elif isinstance(data, dict):
        return {k: convertir_a_mayusculas(v) for k, v in data.items()}
    elif isinstance(data, list):
        return [convertir_a_mayusculas(v) for v in data]
    elif isinstance(data, tuple):
        return tuple(convertir_a_mayusculas(v) for v in data)
    else:
        return data  # cualquier otro tipo se deja igual

# Inicialización de estado
if 'step' not in st.session_state:
    st.session_state.step = 1
    st.session_state.data = {}

st.title("Formulario de Termografía - Word Automatizado")

# Funciones de navegación
def next_step():
    missing = [k for k, v in st.session_state.data.items() if v is None or v == ""]
    if missing:
        st.error("Por favor completa todos los campos antes de continuar.")
    else:
        st.session_state.step += 1
        st.rerun()

def prev_step():
    if st.session_state.step > 1:
        st.session_state.step -= 1
        st.rerun()

# Paso 1: Información General
if st.session_state.step == 1:
    st.header("Paso 1: Información General")
    st.session_state.data['nombreProyecto'] = st.text_input("Nombre del Proyecto", key='nombreProyecto')
    st.session_state.data['nombreCiudadoMunicipio'] = st.text_input("Ciudad o Municipio", key='ciudad')
    st.session_state.data['nombreDepartamento'] = st.text_input("Departamento", key='departamento')
    st.session_state.data['tipoCoordenada'] = st.selectbox(f"Tipo de Imagen para las Coordenadas", ["Urbano", "Rural"], key=f'tipo_coordenada')
    st.session_state.data['nombreCompleto'] = st.text_input("Nombre Completo", key='nombre')
    st.session_state.data['nroConteoTarjeta'] = st.text_input("Número de CONTE o Tarjeta Profesional", key='conte_tarjeta')
    st.session_state.data['nombreCargo'] = st.text_input("Nombre del Cargo", key='cargo')
    st.session_state.data['fechaCreacionSinFormato'] = st.date_input("Fecha de Creación", key='fecha_creacion', value=datetime.now())
    st.session_state.data['fechaCreacion'] = st.session_state.data['fechaCreacionSinFormato'].strftime("%Y-%m-%d")
    st.session_state.data['fechaImagenSinFormato'] = st.date_input("Fecha de Imágenes", key='fecha_imagen', value=datetime.now())
    st.session_state.data['fechaImagen'] = st.session_state.data['fechaImagenSinFormato'].strftime("%Y-%m-%d")
    st.session_state.data['direccionProyecto'] = st.text_input("Dirección", key='direccion')
    st.session_state.data['cantidadObjetos'] = st.selectbox("Cantidad de Objetos", [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20], key='cantidad_perfiles')
    st.session_state.data['latitud'] = st.number_input("Latitud", key='latitud', format="%.6f")
    st.session_state.data['longitud'] = st.number_input("Longitud", key='longitud', format="%.6f")
    #Agregar el campo de selección de Rural o Urbano para la generación de la imagen 

    cols = st.columns([1,1])
    if cols[1].button("Siguiente"):
        next_step()

# Paso 2: Datos Técnicos
elif st.session_state.step == 2:
    st.header("Paso 2: Datos Técnicos de los Objetos")
    
    datos_Sin_Mayuscula = st.session_state.data.copy()
    datos = convertir_a_mayusculas(datos_Sin_Mayuscula)
    
    cantidad_objetos = int(st.session_state.data['cantidadObjetos'])
    
    template_path = obtener_template_path(cantidad_objetos)
    
    try:
        
        st.session_state.doc = DocxTemplate(template_path)
        
    except FileNotFoundError:
        
        st.error(f"No se encontró la plantilla: {template_path}")


    for i in range(1, cantidad_objetos + 1):
        suf = f"N{i}"

        with st.expander(f"Termografía - Objeto #{i}", expanded=True):
            st.markdown("---")

            # --- FILA 1: Datos principales ---
            col1, col2, col3 = st.columns(3)
            with col1:
                st.subheader(f"Equipo {suf}")
                st.session_state.data[f'equipoEvaluado{suf}'] = st.text_input(
                    f"Equipo Evaluado {suf}", key=f'equipoEvaluado{suf}'
                )
            with col2:
                st.subheader(f"Marca {suf}")
                st.session_state.data[f'marcaEquipoEvaluado{suf}'] = st.text_input(
                    f"Marca del Equipo {suf}", key=f'marcaEquipoEvaluado{suf}'
                )
                if st.session_state.data[f'marcaEquipoEvaluado{suf}'] == "":
                    st.session_state.data[f'marcaEquipoEvaluado{suf}'] = 'N/A'
            with col3:
                st.subheader(f"Objeto {suf}")
                st.session_state.data[f'objetoEquipoEvaluado{suf}'] = st.text_input(
                    f"Objeto Evaluado {suf}", key=f'objetoEquipoEvaluado{suf}'
                )

            # --- FILA 2: Imágenes ---
            col1, col2 = st.columns(2)
            with col1:
                st.subheader(f"Imagen Termográfica {suf}")
                key_ImgTermo = f'imgTermografica{suf}'
                st.session_state.data[key_ImgTermo] = st.file_uploader(
                    f"Seleccione la Imagen Termográfica {suf}",
                    type=['png', 'jpg', 'jpeg'],
                    key=key_ImgTermo
                )
                if st.session_state.data[key_ImgTermo] is None:
                    st.warning(f"Por favor, suba la Imagen Termográfica {suf} para continuar.")
                    continue
                else:
                    st.success(f"Imagen Termográfica {suf} cargada correctamente.")

            with col2:
                st.subheader(f"Imagen del Espacio {suf}")
                key_ImgEsp = f'imgEspacio{suf}'
                st.session_state.data[key_ImgEsp] = st.file_uploader(
                    f"Seleccione la Imagen del Espacio {suf}",
                    type=['png', 'jpg', 'jpeg'],
                    key=key_ImgEsp
                )
                if st.session_state.data[key_ImgEsp] is None:
                    st.warning(f"Por favor, suba la Imagen del Espacio {suf} para continuar.")
                    continue
                else:
                    st.success(f"Imagen del Espacio {suf} cargada correctamente.")

            # --- FILA 3: Temperaturas principales ---
            st.subheader(f"Análisis Termográfico {suf}")
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.session_state.data[f'tempMaxImgTermo{suf}'] = st.number_input(
                    f"Temperatura Máxima {suf} [°C]", key=f'tempMaxImgTermo{suf}', min_value=0.0, format="%.2f"
                )
            with col2:
                st.session_state.data[f'tempMinImgTermo{suf}'] = st.number_input(
                    f"Temperatura Mínima {suf} [°C]", key=f'tempMinImgTermo{suf}', min_value=0.0, format="%.2f"
                )
            with col3:
                st.session_state.data[f'tempPromImgTermo{suf}'] = st.number_input(
                    f"Temperatura Promedio {suf} [°C]", key=f'tempPromImgTermo{suf}', format="%.2f"
                )
            with col4:
                st.session_state.data[f'emisividadImgTermo{suf}'] = st.number_input(
                    f"Emisividad {suf}", key=f'emisividadImgTermo{suf}', min_value=0.0, format="%.2f"
                )

            # --- FILA 4: Otros análisis termográficos ---
            col1, col2, col3 = st.columns(3)
            with col1:
                st.session_state.data[f'bgTemp{suf}'] = st.number_input(
                    f"BG Temp {suf} [°C]", key=f'bgTemp{suf}', min_value=0.0, format="%.2f"
                )
                if st.session_state.data[f'bgTemp{suf}'] == 0.0:
                    st.session_state.data[f'bgTemp{suf}'] = 'N/A'
            with col2:
                st.session_state.data[f'desvEst{suf}'] = st.number_input(
                    f"Desviación Estándar {suf}", key=f'desvEst{suf}', min_value=0.0, format="%.2f"
                )
                if st.session_state.data[f'desvEst{suf}'] == 0.0:
                    st.session_state.data[f'desvEst{suf}'] = 'N/A'
            with col3:
                st.session_state.data[f'deltaT{suf}'] = st.number_input(
                    f"Delta T {suf}", key=f'deltaT{suf}', min_value=0.0, format="%.2f"
                )
                if st.session_state.data[f'deltaT{suf}'] == 0.0:
                    st.session_state.data[f'deltaT{suf}'] = 'N/A'

            # --- FILA 5: Temperaturas de fase ---
            col1, col2, col3 = st.columns(3)
            with col1:
                st.session_state.data[f'tfaseR{suf}'] = st.number_input(
                    f"T-FASE R {suf} [°C]", key=f'tfaseR{suf}', format="%.2f"
                )
            with col2:
                st.session_state.data[f'tfaseS{suf}'] = st.number_input(
                    f"T-FASE S {suf} [°C]", key=f'tfaseS{suf}', format="%.2f"
                )
            with col3:
                st.session_state.data[f'tfaseT{suf}'] = st.number_input(
                    f"T-FASE T {suf} [°C]", key=f'tfaseT{suf}', format="%.2f"
                )
                

            # --- FILA 6: Otros datos finales ---
            col1, col2 = st.columns(2)
            with col1:
                st.session_state.data[f'tempFondo{suf}'] = st.number_input(
                    f"Temperatura de Fondo {suf} [°C]", key=f'tempFondo{suf}', min_value=0.0, format="%.2f"
                )
            with col2:
                st.session_state.data[f'conclusiones{suf}'] = st.text_area(
                    f"Conclusiones de la Tabla {suf}", key=f'conclusiones{suf}'
                )

            st.markdown("---")
    
    
    cols = st.columns(1)
    if cols[0].button("Finalizar Formulario y Generar Word"):
        
        cantidad_objetos = int(st.session_state.data['cantidadObjetos'])
        
        datos = convertir_a_mayusculas(st.session_state.data.copy())
            
        for i in range(1, cantidad_objetos + 1):
            suf = f"N{i}"
                
            key_ImgTermo = f'imgTermografica{suf}'
            buf_ImgTermo = io.BytesIO(st.session_state.data[key_ImgTermo].read()) if st.session_state.data[key_ImgTermo] else None
            buf_ImgTermo.seek(0)
            datos[key_ImgTermo] = InlineImage(st.session_state.doc, buf_ImgTermo, Cm(7.5), Cm(6.5))
                
            key_ImgEsp = f'imgEspacio{suf}'
            buf_ImgEsp = io.BytesIO(st.session_state.data[key_ImgEsp].read()) if st.session_state.data[key_ImgEsp] else None
            buf_ImgEsp.seek(0)
            datos[key_ImgEsp] = InlineImage(st.session_state.doc, buf_ImgEsp, Cm(7.5), Cm(6.5))
            
            st.session_state.data[f'tNfaseR{suf}'] = st.session_state.data.get(f'tfaseR{suf}')
            st.session_state.data[f'tNfaseS{suf}'] = st.session_state.data.get(f'tfaseS{suf}')
            st.session_state.data[f'tNfaseT{suf}'] = st.session_state.data.get(f'tfaseT{suf}')
            st.session_state.data[f'tempNPromImgTermo{suf}'] = st.session_state.data.get(f'tempPromImgTermo{suf}')
                
            st.session_state.data[f'valNumResDeltaRs{suf}'] = float(st.session_state.data.get(f'tNfaseR{suf}')) - float(st.session_state.data.get(f'tNfaseS{suf}'))
            st.session_state.data[f'valNumResDeltaSt{suf}'] = float(st.session_state.data.get(f'tNfaseS{suf}')) - float(st.session_state.data.get(f'tNfaseT{suf}'))
            st.session_state.data[f'valNumResDeltaTr{suf}'] = float(st.session_state.data.get(f'tNfaseT{suf}')) - float(st.session_state.data.get(f'tNfaseR{suf}'))
                
            st.session_state.data[f'valNumDeltaRs{suf}'] = round(abs(st.session_state.data.get(f'valNumResDeltaRs{suf}')), 2)
            st.session_state.data[f'valNumDeltaSt{suf}'] = round(abs(st.session_state.data.get(f'valNumResDeltaSt{suf}')), 2)
            st.session_state.data[f'valNumDeltaTr{suf}'] = round(abs(st.session_state.data.get(f'valNumResDeltaTr{suf}')), 2)
            
            st.session_state.data[f'clasificacionDeltaRs{suf}'] = clasificar_delta(float(st.session_state.data.get(f'valNumDeltaRs{suf}')), float(st.session_state.data.get(f'tempNPromImgTermo{suf}')))[0]
            st.session_state.data[f'clasificacionDeltaSt{suf}'] = clasificar_delta(float(st.session_state.data.get(f'valNumDeltaSt{suf}')), float(st.session_state.data.get(f'tempNPromImgTermo{suf}')))[0]
            st.session_state.data[f'clasificacionDeltaTr{suf}'] = clasificar_delta(float(st.session_state.data.get(f'valNumDeltaTr{suf}')), float(st.session_state.data.get(f'tempNPromImgTermo{suf}')))[0]

            st.session_state.data[f'accionDeltaRs{suf}'] = clasificar_delta(float(st.session_state.data.get(f'valNumDeltaRs{suf}')), float(st.session_state.data.get(f'tempNPromImgTermo{suf}')))[1]
            st.session_state.data[f'accionDeltaSt{suf}'] = clasificar_delta(float(st.session_state.data.get(f'valNumDeltaSt{suf}')), float(st.session_state.data.get(f'tempNPromImgTermo{suf}')))[1]
            st.session_state.data[f'accionDeltaTr{suf}'] = clasificar_delta(float(st.session_state.data.get(f'valNumDeltaTr{suf}')), float(st.session_state.data.get(f'tempNPromImgTermo{suf}')))[1]

            st.session_state.data[f'deltaRs{suf}'] = f"{st.session_state.data.get(f'valNumDeltaRs{suf}')} °C ({st.session_state.data.get(f'clasificacionDeltaRs{suf}')} - {st.session_state.data.get(f'accionDeltaRs{suf}')})"
            st.session_state.data[f'deltaSt{suf}'] = f"{st.session_state.data.get(f'valNumDeltaSt{suf}')} °C ({st.session_state.data.get(f'clasificacionDeltaSt{suf}')} - {st.session_state.data.get(f'accionDeltaSt{suf}')})"
            st.session_state.data[f'deltaTr{suf}'] = f"{st.session_state.data.get(f'valNumDeltaTr{suf}')} °C ({st.session_state.data.get(f'clasificacionDeltaTr{suf}')} - {st.session_state.data.get(f'accionDeltaTr{suf}')})"

            
                
                
        todos_los_datos_completos = True
        
        # 1. Bucle de validación y cálculo (Ejecutar solo al presionar el botón)
        for i in range(1, cantidad_objetos + 1):
            suf = f"N{i}"
            
            # Campos críticos que deben ser != None para el cálculo
            campos_criticos = [
                f'tfaseR{suf}', f'tfaseS{suf}', f'tfaseT{suf}', f'tempPromImgTermo{suf}'
            ]
            
            # Validar que los campos críticos no sean None (o 0.0 si no pudiste evitarlo)
            # Usando 'is None' si modificaste los st.number_input (Recomendado)
            if any(st.session_state.data.get(k) is None for k in campos_criticos):
                st.error(f"Faltan valores en las temperaturas de fase o promedio para el Objeto #{i}. Por favor, verifique y complete.")
                todos_los_datos_completos = False
                break # Detiene el bucle en el primer objeto incompleto

        # 2. Generación del Word (SOLO si todos los objetos están completos)
        if todos_los_datos_completos:
            
            try:
                
                ahora = datetime.now()
                meses = ["ENERO","FEBRERO","MARZO","ABRIL","MAYO","JUNIO","JULIO","AGOSTO","SEPTIEMBRE","OCTUBRE","NOVIEMBRE","DICIEMBRE"]
                datos['dia'] = ahora.day
                datos['mes'] = meses[ahora.month-1]
                datos['anio'] = ahora.year
                
                
                
                #st.session_state.doc.render(datos)
                #output_path = f"reporteProtocoloTermografia.docx"
                #st.session_state.doc.save(output_path)
                
                try:
                
                    if st.session_state.data['tipoCoordenada'] == "Urbano":
                    
                        if st.session_state.data['latitud'] and st.session_state.data['longitud']:
                            try:
                                lat = float(str(datos['latitud']).replace(',', '.'))
                                lon = float(str(datos['longitud']).replace(',', '.'))
                                mapa = StaticMap(600, 400)
                                mapa.add_marker(CircleMarker((lon, lat), 'red', 12))
                                img_map = mapa.render()
                                buf_map = io.BytesIO()
                                img_map.save(buf_map, format='PNG')
                                buf_map.seek(0)
                                datos['imgMapsProyecto'] = InlineImage(st.session_state.doc, buf_map, Cm(15), Cm(10))
                            except Exception as e:
                                st.error(f"Coordenadas inválidas para el mapa. {e}")
                        else:
                            st.error("Faltan coordenadas para el mapa.")
                                    
                    else:
                                
                        if st.session_state.data['latitud'] and st.session_state.data['longitud']:
                            try:
                                lat = float(str(st.session_state.data['latitud']).replace(',', '.'))
                                    
                                lon = float(str(st.session_state.data['longitud']).replace(',', '.'))
                                    
                                st.warning(f"Prueba de coordenada en modo rural (latitud): {lat}")
                                st.warning(f"Prueba de coordenada en modo rural (longitud): {lon}")
                                        
                                png_bytes = get_map_png_bytes(lon, lat, buffer_m=300, zoom=17)
                                        
                                buf_map = io.BytesIO(png_bytes)
                                buf_map.seek(0)
                                datos['imgMapsProyecto'] = InlineImage(st.session_state.doc, buf_map, Cm(15), Cm(10))
                            except Exception as e:
                                st.error(f"Coordenadas inválidas para el mapa. {e}")
                        else:
                            st.error("Faltan coordenadas para el mapa.")
                            
                            
                    st.session_state.doc.render(datos)
                    output_path = f"reporteProtocoloTermografia.docx"
                    st.session_state.doc.save(output_path)
                    
                    st.success(f"Documento generado exitosamente: {output_path}")
                    with open(output_path, "rb") as file:
                        btn = st.download_button(
                            label="Descargar Informe Word",
                            data=file,
                            file_name=output_path,
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        )
                        
                except Exception as e:
                    
                    st.error(f"Error al generar el documento: {e}")
                
            except Exception as e:
                st.error(f"Error al generar el documento: {e}")

    
