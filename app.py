import streamlit as st
import pandas as pd
import os
import unicodedata

# Configuración de la página web
st.set_page_config(page_title="Alerta de Vencimiento de Expedientes", layout="wide")
st.title("📊 Alerta de Vencimiento de Expedientes")

ARCHIVOMAESTRO = "Consolidado_Maestro_Expedientes.xlsx"

# Nombres estándar que usaremos en la App
COLUMNAS_REQUERIDAS = [
    "Expediente", "Fecha Vencimiento", "Remitente", "Asunto", 
    "Alerta Vencimiento", "Nombre Usuario Asignado", "Dias Profesional"
]

# Función para normalizar texto
def normalizar_texto(texto):
    if not isinstance(texto, str):
        return str(texto)
    texto = texto.strip().lower()
    texto = "".join(c for c in unicodedata.normalize('NFD', texto) if unicodedata.category(c) != 'Mn')
    return texto

# 1. SECCIÓN DE CARGA (Para la Secretaria)
st.sidebar.header("📁 Carga Diaria de Archivos")
archivo_subido = st.sidebar.file_uploader("Secretaria: Deposite el nuevo archivo Excel aquí", type=["xlsx"])

if archivo_subido is not None:
    try:
        exito = False
        for fila_cabecera in [0, 1, 2, 3]:
            # Forzamos a pandas a no parsear fechas automáticamente al leer para evitar pérdidas de formato mixtos
            df_prueba = pd.read_excel(archivo_subido, header=fila_cabecera)
            
            columnas_normalizadas = {normalizar_texto(col): col for col in df_prueba.columns}
            
            mapeo_encontrado = {}
            for col_req in COLUMNAS_REQUERIDAS:
                norm_req = normalizar_texto(col_req)
                if norm_req in columnas_normalizadas:
                    mapeo_encontrado[col_req] = columnas_normalizadas[norm_req]
            
            if "Expediente" in mapeo_encontrado:
                df_nuevo = df_prueba.rename(columns={v: k for k, v in mapeo_encontrado.items()})
                columnas_a_guardar = [c for c in COLUMNAS_REQUERIDAS if c in df_nuevo.columns]
                df_filtrado = df_nuevo[columnas_a_guardar].dropna(subset=["Expediente"])
                
                # Guardar en el Maestro
                df_filtrado.to_excel(ARCHIVOMAESTRO, index=False)
                st.sidebar.success("✅ ¡Archivo Maestro actualizado con éxito!")
                exito = True
                break
        
        if not exito:
            st.sidebar.error("❌ No se encontraron las columnas correctas.")
            df_error = pd.read_excel(archivo_subido)
            st.sidebar.write("Columnas detectadas en tu archivo actualmente:")
            st.sidebar.write(list(df_error.columns))

    except Exception as e:
        st.sidebar.error(f"Error al procesar el archivo: {e}")

# --- PROCESAMIENTO Y VISUALIZACIÓN DEL REPORTE ---
if os.path.exists(ARCHIVOMAESTRO):
    df_maestro = pd.read_excel(ARCHIVOMAESTRO)
    
    # CORRECCIÓN AVANZADA DE FECHAS MIXTAS (Elimina el "Ver Origen")
    if "Fecha Vencimiento" in df_maestro.columns:
        # Primero intentamos una conversión directa ignorando errores de texto
        df_maestro["Fecha Vencimiento"] = pd.to_datetime(df_maestro["Fecha Vencimiento"], errors='coerce')
        
        # Si la fecha viene como string con hora (ej: 2026-07-19 00:00:00), nos quedamos solo con la fecha limpia YYYY-MM-DD
        # Si la celda venía completamente vacía, la dejamos en blanco de manera ordenada sin textos raros
        df_maestro["Fecha Vencimiento"] = df_maestro["Fecha Vencimiento"].dt.strftime('%d/%m/%Y').fillna("-")

    # Asegurar que Dias Profesional sea numérico para poder ordenarlo correctamente
    if "Dias Profesional" in df_maestro.columns:
        df_maestro["Dias Profesional"] = pd.to_numeric(df_maestro["Dias Profesional"], errors='coerce').fillna(0)

    st.subheader("Filtrar Expedientes Críticos")
    col1, col2, col3 = st.columns([1, 1, 4])
    
    if "filtro" not in st.session_state:
        st.session_state.filtro = "Todos"
        
    with col1:
        if st.button("🚨 Vencidos", use_container_width=True):
            st.session_state.filtro = "Vencido"
    with col2:
        if st.button("📅 Por Vencer", use_container_width=True):
            st.session_state.filtro = "Por Vencer"
            
    # Lógica de filtrado y CORRECCIÓN DE "POR VENCER + HOY"
    if "Alerta Vencimiento" in df_maestro.columns:
        if st.session_state.filtro == "Vencido":
            df_mostrar = df_maestro[df_maestro["Alerta Vencimiento"].astype(str).str.contains("Vencido", case=False, na=False)]
            st.markdown("### 🔴 Mostrando: Expedientes Vencidos")
        elif st.session_state.filtro == "Por Vencer":
            # Filtra si contiene "Por Vencer" O contiene "Hoy"
            condicion = df_maestro["Alerta Vencimiento"].astype(str).str.contains("Por Vencer", case=False, na=False) | \
                        df_maestro["Alerta Vencimiento"].astype(str).str.contains("Hoy", case=False, na=False)
            df_mostrar = df_maestro[condicion]
            st.markdown("### 🟡 Mostrando: Expedientes Por Vencer")
        else:
            df_mostrar = df_maestro
            st.markdown("### 📋 Mostrando: Todos los expedientes")
    else:
        df_mostrar = df_maestro
        st.markdown("### 📋 Mostrando: Todos los expedientes")

    # CORRECCIÓN DE ORDEN DESCENDENTE POR "DIAS PROFESIONAL"
    if "Dias Profesional" in df_mostrar.columns:
        df_mostrar = df_mostrar.sort_values(by="Dias Profesional", ascending=False)

    # Mostrar tabla con las columnas requeridas ordenadas
    cols_visibles = [c for c in COLUMNAS_REQUERIDAS if c in df_mostrar.columns]
    
    if not df_mostrar.empty:
        # Renderizado optimizado para mostrar los registros preliminares con barra horizontal forzada
        st.dataframe(
            df_mostrar[cols_visibles], 
            use_container_width=True, 
            hide_index=True
        )
        st.metric(label="Total de Expedientes en Vista", value=len(df_mostrar))
    else:
        st.info("No se encontraron expedientes para este filtro.")
else:
    st.info("👋 Bienvenido. Esperando que la secretaria deposite el primer archivo.")