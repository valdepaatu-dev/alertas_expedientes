import os
import unicodedata
import pandas as pd
import streamlit as st

# Configuración de la página web
st.set_page_config(
    page_title="Alerta de Vencimiento de Expedientes", layout="wide"
)
st.title("📊 Alerta de Vencimiento de Expedientes")

ARCHIVOMAESTRO = "Consolidado_Maestro_Expedientes.xlsx"

# Nombres estándar que usaremos en la App (Se agrega 'Tipo de Procedimiento')
COLUMNAS_REQUERIDAS = [
    "Expediente",
    "Fecha Vencimiento",
    "Remitente",
    "Asunto",
    "Alerta Vencimiento",
    "Nombre Usuario Asignado",
    "Dias Profesional",
    "Tipo de Procedimiento",
]


# Función para normalizar texto
def normalizar_texto(texto):
    if not isinstance(texto, str):
        texto = str(texto) if pd.notna(texto) else ""
    texto = texto.strip().lower()
    return "".join(
        c
        for c in unicodedata.normalize("NFD", texto)
        if unicodedata.category(c) != "Mn"
    )


# Carga optimizada del archivo maestro utilizando caché
@st.cache_data
def cargar_datos_maestros(ruta_archivo):
    if os.path.exists(ruta_archivo):
        return pd.read_excel(ruta_archivo)
    return None


# --- CONTROL DE ACCESO (MODO ADMINISTRADOR) ---
st.sidebar.header("🔐 Control de Acceso")
modo_admin = st.sidebar.checkbox("Modo Administrador (Carga Diaria)")

es_admin = False
if modo_admin:
    contrasena = st.sidebar.text_input(
        "Contraseña de Administrador:", type="password"
    )
    if contrasena == "atu2026":
        st.sidebar.success("🔓 Modo Administrador Activo")
        es_admin = True
    elif contrasena != "":
        st.sidebar.error("❌ Contraseña incorrecta")

# --- 1. SECCIÓN DE CARGA (Visible SOLO para el Administrador) ---
if es_admin:
    st.sidebar.markdown("---")
    st.sidebar.subheader("📁 Carga Diaria de Archivos")
    archivo_subido = st.sidebar.file_uploader(
        "Deposite el nuevo archivo Excel aquí", type=["xlsx"]
    )

    if archivo_subido is not None:
        try:
            exito = False
            for fila_cabecera in [0, 1, 2, 3]:
                df_prueba = pd.read_excel(archivo_subido, header=fila_cabecera)
                columnas_normalizadas = {
                    normalizar_texto(col): col for col in df_prueba.columns
                }

                mapeo_encontrado = {}
                for col_req in COLUMNAS_REQUERIDAS:
                    norm_req = normalizar_texto(col_req)
                    if norm_req in columnas_normalizadas:
                        mapeo_encontrado[col_req] = columnas_normalizadas[norm_req]

                if "Expediente" in mapeo_encontrado:
                    df_nuevo = df_prueba.rename(
                        columns={v: k for k, v in mapeo_encontrado.items()}
                    )
                    columnas_a_guardar = [
                        c for c in COLUMNAS_REQUERIDAS if c in df_nuevo.columns
                    ]
                    df_filtrado = df_nuevo[columnas_a_guardar].dropna(
                        subset=["Expediente"]
                    )

                    # Guardar archivo localmente y limpiar caché de Streamlit
                    df_filtrado.to_excel(ARCHIVOMAESTRO, index=False)
                    st.cache_data.clear()
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
else:
    st.sidebar.info(
        "💡 Los usuarios generales solo pueden visualizar los reportes."
    )

# --- PROCESAMIENTO Y VISUALIZACIÓN DEL REPORTE ---
df_maestro_raw = cargar_datos_maestros(ARCHIVOMAESTRO)

if df_maestro_raw is not None:
    df_maestro = df_maestro_raw.copy()

    # Formateo de fechas
    if "Fecha Vencimiento" in df_maestro.columns:
        df_maestro["Fecha Vencimiento"] = pd.to_datetime(
            df_maestro["Fecha Vencimiento"], errors="coerce"
        ).dt.strftime("%d/%m/%Y").fillna("-")

    # Asegurar que Dias Profesional sea numérico
    if "Dias Profesional" in df_maestro.columns:
        df_maestro["Dias Profesional"] = pd.to_numeric(
            df_maestro["Dias Profesional"], errors="coerce"
        ).fillna(0)

    # Regla Congreso -> Etiqueta "Urgente"
    if "Tipo de Procedimiento" in df_maestro.columns and "Alerta Vencimiento" in df_maestro.columns:
        asunto_norm = df_maestro["Tipo de Procedimiento"].astype(str).apply(normalizar_texto)
        es_congreso = asunto_norm.str.contains("congreso", na=False)
        df_maestro.loc[es_congreso, "Alerta Vencimiento"] = "Urgente"

    # --- FILTROS DE INTERFAZ ---
    st.subheader("Filtrar Expedientes Críticos")
    col1, col2, col3, _ = st.columns([1, 1, 1, 3])

    if "filtro" not in st.session_state:
        st.session_state.filtro = "Todos"

    with col1:
        if st.button("🚨 Vencidos", use_container_width=True):
            st.session_state.filtro = "Vencido"
    with col2:
        if st.button("📅 Por Vencer", use_container_width=True):
            st.session_state.filtro = "Por Vencer"
    with col3:
        if st.button("📋 Ver Todos", use_container_width=True):
            st.session_state.filtro = "Todos"

    # Lógica de filtrado
    if "Alerta Vencimiento" in df_maestro.columns:
        alerta_str = df_maestro["Alerta Vencimiento"].astype(str)

        if st.session_state.filtro == "Vencido":
            df_mostrar = df_maestro[alerta_str.str.contains("Vencido", case=False, na=False)]
            st.markdown("### 🔴 Expedientes Vencidos")

        elif st.session_state.filtro == "Por Vencer":
            condicion = (
                alerta_str.str.contains("Por Vencer", case=False, na=False)
                | alerta_str.str.contains("Hoy", case=False, na=False)
                | (alerta_str == "Urgente")
            )
            df_mostrar = df_maestro[condicion]
            st.markdown("### 🟡 Expedientes Por Vencer y Urgentes")

        else:
            df_mostrar = df_maestro.copy()
            st.markdown("### 📋 Mostrando: Todos los expedientes")
    else:
        df_mostrar = df_maestro.copy()
        st.markdown("### 📋 Mostrando: Todos los expedientes")

    # Ordenamiento prioritario ("Urgente" primero, luego "Dias Profesional")
    if "Alerta Vencimiento" in df_mostrar.columns:
        df_mostrar["Es_Urgente"] = (
            df_mostrar["Alerta Vencimiento"].astype(str) == "Urgente"
        )
        criterios_orden = ["Es_Urgente"]
        ascendente = [False]

        if "Dias Profesional" in df_mostrar.columns:
            criterios_orden.append("Dias Profesional")
            ascendente.append(False)

        df_mostrar = df_mostrar.sort_values(by=criterios_orden, ascending=ascendente).drop(
            columns=["Es_Urgente"]
        )

    # Filtrar solo columnas requeridas para la vista limpia de usuario
    cols_visibles = [c for c in COLUMNAS_REQUERIDAS if c in df_mostrar.columns and c != "Tipo de Procedimiento"]

    if not df_mostrar.empty:
        st.dataframe(
            df_mostrar[cols_visibles], use_container_width=True, hide_index=True
        )
        st.metric(label="Total de Expedientes en Vista", value=len(df_mostrar))
    else:
        st.info("No se encontraron expedientes para este filtro.")
else:
    st.info(
        "👋 Bienvenido. Esperando que el administrador realice la primera carga del archivo base."
    )
