import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows

# ================= RUTAS =================
PLANTILLA_PATH = "C:/Users/denni/OneDrive/Escritorio/consolidador_tool/Inventarios/PLANTILLAS/plantilla_base.xlsm"
# Ruta opcional para ejemplo
EJEMPLO_PATH = "C:/Users/denni/OneDrive/Escritorio/consolidador_tool/Inventarios/PLANTILLAS/ejemplo.xlsm"

# ================= CONFIGURACIÓN =================
st.set_page_config(page_title="Consolidador de Inventarios", page_icon="🧾", layout="centered")
st.title("Consolidador de Inventarios por Marca")

# ================= INSTRUCCIONES Y BOTONES =================
col1, col2 = st.columns([3, 1])

with col1:
    st.markdown("### 🧠 ¿Cómo usar esta herramienta?")
    st.markdown("""
    1. **Usa la plantilla oficial** para recolectar los datos.
    2. Asegúrate de que todos los archivos subidos contengan la columna **Marca** (pueden ser distintas).
    3. Sube los archivos y la herramienta detectará todas las marcas.
    4. Selecciona las marcas que quieres consolidar.
    5. Para cada marca seleccionada, ingresa los datos fijos del cliente y selecciona mes/año.
    6. Se mostrará siempre 1 registro por fila, aunque algunas Cajas sean 0 o estén vacías.
    7. Las fechas se guardan en el consolidado en formato `YYYY-MM-DD` (sin hora).
    """)

with col2:
    # Botón plantilla oficial
    with open(PLANTILLA_PATH, "rb") as f:
        plantilla_bytes = f.read()
    st.download_button(
        label="📥 Plantilla oficial",
        data=plantilla_bytes,
        file_name="plantilla_base_formato_cliente.xlsm",
        mime="application/vnd.ms-excel.sheet.macroEnabled.12"
    )

    # Botón ejemplo (opcional)
    try:
        with open(EJEMPLO_PATH, "rb") as f:
            ejemplo_bytes = f.read()
        st.download_button(
            label="📄 Ejemplo lleno",
            data=ejemplo_bytes,
            file_name="plantilla_ejemplo.xlsm",
            mime="application/vnd.ms-excel.sheet.macroEnabled.12"
        )
    except FileNotFoundError:
        st.write("")

# ================= REINICIO =================
if "file_uploader_key" not in st.session_state:
    st.session_state["file_uploader_key"] = "initial"

if st.button("🔁 Limpiar todo y comenzar de nuevo"):
    st.session_state.clear()
    st.session_state["file_uploader_key"] = str(datetime.now())
    st.rerun()

# ================= CARGA DE ARCHIVOS =================
uploaded_files = st.file_uploader(
    "📂 Sube uno o más archivos Excel con la plantilla oficial",
    type=["xlsm", "xlsx"],
    accept_multiple_files=True,
    key=st.session_state["file_uploader_key"]
)

registros_finales = []
marcas_detectadas = set()

if uploaded_files:
    st.success(f"Has subido {len(uploaded_files)} archivo(s).")

    for file in uploaded_files:
        try:
            df = pd.read_excel(file)
            df.columns = df.columns.str.strip()

            columnas_base = ["Nombre Comercial", "Tipo de cliente", "Marca", "Codigo de producto", "Descripción"]
            if not all(col in df.columns for col in columnas_base):
                st.error(f"❌ El archivo {file.name} no tiene las columnas necesarias.")
                continue

            marcas_detectadas.update(df["Marca"].dropna().unique())

            cajas_cols = [col for col in df.columns if 'Cajas' in col]
            fechas_cols = [col for col in df.columns if 'Fecha' in col]

            for _, row in df.iterrows():
                base_data = {col: row[col] for col in columnas_base}
                for i in range(min(len(cajas_cols), len(fechas_cols))):
                    base_data[f"Cajas_{i+1}"] = row[cajas_cols[i]]
                    base_data[f"Fecha_{i+1}"] = row[fechas_cols[i]]
                registros_finales.append(base_data)

        except Exception as e:
            st.error(f"❌ Error al procesar {file.name}: {e}")

# ================= SELECCIÓN DE MARCAS =================
marcas_seleccionadas = []
if marcas_detectadas:
    st.markdown("### ✅ Selecciona las marcas que quieres consolidar (Todas están marcadas por defecto)")
    for marca in sorted(marcas_detectadas):
        if st.checkbox(marca, value=True, key=f"chk_{marca}"):
            marcas_seleccionadas.append(marca)

# ================= FORMULARIOS POR MARCA =================
datos_por_marca = {}

if marcas_seleccionadas:
    st.markdown("### 📝 Datos por marca seleccionada")
    for marca in marcas_seleccionadas:
        with st.expander(f"📌 Datos para {marca}", expanded=True):
            col1, col2 = st.columns(2)
            with col1:
                mes = st.selectbox(f"📆 Mes ({marca})", [
                    "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
                    "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"
                ], key=f"mes_{marca}")
                ruta = st.text_input(f"🛣️ Ruta ({marca})", key=f"ruta_{marca}")
                codigo_cliente = st.text_input(f"🔢 Código de cliente ({marca})", key=f"codigo_cliente_{marca}")
            with col2:
                año = st.number_input(f"📅 Año ({marca})", min_value=2023, max_value=2100,
                                       value=datetime.now().year, key=f"año_{marca}")
                zona = st.text_input(f"📍 Zona ({marca})", key=f"zona_{marca}")
                tienda = st.text_input(f"🏪 Nombre de la tienda ({marca})", key=f"tienda_{marca}")

            datos_por_marca[marca] = {
                "Mes": mes,
                "Ruta": ruta,
                "Codigo de cliente": codigo_cliente,
                "Año": año,
                "Zona": zona,
                "Nombre de la tienda": tienda
            }

# ================= GENERACIÓN DE ARCHIVOS =================
if registros_finales and datos_por_marca and st.button("📦 Generar archivo consolidado por marca"):
    with st.spinner("⏳ Generando los archivos consolidados... por favor espera..."):
        df_consolidado = pd.DataFrame(registros_finales)

        st.session_state.downloads_por_marca = []

        for marca, datos in datos_por_marca.items():
            df_marca = df_consolidado[df_consolidado["Marca"] == marca].copy()

            if df_marca.empty:
                continue

            for col in df_marca.columns:
                if col.startswith("Fecha_"):
                    df_marca[col] = pd.to_datetime(df_marca[col], errors="coerce").dt.strftime("%Y-%m-%d")

            for col, valor in datos.items():
                df_marca[col] = valor

            orden_columnas = [
                "Mes", "Ruta", "Zona", "Codigo de cliente", "Nombre de la tienda",
                "Nombre Comercial", "Tipo de cliente", "Marca", "Codigo de producto", "Descripción",
                "Cajas_1", "Fecha_1", "Cajas_2", "Fecha_2", "Cajas_3", "Fecha_3", "Cajas_4", "Fecha_4"
            ]
            df_marca = df_marca[[col for col in orden_columnas if col in df_marca.columns]]

            wb = load_workbook(PLANTILLA_PATH, keep_vba=True)
            ws = wb.active
            if ws.max_row > 1:
                ws.delete_rows(2, ws.max_row)

            for row in dataframe_to_rows(df_marca, index=False, header=False):
                ws.append(row)

            output = BytesIO()
            wb.save(output)
            output.seek(0)

            file_name = f"{datos['Ruta']}_Inventario_{marca}_{datos['Mes']}_{datos['Año']}.xlsm"
            st.session_state.downloads_por_marca.append((file_name, output))

    st.success("✅ Archivos por marca generados correctamente.")

# ================= DESCARGAS =================
if "downloads_por_marca" in st.session_state and st.session_state.downloads_por_marca:
    st.markdown("### 📦 Archivos disponibles:")
    for file_name, file_data in st.session_state.downloads_por_marca:
        st.download_button(
            label=f"📥 Descargar {file_name}",
            data=file_data,
            file_name=file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=file_name
        )
