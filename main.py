import datetime
import os
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Control de Procesos - Líneas",
    page_icon="📋",
    layout="wide",
)

# Rutas de archivos de datos
FILE_PATH = "procesos.xlsx"
DB_FILE = "registros_control_procesos.csv"

# Valores por defecto extraídos del archivo Excel
PRODUCTOS_DEFAULT = [
    "MUFFIN DE MANZANA - FRESCO - (UND)",
    "MUFFIN DE NARANJA & CHOCOCHIPS - FRESCO - (UND)",
    "MUFFIN DE BERRIES - FRESCO - (UND)",
    "MUFFIN DE CHOCOLATE - FRESCO - (UND)",
    "MUFFIN DE RED VELVET (SIN DECORAR)",
    "GALLETA CHOCOCHIP CON AVELLANA STBX",
    "CINNAMON ROLL - STB",
    "BASE DE TARTALETA CIRCULAR",
    "CROUMBLE BERRIES HORNEADO",
    "BASE CROUMBLE DE BERRIES STBX",
    "CROUMBLE BERRY KG STBX",
    "ROLL CINNAMON STARBUCKS",
    "TORTA DE CHOCOLATE CON FUDGE Y MANJAR",
    "PYE DE LIMON - FRESCO - (UND) (STBX)",
    "KEKE RECTANGULAR DE LIMON STBX",
    "KEKE RECTAGULAR DE ZANAHORIA STBX",
    "KEKE RECTANGULAR GINGER DECORADO - B2B",
    "MUFFIN DE RED VELVET - UND",
    "KEKE RECTANGULAR DE LIMON STB",
    "KEKE DE ZANAHORIA INTEGRAL Y PANELA STB",
]

LINEAS_DEFAULT = ["MUFFINS", "AMASADO", "DECORADO", "BATIDOS", "OTRO"]

EQUIPOS_DEFAULT = [
    "BATIDORA",
    "ROBOTCOUPE",
    "AMASADORA",
    "GALLETERA",
    "DIVISORA",
    "DOSIFICADORA",
    "LAMINADORA",
    "NO APLICA",
    "OTRO",
]


@st.cache_data(ttl=60)
def cargar_catalogos():
    """Carga los catálogos de forma eficiente desde el Excel o usa valores por defecto."""
    productos = PRODUCTOS_DEFAULT
    lineas = LINEAS_DEFAULT
    equipos = EQUIPOS_DEFAULT

    if os.path.exists(FILE_PATH):
        try:
            xls = pd.ExcelFile(FILE_PATH)
            for sheet in xls.sheet_names:
                df = pd.read_excel(FILE_PATH, sheet_name=sheet)
                for idx, row in df.iterrows():
                    row_vals = [str(v).strip() for v in row.values if pd.notna(v)]
                    if "PRODUCTO" in row_vals:
                        header_idx = idx
                        df_data = df.iloc[header_idx + 1 :].copy()
                        df_data.columns = [
                            str(c).strip() for c in df.iloc[header_idx].values
                        ]

                        if "PRODUCTO" in df_data.columns:
                            prods = (
                                df_data["PRODUCTO"]
                                .dropna()
                                .astype(str)
                                .str.strip()
                                .unique()
                                .tolist()
                            )
                            if prods:
                                productos = prods

                        if "LÍNEA DE PROCESO" in df_data.columns:
                            lins = (
                                df_data["LÍNEA DE PROCESO"]
                                .dropna()
                                .astype(str)
                                .str.strip()
                                .unique()
                                .tolist()
                            )
                            if lins:
                                lineas = list(dict.fromkeys(lins + ["OTRO"]))

                        if "EQUIPO UTILIZADO" in df_data.columns:
                            eqs = (
                                df_data["EQUIPO UTILIZADO"]
                                .dropna()
                                .astype(str)
                                .str.strip()
                                .unique()
                                .tolist()
                            )
                            if eqs:
                                equipos = list(dict.fromkeys(eqs + ["OTRO"]))
                        break
        except Exception:
            pass

    return productos, lineas, equipos


lista_productos, lista_lineas, lista_equipos = cargar_catalogos()

# --- TÍTULO DE LA APLICACIÓN ---
st.title("📋 FORMATO CONTROL DE PROCESOS LÍNEAS")
st.markdown("---")

# --- FORMULARIO COMPLETO DE REGISTRO (clear_on_submit=True limpia las casillas al guardar) ---
with st.form("form_control_proceso", clear_on_submit=True):
    st.subheader("1. Información General del Proceso")
    col1, col2, col3, col4 = st.columns(4)

    with col1:
        fecha_p = st.date_input(
            "F.P (Fecha de Producción)", value=datetime.date.today()
        )
    with col2:
        lote = st.text_input(
            "LOTE", placeholder="Ej. L-20260901", key="input_lote"
        )
    with col3:
        batch = st.text_input("BATCH", placeholder="Ej. B-01", key="input_batch")
    with col4:
        # Campo Abierto (Texto libre para llenar manualmente)
        responsable = st.text_input(
            "RESPONSABLE",
            placeholder="Ingrese el nombre del responsable",
            key="input_responsable",
        )

    st.subheader("2. Selección de Producto, Línea y Equipo")
    col5, col6, col7 = st.columns(3)

    with col5:
        producto = st.selectbox("PRODUCTO", options=lista_productos)

    with col6:
        linea_sel = st.selectbox("LÍNEA DE PROCESO", options=lista_lineas)
        if linea_sel == "OTRO":
            linea_text = st.text_input(
                "Especifique Línea de Proceso", placeholder="Nombre de línea"
            )
            linea_final = linea_text
        else:
            linea_final = linea_sel

    with col7:
        equipo_sel = st.selectbox("EQUIPO UTILIZADO", options=lista_equipos)
        if equipo_sel == "OTRO":
            equipo_text = st.text_input(
                "Especifique Equipo", placeholder="Nombre del equipo"
            )
            equipo_final = equipo_text
        else:
            equipo_final = equipo_sel

    st.subheader("3. Condiciones del Entorno e Insumos")
    col8, col9, col10, col11 = st.columns(4)

    with col8:
        cond_area = st.radio(
            "CONDICIONES DEL AREA DE TRABAJO",
            options=["CONFORME", "NO CONFORME"],
            horizontal=True,
        )
    with col9:
        cond_equipo = st.radio(
            "CONDICIONES DEL EQUIPO",
            options=["CONFORME", "NO CONFORME"],
            horizontal=True,
        )
    with col10:
        cond_insumos = st.radio(
            "CONDICIONES DE LOS INSUMOS",
            options=["CONFORME", "NO CONFORME"],
            horizontal=True,
        )
    with col11:
        caract_producto = st.radio(
            "CARACTERISTICAS DEL PRODUCTO",
            options=["CONFORME", "NO CONFORME"],
            horizontal=True,
        )

    st.subheader("4. Horarios y Tiempo de Proceso")
    col12, col13, col14 = st.columns(3)

    with col12:
        hora_inicio = st.time_input("HORA INICIO", value=datetime.time(8, 0))
    with col13:
        hora_termino = st.time_input("HORA TÉRMINO", value=datetime.time(8, 15))

    # Cálculo automático de tiempo transcurrido
    dt_inicio = datetime.datetime.combine(datetime.date.today(), hora_inicio)
    dt_termino = datetime.datetime.combine(datetime.date.today(), hora_termino)
    if dt_termino < dt_inicio:
        dt_termino += datetime.timedelta(days=1)

    minutos_totales = int((dt_termino - dt_inicio).total_seconds() / 60)
    tiempo_calculado = f"{minutos_totales} min"

    with col14:
        st.text_input(
            "TIEMPO CALCULADO", value=tiempo_calculado, disabled=True
        )

    st.subheader("5. Observaciones Finales")
    col15, col16 = st.columns([1, 2])

    with col15:
        estado_obs = st.radio(
            "OBSERVACIÓN",
            options=["CONFORME", "NO CONFORME", "OTRO (Texto Libre)"],
        )

    with col16:
        if estado_obs == "OTRO (Texto Libre)":
            obs_detalle = st.text_area(
                "Detalle de Observación",
                placeholder="Escriba aquí la observación personalizada...",
                key="input_obs_detalle",
            )
            observacion_final = obs_detalle
        else:
            obs_adicional = st.text_input(
                "Comentario Adicional (Opcional)",
                placeholder="Escriba detalles si aplica...",
                key="input_obs_adic",
            )
            observacion_final = (
                f"{estado_obs} - {obs_adicional}".strip(" -")
                if obs_adicional
                else estado_obs
            )

    st.markdown("---")
    btn_guardar = st.form_submit_button(
        "💾 Guardar Registro de Control", use_container_width=True
    )

# --- PROCESAMIENTO Y GUARDADO ---
if btn_guardar:
    nuevo_registro = {
        "PRODUCTO": producto,
        "LÍNEA DE PROCESO": linea_final,
        "CONDICIONES DEL AREA DE TRABAJO": cond_area,
        "F.P": fecha_p.strftime("%Y-%m-%d"),
        "LOTE": lote,
        "BATCH": batch,
        "EQUIPO UTILIZADO": equipo_final,
        "HORA INICIO": hora_inicio.strftime("%H:%M"),
        "CONDICIONES DEL EQUIPO": cond_equipo,
        "CONDICIONES DE LOS INSUMOS": cond_insumos,
        "CARACTERISTICAS DEL PRODUCTO": caract_producto,
        "HORA TÉRMINO": hora_termino.strftime("%H:%M"),
        "TIEMPO": tiempo_calculado,
        "RESPONSABLE": responsable,
        "OBSERVACIÓN": observacion_final,
    }

    df_nuevo = pd.DataFrame([nuevo_registro])

    if os.path.exists(DB_FILE):
        df_hist = pd.read_csv(DB_FILE)
        df_hist = pd.concat([df_hist, df_nuevo], ignore_index=True)
    else:
        df_hist = df_nuevo

    df_hist.to_csv(DB_FILE, index=False)
    st.success("✅ Registro guardado exitosamente.")

# --- HISTORIAL Y EXPORTACIÓN ---
st.markdown("---")
st.subheader("📊 Historial de Registros Guardados")

if os.path.exists(DB_FILE):
    df_registros = pd.read_csv(DB_FILE)
    st.dataframe(df_registros, use_container_width=True)

    csv_data = df_registros.to_csv(index=False).encode("utf-8")
    st.download_button(
        label="📥 Descargar Historial Completo (CSV)",
        data=csv_data,
        file_name="FORMATO_CONTROL_PROCESOS_LINEAS.csv",
        mime="text/csv",
    )
else:
    st.info("Aún no hay registros guardados en la base de datos.")