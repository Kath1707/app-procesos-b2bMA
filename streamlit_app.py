"""
App de Registro de Parámetros de Calidad - Producto Intermedio (PI) B2B Starbucks
====================================================================================
Streamlit. Por ahora SIN conexión a Google Sheets (solo exportación local CSV/Excel).

Estructura esperada del repo (el Excel de especificaciones va al MISMO NIVEL que
este archivo y que requirements.txt, no dentro de una carpeta "data/"):

  streamlit_app_pi.py
  requirements.txt
  MA-PL-019_PLAN_CALIDAD_DE_PRODUCTOS.xlsx
"""

import io
import re
from datetime import date

import pandas as pd
import streamlit as st

# ----------------------------------------------------------------------------
# CONFIGURACIÓN GENERAL
# ----------------------------------------------------------------------------
st.set_page_config(
    page_title="Registro de Calidad - Producto Intermedio B2B Starbucks",
    page_icon="🧪",
    layout="wide",
)

EXCEL_PATH = "MA-PL-019_PLAN_CALIDAD_DE_PRODUCTOS.xlsx"  # mismo nivel que este .py
SHEET_NAME = "PROCESO-B2B-STB"
CLIENTE_FIJO = "STARBUCKS"
AREA_FIJA = "EMPAQUE"
MAX_MUESTRAS_COLUMNAS = 8  # sub-columnas fijas preparadas por parámetro en el exportable

EQUIPO_CALIDAD = [
    "Verónica Iriarte",
    "Cristina Merino",
    "Lisseth Aspíllaga",
    "Sandra Chavez",
    "Alejandro Herrera",
    "Katherin Hidalgo",
]

TURNOS = ["Día", "Tarde", "Madrugada"]

# (clave, etiqueta) en el orden fijo que tendrán las columnas del exportable
ALL_PARAM_DEFS = [
    ("peso", "Peso (g)"),
    ("diametro", "Diámetro (cm)"),
    ("altura", "Altura/Espesor (cm)"),
    ("temperatura", "Temperatura (°C)"),
    ("tamizado", "Tamizado"),
    ("tiempo_mezcla", "Tiempo Mezcla (min)"),
    ("decorado", "Decorado"),
    ("brix", "Brix"),
    ("organolepticas", "Características Organolépticas"),
    ("estado_material", "Estado del Material/Equipo/Herramienta"),
    ("observacion", "Observación/Corrección"),
]


def iniciales(nombre_completo: str) -> str:
    partes = nombre_completo.strip().split()
    if len(partes) < 2:
        return nombre_completo[:2].upper()
    return (partes[0][0] + partes[-1][0]).upper()


# ----------------------------------------------------------------------------
# TABLA MIL-STD-105E - Nivel de Inspección Especial S-2, Inspección Rigurosa
# (Tightened, Tabla II-B), AQL 4.0% — misma tabla que usamos en la app de PT
# ----------------------------------------------------------------------------
SAMPLING_TABLE_S2 = [
    (2, 8, "A", 2, 0, 1),
    (9, 15, "A", 2, 0, 1),
    (16, 25, "B", 3, 0, 1),
    (26, 50, "B", 3, 0, 1),
    (51, 90, "B", 3, 0, 1),
    (91, 150, "C", 5, 0, 1),
    (151, 280, "C", 5, 0, 1),
    (281, 500, "C", 5, 0, 1),
    (501, 1200, "D", 8, 1, 2),
    (1201, 3200, "D", 8, 1, 2),
    (3201, 10000, "D", 8, 1, 2),
    (10001, 35000, "E", 13, 1, 2),
    (35001, 150000, "E", 13, 1, 2),
    (150001, 10_000_000, "E", 13, 1, 2),
]


def get_sample_size(lot_size: int):
    for low, high, letter, n, ac, re_ in SAMPLING_TABLE_S2:
        if low <= lot_size <= high:
            return letter, n, ac, re_
    if lot_size < 2:
        return "A", 2, 0, 1
    return "E", 13, 1, 2


# ----------------------------------------------------------------------------
# CARGA Y LIMPIEZA DE DATOS DEL EXCEL
# ----------------------------------------------------------------------------
@st.cache_data(show_spinner=False)
def load_specs(excel_bytes: bytes) -> pd.DataFrame:
    df_raw = pd.read_excel(
        io.BytesIO(excel_bytes),
        sheet_name=SHEET_NAME,
        header=None,
        skiprows=7,
    )

    cols = [
        "producto", "linea_produccion", "linea_haccp", "componente", "actividad",
        "tipo", "peso", "diametro", "altura", "temperatura", "tamizado",
        "tiempo_mezcla", "decorado", "brix", "organolepticas", "estado_material",
        "responsable", "observacion",
    ]
    df_raw = df_raw.iloc[:, : len(cols)]
    df_raw.columns = cols

    # Solo nos quedamos con filas que tengan actividad (evita filas totalmente vacías)
    df_raw = df_raw[df_raw["actividad"].notna()].copy()

    def clean_txt(x):
        if pd.isna(x):
            return x
        return re.sub(r"\s+", " ", str(x)).strip()

    text_cols = [
        "producto", "linea_produccion", "linea_haccp", "componente", "actividad",
        "peso", "diametro", "altura", "temperatura", "tamizado", "tiempo_mezcla",
        "decorado", "brix", "organolepticas", "estado_material", "observacion",
    ]
    for c in text_cols:
        df_raw[c] = df_raw[c].apply(clean_txt)

    # Forward-fill: producto/línea de producción/línea HACCP solo aparecen en la
    # primera fila de cada grupo, el resto queda en blanco en el Excel original.
    for c in ["producto", "linea_produccion", "linea_haccp"]:
        df_raw[c] = df_raw[c].ffill()

    df_raw["producto"] = df_raw["producto"].str.upper()
    df_raw["linea_haccp"] = df_raw["linea_haccp"].str.upper()

    df_raw = df_raw[df_raw["tipo"].astype(str).str.upper().str.strip() == "PI"]

    return df_raw.reset_index(drop=True)


def campo_aplica(valor) -> bool:
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return False
    v = str(valor).strip()
    return v not in ("", "-", "—", "nan", "None")


def get_excel_bytes():
    try:
        with open(EXCEL_PATH, "rb") as f:
            return f.read()
    except FileNotFoundError:
        st.warning(
            f"No encontré el archivo `{EXCEL_PATH}` en la raíz del repositorio."
        )
        up = st.file_uploader("Sube el Excel de especificaciones (.xlsx)", type=["xlsx"])
        if up is not None:
            return up.read()
        st.stop()


# ----------------------------------------------------------------------------
# ESTADO / WIZARD
# ----------------------------------------------------------------------------
if "step" not in st.session_state:
    st.session_state.step = 1


def go_next():
    st.session_state.step += 1


def go_back():
    st.session_state.step -= 1


excel_bytes = get_excel_bytes()
specs_df = load_specs(excel_bytes)

# ============================================================================
# PASO 1 - PORTADA
# ============================================================================
if st.session_state.step == 1:
    st.title("🧪 Registro de Parámetros de Calidad")
    st.subheader("Producto Intermedio (PI) - Línea de Producción B2B Starbucks")
    st.markdown(
        """
        Esta aplicación permite al **equipo de calidad** registrar la inspección
        de producto **intermedio** (etapas del proceso, no producto terminado)
        de la línea B2B Starbucks, siguiendo el plan de muestreo
        **MIL-STD-105E (Nivel Especial S-2, Inspección Rigurosa, AQL 4.0%)**
        cuando aplica número de batch.
        """
    )
    st.button("Comenzar registro ➜", on_click=go_next, type="primary")

# ============================================================================
# PASO 2 - EQUIPO DE CALIDAD, TURNO Y FECHA
# ============================================================================
elif st.session_state.step == 2:
    st.header("1️⃣ Datos del registro")

    responsable = st.selectbox(
        "Equipo de calidad (responsable del registro)",
        EQUIPO_CALIDAD,
        index=EQUIPO_CALIDAD.index(st.session_state.get("responsable", EQUIPO_CALIDAD[0]))
        if st.session_state.get("responsable") in EQUIPO_CALIDAD else 0,
    )

    turno = st.selectbox(
        "Turno",
        TURNOS,
        index=TURNOS.index(st.session_state.get("turno", TURNOS[0]))
        if st.session_state.get("turno") in TURNOS else 0,
    )

    fecha_produccion = st.date_input(
        "Fecha de producción (= fecha de registro)",
        value=st.session_state.get("fecha_produccion", None),
        format="DD/MM/YYYY",
    )
    if fecha_produccion is None:
        st.caption("⚠️ Selecciona la fecha para continuar (no se precarga sola).")

    st.info(f"**Cliente:** {CLIENTE_FIJO}  |  **Área:** {AREA_FIJA}")

    col1, col2 = st.columns(2)
    with col1:
        st.button("⬅ Atrás", on_click=go_back)
    with col2:
        if st.button("Siguiente ➜", type="primary", disabled=fecha_produccion is None):
            st.session_state.responsable = responsable
            st.session_state.turno = turno
            st.session_state.fecha_produccion = fecha_produccion
            go_next()
            st.rerun()

# ============================================================================
# PASO 3 - LÍNEA HACCP Y PRODUCTO
# ============================================================================
elif st.session_state.step == 3:
    st.header("2️⃣ Línea HACCP y producto")

    lineas = sorted(specs_df["linea_haccp"].dropna().unique().tolist())
    linea_sel = st.selectbox("Línea de producción HACCP", lineas)

    productos_linea = sorted(
        specs_df.loc[specs_df["linea_haccp"] == linea_sel, "producto"].unique().tolist()
    )
    producto_sel = st.selectbox("Producto", productos_linea)

    col1, col2 = st.columns(2)
    with col1:
        st.button("⬅ Atrás", on_click=go_back)
    with col2:
        if st.button("Siguiente ➜", type="primary"):
            st.session_state.linea_haccp = linea_sel
            st.session_state.producto = producto_sel
            go_next()
            st.rerun()

# ============================================================================
# PASO 4 - ACTIVIDAD (filtrada por producto, distinguiendo repetidas)
# ============================================================================
elif st.session_state.step == 4:
    st.header("3️⃣ Actividad realizada")

    filas_prod = specs_df[specs_df["producto"] == st.session_state.producto].reset_index(drop=True)

    opciones = []
    for idx, fila in filas_prod.iterrows():
        etiqueta = f"{fila['actividad']} — {fila['componente']}"
        opciones.append((idx, etiqueta))

    idx_sel = st.selectbox(
        "Selecciona la actividad realizada",
        options=[idx for idx, _ in opciones],
        format_func=lambda i: dict(opciones)[i],
    )
    fila_actividad = filas_prod.loc[idx_sel]

    with st.expander("📋 Ficha de referencia de esta actividad", expanded=True):
        for clave, label in ALL_PARAM_DEFS:
            valor = fila_actividad[clave]
            if campo_aplica(valor):
                st.markdown(f"**{label}:** {valor}")

    col1, col2 = st.columns(2)
    with col1:
        st.button("⬅ Atrás", on_click=go_back)
    with col2:
        if st.button("Siguiente ➜", type="primary"):
            st.session_state.fila_actividad = fila_actividad.to_dict()
            go_next()
            st.rerun()

# ============================================================================
# PASO 5 - ¿APLICA NÚMERO DE BATCH? Y MUESTREO
# ============================================================================
elif st.session_state.step == 5:
    st.header("4️⃣ Batch y muestreo")

    aplica_batch = st.radio(
        "¿Aplica registrar un número de batch para esta actividad?",
        ["Sí", "No"],
        horizontal=True,
        help="Elige 'No' para procesos generales que aún no se han separado en "
             "batches (ej. un batido que luego se dosifica). En ese caso se evalúa 1 sola muestra.",
    )

    if aplica_batch == "Sí":
        batch_size = st.number_input("¿Cuántas unidades tiene el batch?", min_value=1, step=1, value=300)
        letra, n_muestras, ac, re_ = get_sample_size(int(batch_size))
        st.success(
            f"Para un batch de **{int(batch_size)}** unidades: letra código **{letra}** → "
            f"**n = {n_muestras}** muestras. Criterio: Aceptar con **{ac}** o menos no conformes, "
            f"Rechazar con **{re_}** o más."
        )
    else:
        batch_size = None
        letra, n_muestras, ac, re_ = None, 1, None, None
        st.info("No aplica número de batch — se evaluará como **1 sola muestra**.")

    col1, col2 = st.columns(2)
    with col1:
        st.button("⬅ Atrás", on_click=go_back)
    with col2:
        if st.button("Siguiente ➜", type="primary"):
            st.session_state.aplica_batch = aplica_batch
            st.session_state.batch_size = batch_size
            st.session_state.n_muestras = n_muestras
            st.session_state.letra_codigo = letra
            st.session_state.ac = ac
            st.session_state.re_ = re_
            go_next()
            st.rerun()

# ============================================================================
# PASO 6 - REGISTRO DE PARÁMETROS POR MUESTRA
# ============================================================================
elif st.session_state.step == 6:
    st.header("5️⃣ Registro de parámetros por muestra")

    fila = st.session_state.fila_actividad
    n = st.session_state.n_muestras

    st.markdown(f"**Producto:** {st.session_state.producto} &nbsp;|&nbsp; "
                f"**Actividad:** {fila['actividad']} — {fila['componente']} &nbsp;|&nbsp; "
                f"**N° de muestras a evaluar:** {n}")

    parametros = []
    for clave, label in ALL_PARAM_DEFS:
        valor = fila[clave]
        if not campo_aplica(valor):
            continue  # se omite: tenía "-" o estaba vacío
        if clave == "organolepticas":
            pregunta = "¿Las características organolépticas cumplen con lo especificado?"
        elif clave == "estado_material":
            pregunta = f"¿El estado del material/equipo/herramienta cumple con lo especificado? ({valor})"
        elif clave == "observacion":
            pregunta = f"¿Se cumple con la observación/corrección indicada? ({valor})"
        else:
            pregunta = f"¿{label} cumple con lo especificado? ({valor})"
        parametros.append((clave, label, pregunta))

    if not parametros:
        st.warning("Esta actividad no tiene parámetros de control aplicables.")
    else:
        if "organolepticas" in [p[0] for p in parametros]:
            with st.expander("Ver detalle completo de características organolépticas"):
                st.text(fila["organolepticas"])

        st.caption("Completa cada muestra tocando 'Conforme' o 'No conforme'. Si algún parámetro "
                   "tiene 1 o más 'No conforme', se habilitará un cuadro de comentario al final de esa pestaña.")

        if "respuestas" not in st.session_state:
            st.session_state.respuestas = {}
        if "comentarios_parametro" not in st.session_state:
            st.session_state.comentarios_parametro = {}

        tabs = st.tabs([label for _, label, _ in parametros])
        for (clave, label, pregunta), tab in zip(parametros, tabs):
            with tab:
                st.markdown(f"**{pregunta}**")
                hay_no_conforme = False
                for i in range(n):
                    key = f"resp_{clave}_{i}"
                    if key not in st.session_state.respuestas:
                        st.session_state.respuestas[key] = "Conforme"
                    valor_actual = st.session_state.respuestas[key]
                    idx_default = 0 if valor_actual == "Conforme" else 1
                    st.session_state.respuestas[key] = st.radio(
                        f"Muestra {i + 1}",
                        ["Conforme", "No conforme"],
                        index=idx_default,
                        horizontal=True,
                        key=f"widget_{key}",
                    )
                    if st.session_state.respuestas[key] == "No conforme":
                        hay_no_conforme = True

                if hay_no_conforme:
                    st.session_state.comentarios_parametro[clave] = st.text_area(
                        f"Comentario / corrección para '{label}' (hay al menos 1 muestra No conforme)",
                        value=st.session_state.comentarios_parametro.get(clave, ""),
                        key=f"comentario_{clave}",
                    )
                else:
                    st.session_state.comentarios_parametro[clave] = ""

    col1, col2 = st.columns(2)
    with col1:
        st.button("⬅ Atrás", on_click=go_back)
    with col2:
        if st.button("Siguiente ➜", type="primary"):
            st.session_state.parametros_muestra = parametros
            go_next()
            st.rerun()

# ============================================================================
# PASO 7 - CONCLUSIÓN DEL REGISTRO
# ============================================================================
elif st.session_state.step == 7:
    st.header("6️⃣ Conclusión del registro")

    conclusion = st.radio(
        "Conclusión",
        ["Conforme", "No conforme"],
        index=0 if st.session_state.get("conclusion", "Conforme") == "Conforme" else 1,
        horizontal=True,
    )

    col1, col2 = st.columns(2)
    with col1:
        st.button("⬅ Atrás", on_click=go_back)
    with col2:
        if st.button("Generar resumen ➜", type="primary"):
            st.session_state.conclusion = conclusion
            go_next()
            st.rerun()

# ============================================================================
# PASO 8 - RESUMEN FINAL Y EXPORTACIÓN
# ============================================================================
elif st.session_state.step == 8:
    st.header("7️⃣ Resumen final")

    fila = st.session_state.fila_actividad
    n = st.session_state.n_muestras
    parametros = st.session_state.parametros_muestra
    claves_activas = [clave for clave, _, _ in parametros]

    st.subheader("Datos del registro")
    resumen_info = pd.DataFrame(
        {
            "Campo": [
                "Equipo de calidad", "Turno", "Fecha de producción", "Cliente", "Área",
                "Línea HACCP", "Producto", "Componente", "Actividad",
                "¿Aplica batch?", "Tamaño de batch", "Letra código muestreo",
                "N° de muestras", "Conclusión",
            ],
            "Valor": [
                st.session_state.responsable, st.session_state.turno,
                st.session_state.fecha_produccion.strftime("%d/%m/%Y"),
                CLIENTE_FIJO, AREA_FIJA,
                st.session_state.linea_haccp, st.session_state.producto,
                fila["componente"], fila["actividad"],
                st.session_state.aplica_batch,
                st.session_state.batch_size if st.session_state.batch_size else "No aplica",
                st.session_state.letra_codigo if st.session_state.letra_codigo else "No aplica",
                n, st.session_state.conclusion,
            ],
        }
    )
    st.dataframe(resumen_info, use_container_width=True, hide_index=True)

    if parametros:
        st.subheader("Resultados por parámetro (muestras)")
        conteo = {}
        for clave, label, _ in parametros:
            valores = [st.session_state.respuestas[f"resp_{clave}_{i}"] for i in range(n)]
            conteo[label] = {
                "Conforme": valores.count("Conforme"),
                "No conforme": valores.count("No conforme"),
            }
        st.dataframe(pd.DataFrame(conteo).T, use_container_width=True)

    # ------------------------------------------------------------------
    # Armado de la fila exportable (UNA fila por registro, con columnas
    # fijas ampliadas Muestra 1..8 por cada parámetro posible)
    # ------------------------------------------------------------------
    fila_export = {
        "Fecha": st.session_state.fecha_produccion.strftime("%d/%m/%Y"),
        "Turno": st.session_state.turno,
        "Cliente": CLIENTE_FIJO,
        "Área": AREA_FIJA,
        "Línea HACCP": st.session_state.linea_haccp,
        "Producto": st.session_state.producto,
        "Componente": fila["componente"],
        "Actividad": fila["actividad"],
        "¿Aplica Batch?": st.session_state.aplica_batch,
        "Tamaño de Batch": st.session_state.batch_size if st.session_state.batch_size else "",
        "Letra código": st.session_state.letra_codigo if st.session_state.letra_codigo else "",
        "N° de muestras": n,
    }

    for clave, label in ALL_PARAM_DEFS:
        valores_muestra = (
            [st.session_state.respuestas[f"resp_{clave}_{i}"] for i in range(n)]
            if clave in claves_activas else []
        )
        for i in range(MAX_MUESTRAS_COLUMNAS):
            col_name = f"{label} M{i + 1}"
            fila_export[col_name] = valores_muestra[i] if i < len(valores_muestra) else ""

    fila_export["Conclusión"] = st.session_state.conclusion
    fila_export["Iniciales"] = iniciales(st.session_state.responsable)

    export_df = pd.DataFrame([fila_export])

    with st.expander("Ver la fila completa que se exportará"):
        st.dataframe(export_df, use_container_width=True, hide_index=True)

    csv_bytes = export_df.to_csv(index=False).encode("utf-8-sig")
    excel_buffer = io.BytesIO()
    with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
        export_df.to_excel(writer, index=False, sheet_name="Registro")
    excel_buffer.seek(0)

    st.divider()
    col1, col2, col3 = st.columns(3)
    with col1:
        st.button("⬅ Atrás", on_click=go_back)
    with col2:
        st.download_button(
            "⬇ Descargar CSV",
            data=csv_bytes,
            file_name=f"registro_PI_{st.session_state.producto}_{st.session_state.fecha_produccion}.csv",
            mime="text/csv",
        )
    with col3:
        st.download_button(
            "⬇ Descargar Excel",
            data=excel_buffer,
            file_name=f"registro_PI_{st.session_state.producto}_{st.session_state.fecha_produccion}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    st.divider()
    if st.button("🔄 Nuevo registro"):
        for key in list(st.session_state.keys()):
            del st.session_state[key]
        st.rerun()
