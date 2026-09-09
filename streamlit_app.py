"""
Registro de Control de Productos en Proceso (B2B)
--------------------------------------------------
App Streamlit para registrar, por PRODUCTO y ACTIVIDAD (1 registro = 1 actividad),
los parametros de control habilitados segun la matriz de especificaciones
(hoja PROCESO-B2B-STB del Excel MA-PL-019), siguiendo el formato de
"CONTROL PRODUCTOS EN PROCESO" (MA-FR-030).

Reglas clave:
- Un registro = un producto + una actividad (no se mezclan actividades).
- Solo se piden los parametros que la matriz tiene habilitados para esa
  actividad especifica (si en el Excel el valor es "-" o esta vacio, el
  parametro no se pide).
- Antes de pedir muestras, se pregunta si la actividad considera BATCH:
    - Si aplica batch -> se pide el tamaño del lote y el numero de muestras
      se calcula con MIL-STD-105E, Nivel de Inspeccion Especial S-2,
      Tabla 2B (inspeccion rigurosa/tightened), AQL 4.0%.
    - Si no aplica batch -> se registra 1 sola muestra general.
- Se evalua conforme / no conforme comparando contra la especificacion
  cuando esta es numerica (ej "75 +/- 5", "7 a 7.5", "0 - 4"); si la
  especificacion es textual (ej "LIBRE DE MATERIAL EXTRAÑO") se marca
  conforme/no conforme de forma manual.
- Cada registro guardado se acumula en un HISTORIAL (tabla de muestras)
  descargable en Excel/CSV con el mismo esqueleto de columnas de MA-FR-030.

Requisitos: streamlit, pandas, openpyxl, xlsxwriter
    pip install streamlit pandas openpyxl xlsxwriter
Ejecutar:
    streamlit run app.py
"""

import re
import io
from datetime import date, datetime

import pandas as pd
import streamlit as st

# --------------------------------------------------------------------------
# CONFIGURACION GENERAL
# --------------------------------------------------------------------------

st.set_page_config(page_title="Control de Productos en Proceso", layout="wide")

MASTER_XLSX_PATH = "MA-PL-019_PLAN_CALIDAD_DE_PRODUCTOS.xlsx"  # subir al repo junto al app.py
MASTER_SHEET = "PROCESO-B2B-STB"

# Nombre interno -> etiqueta visible + si requiere spec numerica o es texto/checklist
PARAM_DEFS = [
    ("peso",          "Peso (g)",                         "numeric"),
    ("diametro",      "Diametro (cm)",                     "numeric"),
    ("altura",        "Altura / Espesor (cm)",              "numeric"),
    ("temperatura",   "Temperatura (C)",                    "numeric"),
    ("tamizado",      "Tamizado",                           "text"),
    ("tiempo_mezcla", "Tiempo de mezcla (min)",              "numeric"),
    ("decorado",      "Decorado / detalle de forma",         "text"),
    ("brix",          "Brix",                               "numeric"),
    ("organolepticas","Caracteristicas organolepticas",      "text"),
]

# --------------------------------------------------------------------------
# MIL-STD-105E — Nivel de Inspeccion Especial S-2 / Tabla 2B (rigurosa) / AQL 4.0%
# --------------------------------------------------------------------------

# Rangos de tamaño de lote -> letra codigo, columna Nivel Especial S-2
# (tabla general de letras codigo de MIL-STD-105E)
LETRA_CODIGO_S2 = [
    (2, 8, "A"),
    (9, 15, "A"),
    (16, 25, "A"),
    (26, 50, "B"),
    (51, 90, "B"),
    (91, 150, "B"),
    (151, 280, "C"),
    (281, 500, "C"),
    (501, 1200, "C"),
    (1201, 3200, "D"),
    (3201, 10000, "D"),
    (10001, 35000, "D"),
    (35001, 150000, "E"),
    (150001, 500000, "E"),
    (500001, float("inf"), "E"),
]

# Tamaño de muestra por letra codigo (Tabla II — igual para normal/rigurosa/reducida)
TAMANO_MUESTRA_POR_LETRA = {
    "A": 2, "B": 3, "C": 5, "D": 8, "E": 13, "F": 20, "G": 32, "H": 50,
    "J": 80, "K": 125, "L": 200, "M": 315, "N": 500, "P": 800, "Q": 1250, "R": 2000,
}

# Ac/Re (aceptacion/rechazo) para inspeccion RIGUROSA (Tabla 2B), AQL 4.0%
# "-" indica que a esa letra le corresponde usar el plan de la siguiente flecha
# hacia abajo (segun la tabla original); aqui se deja el plan efectivo ya resuelto
# para las letras que puede producir S-2 (A-E).
AC_RE_RIGUROSA_AQL_4_0 = {
    "A": (0, 1),
    "B": (0, 1),
    "C": (0, 1),
    "D": (1, 2),
    "E": (1, 2),
}


def calcular_muestreo_mil_std_105e(tamano_lote: int):
    """Devuelve (letra_codigo, n_muestras, ac, re) para Nivel S-2,
    Tabla 2B (rigurosa), AQL 4.0%, segun el tamaño de lote (batch)."""
    letra = None
    for low, high, l in LETRA_CODIGO_S2:
        if low <= tamano_lote <= high:
            letra = l
            break
    if letra is None:
        letra = "E"
    n = TAMANO_MUESTRA_POR_LETRA[letra]
    ac, re = AC_RE_RIGUROSA_AQL_4_0[letra]
    return letra, n, ac, re


COL_MAP = {
    # columna Excel (indice 0-based dentro de la fila) -> clave interna
    0: "producto", 1: "linea_producto", 2: "linea_haccp", 3: "sub_producto",
    4: "actividad", 5: "tipo_producto",
    6: "peso", 7: "diametro", 8: "altura", 9: "temperatura", 10: "tamizado",
    11: "tiempo_mezcla", 12: "decorado", 13: "brix", 14: "organolepticas",
    15: "estado_material", 16: "responsable", 17: "correccion_ref",
}


# --------------------------------------------------------------------------
# CARGA Y PARSEO DE LA MATRIZ MAESTRA
# --------------------------------------------------------------------------

@st.cache_data
def cargar_matriz(path: str) -> pd.DataFrame:
    """Lee la hoja PROCESO-B2B-STB, hace forward-fill de las columnas
    fusionadas (producto, linea, linea_haccp) y devuelve un registro
    por fila = producto + actividad."""
    raw = pd.read_excel(path, sheet_name=MASTER_SHEET, header=None, skiprows=7)
    filas = []
    cur_producto = cur_linea_producto = cur_linea_haccp = None
    for _, row in raw.iterrows():
        vals = row.tolist()
        if not any(pd.notna(v) for v in vals[:18]):
            continue
        if pd.notna(vals[0]):
            cur_producto = vals[0]
        if pd.notna(vals[1]):
            cur_linea_producto = vals[1]
        if pd.notna(vals[2]):
            cur_linea_haccp = vals[2]
        if pd.isna(vals[4]):  # sin actividad -> fila vacia de relleno
            continue
        registro = {"producto": cur_producto, "linea_producto": cur_linea_producto,
                    "linea_haccp": cur_linea_haccp}
        for idx, clave in COL_MAP.items():
            if clave in ("producto", "linea_producto", "linea_haccp"):
                continue
            registro[clave] = vals[idx] if idx < len(vals) else None
        filas.append(registro)
    return pd.DataFrame(filas)


def parametros_habilitados(fila: pd.Series):
    """Devuelve la lista de (clave, etiqueta, tipo, spec_original) habilitados
    para esta actividad: el Excel trae '-' o vacio cuando no aplica."""
    habilitados = []
    for clave, etiqueta, tipo in PARAM_DEFS:
        spec = fila.get(clave)
        if spec is None or (isinstance(spec, str) and spec.strip() in ("-", "")):
            continue
        habilitados.append((clave, etiqueta, tipo, spec))
    return habilitados


def parsear_rango(spec):
    """Intenta extraer un rango numerico [low, high] de la especificacion.
    Soporta '75 +/- 5', '7 a 7.5', '0 - 4', numero suelto. Devuelve None si
    la especificacion es puramente textual (ej 'LIBRE DE MATERIAL EXTRAÑO')."""
    if spec is None:
        return None
    if isinstance(spec, (int, float)):
        return (spec, spec)
    s = str(spec).strip()

    m = re.match(r"^(-?\d+(?:\.\d+)?)\s*\+/-\s*(\d+(?:\.\d+)?)$", s)
    if m:
        base, tol = float(m.group(1)), float(m.group(2))
        return (base - tol, base + tol)

    m = re.match(r"^(-?\d+(?:\.\d+)?)\s*a\s*(-?\d+(?:\.\d+)?)$", s, re.IGNORECASE)
    if m:
        return (float(m.group(1)), float(m.group(2)))

    m = re.match(r"^(-?\d+(?:\.\d+)?)\s*-\s*(-?\d+(?:\.\d+)?)$", s)
    if m:
        return (float(m.group(1)), float(m.group(2)))

    m = re.match(r"^(-?\d+(?:\.\d+)?)$", s)
    if m:
        v = float(m.group(1))
        return (v, v)

    return None  # texto puro


def evaluar_conformidad(valor, spec):
    """True/False/None (None = no se pudo evaluar automaticamente, requiere
    marca manual)."""
    rango = parsear_rango(spec)
    if rango is None or valor is None or valor == "":
        return None
    try:
        v = float(valor)
    except (TypeError, ValueError):
        return None
    low, high = rango
    return low <= v <= high


# --------------------------------------------------------------------------
# ESTADO DE SESION
# --------------------------------------------------------------------------

if "historial" not in st.session_state:
    st.session_state.historial = []  # lista de dicts, un dict por REGISTRO guardado
if "paso" not in st.session_state:
    st.session_state.paso = 1


def reiniciar_wizard():
    st.session_state.paso = 1
    for k in list(st.session_state.keys()):
        if k.startswith("form_"):
            del st.session_state[k]


# --------------------------------------------------------------------------
# CARGA DE DATOS
# --------------------------------------------------------------------------

st.title("Control de productos en proceso — Registro por actividad")

try:
    matriz = cargar_matriz(MASTER_XLSX_PATH)
except FileNotFoundError:
    st.warning(
        f"No se encontro '{MASTER_XLSX_PATH}' junto al app. Sube el Excel maestro "
        "para continuar (debe tener la hoja 'PROCESO-B2B-STB')."
    )
    subido = st.file_uploader("Subir Excel maestro (MA-PL-019)", type=["xlsx"])
    if subido is None:
        st.stop()
    matriz = cargar_matriz(subido)

productos = sorted(matriz["producto"].dropna().unique().tolist())

# --------------------------------------------------------------------------
# WIZARD
# --------------------------------------------------------------------------

st.progress(min(st.session_state.paso, 6) / 6)

# ---- Paso 1: datos generales -----------------------------------------
if st.session_state.paso == 1:
    st.subheader("1. Datos generales del registro")
    c1, c2, c3 = st.columns(3)
    with c1:
        fecha = st.date_input("Fecha", value=date.today(), key="form_fecha")
    with c2:
        area = st.text_input("Area", key="form_area")
    with c3:
        cliente = st.text_input("Cliente", value="Starbucks", key="form_cliente")

    if st.button("Continuar", type="primary"):
        if not area or not cliente:
            st.error("Completa Area y Cliente antes de continuar.")
        else:
            st.session_state.paso = 2
            st.rerun()

# ---- Paso 2: producto + actividad --------------------------------------
elif st.session_state.paso == 2:
    st.subheader("2. Producto y actividad a registrar")
    producto = st.selectbox("Producto", productos, key="form_producto")

    actividades_producto = matriz[matriz["producto"] == producto]
    opciones_actividad = actividades_producto["actividad"].tolist()
    actividad = st.selectbox(
        "Actividad / etapa del proceso (solo se registra UNA por sesion)",
        opciones_actividad, key="form_actividad",
    )

    fila_sel = actividades_producto[actividades_producto["actividad"] == actividad].iloc[0]
    st.session_state.form_fila = fila_sel.to_dict()

    st.caption(f"Linea HACCP: {fila_sel['linea_haccp']}  |  Tipo: {fila_sel.get('tipo_producto', '-')}")

    habilitados = parametros_habilitados(fila_sel)
    if habilitados:
        st.markdown("**Parametros habilitados para esta actividad:**")
        st.table(pd.DataFrame(
            [(etq, spec) for _, etq, _, spec in habilitados],
            columns=["Parametro", "Especificacion"],
        ))
    else:
        st.info("Esta actividad no tiene parametros numericos/organolepticos habilitados en la matriz.")

    c1, c2 = st.columns(2)
    with c1:
        if st.button("Atras"):
            st.session_state.paso = 1
            st.rerun()
    with c2:
        if st.button("Continuar", type="primary"):
            st.session_state.paso = 3
            st.rerun()

# ---- Paso 3: aplica batch? ---------------------------------------------
elif st.session_state.paso == 3:
    st.subheader("3. Muestreo: ¿la actividad considera batch?")
    st.write(
        "Si la actividad se revisa por lote/batch (ej. decorado de galleta, "
        "con muchas unidades por lote), marca que SI considera batch: el "
        "numero de muestras se calcula con MIL-STD-105E (Nivel Especial S-2, "
        "Tabla 2B rigurosa, AQL 4.0%) segun el tamaño del lote. Si es un "
        "proceso general sin batch definido (ej. tamizado del jugo de "
        "naranja), marca que NO y se registra 1 sola muestra representativa."
    )
    aplica_batch = st.radio(
        "¿Esta actividad considera batch?",
        ["Si, considera batch (calcular muestreo MIL-STD-105E)", "No considera batch (1 muestra)"],
        key="form_aplica_batch_radio",
    )
    st.session_state.form_aplica_batch = aplica_batch.startswith("Si")

    if st.session_state.form_aplica_batch:
        tamano_lote = st.number_input(
            "Tamaño del batch (número de unidades del lote)",
            min_value=2, step=1, value=90, key="form_tamano_lote",
        )
        letra, n_muestras, ac, re = calcular_muestreo_mil_std_105e(int(tamano_lote))
        st.session_state.form_letra_codigo = letra
        st.session_state.form_ac = ac
        st.session_state.form_re = re
        st.info(
            f"MIL-STD-105E · Nivel S-2 · Tabla 2B (rigurosa) · AQL 4.0%\n\n"
            f"- Letra codigo: **{letra}**\n"
            f"- Tamaño de muestra: **{n_muestras}**\n"
            f"- Criterio Ac/Re: aceptar con **{ac}** o menos no conformes, "
            f"rechazar con **{re}** o mas."
        )
    else:
        n_muestras = 1
        st.session_state.form_letra_codigo = None
        st.session_state.form_ac = None
        st.session_state.form_re = None

    st.session_state.form_n_muestras = n_muestras
    st.caption(f"Se pedira{'n' if n_muestras > 1 else ''} {n_muestras} muestra{'s' if n_muestras > 1 else ''}.")

    c1, c2 = st.columns(2)
    with c1:
        if st.button("Atras", key="atras3"):
            st.session_state.paso = 2
            st.rerun()
    with c2:
        if st.button("Continuar", type="primary", key="cont3"):
            st.session_state.paso = 4
            st.rerun()

# ---- Paso 4: registro de parametros -------------------------------------
elif st.session_state.paso == 4:
    st.subheader("4. Registro de parametros de control")
    fila_sel = pd.Series(st.session_state.form_fila)
    habilitados = parametros_habilitados(fila_sel)
    n_muestras = st.session_state.form_n_muestras

    if not habilitados:
        st.info("No hay parametros que registrar para esta actividad; puedes continuar.")

    valores = {}  # clave -> lista de valores por muestra
    conformidades = {}  # clave -> lista de bool/None por muestra

    MUESTRAS_POR_FILA = 6  # evita columnas demasiado angostas cuando n_muestras es grande

    for clave, etiqueta, tipo, spec in habilitados:
        st.markdown(f"**{etiqueta}**  ·  especificacion: `{spec}`")
        vals_param = []
        conf_param = []
        for inicio in range(0, n_muestras, MUESTRAS_POR_FILA):
            indices_fila = list(range(inicio, min(inicio + MUESTRAS_POR_FILA, n_muestras)))
            cols = st.columns(len(indices_fila))
            for col, i in zip(cols, indices_fila):
                with col:
                    label_muestra = f"Muestra {i + 1}" if n_muestras > 1 else "Valor"
                    if tipo == "numeric":
                        v = st.number_input(
                            label_muestra, key=f"form_{clave}_{i}", format="%.2f",
                            step=0.1,
                        )
                        conforme = evaluar_conformidad(v, spec)
                    else:
                        v = st.text_input(label_muestra, key=f"form_{clave}_{i}")
                        auto = evaluar_conformidad(v, spec)
                        if auto is None:
                            conforme_manual = st.selectbox(
                                "Conforme?", ["Conforme", "No conforme"],
                                key=f"form_{clave}_conf_{i}",
                            )
                            conforme = conforme_manual == "Conforme"
                        else:
                            conforme = auto
                    vals_param.append(v)
                    conf_param.append(conforme)
                    if conforme is False:
                        st.error("No conforme")
                    elif conforme is True:
                        st.success("Conforme")
        valores[clave] = vals_param
        conformidades[clave] = conf_param

    st.session_state.form_valores = valores
    st.session_state.form_conformidades = conformidades

    c1, c2 = st.columns(2)
    with c1:
        if st.button("Atras", key="atras4"):
            st.session_state.paso = 3
            st.rerun()
    with c2:
        if st.button("Continuar", type="primary", key="cont4"):
            st.session_state.paso = 5
            st.rerun()

# ---- Paso 5: checklist + responsable + observaciones ---------------------
elif st.session_state.paso == 5:
    st.subheader("5. Verificaciones adicionales y cierre del registro")

    c1, c2 = st.columns(2)
    with c1:
        empaque_ok = st.radio(
            "Correcto manejo de empaques", ["Si (Correcto)", "No (Incorrecto)"],
            key="form_empaque_ok",
        )
    with c2:
        estado_material = st.session_state.form_fila.get("estado_material", "-")
        st.text_input(
            "Estado de material/equipo/herramienta esperado (referencia matriz)",
            value=str(estado_material), disabled=True,
        )
        material_ok = st.radio(
            "Se verifico el estado del material/equipo/herramienta",
            ["Si, conforme", "No, con observaciones"], key="form_material_ok",
        )

    responsable = st.text_input("Responsable de calidad", key="form_responsable")
    observaciones = st.text_area(
        "Observaciones / acciones correctivas", key="form_observaciones",
        placeholder="Dejar vacio si no hay no conformidades",
    )

    c1, c2 = st.columns(2)
    with c1:
        if st.button("Atras", key="atras5"):
            st.session_state.paso = 4
            st.rerun()
    with c2:
        if st.button("Guardar registro", type="primary", key="guardar5"):
            if not responsable:
                st.error("Ingresa el responsable de calidad antes de guardar.")
            else:
                fila_sel = st.session_state.form_fila
                valores = st.session_state.form_valores
                conformidades = st.session_state.form_conformidades
                n_muestras = st.session_state.form_n_muestras

                cualquier_no_conforme = any(
                    c is False for lst in conformidades.values() for c in lst
                ) or empaque_ok.startswith("No") or material_ok.startswith("No")

                registro = {
                    "fecha": st.session_state.form_fecha,
                    "area": st.session_state.form_area,
                    "cliente": st.session_state.form_cliente,
                    "producto": fila_sel["producto"],
                    "actividad": fila_sel["actividad"],
                    "linea_haccp": fila_sel["linea_haccp"],
                    "aplica_batch": st.session_state.form_aplica_batch,
                    "tamano_lote": st.session_state.get("form_tamano_lote"),
                    "letra_codigo_mil_std": st.session_state.get("form_letra_codigo"),
                    "ac_mil_std": st.session_state.get("form_ac"),
                    "re_mil_std": st.session_state.get("form_re"),
                    "n_muestras": n_muestras,
                    "empaque_ok": empaque_ok.startswith("Si"),
                    "material_ok": material_ok.startswith("Si"),
                    "responsable": responsable,
                    "observaciones": observaciones,
                    "conforme_general": not cualquier_no_conforme,
                    "registrado_en": datetime.now(),
                }
                # aplanar muestras: peso_1, peso_2, peso_3, ...
                for clave, lst in valores.items():
                    for i, v in enumerate(lst, start=1):
                        registro[f"{clave}_{i}"] = v
                for clave, lst in conformidades.items():
                    for i, c in enumerate(lst, start=1):
                        registro[f"{clave}_{i}_conforme"] = c

                st.session_state.historial.append(registro)
                st.success("Registro guardado en el historial.")
                st.session_state.paso = 6
                st.rerun()

# ---- Paso 6: resumen y opcion de nuevo registro ---------------------------
elif st.session_state.paso == 6:
    st.subheader("6. Registro guardado")
    ultimo = st.session_state.historial[-1]
    st.json({k: str(v) for k, v in ultimo.items()})

    c1, c2 = st.columns(2)
    with c1:
        if st.button("Registrar otra actividad (mismo producto u otro)", type="primary"):
            reiniciar_wizard()
            st.rerun()
    with c2:
        if st.button("Ir al historial completo"):
            st.session_state.paso = 7
            st.rerun()

# ---- Paso 7: historial + exportacion --------------------------------------
if st.session_state.paso == 7 or (st.session_state.historial and st.session_state.paso not in range(1, 7)):
    pass

st.divider()
st.subheader("Historial de registros (tabla de muestras)")

if not st.session_state.historial:
    st.caption("Aun no hay registros guardados en esta sesion.")
else:
    df_hist = pd.DataFrame(st.session_state.historial)
    st.dataframe(df_hist, use_container_width=True)

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        df_hist.to_excel(writer, index=False, sheet_name="Historial")
    st.download_button(
        "Descargar historial (Excel)", data=buffer.getvalue(),
        file_name=f"historial_control_proceso_{date.today().isoformat()}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    st.download_button(
        "Descargar historial (CSV)", data=df_hist.to_csv(index=False).encode("utf-8"),
        file_name=f"historial_control_proceso_{date.today().isoformat()}.csv",
        mime="text/csv",
    )

    if st.button("Iniciar un registro nuevo"):
        reiniciar_wizard()
        st.rerun()
