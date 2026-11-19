import html
import io
import re
import unicodedata
from datetime import datetime

import openpyxl
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Validador de archivos Excel",
    page_icon="📋",
    layout="wide"
)

# Fuentes y colores base vienen de .streamlit/config.toml.
# Este CSS solo da forma a los bloques HTML propios de la app.
st.markdown("""
<style>
    :root {
        --superficie: #FFFFFF;
        --tinta: #172033;
        --tinta-2: #566173;
        --linea: #D5DBE3;
        --hover: #E2E7EC;
        --acento: #0E6973;
        --ok: #1D7A4E;
        --error: #B42318;
        --error-suave: #F8DEDA;
        --claro: #D52B1E;
        --etb: #0057B8;
    }
    [data-testid="stMainBlockContainer"], .block-container {
        max-width: 1320px; padding-top: 2rem; padding-bottom: 3rem;
    }
    header[data-testid="stHeader"] { background: transparent; }

    /* ── Barra superior ── */
    .encabezado h1 {
        font-size: 1.75rem; line-height: 1.15; letter-spacing: -0.02em; margin: 0; padding: 0;
    }
    .encabezado p { color: var(--tinta-2); margin: 0.2rem 0 0 0; }
    .marca {
        display: inline-block; width: 0.6rem; height: 0.6rem; border-radius: 50%;
        margin-right: 0.4rem; vertical-align: 0.05rem;
    }
    .nota-operador { color: var(--tinta-2); font-size: 0.9rem; margin: 0; }
    [data-testid="stFileUploaderDropzone"] {
        background: var(--superficie); border: 1.5px dashed #AEB8C5;
    }

    /* ── Estado vacío ── */
    .vacio {
        background: var(--superficie); border-radius: 10px; padding: 1.25rem 1.5rem;
        display: grid; grid-template-columns: minmax(180px, 240px) 1fr; gap: 1.5rem; align-items: start;
    }
    .vacio h3 { font-size: 1.05rem; margin: 0 0 0.3rem 0; padding: 0; }
    .vacio p { color: var(--tinta-2); font-size: 0.9rem; margin: 0; }

    /* ── Chips de columnas ── */
    .chips { display: flex; flex-wrap: wrap; gap: 0.35rem; }
    .chip {
        font-size: 0.82rem; padding: 0.15rem 0.55rem; border-radius: 999px;
        background: var(--superficie); border: 1px solid var(--linea); white-space: nowrap;
    }
    .chip.falta { border-color: var(--error); color: var(--error); font-weight: 700; }
    .chip.opcional { border-style: dashed; color: var(--tinta-2); }

    /* ── Tabla de reglas (dentro del panel desplegable) ── */
    .tabla-reglas { width: 100%; border-collapse: collapse; font-size: 0.88rem; }
    .tabla-reglas, .tabla-reglas th, .tabla-reglas td {
        border-left: none !important; border-right: none !important; border-top: none !important;
    }
    .tabla-reglas tr, .tabla-reglas thead { background: transparent !important; }
    .tabla-reglas th {
        text-align: left; font-weight: 400; color: var(--tinta-2);
        border-bottom: 1px solid var(--linea); padding: 0.3rem 0.75rem 0.3rem 0;
    }
    .tabla-reglas td {
        border-bottom: 1px solid var(--linea); padding: 0.4rem 0.75rem 0.4rem 0; vertical-align: top;
    }
    .tabla-reglas td:first-child { font-weight: 700; white-space: nowrap; }
    .tabla-reglas tr.opcional td { color: var(--tinta-2); }

    /* ── Veredicto: una franja horizontal ── */
    .st-key-veredicto_ok, .st-key-veredicto_err {
        background: var(--superficie); border-radius: 10px;
        border-left: 7px solid var(--ok); padding: 1.25rem 1.5rem 1rem 1.5rem;
    }
    .st-key-veredicto_err { border-left-color: var(--error); }
    .estado h2 {
        font-size: 1.6rem; line-height: 1.15; letter-spacing: -0.02em; margin: 0 0 0.3rem 0; padding: 0;
    }
    .estado p { color: var(--tinta-2); font-size: 0.95rem; margin: 0; }
    .cifras { display: grid; grid-template-columns: repeat(4, auto); gap: 0 1.75rem; justify-content: start; }
    .cifras div { display: flex; flex-direction: column; }
    .cifras b { font-size: 1.5rem; line-height: 1.2; font-variant-numeric: tabular-nums; }
    .cifras span { font-size: 0.82rem; color: var(--tinta-2); }
    .cifras .rojo b { color: var(--error); }
    .barra {
        display: flex; height: 8px; border-radius: 4px; overflow: hidden;
        background: var(--error-suave); margin: 0.9rem 0 0.4rem 0;
    }
    .barra .ok { background: var(--ok); }
    .leyenda { display: flex; gap: 0.3rem 1.5rem; flex-wrap: wrap; font-size: 0.82rem; color: var(--tinta-2); padding-bottom: 0.35rem; }
    .leyenda span::before {
        content: ""; display: inline-block; width: 0.6rem; height: 0.6rem;
        border-radius: 2px; margin-right: 0.35rem; vertical-align: -0.03rem;
    }
    .leyenda .l-ok::before { background: var(--ok); }
    .leyenda .l-err::before { background: var(--error); }

    /* ── Explorador de errores: lista de campos a la izquierda ── */
    .titulo-seccion { font-size: 1.15rem; font-weight: 700; margin: 0.5rem 0 0 0; }
    .st-key-campo_sel, .st-key-campo_sel [data-testid="stRadio"], .st-key-campo_sel [data-testid="stRadio"] > div { width: 100%; }
    .st-key-campo_sel [role="radiogroup"] { display: flex; flex-direction: column; align-items: stretch; gap: 2px; width: 100%; }
    .st-key-campo_sel label {
        width: 100%; margin: 0; padding: 0.45rem 0.7rem; border-radius: 6px; cursor: pointer;
    }
    .st-key-campo_sel label:hover { background: var(--hover); }
    .st-key-campo_sel label:has(input:checked) {
        background: var(--superficie); box-shadow: inset 3px 0 0 var(--acento);
    }
    .st-key-campo_sel label:has(input:checked) p { font-weight: 700; }
    .st-key-campo_sel label > div:first-of-type:not(:last-child) { display: none; }
    .regla { color: var(--tinta-2); font-size: 0.9rem; margin: 0; }

    [data-testid="stPopoverButton"] p, .stDownloadButton button p { white-space: nowrap; }
    .estado { min-width: 260px; }

    @media (max-width: 640px) {
        .st-key-barra { flex-wrap: wrap; }
        .st-key-titulo { flex: 1 1 100% !important; min-width: 100%; }
        .encabezado h1 { font-size: 1.5rem; }
        .vacio { grid-template-columns: 1fr; gap: 0.75rem; }
        .cifras { grid-template-columns: repeat(2, auto); gap: 0.5rem 1.5rem; }
    }
</style>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# FUNCIONES COMUNES
# ══════════════════════════════════════════════════════════════════════════════

def not_empty(val):
    return not (pd.isna(val) or str(val).strip() in ("", "nan", "None"))

def no_special_chars(val):
    if not not_empty(val): return None
    return bool(re.match(r'^[A-Za-záéíóúÁÉÍÓÚñÑüÜ0-9\s\-_.]+$', str(val).strip()))

def no_special_chars_no_spaces(val):
    if not not_empty(val): return None
    v = str(val).strip()
    if " " in v: return False
    return bool(re.match(r'^[A-Za-záéíóúÁÉÍÓÚñÑüÜ0-9\-_.]+$', v))

def no_spaces_only(val):
    if not not_empty(val): return None
    return " " not in str(val).strip()

def valid_blockname(val):
    if not not_empty(val): return None
    return str(val).strip() in ("N_Casa", "N_Edificio")

def valid_latlon(val, must_be_negative=False):
    if not not_empty(val): return False, "vacío"
    v = str(val).strip()
    if " " in v: return False, "contiene espacios"
    if "," in v: return False, "usa coma en lugar de punto decimal"
    try:
        num = float(v)
    except ValueError:
        return False, "no es un número válido"
    if must_be_negative and num >= 0: return False, "debe ser un valor negativo"
    if "." not in v: return False, "no tiene punto decimal"
    decimals = v.split(".")[1]
    if len(decimals) < 6: return False, f"solo tiene {len(decimals)} decimal(es), necesita mínimo 6"
    return True, "ok"

def valid_puerto_olt(val):
    """
    PUERTO OLT debe tener exactamente el formato N/N/N
    donde N es uno o más dígitos. Ejemplo válido: 1/3/3, 0/1/15, 10/2/8
    Cualquier otro valor (entero, fecha, texto, etc.) es inválido.
    Retorna (True, "ok") si es válido, (False, motivo) si no.
    """
    if not not_empty(val):
        return False, "campo vacío"
    # Si pandas lo convirtió a Timestamp antes de llegar aquí
    if isinstance(val, (datetime, pd.Timestamp)):
        return False, f"fue interpretado como fecha por Excel: '{val}'"
    v = str(val).strip()
    # Formato esperado: solo dígitos separados por exactamente dos barras
    if re.fullmatch(r'\d+/\d+/\d+', v):
        return True, "ok"
    # Diagnosticar el motivo del error para mensaje útil
    if re.fullmatch(r'\d+', v):
        return False, f"es un número entero '{v}', debe ser N/N/N (ej: 1/3/3)"
    if re.search(r'[a-zA-Z]', v):
        return False, f"contiene letras '{v}', debe ser solo dígitos separados por '/' (ej: 1/3/3)"
    if '-' in v or re.search(r'\d{4}', v):
        return False, f"parece una fecha '{v}', debe ser N/N/N (ej: 1/3/3)"
    return False, f"formato inválido '{v}', se esperaba N/N/N (ej: 1/3/3)"

def es_numerico(val):
    if not not_empty(val): return None
    try:
        float(str(val).strip()); return True
    except ValueError:
        return False

def solo_letras_espacios(val):
    if not not_empty(val): return None
    return bool(re.match(r'^[A-Za-záéíóúÁÉÍÓÚñÑüÜ\s]+$', str(val).strip()))


# ── Errores de fórmula de Excel (#REF!, #N/A, ...) ───────────────────────────

# Inglés: así se guardan en el archivo, sin importar el idioma de Excel.
# Español: así quedan si alguien copió el resultado como texto en un Excel en español.
ERRORES_EXCEL = (
    "#REF!", "#N/A", "#VALUE!", "#DIV/0!", "#NAME?", "#NULL!", "#NUM!",
    "#SPILL!", "#CALC!", "#FIELD!", "#BLOCKED!", "#CONNECT!", "#BUSY!",
    "#UNKNOWN!", "#GETTING_DATA",
    "#¡REF!", "#N/D", "#¡VALOR!", "#¡DIV/0!", "#¿NOMBRE?", "#¡NULO!",
    "#¡NUM!", "#¡DESBORDAMIENTO!", "#¡CALC!",
)

# DISTRIBUCION es columna obligatoria en CLARO. Por ahora sus celdas pueden ir vacías;
# cambiar a False si deben tener valor siempre (las vacías se reportarán como error).
DISTRIBUCION_PERMITE_VACIOS = True

def error_excel_en(val):
    """Retorna el código de error de Excel que aparece en el valor (texto), o None."""
    if not not_empty(val): return None
    v = str(val).strip().upper()
    for codigo in ERRORES_EXCEL:
        if codigo in v:
            return codigo
    return None

def clave_columna(nombre):
    """Nombre de columna normalizado: sin espacios extremos, en mayúsculas y sin tildes."""
    s = unicodedata.normalize("NFKD", str(nombre).strip().upper())
    return "".join(ch for ch in s if not unicodedata.combining(ch))

def celdas_con_error_excel(contenido, columnas):
    """
    pandas lee las celdas con error de fórmula (#REF!, #N/A, ...) como vacías,
    así que una columna opcional las dejaría pasar. Esta función abre el archivo
    con openpyxl y reporta esas celdas en las columnas indicadas, con la fila real de Excel.
    """
    errores = []
    objetivo = {clave_columna(c): c for c in columnas}
    try:
        wb = openpyxl.load_workbook(io.BytesIO(contenido), read_only=True, data_only=True)
    except Exception:
        return errores  # p. ej. un .xls antiguo, que openpyxl no abre
    try:
        filas = wb.worksheets[0].iter_rows()
        encabezado = next(filas, None)
        if encabezado is None:
            return errores
        indices = {
            i: objetivo[clave_columna(c.value)]
            for i, c in enumerate(encabezado)
            if c.value is not None and clave_columna(c.value) in objetivo
        }
        for fila in filas:
            for i, campo in indices.items():
                if i < len(fila) and fila[i].data_type == "e":
                    errores.append({
                        "Fila": fila[i].row, "Campo": campo, "Valor": str(fila[i].value),
                        "Error": f"La celda tiene el error de Excel {fila[i].value}; revisa la fórmula o el valor de origen",
                    })
    finally:
        wb.close()
    return errores

def validar_distribucion(col, add):
    """DISTRIBUCION (CLARO): sin errores de Excel escritos como texto; vacía según DISTRIBUCION_PERMITE_VACIOS."""
    if not DISTRIBUCION_PERMITE_VACIOS and not not_empty(col("DISTRIBUCION")):
        add("DISTRIBUCION", "Campo vacío")
        return
    codigo = error_excel_en(col("DISTRIBUCION"))
    if codigo:
        add("DISTRIBUCION", f"Contiene el error de Excel {codigo}; revisa la fórmula o el valor de origen")


# ══════════════════════════════════════════════════════════════════════════════
# VALIDACIÓN CLARO
# ══════════════════════════════════════════════════════════════════════════════

COLUMNAS_CLARO = [
    "CENTRAL", "HANDLE", "ORDEN DE TRABAJO", "BLOCKNAME",
    "BARRIO", "LATITUD", "LONGITUD", "CIUDAD", "DEPARTAMENTO",
    "CTO TERRENO", "CTO x HHPP", "OLT", "PUERTO OLT",
    "LATITUD2", "LONGITUD2", "CANTIDAD PUERTOS", "TIPO CTO", "DISTRIBUCION"
]
OPCIONALES_CLARO = []

REGLAS_CLARO = {
    "CENTRAL":           "Sin caracteres especiales",
    "HANDLE":            "Obligatorio, sin espacios ni caracteres especiales y sin repetirse",
    "ORDEN DE TRABAJO":  "Obligatorio, sin espacios",
    "BLOCKNAME":         "Obligatorio, solo N_Casa o N_Edificio",
    "BARRIO":            "Obligatorio",
    "LATITUD":           "Obligatoria, con punto decimal, mínimo 6 decimales y sin espacios",
    "LONGITUD":          "Obligatoria y negativa, con punto decimal, mínimo 6 decimales y sin espacios",
    "CIUDAD":            "Sin caracteres especiales",
    "DEPARTAMENTO":      "Sin caracteres especiales",
    "CTO TERRENO":       "Obligatorio",
    "CTO x HHPP":        "Obligatorio",
    "OLT":               "Obligatorio",
    "PUERTO OLT":        "Obligatorio, formato N/N/N, por ejemplo 1/3/3",
    "LATITUD2":          "Obligatoria, con punto decimal, mínimo 6 decimales y sin espacios",
    "LONGITUD2":         "Obligatoria y negativa, con punto decimal, mínimo 6 decimales y sin espacios",
    "CANTIDAD PUERTOS":  "Obligatorio",
    "TIPO CTO":          "Obligatorio, sin caracteres especiales",
    "DISTRIBUCION":      "Columna obligatoria; las celdas pueden ir vacías, pero sin errores de Excel como #REF! o #N/A",
}

def validar_fila_claro(row, idx):
    errores = []
    fila_num = idx + 2
    def col(n): return row.get(n, None)
    def add(campo, msg): errores.append({"Fila": fila_num, "Campo": campo, "Valor": str(col(campo))[:80], "Error": msg})

    r = no_special_chars(col("CENTRAL"))
    if r is False: add("CENTRAL", "Contiene caracteres especiales")

    r = no_special_chars_no_spaces(col("HANDLE"))
    if r is False:  add("HANDLE", "Contiene caracteres especiales o espacios")
    elif r is None: add("HANDLE", "Campo vacío")

    r = no_spaces_only(col("ORDEN DE TRABAJO"))
    if r is False:  add("ORDEN DE TRABAJO", "Contiene espacios")
    elif r is None: add("ORDEN DE TRABAJO", "Campo vacío")

    r = valid_blockname(col("BLOCKNAME"))
    if r is False:  add("BLOCKNAME", f"Valor '{col('BLOCKNAME')}' inválido. Debe ser 'N_Casa' o 'N_Edificio'")
    elif r is None: add("BLOCKNAME", "Campo vacío")

    if not not_empty(col("BARRIO")): add("BARRIO", "Campo vacío")

    ok, msg = valid_latlon(col("LATITUD"))
    if not ok: add("LATITUD", f"Error: {msg}")

    ok, msg = valid_latlon(col("LONGITUD"), must_be_negative=True)
    if not ok: add("LONGITUD", f"Error: {msg}")

    r = no_special_chars(col("CIUDAD"))
    if r is False: add("CIUDAD", "Contiene caracteres especiales")

    r = no_special_chars(col("DEPARTAMENTO"))
    if r is False: add("DEPARTAMENTO", "Contiene caracteres especiales")

    for campo_vacio in ["CTO TERRENO", "CTO x HHPP", "OLT", "CANTIDAD PUERTOS"]:
        if not not_empty(col(campo_vacio)): add(campo_vacio, "Campo vacío")

    ok, msg = valid_puerto_olt(col("PUERTO OLT"))
    if not ok:
        add("PUERTO OLT", f"Puerto OLT inválido: {msg}")

    ok, msg = valid_latlon(col("LATITUD2"))
    if not ok: add("LATITUD2", f"Error: {msg}")

    ok, msg = valid_latlon(col("LONGITUD2"), must_be_negative=True)
    if not ok: add("LONGITUD2", f"Error: {msg}")

    r = no_special_chars(col("TIPO CTO"))
    if r is None:   add("TIPO CTO", "Campo vacío")
    elif r is False: add("TIPO CTO", "Contiene caracteres especiales")

    validar_distribucion(col, add)

    return errores


# ══════════════════════════════════════════════════════════════════════════════
# VALIDACIÓN ETB
# ══════════════════════════════════════════════════════════════════════════════

COLUMNAS_ETB = [
    "NUMERO DEL", "NO_DEL_CLU", "CTO", "CHIP", "CUNU", "LOTE",
    "LOCALIDAD", "BARRIO", "DIRECCION", "DIRECCION ANTERIOR",
    "LLAVE", "ESTRATO", "DESTINO", "CLASE_PRED",
    "CODIGO_BAR", "CODIGO_USO", "CODIGO_MUN", "ID PLACA",
    "LATITUD", "LONGITUD", "ID_POLIGON", "TIPO_INCOR",
    "COSECHA-ETB", "ESTADO", "DIRE NUEVA", "ACTUALIZAR"
]
OPCIONALES_ETB = []

REGLAS_ETB = {
    "NUMERO DEL":         "Obligatorio",
    "NO_DEL_CLU":         "Obligatorio, sin caracteres especiales",
    "CTO":                "Obligatorio, formato XXXXX-CTN-NNN-NN-CON",
    "CHIP":               "Obligatorio, numérico",
    "CUNU":               "Opcional",
    "LOTE":               "Obligatorio, numérico",
    "LOCALIDAD":          "Obligatoria, solo letras y espacios",
    "BARRIO":             "Obligatorio",
    "DIRECCION":          "Obligatoria",
    "DIRECCION ANTERIOR": "Opcional",
    "LLAVE":              "Obligatoria",
    "ESTRATO":            "Obligatorio, número entre 1 y 6",
    "DESTINO":            "Obligatorio, solo letras",
    "CLASE_PRED":         "Obligatoria",
    "CODIGO_BAR":         "Opcional, numérico si tiene valor",
    "CODIGO_USO":         "Obligatorio, solo letras",
    "CODIGO_MUN":         "Obligatorio, código DANE de 5 dígitos",
    "ID PLACA":           "Obligatorio, numérico",
    "LATITUD":            "Obligatoria, con punto decimal, mínimo 6 decimales y dentro de Colombia",
    "LONGITUD":           "Obligatoria y negativa, con punto decimal, mínimo 6 decimales y dentro de Colombia",
    "ID_POLIGON":         "Obligatorio, formato XXXX-XXXN-NN-X-XXXNN",
    "TIPO_INCOR":         "Obligatorio",
    "COSECHA-ETB":        "Obligatoria, formato COSECHA-ETB-NN",
    "ESTADO":             "Opcional",
    "DIRE NUEVA":         "Opcional",
    "ACTUALIZAR":         "Opcional",
}

def valid_cto_etb(val):
    if not not_empty(val): return None
    v = str(val).strip().upper()
    return bool(re.match(r'^[A-Z0-9]+-[A-Z0-9]+-[A-Z0-9]+-[A-Z0-9]+-[A-Z0-9]+$', v))

def valid_id_poligon(val):
    if not not_empty(val): return None
    v = str(val).strip().upper()
    return bool(re.match(r'^[A-Z0-9]+-[A-Z0-9]+-[A-Z0-9]+-[A-Z0-9]+-[A-Z0-9]+$', v))

def valid_cosecha_etb(val):
    if not not_empty(val): return None
    return bool(re.match(r'^COSECHA-ETB-\d+$', str(val).strip(), re.IGNORECASE))

def valid_estrato(val):
    if not not_empty(val): return None, "vacío"
    try:
        n = int(float(str(val).strip()))
        if 1 <= n <= 6: return True, "ok"
        return False, f"valor {n} fuera de rango (debe ser 1-6)"
    except ValueError:
        return False, "no es un número entero"

def valid_codigo_mun(val):
    if not not_empty(val): return False, "vacío"
    try:
        v_int = str(int(float(str(val).strip())))
    except Exception:
        return False, "no es numérico"
    if not re.match(r'^\d{5}$', v_int):
        return False, f"debe tener 5 dígitos, tiene {len(v_int)}: '{v_int}'"
    return True, "ok"

def validar_fila_etb(row, idx):
    errores = []
    fila_num = idx + 2
    def col(n): return row.get(n, None)
    def add(campo, msg): errores.append({"Fila": fila_num, "Campo": campo, "Valor": str(col(campo))[:80], "Error": msg})

    # Campos obligatorios simples
    for c in ["NUMERO DEL", "BARRIO", "DIRECCION", "LLAVE", "CLASE_PRED", "TIPO_INCOR"]:
        if not not_empty(col(c)): add(c, "Campo vacío")

    # NO_DEL_CLU
    if not not_empty(col("NO_DEL_CLU")):
        add("NO_DEL_CLU", "Campo vacío")
    else:
        r = no_special_chars(col("NO_DEL_CLU"))
        if r is False: add("NO_DEL_CLU", "Contiene caracteres especiales")

    # CTO
    r = valid_cto_etb(col("CTO"))
    if r is None:    add("CTO", "Campo vacío")
    elif r is False: add("CTO", f"Formato inválido '{col('CTO')}'. Esperado: XXXXX-CTN-NNN-NN-CON (separado por guiones)")

    # CHIP
    if not not_empty(col("CHIP")):
        add("CHIP", "Campo vacío")
    elif es_numerico(col("CHIP")) is False:
        add("CHIP", f"Debe ser numérico, se encontró: '{col('CHIP')}'")

    # LOTE
    if not not_empty(col("LOTE")):
        add("LOTE", "Campo vacío")
    elif es_numerico(col("LOTE")) is False:
        add("LOTE", f"Debe ser numérico, se encontró: '{col('LOTE')}'")

    # LOCALIDAD
    if not not_empty(col("LOCALIDAD")):
        add("LOCALIDAD", "Campo vacío")
    elif solo_letras_espacios(col("LOCALIDAD")) is False:
        add("LOCALIDAD", "Solo debe contener letras y espacios")

    # ESTRATO
    ok, msg = valid_estrato(col("ESTRATO"))
    if ok is None:    add("ESTRATO", "Campo vacío")
    elif ok is False: add("ESTRATO", f"Estrato inválido: {msg}")

    # DESTINO
    if not not_empty(col("DESTINO")):
        add("DESTINO", "Campo vacío")
    elif solo_letras_espacios(col("DESTINO")) is False:
        add("DESTINO", "Solo debe contener letras")

    # CODIGO_BAR — opcional, pero si tiene valor debe ser numérico
    if not_empty(col("CODIGO_BAR")) and es_numerico(col("CODIGO_BAR")) is False:
        add("CODIGO_BAR", f"Debe ser numérico si está presente, se encontró: '{col('CODIGO_BAR')}'")

    # CODIGO_USO
    if not not_empty(col("CODIGO_USO")):
        add("CODIGO_USO", "Campo vacío")
    elif solo_letras_espacios(col("CODIGO_USO")) is False:
        add("CODIGO_USO", "Solo debe contener letras")

    # CODIGO_MUN
    ok, msg = valid_codigo_mun(col("CODIGO_MUN"))
    if not ok: add("CODIGO_MUN", f"Código DANE inválido: {msg}")

    # ID PLACA
    if not not_empty(col("ID PLACA")):
        add("ID PLACA", "Campo vacío")
    elif es_numerico(col("ID PLACA")) is False:
        add("ID PLACA", f"Debe ser numérico, se encontró: '{col('ID PLACA')}'")

    # LATITUD — Colombia ~4 a 13
    ok, msg = valid_latlon(col("LATITUD"), must_be_negative=False)
    if not ok:
        add("LATITUD", f"Error: {msg}")
    else:
        try:
            lat = float(str(col("LATITUD")).strip())
            if not (-5 <= lat <= 15):
                add("LATITUD", f"Valor {lat} fuera del rango geográfico de Colombia")
        except: pass

    # LONGITUD — Colombia ~ -67 a -82
    ok, msg = valid_latlon(col("LONGITUD"), must_be_negative=True)
    if not ok:
        add("LONGITUD", f"Error: {msg}")
    else:
        try:
            lon = float(str(col("LONGITUD")).strip())
            if not (-85 <= lon <= -65):
                add("LONGITUD", f"Valor {lon} fuera del rango geográfico de Colombia")
        except: pass

    # ID_POLIGON
    r = valid_id_poligon(col("ID_POLIGON"))
    if r is None:    add("ID_POLIGON", "Campo vacío")
    elif r is False: add("ID_POLIGON", f"Formato inválido '{col('ID_POLIGON')}'. Esperado: XXXX-XXXN-NN-X-XXXNN")

    # COSECHA-ETB
    r = valid_cosecha_etb(col("COSECHA-ETB"))
    if r is None:    add("COSECHA-ETB", "Campo vacío")
    elif r is False: add("COSECHA-ETB", f"Formato inválido '{col('COSECHA-ETB')}'. Esperado: COSECHA-ETB-NN")

    return errores


# ══════════════════════════════════════════════════════════════════════════════
# MOTOR GENÉRICO
# ══════════════════════════════════════════════════════════════════════════════

OPERADORES = {
    "CLARO": {
        "columnas": COLUMNAS_CLARO, "opcionales": OPCIONALES_CLARO, "reglas": REGLAS_CLARO,
        "color": "var(--claro)", "descripcion": "Estructura FTTH con CTO, OLT y puertos de red.",
        # Columnas donde además se buscan celdas con error de fórmula (#REF!, #N/A, ...)
        "sin_errores_excel": ["DISTRIBUCION"],
    },
    "ETB": {
        "columnas": COLUMNAS_ETB, "opcionales": OPCIONALES_ETB, "reglas": REGLAS_ETB,
        "color": "var(--etb)", "descripcion": "Estructura catastral con predios, polígonos y cosechas.",
        "sin_errores_excel": [],
    },
}

def validar_duplicados_handle(df):
    """
    Detecta valores duplicados en la columna HANDLE.
    Retorna lista de errores indicando en qué filas se repite el valor.
    """
    errores = []
    if "HANDLE" not in df.columns:
        return errores
    handle_series = df["HANDLE"].apply(lambda v: str(v).strip() if not_empty(v) else None)
    grupos = {}
    for idx, val in handle_series.items():
        if val is None:
            continue
        grupos.setdefault(val, []).append(idx)
    for val, indices in grupos.items():
        if len(indices) > 1:
            filas_duplicadas = sorted([i + 2 for i in indices])
            for idx in indices:
                errores.append({
                    "Fila":  idx + 2,
                    "Campo": "HANDLE",
                    "Valor": val,
                    "Error": f"DUPLICADO — mismo HANDLE encontrado en filas: {', '.join(map(str, filas_duplicadas))}"
                })
    return errores


def validar_dataframe(df, operador, errores_extra=()):
    todos = []
    fn = validar_fila_claro if operador == "CLARO" else validar_fila_etb
    for idx, row in df.iterrows():
        todos.extend(fn(row, idx))

    # Validación de duplicados en HANDLE (solo aplica a CLARO)
    if operador == "CLARO":
        todos.extend(validar_duplicados_handle(df))

    # Errores detectados fuera de pandas (celdas con #REF!, #N/A, ...)
    todos.extend(errores_extra)

    result = pd.DataFrame(todos)
    if result.empty:
        result = pd.DataFrame(columns=["Fila", "Campo", "Valor", "Error"])
    result = result.sort_values("Fila", kind="stable").reset_index(drop=True)
    return result

def normalizar_columnas(df, columnas_req, columnas_opc=()):
    """Renombra las columnas del archivo a los nombres oficiales (ignora mayúsculas, tildes y espacios extremos)."""
    df.columns = [str(c).strip() for c in df.columns]
    por_clave = {clave_columna(c): c for c in df.columns}
    rename = {}
    for req in [*columnas_req, *columnas_opc]:
        if clave_columna(req) in por_clave:
            rename[por_clave[clave_columna(req)]] = req
    df = df.rename(columns=rename)
    missing = [c for c in columnas_req if c not in df.columns]
    return df, missing


# ══════════════════════════════════════════════════════════════════════════════
# UI
# ══════════════════════════════════════════════════════════════════════════════

def fmt(n):
    """Número con separador de miles al estilo colombiano: 1.240"""
    return f"{n:,}".replace(",", ".")

def plural(n, singular, plural_):
    return f"{fmt(n)} {singular if n == 1 else plural_}"

def tabla_reglas(reglas):
    filas = []
    for campo, regla in reglas.items():
        clase = ' class="opcional"' if regla.startswith("Opcional") else ""
        filas.append(f"<tr{clase}><td>{html.escape(campo)}</td><td>{html.escape(regla)}</td></tr>")
    return (
        '<table class="tabla-reglas"><thead><tr><th>Columna</th><th>Qué se revisa</th></tr></thead>'
        f'<tbody>{"".join(filas)}</tbody></table>'
    )

def chips_columnas(requeridas, opcionales, faltantes=(), opc_ausentes=()):
    chips = []
    for c in requeridas:
        if c in faltantes:
            chips.append(f'<span class="chip falta">{html.escape(c)}: falta</span>')
        else:
            chips.append(f'<span class="chip">{html.escape(c)}</span>')
    for c in opcionales:
        estado = "opcional, no incluida" if c in opc_ausentes else "opcional"
        chips.append(f'<span class="chip opcional">{html.escape(c)} ({estado})</span>')
    return f'<div class="chips">{"".join(chips)}</div>'


# ── Barra superior: título | operador + reglas ───────────────────────────────
barra = st.container(horizontal=True, vertical_alignment="bottom", gap="medium", key="barra")
with barra.container(width="stretch", key="titulo"):
    st.markdown("""
    <div class="encabezado">
        <h1>Validador de archivos Excel</h1>
        <p>Revisa cada fila del archivo antes de enviarlo al operador.</p>
    </div>
    """, unsafe_allow_html=True)
with barra:
    operador = st.segmented_control(
        "Operador", options=list(OPERADORES), default="CLARO", required=True, key="operador"
    )
conf = OPERADORES[operador]
COLUMNAS_REQ = conf["columnas"]
COLUMNAS_OPC = conf["opcionales"]
REGLAS       = conf["reglas"]
with barra:
    with st.popover("Reglas", icon=":material/rule:"):
        st.markdown(f"**Reglas de {operador}**")
        st.markdown(tabla_reglas(REGLAS), unsafe_allow_html=True)

st.markdown(
    f'<p class="nota-operador"><span class="marca" style="background:{conf["color"]}"></span>'
    f'{operador}: {conf["descripcion"]}</p>',
    unsafe_allow_html=True
)

uploaded_file = st.file_uploader(
    f"Archivo Excel para {operador}",
    type=["xlsx", "xls"],
    key=f"upload_{operador}",
    label_visibility="collapsed",
    help="Se lee la primera hoja del archivo.",
)

if not uploaded_file:
    st.markdown(f"""
    <div class="vacio">
        <div>
            <h3>Sube el archivo de {operador}</h3>
            <p>Debe traer estas {len(COLUMNAS_REQ)} columnas en la primera fila. Las reglas de cada una están en el botón Reglas.</p>
        </div>
        {chips_columnas(COLUMNAS_REQ, COLUMNAS_OPC)}
    </div>
    """, unsafe_allow_html=True)
    st.stop()


# ── Lectura y validación ─────────────────────────────────────────────────────
contenido = uploaded_file.getvalue()
with st.spinner("Leyendo el archivo..."):
    try:
        df_raw = pd.read_excel(io.BytesIO(contenido), dtype=str)
    except Exception as e:
        st.error(f"No se pudo leer el archivo. Verifica que sea un Excel válido y que no esté protegido. Detalle: {e}")
        st.stop()

df, missing_cols = normalizar_columnas(df_raw.copy(), COLUMNAS_REQ, COLUMNAS_OPC)
opc_ausentes = [c for c in COLUMNAS_OPC if c not in df.columns]

with st.spinner("Validando datos..."):
    errores_celdas = celdas_con_error_excel(contenido, conf["sin_errores_excel"])
    df_errores = validar_dataframe(df, operador, errores_celdas)

total_filas    = len(df)
total_errores  = len(df_errores)
filas_con_err  = df_errores["Fila"].nunique() if total_errores > 0 else 0
filas_ok       = max(total_filas - filas_con_err, 0)
campos_con_err = df_errores["Campo"].nunique() if total_errores > 0 else 0
encontradas    = len(COLUMNAS_REQ) - len(missing_cols)

if missing_cols:
    st.warning(
        f"Faltan columnas obligatorias: **{', '.join(missing_cols)}**. "
        "Revisa el nombre del encabezado; mientras falten, cada fila aparecerá con error en esos campos."
    )

if total_filas == 0:
    st.info("El archivo no tiene filas de datos debajo del encabezado. Revisa que la información esté en la primera hoja.")
    st.stop()

# ── Veredicto: estado | cifras | acciones ────────────────────────────────────
pct_ok = filas_ok / total_filas * 100
nombre = html.escape(uploaded_file.name)
listo = total_errores == 0 and not missing_cols
if listo:
    titulo = f"Listo para enviar a {operador}"
    detalle = f"{nombre}: todas las filas cumplen las reglas."
elif total_errores == 0:
    titulo = "Faltan columnas obligatorias"
    detalle = f"{nombre}: agrega {', '.join(html.escape(c) for c in missing_cols)} antes de enviarlo."
else:
    titulo = "Corrige el archivo antes de enviarlo"
    detalle = f"{nombre}: descarga el reporte o revisa el detalle abajo."

with st.container(key="veredicto_ok" if listo else "veredicto_err"):
    franja = st.container(horizontal=True, vertical_alignment="center", gap="large")
    with franja.container(width="stretch"):
        st.markdown(f'<div class="estado"><h2>{titulo}</h2><p>{detalle}</p></div>', unsafe_allow_html=True)
    with franja.container(width="content"):
        rojo = ' class="rojo"' if total_errores else ""
        st.markdown(f"""
        <div class="cifras">
            <div><b>{fmt(total_filas)}</b><span>filas</span></div>
            <div{rojo}><b>{fmt(filas_con_err)}</b><span>con errores</span></div>
            <div{rojo}><b>{fmt(total_errores)}</b><span>errores</span></div>
            <div{rojo}><b>{fmt(campos_con_err)}</b><span>campos</span></div>
        </div>
        """, unsafe_allow_html=True)
    with franja.container(width="content", key="acciones"):
        if total_errores:
            buffer = io.BytesIO()
            ts = datetime.now().strftime('%Y%m%d_%H%M%S')
            with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                df_errores.to_excel(writer, index=False, sheet_name="Errores Detalle")
                resumen = df_errores.groupby("Campo").agg(
                    Total_Errores=("Error", "count"),
                    Filas_afectadas=("Fila", "nunique")
                ).reset_index().sort_values("Total_Errores", ascending=False)
                resumen.to_excel(writer, index=False, sheet_name="Resumen por Campo")
            st.download_button(
                label="Descargar reporte",
                data=buffer.getvalue(),
                file_name=f"reporte_{operador}_{ts}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                icon=":material/download:",
                width="stretch",
            )
        with st.popover(f"Columnas {encontradas}/{len(COLUMNAS_REQ)}", icon=":material/view_column:", width="stretch"):
            st.markdown(chips_columnas(COLUMNAS_REQ, COLUMNAS_OPC, missing_cols, opc_ausentes), unsafe_allow_html=True)

    st.markdown(f"""
    <div class="barra" role="img" aria-label="{pct_ok:.1f}% de filas sin errores">
        <div class="ok" style="width:{pct_ok:.3f}%"></div>
    </div>
    <div class="leyenda">
        <span class="l-ok">{plural(filas_ok, "fila sin errores", "filas sin errores")} ({pct_ok:.1f}%)</span>
        <span class="l-err">{plural(filas_con_err, "fila con errores", "filas con errores")}</span>
    </div>
    """, unsafe_allow_html=True)

if total_errores == 0:
    st.stop()

# ── Explorador: campos a la izquierda, errores a la derecha ──────────────────
st.markdown('<p class="titulo-seccion">Errores por corregir</p>', unsafe_allow_html=True)

conteo = df_errores["Campo"].value_counts()
TODOS = "__todos__"
opciones = [TODOS, *conteo.index]
if st.session_state.get("campo_sel") not in opciones:
    st.session_state.pop("campo_sel", None)

e_campos, e_tabla = st.columns([1.2, 4], gap="medium")
with e_campos:
    campo_sel = st.radio(
        "Campo", options=opciones, key="campo_sel", label_visibility="collapsed",
        format_func=lambda c: f"Todos los campos ({fmt(total_errores)})" if c == TODOS else f"{c} ({fmt(conteo[c])})",
    )

with e_tabla:
    vista = df_errores if campo_sel == TODOS else df_errores[df_errores["Campo"] == campo_sel]
    f_regla, f_buscar = st.columns([3, 2], vertical_alignment="center")
    with f_regla:
        if campo_sel == TODOS:
            st.markdown(f'<p class="regla">{plural(len(vista), "error", "errores")} en {plural(vista["Fila"].nunique(), "fila", "filas")}. Elige un campo a la izquierda para ver su regla.</p>', unsafe_allow_html=True)
        else:
            st.markdown(f'<p class="regla"><b>Regla:</b> {html.escape(REGLAS.get(campo_sel, "sin descripción"))}</p>', unsafe_allow_html=True)
    with f_buscar:
        buscar = st.text_input("Buscar", placeholder="Buscar fila o valor", label_visibility="collapsed",
                               icon=":material/search:", key="buscar").strip()
    if buscar:
        q = buscar.lower()
        vista = vista[
            (vista["Fila"].astype(str) == q)
            | vista["Valor"].astype(str).str.lower().str.contains(q, regex=False)
            | vista["Error"].astype(str).str.lower().str.contains(q, regex=False)
        ]
    columnas = ["Fila", "Valor", "Error"] if campo_sel != TODOS else ["Fila", "Campo", "Valor", "Error"]
    st.dataframe(
        vista[columnas].reset_index(drop=True),
        width="stretch", hide_index=True, height=460,
        column_config={
            "Fila":  st.column_config.NumberColumn("Fila", width="small", format="%d"),
            "Valor": st.column_config.TextColumn("Valor encontrado", width="medium"),
            "Error": st.column_config.TextColumn("Qué corregir", width="large"),
        },
    )
    if buscar and vista.empty:
        st.caption(f"Ningún error coincide con «{buscar}».")
