import streamlit as st
import pandas as pd
import re
import io
from datetime import datetime

st.set_page_config(
    page_title="Validador de Archivos Excel",
    page_icon="🔍",
    layout="wide"
)

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=IBM+Plex+Sans:wght@300;400;600&display=swap');
    html, body, [class*="css"] { font-family: 'IBM Plex Sans', sans-serif; }
    .stApp { background-color: #0f0f0f; color: #e0e0e0; }
    h1, h2, h3 { font-family: 'IBM Plex Mono', monospace !important; color: #00e5a0 !important; }
    .op-tag-claro {
        display: inline-block; background: #ff000022; border: 1px solid #ff4444;
        color: #ff4444; font-family: 'IBM Plex Mono', monospace; font-size: 0.8rem;
        padding: 3px 12px; border-radius: 20px; letter-spacing: 1px; margin-right: 8px;
    }
    .op-tag-etb {
        display: inline-block; background: #0066ff22; border: 1px solid #4488ff;
        color: #4488ff; font-family: 'IBM Plex Mono', monospace; font-size: 0.8rem;
        padding: 3px 12px; border-radius: 20px; letter-spacing: 1px; margin-right: 8px;
    }
    .metric-box {
        background: #1a1a1a; border: 1px solid #2a2a2a; border-radius: 8px;
        padding: 20px; text-align: center;
    }
    .metric-box .number { font-family: 'IBM Plex Mono', monospace; font-size: 2.2rem; font-weight: 600; }
    .metric-box .label { font-size: 0.75rem; text-transform: uppercase; letter-spacing: 2px; color: #888; margin-top: 4px; }
    .success-banner {
        background: #0a1a0f; border: 1px solid #00e5a0; border-radius: 8px;
        padding: 24px; text-align: center; font-family: 'IBM Plex Mono', monospace;
    }
    .field-header {
        background: #111; border: 1px solid #222; border-radius: 6px;
        padding: 9px 14px; margin: 3px 0;
        font-family: 'IBM Plex Mono', monospace; font-size: 0.82rem;
        display: flex; justify-content: space-between; align-items: center;
    }
    .badge-ok { color: #00e5a0; font-weight: 600; }
    .badge-err { color: #ff4444; font-weight: 600; }
    .stButton>button {
        background: #00e5a0; color: #000; font-family: 'IBM Plex Mono', monospace;
        font-weight: 600; border: none; border-radius: 4px; padding: 10px 24px; width: 100%;
    }
    .stButton>button:hover { background: #00c080; color: #000; }
    div[data-testid="stExpander"] { background: #141414; border: 1px solid #222; border-radius: 6px; }
    .upload-hint { color: #555; font-size: 0.8rem; font-family: 'IBM Plex Mono', monospace; text-align: center; margin-bottom: 16px; }
    .rule-hint { color: #555; font-size: 0.75rem; }
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


# ══════════════════════════════════════════════════════════════════════════════
# VALIDACIÓN CLARO
# ══════════════════════════════════════════════════════════════════════════════

COLUMNAS_CLARO = [
    "CENTRAL", "HANDLE", "ORDEN DE TRABAJO", "BLOCKNAME",
    "BARRIO", "LATITUD", "LONGITUD", "CIUDAD", "DEPARTAMENTO",
    "CTO TERRENO", "CTO x HHPP", "OLT", "PUERTO OLT",
    "LATITUD2", "LONGITUD2", "CANTIDAD PUERTOS", "TIPO CTO"
]

REGLAS_CLARO = {
    "CENTRAL":           "Sin caracteres especiales",
    "HANDLE":            "Sin caracteres especiales ni espacios",
    "ORDEN DE TRABAJO":  "Sin espacios",
    "BLOCKNAME":         "Solo 'N_Casa' o 'N_Edificio'",
    "BARRIO":            "No vacío",
    "LATITUD":           "Punto decimal · 6+ decimales · sin espacios",
    "LONGITUD":          "Punto decimal · 6+ decimales · sin espacios · valor negativo",
    "CIUDAD":            "Sin caracteres especiales",
    "DEPARTAMENTO":      "Sin caracteres especiales",
    "CTO TERRENO":       "No vacío",
    "CTO x HHPP":        "No vacío",
    "OLT":               "No vacío",
    "PUERTO OLT":        "Formato N/N/N obligatorio · ej: 1/3/3",
    "LATITUD2":          "Punto decimal · 6+ decimales · sin espacios",
    "LONGITUD2":         "Punto decimal · 6+ decimales · sin espacios · valor negativo",
    "CANTIDAD PUERTOS":  "No vacío",
    "TIPO CTO":          "No vacío · sin caracteres especiales",
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

REGLAS_ETB = {
    "NUMERO DEL":         "No vacío",
    "NO_DEL_CLU":         "No vacío · sin caracteres especiales",
    "CTO":                "No vacío · formato: XXXXX-CTN-NNN-NN-CON",
    "CHIP":               "No vacío · numérico",
    "CUNU":               "Opcional",
    "LOTE":               "No vacío · numérico",
    "LOCALIDAD":          "No vacío · solo letras y espacios",
    "BARRIO":             "No vacío",
    "DIRECCION":          "No vacío",
    "DIRECCION ANTERIOR": "Opcional",
    "LLAVE":              "No vacío",
    "ESTRATO":            "Numérico entre 1 y 6",
    "DESTINO":            "No vacío · solo letras",
    "CLASE_PRED":         "No vacío",
    "CODIGO_BAR":         "Opcional · numérico si presente",
    "CODIGO_USO":         "No vacío · solo letras",
    "CODIGO_MUN":         "No vacío · código DANE 5 dígitos",
    "ID PLACA":           "No vacío · numérico",
    "LATITUD":            "Punto decimal · 6+ decimales · rango Colombia",
    "LONGITUD":           "Punto decimal · 6+ decimales · negativo · rango Colombia",
    "ID_POLIGON":         "No vacío · formato: XXXX-XXXN-NN-X-XXXNN",
    "TIPO_INCOR":         "No vacío",
    "COSECHA-ETB":        "No vacío · formato: COSECHA-ETB-NN",
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

def validar_dataframe(df, operador):
    todos = []
    fn = validar_fila_claro if operador == "CLARO" else validar_fila_etb
    for idx, row in df.iterrows():
        todos.extend(fn(row, idx))
    result = pd.DataFrame(todos)
    if result.empty:
        result = pd.DataFrame(columns=["Fila", "Campo", "Valor", "Error"])
    return result

def normalizar_columnas(df, columnas_req):
    df.columns = [str(c).strip() for c in df.columns]
    col_map_upper = {c.upper(): c for c in df.columns}
    rename = {}
    for req in columnas_req:
        if req.upper() in col_map_upper:
            rename[col_map_upper[req.upper()]] = req
    df = df.rename(columns=rename)
    missing = [c for c in columnas_req if c not in df.columns]
    return df, missing


# ══════════════════════════════════════════════════════════════════════════════
# UI
# ══════════════════════════════════════════════════════════════════════════════

st.markdown("# 🔍 Validador de Archivo Excel")
st.markdown("##### Detecta errores campo por campo antes de enviar")
st.markdown("---")

# ── Selector operador ────────────────────────────────────────────────────────
st.markdown("### Selecciona el operador")
operador = st.radio("Operador", options=["CLARO", "ETB"], horizontal=True, label_visibility="collapsed")

if operador == "CLARO":
    COLUMNAS_REQ = COLUMNAS_CLARO
    REGLAS       = REGLAS_CLARO
    st.markdown('<span class="op-tag-claro">🔴 CLARO</span> — Estructura FTTH con CTOs, OLTs y polígonos de red', unsafe_allow_html=True)
else:
    COLUMNAS_REQ = COLUMNAS_ETB
    REGLAS       = REGLAS_ETB
    st.markdown('<span class="op-tag-etb">🔵 ETB</span> — Estructura catastral con cosechas, polígonos y predios', unsafe_allow_html=True)

st.markdown("---")

# ── Uploader ─────────────────────────────────────────────────────────────────
uploaded_file = st.file_uploader(
    f"Arrastra o selecciona tu archivo Excel para **{operador}**",
    type=["xlsx", "xls"],
    label_visibility="visible",
    key=f"upload_{operador}"
)
st.markdown(f'<p class="upload-hint">Se validarán {len(COLUMNAS_REQ)} columnas según las reglas de {operador}</p>', unsafe_allow_html=True)

if uploaded_file:
    with st.spinner("Leyendo archivo..."):
        try:
            df_raw = pd.read_excel(uploaded_file, dtype=str)
        except Exception as e:
            st.error(f"❌ No se pudo leer el archivo: {e}")
            st.stop()

    df, missing_cols = normalizar_columnas(df_raw.copy(), COLUMNAS_REQ)

    # ── Estado de columnas
    st.markdown("### 📋 Columnas detectadas")
    n_cols_ui = 4 if operador == "CLARO" else 5
    cols_grid = st.columns(n_cols_ui)
    for i, req in enumerate(COLUMNAS_REQ):
        with cols_grid[i % n_cols_ui]:
            if req in missing_cols:
                st.markdown(f'<div class="field-header"><span>{req}</span><span class="badge-err">✗ FALTA</span></div>', unsafe_allow_html=True)
            else:
                st.markdown(f'<div class="field-header"><span>{req}</span><span class="badge-ok">✓</span></div>', unsafe_allow_html=True)

    if missing_cols:
        st.warning(f"⚠️ Columnas faltantes: **{', '.join(missing_cols)}** — Esas validaciones serán omitidas.")

    st.markdown("---")

    # ── Validar
    with st.spinner("Validando datos..."):
        df_errores = validar_dataframe(df, operador)

    total_filas   = len(df)
    total_errores = len(df_errores)
    filas_con_err = df_errores["Fila"].nunique() if total_errores > 0 else 0
    filas_ok      = total_filas - filas_con_err
    pct_ok        = round(filas_ok / total_filas * 100, 1) if total_filas > 0 else 0

    op_color = "#ff4444" if operador == "CLARO" else "#4488ff"

    st.markdown("### 📊 Resumen de validación")
    m1, m2, m3, m4, m5 = st.columns(5)
    with m1: st.markdown(f'<div class="metric-box"><div class="number" style="color:#888">{total_filas}</div><div class="label">Filas totales</div></div>', unsafe_allow_html=True)
    with m2: st.markdown(f'<div class="metric-box"><div class="number" style="color:#00e5a0">{filas_ok}</div><div class="label">Filas OK</div></div>', unsafe_allow_html=True)
    with m3: st.markdown(f'<div class="metric-box"><div class="number" style="color:#ff4444">{filas_con_err}</div><div class="label">Filas con error</div></div>', unsafe_allow_html=True)
    with m4: st.markdown(f'<div class="metric-box"><div class="number" style="color:#ffaa00">{total_errores}</div><div class="label">Errores totales</div></div>', unsafe_allow_html=True)
    with m5: st.markdown(f'<div class="metric-box"><div class="number" style="color:{op_color}">{pct_ok}%</div><div class="label">Aprobación</div></div>', unsafe_allow_html=True)

    st.markdown("---")

    if total_errores == 0:
        st.markdown(f"""
        <div class="success-banner">
            <div style="font-size:3rem">✅</div>
            <div style="font-size:1.4rem; color:#00e5a0; margin-top:12px">¡Archivo válido para {operador}!</div>
            <div style="color:#666; margin-top:8px">Todos los campos cumplen las reglas. El archivo puede ser enviado.</div>
        </div>
        """, unsafe_allow_html=True)

    else:
        st.markdown("### ❌ Detalle de errores encontrados")
        tab1, tab2, tab3 = st.tabs(["📋 Por campo", "📄 Por fila", "🗂️ Vista completa"])

        with tab1:
            campos_conteo = df_errores["Campo"].value_counts()
            for campo, count in campos_conteo.items():
                regla_desc = REGLAS.get(campo, "—")
                with st.expander(f"🔴 **{campo}** — {count} error{'es' if count > 1 else ''} · *{regla_desc}*"):
                    sub = df_errores[df_errores["Campo"] == campo][["Fila", "Valor", "Error"]]
                    st.dataframe(
                        sub.reset_index(drop=True), use_container_width=True,
                        column_config={
                            "Fila":  st.column_config.NumberColumn("Fila #", width="small"),
                            "Valor": st.column_config.TextColumn("Valor encontrado", width="medium"),
                            "Error": st.column_config.TextColumn("Descripción del error", width="large"),
                        }
                    )

        with tab2:
            filas_unicas = sorted(df_errores["Fila"].unique())
            st.markdown(f"**{len(filas_unicas)} fila(s) con errores:** " +
                        ", ".join(map(str, filas_unicas[:60])) +
                        (" ..." if len(filas_unicas) > 60 else ""))
            fila_sel = st.selectbox("🔍 Ver errores de la fila:", options=filas_unicas)
            sub_fila = df_errores[df_errores["Fila"] == fila_sel][["Campo", "Valor", "Error"]]
            st.dataframe(sub_fila.reset_index(drop=True), use_container_width=True)

        with tab3:
            st.dataframe(
                df_errores[["Fila", "Campo", "Valor", "Error"]].reset_index(drop=True),
                use_container_width=True, height=440
            )

        # ── Descarga reporte
        st.markdown("---")
        st.markdown("### 📥 Descargar reporte de errores")
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
            label=f"⬇️ Descargar reporte {operador}",
            data=buffer.getvalue(),
            file_name=f"reporte_{operador}_{ts}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

else:
    # Pantalla de bienvenida
    op_color = "#ff4444" if operador == "CLARO" else "#4488ff"
    op_bg    = "#1a0000"  if operador == "CLARO" else "#00061a"
    st.markdown(f"""
    <div style="background:{op_bg}; border:1px solid {op_color}33; border-radius:12px;
         padding:32px; text-align:center; margin:20px 0;">
        <div style="font-size:2.5rem">📂</div>
        <div style="font-family:'IBM Plex Mono',monospace; color:{op_color}; margin-top:12px; font-size:1rem">
            Esperando archivo {operador}...
        </div>
        <div style="color:#555; margin-top:8px; font-size:0.82rem">
            Sube tu Excel con la estructura {operador} para comenzar la validación
        </div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown(f"#### Reglas activas — {operador} ({len(REGLAS)} campos)")
    reg_cols = st.columns(2)
    for i, (campo, regla) in enumerate(REGLAS.items()):
        with reg_cols[i % 2]:
            st.markdown(
                f'<div class="field-header">'
                f'<span style="color:#ccc"><b>{campo}</b></span>'
                f'<span class="rule-hint">{regla}</span>'
                f'</div>',
                unsafe_allow_html=True
            )