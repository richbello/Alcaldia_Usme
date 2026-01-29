import streamlit as st
import os
import pdfplumber
import pandas as pd
import re
from datetime import datetime
import logging
from time import sleep

# -----------------------
# Configuración general
# -----------------------
st.set_page_config(page_title="Alcaldia Local de Usme", layout="wide")
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SALIDAS_DIR = os.path.join(BASE_DIR, "salidas")
LOG_DIR = os.path.join(BASE_DIR, "logs")
os.makedirs(SALIDAS_DIR, exist_ok=True)
os.makedirs(LOG_DIR, exist_ok=True)
LOG_PATH = os.path.join(LOG_DIR, "accesos.log")
logging.basicConfig(filename=LOG_PATH, level=logging.INFO, format="%(asctime)s - %(message)s")

# -----------------------
# Inicializar session_state
# -----------------------
if "usuario" not in st.session_state:
    st.session_state["usuario"] = None
if "df_final" not in st.session_state:
    st.session_state["df_final"] = None
if "uploaded_flag" not in st.session_state:
    st.session_state["uploaded_flag"] = False
if "processing" not in st.session_state:
    st.session_state["processing"] = False
if "auto_start" not in st.session_state:
    st.session_state["auto_start"] = False

# -----------------------
# Utilidades
# -----------------------
def limpiar_numero(s):
    if not s or s in ["-", ""]:
        return 0
    s = str(s).replace(".", "").replace(",", "").replace("$", "").strip()
    return int(re.sub(r"\D", "", s)) if any(ch.isdigit() for ch in s) else 0

def normalizar_texto(t):
    return re.sub(r"\s+", " ", str(t).strip()) if t else ""

def tipo_compromiso(obj):
    obj = obj.lower()
    if "servicios profesionales" in obj:
        return 145
    if "servicios de apoyo" in obj:
        return 148
    return 0

# -----------------------
# Lectura de Excel cacheada
# -----------------------
@st.cache_data(ttl=300)
def leer_excel_bytes(file_bytes):
    return pd.read_excel(file_bytes, engine="openpyxl")

# -----------------------
# Estilos con colores bandera de Bogotá (mejorados)
# -----------------------
# Amarillo: #FFD100  Rojo: #CE1126
st.markdown(
    """
    <style>
    /* Contenedor general para separar visualmente */
    .app-container {
        padding: 8px;
    }

    /* Panel izquierdo: franja superior amarilla y franja inferior roja (bandera de Bogotá) */
    .left-panel {
      background: linear-gradient(180deg, #FFD100 0% 50%, #CE1126 50% 100%) !important;
      padding: 20px;
      border-radius: 10px;
      color: #0f1724;
      min-height: 520px;
      display: flex;
      flex-direction: column;
      justify-content: flex-start;
      box-shadow: 0 6px 12px rgba(0,0,0,0.08);
    }

    /* Forzar que el fondo de inputs y select muestre el degradado (transparente o semi-transparente) */
    .left-panel input[type="text"],
    .left-panel input[type="password"],
    .left-panel textarea,
    .left-panel .stFileUploader,
    .left-panel .stTextInput>div>div>input,
    .left-panel .stTextInput>div>div>textarea,
    .left-panel .stNumberInput>div>input,
    .left-panel .stSelectbox>div>div,
    .left-panel .stMultiSelect>div>div {
        background-color: rgba(255,255,255,0.92) !important;
        color: #0f1724 !important;
        border-radius: 6px !important;
        border: 1px solid rgba(0,0,0,0.08) !important;
    }

    /* Ajustes para uploader y contenedores generados por Streamlit (clases variables incluidas) */
    .left-panel .stFileUploader, 
    .left-panel .css-1d391kg, 
    .left-panel .css-1y4p8pa {
        background-color: rgba(255,255,255,0.92) !important;
        color: #0f1724 !important;
        border-radius: 6px !important;
    }

    /* Botones legibles sobre la bandera */
    .left-panel .stButton>button {
        background-color: rgba(15,23,36,0.06) !important;
        color: #0f1724 !important;
        border: 1px solid rgba(0,0,0,0.12) !important;
        border-radius: 6px !important;
    }

    /* Etiquetas y texto: color oscuro para la franja amarilla (mejor contraste) */
    .left-panel label,
    .left-panel .stMarkdown,
    .left-panel .stText,
    .left-panel p {
        color: #0f1724 !important;
    }

    /* Asegurar que el panel izquierdo ocupa la altura mínima en la columna */
    [data-testid="column"] > .left-panel { min-height: 520px; }

    /* Panel principal estilo oscuro (sin cambios funcionales) */
    .main-panel {
        background: #0f1724;
        color: #ffffff;
        padding: 20px;
        border-radius: 10px;
        box-shadow: 0 6px 18px rgba(0,0,0,0.18);
        min-height: 320px;
    }

    /* Forzar que el header de Streamlit no tape el estilo si lo deseas */
    header { display: none; }
    </style>
    """,
    unsafe_allow_html=True,
)

# -----------------------
# Layout: panel izquierdo (login) y panel principal
# -----------------------
st.markdown('<div class="app-container">', unsafe_allow_html=True)
col_left, col_main = st.columns([1, 3])

with col_left:
    st.markdown('<div class="left-panel">', unsafe_allow_html=True)
    st.markdown("### 🔒 Inicio de sesión")
    # Si quieres mostrar el escudo de Bogotá, descomenta y coloca la ruta o URL:
    # st.image("ruta_al_escudo.png", width=120)
    if st.session_state["usuario"] is None:
        usuario_input = st.text_input("Usuario", key="u_input")
        clave_input = st.text_input("Clave", type="password", key="p_input")
        if st.button("Ingresar"):
            # Credenciales demo; reemplazar por autenticación real si aplica
            ROLES = {"admin": "admin123", "auditor": "audit456", "usuario": "user789"}
            if usuario_input in ROLES and ROLES[usuario_input] == clave_input:
                st.session_state["usuario"] = usuario_input
                logging.info(f"Login exitoso: {usuario_input}")
                st.experimental_rerun()
            else:
                st.error("Credenciales inválidas")
    else:
        st.markdown(f"**Sesión activa**  \n{st.session_state['usuario']}")
        if st.button("Cerrar sesión"):
            st.session_state.clear()
            st.experimental_rerun()
        if st.button("Ver Reporte de Seguridad"):
            if os.path.exists(LOG_PATH):
                with open(LOG_PATH, "r", encoding="utf-8") as f:
                    st.text_area("Reporte de Seguridad", f.read(), height=300)
            else:
                st.info("No hay registros de seguridad.")
    st.markdown('</div>', unsafe_allow_html=True)

with col_main:
    st.markdown('<div class="main-panel">', unsafe_allow_html=True)
    st.markdown("## 🛡️ Interfaz con Ciberseguridad")
    st.markdown("#### Bienvenido al dashboard seguro")
    st.markdown("---")
    st.markdown("### 📊 Generador de Plantilla Cargue Masivo CRP")
    st.markdown("**Alcaldía Local de Usme**")
    st.markdown("")

    # Uploaders con on_change que marcan estado
    def on_upload():
        st.session_state["uploaded_flag"] = True
        st.session_state["auto_start"] = True

    st.markdown("**Sube los PDFs de contratos**  \nDrag and drop o Browse. Límite 200MB por archivo.")
    pdfs = st.file_uploader("PDFs contratos", type=["pdf"], accept_multiple_files=True, key="pdfs", on_change=on_upload)

    st.markdown("**Sube el Excel de equivalencias CDP**  \nDrag and drop o Browse. Límite 200MB.")
    excel_equiv = st.file_uploader("Excel equivalencias CDP", type=["xlsx"], key="excel", on_change=on_upload)

    st.markdown("---")

    # Mensaje cuando hay archivos nuevos
    if st.session_state["uploaded_flag"] and not st.session_state["processing"]:
        st.info("Se detectaron archivos nuevos. Pulsa Generar Plantilla o espera 2 segundos para inicio automático.")
        sleep(2)

    # Generación automática o por botón
    start_now = st.button("Generar Plantilla") or st.session_state.get("auto_start", False)
    if start_now:
        st.session_state["auto_start"] = False
        if not pdfs or not excel_equiv:
            st.error("❌ Debes subir los PDFs y el Excel de equivalencias.")
        else:
            st.session_state["processing"] = True
            with st.spinner("⏳ Procesando archivos..."):
                try:
                    df_cdp = leer_excel_bytes(excel_equiv)
                except Exception as e:
                    st.error(f"Error leyendo Excel: {e}")
                    st.session_state["processing"] = False
                    raise

                col_cdp = next((col for col in df_cdp.columns if "cdp" in col.lower()), None)
                col_interno = next((col for col in df_cdp.columns if "interno" in col.lower()), None)
                col_objeto = next((col for col in df_cdp.columns if "objeto" in col.lower()), None)
                if not col_cdp or not col_interno or not col_objeto:
                    st.error("❌ El Excel no tiene las columnas esperadas (CDP, Interno, Objeto).")
                    st.session_state["processing"] = False
                else:
                    mapa_cdp = {}
                    for _, fila in df_cdp.iterrows():
                        clave = str(fila[col_cdp]).strip()
                        mapa_cdp[clave] = {"NoInterno": str(fila[col_interno]).strip(), "Objeto": str(fila[col_objeto]).strip()}

                    fecha_actual = datetime.today().strftime("%d.%m.%Y")
                    fijos = {"Posición": "1", "Sociedad": "1001", "Clase Documento": "RP", "Moneda": "COP",
                             "Fecha Documento": fecha_actual, "Fecha Contabilización": fecha_actual,
                             "Fecha Inicial": fecha_actual, "Fecha Final": "31.12.2026",
                             "Tipo de Pago": "02", "Modo Selección": "10", "Tipo Documento Beneficiario": "CC",
                             "ID Solicitante": "1000131265", "ID Responsable": "1000835316"}

                    datos = []
                    total_pdfs = len(pdfs)
                    progress = st.progress(0)
                    for i, archivo in enumerate(pdfs, start=1):
                        try:
                            with pdfplumber.open(archivo) as pdf:
                                for page in pdf.pages:
                                    for tabla in page.extract_tables() or []:
                                        for fila in tabla:
                                            if fila and len(fila) >= 10:
                                                cdp_valor = str(fila[7]).strip()
                                                datos_cdp = mapa_cdp.get(cdp_valor, {"NoInterno": "NO ENCONTRADO", "Objeto": "NO ENCONTRADO"})
                                                datos.append({
                                                    "Importe": limpiar_numero(fila[9]),
                                                    "CDP": datos_cdp["NoInterno"],
                                                    "Posición del CDP": "1",
                                                    "Objeto": normalizar_texto(datos_cdp["Objeto"]),
                                                    "Tipo de compromiso": tipo_compromiso(datos_cdp["Objeto"]),
                                                    "No. Compromiso": normalizar_texto(fila[0]),
                                                    "Identificación Beneficiario": normalizar_texto(fila[4]),
                                                    **fijos
                                                })
                            logging.info(f"Procesado PDF: {archivo.name}")
                        except Exception as e:
                            st.warning(f"⚠️ Error procesando {archivo.name}: {e}")
                            logging.error(f"Error en PDF {archivo.name}: {e}")
                        progress.progress(int(i / total_pdfs * 100))

                    if datos:
                        df = pd.DataFrame(datos)
                        df["CRP"] = range(1, len(df) + 1)
                        df["Num. Ext. Entidad"] = range(1, len(df) + 1)
                        columnas_finales = ["CRP", "Posición", "Fecha Documento", "Fecha Contabilización", "Sociedad", "Clase Documento",
                                            "Moneda", "Importe", "CDP", "Posición del CDP", "Objeto", "Tipo de compromiso",
                                            "No. Compromiso", "Fecha Inicial", "Fecha Final", "Tipo de Pago", "Modo Selección",
                                            "Tipo Documento Beneficiario", "Identificación Beneficiario", "ID Solicitante", "ID Responsable",
                                            "Num. Ext. Entidad"]
                        df_final = df[columnas_finales]
                        st.session_state["df_final"] = df_final
                        salida = os.path.join(SALIDAS_DIR, f"Plantilla_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
                        df_final.to_excel(salida, index=False)
                        st.success("✅ Plantilla generada correctamente")
                        with open(salida, "rb") as f:
                            st.download_button("📥 Descargar Excel", f, file_name=os.path.basename(salida))
                        logging.info("Plantilla generada y guardada.")
                    else:
                        st.warning("⚠️ No se encontraron registros válidos en los PDFs.")
                    st.session_state["processing"] = False
                    st.session_state["uploaded_flag"] = False

    # Mostrar resultados persistentes
    if st.session_state["df_final"] is not None:
        st.markdown("### Resultado actual")
        st.dataframe(st.session_state["df_final"], use_container_width=True)

    st.markdown('</div>', unsafe_allow_html=True)

st.markdown('</div>', unsafe_allow_html=True)






