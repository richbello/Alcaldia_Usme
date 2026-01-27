import streamlit as st
import os
import pdfplumber
import pandas as pd
from datetime import datetime
import re

# --- Funciones auxiliares ---
def limpiar_numero(s: str) -> int:
    if not s or s in ["-", ""]:
        return 0
    s = str(s).replace(".", "").replace(",", "").replace("$", "").strip()
    return int(re.sub(r"\D", "", s)) if any(ch.isdigit() for ch in s) else 0

def normalizar_texto(texto: str) -> str:
    if not texto:
        return ""
    texto = str(texto).strip()
    texto = re.sub(r"\s+", " ", texto)
    return texto

def tipo_compromiso(objeto: str) -> int:
    objeto = objeto.lower()
    if "servicios profesionales" in objeto:
        return 145
    elif "servicios de apoyo" in objeto:
        return 148
    return 0

# --- Interfaz gráfica ---
st.set_page_config(page_title="Generador Plantilla Financiera", layout="wide")

# Fondo e identidad visual
page_bg = """
<style>
[data-testid="stAppViewContainer"] {
    background-image: url("https://upload.wikimedia.org/wikipedia/commons/5/5c/Bogota_City.jpg");
    background-size: cover;
}
[data-testid="stHeader"] {
    background-color: rgba(255, 0, 0, 0.8);
}
[data-testid="stSidebar"] {
    background-color: rgba(255, 255, 0, 0.9);
}
</style>
"""
st.markdown(page_bg, unsafe_allow_html=True)

st.title("📊 Generador de Plantilla Cargue Masivo CRP")
st.title("Alcaldia Local de Usme.")

# Subida de archivos
pdfs = st.file_uploader("Sube los PDFs de contratos", type="pdf", accept_multiple_files=True)
excel_equiv = st.file_uploader("Sube el Excel de equivalencias CDP", type="xlsx")

if st.button("Generar Plantilla"):
    if pdfs and excel_equiv:
        # Leer equivalencias
        df_cdp = pd.read_excel(excel_equiv)
        col_cdp = next((col for col in df_cdp.columns if "cdp" in col.lower()), None)
        col_interno = next((col for col in df_cdp.columns if "interno" in col.lower()), None)
        col_objeto = next((col for col in df_cdp.columns if "objeto" in col.lower()), None)

        mapa_cdp = {}
        for _, fila in df_cdp.iterrows():
            clave = str(fila[col_cdp]).strip()
            mapa_cdp[clave] = {
                "NoInterno": str(fila[col_interno]).strip(),
                "Objeto": str(fila[col_objeto]).strip()
            }

        # Datos fijos
        fecha_actual = datetime.today().strftime("%d.%m.%Y")
        fecha_final = "31.12.2026"
        fijos = {
            "Posición": "1",
            "Sociedad": "1001",
            "Clase Documento": "RP",
            "Moneda": "COP",
            "Fecha Documento": fecha_actual,
            "Fecha Contabilización": fecha_actual,
            "Fecha Inicial": fecha_actual,
            "Fecha Final": fecha_final,
            "Tipo de Pago": "02",
            "Modo Selección": "10",
            "Tipo Documento Beneficiario": "CC",
            "ID Solicitante": "1000131265",
            "ID Responsable": "1000835316"
        }

        datos = []
        for archivo in pdfs:
            with pdfplumber.open(archivo) as pdf:
                for page in pdf.pages:
                    tablas = page.extract_tables()
                    for tabla in tablas:
                        for fila in tabla:
                            if fila and len(fila) >= 10:
                                cdp_valor = str(fila[7]).strip()
                                datos_cdp = mapa_cdp.get(cdp_valor, {"NoInterno": "NO ENCONTRADO", "Objeto": "NO ENCONTRADO"})
                                valor_contrato = limpiar_numero(fila[9])
                                objeto_norm = normalizar_texto(datos_cdp["Objeto"])
                                contrato_norm = normalizar_texto(fila[0])
                                tipo = tipo_compromiso(objeto_norm)

                                datos.append({
                                    "Importe": valor_contrato,
                                    "CDP": datos_cdp["NoInterno"],
                                    "Posición del CDP": "1",
                                    "Objeto": objeto_norm,
                                    "Tipo de compromiso": tipo,
                                    "No. Compromiso": contrato_norm,
                                    "Identificación Beneficiario": normalizar_texto(fila[4]),
                                    **fijos
                                })

        if datos:
            df = pd.DataFrame(datos)
            df["CRP"] = range(1, len(df) + 1)
            df["Num. Ext. Entidad"] = range(1, len(df) + 1)

            columnas_finales = [
                "CRP", "Posición", "Fecha Documento", "Fecha Contabilización", "Sociedad", "Clase Documento",
                "Moneda", "Importe", "CDP", "Posición del CDP", "Objeto", "Tipo de compromiso",
                "No. Compromiso", "Fecha Inicial", "Fecha Final", "Tipo de Pago", "Modo Selección",
                "Tipo Documento Beneficiario", "Identificación Beneficiario", "ID Solicitante", "ID Responsable",
                "Num. Ext. Entidad"
            ]

            df_final = df[columnas_finales]
            st.success("✅ Plantilla generada correctamente")
            st.dataframe(df_final)

            # Descargar Excel
            salida = "Plantilla_Financiera_Generada.xlsx"
            df_final.to_excel(salida, index=False)
            with open(salida, "rb") as f:
                st.download_button("📥 Descargar Excel", f, file_name=salida)
        else:
            st.warning("⚠️ No se encontraron registros válidos en los PDFs.")
    else:
        st.error("❌ Debes subir los PDFs y el Excel de equivalencias.")
