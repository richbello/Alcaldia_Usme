import os
import pdfplumber
import pandas as pd
from datetime import datetime
import re

# Ruta de la carpeta con los PDFs
ruta = r"C:\RICHARD\FDL\Usme\2026\CRP_vigencia\Enero\Grupo5_vig"

# Ruta del Excel con la equivalencia CDP → No.Interno CDP y Objeto
ruta_cdp_excel = os.path.join(ruta, "Reporte_CDP_Ene26_11.xlsx")

# Cargar el archivo de equivalencias
mapa_cdp = {}
if os.path.exists(ruta_cdp_excel):
    df_cdp = pd.read_excel(ruta_cdp_excel)
    col_cdp = next((col for col in df_cdp.columns if "cdp" in col.lower()), None)
    col_interno = next((col for col in df_cdp.columns if "interno" in col.lower()), None)
    col_objeto = next((col for col in df_cdp.columns if "objeto" in col.lower()), None)

    if col_cdp and col_interno and col_objeto:
        for _, fila in df_cdp.iterrows():
            clave = str(fila[col_cdp]).strip()
            mapa_cdp[clave] = {
                "NoInterno": str(fila[col_interno]).strip(),
                "Objeto": str(fila[col_objeto]).strip()
            }
    else:
        print("❌ Columnas necesarias no encontradas en el Excel de equivalencias.")
else:
    print("❌ Archivo de equivalencias no encontrado.")

def limpiar_numero(s: str) -> int:
    if not s or s in ["-", ""]:
        return 0
    s = str(s).replace(".", "").replace(",", "").replace("$", "").strip()
    return int(re.sub(r"\D", "", s)) if any(ch.isdigit() for ch in s) else 0

def normalizar_texto(texto: str) -> str:
    if not texto:
        return ""
    texto = str(texto).strip()
    texto = re.sub(r"\s+", " ", texto)  # espacios múltiples → uno solo
    return texto

def tipo_compromiso(objeto: str) -> int:
    objeto = objeto.lower()
    if "servicios profesionales" in objeto:
        return 145
    elif "servicios de apoyo" in objeto:
        return 148
    return 0

# Fecha actual y final
fecha_actual = datetime.today().strftime("%d.%m.%Y")
fecha_final = "31.12.2026"

# Datos fijos
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

# Extraer datos desde PDFs
datos = []
if os.path.exists(ruta):
    for archivo in os.listdir(ruta):
        if archivo.lower().endswith(".pdf"):
            path_pdf = os.path.join(ruta, archivo)
            with pdfplumber.open(path_pdf) as pdf:
                for page in pdf.pages:
                    tablas = page.extract_tables()
                    for tabla in tablas:
                        for fila in tabla:
                            if fila and len(fila) >= 10:
                                cdp_valor = str(fila[7]).strip()
                                datos_cdp = mapa_cdp.get(cdp_valor, {"NoInterno": "NO ENCONTRADO", "Objeto": "NO ENCONTRADO"})
                                valor_contrato = limpiar_numero(fila[9])
                                objeto_norm = normalizar_texto(datos_cdp["Objeto"])
                                contrato_norm = normalizar_texto(fila[0])  # No DE CONTRATO
                                tipo = tipo_compromiso(objeto_norm)

                                datos.append({
                                    "Importe": valor_contrato,
                                    "CDP": datos_cdp["NoInterno"],  # conversión al No.Interno CDP
                                    "Posición del CDP": "1",
                                    "Objeto": objeto_norm,
                                    "Tipo de compromiso": tipo,
                                    "No. Compromiso": contrato_norm,
                                    "Identificación Beneficiario": normalizar_texto(fila[4]),  # CEDULA DEL CONTRATISTA
                                    **fijos
                                })
else:
    print("❌ Ruta de PDFs no encontrada.")

# Generar Excel final
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
    salida = os.path.join(ruta, "Plantilla_Financiera_Generada.xlsx")
    df_final.to_excel(salida, index=False)
    print(f"✅ Plantilla generada en: {salida}")
else:
    print("⚠️ No se encontraron registros válidos en los PDFs.")



