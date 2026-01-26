import os
import pdfplumber
import pandas as pd

# Ruta de la carpeta con los PDFs
ruta = r"C:\RICHARD\FDL\Usme\2026\CRP_vigencia\Enero\Grupo5_vig"

# Ruta del Excel con la equivalencia CDP → No.Interno CDP y Objeto
ruta_cdp_excel = os.path.join(ruta, "Reporte_CDP_Ene26_11.xlsx")

# Cargar el archivo de equivalencias
mapa_cdp = {}
if os.path.exists(ruta_cdp_excel):
    df_cdp = pd.read_excel(ruta_cdp_excel)
    print("🧪 Columnas detectadas en el archivo de equivalencias:")
    print(df_cdp.columns.tolist())  # Para verificar nombres exactos

    # Buscar columnas relevantes
    col_cdp = next((col for col in df_cdp.columns if "cdp" in col.lower()), None)
    col_interno = next((col for col in df_cdp.columns if "interno" in col.lower()), None)
    col_objeto = next((col for col in df_cdp.columns if "objeto" in col.lower()), None)

    if col_cdp and col_interno and col_objeto:
        for _, fila in df_cdp.iterrows():
            clave = str(fila[col_cdp]).strip()
            valor_interno = str(fila[col_interno]).strip()
            valor_objeto = str(fila[col_objeto]).strip()
            mapa_cdp[clave] = {
                "NoInterno": valor_interno,
                "Objeto": valor_objeto
            }
    else:
        print("❌ No se encontraron columnas 'CDP', 'No.Interno CDP' y 'Objeto' en el Excel.")
else:
    print(f"⚠️ No se encontró el archivo de equivalencias: {ruta_cdp_excel}")

def limpiar_numero(s: str) -> int:
    if not s:
        return 0
    s = str(s).strip()
    if s in ["-", ""]:
        return 0
    if not any(ch.isdigit() for ch in s):
        return 0
    s = s.replace(".", "").replace(",", "").replace("$", "")
    try:
        return int(s)
    except ValueError:
        return 0

datos = []

# Validar que la ruta existe
if not os.path.exists(ruta):
    print(f"❌ La ruta no existe: {ruta}")
else:
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

                                datos.append({
                                    "Archivo": archivo,
                                    "No DE CONTRATO": fila[0],
                                    "MODALIDAD": fila[1],
                                    "TIPOLOGIA DEL CONTRATO": fila[2],
                                    "NOMBRE DEL CONTRATISTA": fila[3],
                                    "CEDULA DEL CONTRATISTA": fila[4],
                                    "RUBRO": fila[5],
                                    "SIPSE": fila[6],
                                    "CDP": cdp_valor,
                                    "No.Interno CDP": datos_cdp["NoInterno"],
                                    "VALOR CDP": limpiar_numero(fila[8]),
                                    "VALOR DEL CONTRATO": limpiar_numero(fila[9]),
                                    "OBJETO": datos_cdp["Objeto"]
                                })

    if datos:
        df = pd.DataFrame(datos)
        salida = os.path.join(ruta, "Reporte_Contratos_Extraidos.xlsx")
        df.to_excel(salida, index=False)
        print(f"✅ Reporte generado en: {salida}")
        for d in datos:
            print(f"{d['No DE CONTRATO']} → CDP: {d['CDP']} → Interno: {d['No.Interno CDP']} → Objeto: {d['OBJETO']}")
    else:
        print("⚠️ No se encontraron contratos en los PDFs.")


