import os
import pdfplumber
import pandas as pd

ruta = r"C:\RICHARD\FDL\Usme\2026\CRP_vigencia\Enero"
datos = []

for archivo in os.listdir(ruta):
    if archivo.lower().endswith(".pdf"):
        path_pdf = os.path.join(ruta, archivo)
        with pdfplumber.open(path_pdf) as pdf:
            for page in pdf.pages:
                tablas = page.extract_tables()
                for tabla in tablas:
                    for fila in tabla:
                        # Validar que la fila tenga al menos 10 columnas (contrato completo)
                        if fila and len(fila) >= 10:
                            datos.append({
                                "Archivo": archivo,
                                "No DE CONTRATO": fila[0],
                                "MODALIDAD": fila[1],
                                "TIPOLOGIA DEL CONTRATO": fila[2],
                                "NOMBRE DEL CONTRATISTA": fila[3],
                                "CEDULA DEL CONTRATISTA": fila[4],
                                "RUBRO": fila[5],
                                "SIPSE": fila[6],
                                "CDP": fila[7],
                                "VALOR CDP": fila[8],
                                "VALOR DEL CONTRATO": fila[9]
                            })

# Exportar a Excel
df = pd.DataFrame(datos)
salida = os.path.join(ruta, "Reporte_Contratos_Extraidos.xlsx")
df.to_excel(salida, index=False)

print(f"✅ Reporte generado en: {salida}")






