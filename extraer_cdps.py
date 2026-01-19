import re
import fitz  # PyMuPDF
import pandas as pd
import os

def extraer_cdps(pdf_path):
    doc = fitz.open(pdf_path)
    resultados = []

    for page in doc:
        texto = page.get_text("text")

        # Extraer Objeto
        objeto_match = re.search(r"OBJETO:\s*(.+?)(?=VALOR:)", texto, re.S)
        objeto = objeto_match.group(1).strip() if objeto_match else ""

        # Extraer Valor
        valor_match = re.search(r"VALOR:\s*\$?\s*([\d\.,]+)", texto)
        valor = valor_match.group(1).strip() if valor_match else ""

        # Extraer Descripción de Actividades
        descripcion_match = re.search(
            r"DESCRIPCIÓN DE ACTIVIDADES DE LA SOLICITUD\s*(.+?)(OBJETO:|VALOR:|$)",
            texto,
            re.S
        )
        descripcion = descripcion_match.group(1).strip() if descripcion_match else ""

        # Extraer Código de Proyecto (ej: 2685 → O230117459920242268501000)
        numero_match = re.search(r"\b(\d{4})\b", descripcion)
        if numero_match:
            numero = numero_match.group(1)
            codigo = f"O230117459920242{numero}01000"
        else:
            codigo = "No encontrado"

        # Extraer Número de Solicitud
        solicitud_match = re.search(r"PARA LA SOLICITUD No\.?\s*(\d+)", texto)
        solicitud = solicitud_match.group(1) if solicitud_match else "No encontrado"

        # Extraer Imagen (guardar primera imagen encontrada)
        imagenes = []
        for img in page.get_images(full=True):
            xref = img[0]
            pix = fitz.Pixmap(doc, xref)
            nombre_img = f"{os.path.basename(pdf_path)}_img_{xref}.png"
            if not pix.alpha:
                pix.save(nombre_img)
            else:
                fitz.Pixmap(pix, 0).save(nombre_img)
            imagenes.append(nombre_img)

        resultados.append({
            "Archivo": os.path.basename(pdf_path),
            "Objeto": objeto,
            "Valor": valor,
            "Descripción Actividades": descripcion,
            "Código Proyecto": codigo,
            "Solicitud No.": solicitud,
            "Imagen": ", ".join(imagenes) if imagenes else "No encontrada"
        })

    return resultados

# 📂 Ruta donde están todos los PDFs
carpeta = r"C:\RICHARD\FDL\Usme\2026\CDP_Vigencia\Enero\tercer_grupo"

archivos = [os.path.join(carpeta, f) for f in os.listdir(carpeta) if f.endswith(".pdf")]

todos = []
for archivo in archivos:
    todos.extend(extraer_cdps(archivo))

# Exportar a Excel en la misma carpeta
excel_path = os.path.join(carpeta, "cdps2.xlsx")

# Eliminar archivo anterior si existe
if os.path.exists(excel_path):
    os.remove(excel_path)

df = pd.DataFrame(todos)
df.to_excel(excel_path, index=False)
print(f"✅ Archivo Excel generado: {excel_path}")



