import pdfplumber
import re

# Ruta del PDF
pdf_file = "CPS 482-2025-1.pdf"

# Diccionario para resultados
datos = {
    "Contrato No": None,
    "Contratista": None,
    "NIT o CC": None,
    "Pago No": None,
    "Valor Bruto": None,
    "% ReteFuentes": [],
    "Valor ReteFuentes": [],
    "Neto a Pagar": None
}

with pdfplumber.open(pdf_file) as pdf:
    texto = ""
    for page in pdf.pages:
        texto += page.extract_text() + "\n"

    # Contrato No
    contrato = re.search(r"CONTRATO No\.?\s*(CPS\s*\d+-\d+)", texto)
    if contrato:
        datos["Contrato No"] = contrato.group(1)

    # Contratista
    contratista = re.search(r"CONTRATISTA:\s*(.+)", texto)
    if contratista:
        datos["Contratista"] = contratista.group(1).strip()

    # NIT o CC
    nit = re.search(r"NIT\. o C\.C\.\s*(\d[\d\.\-]+)", texto)
    if nit:
        datos["NIT o CC"] = nit.group(1)

    # Pago No
    pago = re.search(r"PAGO No\.\s*(\d+)", texto)
    if pago:
        datos["Pago No"] = pago.group(1)

    # Valor Bruto
    valor_bruto = re.search(r"VALOR BRUTO.*?\$ ?([\d\.,]+)", texto)
    if valor_bruto:
        datos["Valor Bruto"] = valor_bruto.group(1)

    # % y valor retefuentes (busca todas las filas con % y valor)
    retefuentes = re.findall(r"(Retefuente.*?|Reteica).*?(\d+[\.,]?\d*%?).*?\$ ?([\d\.,-]+)", texto)
    for rf in retefuentes:
        datos["% ReteFuentes"].append(rf[1])
        datos["Valor ReteFuentes"].append(rf[2])

    # Neto a Pagar
    neto = re.search(r"NETO A PAGAR.*?\$ ?([\d\.,]+)", texto)
    if neto:
        datos["Neto a Pagar"] = neto.group(1)

# Mostrar resultados
for k, v in datos.items():
    print(f"{k}: {v}")
