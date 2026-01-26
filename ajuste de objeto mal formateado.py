import pandas as pd
import os

# Ruta completa al archivo original
ruta = r"C:\RICHARD\FDL\Usme\2026\CRP_vigencia\Enero\Grupo5_vig\ajustar_objeto_justificar.xlsx"

# Verificar que el archivo existe
if not os.path.exists(ruta):
    print("❌ No se encontró el archivo en la ruta especificada.")
else:
    # Leer el archivo
    df = pd.read_excel(ruta)

    # Función para limpiar el texto del Objeto
    def limpiar_objeto(texto):
        if pd.isna(texto):
            return ""
        limpio = str(texto).replace("\n", " ").replace("\r", " ")
        while "  " in limpio:  # eliminar espacios duplicados
            limpio = limpio.replace("  ", " ")
        return limpio.strip()

    # Crear columna ajustada
    df["OBJETO AJUSTADO"] = df["Objeto"].apply(limpiar_objeto)

    # Guardar nuevo archivo
    salida = r"C:\RICHARD\FDL\Usme\2026\CRP_vigencia\Enero\Grupo5_vig\ajustar_objeto_normalizado.xlsx"
    df.to_excel(salida, index=False)

    print(f"✅ Archivo generado correctamente en: {salida}")
