import os
import pandas as pd
from .file_utils import limpiar_ruta

def leer_excel_o_csv(ruta: str, dtype=str, hoja: str = None) -> pd.DataFrame | None:
    """
    Lee un archivo Excel o CSV y retorna un DataFrame.
    Maneja diferentes separadores y encodings para CSV.
    """
    ruta = limpiar_ruta(ruta)
    if not os.path.exists(ruta):
        print(f"El archivo {ruta} no existe.")
        return None

    try:
        extension = os.path.splitext(ruta)[1].lower()

        if extension == ".csv":
            print(f"📄 Leyendo CSV: {ruta}")
            # Intentar diferentes configuraciones
            for sep, enc in [(';', 'utf-8-sig'), (',', 'utf-8-sig'), (';', 'latin-1'), (',', 'latin-1')]:
                try:
                    df = pd.read_csv(ruta, sep=sep, encoding=enc, dtype=dtype, on_bad_lines="skip")
                    print(f"  ✓ Leído con sep='{sep}' y encoding='{enc}'")
                    return df
                except Exception:
                    continue
            print(f"❌ No se pudo leer el CSV con ninguna configuración")
            return None

        elif extension in [".xlsx", ".xls"]:
            if hoja:
                return pd.read_excel(ruta, sheet_name=hoja, dtype=dtype)
            return pd.read_excel(ruta, dtype=dtype)
        else:
            print(f"Extensión '{extension}' no soportada")
            return None

    except Exception as e:
        print(f"❌ Error al leer archivo {ruta}: {e}")
        return None
