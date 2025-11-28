import os
import pandas as pd
from ..core import limpiar_ruta, leer_instrucciones, leer_excel_o_csv
from ..core.file_utils import asegurar_carpeta
from .filtros import aplicar_criterio

def procesar_archivo(ruta: str, columna_criterio: str, criterio: str, columna_id: str) -> pd.DataFrame | None:
    """
    Procesa un archivo y aplica filtros según criterio.
    """
    df = leer_excel_o_csv(ruta)
    if df is None:
        return None

    # Normalizar columnas
    df.columns = [str(c).strip().lower() for c in df.columns]
    columna_criterio = columna_criterio.strip().lower()
    columna_id = columna_id.strip().lower()

    print(f"  [DEBUG] Columnas: {df.columns.tolist()[:10]}")

    # Validar columnas
    if columna_criterio not in df.columns:
        print(f"❌ Columna criterio '{columna_criterio}' no existe")
        return None
    if columna_id not in df.columns:
        print(f"❌ Columna ID '{columna_id}' no existe")
        return None

    # Aplicar filtros
    df_filtrado = aplicar_criterio(df, columna_criterio, criterio)
    
    # Reordenar columnas (columna ID primero)
    if not df_filtrado.empty:
        cols = [columna_id] + [c for c in df_filtrado.columns if c != columna_id]
        df_filtrado = df_filtrado[cols]
    
    return df_filtrado

def generar_excel_salida(ruta_instrucciones: str, ruta_salida: str):
    """
    Genera Excel de salida procesando todas las instrucciones.
    """
    ruta_instrucciones = limpiar_ruta(ruta_instrucciones)
    ruta_salida = limpiar_ruta(ruta_salida)
    asegurar_carpeta(ruta_salida)

    instrucciones = leer_instrucciones(ruta_instrucciones)
    hojas_creadas = False

    try:
        with pd.ExcelWriter(ruta_salida, engine='openpyxl') as writer:
            for idx, instruccion in enumerate(instrucciones):
                nombre_hoja = instruccion.get("archivo") or f"Hoja_{idx + 1}"
                ruta = instruccion.get("ruta", "").strip()
                columna_criterio = instruccion.get("columna criterio", "")
                criterio = instruccion.get("criterio", "")
                columna_id = instruccion.get("columna id", "")

                print(f"Procesando hoja: {nombre_hoja}")

                if ruta:
                    df_filtrado = procesar_archivo(ruta, columna_criterio, criterio, columna_id)
                    if df_filtrado is not None:
                        df_filtrado.to_excel(writer, sheet_name=nombre_hoja, index=False)
                        hojas_creadas = True
                        if df_filtrado.empty:
                            print(f"⚠ Hoja '{nombre_hoja}' creada vacía (sin coincidencias)")
                else:
                    # Hoja vacía
                    pd.DataFrame().to_excel(writer, sheet_name=nombre_hoja, index=False)
                    print(f"Hoja '{nombre_hoja}' creada vacía (sin ruta)")
                    hojas_creadas = True

            if not hojas_creadas:
                df_vacio = pd.DataFrame({"Mensaje": ["No se generaron datos válidos"]})
                df_vacio.to_excel(writer, sheet_name="Hoja_Predeterminada", index=False)

        print(f"✅ Archivo Excel generado en: {ruta_salida}")
    except PermissionError:
        print(f"❌ Error: No se pudo escribir en '{ruta_salida}'. Cierre el archivo.")
    except Exception as e:
        print(f"❌ Error inesperado: {e}")

def crear_hoja_resumen(ruta_excel: str, ruta_instrucciones: str):
    """
    Crea hoja Resumen con IDs de todas las hojas.
    """
    configuraciones = leer_instrucciones(ruta_instrucciones)
    ruta_excel = limpiar_ruta(ruta_excel)

    if not os.path.exists(ruta_excel):
        print(f"Excel '{ruta_excel}' no existe. No se puede crear resumen.")
        return

    try:
        with pd.ExcelWriter(ruta_excel, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
            excel = pd.ExcelFile(ruta_excel)
            resumen = {}

            for configuracion in configuraciones:
                hoja = configuracion.get("archivo")
                columna_id = configuracion.get("columna id")

                if not hoja or not columna_id or hoja not in excel.sheet_names:
                    continue

                df_hoja = excel.parse(hoja)
                df_hoja.columns = [str(c).strip().lower() for c in df_hoja.columns]
                columna_id = columna_id.strip().lower()

                if columna_id not in df_hoja.columns or df_hoja.empty:
                    continue

                ids = df_hoja[columna_id].dropna().unique()
                for id_val in ids:
                    id_str = str(id_val).strip()
                    if id_str not in resumen:
                        resumen[id_str] = {}
                    resumen[id_str][hoja] = "X"

            # Crear DataFrame resumen
            if resumen:
                df_resumen = pd.DataFrame.from_dict(resumen, orient="index").fillna("")
                df_resumen.index.name = "ID"
                df_resumen.reset_index(inplace=True)
            else:
                df_resumen = pd.DataFrame(columns=["ID"])

            df_resumen.to_excel(writer, sheet_name="Resumen", index=False)
            print("Hoja 'Resumen' creada exitosamente.")

    except Exception as e:
        print(f"Error al crear hoja resumen: {e}")
