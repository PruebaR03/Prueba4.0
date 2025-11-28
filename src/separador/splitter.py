import os
import pandas as pd
from ..core import limpiar_ruta, leer_configuracion_separacion, leer_instrucciones

def celda_contiene_identificador(valor, identificadores: list) -> bool:
    """
    Verifica si celda coincide con identificador (maneja N/A, vacíos, etc).
    """
    identificadores_lower = [str(i).lower().strip() for i in identificadores]
    
    if "n/a" in identificadores_lower:
        if pd.isna(valor) or valor is None:
            return True
        valor_str = str(valor).strip().upper()
        if valor_str in ["N/A", "NA", ""]:
            return True
    
    if pd.isna(valor) or valor is None:
        return False
    
    v = str(valor).lower().strip()
    if not v:
        return "" in identificadores_lower
    
    for ident in identificadores_lower:
        if ident in v:
            return True
    
    return False

def _map_id_por_hoja_desde_instrucciones(ruta_instrucciones: str) -> dict:
    """
    Mapea hoja -> columna id desde instrucciones.
    """
    try:
        instrucciones = leer_instrucciones(ruta_instrucciones)
    except Exception as e:
        print(f"⚠ No se pudo leer instrucciones: {e}")
        return {}
    
    mapa = {}
    for inst in instrucciones:
        hoja = inst.get("archivo")
        col_id = inst.get("columna id")
        if hoja and col_id:
            mapa[str(hoja)] = str(col_id).strip().lower()
    return mapa

def flujo_separacion(ruta_excel: str, ruta_configuracion: str, carpeta_salida: str = "output/separados", ruta_instrucciones: str | None = None):
    """
    Separa Excel en múltiples archivos según identificadores.
    """
    ruta_excel = limpiar_ruta(ruta_excel)
    config = leer_configuracion_separacion(ruta_configuracion)

    if not config['hojas_calculo'] or not config['hojas']:
        print("Configuración inválida o incompleta.")
        return

    if not os.path.exists(ruta_excel):
        print(f"Excel base no existe: {ruta_excel}")
        return

    id_por_hoja = _map_id_por_hoja_desde_instrucciones(ruta_instrucciones) if ruta_instrucciones else {}
    carpeta_salida = limpiar_ruta(carpeta_salida)
    os.makedirs(carpeta_salida, exist_ok=True)

    try:
        print("\n" + "="*60)
        print("INICIANDO SEPARACIÓN DE EXCEL")
        print("="*60 + "\n")

        excel_base = pd.ExcelFile(ruta_excel)
        hojas_existentes = set(excel_base.sheet_names)

        for calc in config['hojas_calculo']:
            nombre_salida = calc['name']
            identificadores = calc.get('identificadores', [])
            
            if not identificadores:
                print(f"⚠ '{nombre_salida}' sin identificadores. Omitido.")
                continue

            print(f"\n--- Procesando: {nombre_salida} ---")
            print(f"Identificadores: {', '.join(identificadores)}")

            ruta_out = os.path.join(carpeta_salida, f"{nombre_salida}.xlsx")

            with pd.ExcelWriter(ruta_out, engine='openpyxl') as writer:
                resumen = {}

                for hoja_cfg in config['hojas']:
                    hoja = hoja_cfg.get('hoja')
                    col_sep = hoja_cfg.get('columna_id', '').strip()
                    
                    if not hoja or not col_sep or hoja not in hojas_existentes:
                        continue

                    df = excel_base.parse(hoja)
                    df.columns = [str(c).strip().lower() for c in df.columns]
                    col_sep_norm = col_sep.lower()
                    
                    if col_sep_norm not in df.columns:
                        print(f"  ⚠ Columna '{col_sep}' no existe en '{hoja}'")
                        continue

                    # Filtrar por identificadores
                    mask = df[col_sep_norm].apply(lambda v: celda_contiene_identificador(v, identificadores))
                    df_filtrado = df[mask]

                    if df_filtrado.empty:
                        pd.DataFrame(columns=df.columns).to_excel(writer, sheet_name=hoja, index=False)
                        print(f"  ⚠ Hoja '{hoja}': sin coincidencias (vacía)")
                    else:
                        df_filtrado.to_excel(writer, sheet_name=hoja, index=False)
                        print(f"  ✓ Hoja '{hoja}': {len(df_filtrado)} filas")

                        # Columna para resumen
                        col_resumen = id_por_hoja.get(hoja, col_sep_norm)
                        if col_resumen not in df_filtrado.columns:
                            col_resumen = col_sep_norm

                        # Acumular IDs
                        if "n/a" in [str(i).lower() for i in identificadores]:
                            ids_validos = df_filtrado[col_resumen].unique()
                        else:
                            ids_validos = df_filtrado[col_resumen].dropna().unique()
                        
                        for id_val in ids_validos:
                            key = "N/A" if pd.isna(id_val) else str(id_val).strip()
                            if key not in resumen:
                                resumen[key] = {}
                            resumen[key][hoja] = "X"

                # Hoja Resumen
                if resumen:
                    df_resumen = pd.DataFrame.from_dict(resumen, orient="index").fillna("")
                    df_resumen.index.name = "ID"
                    df_resumen.reset_index(inplace=True)
                else:
                    df_resumen = pd.DataFrame(columns=["ID"])
                
                df_resumen.to_excel(writer, sheet_name="Resumen", index=False)
                print(f"  → Resumen agregado ({len(df_resumen)} filas)")

            print(f"  → Archivo guardado: {ruta_out}")

        print("\n" + "="*60)
        print("✓ SEPARACIÓN COMPLETADA")
        print("="*60 + "\n")

    except Exception as e:
        print(f"\n✗ Error durante la separación: {e}")
        import traceback
        traceback.print_exc()
