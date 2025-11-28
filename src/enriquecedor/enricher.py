import os
import re
import pandas as pd
from ..core import limpiar_ruta, leer_configuracion_enriquecimiento, leer_excel_o_csv
from ..utils import parse_lista_columnas
from .lookup import buscar_coincidencia_parcial
from .formula_engine import evaluar_formula

_cache_hojas_lookup = {}

def _preparar_cache_lookup(excel_base: pd.ExcelFile):
    """
    Prepara caché de hojas para funciones LOOKUP.
    """
    _cache_hojas_lookup.clear()
    for hoja in excel_base.sheet_names:
        try:
            _cache_hojas_lookup[hoja.lower()] = excel_base.parse(hoja)
        except Exception:
            pass

def enriquecer_hojas(ruta_excel: str, ruta_configuracion: str):
    """
    Enriquece hojas Excel con datos externos y fórmulas calculadas.
    """
    data_conf = leer_configuracion_enriquecimiento(ruta_configuracion)
    bloques = data_conf["bloques"]
    parametros = data_conf["parametros"]

    ruta_excel = limpiar_ruta(ruta_excel)
    if not os.path.exists(ruta_excel):
        print(f"Excel base no existe: {ruta_excel}")
        return

    try:
        excel_base = pd.ExcelFile(ruta_excel)
        hojas_existentes = set(excel_base.sheet_names)
        _preparar_cache_lookup(excel_base)

        # Agrupar bloques por hoja
        bloques_por_hoja = {}
        for b in bloques:
            hoja = (b.get("hoja") or "").strip().strip('"')
            if not hoja:
                continue
            bloques_por_hoja.setdefault(hoja, []).append(b)

        with pd.ExcelWriter(ruta_excel, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            for hoja, lista_ops in bloques_por_hoja.items():
                if hoja not in hojas_existentes:
                    print(f"Hoja '{hoja}' no existe. Omitida.")
                    continue

                df_base = excel_base.parse(hoja)
                df_base.columns = [str(c).strip().lower() for c in df_base.columns]

                print(f"Procesando hoja '{hoja}' con {len(lista_ops)} bloque(s)...")

                for cfg in lista_ops:
                    # Enriquecimiento desde archivo externo
                    _enriquecer_desde_archivo(df_base, cfg, _cache_hojas_lookup)
                    
                    # Aplicar fórmulas calculadas
                    _aplicar_calculos(df_base, cfg, parametros, _cache_hojas_lookup)

                df_base.to_excel(writer, sheet_name=hoja, index=False)
                print(f"Hoja '{hoja}' escrita/enriquecida.")

        print("Enriquecimiento finalizado.")
    except Exception as e:
        print(f"Error enriqueciendo: {e}")

def _enriquecer_desde_archivo(df_base: pd.DataFrame, cfg: dict, cache_hojas: dict):
    """
    Enriquece DataFrame base con datos de archivo externo.
    """
    columna_base = (cfg.get("columna base") or "").strip().strip('"')
    ruta_ext = limpiar_ruta(cfg.get("ruta") or "")
    columna_cruzar = (cfg.get("columna cruzar") or "").strip().strip('"')
    columnas_extraer_raw = cfg.get("columna extraer") or ""

    columna_base_norm = columna_base.lower() if columna_base else ""

    if not (ruta_ext and os.path.exists(ruta_ext) and columna_cruzar):
        return

    # Leer archivo externo
    hoja_lookup = (cfg.get("hoja lookup") or cfg.get("hoja externa") or "").strip().strip('"').strip("'")
    df_ext = leer_excel_o_csv(ruta_ext, hoja=hoja_lookup if hoja_lookup else None)
    
    if df_ext is None:
        return

    df_ext.columns = [str(c).strip().lower() for c in df_ext.columns]
    
    # Agregar al caché con alias
    alias_cfg = (cfg.get("alias lookup") or "").strip().strip('"').strip("'")
    alias_base = alias_cfg.lower() if alias_cfg else os.path.splitext(os.path.basename(ruta_ext))[0].lower()
    variantes = {alias_base}
    sin_paren = re.sub(r'\(.*?\)', '', alias_base)
    compacto = re.sub(r'[^a-z0-9]+', ' ', sin_paren).strip()
    variantes.update({sin_paren.strip(), compacto})
    
    for a in variantes:
        cache_hojas[a] = df_ext

    # Validar columnas
    columna_cruzar_norm = columna_cruzar.lower()
    columnas_extraer = parse_lista_columnas(columnas_extraer_raw)
    faltantes = [c for c in columnas_extraer if c not in df_ext.columns]
    
    if columna_cruzar_norm not in df_ext.columns or faltantes:
        print(f"  ⚠ Columnas inválidas en archivo externo")
        return
    
    if columna_base_norm and columna_base_norm not in df_base.columns:
        print(f"  ⚠ Columna base '{columna_base_norm}' no existe")
        return

    # Agregar columnas si no existen
    for col in columnas_extraer:
        if col not in df_base.columns:
            df_base[col] = "N/A"

    # Buscar coincidencias y enriquecer
    if columna_base_norm:
        base_series = df_base[columna_base_norm]
        coincidencias = 0
        
        for idx_base, valor_base in base_series.items():
            idx_match = buscar_coincidencia_parcial(valor_base, df_ext[columna_cruzar_norm])
            if idx_match is not None:
                coincidencias += 1
                for col_extraer in columnas_extraer:
                    df_base.at[idx_base, col_extraer] = df_ext.at[idx_match, col_extraer]
        
        print(f"  → {coincidencias} coincidencias encontradas (archivo: {os.path.basename(ruta_ext)})")

def _aplicar_calculos(df_base: pd.DataFrame, cfg: dict, parametros: dict, cache_hojas: dict):
    """
    Aplica fórmulas calculadas a DataFrame.
    """
    calculos = cfg.get("calculos", [])
    
    if not calculos:
        return

    print(f"  Aplicando {len(calculos)} fórmula(s)...")
    
    for calc in calculos:
        nombre = calc["nombre"]
        formula = calc["formula"]
        resultados = []
        
        for _, fila in df_base.iterrows():
            val = evaluar_formula(formula, fila, parametros, cache_hojas)
            if val is None or (isinstance(val, float) and pd.isna(val)):
                val = "N/A"
            resultados.append(val)
        
        df_base[nombre] = resultados
        print(f"    ✓ '{nombre}' aplicada")
