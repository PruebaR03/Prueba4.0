import os
import re
from typing import Dict, List, Any
from .file_utils import limpiar_ruta

def leer_instrucciones(ruta_instrucciones: str) -> List[Dict[str, str]]:
    """
    Lee el archivo de instrucciones y retorna lista de configuraciones.
    """
    ruta_instrucciones = limpiar_ruta(ruta_instrucciones)
    instrucciones = []
    instruccion = {}
    
    with open(ruta_instrucciones, 'r', encoding='utf-8') as archivo:
        for linea in archivo.read().splitlines():
            if linea.strip():
                if linea.lower().startswith("archivo:"):
                    if instruccion:
                        instrucciones.append(instruccion)
                    instruccion = {}
                if ":" in linea:
                    clave, valor = linea.split(":", 1)
                    instruccion[clave.strip().lower()] = valor.strip().strip('"')
        if instruccion:
            instrucciones.append(instruccion)
    
    return instrucciones

def leer_configuracion_enriquecimiento(ruta_configuracion: str) -> Dict[str, Any]:
    """
    Lee configuración de enriquecimiento con bloques y parámetros.
    """
    ruta_configuracion = limpiar_ruta(ruta_configuracion)
    configuraciones = []
    parametros: Dict[str, Any] = {}
    
    if not os.path.exists(ruta_configuracion):
        print(f"Archivo configuración no encontrado: {ruta_configuracion}")
        return {"bloques": configuraciones, "parametros": parametros}

    actual = {}
    calculos = []
    
    with open(ruta_configuracion, 'r', encoding='utf-8') as archivo:
        for linea in archivo.read().splitlines():
            l = linea.strip()
            if not l or l.startswith("#"):
                continue

            # Parsear parámetros globales
            if l.lower().startswith("parametro:"):
                parte = l.split(":", 1)[1].strip()
                if "=" in parte:
                    nombre, valor = parte.split("=", 1)
                    nombre = nombre.strip().lower()
                    valor = valor.strip().strip('"').strip("'")
                    parametros[nombre] = _cast_valor(valor)
                continue

            # Nueva hoja
            if l.lower().startswith("hoja:"):
                if actual:
                    if calculos:
                        actual["calculos"] = calculos
                    configuraciones.append(actual)
                actual = {}
                calculos = []

            # Parsear líneas con clave:valor
            if ":" in l:
                clave, valor = l.split(":", 1)
                clave_l = clave.strip().lower()
                valor = valor.strip()
                
                if clave_l == "columna calcular":
                    if "=" in valor:
                        nombre_col, formula = valor.split("=", 1)
                        calculos.append({
                            "nombre": nombre_col.strip().strip('"').strip("'").lower(),
                            "formula": formula.strip()
                        })
                else:
                    actual[clave_l] = valor.strip()
        
        if actual:
            if calculos:
                actual["calculos"] = calculos
            configuraciones.append(actual)

    return {"bloques": configuraciones, "parametros": parametros}

def leer_configuracion_separacion(ruta_configuracion: str) -> dict:
    """
    Lee configuración para separar archivos Excel.
    """
    ruta_configuracion = limpiar_ruta(ruta_configuracion)
    if not os.path.exists(ruta_configuracion):
        print(f"Archivo configuración no encontrado: {ruta_configuracion}")
        return {'hojas_calculo': [], 'hojas': []}

    hojas_calculo = []
    hojas = []
    modo = 'hojas_calculo'
    actual_calc = None
    actual_hoja = None

    with open(ruta_configuracion, 'r', encoding='utf-8') as f:
        for linea in f.read().splitlines():
            l = linea.strip()
            if not l or l.startswith("#"):
                continue

            if re.fullmatch(r'-{3,}', l):
                if actual_calc:
                    hojas_calculo.append(actual_calc)
                    actual_calc = None
                modo = 'hojas'
                continue

            if modo == 'hojas_calculo':
                if re.match(r'(?i)^(name|hoja de calculo ?\d*)\s*:', l):
                    if actual_calc:
                        hojas_calculo.append(actual_calc)
                    valor = l.split(":", 1)[1].strip().strip('"').strip("'")
                    actual_calc = {'name': valor, 'identificadores': []}
                    continue
                if l.lower().startswith("identificadores:"):
                    if not actual_calc:
                        continue
                    ids_raw = l.split(":", 1)[1].strip()
                    ids = [x.strip().strip('"').strip("'") for x in ids_raw.split(",") if x.strip()]
                    actual_calc['identificadores'] = ids
                    continue

            if modo == 'hojas':
                if l.lower().startswith("hoja:"):
                    if actual_hoja:
                        hojas.append(actual_hoja)
                    nombre = l.split(":", 1)[1].strip().strip('"').strip("'")
                    actual_hoja = {'hoja': nombre, 'columna_id': ''}
                    continue
                if l.lower().startswith("columna id:"):
                    if not actual_hoja:
                        continue
                    col = l.split(":", 1)[1].strip().strip('"').strip("'")
                    actual_hoja['columna_id'] = col
                    continue

    if actual_calc:
        hojas_calculo.append(actual_calc)
    if actual_hoja:
        hojas.append(actual_hoja)

    return {'hojas_calculo': hojas_calculo, 'hojas': hojas}

def _cast_valor(valor: str) -> Any:
    """
    Intenta convertir un string a int o float, sino devuelve string.
    """
    if re.fullmatch(r'\d+', valor):
        return int(valor)
    elif re.fullmatch(r'\d+\.\d+', valor):
        return float(valor)
    return valor
