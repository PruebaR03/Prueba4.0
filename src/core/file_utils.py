import os

def limpiar_ruta(ruta: str) -> str:
    """
    Limpia rutas eliminando prefijos file:/// y caracteres especiales.
    """
    if ruta is None:
        return ""
    ruta = ruta.strip().strip('"')
    if ruta.startswith("file:///"):
        ruta = ruta[8:].replace("%20", " ")
    elif ruta.startswith("file:"):
        ruta = ruta[5:].replace("%20", " ")
    return ruta

def asegurar_carpeta(ruta: str):
    """
    Crea la carpeta si no existe.
    """
    carpeta = os.path.dirname(ruta)
    if carpeta and not os.path.exists(carpeta):
        os.makedirs(carpeta)
