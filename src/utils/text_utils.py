import re
import unicodedata

def parse_lista_columnas(valor: str) -> list:
    """
    Parsea lista de columnas separadas por + o comas.
    """
    if not valor:
        return []
    bruto = valor.strip().replace("+", ",")
    partes = [p.strip() for p in bruto.split(",") if p.strip()]
    limpio = []
    for p in partes:
        if (p.startswith('"') and p.endswith('"')) or (p.startswith("'") and p.endswith("'")):
            p = p[1:-1].strip()
        if p:
            limpio.append(p.lower())
    return limpio

def sin_acentos(s: str) -> str:
    """
    Elimina acentos de un string.
    """
    nfkd = unicodedata.normalize("NFKD", s)
    return "".join(c for c in nfkd if not unicodedata.combining(c))

def alias_columna(nombre: str) -> str:
    """
    Normaliza nombre de columna.
    """
    s = sin_acentos(str(nombre).lower().strip())
    s = re.sub(r"\s+", " ", s)
    return s
