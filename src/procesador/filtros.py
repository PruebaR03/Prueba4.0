import pandas as pd

def evaluar_subcriterio(df: pd.DataFrame, columna_criterio: str, subcriterio: str) -> pd.DataFrame:
    """
    Evalúa un subcriterio y retorna filas que lo cumplen.
    
    Operadores:
    - *+valor+*  → Diferente de lo que contiene "valor"
    - *valor*    → Diferente de "valor"
    - +valor+    → Contiene "valor"
    - "valor"    → Igual a "valor"
    - []         → Vacío
    - *[]*       → No vacío
    """
    df[columna_criterio] = df[columna_criterio].astype(str).str.strip()
    
    # *+valor+* - Diferente de contiene
    if subcriterio.startswith("*+") and subcriterio.endswith("+*"):
        valor = subcriterio[2:-2]
        print(f"  → Filtrando: diferente de lo que contiene '{valor}'")
        return df[~df[columna_criterio].str.contains(valor, na=False, case=False)]
    
    # *[]* - No vacío
    elif subcriterio == "*[]*":
        print(f"  → Filtrando: valores no vacíos")
        return df[(df[columna_criterio] != "") & (df[columna_criterio] != "[]") & (~df[columna_criterio].isna())]
    
    # *valor* - Diferente exacto
    elif subcriterio.startswith("*") and subcriterio.endswith("*"):
        valor = subcriterio.strip("*")
        print(f"  → Filtrando: diferente de '{valor}'")
        return df[df[columna_criterio] != valor]
    
    # [] - Vacío
    elif subcriterio == "[]":
        print(f"  → Filtrando: valores vacíos o '[]'")
        return df[(df[columna_criterio] == "") | (df[columna_criterio] == "[]") | (df[columna_criterio].isna())]
    
    # +valor+ - Contiene
    elif subcriterio.startswith("+") and subcriterio.endswith("+"):
        valor = subcriterio.strip("+")
        print(f"  → Filtrando: contiene '{valor}'")
        return df[df[columna_criterio].str.contains(valor, na=False, case=False)]
    
    # "valor" - Igual con comillas
    elif subcriterio.startswith('"') and subcriterio.endswith('"'):
        valor = subcriterio.strip('"')
        print(f"  → Filtrando: igual a '{valor}'")
        return df[df[columna_criterio] == valor]
    
    # valor - Igual por defecto
    else:
        print(f"  → Filtrando: igual a '{subcriterio}'")
        return df[df[columna_criterio] == subcriterio]

def aplicar_criterio(df: pd.DataFrame, columna_criterio: str, criterio: str) -> pd.DataFrame:
    """
    Aplica criterio completo con soporte para OR (||) y AND (&&).
    """
    if not criterio:
        print("El criterio está vacío. Se incluirán todas las filas.")
        return df

    print(f"  Criterio: '{criterio}'")

    # OR lógico
    if "||" in criterio:
        subcriterios = [c.strip() for c in criterio.split("||")]
        df_filtrado = pd.DataFrame()
        for subcriterio in subcriterios:
            sub_df = evaluar_subcriterio(df, columna_criterio, subcriterio)
            if sub_df is not None:
                df_filtrado = pd.concat([df_filtrado, sub_df]).drop_duplicates()
        print(f"  ✅ {len(df_filtrado)} filas cumplen al menos un subcriterio\n")
        return df_filtrado
    
    # AND lógico
    elif "&&" in criterio:
        subcriterios = [c.strip() for c in criterio.split("&&")]
        df_filtrado = df.copy()
        for subcriterio in subcriterios:
            df_filtrado = evaluar_subcriterio(df_filtrado, columna_criterio, subcriterio)
        print(f"  ✅ {len(df_filtrado)} filas cumplen todos los subcriterios\n")
        return df_filtrado
    
    # Criterio simple
    else:
        df_filtrado = evaluar_subcriterio(df, columna_criterio, criterio)
        print(f"  ✅ {len(df_filtrado)} filas cumplen el criterio\n")
        return df_filtrado
