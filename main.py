# Colores ANSI para la consola
CYAN = "\033[96m"
GREEN = "\033[92m"
YELLOW = "\033[93m"
RED = "\033[91m"
RESET = "\033[0m"

from src.procesador import generar_excel_salida, crear_hoja_resumen
from src.enriquecedor import enriquecer_hojas
from src.separador import flujo_separacion

def main():
    print(f"{CYAN}╔════════════════════════════════════════════════════════════╗")
    print(f"║              Bienvenido a BatXcel 🐍📊                     ║")
    print(f"╚════════════════════════════════════════════════════════════╝{RESET}")
    print("\nEste programa te ayudará a procesar y enriquecer archivos Excel.")
    print("Por favor, sigue las instrucciones para configurar tus archivos.\n")

    try:
        print(f"{GREEN}Paso 1: Generar el archivo Excel inicial{RESET}")
        ruta_instrucciones = input("Ingrese la ruta del archivo de instrucciones: ").strip()
        ruta_salida = input("Ingrese la ruta del archivo Excel de salida: ").strip()
        generar_excel_salida(ruta_instrucciones, ruta_salida)
        print(f"{GREEN}✓ Archivo Excel generado exitosamente en: {ruta_salida}{RESET}\n")

        print(f"{GREEN}Paso 1.1: Crear hoja resumen{RESET}")
        crear_hoja_resumen(ruta_salida, ruta_instrucciones)
        print(f"{GREEN}✓ Hoja resumen creada exitosamente en el archivo: {ruta_salida}{RESET}\n")

        print(f"{GREEN}Paso 2: Enriquecer el archivo Excel{RESET}")
        ruta_configuracion = input("Ingrese la ruta del archivo de configuración para el enriquecimiento: ").strip()
        enriquecer_hojas(ruta_salida, ruta_configuracion)
        print(f"{GREEN}✓ Archivo Excel enriquecido exitosamente en: {ruta_salida}{RESET}\n")

        print(f"{GREEN}Paso 3: Separar el archivo Excel en múltiples archivos{RESET}")
        separar = input(f"{YELLOW}¿Desea separar el Excel en múltiples archivos según identificadores? (s/n): {RESET}").strip().lower()

        if separar == 's':
            ruta_config_separacion = input("Ingrese la ruta del archivo de configuración para la separación: ").strip()
            carpeta_salida = input("Ingrese la carpeta de salida para los archivos separados (default: output/separados): ").strip() or "output/separados"
            flujo_separacion(ruta_salida, ruta_config_separacion, carpeta_salida, ruta_instrucciones)
            print(f"{GREEN}✓ Archivos separados generados exitosamente en: {carpeta_salida}{RESET}\n")
        else:
            print(f"{YELLOW}Paso 3 omitido. No se realizó la separación.{RESET}\n")

        print(f"{CYAN}╔════════════════════════════════════════════════════════════╗")
        print(f"║         PROCESO COMPLETADO EXITOSAMENTE                    ║")
        print(f"╚════════════════════════════════════════════════════════════╝{RESET}")

    except FileNotFoundError as e:
        print(f"{RED}✗ Error: Archivo no encontrado - {e}{RESET}")
    except Exception as e:
        print(f"{RED}✗ Error inesperado: {e}{RESET}")

if __name__ == "__main__":
    main()
