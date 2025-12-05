CYAN = "\033[96m"
GREEN = "\033[92m"
YELLOW = "\033[93m"
RED = "\033[91m"
RESET = "\033[0m"

from src.procesador import generar_excel_salida, crear_hoja_resumen
from src.enriquecedor import enriquecer_hojas, limpiar_datos_enriquecidos
from src.separador import flujo_separacion
from src.core.file_utils import limpiar_ruta
from src.ppt.ppt_comparador import generar_ppt_comparativo
import os
import glob

def main():
    print(f"{CYAN}╔════════════════════════════════════════════════════════════╗")
    print(f"║              Bienvenido a BatXcel 🐍📊                     ║")
    print(f"╚════════════════════════════════════════════════════════════╝{RESET}")
    print("\nEste programa te ayudará a procesar y enriquecer archivos Excel.")
    print("Por favor, sigue las instrucciones para configurar tus archivos.\n")

    try:
        print(f"{YELLOW}{'═' * 60}")
        print(f"MENÚ PRINCIPAL")
        print(f"{'═' * 60}{RESET}")
        print("1. Procesar Excel (Pasos 1-3)")
        print("2. Generar presentación PPT comparativa")
        print("3. Salir")
        
        opcion = input(f"\n{YELLOW}Selecciona una opción (1/2/3): {RESET}").strip()
        
        if opcion == "1":
            procesar_excel_completo()
        elif opcion == "2":
            generar_ppt_solo()
        elif opcion == "3":
            print(f"{GREEN}Gracias por usar BatXcel. ¡Hasta luego!{RESET}")
            return
        else:
            print(f"{RED}Opción no válida{RESET}")

    except FileNotFoundError as e:
        print(f"{RED}✗ Error: Archivo no encontrado - {e}{RESET}")
    except KeyboardInterrupt:
        print(f"\n{YELLOW}⚠️  Proceso interrumpido por el usuario{RESET}")
    except Exception as e:
        print(f"{RED}✗ Error inesperado: {e}{RESET}")
        import traceback
        traceback.print_exc()
    finally:
        print(f"{CYAN}╔════════════════════════════════════════════════════════════╗")
        print(f"║                     Proceso finalizado                    ║")
        print(f"╚════════════════════════════════════════════════════════════╝{RESET}")

def procesar_excel_completo():
    """Flujo completo de procesamiento: Pasos 1-3 y opcionalmente PPT."""
    
    # PASO 1: Generación de archivo base y resumen
    print(f"\n{GREEN}{'═' * 60}")
    print(f"PASO 1: Generar archivo Excel base")
    print(f"{'═' * 60}{RESET}")
    
    ruta_instrucciones = input("📝 Ruta del archivo de instrucciones: ").strip()
    ruta_salida = input("📂 Ruta del archivo Excel de salida: ").strip()
    
    generar_excel_salida(ruta_instrucciones, ruta_salida)
    print(f"{GREEN}✓ Archivo Excel generado exitosamente{RESET}\n")

    print(f"{GREEN}{'═' * 60}")
    print(f"PASO 1.1: Crear hoja(s) resumen")
    print(f"{'═' * 60}{RESET}")
    
    crear_hoja_resumen(ruta_salida, ruta_instrucciones)
    print(f"{GREEN}✓ Hoja(s) resumen creada(s) exitosamente{RESET}\n")

    # PASO 2: Enriquecimiento opcional
    print(f"{YELLOW}{'═' * 60}")
    print(f"PASO 2: Enriquecer archivo Excel (OPCIONAL)")
    print(f"{'═' * 60}{RESET}")
    
    enriquecer = input(f"{YELLOW}¿Desea enriquecer el archivo Excel? (s/n): {RESET}").strip().lower()
    
    if enriquecer == 's':
        ruta_configuracion = input("📝 Ruta del archivo de configuración para enriquecimiento: ").strip()
        ruta_configuracion = limpiar_ruta(ruta_configuracion)
        
        if not os.path.exists(ruta_configuracion):
            print(f"{RED}✗ Archivo de configuración no encontrado: {ruta_configuracion}")
            print(f"⚠️  Saltando enriquecimiento...{RESET}\n")
        else:
            enriquecer_hojas(ruta_salida, ruta_configuracion)
            print(f"{GREEN}✓ Archivo Excel enriquecido exitosamente{RESET}\n")
    else:
        print(f"{YELLOW}⊘ Paso 2 omitido. No se realizó enriquecimiento.{RESET}\n")

    # PASO 3: Separación opcional
    carpeta_salida = None
    print(f"{YELLOW}{'═' * 60}")
    print(f"PASO 3: Separar Excel en múltiples archivos (OPCIONAL)")
    print(f"{'═' * 60}{RESET}")
    
    separar = input(f"{YELLOW}¿Desea separar el Excel según identificadores? (s/n): {RESET}").strip().lower()

    if separar == 's':
        ruta_config_separacion = input("📝 Ruta del archivo de configuración para separación: ").strip()
        ruta_config_separacion = limpiar_ruta(ruta_config_separacion)
        
        if not os.path.exists(ruta_config_separacion):
            print(f"{RED}✗ Archivo de configuración no encontrado: {ruta_config_separacion}")
            print(f"⚠️  Saltando separación...{RESET}\n")
        else:
            carpeta_salida = input("📂 Carpeta de salida para archivos separados (Enter=output/separados): ").strip() or "output/separados"
            
            flujo_separacion(ruta_salida, ruta_config_separacion, carpeta_salida, ruta_instrucciones)
            print(f"{GREEN}✓ Archivos separados generados exitosamente{RESET}")
            print(f"{GREEN}ℹ️  Las hojas ya contienen el enriquecimiento del archivo base{RESET}\n")
            
            print(f"{YELLOW}{'─' * 60}")
            print(f"PASO 3.1: Limpieza de datos enriquecidos (OPCIONAL)")
            print(f"{'─' * 60}{RESET}")
            print(f"{YELLOW}ℹ️  Elimina filas según criterios en TODOS los archivos separados.")
            print(f"   Ejemplo: Eliminar dispositivos actualizados y compliant.{RESET}")
            
            limpiar_separados = input(f"{YELLOW}¿Desea limpiar datos enriquecidos? (s/n): {RESET}").strip().lower()
            
            if limpiar_separados == 's':
                ruta_config_limpieza = input("📝 Ruta del archivo de configuración para limpieza: ").strip()
                ruta_config_limpieza = limpiar_ruta(ruta_config_limpieza)
                
                if not os.path.exists(ruta_config_limpieza):
                    print(f"{RED}✗ Archivo de configuración no encontrado: {ruta_config_limpieza}{RESET}\n")
                else:
                    print(f"\n{GREEN}{'═' * 70}")
                    print(f"🧹 LIMPIANDO ARCHIVOS SEPARADOS")
                    print(f"{'═' * 70}{RESET}")
                    
                    archivos_separados = glob.glob(os.path.join(carpeta_salida, "*.xlsx"))
                    
                    print(f"\n📂 Carpeta: {carpeta_salida}")
                    print(f"📋 Archivos a limpiar: {len(archivos_separados)}\n")
                    
                    for idx, archivo in enumerate(archivos_separados, 1):
                        print(f"[{idx}/{len(archivos_separados)}] ", end="")
                        try:
                            limpiar_datos_enriquecidos(archivo, ruta_config_limpieza)
                        except Exception as e:
                            print(f"  {RED}⚠️ Error: {e}{RESET}")
                    
                    print(f"\n{GREEN}{'═' * 70}")
                    print(f"✅ LIMPIEZA COMPLETADA")
                    print(f"{'═' * 70}{RESET}\n")
            else:
                print(f"{YELLOW}⊘ Limpieza omitida.{RESET}\n")
    else:
        print(f"{YELLOW}⊘ Paso 3 omitido. No se realizó separación.{RESET}\n")

    print(f"{CYAN}╔════════════════════════════════════════════════════════════╗")
    print(f"║         ✓ PROCESO COMPLETADO EXITOSAMENTE ✓               ║")
    print(f"╚════════════════════════════════════════════════════════════╝{RESET}")
    
    print(f"\n{GREEN}📊 Resumen de operaciones:{RESET}")
    print(f"  ✓ Archivo base generado: {ruta_salida}")
    if enriquecer == 's':
        print(f"  ✓ Enriquecimiento aplicado")
    if separar == 's':
        print(f"  ✓ Archivos separados generados en: {carpeta_salida}")
    print()

    # PASO 4: Generación de PPT comparativo
    generar_ppt = input(f"{YELLOW}¿Desea generar presentación comparativa (PPT)? (s/n): {RESET}").strip().lower()
    if generar_ppt == 's':
        # Si se ejecutó paso 3 (separación), usar esa carpeta como actual
        # Si NO se ejecutó paso 3, pedir la ruta del Excel actual (separados)
        if carpeta_salida:
            carpeta_actual = carpeta_salida
            print(f"\n📂 Usando carpeta actual: {carpeta_actual}")
        else:
            carpeta_actual = input(f"\n📂 Carpeta con archivos separados actuales (obligatorio): ").strip()
            carpeta_actual = limpiar_ruta(carpeta_actual)
        
        # Siempre pedir la carpeta anterior
        carpeta_anterior = input("📂 Carpeta con archivos de la semana pasada (obligatorio): ").strip()
        carpeta_anterior = limpiar_ruta(carpeta_anterior)

        ruta_plantilla = input("🖼️ Ruta del archivo plantilla PPTX (obligatorio): ").strip()
        ruta_plantilla = limpiar_ruta(ruta_plantilla)

        ruta_salida_ppt = input("📄 Ruta de salida para el PPT generado (ej: output/comparativo.pptx): ").strip()
        if not ruta_salida_ppt:
            ruta_salida_ppt = "output/comparativo_resultado.pptx"
        # Validar que termine en .pptx
        if not ruta_salida_ppt.lower().endswith('.pptx'):
            ruta_salida_ppt = ruta_salida_ppt.rstrip('\\').rstrip('/') + "/comparativo_resultado.pptx"
        ruta_salida_ppt = limpiar_ruta(ruta_salida_ppt)

        # Validaciones
        if not os.path.exists(carpeta_actual):
            print(f"{RED}✗ Carpeta actual no encontrada: {carpeta_actual}{RESET}")
        elif not os.path.exists(carpeta_anterior):
            print(f"{RED}✗ Carpeta anterior no encontrada: {carpeta_anterior}{RESET}")
        elif not os.path.exists(ruta_plantilla):
            print(f"{RED}✗ Plantilla PPT no encontrada: {ruta_plantilla}{RESET}")
        else:
            # Asegurar que la carpeta de salida existe (solo la carpeta, no el archivo)
            carpeta_salida_ppt = os.path.dirname(ruta_salida_ppt)
            if carpeta_salida_ppt and carpeta_salida_ppt.strip():
                os.makedirs(carpeta_salida_ppt, exist_ok=True)
            try:
                print(f"\n{GREEN}Generando presentación comparativa...{RESET}")
                generar_ppt_comparativo(carpeta_actual, carpeta_anterior, ruta_plantilla, ruta_salida_ppt)
            except Exception as e:
                print(f"{RED}✗ Error generando PPT: {e}{RESET}")

def generar_ppt_solo():
    """Genera PPT comparativo sin ejecutar los pasos 1-3."""
    
    print(f"\n{CYAN}{'═' * 60}")
    print(f"GENERADOR DE PRESENTACIÓN COMPARATIVA")
    print(f"{'═' * 60}{RESET}\n")
    
    while True:
        ruta_excel_actual = input("📄 Ruta del Excel ACTUAL: ").strip()
        ruta_excel_actual = limpiar_ruta(ruta_excel_actual)
        
        if not os.path.exists(ruta_excel_actual):
            print(f"{RED}✗ Excel actual no encontrado{RESET}\n")
            continue
        break
    
    while True:
        ruta_excel_anterior = input("📄 Ruta del Excel de la SEMANA PASADA: ").strip()
        ruta_excel_anterior = limpiar_ruta(ruta_excel_anterior)
        
        if not os.path.exists(ruta_excel_anterior):
            print(f"{RED}✗ Excel anterior no encontrado{RESET}\n")
            continue
        break

    while True:
        ruta_plantilla = input("🖼️ Ruta del archivo plantilla PPTX: ").strip()
        ruta_plantilla = limpiar_ruta(ruta_plantilla)
        
        if not os.path.exists(ruta_plantilla):
            print(f"{RED}✗ Plantilla PPT no encontrada{RESET}\n")
            continue
        break

    ruta_instrucciones = input("📝 Ruta del archivo instrucciones.txt (Enter=omitir): ").strip()
    if ruta_instrucciones:
        ruta_instrucciones = limpiar_ruta(ruta_instrucciones)
        if not os.path.exists(ruta_instrucciones):
            print(f"{YELLOW}⚠️  Instrucciones no encontradas, continuando sin categorías{RESET}")
            ruta_instrucciones = None
    else:
        ruta_instrucciones = None

    ruta_salida_ppt = input("📄 Ruta de salida para el PPT (ej: output/comparativo.pptx): ").strip()
    if not ruta_salida_ppt:
        ruta_salida_ppt = "output/comparativo_resultado.pptx"
    # Validar que termine en .pptx
    if not ruta_salida_ppt.lower().endswith('.pptx'):
        # Agregar .pptx solo si no termina en eso
        ruta_salida_ppt = ruta_salida_ppt.rstrip('\\').rstrip('/') + "/comparativo_resultado.pptx"
    ruta_salida_ppt = limpiar_ruta(ruta_salida_ppt)

    # Crear carpeta de salida
    carpeta_salida_ppt = os.path.dirname(ruta_salida_ppt)
    if carpeta_salida_ppt and carpeta_salida_ppt.strip():
        os.makedirs(carpeta_salida_ppt, exist_ok=True)
    
    try:
        print(f"\n{GREEN}Generando presentación comparativa...{RESET}")
        generar_ppt_comparativo(ruta_excel_actual, ruta_excel_anterior, ruta_plantilla, ruta_salida_ppt, ruta_instrucciones)
        print(f"\n{GREEN}{'═' * 60}")
        print(f"✅ PRESENTACIÓN GENERADA EXITOSAMENTE")
        print(f"📁 Ubicación: {ruta_salida_ppt}")
        print(f"{'═' * 60}{RESET}\n")
    except Exception as e:
        print(f"{RED}✗ Error generando PPT: {e}{RESET}\n")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    while True:
        main()
        repetir = input(f"\n{YELLOW}¿Desea realizar otra operación? (s/n): {RESET}").strip().lower()
        if repetir != 's':
            break
    print(f"{GREEN}Gracias por usar BatXcel. ¡Hasta luego!{RESET}")
