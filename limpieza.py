import pandas as pd
import os
import glob
import shutil
from datetime import datetime

def limpiar_inventario():
    ruta_downloads = os.path.expanduser("~/storage/downloads")
    carpeta_imagenes = os.path.join(ruta_downloads, "codigos_generados")

    print("⚠️  Esto ELIMINARÁ todas las imágenes QR/códigos de barras")
    print("   y los archivos de sesión INVENTARIO*.xlsx generados hoy.")
    respuesta = input("¿Estás seguro? (S/N): ").strip().upper()
    if respuesta != "S":
        print("Operación cancelada.")
        return

    # 1. Eliminar SOLO archivos de sesión (INVENTARIO + fecha, ej: INVENTARIO20260413.xlsx)
    #    El patrón anterior incluía "inventario*.xlsx" que capturaba inventario.xlsx — corregido.
    archivos_sesion = glob.glob(os.path.join(ruta_downloads, "INVENTARIO[0-9]*.xlsx"))

    if archivos_sesion:
        for archivo in archivos_sesion:
            try:
                os.remove(archivo)
                print(f"   ✅ Eliminado: {os.path.basename(archivo)}")
            except Exception as e:
                print(f"   ❌ Error eliminando {os.path.basename(archivo)}: {e}")
    else:
        print("ℹ️  No se encontraron archivos de sesión INVENTARIO*.xlsx.")

    # 2. Eliminar carpeta de imágenes generadas
    if os.path.exists(carpeta_imagenes):
        try:
            shutil.rmtree(carpeta_imagenes)
            print(f"✅ Carpeta de imágenes eliminada: {carpeta_imagenes}")
        except Exception as e:
            print(f"❌ No se pudo eliminar carpeta de imágenes: {e}")
    else:
        print("ℹ️  No se encontró la carpeta de imágenes.")

    print("\n🏁 Limpieza completada.")
    print("   inventario.xlsx no fue modificado (es el archivo base de catálogos).")

if __name__ == "__main__":
    limpiar_inventario()