import pandas as pd
import os
import glob
import shutil

def limpiar_inventario():
    # Rutas
    ruta_downloads = os.path.expanduser("~/storage/downloads")
    carpeta_imagenes = os.path.join(ruta_downloads, "codigos_generados")

    # Confirmación
    print("⚠️  Esto ELIMINARÁ todas las imágenes QR/códigos de barras")
    print("   y VACIARÁ el contenido de los archivos Excel de inventario.")
    respuesta = input("¿Estás seguro? (S/N): ").strip().upper()
    if respuesta != "S":
        print("Operación cancelada.")
        return

    # 1. Limpiar archivos Excel (solo los de sesión, no el base)
    patrones_sesion = ["INVENTARIO*.xlsx", "inventario*.xlsx"]  # excluye "inventario.xlsx" fijo
    archivos_excel = []
    for p in patrones_sesion:
        archivos_excel.extend(glob.glob(os.path.join(ruta_downloads, p)))
    archivos_excel = list(set(archivos_excel))

    for archivo in archivos_excel:
        try:
            print(f"Procesando Excel: {os.path.basename(archivo)}...")
            df = pd.read_excel(archivo)
            df_vacio = pd.DataFrame(columns=df.columns)
            df_vacio.to_excel(archivo, index=False)
            print(f"   ✅ Limpiado: solo cabeceras conservadas.")
        except Exception as e:
            print(f"   ❌ Error: {e}")

    # 2. Eliminar carpeta de imágenes (si existe)
    if os.path.exists(carpeta_imagenes):
        try:
            shutil.rmtree(carpeta_imagenes)
            print(f"✅ Carpeta de imágenes eliminada: {carpeta_imagenes}")
        except Exception as e:
            print(f"❌ No se pudo eliminar carpeta de imágenes: {e}")
    else:
        print("ℹ️  No se encontró la carpeta de imágenes.")

    print("\n🏁 Limpieza completada.")

if __name__ == "__main__":
    limpiar_inventario()