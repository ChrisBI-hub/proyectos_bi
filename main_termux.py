#!/usr/bin/env python3
"""
main.py - Orquestador del sistema de inventario en Termux
Ejecuta Codigoqr.py y al terminar (envío exitoso) ejecuta limpieza.py
"""

import subprocess
import sys
import os

def main():
    print("\n" + "=" * 60)
    print("  SISTEMA DE INVENTARIO COMPLETO")
    print("  Captura → Envío → Limpieza automática")
    print("=" * 60)

    # 1. Ejecutar el script de captura
    print("\n📱 Iniciando Codigoqr.py...\n")
    resultado = subprocess.run([sys.executable, "Codigoqr.py"])

    # 2. Verificar si terminó correctamente (código 0)
    if resultado.returncode == 0:
        print("\n✅ Codigoqr.py finalizó exitosamente.")
        print("🧹 Ejecutando limpieza automática...\n")

        # 3. Ejecutar limpieza
        subprocess.run([sys.executable, "limpieza.py"])
    else:
        print("\n⚠️  Codigoqr.py terminó con errores (código {}).".format(resultado.returncode))
        print("   No se ejecutará la limpieza automática para evitar pérdida de datos.")
        print("   Revisa los mensajes anteriores o ejecuta limpieza.py manualmente si estás seguro.")

    print("\n🏁 Proceso completado.")

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n🛑 Interrupción detectada. Saliendo...")
        sys.exit(1)