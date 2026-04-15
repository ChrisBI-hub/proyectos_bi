"""
main.py
=======
Orquestador principal del sistema de inventario QR.
Lanza los tres módulos en hilos paralelos:

  ┌─────────────────────────────────────────────────────────────┐
  │  Hilo 1 → sync_bd.py        Correo → SQL Server  (60 s)    │
  │  Hilo 2 → sinc_qr_drive.py  Correo → Drive       (60 s)    │
  │  Hilo 3 → qr_printer.py     Drive  → PDF → Impresora (1 h) │
  └─────────────────────────────────────────────────────────────┘

Uso:
    python3 main.py          # inicia los tres servicios
    python3 main.py --once   # cada módulo corre una sola vez y sale
"""

import threading
import argparse
import sys
import time
from datetime import datetime

# ──────────────────────────────────────────────────────────────
# Importar los módulos como librerías (no como __main__)
# ──────────────────────────────────────────────────────────────
# Cada script tiene su lógica en funciones, así que se importan
# directamente en lugar de lanzarlos como subprocesos, lo que
# permite compartir el mismo proceso Python y ver todos los logs
# en una sola terminal.
# ──────────────────────────────────────────────────────────────
try:
    import sync_bd
    import sinc_qr_drive as sync_qr
    import qr_printer
except ImportError as e:
    print(f"\n❌ No se pudo importar un módulo: {e}")
    print("   Asegúrate de que main.py está en la misma carpeta que:")
    print("   sync_bd.py  |  sinc_qr_drive.py  |  qr_printer.py\n")
    sys.exit(1)


# ──────────────────────────────────────────────────────────────
# PREFIJO DE LOG POR HILO
# ──────────────────────────────────────────────────────────────
def ts():
    return datetime.now().strftime("%H:%M:%S")


def log(servicio: str, mensaje: str):
    etiquetas = {
        "bd":      "🗄️  [SYNC_BD    ]",
        "drive":   "☁️  [SYNC_DRIVE ]",
        "printer": "🖨️  [QR_PRINTER ]",
        "main":    "🚀 [MAIN       ]",
    }
    print(f"[{ts()}] {etiquetas.get(servicio, servicio)} {mensaje}", flush=True)


# ──────────────────────────────────────────────────────────────
# FUNCIONES DE HILO — envuelven el loop de cada módulo
# ──────────────────────────────────────────────────────────────
def hilo_sync_bd(solo_una_vez: bool):
    """Hilo para sync_bd: actualiza SQL Server desde correo."""
    log("bd", "Iniciando...")
    try:
        if solo_una_vez:
            sync_bd.revisar_una_vez()
        else:
            # El módulo ya tiene su propio while True + sleep
            sync_bd.main_loop()
    except Exception as e:
        log("bd", f"❌ Error inesperado: {e}")


def hilo_sync_drive(solo_una_vez: bool):
    """Hilo para sinc_qr_drive: sube imágenes QR a Google Drive."""
    log("drive", "Iniciando...")
    try:
        # Preparar estado compartido del módulo
        folder_id = sync_qr.extraer_folder_id(sync_qr.DRIVE_FOLDER_ID)
        registro  = sync_qr.cargar_registro()
        service   = sync_qr.conectar_drive()
        log("drive", "✅ Conexión a Drive lista")

        if solo_una_vez:
            sync_qr.revisar_una_vez(service, folder_id, registro)
        else:
            while True:
                sync_qr.revisar_una_vez(service, folder_id, registro)
                log("drive", f"⏳ Próxima revisión en {sync_qr.INTERVALO_SEGUNDOS}s...")
                time.sleep(sync_qr.INTERVALO_SEGUNDOS)

    except Exception as e:
        log("drive", f"❌ Error inesperado: {e}")


def hilo_qr_printer(solo_una_vez: bool):
    """Hilo para qr_printer: descarga Drive → PDF → impresora."""
    log("printer", "Iniciando...")
    try:
        if solo_una_vez:
            qr_printer.ejecutar_ciclo()
        else:
            qr_printer.main_loop()
    except Exception as e:
        log("printer", f"❌ Error inesperado: {e}")


# ──────────────────────────────────────────────────────────────
# MAIN
# ──────────────────────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser(
        description="Orquestador QR — lanza sync_bd, sinc_qr_drive y qr_printer en paralelo"
    )
    parser.add_argument(
        "--once",
        action="store_true",
        help="Cada módulo corre una sola iteración y el programa termina"
    )
    args = parser.parse_args()

    banner = """
╔══════════════════════════════════════════════════════╗
║         SISTEMA QR — ORQUESTADOR PRINCIPAL           ║
║  sync_bd  │  sinc_qr_drive  │  qr_printer            ║
╚══════════════════════════════════════════════════════╝"""
    print(banner)
    log("main", f"Modo: {'--once (una sola vez)' if args.once else 'continuo (Ctrl+C para detener)'}\n")

    # Guardar definición de cada hilo para poder relanzarlos si caen
    definiciones = [
        {"target": hilo_sync_bd,    "args": (args.once,), "name": "sync_bd"},
        {"target": hilo_sync_drive, "args": (args.once,), "name": "sync_drive"},
        {"target": hilo_qr_printer, "args": (args.once,), "name": "qr_printer"},
    ]

    hilos = []
    for d in definiciones:
        h = threading.Thread(target=d["target"], args=d["args"], name=d["name"], daemon=True)
        h.start()
        log("main", f"▶️  Hilo '{h.name}' iniciado (id={h.ident})")
        hilos.append(h)

    try:
        if args.once:
            # Esperar a que todos terminen
            for hilo in hilos:
                hilo.join()
            log("main", "✅ Todos los módulos completaron su ciclo único. Saliendo.")
        else:
            # Mantener el proceso vivo; los hilos son daemon y corren solos
            while True:
                # Reiniciar hilos caídos
        # Reiniciar hilos caídos usando las definiciones originales
                for i, (hilo, d) in enumerate(zip(hilos, definiciones)):
                    if not hilo.is_alive():
                        log("main", f"⚠️  Hilo '{hilo.name}' terminó inesperadamente — reiniciando...")
                        nuevo = threading.Thread(
                            target=d["target"],
                            args=d["args"],
                            name=d["name"],
                            daemon=True,
                        )
                        nuevo.start()
                        hilos[i] = nuevo
                time.sleep(10)

    except KeyboardInterrupt:
        log("main", "\n🛑 Detenido por el usuario. Cerrando hilos...")
        sys.exit(0)


# sync_bd ya expone main_loop() y revisar_una_vez() directamente,
# no se necesita ningún parche dinámico.


if __name__ == "__main__":
    main()