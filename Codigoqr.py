"""
CodigoQR.py  (Termux)
=====================
Captura inventario, genera QR/barras y al finalizar envía el
archivo INVENTARIO{fecha}.xlsx por correo con el asunto
"ACTUALIZACION_BASE_DE_DATOS" para que el script sync_bd.py
en Kali lo detecte y lo suba al servidor SQL.

NO requiere acceso directo a SQL Server.
"""

import sys
import subprocess
import os

# ──────────────────────────────────────────────
# AUTO-INSTALACIÓN (Termux)
# ──────────────────────────────────────────────
def instalar_paquetes_sistema():
    paquetes = ["libjpeg-turbo", "libpng", "zlib", "python"]
    print("📦 Verificando paquetes del sistema Termux...")
    for paquete in paquetes:
        try:
            r = subprocess.run(["pkg", "install", "-y", paquete], capture_output=True, text=True)
            print(f"   {'✅' if r.returncode == 0 else '⚠️ '} {paquete}")
        except FileNotFoundError:
            break  # no estamos en Termux, ignorar

def instalar_pip_paquetes():
    requeridos = {
        "pandas":   "pandas",
        "qrcode":   "qrcode[pil]",
        "barcode":  "python-barcode",
        "openpyxl": "openpyxl",
        "Pillow":   "Pillow",
    }
    print("\n🐍 Verificando paquetes Python...")
    faltantes = []
    for modulo, paquete in requeridos.items():
        try:
            __import__(modulo)
            print(f"   ✅ {modulo}")
        except ImportError:
            print(f"   ⬇️  {modulo} → se instalará")
            faltantes.append(paquete)

    if faltantes:
        print(f"\n⚙️  Instalando {len(faltantes)} paquete(s)...")
        for paquete in faltantes:
            print(f"   {paquete}...", end=" ", flush=True)
            r = subprocess.run(
                [sys.executable, "-m", "pip", "install", "--quiet", paquete],
                capture_output=True, text=True
            )
            print("✅" if r.returncode == 0 else f"❌ {r.stderr.strip()}")
            if r.returncode != 0:
                sys.exit(1)
        print("\n✅ Dependencias listas.\n")
    else:
        print("   Sin instalaciones pendientes.\n")

instalar_paquetes_sistema()
instalar_pip_paquetes()

# ──────────────────────────────────────────────
# IMPORTS
# ──────────────────────────────────────────────
import smtplib
from datetime import datetime
from email.message import EmailMessage

import pandas as pd
import qrcode
import barcode
from barcode.writer import ImageWriter

# ──────────────────────────────────────────────
# CONFIGURACIÓN
# ──────────────────────────────────────────────
archivo_excel  = '/sdcard/Download/inventario.xlsx'
carpeta_salida = '/sdcard/Download/codigos_generados'

FECHA_HOY      = datetime.now().strftime("%Y%m%d")
ARCHIVO_SESION = f"INVENTARIO{FECHA_HOY}.xlsx"

EDIFICIOS = ["NORTE 180", "NORTE 182", "PATIO SEC"]

# ── Correo ────────────────────────────────────
SMTP_SERVER   = "smtp.gmail.com"
SMTP_PORT     = 465
EMAIL_USER    = "reportes.bi@abcsc.mx"
EMAIL_PASS    = "jwvjdrvmprzrwzxy"

# Asunto especial que sync_bd.py espera para disparar la actualización
ASUNTO_TRIGGER = "ACTUALIZACION_BASE_DE_DATOS"

DESTINATARIOS = [
    "ccarbajal@abcsc.mx",
    "myanez@abcsc.mx",
    "sgonzalez@abcsc.mx",
    "ymontoya@abcsc.mx",
    "reportes.bi@abcsc.mx",
]

os.makedirs(carpeta_salida, exist_ok=True)

# ──────────────────────────────────────────────
# HELPERS
# ──────────────────────────────────────────────
def extraer_tres_letras(nombre: str) -> str:
    limpio = nombre.replace(" ", "").upper()
    if len(limpio) == 0: return "XXX"
    if len(limpio) == 1: return limpio * 3
    if len(limpio) == 2: return limpio + "X"
    return f"{limpio[0]}{limpio[len(limpio) // 2]}{limpio[-1]}"


def seleccionar_opcion(lista: list, titulo: str) -> str:
    print(f"\n{titulo}")
    for i, item in enumerate(lista, 1):
        print(f"   {i}. {item}")
    while True:
        try:
            idx = int(input("   Selecciona número: ")) - 1
            if 0 <= idx < len(lista):
                return lista[idx]
            print(f"   ⚠️  Elige entre 1 y {len(lista)}.")
        except ValueError:
            print("   ⚠️  Ingresa un número válido.")


# ──────────────────────────────────────────────
# GUARDAR SESIÓN EN INVENTARIO{fecha}.xlsx
# ──────────────────────────────────────────────
def guardar_sesion_excel(registros: list) -> str:
    df_nuevo = pd.DataFrame(registros)
    if os.path.exists(ARCHIVO_SESION):
        df_final = pd.concat([pd.read_excel(ARCHIVO_SESION), df_nuevo], ignore_index=True)
    else:
        df_final = df_nuevo
    df_final.to_excel(ARCHIVO_SESION, index=False)
    print(f"💾 Sesión guardada → {ARCHIVO_SESION}  ({len(df_final)} registros totales)")
    return ARCHIVO_SESION


# ──────────────────────────────────────────────
# ENVÍO DE CORREO CON ASUNTO TRIGGER
# ──────────────────────────────────────────────
def enviar_correo(archivos_adjuntos: list, total_registros: int):
    """
    Envía el archivo INVENTARIO{fecha}.xlsx con el asunto
    ACTUALIZACION_BASE_DE_DATOS para que sync_bd.py lo detecte.
    """
    print("\n📧 Enviando correo con datos de sesión...")
    hoy = datetime.now().strftime("%d/%m/%Y %H:%M")

    msg = EmailMessage()
    msg["Subject"] = ASUNTO_TRIGGER          # ← clave para sync_bd.py
    msg["From"]    = EMAIL_USER
    msg["To"]      = ", ".join(DESTINATARIOS)
    msg.set_content(
        f"Sesión de inventario: {hoy}\n"
        f"Registros capturados: {total_registros}\n"
        f"Archivo adjunto: {ARCHIVO_SESION}\n\n"
        f"Este correo será procesado automáticamente por sync_bd.py.\n"
        f"— Sistema BI · Inventario"
    )

    for archivo in archivos_adjuntos:
        if not os.path.exists(archivo):
            print(f"   ⚠️  No encontrado, omitido: {archivo}")
            continue
        with open(archivo, "rb") as f:
            datos = f.read()
        nombre = os.path.basename(archivo)
        ext    = nombre.rsplit(".", 1)[-1].lower()
        if ext in ("png", "jpg"):
            msg.add_attachment(datos, maintype="image", subtype=ext, filename=nombre)
        elif ext == "xlsx":
            msg.add_attachment(
                datos, maintype="application",
                subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                filename=nombre,
            )
        else:
            msg.add_attachment(datos, maintype="application", subtype="octet-stream", filename=nombre)

    try:
        with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT) as smtp:
            smtp.login(EMAIL_USER, EMAIL_PASS)
            smtp.send_message(msg)
        print(f"✅ Correo enviado con asunto [{ASUNTO_TRIGGER}]")
        print(f"   Destinatarios: {', '.join(DESTINATARIOS)}")
    except smtplib.SMTPAuthenticationError:
        print("❌ Error de autenticación SMTP.")
    except Exception as e:
        print(f"❌ Error al enviar: {e}")


# ──────────────────────────────────────────────
# FLUJO PRINCIPAL
# ──────────────────────────────────────────────
def ejecutar_sistema():
    archivos_sesion  = []
    registros_sesion = []

    while True:
        try:
            xls      = pd.ExcelFile(archivo_excel)
            df_mat   = pd.read_excel(xls, "Materiales")
            empresas = pd.read_excel(xls, "Empresa")["Empresa"].dropna().tolist()
            areas    = pd.read_excel(xls, "Areas")["Area"].dropna().tolist()

            print("\n" + "=" * 45)
            print(f"  REGISTRO #{len(registros_sesion)+1}  |  {datetime.now().strftime('%d/%m/%Y %H:%M')}")
            print("=" * 45)

            nombre_mat   = input("\nNombre del Material : ").strip()
            letras_auto  = extraer_tres_letras(nombre_mat)
            print(f"✨ Abreviatura: {letras_auto}")

            desc         = input("Descripción         : ").strip()
            emp_sel      = seleccionar_opcion(empresas, "🏢 Empresa:")
            edificio_sel = seleccionar_opcion(EDIFICIOS, "🏗️  Edificio:")
            area_sel     = seleccionar_opcion(areas,    "📍 Área:")

            correlativo = len(df_mat) + len(registros_sesion) + 1
            codigo_id   = f"{letras_auto}{correlativo:04d}"
            timestamp   = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            print(f"\n🔑 Código asignado: {codigo_id}")

            # QR
            ruta_qr  = os.path.join(carpeta_salida, f"qr_{codigo_id}.png")
            datos_qr = (
                f"ID:{codigo_id}\nMAT:{nombre_mat}\nDESC:{desc}\n"
                f"ORG:{emp_sel}\nEDIF:{edificio_sel}\nLOC:{area_sel}"
            )
            qrcode.make(datos_qr).save(ruta_qr)
            archivos_sesion.append(ruta_qr)
            print(f"   📷 QR       → {ruta_qr}")

            # Barras
            bn_class          = barcode.get_barcode_class("code128")
            ruta_barras_base  = os.path.join(carpeta_salida, f"barras_{codigo_id}")
            ruta_barras_final = bn_class(codigo_id, writer=ImageWriter()).save(ruta_barras_base)
            archivos_sesion.append(ruta_barras_final)
            print(f"   📊 Barras   → {ruta_barras_final}")

            # Actualizar catálogo inventario.xlsx
            nueva_fila_cat = {
                "Material": nombre_mat, "Abreviatura": letras_auto,
                "Descripción": desc,    "Empresa": emp_sel,
                "Edificio": edificio_sel, "Area": area_sel,
            }
            with pd.ExcelWriter(archivo_excel, mode="a", if_sheet_exists="overlay", engine="openpyxl") as writer:
                pd.concat([df_mat, pd.DataFrame([nueva_fila_cat])], ignore_index=True
                          ).to_excel(writer, sheet_name="Materiales", index=False)

            registros_sesion.append({
                "CodigoID":    codigo_id,
                "Material":    nombre_mat,
                "Abreviatura": letras_auto,
                "Descripcion": desc,
                "Empresa":     emp_sel,
                "Edificio":    edificio_sel,
                "Area":        area_sel,
                "RutaQR":      ruta_qr,
                "RutaBarras":  ruta_barras_final,
                "FechaHora":   timestamp,
            })

            print(f"\n✅ Registrado: {codigo_id}")

            if input("\n¿Registrar otro? (S/N): ").strip().upper() == "N":
                break

        except (IndexError, ValueError) as e:
            print(f"⚠️  Entrada inválida: {e}. Intenta de nuevo.")
        except KeyboardInterrupt:
            print("\n⚠️  Interrumpido.")
            break
        except Exception as e:
            print(f"❌ Error: {e}")
            break

    # ── CIERRE ────────────────────────────────
    if not registros_sesion:
        print("\nℹ️  Sin registros. Fin.")
        return

    print(f"\n{'='*45}")
    print(f"  CERRANDO — {len(registros_sesion)} registro(s) capturado(s)")
    print(f"{'='*45}")

    archivo_sesion_path = guardar_sesion_excel(registros_sesion)
    # Solo correo — sin SQL directo
    enviar_correo(archivos_sesion + [archivo_sesion_path], len(registros_sesion))

    print("\n🏁 ¡Proceso terminado! sync_bd.py en Kali subirá los datos.")


if __name__ == "__main__":
    ejecutar_sistema()