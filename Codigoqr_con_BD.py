"""
CodigoQR.py
===========
Sistema de generación de códigos QR y de barras para inventario.
Auto-instala todas las dependencias necesarias en Termux.

Funcionalidades:
  - Captura: Material, Descripción, Empresa, Edificio, Área
  - Genera QR y código de barras por cada registro
  - Guarda sesión completa en INVENTARIO{fecha}.xlsx
  - Sube los registros nuevos a SQL Server al finalizar
  - Envía los códigos generados por correo
"""

import sys
import subprocess
import os

# ──────────────────────────────────────────────
# AUTO-INSTALACIÓN DE DEPENDENCIAS (Termux)
# ──────────────────────────────────────────────
def instalar_paquetes_sistema():
    paquetes = ["libjpeg-turbo", "libpng", "zlib", "python"]
    print("📦 Verificando paquetes del sistema Termux...")
    for paquete in paquetes:
        try:
            resultado = subprocess.run(["pkg", "install", "-y", paquete], capture_output=True, text=True)
            print(f"   {'✅' if resultado.returncode == 0 else '⚠️ '} {paquete}")
        except FileNotFoundError:
            break

def instalar_pip_paquetes():
    requeridos = {
        "pandas":     "pandas",
        "qrcode":     "qrcode[pil]",
        "barcode":    "python-barcode",
        "openpyxl":   "openpyxl",
        "Pillow":     "Pillow",
        "sqlalchemy": "sqlalchemy",
        "pyodbc":     "pyodbc",
    }
    print("\n🐍 Verificando paquetes Python...")
    faltantes = []
    for modulo, paquete in requeridos.items():
        try:
            __import__(modulo)
            print(f"   ✅ {modulo}")
        except ImportError:
            print(f"   ⬇️  {modulo} no encontrado → se instalará")
            faltantes.append(paquete)

    if faltantes:
        print(f"\n⚙️  Instalando {len(faltantes)} paquete(s)...")
        for paquete in faltantes:
            print(f"   Instalando {paquete}...", end=" ", flush=True)
            r = subprocess.run([sys.executable, "-m", "pip", "install", "--quiet", paquete], capture_output=True, text=True)
            if r.returncode == 0:
                print("✅")
            else:
                print(f"❌\n   Error: {r.stderr.strip()}")
                sys.exit(1)
        print("\n✅ Todas las dependencias instaladas.\n")
    else:
        print("   Todo listo, sin instalaciones pendientes.\n")

instalar_paquetes_sistema()
instalar_pip_paquetes()

# ──────────────────────────────────────────────
# IMPORTS
# ──────────────────────────────────────────────
import urllib
import smtplib
from datetime import datetime
from email.message import EmailMessage

import pandas as pd
import qrcode
import barcode
from barcode.writer import ImageWriter
from sqlalchemy import create_engine, text

# ──────────────────────────────────────────────
# CONFIGURACIÓN
# ──────────────────────────────────────────────
archivo_excel  = '/home/christian/Documentos/proyectos_bi/inventario.xlsx'
carpeta_salida = '/home/christian/Documentos/proyectos_bi/codigos_generados'

FECHA_HOY      = datetime.now().strftime("%Y%m%d")
ARCHIVO_SESION = f"INVENTARIO{FECHA_HOY}.xlsx"

EDIFICIOS = ["NORTE 180", "NORTE 182", "PATIO SEC"]

SQL_CONFIG = {
    "server":   "150.1.1.152",
    "database": "BI",
    "username": "ConsultaBD",
    "password": "5D$bc#kM&5W2T8J40?s%",
    "driver":   "{ODBC Driver 17 for SQL Server}",
}
TABLA_SQL = "Inventario.Codigos_QR"   # ajusta schema.tabla según tu BD

SMTP_SERVER   = "smtp.gmail.com"
SMTP_PORT     = 465
EMAIL_USER    = "reportes.bi@abcsc.mx"
EMAIL_PASS    = "jwvjdrvmprzrwzxy"
DESTINATARIOS = [
    "ccarbajal@abcsc.mx",
    "myanez@abcsc.mx",
    "sgonzalez@abcsc.mx",
    "ymontoya@abcsc.mx",
    "jperez@abcsc.mx",
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
    """Muestra lista numerada y retorna la opción elegida con validación."""
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
    """
    Guarda todos los registros de la sesión en INVENTARIO{fecha}.xlsx.
    Si ya existe (misma fecha), concatena sin borrar lo previo.
    """
    df_nuevo = pd.DataFrame(registros)

    if os.path.exists(ARCHIVO_SESION):
        df_existente = pd.read_excel(ARCHIVO_SESION)
        df_final = pd.concat([df_existente, df_nuevo], ignore_index=True)
    else:
        df_final = df_nuevo

    df_final.to_excel(ARCHIVO_SESION, index=False)
    print(f"💾 Sesión guardada  → {ARCHIVO_SESION}  ({len(df_final)} registros totales)")
    return ARCHIVO_SESION


# ──────────────────────────────────────────────
# SUBIDA A SQL SERVER
# ──────────────────────────────────────────────
def conectar_sql():
    try:
        c = SQL_CONFIG
        conn_str = (
            f"DRIVER={c['driver']};SERVER={c['server']};"
            f"DATABASE={c['database']};UID={c['username']};PWD={c['password']}"
        )
        params = urllib.parse.quote_plus(conn_str)
        engine = create_engine(f"mssql+pyodbc:///?odbc_connect={params}", fast_executemany=True)
        with engine.connect() as conn:
            conn.execute(text("SELECT 1"))
        print(f"✅ SQL Server conectado → {c['server']} / {c['database']}")
        return engine
    except Exception as e:
        print(f"❌ No se pudo conectar a SQL Server: {e}")
        return None


def subir_a_sql(registros: list):
    print("\n🔄 Subiendo datos a SQL Server...")
    engine = conectar_sql()
    if not engine:
        print("   ⚠️  Los datos NO se subieron (fallo de conexión).")
        return
    try:
        df = pd.DataFrame(registros)
        if "." in TABLA_SQL:
            schema, tabla = TABLA_SQL.split(".", 1)
        else:
            schema, tabla = None, TABLA_SQL

        df.to_sql(name=tabla, con=engine, schema=schema, if_exists="append", index=False)
        print(f"   ✅ {len(df)} registro(s) insertado(s) en [{TABLA_SQL}]")
    except Exception as e:
        print(f"   ❌ Error al insertar en SQL: {e}")
    finally:
        engine.dispose()


# ──────────────────────────────────────────────
# ENVÍO DE CORREO
# ──────────────────────────────────────────────
def enviar_correo(archivos_adjuntos: list, total_registros: int):
    print("\n📧 Enviando reporte por correo...")
    hoy = datetime.now().strftime("%d/%m/%Y %H:%M")

    msg = EmailMessage()
    msg["Subject"] = f"Inventario QR — Sesión {datetime.now().strftime('%d/%m/%Y')}"
    msg["From"]    = EMAIL_USER
    msg["To"]      = ", ".join(DESTINATARIOS)
    msg.set_content(
        f"Buen día,\n\n"
        f"Se registraron {total_registros} artículo(s) en la sesión del {hoy}.\n"
        f"Se adjuntan los códigos QR/barras generados y el archivo {ARCHIVO_SESION}.\n\n"
        f"— Sistema BI · Inventario"
    )

    for archivo in archivos_adjuntos:
        if not os.path.exists(archivo):
            continue
        with open(archivo, "rb") as f:
            datos = f.read()
        nombre = os.path.basename(archivo)
        ext    = nombre.rsplit(".", 1)[-1].lower()
        if ext in ("png", "jpg", "jpeg"):
            msg.add_attachment(datos, maintype="image", subtype=ext, filename=nombre)
        elif ext == "xlsx":
            msg.add_attachment(datos, maintype="application",
                               subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               filename=nombre)
        else:
            msg.add_attachment(datos, maintype="application", subtype="octet-stream", filename=nombre)

    try:
        with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT) as smtp:
            smtp.login(EMAIL_USER, EMAIL_PASS)
            smtp.send_message(msg)
        print(f"✅ Correo enviado a: {', '.join(DESTINATARIOS)}")
    except smtplib.SMTPAuthenticationError:
        print("❌ Error de autenticación SMTP.")
    except Exception as e:
        print(f"❌ Error de envío: {e}")


# ──────────────────────────────────────────────
# FLUJO PRINCIPAL
# ──────────────────────────────────────────────
def ejecutar_sistema():
    archivos_sesion  = []   # imágenes QR y barras generadas
    registros_sesion = []   # datos completos de cada registro

    while True:
        try:
            xls      = pd.ExcelFile(archivo_excel)
            df_mat   = pd.read_excel(xls, "Materiales")
            empresas = pd.read_excel(xls, "Empresa")["Empresa"].dropna().tolist()
            areas    = pd.read_excel(xls, "Areas")["Area"].dropna().tolist()

            print("\n" + "=" * 45)
            print(f"  REGISTRO #{len(registros_sesion) + 1}  |  {datetime.now().strftime('%d/%m/%Y %H:%M')}")
            print("=" * 45)

            # ── Captura ──────────────────────────────────
            nombre_mat  = input("\nNombre del Material : ").strip()
            letras_auto = extraer_tres_letras(nombre_mat)
            print(f"✨ Abreviatura generada: {letras_auto}")

            desc         = input("Descripción         : ").strip()
            emp_sel      = seleccionar_opcion(empresas, "🏢 Empresa:")
            edificio_sel = seleccionar_opcion(EDIFICIOS, "🏗️  Edificio:")
            area_sel     = seleccionar_opcion(areas,    "📍 Área:")

            # ── Código único ──────────────────────────────
            correlativo = len(df_mat) + len(registros_sesion) + 1
            codigo_id   = f"{letras_auto}{correlativo:04d}"
            timestamp   = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            print(f"\n🔑 Código asignado: {codigo_id}")

            # ── 1. QR ─────────────────────────────────────
            ruta_qr  = os.path.join(carpeta_salida, f"qr_{codigo_id}.png")
            datos_qr = (
                f"ID:{codigo_id}\nMAT:{nombre_mat}\nDESC:{desc}\n"
                f"ORG:{emp_sel}\nEDIF:{edificio_sel}\nLOC:{area_sel}"
            )
            qrcode.make(datos_qr).save(ruta_qr)
            archivos_sesion.append(ruta_qr)
            print(f"   📷 QR generado      → {ruta_qr}")

            # ── 2. Código de barras ───────────────────────
            bn_class          = barcode.get_barcode_class("code128")
            ruta_barras_base  = os.path.join(carpeta_salida, f"barras_{codigo_id}")
            ruta_barras_final = bn_class(codigo_id, writer=ImageWriter()).save(ruta_barras_base)
            archivos_sesion.append(ruta_barras_final)
            print(f"   📊 Barras generadas → {ruta_barras_final}")

            # ── 3. Actualizar inventario.xlsx (catálogo) ──
            nueva_fila_cat = {
                "Material": nombre_mat, "Abreviatura": letras_auto,
                "Descripción": desc,    "Empresa": emp_sel,
                "Edificio": edificio_sel, "Area": area_sel,
            }
            with pd.ExcelWriter(archivo_excel, mode="a", if_sheet_exists="overlay", engine="openpyxl") as writer:
                pd.concat([df_mat, pd.DataFrame([nueva_fila_cat])], ignore_index=True
                          ).to_excel(writer, sheet_name="Materiales", index=False)

            # ── 4. Acumular para sesión ───────────────────
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

            print(f"\n✅ Registrado exitosamente: {codigo_id}")

            continuar = input("\n¿Deseas registrar otro? (S/N): ").strip().upper()
            if continuar == "N":
                break

        except (IndexError, ValueError) as e:
            print(f"⚠️  Entrada inválida: {e}. Intenta de nuevo.")
        except KeyboardInterrupt:
            print("\n⚠️  Interrumpido por el usuario.")
            break
        except Exception as e:
            print(f"❌ Error inesperado: {e}")
            break

    # ── CIERRE ────────────────────────────────────────────
    if not registros_sesion:
        print("\nℹ️  No se capturó ningún registro. Fin.")
        return

    print(f"\n{'='*45}")
    print(f"  CERRANDO SESIÓN — {len(registros_sesion)} registro(s)")
    print(f"{'='*45}")

    archivo_sesion_path = guardar_sesion_excel(registros_sesion)
    subir_a_sql(registros_sesion)
    enviar_correo(archivos_sesion + [archivo_sesion_path], len(registros_sesion))

    print("\n🏁 ¡Proceso terminado!")


if __name__ == "__main__":
    ejecutar_sistema()
