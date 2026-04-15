"""
sync_bd.py  (Kali Linux — tu PC con acceso al servidor)
========================================================
Corre en bucle cada 60 segundos. Revisa el buzón IMAP buscando
correos con el asunto exacto "ACTUALIZACION_BASE_DE_DATOS".
Por cada correo nuevo encontrado:
  1. Descarga el adjunto INVENTARIO*.xlsx
  2. Inserta los registros en SQL Server
  3. Marca el correo como leído para no procesarlo dos veces
  4. Guarda el xlsx localmente como respaldo

Uso:
    python3 sync_bd.py              # corre indefinidamente
    python3 sync_bd.py --once       # revisa una sola vez y sale

Instalación de dependencias (ejecutar una vez en Kali):
    pip install pandas openpyxl sqlalchemy pyodbc
"""

import os
import sys
import time
import imaplib
import email
import urllib.parse          # ← FIX: debe ser urllib.parse explícito, no solo urllib
import argparse
import io
from datetime import datetime

import pandas as pd
from sqlalchemy import create_engine, text

# ──────────────────────────────────────────────
# CONFIGURACIÓN
# ──────────────────────────────────────────────

# ── Correo IMAP ───────────────────────────────
IMAP_SERVER  = "imap.gmail.com"
IMAP_PORT    = 993
EMAIL_USER   = "reportes.bi@abcsc.mx"
EMAIL_PASS   = "jwvjdrvmprzrwzxy"

# Asunto exacto que dispara la sincronización
ASUNTO_TRIGGER = "ACTUALIZACION_BASE_DE_DATOS"

# Carpeta local donde guardar los xlsx recibidos (respaldo)
CARPETA_RESPALDO = "./respaldo_inventario"

# ── SQL Server ────────────────────────────────
SQL_CONFIG = {
    "server":   "150.1.1.152",
    "database": "BI",
    "username": "ConsultaBD",
    "password": "5D$bc#kM&5W2T8J40?s%",
    "driver":   "{ODBC Driver 17 for SQL Server}",
}
TABLA_SQL  = "Inventario.Codigos_QR"

# ── Intervalo de revisión ─────────────────────
INTERVALO_SEGUNDOS = 60


# ──────────────────────────────────────────────
# CONEXIÓN SQL
# ──────────────────────────────────────────────
def conectar_sql():
    try:
        c = SQL_CONFIG
        conn_str = (
            f"DRIVER={c['driver']};SERVER={c['server']};"
            f"DATABASE={c['database']};UID={c['username']};PWD={c['password']}"
        )
        # FIX: urllib.parse.quote_plus en lugar de urllib.parse.quote_plus
        params = urllib.parse.quote_plus(conn_str)
        engine = create_engine(
            f"mssql+pyodbc:///?odbc_connect={params}",
            fast_executemany=True,
        )
        # Verificar conexión
        with engine.connect() as conn:
            conn.execute(text("SELECT 1"))
        print("   ✅ Conexión a SQL Server exitosa")
        return engine
    except Exception as e:
        print(f"   ❌ Error conectando a SQL Server: {e}")
        return None


def insertar_en_sql(df: pd.DataFrame) -> bool:
    print(f"   🔄 Conectando a SQL Server [{SQL_CONFIG['server']}]...")
    engine = conectar_sql()
    if not engine:
        return False
    try:
        schema, tabla = (TABLA_SQL.split(".", 1) if "." in TABLA_SQL else (None, TABLA_SQL))
        tabla_completa = f"[{schema}].[{tabla}]" if schema else f"[{tabla}]"

        # Consultar los CodigoID que ya existen en la tabla
        with engine.connect() as conn:
            resultado = conn.execute(text(f"SELECT CodigoID FROM {tabla_completa}"))
            ids_existentes = set(row[0] for row in resultado)

        # Filtrar solo filas nuevas
        df_nuevos = df[~df["CodigoID"].isin(ids_existentes)].copy()

        if df_nuevos.empty:
            print(f"   ℹ️  Sin registros nuevos — todos los {len(df)} del xlsx ya estaban en la tabla.")
            return True

        omitidos = len(df) - len(df_nuevos)
        if omitidos > 0:
            print(f"   ⏭️  {omitidos} registro(s) ya existían en la tabla — omitidos.")

        # Convertir FechaHora a datetime para que SQL Server lo acepte correctamente
        if "FechaHora" in df_nuevos.columns:
            df_nuevos["FechaHora"] = pd.to_datetime(df_nuevos["FechaHora"], errors="coerce")
            nulos = df_nuevos["FechaHora"].isna().sum()
            if nulos > 0:
                print(f"   ⚠️  {nulos} valor(es) de FechaHora no pudieron convertirse y se dejarán como NULL.")

        print(f"   📤 Insertando {len(df_nuevos)} fila(s) nuevas en [{TABLA_SQL}]...")
        df_nuevos.to_sql(name=tabla, con=engine, schema=schema, if_exists="append", index=False)
        print(f"   ✅ {len(df_nuevos)} fila(s) insertada(s) correctamente en [{TABLA_SQL}]")
        return True

    except Exception as e:
        print(f"   ❌ Error al insertar en SQL: {e}")
        return False
    finally:
        engine.dispose()


# ──────────────────────────────────────────────
# LECTURA DE CORREO (IMAP)
# ──────────────────────────────────────────────
def conectar_imap():
    try:
        mail = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
        mail.login(EMAIL_USER, EMAIL_PASS)
        print("   ✅ Conexión IMAP exitosa")
        return mail
    except Exception as e:
        print(f"   ❌ Error IMAP: {e}")
        return None


def buscar_correos_nuevos(mail: imaplib.IMAP4_SSL) -> list:
    mail.select("INBOX")
    _, data = mail.search(None, f'(UNSEEN SUBJECT "{ASUNTO_TRIGGER}")')
    ids = data[0].split()
    return ids


def extraer_xlsx_de_correo(mail: imaplib.IMAP4_SSL, uid: bytes):
    _, data = mail.fetch(uid, "(RFC822)")
    msg = email.message_from_bytes(data[0][1])

    for part in msg.walk():
        filename = part.get_filename()
        if filename and filename.lower().endswith(".xlsx"):
            print(f"   📎 Adjunto encontrado: {filename}")
            return part.get_payload(decode=True), filename

    return None, None


def marcar_como_leido(mail: imaplib.IMAP4_SSL, uid: bytes):
    mail.store(uid, "+FLAGS", "\\Seen")


# ──────────────────────────────────────────────
# PROCESAMIENTO DE UN CORREO
# ──────────────────────────────────────────────
def procesar_correo(mail: imaplib.IMAP4_SSL, uid: bytes):
    print(f"\n   📩 Procesando correo UID {uid.decode()}...")

    xlsx_bytes, filename = extraer_xlsx_de_correo(mail, uid)

    if not xlsx_bytes:
        print("   ⚠️  No se encontró adjunto .xlsx — correo ignorado.")
        marcar_como_leido(mail, uid)
        return

    # Leer el xlsx desde memoria
    try:
        df = pd.read_excel(io.BytesIO(xlsx_bytes))
        print(f"   📊 Archivo: {filename}  |  {len(df)} fila(s)  |  Columnas: {list(df.columns)}")
    except Exception as e:
        print(f"   ❌ Error leyendo xlsx: {e}")
        marcar_como_leido(mail, uid)
        return

    if df.empty:
        print("   ⚠️  El archivo no contiene datos.")
        marcar_como_leido(mail, uid)
        return

    # Guardar respaldo local
    os.makedirs(CARPETA_RESPALDO, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    ruta_respaldo = os.path.join(CARPETA_RESPALDO, f"{ts}_{filename}")
    with open(ruta_respaldo, "wb") as f:
        f.write(xlsx_bytes)
    print(f"   💾 Respaldo guardado → {ruta_respaldo}")

    # Insertar en SQL Server
    exito = insertar_en_sql(df)

    if exito:
        marcar_como_leido(mail, uid)
        print(f"   ✉️  Correo marcado como leído.")
    else:
        print(f"   ⚠️  Inserción fallida — se reintentará en el próximo ciclo.")


# ──────────────────────────────────────────────
# CICLO PRINCIPAL
# ──────────────────────────────────────────────
def revisar_una_vez():
    print(f"[{datetime.now().strftime('%H:%M:%S')}] 🔍 Revisando bandeja...")
    mail = conectar_imap()
    if not mail:
        print("   Sin conexión IMAP. Se reintentará.")
        return

    try:
        uids = buscar_correos_nuevos(mail)
        if not uids:
            print("   Sin correos nuevos con el asunto trigger.")
        else:
            print(f"   📬 {len(uids)} correo(s) encontrado(s).")
            for uid in uids:
                procesar_correo(mail, uid)
    finally:
        try:
            mail.logout()
        except Exception:
            pass


def main_loop():
    """Bucle infinito — llamado por main.py."""
    while True:
        revisar_una_vez()
        print(f"   ⏳ Próxima revisión en {INTERVALO_SEGUNDOS}s...\n")
        time.sleep(INTERVALO_SEGUNDOS)


def main():
    parser = argparse.ArgumentParser(description="sync_bd — Sincronizador de inventario a SQL Server")
    parser.add_argument("--once", action="store_true", help="Revisar una sola vez y salir")
    args = parser.parse_args()

    print("=" * 55)
    print("  SYNC_BD — Sincronizador Inventario → SQL Server")
    print(f"  Trigger  : {ASUNTO_TRIGGER}")
    print(f"  Tabla SQL: {TABLA_SQL}")
    print(f"  Intervalo: {INTERVALO_SEGUNDOS}s")
    print("=" * 55)

    if args.once:
        revisar_una_vez()
        return

    print("▶️  Corriendo en modo continuo. Ctrl+C para detener.\n")
    try:
        main_loop()
    except KeyboardInterrupt:
        print("\n🛑 Detenido por el usuario.")


if __name__ == "__main__":
    main()