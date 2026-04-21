"""
sync_bd.py v2.0  (Kali Linux — tu PC con acceso al servidor)
=============================================================
Versión mejorada para procesar la nueva estructura de Excel de Codigoqr_v2.py

Corre en bucle cada 60 segundos. Revisa el buzón IMAP buscando
correos con el asunto exacto "ACTUALIZACION_BASE_DE_DATOS".
Por cada correo nuevo encontrado:
  1. Descarga el adjunto INVENTARIO*.xlsx
  2. Procesa múltiples hojas (CPU, Software, Periféricos, Otros)
  3. Inserta los registros en SQL Server con estructura adecuada
  4. Marca el correo como leído para no procesarlo dos veces
  5. Guarda el xlsx localmente como respaldo

CAMBIOS v2.0:
  - Soporte para múltiples hojas de Excel
  - Procesamiento inteligente según tipo de equipamiento
  - Mejor validación de datos
  - Logging mejorado

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
import urllib.parse
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

# Tablas SQL por tipo de inventario
TABLAS_SQL = {
    "CPU": {
        "tabla_principal": "Inventario.CPU",
        "tabla_software": "Inventario.CPU_Software",
        "tabla_perifericos": "Inventario.CPU_Perifericos",
        "tabla_relaciones": "Inventario.CPU_Relaciones",
    },
    "OTROS": {
        "tabla_principal": "Inventario.Otros_Equipos",
    }
}

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


def obtener_ids_existentes(engine, tabla: str, columna_id: str = "CodigoID") -> set:
    """
    Obtiene los IDs que ya existen en una tabla SQL.
    Útil para evitar duplicados.
    """
    try:
        with engine.connect() as conn:
            schema, tabla_nombre = (tabla.split(".", 1) if "." in tabla else (None, tabla))
            tabla_completa = f"[{schema}].[{tabla_nombre}]" if schema else f"[{tabla_nombre}]"
            resultado = conn.execute(text(f"SELECT {columna_id} FROM {tabla_completa}"))
            return set(row[0] for row in resultado if row[0])
    except Exception as e:
        print(f"   ⚠️  No se pudo obtener IDs de {tabla}: {e}")
        return set()


def insertar_dataframe_sql(df: pd.DataFrame, tabla: str, schema: str = None, 
                          engine=None, columna_id: str = "CodigoID") -> bool:
    """
    Inserta un DataFrame en SQL Server, evitando duplicados.
    """
    if df.empty:
        return True

    try:
        # Obtener IDs existentes si hay columna de ID
        ids_existentes = set()
        if columna_id in df.columns:
            ids_existentes = obtener_ids_existentes(engine, tabla, columna_id)

            # Filtrar solo filas nuevas
            df_nuevos = df[~df[columna_id].isin(ids_existentes)].copy()

            if df_nuevos.empty:
                print(f"   ℹ️  {tabla}: {len(df)} registro(s) ya existían — ninguno insertado.")
                return True

            omitidos = len(df) - len(df_nuevos)
            if omitidos > 0:
                print(f"   ⏭️  {tabla}: {omitidos} registro(s) ya existían — omitidos.")
        else:
            df_nuevos = df.copy()

        # Convertir timestamps
        for col in df_nuevos.columns:
            if "timestamp" in col.lower() or "fecha" in col.lower():
                try:
                    df_nuevos[col] = pd.to_datetime(df_nuevos[col], errors="coerce")
                except:
                    pass

        print(f"   📤 Insertando {len(df_nuevos)} fila(s) en [{tabla}]...")
        tabla_nombre = tabla.split(".")[-1]
        df_nuevos.to_sql(
            name=tabla_nombre,
            con=engine,
            schema=schema,
            if_exists="append",
            index=False
        )
        print(f"   ✅ {len(df_nuevos)} fila(s) insertada(s) correctamente en [{tabla}]")
        return True

    except Exception as e:
        print(f"   ❌ Error al insertar en {tabla}: {e}")
        return False


# ──────────────────────────────────────────────
# PROCESAMIENTO DE HOJAS EXCEL
# ──────────────────────────────────────────────

def procesar_archivo_excel(xlsx_bytes: bytes) -> dict:
    """
    Lee el archivo Excel y extrae DataFrames de todas las hojas.
    Retorna dict con estructura: {nombre_hoja: dataframe}
    """
    try:
        xls = pd.ExcelFile(io.BytesIO(xlsx_bytes))
        hojas = {}
        
        print(f"   📊 Hojas encontradas: {xls.sheet_names}")
        
        for sheet_name in xls.sheet_names:
            df = pd.read_excel(io.BytesIO(xlsx_bytes), sheet_name=sheet_name)
            hojas[sheet_name.lower()] = df
            print(f"      └─ {sheet_name}: {len(df)} fila(s), {len(df.columns)} columna(s)")
        
        return hojas
    except Exception as e:
        print(f"   ❌ Error leyendo Excel: {e}")
        return {}


def insertar_estructura_cpu(hojas: dict, engine) -> bool:
    """
    Procesa la estructura de CPU con múltiples hojas:
    - CPU (datos principales)
    - Software (software instalado)
    - Perifericos (periféricos asociados)
    - Relaciones (vinculación de códigos)
    """
    exito = True
    
    # Procesar tabla principal de CPU
    if "cpu" in hojas and not hojas["cpu"].empty:
        print("\n   🖥️  Procesando CPU...")
        df_cpu = hojas["cpu"].drop(columns=["Tipo"], errors="ignore")
        exito &= insertar_dataframe_sql(
            df_cpu,
            TABLAS_SQL["CPU"]["tabla_principal"],
            schema="Inventario",
            engine=engine,
            columna_id="Codigo_Barras_CPU"
        )
    
    # Procesar Software
    if "software" in hojas and not hojas["software"].empty:
        print("\n   📦 Procesando Software...")
        exito &= insertar_dataframe_sql(
            hojas["software"],
            TABLAS_SQL["CPU"]["tabla_software"],
            schema="Inventario",
            engine=engine,
            columna_id="Host_CPU"
        )
    
    # Procesar Periféricos
    if "perifericos" in hojas and not hojas["perifericos"].empty:
        print("\n   🖱️  Procesando Periféricos...")
        exito &= insertar_dataframe_sql(
            hojas["perifericos"],
            TABLAS_SQL["CPU"]["tabla_perifericos"],
            schema="Inventario",
            engine=engine,
            columna_id="Codigo_ID"
        )
    
    # Procesar Relaciones
    if "relaciones" in hojas and not hojas["relaciones"].empty:
        print("\n   🔗 Procesando Relaciones...")
        exito &= insertar_dataframe_sql(
            hojas["relaciones"],
            TABLAS_SQL["CPU"]["tabla_relaciones"],
            schema="Inventario",
            engine=engine,
            columna_id=None  # Sin deduplicación para relaciones
        )
    
    return exito


def insertar_otros_equipos(hojas: dict, engine) -> bool:
    """
    Procesa equipos simples (sillas, mesas, etc.)
    """
    if "otros" in hojas and not hojas["otros"].empty:
        print("\n   📦 Procesando Otros Equipamientos...")
        return insertar_dataframe_sql(
            hojas["otros"],
            TABLAS_SQL["OTROS"]["tabla_principal"],
            schema="Inventario",
            engine=engine,
            columna_id="Codigo_ID"
        )
    return True


def insertar_en_sql(xlsx_bytes: bytes) -> bool:
    """
    Coordina todo el proceso de inserción.
    """
    print(f"   🔄 Conectando a SQL Server [{SQL_CONFIG['server']}]...")
    engine = conectar_sql()
    if not engine:
        return False

    try:
        # Leer Excel
        hojas = procesar_archivo_excel(xlsx_bytes)
        
        if not hojas:
            print("   ❌ No se pudieron leer las hojas del Excel")
            return False

        # Procesar según el tipo de contenido
        exito = True
        
        # Si tiene tabla CPU, es estructura CPU
        if "cpu" in hojas:
            exito &= insertar_estructura_cpu(hojas, engine)
        
        # Si tiene tabla Otros, procesar equipos simples
        if "otros" in hojas:
            exito &= insertar_otros_equipos(hojas, engine)

        hojas_reconocidas = {"cpu", "software", "perifericos", "relaciones", "otros"}
        if not any(h in hojas_reconocidas for h in hojas):
            print("   ⚠️  Excel sin hojas reconocidas (CPU u Otros)")

        return exito

    except Exception as e:
        print(f"   ❌ Error en procesamiento: {e}")
        import traceback
        traceback.print_exc()
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
    """Busca correos no leídos con el asunto trigger."""
    mail.select("INBOX")
    _, data = mail.search(None, f'(UNSEEN SUBJECT "{ASUNTO_TRIGGER}")')
    ids = data[0].split()
    return ids


def extraer_xlsx_de_correo(mail: imaplib.IMAP4_SSL, uid: bytes):
    """Extrae el adjunto XLSX del correo."""
    _, data = mail.fetch(uid, "(RFC822)")
    msg = email.message_from_bytes(data[0][1])

    for part in msg.walk():
        filename = part.get_filename()
        if filename and filename.lower().endswith(".xlsx"):
            print(f"   📎 Adjunto encontrado: {filename}")
            return part.get_payload(decode=True), filename

    return None, None


def marcar_como_leido(mail: imaplib.IMAP4_SSL, uid: bytes):
    """Marca un correo como leído."""
    mail.store(uid, "+FLAGS", "\\Seen")


# ──────────────────────────────────────────────
# PROCESAMIENTO DE UN CORREO
# ──────────────────────────────────────────────
def procesar_correo(mail: imaplib.IMAP4_SSL, uid: bytes):
    """Procesa un correo individual."""
    print(f"\n   📩 Procesando correo UID {uid.decode()}...")

    xlsx_bytes, filename = extraer_xlsx_de_correo(mail, uid)

    if not xlsx_bytes:
        print("   ⚠️  No se encontró adjunto .xlsx — correo ignorado.")
        marcar_como_leido(mail, uid)
        return

    # Leer y validar Excel
    try:
        xls = pd.ExcelFile(io.BytesIO(xlsx_bytes))
        print(f"   📊 Archivo: {filename}  |  Hojas: {xls.sheet_names}")
    except Exception as e:
        print(f"   ❌ Error leyendo xlsx: {e}")
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
    exito = insertar_en_sql(xlsx_bytes)

    if exito:
        marcar_como_leido(mail, uid)
        print(f"   ✉️  Correo marcado como leído.")
    else:
        print(f"   ⚠️  Inserción falló — se reintentará en el próximo ciclo.")


# ──────────────────────────────────────────────
# CICLO PRINCIPAL
# ──────────────────────────────────────────────
def revisar_una_vez():
    """Ejecuta una revisión única del buzón."""
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
    """Entry point del script."""
    parser = argparse.ArgumentParser(description="sync_bd — Sincronizador de inventario a SQL Server v2.0")
    parser.add_argument("--once", action="store_true", help="Revisar una sola vez y salir")
    args = parser.parse_args()

    print("=" * 60)
    print("  SYNC_BD v2.0 — Sincronizador Inventario → SQL Server")
    print(f"  Trigger  : {ASUNTO_TRIGGER}")
    print(f"  Tablas   : CPU, Software, Periféricos, Otros")
    print(f"  Intervalo: {INTERVALO_SEGUNDOS}s")
    print("=" * 60)

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