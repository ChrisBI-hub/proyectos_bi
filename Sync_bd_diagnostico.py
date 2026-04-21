"""
sync_bd.py v2.2  (Kali Linux) — VERSIÓN CON DIAGNÓSTICO
========================================================
Versión mejorada con logging detallado para identificar problemas.

Si los datos no se suben a SQL Server, este script te mostrará exactamente
qué está pasando con mensajes más descriptivos.

Uso:
    python3 sync_bd.py              # corre indefinidamente
    python3 sync_bd.py --once       # revisa una sola vez y sale
    python3 sync_bd.py --debug      # modo debug (muy verbose)
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
from sqlalchemy import create_engine, text, inspect

# ──────────────────────────────────────────────
# CONFIGURACIÓN
# ──────────────────────────────────────────────

DEBUG = False  # Se activa con --debug

IMAP_SERVER  = "imap.gmail.com"
IMAP_PORT    = 993
EMAIL_USER   = "reportes.bi@abcsc.mx"
EMAIL_PASS   = "jwvjdrvmprzrwzxy"

ASUNTO_TRIGGER = "ACTUALIZACION_BASE_DE_DATOS"

CARPETA_RESPALDO = "./respaldo_inventario"

SQL_CONFIG = {
    "server":   "150.1.1.152",
    "database": "BI",
    "username": "ConsultaBD",
    "password": "5D$bc#kM&5W2T8J40?s%",
    "driver":   "{ODBC Driver 17 for SQL Server}",
}

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

INTERVALO_SEGUNDOS = 60


# ──────────────────────────────────────────────
# LOGGING CON DEBUG
# ──────────────────────────────────────────────

def dprint(mensaje: str):
    """Imprime solo si DEBUG está activo."""
    if DEBUG:
        print(f"   [DEBUG] {mensaje}")


def log_info(mensaje: str):
    """Información normal."""
    print(f"   ℹ️  {mensaje}")


def log_exito(mensaje: str):
    """Mensaje de éxito."""
    print(f"   ✅ {mensaje}")


def log_error(mensaje: str):
    """Mensaje de error."""
    print(f"   ❌ {mensaje}")


def log_advertencia(mensaje: str):
    """Mensaje de advertencia."""
    print(f"   ⚠️  {mensaje}")


# ──────────────────────────────────────────────
# CONEXIÓN SQL CON DIAGNÓSTICO
# ──────────────────────────────────────────────

def conectar_sql():
    """Conecta a SQL Server con diagnóstico detallado."""
    try:
        c = SQL_CONFIG
        conn_str = (
            f"DRIVER={c['driver']};SERVER={c['server']};"
            f"DATABASE={c['database']};UID={c['username']};PWD={c['password']}"
        )
        
        dprint(f"Connection string: {conn_str[:50]}...")
        
        params = urllib.parse.quote_plus(conn_str)
        engine = create_engine(
            f"mssql+pyodbc:///?odbc_connect={params}",
            fast_executemany=True,
        )
        
        # Verificar conexión
        with engine.connect() as conn:
            resultado = conn.execute(text("SELECT 1"))
            dprint(f"Test query result: {resultado.fetchone()}")
        
        log_exito("Conexión a SQL Server exitosa")
        return engine
        
    except Exception as e:
        log_error(f"Conectando a SQL Server: {str(e)}")
        import traceback
        dprint(traceback.format_exc())
        return None


def verificar_tablas(engine):
    """Verifica si las tablas existen en SQL Server."""
    try:
        inspector = inspect(engine)
        tablas_existentes = inspector.get_table_names(schema="Inventario")
        
        dprint(f"Tablas encontradas en schema 'Inventario': {tablas_existentes}")
        
        if not tablas_existentes:
            log_error("No se encontraron tablas en schema 'Inventario'")
            print("   → Ejecuta el script creacion_de_vistas.sql primero")
            return False
        
        # Verificar tablas específicas
        tablas_requeridas = ["CPU", "CPU_Software", "CPU_Perifericos", "CPU_Relaciones", "Otros_Equipos"]
        faltantes = [t for t in tablas_requeridas if t not in tablas_existentes]
        
        if faltantes:
            log_error(f"Faltan tablas: {faltantes}")
            return False
        
        log_exito("Todas las tablas requeridas existen")
        return True
        
    except Exception as e:
        log_error(f"Verificando tablas: {e}")
        return False


def obtener_ids_existentes(engine, tabla: str, columna_id: str = "CodigoID") -> set:
    """Obtiene IDs existentes con diagnóstico."""
    try:
        with engine.connect() as conn:
            schema, tabla_nombre = (tabla.split(".", 1) if "." in tabla else (None, tabla))
            tabla_completa = f"[{schema}].[{tabla_nombre}]" if schema else f"[{tabla_nombre}]"
            
            dprint(f"Consultando tabla: {tabla_completa}, columna: {columna_id}")
            
            resultado = conn.execute(text(f"SELECT {columna_id} FROM {tabla_completa}"))
            ids = set(row[0] for row in resultado if row[0])
            
            dprint(f"IDs existentes en {tabla}: {len(ids)}")
            return ids
            
    except Exception as e:
        dprint(f"Error obteniendo IDs de {tabla}: {e}")
        return set()


def insertar_dataframe_sql(df: pd.DataFrame, tabla: str, schema: str = None, 
                          engine=None, columna_id: str = "CodigoID") -> bool:
    """Inserta DataFrame en SQL con diagnóstico detallado."""
    
    if df.empty:
        log_advertencia(f"{tabla}: DataFrame vacío, nada que insertar")
        return True

    try:
        dprint(f"\n=== INSERTANDO EN {tabla} ===")
        dprint(f"Filas totales: {len(df)}")
        dprint(f"Columnas: {list(df.columns)}")
        
        # Obtener IDs existentes
        ids_existentes = set()
        if columna_id in df.columns:
            dprint(f"Verificando duplicados con columna: {columna_id}")
            ids_existentes = obtener_ids_existentes(engine, tabla, columna_id)
            
            df_nuevos = df[~df[columna_id].isin(ids_existentes)].copy()
            
            if df_nuevos.empty:
                log_advertencia(f"{tabla}: {len(df)} ya existían - ninguno insertado")
                return True
            
            omitidos = len(df) - len(df_nuevos)
            if omitidos > 0:
                dprint(f"{tabla}: {omitidos} registro(s) ya existían")
        else:
            dprint(f"Columna '{columna_id}' no encontrada en DataFrame")
            df_nuevos = df.copy()

        # Convertir timestamps
        for col in df_nuevos.columns:
            if "timestamp" in col.lower() or "fecha" in col.lower():
                try:
                    df_nuevos[col] = pd.to_datetime(df_nuevos[col], errors="coerce")
                    dprint(f"Convertida columna {col} a datetime")
                except:
                    pass

        # Insertar
        dprint(f"\nInsertando {len(df_nuevos)} fila(s) en {tabla}...")
        tabla_nombre = tabla.split(".")[-1]
        
        dprint(f"Parámetros to_sql: name='{tabla_nombre}', schema='{schema}', if_exists='append'")
        dprint(f"Primeras filas a insertar:")
        dprint(f"{df_nuevos.head().to_string()}")
        
        df_nuevos.to_sql(
            name=tabla_nombre,
            con=engine,
            schema=schema,
            if_exists="append",
            index=False,
            method='multi'  # Más rápido para muchas filas
        )
        
        log_exito(f"{tabla}: {len(df_nuevos)} fila(s) insertada(s)")
        return True

    except Exception as e:
        log_error(f"Insertando en {tabla}: {str(e)}")
        dprint(f"Traceback completo:")
        import traceback
        dprint(traceback.format_exc())
        return False


# ──────────────────────────────────────────────
# PROCESAMIENTO DE EXCEL
# ──────────────────────────────────────────────

def procesar_archivo_excel(xlsx_bytes: bytes) -> dict:
    """Lee Excel con diagnóstico detallado."""
    try:
        xls = pd.ExcelFile(io.BytesIO(xlsx_bytes))
        hojas = {}
        
        dprint(f"Hojas en archivo: {xls.sheet_names}")
        print(f"   📊 Hojas encontradas: {xls.sheet_names}")
        
        for sheet_name in xls.sheet_names:
            df = pd.read_excel(io.BytesIO(xlsx_bytes), sheet_name=sheet_name)
            hojas[sheet_name.lower()] = df
            
            dprint(f"\n--- Hoja: {sheet_name} ---")
            dprint(f"Filas: {len(df)}, Columnas: {len(df.columns)}")
            dprint(f"Nombres de columnas: {list(df.columns)}")
            dprint(f"Datos de muestra:\n{df.head(2).to_string()}")
            
            print(f"      └─ {sheet_name}: {len(df)} fila(s), {len(df.columns)} columna(s)")
        
        return hojas
        
    except Exception as e:
        log_error(f"Leyendo Excel: {e}")
        import traceback
        dprint(traceback.format_exc())
        return {}


def insertar_estructura_cpu(hojas: dict, engine) -> bool:
    """Procesa CPU con diagnóstico."""
    exito = True
    
    # CPU principal
    if "cpu" in hojas and not hojas["cpu"].empty:
        print("\n   🖥️  Procesando CPU...")
        exito &= insertar_dataframe_sql(
            hojas["cpu"],
            TABLAS_SQL["CPU"]["tabla_principal"],
            schema="Inventario",
            engine=engine,
            columna_id="Codigo_Barras_CPU"
        )
    else:
        dprint("Hoja 'cpu' no encontrada o vacía")
    
    # Software
    if "software" in hojas and not hojas["software"].empty:
        print("\n   📦 Procesando Software...")
        exito &= insertar_dataframe_sql(
            hojas["software"],
            TABLAS_SQL["CPU"]["tabla_software"],
            schema="Inventario",
            engine=engine,
            columna_id="Host_CPU"
        )
    else:
        dprint("Hoja 'software' no encontrada o vacía")
    
    # Periféricos
    if "perifericos" in hojas and not hojas["perifericos"].empty:
        print("\n   🖱️  Procesando Periféricos...")
        exito &= insertar_dataframe_sql(
            hojas["perifericos"],
            TABLAS_SQL["CPU"]["tabla_perifericos"],
            schema="Inventario",
            engine=engine,
            columna_id="Codigo_ID"
        )
    else:
        dprint("Hoja 'perifericos' no encontrada o vacía")
    
    # Relaciones
    if "relaciones" in hojas and not hojas["relaciones"].empty:
        print("\n   🔗 Procesando Relaciones...")
        exito &= insertar_dataframe_sql(
            hojas["relaciones"],
            TABLAS_SQL["CPU"]["tabla_relaciones"],
            schema="Inventario",
            engine=engine,
            columna_id=None
        )
    else:
        dprint("Hoja 'relaciones' no encontrada o vacía")
    
    return exito


def insertar_otros_equipos(hojas: dict, engine) -> bool:
    """Procesa otros equipos con diagnóstico."""
    if "otros" in hojas and not hojas["otros"].empty:
        print("\n   📦 Procesando Otros Equipamientos...")
        return insertar_dataframe_sql(
            hojas["otros"],
            TABLAS_SQL["OTROS"]["tabla_principal"],
            schema="Inventario",
            engine=engine,
            columna_id="Codigo_ID"
        )
    else:
        dprint("Hoja 'otros' no encontrada o vacía")
    return True


def insertar_en_sql(xlsx_bytes: bytes, engine) -> bool:
    """Coordina inserción con diagnóstico."""
    print(f"   🔄 Procesando archivo Excel...")
    
    # Leer Excel
    hojas = procesar_archivo_excel(xlsx_bytes)
    
    if not hojas:
        log_error("No se pudieron leer las hojas del Excel")
        return False

    exito = True
    
    # Procesar según contenido
    if "cpu" in hojas:
        exito &= insertar_estructura_cpu(hojas, engine)
    
    if "otros" in hojas:
        exito &= insertar_otros_equipos(hojas, engine)

    if not any(["cpu" in h or "otros" in h for h in hojas]):
        log_advertencia("Excel sin hojas reconocidas (CPU u Otros)")

    return exito


# ──────────────────────────────────────────────
# CORREO IMAP
# ──────────────────────────────────────────────

def conectar_imap():
    try:
        mail = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
        mail.login(EMAIL_USER, EMAIL_PASS)
        return mail
    except Exception as e:
        log_error(f"IMAP: {e}")
        return None


def buscar_correos_nuevos(mail: imaplib.IMAP4_SSL) -> list:
    """Devuelve UIDs de correos no leídos."""
    mail.select("INBOX")
    _, data = mail.search(None, f'(UNSEEN SUBJECT "{ASUNTO_TRIGGER}")')
    ids = data[0].split()
    return ids


def extraer_xlsx_de_correo(mail: imaplib.IMAP4_SSL, uid: bytes):
    """Extrae adjunto XLSX del correo."""
    _, data = mail.fetch(uid, "(RFC822)")
    msg = email.message_from_bytes(data[0][1])

    for part in msg.walk():
        filename = part.get_filename()
        if filename and filename.lower().endswith(".xlsx"):
            dprint(f"Adjunto encontrado: {filename}")
            return part.get_payload(decode=True), filename

    return None, None


def marcar_como_leido(mail: imaplib.IMAP4_SSL, uid: bytes):
    """Marca correo como leído."""
    mail.store(uid, "+FLAGS", "\\Seen")


# ──────────────────────────────────────────────
# PROCESAMIENTO DE CORREO
# ──────────────────────────────────────────────

def procesar_correo(mail: imaplib.IMAP4_SSL, uid: bytes, engine):
    """Procesa un correo individual."""
    print(f"\n   📩 Procesando correo UID {uid.decode()}...")

    xlsx_bytes, filename = extraer_xlsx_de_correo(mail, uid)

    if not xlsx_bytes:
        log_advertencia("Sin adjunto .xlsx - correo ignorado")
        marcar_como_leido(mail, uid)
        return

    # Validar Excel
    try:
        xls = pd.ExcelFile(io.BytesIO(xlsx_bytes))
        dprint(f"Hojas en archivo: {xls.sheet_names}")
        print(f"   📊 Archivo: {filename} | Hojas: {xls.sheet_names}")
    except Exception as e:
        log_error(f"Leyendo xlsx: {e}")
        marcar_como_leido(mail, uid)
        return

    # Guardar respaldo
    os.makedirs(CARPETA_RESPALDO, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    ruta_respaldo = os.path.join(CARPETA_RESPALDO, f"{ts}_{filename}")
    with open(ruta_respaldo, "wb") as f:
        f.write(xlsx_bytes)
    dprint(f"Respaldo guardado: {ruta_respaldo}")

    # Insertar en SQL
    exito = insertar_en_sql(xlsx_bytes, engine)

    if exito:
        marcar_como_leido(mail, uid)
        log_exito("Correo marcado como leído")
    else:
        log_advertencia("Inserción falló - se reintentará en el próximo ciclo")


# ──────────────────────────────────────────────
# CICLO PRINCIPAL
# ──────────────────────────────────────────────

def revisar_una_vez(engine):
    """Ejecuta una revisión."""
    print(f"[{datetime.now().strftime('%H:%M:%S')}] 🔍 Revisando bandeja...")
    
    mail = conectar_imap()
    if not mail:
        log_advertencia("Sin conexión IMAP")
        return

    try:
        uids = buscar_correos_nuevos(mail)
        if not uids:
            print("   Sin correos nuevos con el asunto trigger.")
        else:
            print(f"   📬 {len(uids)} correo(s) encontrado(s).")
            for uid in uids:
                procesar_correo(mail, uid, engine)
    finally:
        try:
            mail.logout()
        except:
            pass


def main_loop(engine):
    """Bucle infinito."""
    while True:
        revisar_una_vez(engine)
        print(f"   ⏳ Próxima revisión en {INTERVALO_SEGUNDOS}s...\n")
        time.sleep(INTERVALO_SEGUNDOS)


def main():
    global DEBUG
    
    parser = argparse.ArgumentParser(description="sync_bd v2.2 — Con diagnóstico")
    parser.add_argument("--once", action="store_true", help="Revisar una sola vez")
    parser.add_argument("--debug", action="store_true", help="Modo debug (verbose)")
    parser.add_argument("--verify", action="store_true", help="Verificar tablas y salir")
    args = parser.parse_args()
    
    DEBUG = args.debug

    print("=" * 60)
    print("  SYNC_BD v2.2 — Sincronizador Inventario → SQL Server")
    print(f"  Trigger  : {ASUNTO_TRIGGER}")
    print(f"  Debug    : {'ACTIVO' if DEBUG else 'Inactivo'}")
    print("=" * 60)

    # Conectar a SQL
    engine = conectar_sql()
    if not engine:
        print("\n❌ No se pudo conectar a SQL Server")
        if DEBUG:
            print("   Ejecuta con --debug para más detalles")
        return 1

    # Verificar tablas si se solicita
    if args.verify:
        print("\n🔍 Verificando tablas en SQL Server...")
        if verificar_tablas(engine):
            print("✅ Todas las tablas existen correctamente")
            return 0
        else:
            print("❌ Hay problemas con las tablas")
            return 1

    # Ejecutar
    if args.once:
        revisar_una_vez(engine)
    else:
        print("▶️  Corriendo en modo continuo. Ctrl+C para detener.\n")
        try:
            main_loop(engine)
        except KeyboardInterrupt:
            print("\n🛑 Detenido por el usuario.")
    
    return 0


if __name__ == "__main__":
    sys.exit(main())