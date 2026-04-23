"""
sync_bd.py v2.1  (Kali Linux — tu PC con acceso al servidor)
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

Uso:
    python3 sync_bd.py              # corre indefinidamente
    python3 sync_bd.py --once       # revisa una sola vez y sale
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

IMAP_SERVER  = "imap.gmail.com"
IMAP_PORT    = 993
EMAIL_USER   = "reportes.bi@abcsc.mx"
EMAIL_PASS   = "jwvjdrvmprzrwzxy"

ASUNTO_TRIGGER   = "ACTUALIZACION_BASE_DE_DATOS"
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
        "tabla_principal":   "Inventario.CPU",
        "tabla_software":    "Inventario.CPU_Software",
        "tabla_perifericos": "Inventario.CPU_Perifericos",
        "tabla_relaciones":  "Inventario.CPU_Relaciones",
    },
    "OTROS": {
        "tabla_principal": "Inventario.Otros_Equipos",
    }
}

INTERVALO_SEGUNDOS = 60

# Límites de caracteres por columna (deben coincidir con creacion_de_vistas.sql v2.1)
# Sirven como red de seguridad si el SQL aún no fue actualizado.
LIMITES_COLUMNAS = {
    "Tipo":              100,
    "Nombre":            500,
    "No_Serie":          200,
    "Marca":             200,
    "Modelo":            200,
    "Empresa":           200,
    "Edificio":          100,
    "Area":              200,
    "Ubicacion_En_Edificio": 250,
    "Estado":             50,
    "Tipo_Sensor":       100,
    "Resolucion_Pantalla": 50,
    "Sistema_Operativo": 100,
    "Codigo_Barras":     255,
    "Codigo_ID":         100,
    "Host":              100,
    "Procesador":        150,
    "RAM":                50,
    "Capacidad_Disco":    50,
    "Tipo_Disco":         20,
    "Codigo_QR":         500,
    "Codigo_Barras_CPU": 255,
    "Host_CPU":          100,
    "SO":                100,
    "Office":            100,
    "Antivirus":         100,
    "Lector_PDF":        100,
    "ERP":               100,
    "Otro_1":            100,
    "Otro_2":            100,
    "Otro_3":            100,
}


# ──────────────────────────────────────────────
# HELPERS
# ──────────────────────────────────────────────
def truncar_strings(df: pd.DataFrame) -> pd.DataFrame:
    """
    Recorta valores de texto a los límites definidos en LIMITES_COLUMNAS.
    Evita errores de truncación en SQL Server si alguna columna aún tiene
    un tamaño antiguo o si llega un valor inesperadamente largo.
    """
    df = df.copy()
    for col in df.columns:
        if col in LIMITES_COLUMNAS and df[col].dtype == object:
            limite = LIMITES_COLUMNAS[col]
            mask = df[col].notna() & (df[col].str.len() > limite)
            if mask.any():
                print(f"   ✂️  Columna '{col}': {mask.sum()} valor(es) recortado(s) a {limite} chars.")
            df[col] = df[col].where(df[col].isna(), df[col].str.slice(0, limite))
    return df


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
        with engine.connect() as conn:
            conn.execute(text("SELECT 1"))
        print("   ✅ Conexión a SQL Server exitosa")
        return engine
    except Exception as e:
        print(f"   ❌ Error conectando a SQL Server: {e}")
        return None


def obtener_ids_existentes(engine, tabla: str, columna_id: str) -> set:
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
    if df.empty:
        return True

    try:
        # Filtrar duplicados
        if columna_id and columna_id in df.columns:
            ids_existentes = obtener_ids_existentes(engine, tabla, columna_id)
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
                except Exception:
                    pass

        # Recortar strings al límite de cada columna (red de seguridad)
        df_nuevos = truncar_strings(df_nuevos)

        print(f"   📤 Insertando {len(df_nuevos)} fila(s) en [{tabla}]...")
        tabla_nombre = tabla.split(".")[-1]

        with engine.begin() as conn:
            df_nuevos.to_sql(
                name=tabla_nombre,
                con=conn,
                schema=schema,
                if_exists="append",
                index=False
            )
        print(f"   ✅ {len(df_nuevos)} fila(s) insertada(s) correctamente en [{tabla}]")
        return True

    except Exception as e:
        print(f"   ❌ Error al insertar en {tabla}: {e}")
        import traceback
        traceback.print_exc()
        return False


# ──────────────────────────────────────────────
# PROCESAMIENTO DE HOJAS EXCEL
# ──────────────────────────────────────────────
def procesar_archivo_excel(xlsx_bytes: bytes) -> dict:
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
    exito = True

    # ── 1. CPU principal ─────────────────────────────────────────────────────
    if "cpu" in hojas and not hojas["cpu"].empty:
        print("\n   🖥️  Procesando CPU...")
        df_cpu = hojas["cpu"].drop(columns=["Tipo"], errors="ignore")
        exito &= insertar_dataframe_sql(
            df_cpu, TABLAS_SQL["CPU"]["tabla_principal"],
            schema="Inventario", engine=engine,
            columna_id="Codigo_Barras_CPU"
        )

    # ── 2. Obtener TODOS los hosts que existen en Inventario.CPU ─────────────
    # Se consulta DESPUÉS de la inserción anterior para incluir los recién
    # agregados. Las tablas hijas solo se insertan si su Host_CPU ya existe
    # en la tabla padre, evitando la violación de FK.
    hosts_en_bd = obtener_ids_existentes(engine, TABLAS_SQL["CPU"]["tabla_principal"], "Host")

    # ── 3. Software ──────────────────────────────────────────────────────────
    if "software" in hojas and not hojas["software"].empty:
        print("\n   📦 Procesando Software...")
        df_sw = hojas["software"].copy()

        # Filtrar filas cuyo Host_CPU no existe en Inventario.CPU
        if "Host_CPU" in df_sw.columns:
            sin_padre = df_sw[~df_sw["Host_CPU"].isin(hosts_en_bd)]
            if not sin_padre.empty:
                print(f"   ⚠️  Software: {len(sin_padre)} fila(s) omitida(s) — "
                      f"Host_CPU no existe en CPU: {sin_padre['Host_CPU'].unique().tolist()}")
            df_sw = df_sw[df_sw["Host_CPU"].isin(hosts_en_bd)]

        if not df_sw.empty:
            # Para Software usamos Host_CPU como clave de dedup (1 fila por host)
            exito &= insertar_dataframe_sql(
                df_sw, TABLAS_SQL["CPU"]["tabla_software"],
                schema="Inventario", engine=engine,
                columna_id="Host_CPU"
            )

    # ── 4. Periféricos ───────────────────────────────────────────────────────
    if "perifericos" in hojas and not hojas["perifericos"].empty:
        print("\n   🖱️  Procesando Periféricos...")
        df_per = hojas["perifericos"].copy()

        # Filtrar filas cuyo Host_CPU no existe en Inventario.CPU
        if "Host_CPU" in df_per.columns:
            sin_padre = df_per[~df_per["Host_CPU"].isin(hosts_en_bd)]
            if not sin_padre.empty:
                print(f"   ⚠️  Periféricos: {len(sin_padre)} fila(s) omitida(s) — "
                      f"Host_CPU no existe en CPU: {sin_padre['Host_CPU'].unique().tolist()}")
            df_per = df_per[df_per["Host_CPU"].isin(hosts_en_bd)]

        if not df_per.empty:
            exito &= insertar_dataframe_sql(
                df_per, TABLAS_SQL["CPU"]["tabla_perifericos"],
                schema="Inventario", engine=engine,
                columna_id="Codigo_ID"
            )

    # ── 5. Relaciones ────────────────────────────────────────────────────────
    if "relaciones" in hojas and not hojas["relaciones"].empty:
        print("\n   🔗 Procesando Relaciones...")
        exito &= insertar_dataframe_sql(
            hojas["relaciones"], TABLAS_SQL["CPU"]["tabla_relaciones"],
            schema="Inventario", engine=engine,
            columna_id=None
        )

    return exito


def insertar_otros_equipos(hojas: dict, engine) -> bool:
    if "otros" in hojas and not hojas["otros"].empty:
        print("\n   📦 Procesando Otros Equipamientos...")
        return insertar_dataframe_sql(
            hojas["otros"], TABLAS_SQL["OTROS"]["tabla_principal"],
            schema="Inventario", engine=engine,
            columna_id="Codigo_ID"
        )
    return True


def insertar_en_sql(xlsx_bytes: bytes) -> bool:
    print(f"   🔄 Conectando a SQL Server [{SQL_CONFIG['server']}]...")
    engine = conectar_sql()
    if not engine:
        return False

    try:
        hojas = procesar_archivo_excel(xlsx_bytes)

        if not hojas:
            print("   ❌ No se pudieron leer las hojas del Excel")
            return False

        exito = True

        if "cpu" in hojas:
            exito &= insertar_estructura_cpu(hojas, engine)

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
# CORREO IMAP
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
    return data[0].split()


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

    try:
        xls = pd.ExcelFile(io.BytesIO(xlsx_bytes))
        print(f"   📊 Archivo: {filename}  |  Hojas: {xls.sheet_names}")
    except Exception as e:
        print(f"   ❌ Error leyendo xlsx: {e}")
        marcar_como_leido(mail, uid)
        return

    os.makedirs(CARPETA_RESPALDO, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    ruta_respaldo = os.path.join(CARPETA_RESPALDO, f"{ts}_{filename}")
    with open(ruta_respaldo, "wb") as f:
        f.write(xlsx_bytes)
    print(f"   💾 Respaldo guardado → {ruta_respaldo}")

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
    while True:
        revisar_una_vez()
        print(f"   ⏳ Próxima revisión en {INTERVALO_SEGUNDOS}s...\n")
        time.sleep(INTERVALO_SEGUNDOS)


def main():
    parser = argparse.ArgumentParser(description="sync_bd v2.1 — Sincronizador Inventario → SQL Server")
    parser.add_argument("--once", action="store_true", help="Revisar una sola vez y salir")
    args = parser.parse_args()

    print("=" * 60)
    print("  SYNC_BD v2.1 — Sincronizador Inventario → SQL Server")
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