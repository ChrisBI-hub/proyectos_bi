"""
sync_bd.py v2.3 — DIAGNÓSTICO MEJORADO
=======================================
Versión con mejor detección de archivos Excel y diagnóstico completo.

Cambios principales:
  ✅ Busca específicamente .xlsx (no imágenes)
  ✅ Lista TODOS los adjuntos encontrados
  ✅ Mejor logging para diagnosticar problemas
  ✅ Tolerante si hay múltiples adjuntos
  ✅ UPSERT para evitar duplicados

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

LIMITES_COLUMNAS = {
    "Tipo": 100, "Nombre": 500, "No_Serie": 200, "Marca": 200, "Modelo": 200,
    "Empresa": 200, "Edificio": 100, "Area": 200, "Ubicacion_En_Edificio": 250,
    "Estado": 50, "Tipo_Sensor": 100, "Resolucion_Pantalla": 50,
    "Sistema_Operativo": 100, "Codigo_Barras": 255, "Codigo_ID": 100,
    "Host": 100, "Procesador": 150, "RAM": 50, "Capacidad_Disco": 50,
    "Tipo_Disco": 20, "Codigo_QR": 500, "Codigo_Barras_CPU": 255,
    "Host_CPU": 100, "SO": 100, "Office": 100, "Antivirus": 100,
    "Lector_PDF": 100, "ERP": 100, "Otro_1": 100, "Otro_2": 100, "Otro_3": 100,
}


# ──────────────────────────────────────────────
# HELPERS
# ──────────────────────────────────────────────
def truncar_strings(df: pd.DataFrame) -> pd.DataFrame:
    """Recorta valores de texto a los límites definidos."""
    df = df.copy()
    for col in df.columns:
        if col in LIMITES_COLUMNAS and df[col].dtype == object:
            limite = LIMITES_COLUMNAS[col]
            mask = df[col].notna() & (df[col].str.len() > limite)
            if mask.any():
                print(f"   ✂️  Columna '{col}': {mask.sum()} valor(es) recortado(s)")
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
    """Obtiene IDs únicos de una tabla para deduplicación."""
    try:
        with engine.connect() as conn:
            schema, tabla_nombre = (tabla.split(".", 1) if "." in tabla else (None, tabla))
            tabla_completa = f"[{schema}].[{tabla_nombre}]" if schema else f"[{tabla_nombre}]"
            resultado = conn.execute(text(f"SELECT {columna_id} FROM {tabla_completa}"))
            return set(row[0] for row in resultado if row[0])
    except Exception as e:
        print(f"   ⚠️  No se pudo obtener IDs de {tabla}: {e}")
        return set()


def hacer_upsert_sql(df: pd.DataFrame, tabla: str, schema: str, engine, 
                     columna_pk: str = None) -> bool:
    """Hace UPSERT: actualiza si existe, inserta si es nuevo."""
    if df.empty:
        return True

    try:
        print(f"   📤 Procesando {len(df)} fila(s) para [{tabla}]...")
        
        schema_tabla = f"[{schema}].[{tabla.split('.')[-1]}]" if schema else f"[{tabla.split('.')[-1]}]"
        
        if not columna_pk or columna_pk not in df.columns:
            df = truncar_strings(df)
            for col in df.columns:
                if "timestamp" in col.lower() or "fecha" in col.lower():
                    try:
                        df[col] = pd.to_datetime(df[col], errors="coerce")
                    except:
                        pass
            
            with engine.begin() as conn:
                tabla_nombre = tabla.split(".")[-1]
                df.to_sql(name=tabla_nombre, con=conn, schema=schema, if_exists="append", index=False)
            print(f"   ✅ {len(df)} fila(s) insertada(s) en [{tabla}]")
            return True

        ids_existentes = obtener_ids_existentes(engine, tabla, columna_pk)
        
        df_nuevos = df[~df[columna_pk].isin(ids_existentes)].copy()
        df_existentes = df[df[columna_pk].isin(ids_existentes)].copy()

        contador_actualizado = 0
        contador_insertado = 0

        # ACTUALIZAR
        if not df_existentes.empty:
            df_existentes = truncar_strings(df_existentes)
            for col in df_existentes.columns:
                if "timestamp" in col.lower() or "fecha" in col.lower():
                    try:
                        df_existentes[col] = pd.to_datetime(df_existentes[col], errors="coerce")
                    except:
                        pass
            
            try:
                with engine.begin() as conn:
                    for idx, row in df_existentes.iterrows():
                        set_clause = ", ".join([f"[{col}] = :{col}" for col in df_existentes.columns if col != columna_pk])
                        sql_update = f"UPDATE {schema_tabla} SET {set_clause} WHERE [{columna_pk}] = :{columna_pk}"
                        params = {col: row[col] for col in df_existentes.columns}
                        conn.execute(text(sql_update), params)
                        contador_actualizado += 1
                print(f"   🔄 {contador_actualizado} fila(s) actualizada(s)")
            except Exception as e:
                print(f"   ⚠️  Error actualizando: {e}")

        # INSERTAR
        if not df_nuevos.empty:
            df_nuevos = truncar_strings(df_nuevos)
            for col in df_nuevos.columns:
                if "timestamp" in col.lower() or "fecha" in col.lower():
                    try:
                        df_nuevos[col] = pd.to_datetime(df_nuevos[col], errors="coerce")
                    except:
                        pass
            
            try:
                with engine.begin() as conn:
                    tabla_nombre = tabla.split(".")[-1]
                    df_nuevos.to_sql(name=tabla_nombre, con=conn, schema=schema, if_exists="append", index=False)
                contador_insertado = len(df_nuevos)
                print(f"   ➕ {contador_insertado} fila(s) insertada(s)")
            except Exception as e:
                print(f"   ⚠️  Error insertando: {e}")

        if contador_actualizado == 0 and contador_insertado == 0:
            print(f"   ℹ️  {tabla}: {len(df)} registro(s) ya existían")
        
        return True

    except Exception as e:
        print(f"   ❌ Error en UPSERT: {e}")
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

    if "cpu" in hojas and not hojas["cpu"].empty:
        print("\n   🖥️  Procesando CPU...")
        df_cpu = hojas["cpu"].copy()
        exito &= hacer_upsert_sql(df_cpu, TABLAS_SQL["CPU"]["tabla_principal"],
                                   schema="Inventario", engine=engine,
                                   columna_pk="Codigo_Barras_CPU")

    hosts_en_bd = obtener_ids_existentes(engine, TABLAS_SQL["CPU"]["tabla_principal"], "Host")

    if "software" in hojas and not hojas["software"].empty:
        print("\n   📦 Procesando Software...")
        df_sw = hojas["software"].copy()
        if "Host_CPU" in df_sw.columns:
            sin_padre = df_sw[~df_sw["Host_CPU"].isin(hosts_en_bd)]
            if not sin_padre.empty:
                print(f"   ⚠️  {len(sin_padre)} fila(s) omitida(s) — Host_CPU no existe")
            df_sw = df_sw[df_sw["Host_CPU"].isin(hosts_en_bd)]
        if not df_sw.empty:
            exito &= hacer_upsert_sql(df_sw, TABLAS_SQL["CPU"]["tabla_software"],
                                       schema="Inventario", engine=engine,
                                       columna_pk="Host_CPU")

    if "perifericos" in hojas and not hojas["perifericos"].empty:
        print("\n   🖱️  Procesando Periféricos...")
        df_per = hojas["perifericos"].copy()
        if "Host_CPU" in df_per.columns:
            sin_padre = df_per[~df_per["Host_CPU"].isin(hosts_en_bd)]
            if not sin_padre.empty:
                print(f"   ⚠️  {len(sin_padre)} fila(s) omitida(s) — Host_CPU no existe")
            df_per = df_per[df_per["Host_CPU"].isin(hosts_en_bd)]
        if not df_per.empty:
            exito &= hacer_upsert_sql(df_per, TABLAS_SQL["CPU"]["tabla_perifericos"],
                                       schema="Inventario", engine=engine,
                                       columna_pk="Codigo_ID")

    if "relaciones" in hojas and not hojas["relaciones"].empty:
        print("\n   🔗 Procesando Relaciones...")
        exito &= hacer_upsert_sql(hojas["relaciones"], TABLAS_SQL["CPU"]["tabla_relaciones"],
                                   schema="Inventario", engine=engine, columna_pk=None)

    return exito


def insertar_otros_equipos(hojas: dict, engine) -> bool:
    if "otros" in hojas and not hojas["otros"].empty:
        print("\n   📦 Procesando Otros Equipamientos...")
        return hacer_upsert_sql(hojas["otros"], TABLAS_SQL["OTROS"]["tabla_principal"],
                                 schema="Inventario", engine=engine, columna_pk="Codigo_ID")
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
        return False
    finally:
        engine.dispose()


# ──────────────────────────────────────────────
# CORREO IMAP — CON MEJOR DIAGNÓSTICO
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


def extraer_adjuntos_de_correo(mail: imaplib.IMAP4_SSL, uid: bytes) -> dict:
    """
    Extrae TODOS los adjuntos y los clasifica.
    Retorna {'xlsx': [bytes, filename], 'otros': [list of (filename, type)]}
    """
    _, data = mail.fetch(uid, "(RFC822)")
    msg = email.message_from_bytes(data[0][1])
    
    xlsx_encontrado = None
    otros_adjuntos = []

    for part in msg.walk():
        filename = part.get_filename()
        if not filename:
            continue
        
        # Buscar .xlsx
        if filename.lower().endswith((".xlsx", ".xls")):
            print(f"   📎 Excel encontrado: {filename}")
            try:
                xlsx_encontrado = (part.get_payload(decode=True), filename)
            except Exception as e:
                print(f"   ⚠️  Error decodificando {filename}: {e}")
        else:
            # Catalogar otros adjuntos
            content_type = part.get_content_type()
            otros_adjuntos.append((filename, content_type))
            print(f"   📄 Otro adjunto: {filename} ({content_type})")

    return {
        'xlsx': xlsx_encontrado,
        'otros': otros_adjuntos
    }


def marcar_como_leido(mail: imaplib.IMAP4_SSL, uid: bytes):
    mail.store(uid, "+FLAGS", "\\Seen")


# ──────────────────────────────────────────────
# PROCESAMIENTO DE UN CORREO
# ──────────────────────────────────────────────
def procesar_correo(mail: imaplib.IMAP4_SSL, uid: bytes):
    print(f"\n   📩 Procesando correo UID {uid.decode()}...")

    adjuntos = extraer_adjuntos_de_correo(mail, uid)
    
    if not adjuntos['xlsx']:
        print("   ⚠️  No se encontró archivo .xlsx — correo ignorado.")
        if adjuntos['otros']:
            print(f"      (Adjuntos encontrados: {', '.join([a[0] for a in adjuntos['otros']])})")
        marcar_como_leido(mail, uid)
        return

    xlsx_bytes, filename = adjuntos['xlsx']

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
        print(f"   ⚠️  Inserción falló — se reintentará próximamente.")


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
    parser = argparse.ArgumentParser(description="sync_bd v2.3 — Sincronizador Inventario → SQL Server")
    parser.add_argument("--once", action="store_true", help="Revisar una sola vez y salir")
    args = parser.parse_args()

    print("=" * 60)
    print("  SYNC_BD v2.3 — Sincronizador Inventario → SQL Server")
    print("  ✅ Diagnóstico mejorado + UPSERT")
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