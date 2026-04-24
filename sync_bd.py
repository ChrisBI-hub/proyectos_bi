"""
sync_bd.py
==========
Sincroniza archivos Excel de inventario recibidos por correo hacia SQL Server.

El archivo recibido puede venir sin IDs finales; este script normaliza las hojas,
asigna los identificadores del lado receptor y hace UPSERT en la base de datos.
"""

import argparse
import email
import imaplib
import io
import os
import time
import urllib.parse
from datetime import datetime

import pandas as pd
from sqlalchemy import create_engine, text

from asignar_ids_inventario import preparar_hojas_para_sql


IMAP_SERVER = "imap.gmail.com"
IMAP_PORT = 993
EMAIL_USER = "reportes.bi@abcsc.mx"
EMAIL_PASS = "jwvjdrvmprzrwzxy"

ASUNTO_TRIGGER = "ACTUALIZACION_BASE_DE_DATOS"
CARPETA_RESPALDO = "./respaldo_inventario"

SQL_CONFIG = {
    "server": "150.1.1.152",
    "database": "BI",
    "username": "ConsultaBD",
    "password": "5D$bc#kM&5W2T8J40?s%",
    "driver": "{ODBC Driver 17 for SQL Server}",
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
    },
}

INTERVALO_SEGUNDOS = 60

LIMITES_COLUMNAS = {
    "Tipo": 100,
    "Nombre": 500,
    "No_Serie": 200,
    "Marca": 200,
    "Modelo": 200,
    "Empresa": 200,
    "Edificio": 100,
    "Area": 200,
    "Ubicacion_En_Edificio": 250,
    "Estado": 50,
    "Tipo_Sensor": 100,
    "Resolucion_Pantalla": 50,
    "Sistema_Operativo": 100,
    "Codigo_Barras": 255,
    "Codigo_ID": 100,
    "Periferico_UID": 100,
    "Host": 100,
    "Procesador": 150,
    "RAM": 50,
    "Capacidad_Disco": 50,
    "Tipo_Disco": 20,
    "Codigo_QR": 500,
    "Codigo_Barras_CPU": 255,
    "Host_CPU": 100,
    "SO": 100,
    "Office": 100,
    "Antivirus": 100,
    "Lector_PDF": 100,
    "ERP": 100,
    "Otro_1": 100,
    "Otro_2": 100,
    "Otro_3": 100,
}


def truncar_strings(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    for col in df.columns:
        if col in LIMITES_COLUMNAS and df[col].dtype == object:
            limite = LIMITES_COLUMNAS[col]
            mask = df[col].notna() & (df[col].str.len() > limite)
            if mask.any():
                print(f"   ✂️  Columna '{col}': {mask.sum()} valor(es) recortado(s)")
            df[col] = df[col].where(df[col].isna(), df[col].str.slice(0, limite))
    return df


def preparar_fechas(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    for col in df.columns:
        if "timestamp" in col.lower() or "fecha" in col.lower():
            try:
                df[col] = pd.to_datetime(df[col], errors="coerce")
            except Exception:
                pass
    return df


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
            resultado = conn.execute(text(f"SELECT [{columna_id}] FROM {tabla_completa}"))
            return {row[0] for row in resultado if row[0] is not None}
    except Exception as e:
        print(f"   ⚠️  No se pudo obtener IDs de {tabla}: {e}")
        return set()


def obtener_columnas_tabla(engine, schema: str, tabla: str) -> list[str]:
    tabla_nombre = tabla.split(".")[-1]
    with engine.connect() as conn:
        rows = conn.execute(
            text(
                """
                SELECT COLUMN_NAME
                FROM INFORMATION_SCHEMA.COLUMNS
                WHERE TABLE_SCHEMA = :schema AND TABLE_NAME = :tabla
                ORDER BY ORDINAL_POSITION
                """
            ),
            {"schema": schema, "tabla": tabla_nombre},
        ).fetchall()
    return [row[0] for row in rows]


def alinear_dataframe_a_tabla(df: pd.DataFrame, engine, schema: str, tabla: str) -> pd.DataFrame:
    if df.empty:
        return df
    columnas_tabla = obtener_columnas_tabla(engine, schema, tabla)
    columnas_validas = [col for col in df.columns if col in columnas_tabla]
    columnas_ignoradas = [col for col in df.columns if col not in columnas_tabla]
    if columnas_ignoradas:
        print(f"   ⚠️  Columnas ignoradas en [{tabla}]: {', '.join(columnas_ignoradas)}")
    return df[columnas_validas].copy()


def hacer_upsert_sql(df: pd.DataFrame, tabla: str, schema: str, engine, columna_pk: str | None = None) -> bool:
    if df.empty:
        return True

    try:
        tabla_corta = tabla.split(".")[-1]
        schema_tabla = f"[{schema}].[{tabla_corta}]" if schema else f"[{tabla_corta}]"
        df = alinear_dataframe_a_tabla(df, engine, schema, tabla)
        df = truncar_strings(df)
        df = preparar_fechas(df)

        if df.empty:
            print(f"   ℹ️  [{tabla}] sin columnas válidas para insertar.")
            return True

        print(f"   📤 Procesando {len(df)} fila(s) para [{tabla}]...")

        if not columna_pk or columna_pk not in df.columns:
            with engine.begin() as conn:
                df.to_sql(name=tabla_corta, con=conn, schema=schema, if_exists="append", index=False)
            print(f"   ✅ {len(df)} fila(s) insertada(s) en [{tabla}]")
            return True

        ids_existentes = obtener_ids_existentes(engine, tabla, columna_pk)
        df_existentes = df[df[columna_pk].isin(ids_existentes)].copy()
        df_nuevos = df[~df[columna_pk].isin(ids_existentes)].copy()

        actualizados = 0
        insertados = 0

        if not df_existentes.empty:
            set_clause = ", ".join(f"[{col}] = :{col}" for col in df_existentes.columns if col != columna_pk)
            if set_clause:
                sql_update = f"UPDATE {schema_tabla} SET {set_clause} WHERE [{columna_pk}] = :{columna_pk}"
                try:
                    with engine.begin() as conn:
                        for _, row in df_existentes.iterrows():
                            conn.execute(text(sql_update), row.to_dict())
                            actualizados += 1
                    print(f"   🔄 {actualizados} fila(s) actualizada(s)")
                except Exception as e:
                    print(f"   ❌ Error actualizando [{tabla}]: {e}")
                    return False

        if not df_nuevos.empty:
            try:
                with engine.begin() as conn:
                    df_nuevos.to_sql(name=tabla_corta, con=conn, schema=schema, if_exists="append", index=False)
                insertados = len(df_nuevos)
                print(f"   ➕ {insertados} fila(s) insertada(s)")
            except Exception as e:
                print(f"   ❌ Error insertando [{tabla}]: {e}")
                return False

        if actualizados == 0 and insertados == 0:
            print(f"   ℹ️  {tabla}: {len(df)} registro(s) ya existían")

        return True
    except Exception as e:
        print(f"   ❌ Error en UPSERT de [{tabla}]: {e}")
        return False


def procesar_archivo_excel(xlsx_bytes: bytes) -> dict[str, pd.DataFrame]:
    try:
        xls = pd.ExcelFile(io.BytesIO(xlsx_bytes))
        hojas: dict[str, pd.DataFrame] = {}
        print(f"   📊 Hojas encontradas: {xls.sheet_names}")
        for sheet_name in xls.sheet_names:
            df = pd.read_excel(io.BytesIO(xlsx_bytes), sheet_name=sheet_name)
            hojas[sheet_name.lower()] = df
            print(f"      └─ {sheet_name}: {len(df)} fila(s), {len(df.columns)} columna(s)")
        return hojas
    except Exception as e:
        print(f"   ❌ Error leyendo Excel: {e}")
        return {}


def insertar_estructura_cpu(hojas: dict[str, pd.DataFrame], engine) -> bool:
    exito = True

    if "cpu" in hojas and not hojas["cpu"].empty:
        print("\n   🖥️  Procesando CPU...")
        exito &= hacer_upsert_sql(
            hojas["cpu"].copy(),
            TABLAS_SQL["CPU"]["tabla_principal"],
            schema="Inventario",
            engine=engine,
            columna_pk="Host",
        )

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
            exito &= hacer_upsert_sql(
                df_sw,
                TABLAS_SQL["CPU"]["tabla_software"],
                schema="Inventario",
                engine=engine,
                columna_pk="Host_CPU",
            )

    if "perifericos" in hojas and not hojas["perifericos"].empty:
        print("\n   🖱️  Procesando Periféricos...")
        df_per = hojas["perifericos"].copy()
        if "Host_CPU" in df_per.columns:
            sin_padre = df_per[~df_per["Host_CPU"].isin(hosts_en_bd)]
            if not sin_padre.empty:
                print(f"   ⚠️  {len(sin_padre)} fila(s) omitida(s) — Host_CPU no existe")
            df_per = df_per[df_per["Host_CPU"].isin(hosts_en_bd)]
        if not df_per.empty:
            exito &= hacer_upsert_sql(
                df_per,
                TABLAS_SQL["CPU"]["tabla_perifericos"],
                schema="Inventario",
                engine=engine,
                columna_pk="Periferico_UID",
            )

    if "relaciones" in hojas and not hojas["relaciones"].empty:
        print("\n   🔗 Procesando Relaciones...")
        exito &= hacer_upsert_sql(
            hojas["relaciones"].copy(),
            TABLAS_SQL["CPU"]["tabla_relaciones"],
            schema="Inventario",
            engine=engine,
            columna_pk=None,
        )

    return exito


def insertar_otros_equipos(hojas: dict[str, pd.DataFrame], engine) -> bool:
    if "otros" in hojas and not hojas["otros"].empty:
        print("\n   📦 Procesando Otros Equipamientos...")
        return hacer_upsert_sql(
            hojas["otros"].copy(),
            TABLAS_SQL["OTROS"]["tabla_principal"],
            schema="Inventario",
            engine=engine,
            columna_pk="Codigo_ID",
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

        hojas = preparar_hojas_para_sql(hojas, engine)

        exito = True
        if not hojas["cpu"].empty:
            exito &= insertar_estructura_cpu(hojas, engine)
        if not hojas["otros"].empty:
            exito &= insertar_otros_equipos(hojas, engine)

        if all(hojas[nombre].empty for nombre in ("cpu", "software", "perifericos", "relaciones", "otros")):
            print("   ⚠️  Excel sin hojas reconocidas con datos para sincronizar")

        return exito
    except Exception as e:
        print(f"   ❌ Error en procesamiento: {e}")
        return False
    finally:
        engine.dispose()


def conectar_imap():
    try:
        mail = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
        mail.login(EMAIL_USER, EMAIL_PASS)
        print("   ✅ Conexión IMAP exitosa")
        return mail
    except Exception as e:
        print(f"   ❌ Error IMAP: {e}")
        return None


def buscar_correos_nuevos(mail: imaplib.IMAP4_SSL) -> list[bytes]:
    mail.select("INBOX")
    _, data = mail.search(None, f'(UNSEEN SUBJECT "{ASUNTO_TRIGGER}")')
    return data[0].split()


def extraer_adjuntos_de_correo(mail: imaplib.IMAP4_SSL, uid: bytes) -> dict:
    _, data = mail.fetch(uid, "(RFC822)")
    msg = email.message_from_bytes(data[0][1])

    xlsx_encontrado = None
    otros_adjuntos = []

    for part in msg.walk():
        filename = part.get_filename()
        if not filename:
            continue

        if filename.lower().endswith((".xlsx", ".xls")):
            print(f"   📎 Excel encontrado: {filename}")
            try:
                xlsx_encontrado = (part.get_payload(decode=True), filename)
            except Exception as e:
                print(f"   ⚠️  Error decodificando {filename}: {e}")
        else:
            content_type = part.get_content_type()
            otros_adjuntos.append((filename, content_type))
            print(f"   📄 Otro adjunto: {filename} ({content_type})")

    return {"xlsx": xlsx_encontrado, "otros": otros_adjuntos}


def marcar_como_leido(mail: imaplib.IMAP4_SSL, uid: bytes):
    mail.store(uid, "+FLAGS", "\\Seen")


def procesar_correo(mail: imaplib.IMAP4_SSL, uid: bytes):
    print(f"\n   📨 Procesando correo UID {uid.decode()}...")
    adjuntos = extraer_adjuntos_de_correo(mail, uid)

    if not adjuntos["xlsx"]:
        print("   ⚠️  No se encontró archivo .xlsx — correo ignorado.")
        if adjuntos["otros"]:
            print(f"      (Adjuntos encontrados: {', '.join(a[0] for a in adjuntos['otros'])})")
        marcar_como_leido(mail, uid)
        return

    xlsx_bytes, filename = adjuntos["xlsx"]

    try:
        xls = pd.ExcelFile(io.BytesIO(xlsx_bytes))
        print(f"   📊 Archivo: {filename} | Hojas: {xls.sheet_names}")
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
        print("   ✉️  Correo marcado como leído.")
    else:
        print("   ⚠️  Inserción falló — se reintentará próximamente.")


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
    parser = argparse.ArgumentParser(description="sync_bd — Sincronizador Inventario → SQL Server")
    parser.add_argument("--once", action="store_true", help="Revisar una sola vez y salir")
    args = parser.parse_args()

    print("=" * 60)
    print("  SYNC_BD — Sincronizador Inventario → SQL Server")
    print("  ✅ IDs asignados del lado receptor + UPSERT")
    print(f"  Trigger  : {ASUNTO_TRIGGER}")
    print("  Tablas   : CPU, Software, Periféricos, Relaciones, Otros")
    print(f"  Intervalo: {INTERVALO_SEGUNDOS}s")
    print("=" * 60)

    if args.once:
        revisar_una_vez()
        return

    print("▶️  Corriendo en modo continuo. Ctrl+C para detener.\n")
    try:
        main_loop()
    except KeyboardInterrupt:
        print("\n🛑 Detenido por el usuario. revisar continuidad")


if __name__ == "__main__":
    main()