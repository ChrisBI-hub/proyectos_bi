"""
==============================================================
  migrar_inventario.py  –  Migración Excel → SQL Server
  Base de datos : BI  |  Schema : Inventario
  Tablas        : CPU  /  CPU_Software  /  CPU_Perifericos
  Requiere      : pip install pandas openpyxl pyodbc
==============================================================
"""

import sys
import socket
import getpass
import logging
from datetime import datetime
from pathlib import Path

import pandas as pd
import pyodbc




SQL_CONFIG = {
    "server":   "150.1.1.152",
    "database": "BI",
    "username": "ConsultaBD",
    "password": "5D$bc#kM&5W2T8J40?s%",
    "driver":   "{ODBC Driver 17 for SQL Server}",
}


# ──────────────────────────────────────────────────────────────
# 1.  CONFIGURACIÓN  ← edita solo esta sección
# ──────────────────────────────────────────────────────────────
EXCEL_PATH = r"/home/christian/Documentos/proyectos_bi/respaldo_inventario/20260423_173804_INVENTARIO20260423.xlsx"      # ruta al archivo Excel

SQL_SERVER  = SQL_CONFIG["server"]
SQL_DB      = SQL_CONFIG["database"]
SQL_USER    = SQL_CONFIG["username"]
SQL_PASS    = SQL_CONFIG["password"]
SQL_DRIVER  = SQL_CONFIG["driver"]

MODO_UPSERT = True   # True = MERGE (insertar+actualizar)  |  False = solo insertar nuevos
# ──────────────────────────────────────────────────────────────


# ── logging ───────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler(f"migracion_{datetime.now():%Y%m%d_%H%M%S}.log", encoding="utf-8"),
    ],
)
log = logging.getLogger(__name__)


# ══════════════════════════════════════════════════════════════
# 2.  CONEXIÓN
# ══════════════════════════════════════════════════════════════
def get_connection() -> pyodbc.Connection:
    # Nota: SQL_DRIVER ya incluye las llaves → "{ODBC Driver 17 for SQL Server}"
    conn_str = (
        f"DRIVER={SQL_DRIVER};"
        f"SERVER={SQL_SERVER};"
        f"DATABASE={SQL_DB};"
        f"UID={SQL_USER};"
        f"PWD={{{SQL_PASS}}};"          # 👈 doble llave para encerrar la contraseña
        "Encrypt=no;"
        "TrustServerCertificate=yes;"
    )
    try:
        conn = pyodbc.connect(conn_str, timeout=10)
        conn.autocommit = False
        log.info("✅  Conexión establecida → %s / %s", SQL_SERVER, SQL_DB)
        return conn
    except pyodbc.Error as e:
        log.error("❌  No se pudo conectar al servidor: %s", e)
        sys.exit(1)


# ══════════════════════════════════════════════════════════════
# 3.  LECTURA DEL EXCEL
# ══════════════════════════════════════════════════════════════
def leer_excel() -> dict[str, pd.DataFrame]:
    path = Path(EXCEL_PATH)
    if not path.exists():
        log.error("❌  Archivo no encontrado: %s", path.resolve())
        sys.exit(1)

    log.info("📂  Leyendo Excel: %s", path.resolve())
    hojas = {
        "CPU":          "CPU",
        "Software":     "CPU_Software",
        "Perifericos":  "CPU_Perifericos",
    }

    datos: dict[str, pd.DataFrame] = {}
    for hoja, tabla in hojas.items():
        try:
            df = pd.read_excel(path, sheet_name=hoja, dtype=str)
            # Eliminar columnas y filas completamente vacías
            df.dropna(how="all", axis=0, inplace=True)
            df.dropna(how="all", axis=1, inplace=True)
            # Reemplazar NaN por None (NULL en SQL)
            df = df.where(pd.notna(df), None)
            # Limpiar espacios al inicio/fin de strings
            df = df.apply(lambda col: col.str.strip() if col.dtype == "object" else col)
            # Convertir Timestamp de string a datetime (evita error 22007 en SQL Server)
            if "Timestamp" in df.columns:
                df["Timestamp"] = pd.to_datetime(df["Timestamp"], errors="coerce")
                df["Timestamp"] = df["Timestamp"].where(pd.notna(df["Timestamp"]), None)
            datos[tabla] = df
            log.info("   %-20s → %d filas, %d columnas", hoja, len(df), len(df.columns))
        except Exception as e:
            log.error("❌  Error al leer hoja '%s': %s", hoja, e)
            sys.exit(1)

    return datos


# ══════════════════════════════════════════════════════════════
# 4.  MIGRACIÓN CPU  (tabla principal – sin FK)
# ══════════════════════════════════════════════════════════════
def migrar_cpu(cursor: pyodbc.Cursor, df: pd.DataFrame) -> tuple[int, int, int]:
    """Devuelve (insertados, actualizados, errores)."""
    # Columnas que existen en la tabla DB (ignorar 'Tipo' del Excel)
    cols_db = [
        "Host", "No_Serie", "Empresa", "Edificio", "Area", "Estado",
        "Marca", "Modelo", "Procesador", "RAM", "Capacidad_Disco",
        "Tipo_Disco", "Observaciones", "Codigo_QR", "Codigo_Barras_CPU", "Timestamp",
    ]
    # Solo conservar columnas que existan en el dataframe
    cols_disponibles = [c for c in cols_db if c in df.columns]
    cols_sin_host    = [c for c in cols_disponibles if c != "Host"]

    ins = upd = err = 0
    for _, row in df.iterrows():
        host = row.get("Host")
        if not host:
            log.warning("   ⚠️  Fila sin Host en CPU — se omite")
            continue

        try:
            cursor.execute("SELECT COUNT(1) FROM Inventario.CPU WHERE Host = ?", host)
            existe = cursor.fetchone()[0] > 0

            if existe and MODO_UPSERT:
                # ── UPDATE si ya existe ──
                set_clause = ", ".join(f"{c} = ?" for c in cols_sin_host)
                valores    = [row.get(c) for c in cols_sin_host]
                cursor.execute(
                    f"UPDATE Inventario.CPU SET {set_clause}, Fecha_Modificacion = GETDATE() "
                    f"WHERE Host = ?",
                    valores + [host],
                )
                upd += 1
            elif not existe:
                # ── INSERT si no existe ──
                col_list  = ", ".join(cols_disponibles)
                val_marks = ", ".join("?" * len(cols_disponibles))
                cursor.execute(
                    f"INSERT INTO Inventario.CPU ({col_list}, Fecha_Creacion, Fecha_Modificacion) "
                    f"VALUES ({val_marks}, GETDATE(), GETDATE())",
                    [row.get(c) for c in cols_disponibles],
                )
                ins += 1

        except pyodbc.Error as e:
            log.error("   ❌  Error CPU Host='%s': %s", host, e)
            err += 1

    return ins, upd, err


# ══════════════════════════════════════════════════════════════
# 5.  MIGRACIÓN CPU_Software
# ══════════════════════════════════════════════════════════════
def migrar_software(cursor: pyodbc.Cursor, df: pd.DataFrame) -> tuple[int, int, int]:
    cols_db = [
        "Host_CPU", "SO", "Office", "Antivirus", "Lector_PDF",
        "ERP", "Otro_1", "Otro_2", "Otro_3", "Timestamp",
    ]
    cols_disponibles = [c for c in cols_db if c in df.columns]

    ins = upd = err = 0
    for _, row in df.iterrows():
        host = row.get("Host_CPU")
        if not host:
            continue

        try:
            # Verificar que el CPU padre exista
            cursor.execute("SELECT COUNT(1) FROM Inventario.CPU WHERE Host = ?", host)
            if cursor.fetchone()[0] == 0:
                log.warning("   ⚠️  Software: Host_CPU='%s' no existe en CPU — se omite", host)
                err += 1
                continue

            # Software: usamos Host_CPU como clave de upsert
            cursor.execute(
                "SELECT COUNT(1) FROM Inventario.CPU_Software WHERE Host_CPU = ?", host
            )
            existe = cursor.fetchone()[0] > 0

            if existe and MODO_UPSERT:
                set_clause = ", ".join(
                    f"{c} = ?" for c in cols_disponibles if c != "Host_CPU"
                )
                valores = [row.get(c) for c in cols_disponibles if c != "Host_CPU"]
                cursor.execute(
                    f"UPDATE Inventario.CPU_Software SET {set_clause} WHERE Host_CPU = ?",
                    valores + [host],
                )
                upd += 1
            elif not existe:
                col_list  = ", ".join(cols_disponibles)
                val_marks = ", ".join("?" * len(cols_disponibles))
                cursor.execute(
                    f"INSERT INTO Inventario.CPU_Software ({col_list}, Fecha_Creacion) "
                    f"VALUES ({val_marks}, GETDATE())",
                    [row.get(c) for c in cols_disponibles],
                )
                ins += 1

        except pyodbc.Error as e:
            log.error("   ❌  Error Software Host_CPU='%s': %s", host, e)
            err += 1

    return ins, upd, err


# ══════════════════════════════════════════════════════════════
# 6.  MIGRACIÓN CPU_Perifericos
# ══════════════════════════════════════════════════════════════
def migrar_perifericos(cursor: pyodbc.Cursor, df: pd.DataFrame) -> tuple[int, int, int]:
    cols_db = [
        "Host_CPU", "Tipo", "Modelo", "No_Serie", "Marca",
        "Estado", "Observaciones", "Codigo_Barras", "Codigo_ID", "Timestamp",
    ]
    cols_disponibles = [c for c in cols_db if c in df.columns]

    ins = upd = err = 0
    for _, row in df.iterrows():
        host     = row.get("Host_CPU")
        codigo   = row.get("Codigo_ID")

        if not host:
            continue

        try:
            # Verificar CPU padre
            cursor.execute("SELECT COUNT(1) FROM Inventario.CPU WHERE Host = ?", host)
            if cursor.fetchone()[0] == 0:
                log.warning("   ⚠️  Periferico: Host_CPU='%s' no existe en CPU — se omite", host)
                err += 1
                continue

            if codigo:
                cursor.execute(
                    "SELECT COUNT(1) FROM Inventario.CPU_Perifericos WHERE Codigo_ID = ?", codigo
                )
                existe = cursor.fetchone()[0] > 0
            else:
                existe = False

            if existe and MODO_UPSERT:
                set_clause = ", ".join(
                    f"{c} = ?" for c in cols_disponibles if c != "Codigo_ID"
                )
                valores = [row.get(c) for c in cols_disponibles if c != "Codigo_ID"]
                cursor.execute(
                    f"UPDATE Inventario.CPU_Perifericos SET {set_clause} WHERE Codigo_ID = ?",
                    valores + [codigo],
                )
                upd += 1
            elif not existe:
                col_list  = ", ".join(cols_disponibles)
                val_marks = ", ".join("?" * len(cols_disponibles))
                cursor.execute(
                    f"INSERT INTO Inventario.CPU_Perifericos ({col_list}, Fecha_Creacion) "
                    f"VALUES ({val_marks}, GETDATE())",
                    [row.get(c) for c in cols_disponibles],
                )
                ins += 1

        except pyodbc.Error as e:
            log.error("   ❌  Error Periferico Codigo_ID='%s': %s", codigo, e)
            err += 1

    return ins, upd, err


# ══════════════════════════════════════════════════════════════
# 7.  REGISTRO DE AUDITORÍA
# ══════════════════════════════════════════════════════════════
def registrar_auditoria(cursor: pyodbc.Cursor, tabla: str, tipo_op: str,
                         cantidad: int, detalles: str) -> None:
    try:
        cursor.execute(
            """
            INSERT INTO Inventario.Auditoria
                (Tabla_Afectada, Tipo_Operacion, Cantidad_Registros,
                 Usuario, IP_Origen, Detalles, Fecha_Operacion)
            VALUES (?, ?, ?, ?, ?, ?, GETDATE())
            """,
            tabla, tipo_op, cantidad,
            getpass.getuser(), socket.gethostbyname(socket.gethostname()),
            detalles,
        )
    except Exception:
        pass  # La auditoría nunca debe interrumpir la migración


# ══════════════════════════════════════════════════════════════
# 8.  MAIN
# ══════════════════════════════════════════════════════════════
def main():
    inicio = datetime.now()
    log.info("═" * 60)
    log.info("  MIGRACIÓN INVENTARIO EXCEL → SQL SERVER")
    log.info("  Inicio: %s", inicio.strftime("%Y-%m-%d %H:%M:%S"))
    log.info("═" * 60)

    # Leer Excel
    datos = leer_excel()

    # Conectar
    conn   = get_connection()
    cursor = conn.cursor()

    resumen: dict[str, tuple[int, int, int]] = {}

    try:
        # ── CPU (primero, sin FK) ──────────────────────────────
        log.info("\n📋  Migrando hoja CPU → Inventario.CPU ...")
        r = migrar_cpu(cursor, datos["CPU"])
        resumen["CPU"] = r
        log.info("   Insertados: %d | Actualizados: %d | Errores: %d", *r)
        registrar_auditoria(cursor, "Inventario.CPU", "UPSERT", r[0]+r[1],
                            f"Ins={r[0]}, Upd={r[1]}, Err={r[2]}")

        # ── Software ───────────────────────────────────────────
        log.info("\n💿  Migrando hoja Software → Inventario.CPU_Software ...")
        r = migrar_software(cursor, datos["CPU_Software"])
        resumen["CPU_Software"] = r
        log.info("   Insertados: %d | Actualizados: %d | Errores: %d", *r)
        registrar_auditoria(cursor, "Inventario.CPU_Software", "UPSERT", r[0]+r[1],
                            f"Ins={r[0]}, Upd={r[1]}, Err={r[2]}")

        # ── Periféricos ────────────────────────────────────────
        log.info("\n🖥️   Migrando hoja Perifericos → Inventario.CPU_Perifericos ...")
        r = migrar_perifericos(cursor, datos["CPU_Perifericos"])
        resumen["CPU_Perifericos"] = r
        log.info("   Insertados: %d | Actualizados: %d | Errores: %d", *r)
        registrar_auditoria(cursor, "Inventario.CPU_Perifericos", "UPSERT", r[0]+r[1],
                            f"Ins={r[0]}, Upd={r[1]}, Err={r[2]}")

        # ── Confirmar todo ─────────────────────────────────────
        conn.commit()
        log.info("\n✅  Transacción confirmada (COMMIT)")

    except Exception as e:
        conn.rollback()
        log.error("\n❌  Error inesperado — se realizó ROLLBACK: %s", e)
        raise
    finally:
        cursor.close()
        conn.close()

    # ── Resumen final ──────────────────────────────────────────
    fin      = datetime.now()
    duracion = (fin - inicio).total_seconds()

    log.info("\n%s", "═" * 60)
    log.info("  RESUMEN FINAL")
    log.info("  Duración: %.1f seg", duracion)
    log.info("  %-25s  %6s  %6s  %6s", "Tabla", "Inserts", "Updates", "Errores")
    log.info("  %s", "-" * 50)
    for tabla, (i, u, e) in resumen.items():
        log.info("  %-25s  %6d  %6d  %6d", tabla, i, u, e)

    total_err = sum(e for _, _, e in resumen.values())
    if total_err > 0:
        log.warning("\n  ⚠️  Hubo %d errores. Revisa el log para detalles.", total_err)
    else:
        log.info("\n  🎉  Migración completada sin errores.")
    log.info("═" * 60)


if __name__ == "__main__":
    main()