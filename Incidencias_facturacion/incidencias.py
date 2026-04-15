"""
script_incidencias_v2.py
========================
Analiza INCIDENCIAS_FACTURACION.xlsx y completa los campos
Ejecutivo_ABC y Filtro mediante consulta a SQL Server.

Lógica de búsqueda:
  1. Si Pedimento es válido  → busca por Cliente + Pedimento
  2. Si Pedimento es nulo/genérico/NULL → busca por Cliente + Referencia
  3. Si no hay ninguno de los dos → fila sin resultado
"""

import os
import warnings
import urllib
import smtplib
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from typing import Optional

import pandas as pd
from sqlalchemy import create_engine, text

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# ──────────────────────────────────────────────
# CONFIGURACIÓN CENTRALIZADA
# ──────────────────────────────────────────────
SQL_CONFIG = {
    "server":   "150.1.1.152",
    "database": "SIR",
    "username": "ConsultaBD",
    "password": "5D$bc#kM&5W2T8J40?s%",
    "driver":   "{ODBC Driver 17 for SQL Server}",
}

ARCHIVO_ENTRADA = "INCIDENCIAS_FACTURACION.xlsx"
ARCHIVO_SALIDA  = "Incidencias_facturacion_final.xlsx"
ARCHIVO_CSV     = "Incidencias_facturacion_final.csv"

# Pedimentos que NO sirven como clave de búsqueda
PEDIMENTOS_GENERICOS = {
    "0", "00", "000", "0000000",
    "4000000", "4M",
    "5000000", "5M",
    "6000000", "6M",
    "NULL", "NONE", "N/A", "NA", "S/N",
}

# ──────────────────────────────────────────────
# CONFIGURACIÓN DE CORREO
# ──────────────────────────────────────────────
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT   = 465
EMAIL_USER  = "reportes.bi@abcsc.mx"
EMAIL_PASS  = "jwvjdrvmprzrwzxy"

PARA = ["ccarbajal@abcsc.mx"]
CC   = ["ymontoya@abcsc.mx"]
CCO  = ["sgonzalez@abcsc.mx"]

# Posibles nombres de la columna cliente en el Excel
NOMBRES_COLUMNA_CLIENTE = ["Cliente_nombre_largo", "Cliente nombre largo"]

TABLA_SQL = "[Admin].[SIR_VT_Sabana_Pedimento_ABC]"


# ──────────────────────────────────────────────
# CONEXIÓN
# ──────────────────────────────────────────────
def conectar_sql_server():
    """Crea y valida la conexión a SQL Server. Retorna engine o None."""
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
        print(f"✅ Conexión exitosa → {c['server']} / {c['database']}")
        return engine
    except Exception as e:
        print(f"❌ Error de conexión: {e}")
        return None


# ──────────────────────────────────────────────
# UTILIDADES
# ──────────────────────────────────────────────
def limpiar_dato(valor) -> Optional[str]:
    """
    Normaliza un valor de celda Excel a string limpio.
    Retorna None si el valor está vacío, es NaN, o es uno de los
    strings que representan ausencia de dato ("NULL", "NONE", etc.).
    También elimina el sufijo '.0' que Pandas agrega a números leídos como float.
    """
    if pd.isna(valor):
        return None
    s = str(valor).strip().replace("\xa0", " ")
    # Eliminar sufijo decimal inútil (e.g. "4000000.0" → "4000000")
    if s.endswith(".0"):
        s = s[:-2]
    if s.upper() in {"NAN", "NONE", "NULL", "N/A", "NA", "S/N", ""}:
        return None
    return s


def es_pedimento_valido(pedimento: Optional[str]) -> bool:
    """Retorna True sólo si el pedimento es una clave real de búsqueda."""
    return bool(pedimento) and pedimento.upper() not in PEDIMENTOS_GENERICOS


def detectar_columna_cliente(columnas) -> Optional[str]:
    """Encuentra el nombre exacto de la columna cliente en el DataFrame."""
    for nombre in NOMBRES_COLUMNA_CLIENTE:
        if nombre in columnas:
            return nombre
    return None


# ──────────────────────────────────────────────
# CONSULTA SQL  (individual, con fallback)
# ──────────────────────────────────────────────
QUERY_BASE = f"""
    SELECT TOP 1
        [Referencia],
        [Pedimento],
        [Cliente],
        [Ejecutivo_ABC],
        [Ejecutivo] AS [FILTRO]
    FROM {TABLA_SQL}
    WHERE [Cliente] = :cliente
      AND {{clausula}}
    ORDER BY
        CASE WHEN [Ejecutivo_ABC] IS NOT NULL THEN 0 ELSE 1 END,
        CASE WHEN [Ejecutivo]     IS NOT NULL THEN 0 ELSE 1 END
"""

def buscar_registro(engine, cliente: str, pedimento: Optional[str], referencia: Optional[str]):
    """
    Busca en SQL el ejecutivo y filtro para una fila.

    Estrategia:
      - Pedimento válido  → filtra por Pedimento
      - Pedimento inválido pero Referencia existe → filtra por Referencia
      - Nada disponible   → retorna None
    """
    cliente_limpio   = limpiar_dato(cliente)
    pedimento_limpio = limpiar_dato(pedimento)
    referencia_limpia = limpiar_dato(referencia)

    if not cliente_limpio:
        return None

    if es_pedimento_valido(pedimento_limpio):
        clausula = "[Pedimento] = :valor"
        params   = {"cliente": cliente_limpio, "valor": pedimento_limpio}
        modo     = f"Pedimento={pedimento_limpio}"
    elif referencia_limpia:
        clausula = "[Referencia] = :valor"
        params   = {"cliente": cliente_limpio, "valor": referencia_limpia}
        modo     = f"Referencia={referencia_limpia}"
    else:
        return None  # Sin datos suficientes para buscar

    try:
        query = text(QUERY_BASE.format(clausula=clausula))
        with engine.connect() as conn:
            resultado = pd.read_sql_query(query, conn, params=params)
        if resultado.empty:
            return None
        fila = resultado.iloc[0]
        fila["_modo_busqueda"] = modo      # metadata para el log
        return fila
    except Exception as e:
        print(f"   ⚠️  Error consultando [{modo}] cliente [{cliente_limpio}]: {e}")
        return None


# ──────────────────────────────────────────────
# CONSULTA EN LOTE  (más eficiente)
# ──────────────────────────────────────────────
def construir_consulta_lote(pares_ped, pares_ref):
    """
    Genera una sola query que trae resultados para múltiples
    (Cliente, Pedimento) y (Cliente, Referencia) a la vez.
    Retorna (query_text, params_dict) o None si no hay pares.
    """
    condiciones = []
    params: dict = {}

    for idx, (cliente, valor) in enumerate(pares_ped):
        c_key = f"cp{idx}"
        v_key = f"vp{idx}"
        condiciones.append(f"([Cliente] = :{c_key} AND [Pedimento] = :{v_key})")
        params[c_key] = cliente
        params[v_key] = valor

    for idx, (cliente, valor) in enumerate(pares_ref):
        c_key = f"cr{idx}"
        v_key = f"vr{idx}"
        condiciones.append(f"([Cliente] = :{c_key} AND [Referencia] = :{v_key})")
        params[c_key] = cliente
        params[v_key] = valor

    if not condiciones:
        return None, None

    where_clause = " OR ".join(condiciones)
    query = text(f"""
        SELECT
            [Referencia], [Pedimento], [Cliente],
            [Ejecutivo_ABC], [Ejecutivo] AS [FILTRO]
        FROM {TABLA_SQL}
        WHERE {where_clause}
    """)
    return query, params


# ──────────────────────────────────────────────
# ENVÍO DE CORREO
# ──────────────────────────────────────────────
def enviar_correo(archivos: list[str], total: int, encontrados: int):
    """
    Envía el reporte por correo con los archivos adjuntos.
    archivos: lista de rutas a adjuntar (xlsx y csv).
    """
    hoy        = datetime.now().strftime("%d/%m/%Y %H:%M")
    sin_result = total - encontrados
    pct        = encontrados / total * 100 if total else 0

    asunto = f"Incidencias Facturación — Reporte {datetime.now().strftime('%d/%m/%Y')}"

    cuerpo_html = f"""
    <html><body style="font-family:Arial,sans-serif;font-size:14px;color:#333;">
        <p>Buen día,</p>
        <p>Se adjunta el reporte de <b>Incidencias de Facturación</b> procesado el {hoy}.</p>
        <table border="1" cellpadding="6" cellspacing="0"
               style="border-collapse:collapse;width:340px;">
            <tr style="background:#003366;color:#fff;">
                <th>Concepto</th><th>Valor</th>
            </tr>
            <tr><td>Total de filas</td><td><b>{total}</b></td></tr>
            <tr style="background:#f2f2f2;">
                <td>Con ejecutivo asignado</td>
                <td><b>{encontrados}</b> ({pct:.1f}%)</td>
            </tr>
            <tr><td>Sin resultado</td><td><b>{sin_result}</b></td></tr>
        </table>
        <p style="margin-top:16px;">Se adjuntan los archivos en formato <b>Excel</b> y <b>CSV</b>.</p>
        <p style="color:#888;font-size:12px;">— BI Extractor · Módulo Incidencias Facturación</p>
    </body></html>
    """

    todos_para = PARA + CC  # smtplib necesita todos los destinatarios visibles juntos

    msg = MIMEMultipart("mixed")
    msg["Subject"] = asunto
    msg["From"]    = EMAIL_USER
    msg["To"]      = ", ".join(PARA)
    msg["Cc"]      = ", ".join(CC)
    # CCO no se declara en headers para que sea invisible

    msg.attach(MIMEText(cuerpo_html, "html", "utf-8"))

    # Adjuntar archivos
    for ruta in archivos:
        if not os.path.exists(ruta):
            print(f"   ⚠️  Adjunto no encontrado, se omite: {ruta}")
            continue
        with open(ruta, "rb") as f:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(f.read())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f'attachment; filename="{os.path.basename(ruta)}"')
        msg.attach(part)

    destinatarios = PARA + CC + CCO   # incluye CCO en el envío real

    try:
        with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT) as server:
            server.login(EMAIL_USER, EMAIL_PASS)
            server.sendmail(EMAIL_USER, destinatarios, msg.as_string())
        print(f"📧 Correo enviado correctamente a: {', '.join(PARA + CC)}")
        if CCO:
            print(f"   (CCO: {', '.join(CCO)})")
    except smtplib.SMTPAuthenticationError:
        print("❌ Error de autenticación SMTP: verifica EMAIL_USER y EMAIL_PASS.")
    except smtplib.SMTPException as e:
        print(f"❌ Error SMTP al enviar correo: {e}")
    except Exception as e:
        print(f"❌ Error inesperado al enviar correo: {e}")


# ──────────────────────────────────────────────
# PROCESO PRINCIPAL
# ──────────────────────────────────────────────
def main():
    print("=" * 55)
    print("  BI EXTRACTOR — MÓDULO INCIDENCIAS FACTURACIÓN v2")
    print("=" * 55)

    # ── 1. Validar archivo ──────────────────────────────────
    if not os.path.exists(ARCHIVO_ENTRADA):
        print(f"❌ No se encontró el archivo: {ARCHIVO_ENTRADA}")
        return

    # ── 2. Cargar Excel ─────────────────────────────────────
    # Forzamos Pedimento y Referencia a string para evitar LossySetitem
    df = pd.read_excel(ARCHIVO_ENTRADA, dtype={"Pedimento": str, "Referencia": str})
    print(f"📄 Archivo cargado: {ARCHIVO_ENTRADA}  ({len(df)} filas, {len(df.columns)} columnas)")

    # ── 3. Detectar columna cliente ─────────────────────────
    col_cliente = detectar_columna_cliente(df.columns)
    if not col_cliente:
        print(f"❌ No se encontró la columna de cliente. Esperadas: {NOMBRES_COLUMNA_CLIENTE}")
        print(f"   Columnas disponibles: {list(df.columns)}")
        return
    print(f"✔  Columna cliente detectada: '{col_cliente}'")

    # ── 4. Preparar columnas de salida ──────────────────────
    for col in ["Ejecutivo_ABC", "Filtro"]:
        if col not in df.columns:
            df[col] = None
        df[col] = df[col].astype(object)

    # Limpiar columna residual de corridas anteriores
    if "Referencia_SQL" in df.columns:
        df.drop(columns=["Referencia_SQL"], inplace=True)

    # ── 5. Conectar ─────────────────────────────────────────
    engine = conectar_sql_server()
    if not engine:
        return

    # ── 6. Clasificar filas según estrategia de búsqueda ───
    pares_ped: list[tuple] = []   # (cliente, pedimento) → índices de fila
    pares_ref: list[tuple] = []   # (cliente, referencia) → índices de fila
    sin_datos: list[int]   = []

    # Mapas para cruzar resultados del lote con filas del df
    mapa_ped: dict[tuple, list[int]] = {}  # (cliente, pedimento) → [filas]
    mapa_ref: dict[tuple, list[int]] = {}  # (cliente, referencia) → [filas]

    for i, row in df.iterrows():
        cliente   = limpiar_dato(row[col_cliente])
        pedimento = limpiar_dato(row.get("Pedimento"))
        referencia = limpiar_dato(row.get("Referencia"))

        if not cliente:
            sin_datos.append(i)
            continue

        if es_pedimento_valido(pedimento):
            clave = (cliente, pedimento)
            mapa_ped.setdefault(clave, []).append(i)
        elif referencia:
            clave = (cliente, referencia)
            mapa_ref.setdefault(clave, []).append(i)
        else:
            sin_datos.append(i)

    pares_ped_unicos = list(mapa_ped.keys())
    pares_ref_unicos = list(mapa_ref.keys())

    print(f"\n📊 Distribución de búsqueda:")
    print(f"   Por Pedimento : {len(pares_ped_unicos)} claves únicas ({sum(len(v) for v in mapa_ped.values())} filas)")
    print(f"   Por Referencia: {len(pares_ref_unicos)} claves únicas ({sum(len(v) for v in mapa_ref.values())} filas)")
    print(f"   Sin datos     : {len(sin_datos)} filas")

    # ── 7. Consulta en lote ─────────────────────────────────
    BATCH_SIZE = 200   # Cuántas condiciones OR por query (ajustable)

    resultados_ped: dict[tuple, pd.Series] = {}
    resultados_ref: dict[tuple, pd.Series] = {}

    def ejecutar_lote(pares_p, pares_r):
        query, params = construir_consulta_lote(pares_p, pares_r)
        if query is None:
            return {}, {}
        res_p, res_r = {}, {}
        try:
            with engine.connect() as conn:
                df_lote = pd.read_sql_query(query, conn, params=params)

            # Indexar por (Cliente, Pedimento)
            for _, fila in df_lote.iterrows():
                clave_p = (fila["Cliente"], str(fila["Pedimento"]).strip()
                           if pd.notna(fila["Pedimento"]) else "")
                if clave_p in {k: None for k in pares_p}:
                    res_p.setdefault(clave_p, fila)

                clave_r = (fila["Cliente"], str(fila["Referencia"]).strip()
                           if pd.notna(fila["Referencia"]) else "")
                if clave_r in {k: None for k in pares_r}:
                    res_r.setdefault(clave_r, fila)

        except Exception as e:
            print(f"   ⚠️  Error en lote: {e} → cambiando a modo individual")
            # Fallback a consultas individuales
            for cliente, valor in pares_p:
                row_res = buscar_registro(engine, cliente, valor, None)
                if row_res is not None:
                    res_p[(cliente, valor)] = row_res
            for cliente, valor in pares_r:
                row_res = buscar_registro(engine, cliente, None, valor)
                if row_res is not None:
                    res_r[(cliente, valor)] = row_res

        return res_p, res_r

    print(f"\n🔍 Ejecutando consultas en lote (batch={BATCH_SIZE})...")

    for start in range(0, max(len(pares_ped_unicos), len(pares_ref_unicos), 1), BATCH_SIZE):
        lote_p = pares_ped_unicos[start:start + BATCH_SIZE]
        lote_r = pares_ref_unicos[start:start + BATCH_SIZE]
        rp, rr = ejecutar_lote(lote_p, lote_r)
        resultados_ped.update(rp)
        resultados_ref.update(rr)
        print(f"   Lote {start//BATCH_SIZE + 1}: {len(rp)} hits por pedimento, {len(rr)} hits por referencia")

    # ── 8. Aplicar resultados al DataFrame ──────────────────
    encontrados = 0

    def aplicar_resultado(idx_filas: list, resultado: pd.Series, via_referencia: bool):
        nonlocal encontrados
        for i in idx_filas:
            df.at[i, "Ejecutivo_ABC"] = str(resultado["Ejecutivo_ABC"]) if pd.notna(resultado["Ejecutivo_ABC"]) else None
            df.at[i, "Filtro"]        = str(resultado["FILTRO"])        if pd.notna(resultado["FILTRO"])        else None

            # ── Completar Referencia si la celda original está vacía ──────
            ref_sql    = limpiar_dato(resultado["Referencia"])
            ref_actual = limpiar_dato(df.at[i, "Referencia"])
            if ref_sql and not ref_actual:
                df.at[i, "Referencia"] = ref_sql

            # ── Reemplazar Pedimento si era genérico, nulo o exactamente "0"
            ped_actual = limpiar_dato(df.at[i, "Pedimento"])
            ped_sql    = limpiar_dato(resultado["Pedimento"])
            if ped_sql and es_pedimento_valido(ped_sql):
                if not ped_actual or ped_actual in PEDIMENTOS_GENERICOS or ped_actual == "0":
                    df.at[i, "Pedimento"] = ped_sql

            encontrados += 1

    for clave, filas in mapa_ped.items():
        resultado = resultados_ped.get(clave)
        if resultado is not None:
            aplicar_resultado(filas, resultado, via_referencia=False)

    for clave, filas in mapa_ref.items():
        resultado = resultados_ref.get(clave)
        if resultado is not None:
            aplicar_resultado(filas, resultado, via_referencia=True)

    # ── 9. Resumen ──────────────────────────────────────────
    total = len(df)
    sin_resultado = total - encontrados
    print(f"\n📈 RESUMEN:")
    print(f"   Total filas      : {total}")
    print(f"   Con resultado    : {encontrados}  ({encontrados/total*100:.1f}%)")
    print(f"   Sin resultado    : {sin_resultado}  ({sin_resultado/total*100:.1f}%)")

    # ── 10. Guardar Excel y CSV ─────────────────────────────
    archivos_generados = []
    try:
        df.to_excel(ARCHIVO_SALIDA, index=False)
        print(f"\n💾 Excel guardado  → {ARCHIVO_SALIDA}")
        archivos_generados.append(ARCHIVO_SALIDA)
    except Exception as e:
        print(f"❌ Error al guardar Excel: {e}")

    try:
        df.to_csv(ARCHIVO_CSV, index=False, encoding="utf-8-sig")
        print(f"💾 CSV guardado    → {ARCHIVO_CSV}")
        archivos_generados.append(ARCHIVO_CSV)
    except Exception as e:
        print(f"❌ Error al guardar CSV: {e}")

    engine.dispose()
    print("🔌 Conexión cerrada.")

    # ── 11. Enviar correo ────────────────────────────────────
    if archivos_generados:
        print("\n📤 Enviando correo...")
        enviar_correo(archivos_generados, total, encontrados)

    print("\n✅ PROCESO TERMINADO")


if __name__ == "__main__":
    main()