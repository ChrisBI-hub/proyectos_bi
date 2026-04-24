from __future__ import annotations

import re
from collections import defaultdict

import pandas as pd
from sqlalchemy import text


CPU_COLUMNAS = [
    "Host",
    "No_Serie",
    "Empresa",
    "Edificio",
    "Area",
    "Estado",
    "Marca",
    "Modelo",
    "Procesador",
    "RAM",
    "Capacidad_Disco",
    "Tipo_Disco",
    "Observaciones",
    "Codigo_QR",
    "Codigo_Barras_CPU",
    "Timestamp",
]

SOFTWARE_COLUMNAS = [
    "Host_CPU",
    "SO",
    "Office",
    "Antivirus",
    "Lector_PDF",
    "ERP",
    "Otro_1",
    "Otro_2",
    "Otro_3",
    "Timestamp",
]

PERIFERICOS_COLUMNAS = [
    "Host_CPU",
    "Tipo",
    "Modelo",
    "No_Serie",
    "Marca",
    "Estado",
    "Observaciones",
    "Codigo_Barras",
    "Codigo_ID",
    "Periferico_UID",
    "Timestamp",
]

RELACIONES_COLUMNAS = [
    "Codigo_QR",
    "Codigo_Barras_CPU",
    "Tipo_Periferico",
    "Codigo_Barras_Periferico",
    "Timestamp",
]

OTROS_COLUMNAS = [
    "Tipo",
    "Nombre",
    "No_Serie",
    "Marca",
    "Modelo",
    "Empresa",
    "Edificio",
    "Area",
    "Ubicacion_En_Edificio",
    "Estado",
    "Tipo_Sensor",
    "Resolucion_Pantalla",
    "Sistema_Operativo",
    "Observaciones",
    "Codigo_Barras",
    "Codigo_ID",
    "Timestamp",
]


def preparar_hojas_para_sql(hojas: dict[str, pd.DataFrame], engine) -> dict[str, pd.DataFrame]:
    hojas = {nombre: _sanear_dataframe(df) for nombre, df in hojas.items()}
    cpu = _preparar_cpu(hojas.get("cpu", pd.DataFrame()), engine)
    software = _preparar_software(hojas.get("software", pd.DataFrame()))
    perifericos = _preparar_perifericos(hojas.get("perifericos", pd.DataFrame()), engine)
    otros = _preparar_otros(hojas.get("otros", pd.DataFrame()), engine)
    relaciones = _construir_relaciones(cpu, perifericos)

    return {
        "cpu": cpu,
        "software": software,
        "perifericos": perifericos,
        "relaciones": relaciones,
        "otros": otros,
    }


def extraer_tres_letras(valor: str) -> str:
    limpio = re.sub(r"[^A-Z0-9]", "", str(valor or "").upper())
    if not limpio:
        return "GEN"
    if len(limpio) == 1:
        return f"{limpio}XX"
    if len(limpio) == 2:
        return f"{limpio}X"
    return f"{limpio[0]}{limpio[len(limpio) // 2]}{limpio[-1]}"


def _sanear_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame()
    df = df.copy()
    df.columns = [str(col).strip() for col in df.columns]
    df = df.where(pd.notna(df), None)
    return df


def _asegurar_columnas(df: pd.DataFrame, columnas: list[str]) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=columnas)
    df = df.copy()
    for col in columnas:
        if col not in df.columns:
            df[col] = None
    return df[columnas]


def _normalizar_texto(valor) -> str:
    if valor is None:
        return ""
    if isinstance(valor, float) and pd.isna(valor):
        return ""
    return str(valor).strip()


def _a_timestamp(df: pd.DataFrame) -> pd.Series:
    if "Timestamp" not in df.columns:
        return pd.Series([pd.NaT] * len(df), index=df.index)
    return pd.to_datetime(df["Timestamp"], errors="coerce")


def _deduplicar_por_ultima_fila(df: pd.DataFrame, columnas_clave: list[str]) -> pd.DataFrame:
    if df.empty:
        return df
    trabajo = df.copy()
    trabajo["_ts"] = _a_timestamp(trabajo)
    trabajo = trabajo.sort_values(by=["_ts"], kind="stable")
    trabajo = trabajo.drop_duplicates(subset=columnas_clave, keep="last")
    return trabajo.drop(columns="_ts").reset_index(drop=True)


def _limpiar_vacias(df: pd.DataFrame, columnas_utiles: list[str]) -> pd.DataFrame:
    if df.empty:
        return df
    trabajo = df.copy()
    mask = trabajo[columnas_utiles].applymap(_normalizar_texto).ne("").any(axis=1)
    return trabajo[mask].reset_index(drop=True)


def _siguiente_consecutivo(codigos: list[str], prefijo: str, ancho: int = 4) -> int:
    patron = re.compile(rf"^{re.escape(prefijo)}(\d+)$")
    maximo = 0
    for codigo in codigos:
        match = patron.match(_normalizar_texto(codigo))
        if match:
            maximo = max(maximo, int(match.group(1)))
    return maximo + 1 if maximo else 1


def _preparar_cpu(df: pd.DataFrame, engine) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=CPU_COLUMNAS)

    df = df.copy()
    if "Tipo" in df.columns:
        df = df.drop(columns=["Tipo"])
    df = _asegurar_columnas(df, CPU_COLUMNAS)
    df = _limpiar_vacias(df, ["Host", "No_Serie", "Marca", "Modelo"])
    df["Host"] = df["Host"].map(_normalizar_texto)
    df = df[df["Host"] != ""].reset_index(drop=True)
    df = _deduplicar_por_ultima_fila(df, ["Host"])

    with engine.connect() as conn:
        rows = conn.execute(
            text(
                """
                SELECT Host, Codigo_Barras_CPU, Codigo_QR
                FROM Inventario.CPU
                """
            )
        ).mappings().all()

    existentes = {
        _normalizar_texto(row["Host"]): {
            "Codigo_Barras_CPU": _normalizar_texto(row["Codigo_Barras_CPU"]),
            "Codigo_QR": _normalizar_texto(row["Codigo_QR"]),
        }
        for row in rows
    }
    siguiente_cpu = _siguiente_consecutivo(
        [row["Codigo_Barras_CPU"] for row in rows],
        "CPU",
    )

    codigos_cpu = []
    codigos_qr = []
    for _, fila in df.iterrows():
        host = _normalizar_texto(fila["Host"])
        if host in existentes and existentes[host]["Codigo_Barras_CPU"]:
            codigo_cpu = existentes[host]["Codigo_Barras_CPU"]
            codigo_qr = existentes[host]["Codigo_QR"] or codigo_cpu
        else:
            codigo_cpu = f"CPU{siguiente_cpu:04d}"
            codigo_qr = codigo_cpu
            existentes[host] = {
                "Codigo_Barras_CPU": codigo_cpu,
                "Codigo_QR": codigo_qr,
            }
            siguiente_cpu += 1
        codigos_cpu.append(codigo_cpu)
        codigos_qr.append(codigo_qr)

    df["Codigo_Barras_CPU"] = codigos_cpu
    df["Codigo_QR"] = codigos_qr
    return df.reset_index(drop=True)


def _preparar_software(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=SOFTWARE_COLUMNAS)
    df = _asegurar_columnas(df, SOFTWARE_COLUMNAS)
    df = _limpiar_vacias(df, ["Host_CPU", "SO", "Office", "ERP"])
    df["Host_CPU"] = df["Host_CPU"].map(_normalizar_texto)
    df = df[df["Host_CPU"] != ""].reset_index(drop=True)
    return _deduplicar_por_ultima_fila(df, ["Host_CPU"])


def _periferico_base(fila: pd.Series) -> str:
    partes = [
        fila.get("Host_CPU"),
        fila.get("Tipo"),
        fila.get("No_Serie"),
        fila.get("Marca"),
        fila.get("Modelo"),
    ]
    return "|".join(_normalizar_texto(valor).upper() for valor in partes)


def _preparar_perifericos(df: pd.DataFrame, engine) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=PERIFERICOS_COLUMNAS)

    df = _asegurar_columnas(df, PERIFERICOS_COLUMNAS)
    df = _limpiar_vacias(df, ["Host_CPU", "Tipo", "Modelo", "No_Serie", "Marca"])
    df["Host_CPU"] = df["Host_CPU"].map(_normalizar_texto)
    df["Tipo"] = df["Tipo"].map(_normalizar_texto)
    df = df[(df["Host_CPU"] != "") & (df["Tipo"] != "")].reset_index(drop=True)

    if df.empty:
        return pd.DataFrame(columns=PERIFERICOS_COLUMNAS)

    trabajo = df.copy()
    trabajo["_ts"] = _a_timestamp(trabajo)
    trabajo["_base"] = trabajo.apply(_periferico_base, axis=1)
    trabajo = trabajo.sort_values(by=["_base", "_ts"], kind="stable").reset_index(drop=True)
    trabajo["_ocurrencia"] = trabajo.groupby("_base").cumcount() + 1

    with engine.connect() as conn:
        rows = conn.execute(
            text(
                """
                SELECT Host_CPU, Tipo, No_Serie, Marca, Modelo, Periferico_UID, Codigo_ID
                FROM Inventario.CPU_Perifericos
                """
            )
        ).mappings().all()

    existentes_df = pd.DataFrame(rows)
    lookup_existentes: dict[tuple[str, int], tuple[str, str]] = {}
    if not existentes_df.empty:
        existentes_df = existentes_df.fillna("")
        existentes_df["_base"] = existentes_df.apply(_periferico_base, axis=1)
        existentes_df = existentes_df.sort_values(
            by=["_base", "Periferico_UID", "Codigo_ID"],
            kind="stable",
        ).reset_index(drop=True)
        existentes_df["_ocurrencia"] = existentes_df.groupby("_base").cumcount() + 1
        for _, fila in existentes_df.iterrows():
            clave = (fila["_base"], int(fila["_ocurrencia"]))
            lookup_existentes[clave] = (
                _normalizar_texto(fila["Periferico_UID"]),
                _normalizar_texto(fila["Codigo_ID"]),
            )

    siguiente_uid = _siguiente_consecutivo(
        [row["Periferico_UID"] for row in rows],
        "PER",
        ancho=6,
    )

    codigos_existentes = [_normalizar_texto(row["Codigo_ID"]) for row in rows]
    consecutivos_tipo = defaultdict(int)
    for codigo in codigos_existentes:
        prefijo = re.sub(r"\d+$", "", codigo)
        if prefijo:
            consecutivos_tipo[prefijo] = max(
                consecutivos_tipo[prefijo],
                int(re.search(r"(\d+)$", codigo).group(1)),
            )

    uids = []
    codigos = []
    for _, fila in trabajo.iterrows():
        clave = (fila["_base"], int(fila["_ocurrencia"]))
        if clave in lookup_existentes:
            uid, codigo = lookup_existentes[clave]
        else:
            prefijo = extraer_tres_letras(fila["Tipo"])
            consecutivos_tipo[prefijo] += 1
            uid = f"PER{siguiente_uid:06d}"
            codigo = f"{prefijo}{consecutivos_tipo[prefijo]:04d}"
            lookup_existentes[clave] = (uid, codigo)
            siguiente_uid += 1
        uids.append(uid)
        codigos.append(codigo)

    trabajo["Periferico_UID"] = uids
    trabajo["Codigo_ID"] = codigos
    trabajo["Codigo_Barras"] = trabajo["Codigo_ID"]
    return trabajo[PERIFERICOS_COLUMNAS].reset_index(drop=True)


def _otro_base(fila: pd.Series) -> str:
    partes = [
        fila.get("Tipo"),
        fila.get("Nombre"),
        fila.get("No_Serie"),
        fila.get("Marca"),
        fila.get("Modelo"),
        fila.get("Empresa"),
        fila.get("Edificio"),
        fila.get("Area"),
        fila.get("Ubicacion_En_Edificio"),
    ]
    return "|".join(_normalizar_texto(valor).upper() for valor in partes)


def _preparar_otros(df: pd.DataFrame, engine) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=OTROS_COLUMNAS)

    df = _asegurar_columnas(df, OTROS_COLUMNAS)
    df = _limpiar_vacias(df, ["Tipo", "Nombre", "No_Serie", "Marca", "Modelo"])
    df["Tipo"] = df["Tipo"].map(_normalizar_texto)
    df["Nombre"] = df["Nombre"].map(_normalizar_texto)
    df = df[(df["Tipo"] != "") | (df["Nombre"] != "")].reset_index(drop=True)

    if df.empty:
        return pd.DataFrame(columns=OTROS_COLUMNAS)

    trabajo = df.copy()
    trabajo["_ts"] = _a_timestamp(trabajo)
    trabajo["_base"] = trabajo.apply(_otro_base, axis=1)
    trabajo = trabajo.sort_values(by=["_base", "_ts"], kind="stable").reset_index(drop=True)
    trabajo["_ocurrencia"] = trabajo.groupby("_base").cumcount() + 1

    with engine.connect() as conn:
        rows = conn.execute(
            text(
                """
                SELECT Tipo, Nombre, No_Serie, Marca, Modelo, Empresa, Edificio, Area,
                       Ubicacion_En_Edificio, Codigo_ID
                FROM Inventario.Otros_Equipos
                """
            )
        ).mappings().all()

    existentes_df = pd.DataFrame(rows)
    lookup_existentes: dict[tuple[str, int], str] = {}
    if not existentes_df.empty:
        existentes_df = existentes_df.fillna("")
        existentes_df["_base"] = existentes_df.apply(_otro_base, axis=1)
        existentes_df = existentes_df.sort_values(by=["_base", "Codigo_ID"], kind="stable").reset_index(drop=True)
        existentes_df["_ocurrencia"] = existentes_df.groupby("_base").cumcount() + 1
        for _, fila in existentes_df.iterrows():
            lookup_existentes[(fila["_base"], int(fila["_ocurrencia"]))] = _normalizar_texto(fila["Codigo_ID"])

    codigos_existentes = [_normalizar_texto(row["Codigo_ID"]) for row in rows]
    consecutivos_tipo = defaultdict(int)
    for codigo in codigos_existentes:
        prefijo = re.sub(r"\d+$", "", codigo)
        if prefijo:
            consecutivos_tipo[prefijo] = max(
                consecutivos_tipo[prefijo],
                int(re.search(r"(\d+)$", codigo).group(1)),
            )

    codigos = []
    for _, fila in trabajo.iterrows():
        clave = (fila["_base"], int(fila["_ocurrencia"]))
        if clave in lookup_existentes:
            codigo = lookup_existentes[clave]
        else:
            prefijo = extraer_tres_letras(fila["Tipo"] or fila["Nombre"])
            consecutivos_tipo[prefijo] += 1
            codigo = f"{prefijo}{consecutivos_tipo[prefijo]:04d}"
            lookup_existentes[clave] = codigo
        codigos.append(codigo)

    trabajo["Codigo_ID"] = codigos
    trabajo["Codigo_Barras"] = trabajo["Codigo_ID"]
    return trabajo[OTROS_COLUMNAS].reset_index(drop=True)


def _construir_relaciones(cpu: pd.DataFrame, perifericos: pd.DataFrame) -> pd.DataFrame:
    if cpu.empty:
        return pd.DataFrame(columns=RELACIONES_COLUMNAS)

    filas = []
    perifericos_por_host = defaultdict(list)
    if not perifericos.empty:
        for _, fila in perifericos.iterrows():
            perifericos_por_host[_normalizar_texto(fila["Host_CPU"])].append(fila)

    for _, fila_cpu in cpu.iterrows():
        host = _normalizar_texto(fila_cpu["Host"])
        codigo_cpu = _normalizar_texto(fila_cpu["Codigo_Barras_CPU"])
        codigo_qr = _normalizar_texto(fila_cpu["Codigo_QR"]) or codigo_cpu
        timestamp = fila_cpu.get("Timestamp")

        relacionados = perifericos_por_host.get(host, [])
        if not relacionados:
            filas.append(
                {
                    "Codigo_QR": codigo_qr,
                    "Codigo_Barras_CPU": codigo_cpu,
                    "Tipo_Periferico": None,
                    "Codigo_Barras_Periferico": None,
                    "Timestamp": timestamp,
                }
            )
            continue

        for periferico in relacionados:
            filas.append(
                {
                    "Codigo_QR": codigo_qr,
                    "Codigo_Barras_CPU": codigo_cpu,
                    "Tipo_Periferico": periferico.get("Tipo"),
                    "Codigo_Barras_Periferico": periferico.get("Codigo_ID"),
                    "Timestamp": periferico.get("Timestamp") or timestamp,
                }
            )

    return pd.DataFrame(filas, columns=RELACIONES_COLUMNAS)