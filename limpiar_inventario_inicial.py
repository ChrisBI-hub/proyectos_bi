from __future__ import annotations

import argparse
from pathlib import Path

import pandas as pd


HOJAS_SALIDA = {
    "CPU": [
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
    ],
    "Software": [
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
    ],
    "Perifericos": [
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
    ],
    "Relaciones": [
        "Codigo_QR",
        "Codigo_Barras_CPU",
        "Tipo_Periferico",
        "Codigo_Barras_Periferico",
        "Timestamp",
    ],
    "Otros": [
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
    ],
}


def normalizar_texto(valor) -> str:
    if valor is None:
        return ""
    if isinstance(valor, float) and pd.isna(valor):
        return ""
    return str(valor).strip()


def sanear_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame()
    trabajo = df.copy()
    trabajo.columns = [str(c).strip() for c in trabajo.columns]
    trabajo = trabajo.where(pd.notna(trabajo), None)
    return trabajo


def asegurar_columnas(df: pd.DataFrame, columnas: list[str]) -> pd.DataFrame:
    trabajo = df.copy() if not df.empty else pd.DataFrame()
    for col in columnas:
        if col not in trabajo.columns:
            trabajo[col] = None
    return trabajo[columnas]


def serie_timestamp(df: pd.DataFrame) -> pd.Series:
    if "Timestamp" not in df.columns:
        return pd.Series([pd.NaT] * len(df), index=df.index)
    return pd.to_datetime(df["Timestamp"], errors="coerce")


def limpiar_vacias(df: pd.DataFrame, columnas_utiles: list[str]) -> pd.DataFrame:
    if df.empty:
        return df
    texto = df[columnas_utiles].apply(lambda col: col.map(normalizar_texto))
    mask = texto.ne("").any(axis=1)
    return df[mask].reset_index(drop=True)


def deduplicar_por_ultima(df: pd.DataFrame, columnas_clave: list[str]) -> pd.DataFrame:
    if df.empty:
        return df
    trabajo = df.copy()
    trabajo["_ts"] = serie_timestamp(trabajo)
    trabajo = trabajo.sort_values(by=["_ts"], kind="stable")
    trabajo = trabajo.drop_duplicates(subset=columnas_clave, keep="last")
    return trabajo.drop(columns="_ts").reset_index(drop=True)


def limpiar_cpu(df: pd.DataFrame) -> pd.DataFrame:
    df = sanear_dataframe(df)
    if "Tipo" in df.columns:
        df = df.drop(columns=["Tipo"])
    df = asegurar_columnas(df, HOJAS_SALIDA["CPU"])
    df = limpiar_vacias(df, ["Host", "No_Serie", "Marca", "Modelo"])
    df["Host"] = df["Host"].map(normalizar_texto)
    df = df[df["Host"] != ""].reset_index(drop=True)
    df = deduplicar_por_ultima(df, ["Host"])
    df["Codigo_QR"] = ""
    df["Codigo_Barras_CPU"] = ""
    return df


def limpiar_software(df: pd.DataFrame) -> pd.DataFrame:
    df = sanear_dataframe(df)
    df = asegurar_columnas(df, HOJAS_SALIDA["Software"])
    df = limpiar_vacias(df, ["Host_CPU", "SO", "Office", "ERP"])
    df["Host_CPU"] = df["Host_CPU"].map(normalizar_texto)
    df = df[df["Host_CPU"] != ""].reset_index(drop=True)
    return deduplicar_por_ultima(df, ["Host_CPU"])


def limpiar_perifericos(df: pd.DataFrame) -> pd.DataFrame:
    df = sanear_dataframe(df)
    df = asegurar_columnas(df, HOJAS_SALIDA["Perifericos"])
    df = limpiar_vacias(df, ["Host_CPU", "Tipo", "Modelo", "No_Serie", "Marca"])
    for col in ["Host_CPU", "Tipo", "Modelo", "No_Serie", "Marca", "Estado", "Observaciones"]:
        df[col] = df[col].map(normalizar_texto)
    df = df[(df["Host_CPU"] != "") & (df["Tipo"] != "")].reset_index(drop=True)
    df = deduplicar_por_ultima(df, ["Host_CPU", "Tipo", "No_Serie", "Marca", "Modelo"])
    df["Codigo_Barras"] = ""
    df["Codigo_ID"] = ""
    df["Periferico_UID"] = ""
    return df


def limpiar_otros(df: pd.DataFrame) -> pd.DataFrame:
    df = sanear_dataframe(df)
    df = asegurar_columnas(df, HOJAS_SALIDA["Otros"])
    df = limpiar_vacias(df, ["Tipo", "Nombre", "No_Serie", "Marca", "Modelo"])
    for col in [
        "Tipo",
        "Nombre",
        "No_Serie",
        "Marca",
        "Modelo",
        "Empresa",
        "Edificio",
        "Area",
        "Ubicacion_En_Edificio",
    ]:
        df[col] = df[col].map(normalizar_texto)
    df = df[(df["Tipo"] != "") | (df["Nombre"] != "")].reset_index(drop=True)
    df = deduplicar_por_ultima(
        df,
        ["Tipo", "Nombre", "No_Serie", "Marca", "Modelo", "Empresa", "Edificio", "Area", "Ubicacion_En_Edificio"],
    )
    df["Codigo_Barras"] = ""
    df["Codigo_ID"] = ""
    return df


def limpiar_relaciones() -> pd.DataFrame:
    return pd.DataFrame(columns=HOJAS_SALIDA["Relaciones"])


def limpiar_archivo_excel(ruta_entrada: Path, ruta_salida: Path) -> dict[str, tuple[int, int]]:
    xls = pd.ExcelFile(ruta_entrada)
    hojas = {nombre: pd.read_excel(ruta_entrada, sheet_name=nombre) for nombre in xls.sheet_names}

    cpu = limpiar_cpu(hojas.get("CPU", pd.DataFrame()))
    software = limpiar_software(hojas.get("Software", pd.DataFrame()))
    perifericos = limpiar_perifericos(hojas.get("Perifericos", pd.DataFrame()))
    otros = limpiar_otros(hojas.get("Otros", pd.DataFrame()))
    relaciones = limpiar_relaciones()

    with pd.ExcelWriter(ruta_salida, engine="openpyxl") as writer:
        cpu.to_excel(writer, sheet_name="CPU", index=False)
        software.to_excel(writer, sheet_name="Software", index=False)
        perifericos.to_excel(writer, sheet_name="Perifericos", index=False)
        relaciones.to_excel(writer, sheet_name="Relaciones", index=False)
        otros.to_excel(writer, sheet_name="Otros", index=False)

    resumen = {
        "CPU": (len(hojas.get("CPU", pd.DataFrame())), len(cpu)),
        "Software": (len(hojas.get("Software", pd.DataFrame())), len(software)),
        "Perifericos": (len(hojas.get("Perifericos", pd.DataFrame())), len(perifericos)),
        "Relaciones": (len(hojas.get("Relaciones", pd.DataFrame())), len(relaciones)),
        "Otros": (len(hojas.get("Otros", pd.DataFrame())), len(otros)),
    }
    return resumen


def main():
    parser = argparse.ArgumentParser(description="Limpia un inventario Excel para carga inicial.")
    parser.add_argument("entrada", help="Ruta del archivo Excel origen")
    parser.add_argument(
        "--salida",
        help="Ruta del archivo limpio de salida. Si no se indica, agrega _limpio al nombre.",
    )
    args = parser.parse_args()

    ruta_entrada = Path(args.entrada).resolve()
    ruta_salida = Path(args.salida).resolve() if args.salida else ruta_entrada.with_name(
        f"{ruta_entrada.stem}_limpio{ruta_entrada.suffix}"
    )

    resumen = limpiar_archivo_excel(ruta_entrada, ruta_salida)

    print(f"Archivo limpio generado: {ruta_salida}")
    for hoja, (antes, despues) in resumen.items():
        print(f"{hoja}: {antes} -> {despues}")


if __name__ == "__main__":
    main()