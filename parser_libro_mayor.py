from pathlib import Path

import pandas as pd


COLUMNAS_ESPERADAS = ["CTA. CONTABLE", "FECHA", "CONCEPTO", "DEBE", "HABER", "SALDO"]
PALABRAS_TOTAL = ("TOTAL", "SUBTOTAL", "SALDO INICIAL", "SALDO FINAL")


def _normalizar_columna(nombre: object) -> str:
    return str(nombre).strip().upper()


def _detectar_fila_encabezado(ruta_excel: Path) -> int:
    vista = pd.read_excel(ruta_excel, header=None, nrows=30)
    esperadas = {_normalizar_columna(col) for col in COLUMNAS_ESPERADAS}
    for indice, fila in vista.iterrows():
        valores = {_normalizar_columna(valor) for valor in fila.dropna().tolist()}
        if {"CTA. CONTABLE", "FECHA", "CONCEPTO", "DEBE", "HABER"}.issubset(valores):
            return int(indice)
    primera_fila = {_normalizar_columna(col) for col in pd.read_excel(ruta_excel, nrows=0).columns}
    if esperadas.intersection(primera_fila):
        return 0
    raise ValueError("El Excel no contiene las columnas esperadas del Libro Mayor.")


def _limpiar_numero(serie: pd.Series) -> pd.Series:
    texto = serie.astype(str).str.replace(",", "", regex=False).str.strip()
    texto = texto.replace({"": "0", "nan": "0", "None": "0"})
    return pd.to_numeric(texto, errors="coerce").fillna(0.0)


def _separar_cuenta(valor: object) -> tuple[str, str]:
    texto = str(valor).strip()
    partes = texto.split(maxsplit=1)
    if not partes:
        return "", ""
    codigo = partes[0].strip()
    descripcion = partes[1].strip() if len(partes) > 1 else ""
    return codigo, descripcion


def leer_libro_mayor(ruta_excel: str) -> pd.DataFrame:
    ruta = Path(ruta_excel)
    fila_header = _detectar_fila_encabezado(ruta)
    df = pd.read_excel(ruta, header=fila_header)
    df.columns = [_normalizar_columna(col) for col in df.columns]

    faltantes = [col for col in COLUMNAS_ESPERADAS if col not in df.columns]
    if faltantes:
        raise ValueError(f"Excel incompleto. Faltan columnas: {', '.join(faltantes)}")

    df = df[COLUMNAS_ESPERADAS].copy()
    df = df.rename(columns={"CTA. CONTABLE": "CTA_CONTABLE"})
    df = df.dropna(how="all")
    df["CTA_CONTABLE"] = df["CTA_CONTABLE"].ffill().astype(str).str.strip()
    df["CONCEPTO"] = df["CONCEPTO"].fillna("").astype(str).str.strip()

    conceptos_mayus = df["CONCEPTO"].str.upper()
    mascara_total = conceptos_mayus.apply(lambda texto: any(palabra in texto for palabra in PALABRAS_TOTAL))
    df = df[df["CTA_CONTABLE"].ne("") & df["CTA_CONTABLE"].ne("nan") & ~mascara_total].copy()

    df["FECHA"] = pd.to_datetime(df["FECHA"], errors="coerce", dayfirst=True, format="mixed")
    columnas_movimiento = ["DEBE", "HABER", "SALDO"]
    sin_fecha = df["FECHA"].isna()
    sin_movimiento = df[columnas_movimiento].isna().all(axis=1) | df[columnas_movimiento].fillna(0).eq(0).all(axis=1)
    sin_concepto = df["CONCEPTO"].eq("")
    df = df[~(sin_fecha & (sin_movimiento | sin_concepto))].copy()
    if df["FECHA"].isna().any():
        raise ValueError("La columna FECHA contiene valores que no se pueden parsear como fecha.")

    df["MES"] = df["FECHA"].dt.month
    df["AÑO"] = df["FECHA"].dt.year
    df["DEBE"] = _limpiar_numero(df["DEBE"])
    df["HABER"] = _limpiar_numero(df["HABER"])
    df["SALDO"] = _limpiar_numero(df["SALDO"])
    cuentas = df["CTA_CONTABLE"].apply(_separar_cuenta)
    df["CTA_CODIGO"] = cuentas.apply(lambda item: item[0])
    df["DESCRIPCION_CUENTA_LIBRO"] = cuentas.apply(lambda item: item[1])

    return df[
        [
            "CTA_CONTABLE",
            "CTA_CODIGO",
            "DESCRIPCION_CUENTA_LIBRO",
            "FECHA",
            "MES",
            "AÑO",
            "CONCEPTO",
            "DEBE",
            "HABER",
            "SALDO",
        ]
    ]


def agrupar_por_cuenta_periodo(df: pd.DataFrame) -> pd.DataFrame:
    agrupado = (
        df.groupby(["CTA_CONTABLE", "CTA_CODIGO", "DESCRIPCION_CUENTA_LIBRO", "MES", "AÑO"], as_index=False)
        .agg(TOTAL_DEBE=("DEBE", "sum"), TOTAL_HABER=("HABER", "sum"))
        .copy()
    )
    agrupado["MOVIMIENTO_NETO"] = agrupado["TOTAL_DEBE"] - agrupado["TOTAL_HABER"]
    return agrupado
