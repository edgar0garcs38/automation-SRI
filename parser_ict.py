from pathlib import Path

import pandas as pd


HOJA_MAPEO = "MAPEO DE LA DECLARACIÓN A1"


def leer_mapeo_declaracion_ict(ruta_excel: str) -> pd.DataFrame:
    ruta = Path(ruta_excel)
    libro = pd.ExcelFile(ruta)
    hoja = next((nombre for nombre in libro.sheet_names if nombre.strip().upper() == HOJA_MAPEO.upper()), None)
    if hoja is None:
        raise ValueError(f"El Excel no contiene la hoja '{HOJA_MAPEO}'.")

    df = pd.read_excel(ruta, sheet_name=hoja, header=10)
    columnas = {str(col).strip(): col for col in df.columns}
    requeridas = ["Número de Casillero", "Nombre del Casillero", "Valor declarado"]
    faltantes = [col for col in requeridas if col not in columnas]
    if faltantes:
        raise ValueError(f"La hoja A1 no contiene las columnas esperadas: {', '.join(faltantes)}")

    salida = pd.DataFrame(
        {
            "CASILLERO": pd.to_numeric(df[columnas["Número de Casillero"]], errors="coerce"),
            "NOMBRE_CASILLERO": df[columnas["Nombre del Casillero"]].fillna("").astype(str).str.strip(),
            "VALOR_DECLARADO": pd.to_numeric(df[columnas["Valor declarado"]], errors="coerce"),
        }
    )
    salida = salida.dropna(subset=["CASILLERO"]).copy()
    salida["CASILLERO"] = salida["CASILLERO"].astype(int).astype(str)
    salida["VALOR_DECLARADO"] = salida["VALOR_DECLARADO"].fillna(0.0).astype(float)
    return salida.reset_index(drop=True)
