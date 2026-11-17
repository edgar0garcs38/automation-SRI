from typing import Any

import pandas as pd


def _normalizar_casillero(valor: object) -> str:
    texto = str(valor).strip()
    return str(int(texto)) if texto.isdigit() else texto


def comparar_f101(formulario: dict[str, Any], df_ict: pd.DataFrame) -> pd.DataFrame:
    if formulario.get("tipo") != "F101":
        raise ValueError("El PDF seleccionado no corresponde a un formulario F101.")

    periodo = formulario.get("periodo", {})
    anio = int(periodo.get("año") or 0)
    casilleros_pdf = {
        _normalizar_casillero(codigo): float(valor)
        for codigo, valor in formulario.get("casilleros", {}).items()
    }
    ict_idx = {
        _normalizar_casillero(fila["CASILLERO"]): fila
        for _, fila in df_ict.iterrows()
    }

    filas = []
    for casillero in sorted(set(ict_idx), key=lambda x: int(x) if x.isdigit() else x):
        existe_pdf = casillero in casilleros_pdf
        existe_ict = casillero in ict_idx
        valor_pdf = casilleros_pdf.get(casillero, 0.0)
        valor_ict = float(ict_idx[casillero]["VALOR_DECLARADO"]) if existe_ict else 0.0
        if existe_pdf and existe_ict and valor_ict < 0 and valor_pdf > 0:
            valor_pdf *= -1
        diferencia = round(valor_pdf - valor_ict, 2)

        if existe_pdf and not existe_ict:
            estado = "SOLO_FORMULARIO"
        elif existe_ict and not existe_pdf:
            estado = "SOLO_ANEXO"
        elif abs(diferencia) < 0.005:
            estado = "OK"
        else:
            estado = "DIFERENCIA"

        nombre = str(ict_idx[casillero]["NOMBRE_CASILLERO"]) if existe_ict else ""
        filas.append(
            {
                "MES": 12,
                "AÑO": anio,
                "PERIODO_LABEL": periodo.get("label", f"AÑO {anio}"),
                "PERIODO": periodo.get("label", f"AÑO {anio}"),
                "FORMULARIO": "F101",
                "NUMERO_SERIAL": formulario.get("numero_serial", ""),
                "FECHA_RECAUDACION": formulario.get("fecha_recaudacion", ""),
                "CASILLERO": casillero,
                "CAS": casillero,
                "DESCRIPCION_CASILLERO": nombre,
                "CUENTA_CONTABLE": "",
                "DESCRIPCION_CUENTA": "",
                "CONCEPTO": nombre,
                "TP": "IR",
                "VALOR_SRI": valor_pdf,
                "VALOR_LIBRO_MAYOR": valor_ict,
                "VALOR_LIBRO": valor_ict,
                "DIFERENCIA": diferencia,
                "ESTADO": estado,
                "COMENTARIO": "",
            }
        )
    return pd.DataFrame(filas)
