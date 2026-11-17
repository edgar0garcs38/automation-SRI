from typing import Any

import pandas as pd


MESES_NOMBRE = {
    1: "ENERO",
    2: "FEBRERO",
    3: "MARZO",
    4: "ABRIL",
    5: "MAYO",
    6: "JUNIO",
    7: "JULIO",
    8: "AGOSTO",
    9: "SEPTIEMBRE",
    10: "OCTUBRE",
    11: "NOVIEMBRE",
    12: "DICIEMBRE",
}


def _sumar_meses(mes: int, anio: int, cantidad: int) -> tuple[int, int]:
    indice = anio * 12 + (mes - 1) + cantidad
    return indice % 12 + 1, indice // 12


def _valor_libro(filas: pd.DataFrame, columna: str) -> float:
    columna = (columna or "MOVIMIENTO_NETO").upper()
    if columna == "DEBE":
        return float(filas["TOTAL_DEBE"].sum() or 0.0)
    if columna == "HABER":
        return float(filas["TOTAL_HABER"].sum() or 0.0)
    return float(filas["MOVIMIENTO_NETO"].sum() or 0.0)


def _filtrar_libro(libro_mayor_agrupado: pd.DataFrame, cuenta: str, mes: int, anio: int) -> pd.DataFrame:
    cuenta = str(cuenta).strip()
    mascara_periodo = (libro_mayor_agrupado["MES"].astype(int) == mes) & (
        libro_mayor_agrupado["AÑO"].astype(int) == anio
    )
    mascara_cuenta = libro_mayor_agrupado["CTA_CONTABLE"].astype(str).str.strip().eq(cuenta)
    if "CTA_CODIGO" in libro_mayor_agrupado.columns:
        mascara_cuenta = mascara_cuenta | libro_mayor_agrupado["CTA_CODIGO"].astype(str).str.strip().eq(cuenta)
    return libro_mayor_agrupado[mascara_periodo & mascara_cuenta]


def comparar(
    formularios: list[dict[str, Any]],
    libro_mayor_agrupado: pd.DataFrame,
    mapeos: list[dict[str, Any]],
) -> pd.DataFrame:
    if not mapeos:
        raise ValueError("No hay mapeos configurados para comparar.")

    formularios_idx: dict[tuple[str, int, int], dict[str, Any]] = {}
    periodos_formularios = set()
    for formulario in formularios:
        periodo = formulario.get("periodo", {})
        mes = periodo.get("mes")
        anio = periodo.get("año")
        if mes and anio:
            formularios_idx[(formulario.get("tipo"), int(mes), int(anio))] = formulario
            periodos_formularios.add((int(mes), int(anio)))

    filas = []
    for mes, anio in sorted(periodos_formularios, key=lambda item: (item[1], item[0])):
        for mapeo in mapeos:
            formulario_tipo = str(mapeo.get("formulario", "")).upper()
            cuenta = str(mapeo.get("cuenta_contable", "")).strip()
            casillero = str(mapeo.get("casillero", "")).strip()
            if not formulario_tipo or not cuenta or not casillero:
                continue

            formulario = formularios_idx.get((formulario_tipo, mes, anio))
            casilleros = formulario.get("casilleros", {}) if formulario else {}
            valor_sri = casilleros.get(casillero)
            base_imponible = casilleros.get("349") if formulario_tipo == "F103" else None
            valor_retenido = casilleros.get("399") if formulario_tipo == "F103" else None

            desfase_meses = int(mapeo.get("desfase_meses", 0) or 0)
            mes_pago, anio_pago = _sumar_meses(mes, anio, desfase_meses)
            libro_filtrado = _filtrar_libro(libro_mayor_agrupado, cuenta, mes_pago, anio_pago)
            valor_libro = None
            if not libro_filtrado.empty:
                valor_libro = _valor_libro(libro_filtrado, str(mapeo.get("columna_libro", "")))
                valor_libro *= int(mapeo.get("signo", 1) or 1)

            if valor_sri is None and valor_libro is None:
                continue

            sri_num = float(valor_sri or 0.0)
            libro_num = float(valor_libro or 0.0)
            diferencia = round(sri_num - libro_num, 2)
            if valor_sri is not None and valor_libro is None:
                estado = "SOLO_SRI"
            elif valor_sri is None and valor_libro is not None:
                estado = "SOLO_LIBRO"
            elif abs(diferencia) < 0.005:
                estado = "OK"
            else:
                estado = "DIFERENCIA"

            periodo_label = f"{MESES_NOMBRE.get(mes, mes)} {anio}"
            periodo_pago = f"{MESES_NOMBRE.get(mes_pago, mes_pago)} {anio_pago}"
            filas.append(
                {
                    "MES": mes,
                    "AÑO": anio,
                    "PERIODO_LABEL": periodo_label,
                    "PERIODO": periodo_label,
                    "MES_PAGO": mes_pago,
                    "AÑO_PAGO": anio_pago,
                    "PERIODO_PAGO": periodo_pago,
                    "FORMULARIO": formulario_tipo,
                    "NUMERO_SERIAL": (formulario or {}).get("numero_serial", ""),
                    "FECHA_RECAUDACION": (formulario or {}).get("fecha_recaudacion", ""),
                    "CASILLERO": casillero,
                    "CAS": casillero,
                    "DESCRIPCION_CASILLERO": mapeo.get("descripcion_casillero", ""),
                    "CUENTA_CONTABLE": cuenta,
                    "DESCRIPCION_CUENTA": mapeo.get("descripcion_cuenta", ""),
                    "CONCEPTO": mapeo.get("descripcion_casillero", ""),
                    "TP": "IVA" if formulario_tipo == "F104" else "RF",
                    "VALOR_SRI": sri_num if valor_sri is not None else 0.0,
                    "BASE_IMPONIBLE": float(base_imponible) if base_imponible is not None else None,
                    "VALOR_RETENIDO": float(valor_retenido) if valor_retenido is not None else None,
                    "VALOR_LIBRO_MAYOR": libro_num if valor_libro is not None else 0.0,
                    "VALOR_LIBRO": libro_num if valor_libro is not None else 0.0,
                    "DIFERENCIA": diferencia,
                    "ESTADO": estado,
                    "COMENTARIO": "",
                }
            )

    return pd.DataFrame(filas)
