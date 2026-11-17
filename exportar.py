from importlib.util import find_spec
from io import BytesIO

import pandas as pd


COLUMNAS_MONEDA = {
    "VALOR_SRI",
    "VALOR_LIBRO_MAYOR",
    "VALOR_LIBRO",
    "VALOR_EXTRAIDO",
    "VALOR_DECLARADO",
    "BASE_IMPONIBLE",
    "VALOR_RETENIDO",
    "BASE IMPONIBLE",
    "VALOR RETENIDO",
    "VALOR FORMULARIO",
    "DIFERENCIA",
    "DEBE",
    "HABER",
    "SALDO",
}


def _crear_resumen(df_comparacion: pd.DataFrame) -> pd.DataFrame:
    if df_comparacion.empty:
        return pd.DataFrame()
    return (
        df_comparacion.pivot_table(
            index=["PERIODO_LABEL", "FORMULARIO"],
            values=["VALOR_SRI", "VALOR_LIBRO_MAYOR", "DIFERENCIA"],
            aggfunc="sum",
        )
        .reset_index()
        .sort_values(["PERIODO_LABEL", "FORMULARIO"])
    )


def _formatear_xlsxwriter(writer, hojas: dict[str, pd.DataFrame]) -> None:
    libro = writer.book
    header_fmt = libro.add_format({"bold": True, "font_color": "white", "bg_color": "#003087", "border": 1})
    money_fmt = libro.add_format({"num_format": "$ #,##0.00"})
    ok_fmt = libro.add_format({"bg_color": "#C6EFCE", "font_color": "#006100"})
    diff_fmt = libro.add_format({"bg_color": "#FFC7CE", "font_color": "#9C0006"})
    warn_fmt = libro.add_format({"bg_color": "#FFEB9C", "font_color": "#9C6500"})

    for hoja_nombre, df in hojas.items():
        hoja = writer.sheets[hoja_nombre]
        hoja.freeze_panes(1, 0)
        for col_num, col_name in enumerate(df.columns):
            hoja.write(0, col_num, col_name, header_fmt)
            ancho = max(12, min(45, len(str(col_name)) + 4))
            formato = money_fmt if col_name in COLUMNAS_MONEDA else None
            hoja.set_column(col_num, col_num, ancho, formato)

    detalle_df = hojas.get("Detalle", pd.DataFrame())
    if not detalle_df.empty:
        detalle = writer.sheets["Detalle"]
        estado_col = detalle_df.columns.get_loc("ESTADO")
        last_row = len(detalle_df)
        for texto, formato in [("OK", ok_fmt), ("DIFERENCIA", diff_fmt), ("SOLO_", warn_fmt)]:
            detalle.conditional_format(1, estado_col, last_row, estado_col, {
                "type": "text",
                "criteria": "containing",
                "value": texto,
                "format": formato,
            })
        diff_col = detalle_df.columns.get_loc("DIFERENCIA")
        detalle.conditional_format(1, diff_col, last_row, diff_col, {
            "type": "cell",
            "criteria": "!=",
            "value": 0,
            "format": diff_fmt,
        })

    resumen = hojas.get("Resumen", pd.DataFrame())
    if not resumen.empty:
        resumen_ws = writer.sheets["Resumen"]
        diff_col = resumen.columns.get_loc("DIFERENCIA")
        resumen_ws.conditional_format(1, diff_col, len(resumen), diff_col, {
            "type": "cell",
            "criteria": "==",
            "value": 0,
            "format": ok_fmt,
        })
        resumen_ws.conditional_format(1, diff_col, len(resumen), diff_col, {
            "type": "cell",
            "criteria": "!=",
            "value": 0,
            "format": diff_fmt,
        })


def _formatear_openpyxl(writer, hojas: dict[str, pd.DataFrame]) -> None:
    from openpyxl.styles import Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    header_fill = PatternFill("solid", fgColor="003087")
    header_font = Font(color="FFFFFF", bold=True)
    thin = Side(style="thin", color="AAB2BD")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    ok_fill = PatternFill("solid", fgColor="C6EFCE")
    diff_fill = PatternFill("solid", fgColor="FFC7CE")
    warn_fill = PatternFill("solid", fgColor="FFEB9C")

    for hoja_nombre, df in hojas.items():
        hoja = writer.sheets[hoja_nombre]
        hoja.freeze_panes = "A2"
        hoja.auto_filter.ref = hoja.dimensions
        for celda in hoja[1]:
            celda.fill = header_fill
            celda.font = header_font
            celda.border = border

        for indice, col_name in enumerate(df.columns, start=1):
            letra = get_column_letter(indice)
            ancho = max(12, min(45, len(str(col_name)) + 4))
            hoja.column_dimensions[letra].width = ancho
            if col_name in COLUMNAS_MONEDA:
                for fila in range(2, hoja.max_row + 1):
                    hoja.cell(fila, indice).number_format = '$ #,##0.00'

    detalle_df = hojas.get("Detalle", pd.DataFrame())
    if not detalle_df.empty:
        hoja = writer.sheets["Detalle"]
        estado_col = detalle_df.columns.get_loc("ESTADO") + 1
        diff_col = detalle_df.columns.get_loc("DIFERENCIA") + 1
        for fila in range(2, hoja.max_row + 1):
            estado = str(hoja.cell(fila, estado_col).value or "")
            if estado == "OK":
                hoja.cell(fila, estado_col).fill = ok_fill
            elif estado == "DIFERENCIA":
                hoja.cell(fila, estado_col).fill = diff_fill
            elif estado.startswith("SOLO_"):
                hoja.cell(fila, estado_col).fill = warn_fill
            diferencia = hoja.cell(fila, diff_col).value or 0
            if diferencia != 0:
                hoja.cell(fila, diff_col).fill = diff_fill

    resumen = hojas.get("Resumen", pd.DataFrame())
    if not resumen.empty:
        hoja = writer.sheets["Resumen"]
        diff_col = resumen.columns.get_loc("DIFERENCIA") + 1
        for fila in range(2, hoja.max_row + 1):
            diferencia = hoja.cell(fila, diff_col).value or 0
            hoja.cell(fila, diff_col).fill = ok_fill if diferencia == 0 else diff_fill


def generar_reporte(
    df_comparacion: pd.DataFrame,
    df_fuente: pd.DataFrame,
    nombre_hoja_fuente: str = "Libro Mayor Cargado",
    df_casilleros: pd.DataFrame | None = None,
) -> bytes:
    salida = BytesIO()
    resumen = _crear_resumen(df_comparacion)
    detalle = df_comparacion.rename(
        columns={
            "BASE_IMPONIBLE": "BASE IMPONIBLE",
            "VALOR_RETENIDO": "VALOR RETENIDO",
        }
    )
    nombre_hoja_fuente = nombre_hoja_fuente[:31]
    hojas = {
        "Resumen": resumen,
        "Detalle": detalle,
        nombre_hoja_fuente: df_fuente,
    }
    if df_casilleros is not None and not df_casilleros.empty:
        hojas["Casilleros Extraidos"] = df_casilleros

    engine = "xlsxwriter" if find_spec("xlsxwriter") is not None else "openpyxl"
    with pd.ExcelWriter(salida, engine=engine) as writer:
        for hoja_nombre, df in hojas.items():
            df.to_excel(writer, sheet_name=hoja_nombre, index=False)
        if engine == "xlsxwriter":
            _formatear_xlsxwriter(writer, hojas)
        else:
            _formatear_openpyxl(writer, hojas)

    return salida.getvalue()


def generar_reporte_formularios(
    df_formularios: pd.DataFrame,
    df_casilleros: pd.DataFrame,
) -> bytes:
    """Genera el reporte del análisis independiente, sin columnas de conciliación."""
    salida = BytesIO()
    resumen = df_formularios[
        [
            "ARCHIVO",
            "FORMULARIO",
            "PERIODO",
            "NUMERO_SERIAL",
            "FECHA_RECAUDACION",
            "CASILLERO_BASE",
            "BASE_IMPONIBLE",
            "CASILLERO_RETENIDO",
            "VALOR_RETENIDO",
            "COMENTARIO",
        ]
    ].rename(
        columns={
            "NUMERO_SERIAL": "NÚMERO SERIAL",
            "FECHA_RECAUDACION": "FECHA RECAUDACIÓN",
            "CASILLERO_BASE": "CASILLERO BASE",
            "BASE_IMPONIBLE": "BASE IMPONIBLE",
            "CASILLERO_RETENIDO": "CASILLERO RETENIDO",
            "VALOR_RETENIDO": "VALOR RETENIDO",
        }
    )
    casilleros = df_casilleros.rename(columns={"VALOR_EXTRAIDO": "VALOR FORMULARIO"})
    hojas = {
        "Resumen Formularios": resumen,
        "Casilleros Extraidos": casilleros,
    }

    engine = "xlsxwriter" if find_spec("xlsxwriter") is not None else "openpyxl"
    with pd.ExcelWriter(salida, engine=engine) as writer:
        for hoja_nombre, df in hojas.items():
            df.to_excel(writer, sheet_name=hoja_nombre, index=False)
        if engine == "xlsxwriter":
            _formatear_xlsxwriter(writer, hojas)
        else:
            _formatear_openpyxl(writer, hojas)
    return salida.getvalue()
