import json
import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st

from comparador import MESES_NOMBRE, comparar
from exportar import generar_reporte
from mapeo_config import cargar_mapeo, guardar_mapeo
from parser_libro_mayor import agrupar_por_cuenta_periodo, leer_libro_mayor
from parser_pdf import extraer_formulario, validar_formulario


st.set_page_config(
    page_title="Auditor SRI - Conciliacion 103/104 vs Libro Mayor",
    page_icon="🇪🇨",
    layout="wide",
)


CSS = """
<style>
    .main .block-container { padding-top: 1.4rem; }
    h1, h2, h3 { color: #003087; }
    div[data-testid="stMetric"] {
        border-left: 5px solid #003087;
        background: #f7f9fc;
        padding: 12px 14px;
        border-radius: 8px;
    }
    div.stButton > button, div.stDownloadButton > button {
        border-color: #003087;
    }
    .sri-divider {
        border-top: 3px solid #C8102E;
        margin: 0.2rem 0 1rem 0;
    }
</style>
"""


def inicializar_estado() -> None:
    valores = {
        "formularios": [],
        "df_libro_mayor": pd.DataFrame(),
        "df_libro_mayor_agrupado": pd.DataFrame(),
        "df_comparacion": pd.DataFrame(),
        "mapeos": cargar_mapeo(),
        "comentarios": {},
        "alertas_validacion": [],
    }
    for clave, valor in valores.items():
        if clave not in st.session_state:
            st.session_state[clave] = valor


def guardar_upload_temporal(upload, suffix: str) -> str:
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as archivo:
        archivo.write(upload.getvalue())
        return archivo.name


def formato_moneda(valor: float) -> str:
    return f"$ {float(valor):,.2f}"


def aplicar_comentarios(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    df = df.copy()
    for idx, fila in df.iterrows():
        clave = (fila["PERIODO_LABEL"], fila["FORMULARIO"], fila["CASILLERO"], fila["CUENTA_CONTABLE"])
        if clave in st.session_state.comentarios:
            df.at[idx, "COMENTARIO"] = st.session_state.comentarios[clave]
    return df


def cuentas_mapeo_en_libro(mapeos: list[dict], df_agrupado: pd.DataFrame) -> tuple[list[str], list[str]]:
    if df_agrupado.empty:
        return [], []
    cuentas_libro = set(df_agrupado["CTA_CONTABLE"].astype(str).str.strip())
    if "CTA_CODIGO" in df_agrupado.columns:
        cuentas_libro |= set(df_agrupado["CTA_CODIGO"].astype(str).str.strip())

    mapeadas = []
    faltantes = []
    for mapeo in mapeos:
        cuenta = str(mapeo.get("cuenta_contable", "")).strip()
        if not cuenta:
            continue
        if cuenta in cuentas_libro:
            mapeadas.append(cuenta)
        else:
            faltantes.append(cuenta)
    return sorted(set(mapeadas)), sorted(set(faltantes))


def asegurar_mapeo_compatible(df_agrupado: pd.DataFrame) -> None:
    encontradas, faltantes = cuentas_mapeo_en_libro(st.session_state.mapeos, df_agrupado)
    if encontradas:
        if faltantes:
            st.warning(
                "Algunas cuentas del mapeo no existen en el Libro Mayor cargado: "
                + ", ".join(faltantes[:12])
                + ("..." if len(faltantes) > 12 else "")
            )
        return

    mapeo_disco = cargar_mapeo()
    encontradas_disco, faltantes_disco = cuentas_mapeo_en_libro(mapeo_disco, df_agrupado)
    if encontradas_disco:
        st.session_state.mapeos = mapeo_disco
        st.info(
            "El mapeo activo no coincidía con el Libro Mayor. Se recargó automáticamente "
            "el mapeo default actualizado."
        )
        if faltantes_disco:
            st.warning(
                "Cuentas del mapeo default no encontradas en el Libro Mayor: "
                + ", ".join(faltantes_disco[:12])
                + ("..." if len(faltantes_disco) > 12 else "")
            )
        return

    muestra = df_agrupado[["CTA_CONTABLE"]].drop_duplicates().head(15)["CTA_CONTABLE"].astype(str).tolist()
    st.error(
        "Ninguna cuenta del mapeo coincide con el Libro Mayor. "
        "Esto no depende del concepto del movimiento, sino de la columna CTA. CONTABLE. "
        "Primeras cuentas detectadas: " + " | ".join(muestra)
    )


def procesar_archivos(pdfs, libro) -> None:
    if not pdfs:
        st.error("Sube al menos un PDF del SRI.")
        return
    if libro is None:
        st.error("Sube el archivo Excel del Libro Mayor.")
        return
    if not st.session_state.mapeos:
        st.error("No hay mapeos configurados al intentar comparar.")
        return

    formularios = []
    alertas = []
    with st.spinner("Extrayendo PDFs y leyendo Libro Mayor..."):
        for pdf in pdfs:
            ruta_pdf = guardar_upload_temporal(pdf, ".pdf")
            try:
                formulario = extraer_formulario(ruta_pdf)
                formularios.append(formulario)
                alertas.extend(validar_formulario(formulario))
            except Exception as exc:
                st.error(f"{pdf.name}: {exc}")

        ruta_excel = guardar_upload_temporal(libro, Path(libro.name).suffix or ".xlsx")
        try:
            df_libro = leer_libro_mayor(ruta_excel)
            df_agrupado = agrupar_por_cuenta_periodo(df_libro)
        except Exception as exc:
            st.error(str(exc))
            return

        if not formularios:
            st.error("No se pudo extraer ningun formulario valido.")
            return

        asegurar_mapeo_compatible(df_agrupado)

        try:
            df_comparacion = comparar(formularios, df_agrupado, st.session_state.mapeos)
        except Exception as exc:
            st.error(str(exc))
            return

    periodos_pdf = {(f["periodo"].get("mes"), f["periodo"].get("año")) for f in formularios}
    periodos_libro = set(zip(df_agrupado["MES"].astype(int), df_agrupado["AÑO"].astype(int)))
    faltantes = sorted(periodos_pdf - periodos_libro)
    for mes, anio in faltantes:
        st.warning(f"El periodo {MESES_NOMBRE.get(mes, mes)} {anio} del PDF no coincide con ningun periodo del Libro Mayor.")

    st.session_state.formularios = formularios
    st.session_state.df_libro_mayor = df_libro
    st.session_state.df_libro_mayor_agrupado = df_agrupado
    st.session_state.df_comparacion = aplicar_comentarios(df_comparacion)
    st.session_state.alertas_validacion = alertas
    st.success("Archivos procesados correctamente.")


def resumen_carga() -> None:
    formularios = st.session_state.formularios
    df_libro = st.session_state.df_libro_mayor
    if not formularios or df_libro.empty:
        return
    f103 = sum(1 for f in formularios if f["tipo"] == "F103")
    f104 = sum(1 for f in formularios if f["tipo"] == "F104")
    periodos = sorted({f["periodo"].get("label", "") for f in formularios if f.get("periodo")})
    st.caption(
        f"Detectados: F103={f103}, F104={f104} | Periodos: {', '.join(periodos)} | "
        f"Libro Mayor: {len(df_libro):,} filas, {df_libro['CTA_CONTABLE'].nunique():,} cuentas unicas"
    )


def panel_mapeo() -> None:
    st.sidebar.subheader("⚙️ CONFIGURACIÓN DE MAPEO")
    with st.sidebar.expander("Ver/Editar mapeos actuales", expanded=False):
        df_mapeos = pd.DataFrame(st.session_state.mapeos)
        editado = st.data_editor(
            df_mapeos,
            use_container_width=True,
            num_rows="dynamic",
            column_config={
                "formulario": st.column_config.SelectboxColumn(options=["F103", "F104"]),
                "columna_libro": st.column_config.SelectboxColumn(options=["DEBE", "HABER", "MOVIMIENTO_NETO"]),
                "signo": st.column_config.SelectboxColumn(options=[1, -1]),
            },
        )
        if st.button("Guardar cambios de mapeo"):
            st.session_state.mapeos = editado.fillna("").to_dict("records")
            guardar_mapeo(st.session_state.mapeos)
            st.success("Mapeo guardado.")

    with st.sidebar.expander("+ Agregar nuevo mapeo", expanded=False):
        with st.form("nuevo_mapeo"):
            cuenta = st.text_input("Cuenta contable")
            desc_cuenta = st.text_input("Descripcion cuenta")
            formulario = st.selectbox("Formulario", ["F104", "F103"])
            casillero = st.text_input("Casillero")
            desc_casillero = st.text_input("Descripcion casillero")
            columna = st.selectbox("Columna libro", ["HABER", "DEBE", "MOVIMIENTO_NETO"])
            signo = st.selectbox("Signo", [1, -1])
            if st.form_submit_button("Agregar"):
                nuevo = {
                    "id": max([int(m.get("id", 0)) for m in st.session_state.mapeos] or [0]) + 1,
                    "cuenta_contable": cuenta,
                    "descripcion_cuenta": desc_cuenta,
                    "formulario": formulario,
                    "casillero": casillero,
                    "descripcion_casillero": desc_casillero,
                    "columna_libro": columna,
                    "signo": signo,
                }
                st.session_state.mapeos.append(nuevo)
                guardar_mapeo(st.session_state.mapeos)
                st.success("Mapeo agregado.")

    st.sidebar.download_button(
        "Descargar mapeo actual",
        data=json.dumps({"mapeos": st.session_state.mapeos}, ensure_ascii=False, indent=2),
        file_name="mapeo_sri.json",
        mime="application/json",
    )
    if st.sidebar.button("Restaurar mapeo default"):
        st.session_state.mapeos = cargar_mapeo()
        st.sidebar.success("Mapeo default recargado.")

    carga_mapeo = st.sidebar.file_uploader("Subir mapeo guardado", type=["json"])
    if carga_mapeo is not None and st.sidebar.button("Aplicar mapeo subido"):
        datos = json.loads(carga_mapeo.getvalue().decode("utf-8"))
        st.session_state.mapeos = datos.get("mapeos", [])
        guardar_mapeo(st.session_state.mapeos)
        st.sidebar.success("Mapeo cargado.")


def filtrar(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    col1, col2, col3, col4, col5, col6, col7 = st.columns([1, 1, 1, 1, 2, 1, 1.4])
    with col1:
        meses_labels = [MESES_NOMBRE.get(m, str(m)) for m in sorted(df["MES"].dropna().unique())]
        meses_sel = st.multiselect("MES", meses_labels)
    with col2:
        anios_sel = st.multiselect("AÑO", sorted(df["AÑO"].dropna().unique()))
    with col3:
        form_sel = st.multiselect("FORMULARIO", ["Todos", "F103", "F104"], default=["Todos"])
    with col4:
        cas_sel = st.multiselect("CASILLERO", sorted(df["CASILLERO"].astype(str).unique()))
    with col5:
        concepto = st.text_input("CONCEPTO")
    with col6:
        tp_sel = st.multiselect("TP", ["Todos", "IVA", "RF"], default=["Todos"])
    with col7:
        estado_sel = st.multiselect("ESTADO", ["Todos", "OK", "DIFERENCIA", "SOLO_SRI", "SOLO_LIBRO"], default=["Todos"])

    filtrado = df.copy()
    if meses_sel:
        meses_num = [num for num, nombre in MESES_NOMBRE.items() if nombre in meses_sel]
        filtrado = filtrado[filtrado["MES"].isin(meses_num)]
    if anios_sel:
        filtrado = filtrado[filtrado["AÑO"].isin(anios_sel)]
    if form_sel and "Todos" not in form_sel:
        filtrado = filtrado[filtrado["FORMULARIO"].isin(form_sel)]
    if cas_sel:
        filtrado = filtrado[filtrado["CASILLERO"].astype(str).isin(cas_sel)]
    if concepto:
        filtrado = filtrado[filtrado["CONCEPTO"].str.contains(concepto, case=False, na=False)]
    if tp_sel and "Todos" not in tp_sel:
        filtrado = filtrado[filtrado["TP"].isin(tp_sel)]
    if estado_sel and "Todos" not in estado_sel:
        filtrado = filtrado[filtrado["ESTADO"].isin(estado_sel)]
    return filtrado


def pintar_tabla(df: pd.DataFrame):
    def estado_color(valor):
        if valor == "OK":
            return "background-color: #C6EFCE; color: #006100"
        if valor == "DIFERENCIA":
            return "background-color: #FFC7CE; color: #9C0006"
        if str(valor).startswith("SOLO_"):
            return "background-color: #FFEB9C; color: #9C6500"
        return ""

    def diferencia_color(valor):
        if valor > 0:
            return "color: #006100"
        if valor < 0:
            return "color: #C8102E"
        return "color: #666666"

    return (
        df.style.map(estado_color, subset=["ESTADO"])
        .map(diferencia_color, subset=["DIFERENCIA"])
        .format(
            {
                "VALOR SRI": formato_moneda,
                "VALOR LIBRO": formato_moneda,
                "DIFERENCIA": formato_moneda,
            }
        )
    )


def mostrar_metricas(df: pd.DataFrame) -> None:
    total = len(df)
    ok = int((df["ESTADO"] == "OK").sum()) if not df.empty else 0
    diferencia = int((df["ESTADO"] == "DIFERENCIA").sum()) if not df.empty else 0
    diff_total = df["DIFERENCIA"].sum() if not df.empty else 0.0
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Total casilleros comparados", f"{total:,}")
    col2.metric("Casilleros OK", f"🟢 {ok:,}")
    col3.metric("Con diferencia", f"🔴 {diferencia:,}")
    col4.metric("Diferencia total $", formato_moneda(diff_total))


def mostrar_detalle(df: pd.DataFrame) -> None:
    if df.empty:
        return
    opciones = {
        f"{i} | {fila.PERIODO_LABEL} | {fila.FORMULARIO} {fila.CASILLERO} | {fila.CUENTA_CONTABLE}": i
        for i, fila in df.reset_index(drop=True).iterrows()
    }
    seleccionado = st.selectbox("Ver detalle de casillero", list(opciones.keys()))
    fila = df.reset_index(drop=True).iloc[opciones[seleccionado]]

    with st.expander("Detalle de casillero", expanded=True):
        formulario = next(
            (
                f
                for f in st.session_state.formularios
                if f["tipo"] == fila["FORMULARIO"]
                and f["periodo"].get("mes") == int(fila["MES"])
                and f["periodo"].get("año") == int(fila["AÑO"])
            ),
            None,
        )
        evidencia = (formulario or {}).get("evidencias", {}).get(str(fila["CASILLERO"]), {})
        st.write(f"Valor PDF: **{formato_moneda(fila['VALOR_SRI'])}**")
        st.caption(f"Pagina: {evidencia.get('pagina', 'N/D')} | Texto: {evidencia.get('texto_original', 'N/D')}")

        movimientos = st.session_state.df_libro_mayor
        cuenta_fila = str(fila["CUENTA_CONTABLE"]).strip()
        mascara_cuenta = movimientos["CTA_CONTABLE"].astype(str).str.strip().eq(cuenta_fila)
        if "CTA_CODIGO" in movimientos.columns:
            mascara_cuenta = mascara_cuenta | movimientos["CTA_CODIGO"].astype(str).str.strip().eq(cuenta_fila)
        movimientos = movimientos[
            mascara_cuenta
            & (movimientos["MES"].astype(int) == int(fila["MES"]))
            & (movimientos["AÑO"].astype(int) == int(fila["AÑO"]))
        ].copy()
        if not movimientos.empty:
            movimientos["FECHA"] = movimientos["FECHA"].dt.strftime("%d/%m/%Y")
            st.dataframe(movimientos, use_container_width=True, height=220)
        else:
            st.info("No hay movimientos del Libro Mayor para esta cuenta y periodo.")

        clave = (fila["PERIODO_LABEL"], fila["FORMULARIO"], fila["CASILLERO"], fila["CUENTA_CONTABLE"])
        comentario = st.text_area("Comentario de auditoria", value=st.session_state.comentarios.get(clave, ""))
        if st.button("Guardar comentario"):
            st.session_state.comentarios[clave] = comentario
            st.session_state.df_comparacion = aplicar_comentarios(st.session_state.df_comparacion)
            st.success("Comentario guardado.")


def bienvenida() -> None:
    st.info("Carga los PDFs F103/F104 y el Libro Mayor para iniciar la conciliacion.")
    st.markdown(
        """
        **Uso en 5 pasos**

        1. Sube uno o varios formularios PDF del SRI.
        2. Sube el Libro Mayor en Excel con columnas `CTA. CONTABLE`, `FECHA`, `CONCEPTO`, `DEBE`, `HABER`, `SALDO`.
        3. Revisa o ajusta el mapeo cuenta-casillero en la barra lateral.
        4. Pulsa `Procesar archivos` y filtra los resultados.
        5. Descarga el reporte Excel con resumen, detalle y trazabilidad del Libro Mayor.
        """
    )


def main() -> None:
    inicializar_estado()
    st.markdown(CSS, unsafe_allow_html=True)
    st.title("🇪🇨 Auditor SRI — Conciliación 103/104 vs Libro Mayor")
    st.markdown('<div class="sri-divider"></div>', unsafe_allow_html=True)

    st.sidebar.subheader("📂 CARGA DE ARCHIVOS")
    pdfs = st.sidebar.file_uploader("Subir PDFs del SRI", type=["pdf"], accept_multiple_files=True)
    libro = st.sidebar.file_uploader("Subir Libro Mayor", type=["xlsx", "xlsm"])
    if st.sidebar.button("Procesar archivos", type="primary"):
        procesar_archivos(pdfs, libro)
    resumen_carga()
    panel_mapeo()

    for alerta in st.session_state.alertas_validacion:
        st.warning(alerta)

    df = aplicar_comentarios(st.session_state.df_comparacion)
    if df.empty:
        bienvenida()
        return

    df_filtrado = filtrar(df)
    mostrar_metricas(df_filtrado)

    visibles = df_filtrado[
        [
            "PERIODO",
            "FORMULARIO",
            "CAS",
            "CONCEPTO",
            "CUENTA_CONTABLE",
            "TP",
            "VALOR_SRI",
            "VALOR_LIBRO",
            "DIFERENCIA",
            "ESTADO",
            "COMENTARIO",
        ]
    ].rename(columns={"CUENTA_CONTABLE": "CUENTA CONTABLE", "VALOR_SRI": "VALOR SRI", "VALOR_LIBRO": "VALOR LIBRO"})

    st.dataframe(pintar_tabla(visibles), use_container_width=True, height=500)

    anio = int(df_filtrado["AÑO"].iloc[0]) if not df_filtrado.empty else "todos"
    mes = int(df_filtrado["MES"].iloc[0]) if not df_filtrado.empty else "todos"
    reporte = generar_reporte(df, st.session_state.df_libro_mayor)
    st.download_button(
        "📥 Descargar Reporte Excel",
        data=reporte,
        file_name=f"Conciliacion_SRI_{anio}_{mes}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        disabled=df.empty,
    )

    mostrar_detalle(df_filtrado)


if __name__ == "__main__":
    main()
