import re
from pathlib import Path
from typing import Any

import pdfplumber
import pandas as pd


MESES = {
    "ENERO": 1,
    "FEBRERO": 2,
    "MARZO": 3,
    "ABRIL": 4,
    "MAYO": 5,
    "JUNIO": 6,
    "JULIO": 7,
    "AGOSTO": 8,
    "SEPTIEMBRE": 9,
    "SETIEMBRE": 9,
    "OCTUBRE": 10,
    "NOVIEMBRE": 11,
    "DICIEMBRE": 12,
}


CODIGOS_NO_CASILLERO = {"1021", "1031", "2011"}


def _normalizar_texto(texto: str) -> str:
    return re.sub(r"\s+", " ", texto or "").strip()


def _es_casillero_valido(codigo: str, tipo: str | None) -> bool:
    if not codigo.isdigit() or not 2 <= len(codigo) <= 4:
        return False
    numero = int(codigo)
    if codigo in CODIGOS_NO_CASILLERO or 1900 <= numero <= 2100:
        return False
    if tipo == "F101":
        return True
    if tipo == "F104":
        return len(codigo) == 4 or 401 <= numero <= 999
    if tipo == "F103":
        return len(codigo) == 4 or 302 <= numero <= 999
    return True


def _identificar_tipo(texto: str) -> str:
    texto_mayus = texto.upper()
    obligacion = re.search(
        r"OBLIGACI[ÓO]N\s+TRIBUTARIA\s*:?\s*(\d{4})",
        texto_mayus,
    )
    codigo_obligacion = obligacion.group(1) if obligacion else ""
    if codigo_obligacion == "1021" or "DECLARACIÓN DE IMPUESTO A LA RENTA SOCIEDADES" in texto_mayus:
        return "F101"
    if codigo_obligacion == "2011" or "DECLARACION DE IVA" in texto_mayus or "DECLARACIÓN DE IVA" in texto_mayus:
        return "F104"
    if codigo_obligacion == "1031" or "RETENCIONES EN LA FUENTE" in texto_mayus:
        return "F103"
    raise ValueError("El PDF no corresponde a un formulario F103 ni F104 reconocible.")


def _extraer_periodo(texto: str) -> dict[str, Any]:
    anual = re.search(r"Per[ií]odo\s+Fiscal\s*:?\s*A[NÑ]O\s+(\d{4})", texto, flags=re.IGNORECASE)
    if anual:
        anio = int(anual.group(1))
        return {"mes": 12, "año": anio, "label": f"AÑO {anio}"}
    patrones = [
        r"Per[ií]odo\s+Fiscal\s*:?\s*([A-ZÁÉÍÓÚÑ]+)\s+(\d{4})",
        r"Mes\s*:?\s*([A-ZÁÉÍÓÚÑ]+).*?A[nñ]o\s*:?\s*(\d{4})",
        r"\b(ENERO|FEBRERO|MARZO|ABRIL|MAYO|JUNIO|JULIO|AGOSTO|SEPTIEMBRE|SETIEMBRE|OCTUBRE|NOVIEMBRE|DICIEMBRE)\s+(\d{4})\b",
    ]
    texto_mayus = texto.upper()
    for patron in patrones:
        encontrado = re.search(patron, texto_mayus, flags=re.IGNORECASE)
        if encontrado:
            mes_nombre = encontrado.group(1).upper()
            anio = int(encontrado.group(2))
            mes = MESES.get(mes_nombre)
            if mes:
                return {"mes": mes, "año": anio, "label": f"{mes_nombre} {anio}"}
    return {"mes": None, "año": None, "label": "PERIODO NO DETECTADO"}


def _extraer_serial_fecha(texto: str) -> tuple[str, str]:
    patron_footer = re.search(
        r"SRIDEC\w*\s+(\d{10,14})\s+(\d{2}[-/]\d{2}[-/]\d{4})",
        texto,
        flags=re.IGNORECASE,
    )
    if patron_footer:
        return patron_footer.group(1), patron_footer.group(2).replace("/", "-")

    serial = ""
    fecha = ""
    encontrado_serial = re.search(r"N[uú]mero\s+de\s+serie\s*:?\s*(\d{10,14})", texto, flags=re.IGNORECASE)
    if encontrado_serial:
        serial = encontrado_serial.group(1)
    encontrado_fecha = re.search(
        r"Fecha\s+y\s+hora\s+de\s+declaraci[oó]n\s*:?\s*(\d{2}[/\-]\d{2}[/\-]\d{4})",
        texto,
        flags=re.IGNORECASE,
    )
    if encontrado_fecha:
        fecha = encontrado_fecha.group(1).replace("/", "-")
    return serial, fecha


def _extraer_ruc(texto: str) -> str:
    encontrado = re.search(r"\b(?:RUC|No\.\s*RUC|N[uú]mero\s+RUC)\s*:?\s*(\d{13})\b", texto, flags=re.IGNORECASE)
    if encontrado:
        return encontrado.group(1)
    encontrado = re.search(r"\b\d{13}\b", texto)
    return encontrado.group(0) if encontrado else ""


def _extraer_razon_social(texto: str) -> str:
    patrones = [
        r"Raz[oó]n\s+Social\s*:?\s*([^\n\r]+)",
        r"Apellidos\s+y\s+Nombres\s+o\s+Raz[oó]n\s+Social\s*:?\s*([^\n\r]+)",
    ]
    for patron in patrones:
        encontrado = re.search(patron, texto, flags=re.IGNORECASE)
        if encontrado:
            return _normalizar_texto(encontrado.group(1))
    return ""


def _valor_a_float(valor: str) -> float:
    limpio = re.sub(r"[^\d,.\-]", "", valor)
    if "," in limpio and "." in limpio:
        if limpio.rfind(",") > limpio.rfind("."):
            limpio = limpio.replace(".", "").replace(",", ".")
        else:
            limpio = limpio.replace(",", "")
    elif "," in limpio:
        partes = limpio.split(",")
        if len(partes[-1]) == 2:
            limpio = "".join(partes[:-1]).replace(".", "") + "." + partes[-1]
        else:
            limpio = limpio.replace(",", "")
    elif limpio.count(".") > 1:
        partes = limpio.split(".")
        limpio = "".join(partes[:-1]) + "." + partes[-1]
    return float(limpio)


VALOR_RE = r"(?:\$\s*)?(?:\d{1,3}(?:[.,]\d{3})+|\d+)[.,]\d{2}"


def _guardar_casillero(
    casilleros: dict[str, float],
    evidencias: dict[str, dict[str, Any]],
    codigo: str,
    valor_txt: str,
    tipo: str,
    pagina_num: int,
    texto: str,
    inicio: int,
    fin: int,
) -> None:
    if not _es_casillero_valido(codigo, tipo):
        return
    try:
        valor = _valor_a_float(valor_txt)
    except ValueError:
        return
    casilleros[codigo] = valor
    contexto_inicio = max(0, inicio - 80)
    contexto_fin = min(len(texto), fin + 100)
    evidencias[codigo] = {
        "pagina": pagina_num,
        "texto_original": _normalizar_texto(texto[contexto_inicio:contexto_fin]),
    }


def _extraer_casilleros_pagina(
    texto: str,
    tipo: str,
    pagina_num: int,
    casilleros: dict[str, float],
    evidencias: dict[str, dict[str, Any]],
) -> None:
    texto = (texto or "").replace("\u00a0", " ")

    # Algunos PDF del SRI pegan columnas contiguas al extraer el texto. Por
    # ejemplo, "1858.24421 278.74" representa valor 1858.24, casillero 421 y
    # valor 278.74. Se captura el segundo par sin alterar el primer importe.
    patron_valor_casillero = re.compile(rf"{VALOR_RE}(\d{{3,4}})\s+({VALOR_RE})")
    for encontrado in patron_valor_casillero.finditer(texto):
        codigo, valor_txt = encontrado.groups()
        _guardar_casillero(
            casilleros,
            evidencias,
            codigo,
            valor_txt,
            tipo,
            pagina_num,
            texto,
            encontrado.start(1),
            encontrado.end(),
        )

    # También aparecen totales pegados al último operando de una fórmula:
    # "483+484499 278.74" equivale a "483+484  499 278.74".
    patron_formula_casillero = re.compile(
        rf"\b\d{{3,4}}(?:\s*[+\-]\s*\d{{3,4}})*\s*[+\-]\s*\d{{3,4}}(\d{{3}})\s+({VALOR_RE})"
    )
    for encontrado in patron_formula_casillero.finditer(texto):
        codigo, valor_txt = encontrado.groups()
        _guardar_casillero(
            casilleros,
            evidencias,
            codigo,
            valor_txt,
            tipo,
            pagina_num,
            texto,
            encontrado.start(1),
            encontrado.end(),
        )

    # Formato habitual: "401 5264.28", con separadores monetarios locales o símbolo $.
    digitos_codigo = r"\d{2,4}" if tipo == "F101" else r"\d{3,4}"
    patron_directo = re.compile(rf"(\b{digitos_codigo}\b)\s*[:=\-]?\s*({VALOR_RE})")
    for encontrado in patron_directo.finditer(texto):
        codigo, valor_txt = encontrado.groups()
        _guardar_casillero(
            casilleros,
            evidencias,
            codigo,
            valor_txt,
            tipo,
            pagina_num,
            texto,
            encontrado.start(),
            encontrado.end(),
        )

    # Respaldo para líneas donde hay texto entre el código y el importe. Se
    # limita a una sola línea para no emparejar números de secciones distintas.
    patron_eventos = re.compile(rf"\b{digitos_codigo}\b|{VALOR_RE}")
    desplazamiento = 0
    for linea in texto.splitlines(keepends=True):
        codigo_pendiente: tuple[str, int] | None = None
        for evento in patron_eventos.finditer(linea):
            token = evento.group(0)
            if re.fullmatch(digitos_codigo, token):
                if _es_casillero_valido(token, tipo):
                    codigo_pendiente = (token, desplazamiento + evento.start())
                continue
            if codigo_pendiente is None:
                continue
            codigo, inicio = codigo_pendiente
            clave_ya_extraida = codigo in casilleros
            if not clave_ya_extraida and desplazamiento + evento.start() - inicio <= 300:
                _guardar_casillero(
                    casilleros,
                    evidencias,
                    codigo,
                    token,
                    tipo,
                    pagina_num,
                    texto,
                    inicio,
                    desplazamiento + evento.end(),
                )
            codigo_pendiente = None
        desplazamiento += len(linea)


def extraer_formulario(ruta_pdf: str) -> dict[str, Any]:
    ruta = Path(ruta_pdf)
    textos_paginas: list[tuple[int, str]] = []

    try:
        with pdfplumber.open(ruta) as pdf:
            for indice, pagina in enumerate(pdf.pages, start=1):
                texto = pagina.extract_text() or ""
                textos_paginas.append((indice, texto))
    except Exception as exc:
        raise ValueError(f"No se pudo leer el PDF '{ruta.name}': {exc}") from exc

    texto_total = "\n".join(texto for _, texto in textos_paginas)
    if not _normalizar_texto(texto_total):
        raise ValueError("El PDF no tiene texto seleccionable. Puede ser un archivo escaneado.")

    tipo = _identificar_tipo(texto_total)
    numero_serial, fecha_recaudacion = _extraer_serial_fecha(texto_total)
    casilleros: dict[str, float] = {}
    evidencias: dict[str, dict[str, Any]] = {}
    for pagina_num, texto in textos_paginas:
        _extraer_casilleros_pagina(texto, tipo, pagina_num, casilleros, evidencias)

    if not casilleros:
        texto_mayus = texto_total.upper()
        if "COMPROBANTE" in texto_mayus and "ELECTRÓNICO PARA PAGO" in texto_mayus:
            raise ValueError(
                f"'{ruta.name}' es un comprobante electrónico de pago, no el formulario completo. "
                "Contiene número de serie y periodo, pero no incluye casilleros declarados. "
                "Carga el PDF completo de la declaración F104 correspondiente."
            )
        raise ValueError(
            f"No se encontraron casilleros con valores monetarios en '{ruta.name}'. "
            "El PDF puede usar una estructura nueva del SRI o tener el texto fragmentado."
        )

    return {
        "tipo": tipo,
        "periodo": _extraer_periodo(texto_total),
        "ruc": _extraer_ruc(texto_total),
        "razon_social": _extraer_razon_social(texto_total),
        "numero_serial": numero_serial,
        "fecha_recaudacion": fecha_recaudacion,
        "casilleros": casilleros,
        "evidencias": evidencias,
        "archivo": ruta.name,
    }


def validar_formulario(formulario: dict[str, Any], tolerancia: float = 0.01) -> list[str]:
    casilleros = formulario.get("casilleros", {})
    tipo = formulario.get("tipo")

    def val(codigo: str) -> float:
        return float(casilleros.get(codigo, 0.0) or 0.0)

    def existe_alguno(codigos: list[str]) -> bool:
        return any(codigo in casilleros for codigo in codigos)

    reglas = []
    if tipo == "F104":
        reglas = [
            ("409", ["401", "402", "403", "404", "405", "406", "407", "408"]),
            ("499", ["483", "484"]),
            ("859", ["699", "801"]),
            ("999", ["902", "903", "904"]),
        ]
    elif tipo == "F103":
        reglas = [
            ("499", ["399", "498"]),
            ("902", ["499", "-898"]),
            ("999", ["902", "903", "904"]),
        ]

    alertas = []
    for destino, componentes in reglas:
        if destino not in casilleros and not existe_alguno([c.strip("-") for c in componentes]):
            continue
        calculado = 0.0
        partes = []
        for componente in componentes:
            signo = -1 if componente.startswith("-") else 1
            codigo = componente.strip("-")
            calculado += signo * val(codigo)
            partes.append(f"{'-' if signo < 0 else '+'}{codigo}" if partes else componente)
        declarado = val(destino)
        if abs(declarado - calculado) > tolerancia:
            formula = " ".join(partes).replace("+", "+ ")
            alertas.append(
                f"{tipo} {formulario.get('periodo', {}).get('label', '')}: casillero {destino} "
                f"declara {declarado:,.2f}, pero {' '.join(componentes)} suma {calculado:,.2f}."
            )
    return alertas


def casilleros_a_dataframe(formularios: list[dict[str, Any]]) -> pd.DataFrame:
    filas = []
    for formulario in formularios:
        periodo = formulario.get("periodo", {})
        evidencias = formulario.get("evidencias", {})
        for casillero, valor in formulario.get("casilleros", {}).items():
            evidencia = evidencias.get(casillero, {})
            filas.append(
                {
                    "ARCHIVO": formulario.get("archivo", ""),
                    "FORMULARIO": formulario.get("tipo", ""),
                    "PERIODO": periodo.get("label", ""),
                    "NUMERO_SERIAL": formulario.get("numero_serial", ""),
                    "FECHA_RECAUDACION": formulario.get("fecha_recaudacion", ""),
                    "CASILLERO": casillero,
                    "VALOR_EXTRAIDO": float(valor),
                    "PAGINA": evidencia.get("pagina", ""),
                    "TEXTO_ORIGINAL": evidencia.get("texto_original", ""),
                }
            )
    if not filas:
        return pd.DataFrame()
    df = pd.DataFrame(filas)
    df["ORDEN_CASILLERO"] = pd.to_numeric(df["CASILLERO"], errors="coerce")
    df = df.sort_values(["FORMULARIO", "PERIODO", "ORDEN_CASILLERO", "CASILLERO"])
    return df.drop(columns=["ORDEN_CASILLERO"]).reset_index(drop=True)
