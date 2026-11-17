import json
from pathlib import Path
from typing import Any


RUTA_MAPEO_DEFAULT = Path(__file__).with_name("mapeo_default.json")
RUTA_MAPEO = RUTA_MAPEO_DEFAULT


def _normalizar_ids(mapeos: list[dict[str, Any]]) -> list[dict[str, Any]]:
    normalizados = []
    for indice, mapeo in enumerate(mapeos, start=1):
        item = dict(mapeo)
        item["id"] = int(item.get("id") or indice)
        item["casillero"] = str(item.get("casillero", "")).strip()
        item["cuenta_contable"] = str(item.get("cuenta_contable", "")).strip()
        item["formulario"] = str(item.get("formulario", "")).strip().upper()
        item["columna_libro"] = str(item.get("columna_libro", "MOVIMIENTO_NETO")).strip().upper()
        item["signo"] = int(item.get("signo", 1))
        item["desfase_meses"] = int(item.get("desfase_meses", 0))
        normalizados.append(item)
    return normalizados


def _cargar_desde(ruta: Path) -> list[dict[str, Any]]:
    if not ruta.exists():
        return []
    with ruta.open("r", encoding="utf-8") as archivo:
        datos = json.load(archivo)
    return _normalizar_ids(datos.get("mapeos", []))


def cargar_mapeo_default() -> list[dict[str, Any]]:
    return _cargar_desde(RUTA_MAPEO_DEFAULT)


def cargar_mapeo() -> list[dict[str, Any]]:
    mapeo_usuario = _cargar_desde(RUTA_MAPEO)
    return mapeo_usuario or cargar_mapeo_default()


def guardar_mapeo(mapeos: list[dict[str, Any]]) -> None:
    datos = {"mapeos": _normalizar_ids(mapeos)}
    RUTA_MAPEO.parent.mkdir(parents=True, exist_ok=True)
    with RUTA_MAPEO.open("w", encoding="utf-8") as archivo:
        json.dump(datos, archivo, ensure_ascii=False, indent=2)


def agregar_mapeo(nuevo: dict[str, Any]) -> None:
    mapeos = cargar_mapeo()
    siguiente_id = max([int(m.get("id", 0)) for m in mapeos] or [0]) + 1
    item = dict(nuevo)
    item["id"] = siguiente_id
    mapeos.append(item)
    guardar_mapeo(mapeos)


def eliminar_mapeo(id: int) -> None:
    mapeos = [mapeo for mapeo in cargar_mapeo() if int(mapeo.get("id", 0)) != int(id)]
    guardar_mapeo(mapeos)


def editar_mapeo(id: int, datos: dict[str, Any]) -> None:
    mapeos = []
    for mapeo in cargar_mapeo():
        if int(mapeo.get("id", 0)) == int(id):
            actualizado = dict(mapeo)
            actualizado.update(datos)
            mapeos.append(actualizado)
        else:
            mapeos.append(mapeo)
    guardar_mapeo(mapeos)
