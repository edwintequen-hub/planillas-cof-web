import json
from pathlib import Path

from .config import BASE_DIR


# ==========================================================
# CONFIGURACION OPERACIONAL DE PLANILLAS
# ==========================================================

CARPETA_CONFIGURACION = BASE_DIR / "configuracion"

ARCHIVO_CONFIGURACION_PLANILLAS = (
    CARPETA_CONFIGURACION / "configuracion_planillas.json"
)

CONFIGURACION_PREDETERMINADA = {
    "incluir_vacios": True
}


def asegurar_configuracion():

    CARPETA_CONFIGURACION.mkdir(
        parents=True,
        exist_ok=True
    )

    if not ARCHIVO_CONFIGURACION_PLANILLAS.exists():

        guardar_configuracion(
            CONFIGURACION_PREDETERMINADA
        )


def obtener_configuracion():

    asegurar_configuracion()

    try:

        with open(
            ARCHIVO_CONFIGURACION_PLANILLAS,
            "r",
            encoding="utf-8"
        ) as archivo:

            datos = json.load(archivo)

    except (OSError, json.JSONDecodeError):

        datos = {}

    configuracion = (
        CONFIGURACION_PREDETERMINADA.copy()
    )

    configuracion.update(
        datos
        if isinstance(datos, dict)
        else {}
    )

    configuracion["incluir_vacios"] = bool(
        configuracion.get(
            "incluir_vacios",
            True
        )
    )

    return configuracion


def guardar_configuracion(configuracion):

    CARPETA_CONFIGURACION.mkdir(
        parents=True,
        exist_ok=True
    )

    datos = (
        CONFIGURACION_PREDETERMINADA.copy()
    )

    if isinstance(configuracion, dict):

        datos.update(configuracion)

    datos["incluir_vacios"] = bool(
        datos.get(
            "incluir_vacios",
            True
        )
    )

    temporal = (
        ARCHIVO_CONFIGURACION_PLANILLAS
        .with_suffix(".tmp")
    )

    with open(
        temporal,
        "w",
        encoding="utf-8"
    ) as archivo:

        json.dump(
            datos,
            archivo,
            ensure_ascii=False,
            indent=4
        )

    temporal.replace(
        ARCHIVO_CONFIGURACION_PLANILLAS
    )

    return datos


def actualizar_incluir_vacios(incluir_vacios):

    configuracion = obtener_configuracion()

    configuracion["incluir_vacios"] = bool(
        incluir_vacios
    )

    return guardar_configuracion(
        configuracion
    )
