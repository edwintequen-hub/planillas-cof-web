"""
Sistema de logging del proyecto Planillas COF Web.

Autor: Proyecto Planillas COF Web
Python: 3.14+
"""

from __future__ import annotations

import logging
from pathlib import Path


# ==========================================================
# CARPETA DE LOGS
# ==========================================================

BASE_DIR = Path(__file__).resolve().parent.parent

LOG_DIR = BASE_DIR / "logs"
LOG_DIR.mkdir(parents=True, exist_ok=True)

LOG_FILE = LOG_DIR / "planillas.log"


# ==========================================================
# LOGGER
# ==========================================================

def obtener_logger(nombre: str) -> logging.Logger:
    """
    Retorna un logger configurado para el proyecto.

    Parameters
    ----------
    nombre : str
        Nombre del módulo.

    Returns
    -------
    logging.Logger
    """

    logger = logging.getLogger(nombre)

    if logger.handlers:
        return logger

    logger.setLevel(logging.INFO)

    formato = logging.Formatter(
        "%(asctime)s | %(levelname)-8s | %(name)s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S"
    )

    # Archivo
    file_handler = logging.FileHandler(
        LOG_FILE,
        encoding="utf-8"
    )
    file_handler.setFormatter(formato)

    # Consola
    console_handler = logging.StreamHandler()
    console_handler.setFormatter(formato)

    logger.addHandler(file_handler)
    logger.addHandler(console_handler)

    logger.propagate = False

    return logger