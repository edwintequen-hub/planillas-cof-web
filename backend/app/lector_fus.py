"""
Lector de archivos FUS.

Responsabilidades:
- Validar el archivo.
- Leer el Excel.
- Validar columnas.
- Renombrar columnas.
- Limpiar datos.
- Retornar un DataFrame estandarizado.
"""

from __future__ import annotations

from pathlib import Path

import pandas as pd

from .logger import obtener_logger

logger = obtener_logger(__name__)


# ==========================================================
# COLUMNAS OBLIGATORIAS DEL FUS
# ==========================================================

COLUMNAS_OBLIGATORIAS = {
    "Evento",
    "Tipo",
    "Inicio",
    "De",
    "Fin",
    "A",
    "Duración",
    "Bus",
    "Tipo.1",
    "Línea:",
    "km",
    "Tipo Mapeado",
    "Sentido",
}


# ==========================================================
# RENOMBRE DE COLUMNAS
# ==========================================================

COLUMNAS = {
    "Evento": "evento",
    "Tipo": "tipo_evento",
    "Inicio": "inicio",
    "De": "origen",
    "Fin": "fin",
    "A": "destino",
    "Duración": "duracion",
    "Bus": "bus",
    "Tipo.1": "tipo_bus",
    "Línea:": "linea",
    "km": "km",
    "Tipo Mapeado": "tipo_dia",
    "Sentido": "sentido",
}