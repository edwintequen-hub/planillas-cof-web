from pathlib import Path


# ==========================================================
# RUTA BASE DEL PROYECTO
# ==========================================================

BASE_DIR = Path(__file__).resolve().parent.parent



# ==========================================================
# CARPETAS
# ==========================================================

CARPETA_PLANTILLA = BASE_DIR / "plantilla"

CARPETA_UPLOADS = BASE_DIR / "uploads"

CARPETA_SALIDAS = BASE_DIR / "salidas"



# ==========================================================
# ARCHIVOS EXCEL
# ==========================================================

PLANTILLA_EXCEL = CARPETA_PLANTILLA / "Plantilla.xlsx"

INFO_EXCEL = CARPETA_PLANTILLA / "INFO.xlsx"



# ==========================================================
# COMPATIBILIDAD
# ==========================================================

UPLOAD_DIR = CARPETA_UPLOADS

OUTPUT_DIR = CARPETA_SALIDAS



# ==========================================================
# CREAR CARPETAS
# ==========================================================

UPLOAD_DIR.mkdir(
    parents=True,
    exist_ok=True
)


OUTPUT_DIR.mkdir(
    parents=True,
    exist_ok=True
)



# ==========================================================
# VALIDAR ARCHIVOS BASE
# ==========================================================

print("==============================")
print("CONFIG COF WEB")
print("==============================")

print(
    "PLANTILLA:",
    PLANTILLA_EXCEL
)

print(
    "EXISTE:",
    PLANTILLA_EXCEL.exists()
)


print(
    "INFO:",
    INFO_EXCEL
)

print(
    "EXISTE:",
    INFO_EXCEL.exists()
)


print("==============================")



if not PLANTILLA_EXCEL.exists():

    raise FileNotFoundError(
        f"No existe Plantilla.xlsx: {PLANTILLA_EXCEL}"
    )



if not INFO_EXCEL.exists():

    raise FileNotFoundError(
        f"No existe INFO.xlsx: {INFO_EXCEL}"
    )