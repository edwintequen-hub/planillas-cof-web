from pathlib import Path
from datetime import datetime
import zipfile


def crear_zip_planillas(
    output_dir,
    unidad
):

    fecha = datetime.now().strftime(
        "%Y%m%d"
    )


    nombre_zip = (
        f"PlanillasCof_{unidad}_{fecha}.zip"
    )


    ruta_zip = (
        output_dir / nombre_zip
    )


    carpeta_unidad = (
        output_dir / unidad
    )


    if not carpeta_unidad.exists():

        raise FileNotFoundError(
            f"No existe carpeta: {carpeta_unidad}"
        )


    with zipfile.ZipFile(
        ruta_zip,
        "w",
        zipfile.ZIP_DEFLATED
    ) as zipf:


        for archivo in carpeta_unidad.rglob("*"):


            if archivo.is_file():

                ruta_relativa = archivo.relative_to(
                    output_dir
                )


                zipf.write(
                    archivo,
                    ruta_relativa
                )


    return nombre_zip