from pathlib import Path
import shutil


def limpiar_salida_unidad(
    output_dir: Path,
    unidad: str,
) -> None:

    carpeta = output_dir / unidad

    print("=" * 30)
    print("LIMPIANDO SALIDA")
    print("UNIDAD:", unidad)
    print("CARPETA:", carpeta)
    print("=" * 30)

    if carpeta.exists():

        for elemento in carpeta.iterdir():

            try:

                if elemento.is_dir():

                    shutil.rmtree(elemento)

                else:

                    elemento.unlink()

            except Exception as e:

                print(f"No fue posible eliminar: {elemento}")
                print(e)

    carpeta.mkdir(
        parents=True,
        exist_ok=True,
    )

    print("LIMPIEZA TERMINADA")
    print("=" * 30)