import shutil


def limpiar_salida_unidad(
    output_dir,
    unidad
):

    carpeta = output_dir / unidad


    print("==============================")
    print("LIMPIANDO SALIDA")
    print("UNIDAD:", unidad)
    print("CARPETA:", carpeta)
    print("==============================")


    if carpeta.exists():


        for elemento in carpeta.iterdir():


            if elemento.is_dir():

                shutil.rmtree(
                    elemento
                )


            else:

                elemento.unlink()



    carpeta.mkdir(

        parents=True,

        exist_ok=True

    )


    print("LIMPIEZA TERMINADA")
    print("==============================")