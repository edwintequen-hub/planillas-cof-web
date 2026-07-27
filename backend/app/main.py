from pathlib import Path
from .zip import crear_zip_planillas
from .limpieza import limpiar_salida_unidad

import shutil
import uuid
import traceback
import time

from fastapi import (
    FastAPI,
    Request,
    UploadFile,
    File,
    Form,
)

from fastapi.responses import (
    HTMLResponse,
    JSONResponse,
    FileResponse,
)

from fastapi.templating import Jinja2Templates
from fastapi.staticfiles import StaticFiles



from .config import (
    PLANTILLA_EXCEL,
    INFO_EXCEL,
    UPLOAD_DIR,
    OUTPUT_DIR,
)


from .info import (
    obtener_unidades,
    obtener_servicios,
    obtener_terminal,
    obtener_unidad_por_servicio,
)


from .generador import (
    generar_planillas,
    leer_fus,
)



# ==========================================================
# APP
# ==========================================================

app = FastAPI(

    title="Planillas COF Web",

    version="2.0"

)



# ==========================================================
# DIRECTORIOS
# ==========================================================

BASE_DIR = Path(__file__).resolve().parents[2]

FRONTEND_DIR = BASE_DIR / "frontend"


templates = Jinja2Templates(

    directory=str(FRONTEND_DIR)

)



print("==============================")
print("BASE:", BASE_DIR)
print("FRONTEND:", FRONTEND_DIR)
print(
    "INDEX:",
    (FRONTEND_DIR / "index.html").exists()
)
print("==============================")



# ==========================================================
# STATIC
# ==========================================================

STATIC_DIR = FRONTEND_DIR / "static"

print("==============================")
print("STATIC:", STATIC_DIR)
print("EXISTE:", STATIC_DIR.exists())
print("LOGO:", (STATIC_DIR / "img" / "logo_metropol.png").exists())
print("==============================")


if STATIC_DIR.exists():

    app.mount(

        "/static",

        StaticFiles(

            directory=str(STATIC_DIR)

        ),

        name="static"

    )



# ==========================================================
# CARPETAS
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
# VALIDAR UNIDAD SEGUN SERVICIOS DEL FUS
# ==========================================================

def validar_unidad_fus(

    archivos,

    unidad_seleccionada

):


    try:


        registros = leer_fus(

            archivos

        )


        servicios = set()



        for registro in registros:


            servicio = registro.get(
                "servicio"
            )


            if servicio:


                servicios.add(

                    str(servicio).strip()

                )



        unidades_detectadas = set()


        for servicio in servicios:


            unidad = obtener_unidad_por_servicio(

                servicio

            )


            if unidad:


                unidades_detectadas.add(

                    unidad

                )



        if not unidades_detectadas:


            return False, (

                "No se encontraron servicios "

                "válidos en INFO.xlsx."

            )



        if unidad_seleccionada not in unidades_detectadas:


            return False, (

                "El FUS no corresponde a la "

                "unidad seleccionada. "

                f"Unidad seleccionada: "

                f"{unidad_seleccionada}. "

                f"Unidad detectada: "

                f"{', '.join(unidades_detectadas)}. "

                f"Servicios encontrados: "

                f"{', '.join(servicios)}"

            )



        return True, None



    except Exception as e:


        return False, str(e)
    # ==========================================================
# PAGINA PRINCIPAL
# ==========================================================

@app.get(
    "/",
    response_class=HTMLResponse
)
async def inicio(
    request: Request
):

    return templates.TemplateResponse(

        request=request,

        name="index.html",

        context={

            "request": request

        }

    )



# ==========================================================
# API UNIDADES
# ==========================================================

@app.get("/api/unidades")
def api_unidades():

    return {

        "unidades":

            obtener_unidades()

    }



# ==========================================================
# API SERVICIOS
# ==========================================================

@app.get("/api/servicios/{unidad}")
def api_servicios(

    unidad: str

):

    return {

        "servicios":

            obtener_servicios(

                unidad

            )

    }



# ==========================================================
# API TERMINAL
# ==========================================================

@app.get("/api/terminal/{unidad}/{servicio}")
def api_terminal(

    unidad: str,

    servicio: str

):

    return {

        "terminal":

            obtener_terminal(

                unidad,

                servicio

            )

    }



# ==========================================================
# GENERAR PLANILLAS
# ==========================================================

@app.post("/generar")
async def generar(

    archivos: list[UploadFile] = File(...),

    unidad: str = Form(...),

    servicio: str = Form(...),

    tipo_dia: str = Form(...)

):


    print("ENTRO AL ENDPOINT GENERAR")


    print("==============================")
    print("INICIO GENERACION")
    print("Unidad:", unidad)
    print("Servicio:", servicio)
    print("Tipo Día:", tipo_dia)
    print("Cantidad archivos:", len(archivos))
    print("==============================")


    inicio = time.time()


    archivos_guardados = []



    try:


        # ==================================
        # GUARDAR FUS TEMPORALES
        # ==================================

        for archivo in archivos:


            nombre = (

                str(uuid.uuid4())

                +

                ".xlsx"

            )


            ruta = UPLOAD_DIR / nombre



            with open(

                ruta,

                "wb"

            ) as buffer:


                shutil.copyfileobj(

                    archivo.file,

                    buffer

                )


            archivos_guardados.append(

                ruta

            )



        print("ARCHIVOS TEMPORALES:")


        for archivo in archivos_guardados:

            print(archivo)



        # ==================================
        # VALIDAR UNIDAD
        # ==================================

        valido, mensaje = validar_unidad_fus(

            archivos_guardados,

            unidad

        )


        if not valido:


            return JSONResponse(

                status_code=400,

                content={

                    "ok": False,

                    "error": mensaje

                }

            )
        # ==================================
        # LIMPIAR ARCHIVOS ANTERIORES
        # ==================================

        limpiar_salida_unidad(

            OUTPUT_DIR,

            unidad

        )


        # ==================================
        # GENERAR
        # ==================================

        resultado = generar_planillas(

            archivos_guardados,

            PLANTILLA_EXCEL,

            INFO_EXCEL,

            OUTPUT_DIR,

            unidad,

            servicio,

            tipo_dia

        )

        nombre_zip = crear_zip_planillas(

            OUTPUT_DIR,

            unidad

        )

        print("==============================")
        print("ZIP GENERADO:")
        print(nombre_zip)
        print("==============================")

        fin = time.time()


        tiempo = round(

            fin - inicio,

            2

        )



        print("==============================")
        print("RESULTADO:")
        print(resultado)
        print("TIEMPO:", tiempo, "segundos")
        print("==============================")



        return {


            "ok": True,


            "mensaje":

                "Proceso terminado",



            "total":

                resultado.get(

                    "total",

                    0

                ),



            "laboral":

                resultado.get(

                    "laboral",

                    0

                ),



            "sabado":

                resultado.get(

                    "sabado",

                    0

                ),



            "domingo":

                resultado.get(

                    "domingo",

                    0

                ),



            "tiempo":

                tiempo,

            "zip":

                nombre_zip


        }



    except Exception as e:


        print("==============================")
        print("ERROR GENERADOR")
        print("==============================")


        traceback.print_exc()



        return JSONResponse(

            status_code=500,

            content={

                "ok": False,

                "error": str(e)

            }

        )



    finally:


        for archivo in archivos_guardados:


            if archivo.exists():


                archivo.unlink()

# ==========================================================
# DESCARGAR ZIP
# ==========================================================

@app.get("/descargar/{nombre_zip}")
def descargar_zip(
    nombre_zip: str
):


    ruta = OUTPUT_DIR / nombre_zip


    if not ruta.exists():

        return JSONResponse(

            status_code=404,

            content={

                "error":

                    "Archivo no encontrado"

            }

        )


    return FileResponse(

        path=ruta,

        filename=nombre_zip,

        media_type="application/zip"

    )

# ==========================================================
# HEALTH
# ==========================================================

@app.get("/health")
def health():

    return {

        "estado":

            "OK"

    }