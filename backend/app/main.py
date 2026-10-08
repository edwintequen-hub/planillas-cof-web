from pathlib import Path
from .zip import crear_zip_planillas
from .limpieza import limpiar_salida_unidad

import shutil
import uuid
import traceback
import time
from typing import Optional

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

from fastapi.responses import Response
from fastapi.templating import Jinja2Templates
from fastapi.staticfiles import StaticFiles



from .config import (
    PLANTILLA_EXCEL,
    INFO_EXCEL,
    UPLOAD_DIR,
    OUTPUT_DIR,
)

from .catalogo_nodos import (
    obtener_catalogo,
    actualizar_nodo,
)

from .configuracion_planillas import (
    obtener_configuracion,
    actualizar_incluir_vacios,
)


from .info import (
    obtener_unidades,
    obtener_servicios,
    obtener_terminal,
    obtener_unidad_por_servicio,
    obtener_unidad_por_servicio_fus,
)


from .generador import (
    generar_planillas,
    leer_fus,
)

from .anexo4 import (
    validar_unidad_anexo4,
)

from .anexo5 import (
    leer_anexo5,
    crear_indice_anexo5_ts,
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


            unidad = obtener_unidad_por_servicio_fus(

                servicio

            )


            if unidad:


                unidades_detectadas.add(

                    unidad

                )



        if not unidades_detectadas:


            return False, (

                "No se encontraron servicios "

                "vÃ¡lidos en INFO.xlsx."

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

@app.head("/")
async def inicio_head():
    return Response(status_code=200)


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


# ==========================================================
# ACTUALIZAR ANEXO 5 VIGENTE
# ==========================================================


# ==========================================================
# CONFIGURACION OPERACIONAL DE PLANILLAS
# ==========================================================



@app.get("/api/configuracion/nodos")
def consultar_catalogo_nodos(
    unidad: Optional[str] = None
):

    registros = obtener_catalogo()

    if unidad:
        unidad_normalizada = (
            str(unidad)
            .strip()
            .upper()
        )

        registros = [
            registro
            for registro in registros
            if registro.get(
                "unidad",
                ""
            ).upper() == unidad_normalizada
        ]

    return {
        "ok": True,
        "total": len(registros),
        "registros": registros,
    }


@app.post("/api/configuracion/nodos")
async def modificar_nodo(
    request: Request
):

    try:
        datos = await request.json()

    except Exception:

        return JSONResponse(
            status_code=400,
            content={
                "ok": False,
                "error": "JSON invalido.",
            }
        )

    unidad = str(
        datos.get(
            "unidad",
            ""
        )
    ).strip()

    codigo = str(
        datos.get(
            "codigo",
            ""
        )
    ).strip()

    nodo = str(
        datos.get(
            "nodo",
            ""
        )
    ).strip()

    if not unidad or not codigo or not nodo:

        return JSONResponse(
            status_code=400,
            content={
                "ok": False,
                "error":
                    "unidad, codigo y nodo son obligatorios.",
            }
        )

    try:

        registro = actualizar_nodo(
            unidad,
            codigo,
            nodo,
        )

    except ValueError as exc:

        return JSONResponse(
            status_code=404,
            content={
                "ok": False,
                "error": str(exc),
            }
        )

    return {
        "ok": True,
        "registro": registro,
    }


@app.get("/api/configuracion/planillas")
def consultar_configuracion_planillas():

    configuracion = obtener_configuracion()

    return {
        "ok": True,
        "configuracion": configuracion,
    }


@app.post("/api/configuracion/planillas")
async def guardar_configuracion_planillas(
    request: Request
):

    try:

        datos = await request.json()

    except Exception:

        return JSONResponse(
            status_code=400,
            content={
                "ok": False,
                "error": "JSON invalido."
            }
        )

    incluir_vacios = datos.get(
        "incluir_vacios"
    )

    if not isinstance(incluir_vacios, bool):

        return JSONResponse(
            status_code=400,
            content={
                "ok": False,
                "error":
                    "incluir_vacios debe ser true o false."
            }
        )

    configuracion = actualizar_incluir_vacios(
        incluir_vacios
    )

    return {
        "ok": True,
        "configuracion": configuracion,
    }


# ==========================================================
# ESTADO ANEXO 5 VIGENTE
# ==========================================================

@app.get("/api/anexo5/estado")
def estado_anexo5():

    from datetime import datetime

    catalogos_dir = (
        Path(__file__).resolve().parent.parent
        / "catalogos"
    )

    resultado = {}

    for unidad in ("U8", "U9"):

        archivo = (
            catalogos_dir
            / f"Anexo5_{unidad}.xlsx"
        )

        if archivo.exists():

            fecha = datetime.fromtimestamp(
                archivo.stat().st_mtime
            )

            resultado[unidad] = {
                "cargado": True,
                "archivo": archivo.name,
                "ultima_actualizacion":
                    fecha.strftime("%d-%m-%Y %H:%M"),
            }

        else:

            resultado[unidad] = {
                "cargado": False,
                "archivo": None,
                "ultima_actualizacion": None,
            }

    return {
        "ok": True,
        "catalogos": resultado,
    }



# ==========================================================
# RESUMEN E HISTORIAL ANEXO 5
# ==========================================================

@app.get("/api/anexo5/resumen")
def resumen_anexo5():

    from datetime import datetime

    catalogos_dir = (
        Path(__file__).resolve().parent.parent
        / "catalogos"
    )

    historico_dir = (
        catalogos_dir
        / "historico_anexo5"
    )

    resultado = {}

    for unidad in ("U8", "U9"):

        actual = (
            catalogos_dir
            / f"Anexo5_{unidad}.xlsx"
        )

        actual_info = None

        if actual.exists():

            fecha = datetime.fromtimestamp(
                actual.stat().st_mtime
            )

            actual_info = {
                "archivo": actual.name,
                "fecha":
                    fecha.strftime("%d-%m-%Y %H:%M"),
                "cargado": True,
            }

        historial = []

        if historico_dir.exists():

            archivos = sorted(
                historico_dir.glob(
                    f"Anexo5_{unidad}_*.xlsx"
                ),
                key=lambda p: p.stat().st_mtime,
                reverse=True
            )

            for archivo in archivos[:30]:

                fecha = datetime.fromtimestamp(
                    archivo.stat().st_mtime
                )

                historial.append({
                    "archivo": archivo.name,
                    "fecha":
                        fecha.strftime("%d-%m-%Y %H:%M"),
                })

        resultado[unidad] = {
            "actual": actual_info,
            "historial": historial,
        }

    return {
        "ok": True,
        "unidades": resultado,
    }


@app.post("/api/anexo5/actualizar")
async def actualizar_anexo5(
    archivo: UploadFile = File(...),
    unidad: str = Form(...)
):

    unidad_normalizada = str(
        unidad or ""
    ).strip().upper()

    equivalencias = {
        "U8": "U8",
        "8": "U8",
        "ALFA": "U8",
        "ALFAU8": "U8",

        "U9": "U9",
        "9": "U9",
        "OMEGA": "U9",
        "OMEGAU9": "U9",
    }

    unidad_tecnica = equivalencias.get(
        unidad_normalizada
    )

    if unidad_tecnica not in ("U8", "U9"):

        return JSONResponse(
            status_code=400,
            content={
                "ok": False,
                "error":
                    "Unidad invalida. Debe seleccionar U8 o U9."
            }
        )

    if (
        archivo is None
        or not archivo.filename
    ):

        return JSONResponse(
            status_code=400,
            content={
                "ok": False,
                "error":
                    "Debe seleccionar un archivo Anexo 5."
            }
        )

    if not archivo.filename.lower().endswith(".xlsx"):

        return JSONResponse(
            status_code=400,
            content={
                "ok": False,
                "error":
                    "El Anexo 5 debe ser un archivo .xlsx."
            }
        )

    catalogos_dir = (
        Path(__file__).resolve().parent.parent
        / "catalogos"
    )

    catalogos_dir.mkdir(
        parents=True,
        exist_ok=True
    )

    destino = (
        catalogos_dir
        / f"Anexo5_{unidad_tecnica}.xlsx"
    )

    historico_dir = (
        catalogos_dir
        / "historico_anexo5"
    )

    historico_dir.mkdir(
        parents=True,
        exist_ok=True
    )

    temporal = (
        UPLOAD_DIR
        / f"anexo5_validacion_{uuid.uuid4().hex}.xlsx"
    )

    try:

        # ----------------------------------------------
        # Guardar primero como temporal
        # ----------------------------------------------

        with open(temporal, "wb") as buffer:

            shutil.copyfileobj(
                archivo.file,
                buffer
            )

        # ----------------------------------------------
        # VALIDACION REAL CON MOTOR ANEXO 5
        # ----------------------------------------------

        registros = leer_anexo5(
            temporal
        )

        if not registros:

            raise ValueError(
                "El archivo no contiene registros "
                "vigentes validos para IP."
            )

        indice = crear_indice_anexo5_ts(
            registros
        )

        if not indice:

            raise ValueError(
                "No fue posible construir el indice "
                "tecnico del Anexo 5."
            )

        # ----------------------------------------------
        # VALIDAR QUE EL ARCHIVO CORRESPONDA A LA UNIDAD
        # ----------------------------------------------

        unidades_archivo = {
            str(
                registro.get("unidad") or ""
            ).strip().upper()
            for registro in registros
            if registro.get("unidad") is not None
        }

        unidades_normalizadas = set()

        for valor in unidades_archivo:

            valor_limpio = (
                valor
                .replace(" ", "")
            )

            if valor_limpio in (
                "U8",
                "8",
                "ALFA",
                "ALFAU8",
            ):
                unidades_normalizadas.add("U8")

            elif valor_limpio in (
                "U9",
                "9",
                "OMEGA",
                "OMEGAU9",
            ):
                unidades_normalizadas.add("U9")

        if (
            unidades_normalizadas
            and unidad_tecnica not in unidades_normalizadas
        ):

            raise ValueError(
                f"El archivo seleccionado no corresponde "
                f"a la unidad {unidad_tecnica}. "
                f"Unidades detectadas: "
                f"{sorted(unidades_normalizadas)}"
            )

        # ----------------------------------------------
        # BACKUP DEL CATALOGO VIGENTE ANTERIOR
        # ----------------------------------------------

        backup_creado = None

        if destino.exists():

            marca = time.strftime(
                "%Y%m%d_%H%M%S"
            )

            backup_creado = (
                historico_dir
                / (
                    f"Anexo5_{unidad_tecnica}_"
                    f"{marca}.xlsx"
                )
            )

            shutil.copy2(
                destino,
                backup_creado
            )

        # ----------------------------------------------
        # REEMPLAZAR SOLO DESPUES DE VALIDAR
        # ----------------------------------------------

        shutil.copy2(
            temporal,
            destino
        )

        return {
            "ok": True,
            "unidad": unidad_tecnica,
            "archivo_original": archivo.filename,
            "catalogo_vigente": destino.name,
            "registros_validos": len(registros),
            "claves_indice": len(indice),
            "backup_anterior":
                backup_creado.name
                if backup_creado
                else None,
            "mensaje":
                (
                    f"Anexo 5 {unidad_tecnica} actualizado "
                    f"correctamente y queda vigente hasta "
                    f"la proxima actualizacion DTPM."
                )
        }

    except Exception as exc:

        traceback.print_exc()

        return JSONResponse(
            status_code=400,
            content={
                "ok": False,
                "error":
                    f"No se actualizo el Anexo 5: {exc}"
            }
        )

    finally:

        try:

            if temporal.exists():
                temporal.unlink()

        except Exception:

            pass


@app.post("/generar")
async def generar(

    archivos: Optional[list[UploadFile]] = File(default=None),

    anexo4: Optional[UploadFile] = File(default=None),

    unidad: str = Form(...),

    servicio: str = Form(...),

    tipo_dia: str = Form(...)

):


    print("ENTRO AL ENDPOINT GENERAR")

    # Normalizar entradas
    archivos = [a for a in (archivos or []) if a and a.filename]
    tiene_fus = len(archivos) > 0
    tiene_anexo4 = anexo4 is not None and bool(anexo4.filename)

    print("==============================")
    print("INICIO GENERACION")
    print("Unidad:", unidad)
    print("Servicio:", servicio)
    print("Tipo DÃ­a:", tipo_dia)
    print("Tiene FUS:", tiene_fus)
    print("Cantidad FUS:", len(archivos))
    print("Tiene Anexo4:", tiene_anexo4)
    if tiene_anexo4:
        print("Archivo:", anexo4.filename)
    print("==============================")

    inicio = time.time()
    archivos_guardados = []

    # ==================================
    # VALIDAR ORIGEN
    # ==================================

    if tiene_fus and tiene_anexo4:
        return JSONResponse(
            status_code=400,
            content={
                "ok": False,
                "error": "Seleccione Archivos FUS o Archivo Anexo 4, no ambos."
            }
        )

    if not tiene_fus and not tiene_anexo4:
        return JSONResponse(
            status_code=400,
            content={
                "ok": False,
                "error": "Debe seleccionar Archivos FUS o un Archivo Anexo 4."
            }
        )

    try:


        # ==================================
        # GUARDAR ARCHIVOS TEMPORALES
        # ==================================

        if tiene_fus:
            archivos_subidos = archivos
        else:
            archivos_subidos = [anexo4]

        for archivo in archivos_subidos:

            nombre = f"{uuid.uuid4()}.xlsx"

            ruta = UPLOAD_DIR / nombre

            with open(ruta, "wb") as buffer:

                shutil.copyfileobj(
                    archivo.file,
                    buffer
                )

            archivos_guardados.append(ruta)



        print("ARCHIVOS TEMPORALES:")


        for archivo in archivos_guardados:

            print(archivo)



        # ==================================
        # VALIDAR UNIDAD
        # ==================================

        if tiene_fus:

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

        else:

            valido, mensaje = validar_unidad_anexo4(

                archivos_guardados[0],

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

        configuracion_planillas = obtener_configuracion()

        incluir_vacios = configuracion_planillas.get(
            "incluir_vacios",
            True
        )

        print("==============================")
        print("CONFIGURACION ACTIVA")
        print("Incluir vacios:", incluir_vacios)
        print("==============================")

        resultado = generar_planillas(

            archivos_guardados,

            PLANTILLA_EXCEL,

            INFO_EXCEL,

            OUTPUT_DIR,

            unidad,

            servicio,

            tipo_dia,

            usar_anexo4=tiene_anexo4,

            incluir_vacios=incluir_vacios

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
