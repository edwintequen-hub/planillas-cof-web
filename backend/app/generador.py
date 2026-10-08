from pathlib import Path
from copy import copy
from datetime import datetime, time as dt_time

from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

from .info import cargar_info
from .catalogo_nodos import obtener_catalogo
from .anexo4 import leer_anexo4
from .anexo5 import (
    leer_anexo5,
    crear_indice_anexo5_ts,
    buscar_medicion_ip_ts,
)

# ==========================================================
# NORMALIZAR TEXTO
# ==========================================================

def normaliza(valor):

    if valor is None:

        return ""

    return (
        str(valor)
        .strip()
        .upper()
    )



# ==========================================================
# SERVICIO PURO
#
# F30n -> F30N
# F30N-F30N -> F30N
# ==========================================================

def servicio_puro(valor):

    texto = normaliza(valor)


    if "-" in texto:

        texto = texto.split("-")[0]


    return texto.strip()



# ==========================================================
# FORMATO TIPO DIA
# ==========================================================

def formato_tipo_dia(valor):

    texto = normaliza(valor)


    if texto == "LABORAL":

        return "Laboral"



    if texto in (
        "SABADO",
        "SÁBADO"
    ):

        return "Sabado"



    if texto == "DOMINGO":

        return "Domingo"



    return str(valor)




# ==========================================================
# RUTA CATALOGO ANEXO 5 POR UNIDAD
# ==========================================================

def unidad_tecnica(unidad):

    texto = normaliza(unidad)

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

    return equivalencias.get(
        texto,
        texto
    )


def unidad_empresa(unidad):

    tecnica = unidad_tecnica(
        unidad
    )

    if tecnica == "U8":
        return "Alfa"

    if tecnica == "U9":
        return "Omega"

    return str(unidad).strip()


def ruta_catalogo_anexo5(unidad):

    tecnica = unidad_tecnica(
        unidad
    )

    return (
        Path(__file__).resolve().parent.parent
        / "catalogos"
        / f"Anexo5_{tecnica}.xlsx"
    )


# ==========================================================
# COLORES EVENTOS
# ==========================================================

COLOR_VPA = PatternFill(
    fill_type="solid",
    start_color="C6E0B4",
    end_color="C6E0B4"
)

COLOR_VEX = PatternFill(
    fill_type="solid",
    start_color="F4B183",
    end_color="F4B183"
)

# ==========================================================
# COLORES INDICADOR ANEXO 5
# ==========================================================

COLOR_IP = PatternFill(
    fill_type="solid",
    start_color="4472C4",
    end_color="4472C4"
)

COLOR_IE = PatternFill(
    fill_type="solid",
    start_color="C00000",
    end_color="C00000"
)

FUENTE_BLANCA = Font(
    color="FFFFFF",
    bold=True
)

# ==========================================================
# LEER FUS
#
# Columnas FUS:
#
# B  Tipo
# C  Inicio
# E  Fin
# J  Tipo Bus
# K  Servicio
# R  Tipo Mapeado
# T  Sentido
#
# Solo EXP
# ==========================================================

# ==========================================================
# NORMALIZAR HORA PARA ORDENAR
# ==========================================================
def hora_excel(valor):

    if valor is None:
        return None

    if isinstance(valor, dt_time):
        return valor

    if isinstance(valor, datetime):
        return valor.time()

    texto = str(valor).strip()

    if not texto:
        return None

    for formato in (
        "%H:%M:%S",
        "%H:%M"
    ):

        try:

            return datetime.strptime(
                texto,
                formato
            ).time()

        except ValueError:
            pass

    raise ValueError(
        f"Hora Anexo 5 no reconocida: {valor}"
    )


def hora_orden(valor):

    if valor is None:
        return ""

    if hasattr(valor, "strftime"):
        return valor.strftime("%H:%M:%S")

    return str(valor)


# ==========================================================
# LEER FUS
# ==========================================================
def leer_fus(
    archivos_fus
):


    registros = []


    for archivo in archivos_fus:


        print("==============================")
        print("LEYENDO FUS:")
        print(archivo)
        print("==============================")


        wb = load_workbook(
            archivo,
            data_only=True,
            read_only=True
        )


        try:

            ws = wb.active


            for fila in ws.iter_rows(
                min_row=2,
                values_only=True
            ):


                def valor(columna):

                    indice = columna - 1

                    if indice >= len(fila):

                        return None

                    return fila[indice]


                tipo = normaliza(
                    valor(2)
                )


                if tipo not in ("EXP", "VPA", "VEX"):

                    continue


                registro = {

                    "tipo": tipo,

                    "evento":
                        valor(1),

                    "hora":
                        valor(3),

                    "fin":
                        valor(5),

                    "tipo_bus":
                        valor(10),

                    "servicio":
                        servicio_puro(
                            valor(11)
                        ),

                    "linea":
                        valor(11),

                    "tipo_dia":
                        formato_tipo_dia(
                            valor(18)
                        ),

                    "sentido":
                        valor(20)

                }


                registros.append(
                    registro
                )


        finally:

            wb.close()


    registros.sort(

        key=lambda x:(
            x["servicio"] or "",
            x["sentido"] or "",
            hora_orden(x["hora"])
        )

    )


    print("==============================")
    print(
        "TOTAL EXP:",
        len(registros)
    )
    print("==============================")


    return registros


# ==========================================================
# LIMPIAR BLOQUE
# Borra datos y formato de un rango de columnas
# ==========================================================

def limpiar_bloque(ws, fila_inicio, col_inicio, col_fin):

    ULTIMA_FILA_PLANTILLA = 555

    FILA_MODELO = 557

    for fila in range(fila_inicio, ULTIMA_FILA_PLANTILLA + 1):

        for col in range(col_inicio, col_fin + 1):

            celda = ws.cell(fila, col)

            modelo = ws.cell(FILA_MODELO, col)

            celda.value = None

            celda._style = copy(modelo._style)

        
# ==========================================================
# RESOLVER NODO POR CODIGO TS
# ==========================================================

def resolver_nodo_catalogo(unidad, codigo_ts, sentido=None):

    unidad_buscar = normaliza(unidad)

    codigo_original = " ".join(
        str(codigo_ts or "").strip().upper().split()
    )

    if not codigo_original:
        return ""

    sentido_texto = str(sentido or "").strip()

    if sentido_texto == "1":
        sufijo_sentido = "00I"
    elif sentido_texto == "2":
        sufijo_sentido = "00R"
    else:
        sufijo_sentido = ""

    # ------------------------------------------------------
    # Normalizacion para relacionar INFO con catalogo.
    #
    # Ejemplos:
    #   963E -> T963 E0
    #   932C -> T932 C0
    #   945E -> T945 E0
    #   967  -> T967
    #
    # No modifica codigo_ts_a5 utilizado por Anexo 5.
    # Solo se usa para localizar el nodo de cabecera.
    # ------------------------------------------------------

    compacto = (
        codigo_original
        .replace(" ", "")
        .replace("-", "")
    )

    # Si ya viene como codigo completo del catalogo,
    # primero intentamos coincidencia exacta.
    catalogo = obtener_catalogo()

    for registro in catalogo:

        unidad_catalogo = normaliza(
            registro.get("unidad")
        )

        codigo_catalogo = " ".join(
            str(
                registro.get("codigo") or ""
            ).strip().upper().split()
        )

        if (
            unidad_catalogo == unidad_buscar
            and codigo_catalogo == codigo_original
        ):
            return str(
                registro.get("nodo") or ""
            ).strip()

    # ------------------------------------------------------
    # Codigo proveniente de INFO/FUS.
    # ------------------------------------------------------

    import re

    m = re.fullmatch(
        r"T?(\d+)([A-Z])?",
        compacto
    )

    if not m:
        return ""

    numero = m.group(1)
    variante = m.group(2)

    if variante:
        base_catalogo = (
            f"T{numero} {variante}0"
        )
    else:
        base_catalogo = (
            f"T{numero}"
        )

    if sufijo_sentido:
        codigo_objetivo = (
            f"{base_catalogo} {sufijo_sentido}"
        )
    else:
        codigo_objetivo = base_catalogo

    codigo_objetivo = " ".join(
        codigo_objetivo.upper().split()
    )

    for registro in catalogo:

        unidad_catalogo = normaliza(
            registro.get("unidad")
        )

        codigo_catalogo = " ".join(
            str(
                registro.get("codigo") or ""
            ).strip().upper().split()
        )

        if (
            unidad_catalogo == unidad_buscar
            and codigo_catalogo == codigo_objetivo
        ):
            return str(
                registro.get("nodo") or ""
            ).strip()

    return ""


# ==========================================================
# CREAR PLANILLA
#
# Mantiene estructura de Plantilla.xlsx
# ==========================================================

def crear_planilla(

    registros,

    plantilla,

    salida,

    unidad,

    servicio,

    tipo_dia,

    terminal,

    indice_anexo5=None,

    usar_anexo4=False

):
    import time

    inicio_planilla = time.time()

    wb = load_workbook(
        plantilla
    )


    ws = wb["Planilla"]




    # ======================================================
    # ENCABEZADOS
    # ======================================================

    # ======================================================
    # NODO CABECERA POR SENTIDO Y CODIGO TS REAL
    # ======================================================

    codigo_ts_ida = ""
    codigo_ts_regreso = ""

    print("\n=== DEBUG NODOS - REGISTROS RECIBIDOS ===")
    print("Servicio:", servicio)
    print("Unidad:", unidad)
    print("Total:", len(registros))

    for i, r in enumerate(registros[:10]):
        print(
            f"[{i}]",
            "servicio=", repr(r.get("servicio")),
            "| sentido=", repr(r.get("sentido")),
            "| linea=", repr(r.get("linea")),
            "| codigo_ts_a5=", repr(r.get("codigo_ts_a5"))
        )

    print("=== FIN DEBUG NODOS ===\n")

    for registro in registros:

        sentido_registro = str(
            registro.get("sentido") or ""
        ).strip()

        codigo_ts_registro = str(
            registro.get("codigo_ts_a5") or ""
        ).strip()

        if not codigo_ts_registro:
            continue

        if (
            sentido_registro == "1"
            and not codigo_ts_ida
        ):
            codigo_ts_ida = codigo_ts_registro

        elif (
            sentido_registro == "2"
            and not codigo_ts_regreso
        ):
            codigo_ts_regreso = codigo_ts_registro

    nodo_ida = resolver_nodo_catalogo(
        unidad,
        codigo_ts_ida,
        "1"
    )

    nodo_regreso = resolver_nodo_catalogo(
        unidad,
        codigo_ts_regreso,
        "2"
    )

    # Respaldo funcional:
    # si no existe nodo para ese codigo TS,
    # conserva el terminal anterior.
    ws["C3"] = nodo_ida or terminal
    ws["AG3"] = nodo_regreso or terminal

    print(
        "NODOS CABECERA:",
        f"Servicio={servicio}",
        f"IDA={codigo_ts_ida} -> {nodo_ida or terminal}",
        f"REG={codigo_ts_regreso} -> {nodo_regreso or terminal}",
    )


    if str(servicio).isdigit():
        servicio_excel = int(servicio)
    else:
        servicio_excel = servicio

    ws["C4"] = servicio_excel
    ws["AG4"] = servicio_excel


    ws["C6"] = tipo_dia
    ws["AG6"] = tipo_dia




    fila_ida = 13

    fila_reg = 13




    # ======================================================
    # RECORRER REGISTROS
    # ======================================================

    print(f"Total registros recibidos: {len(registros)}")
    
    for registro in registros:




        if servicio_puro(
            registro["servicio"]
        ) != servicio_puro(servicio):

            continue





        if formato_tipo_dia(
            registro["tipo_dia"]
        ) != formato_tipo_dia(tipo_dia):

            continue





        sentido = str(
            registro["sentido"]
        )




        # ==================================================
        # IDA
        # ==================================================

        if sentido == "1":



            ws.cell(
                fila_ida,
                2
            ).value = registro["tipo_bus"]

            celda = ws.cell(fila_ida, 3)

            if registro["tipo"] == "VPA":
                celda.value = "Vacío a Patio IDA"
                celda.fill = COLOR_VPA

            elif registro["tipo"] == "VEX":
                celda.value = "Vacío a Cabezal IDA"
                celda.fill = COLOR_VEX

            else:
                celda.value = servicio + " IDA"



            ws.cell(
                fila_ida,
                7
            ).value = registro["hora"]



            ws.cell(
                fila_ida,
                21
            ).value = registro["fin"]


            # ==============================================
            # ANEXO 5 - IDA
            # Solo expediciones comerciales desde Anexo 4
            # V = IP / IE
            # W = Punto de control teorico
            # ==============================================

            if (
                indice_anexo5 is not None
                and registro.get("tipo") == "EXP"
                and registro.get("codigo_ts_a5")
            ):

                medicion_ip = buscar_medicion_ip_ts(
                    indice_anexo5,
                    unidad,
                    registro.get("codigo_ts_a5"),
                    registro.get("tipo_dia"),
                    sentido,
                    registro.get("hora"),
                )

                if medicion_ip:

                    celda_indicador = ws.cell(
                        fila_ida,
                        22
                    )

                    celda_indicador.value = "IP"
                    celda_indicador.fill = COLOR_IP
                    celda_indicador.font = FUENTE_BLANCA

                    celda_control = ws.cell(
                        fila_ida,
                        23
                    )

                    celda_control.value = hora_excel(
                        medicion_ip.get(
                            "hora_teorica_texto"
                        )
                    )

                    celda_control.number_format = "hh:mm"

                else:

                    celda_indicador = ws.cell(
                        fila_ida,
                        22
                    )

                    celda_indicador.value = "IE"
                    celda_indicador.fill = COLOR_IE
                    celda_indicador.font = FUENTE_BLANCA

                    ws.cell(
                        fila_ida,
                        23
                    ).value = None


            fila_ida += 1






        # ==================================================
        # REGRESO
        # ==================================================

        elif sentido == "2":



            ws.cell(
                fila_reg,
                32
            ).value = registro["tipo_bus"]

            celda = ws.cell(fila_reg, 33)

            if registro["tipo"] == "VPA":

                celda.value = "Vacío a Patio REG"
                celda.fill = COLOR_VPA

            elif registro["tipo"] == "VEX":

                celda.value = "Vacío a Cabezal REG"
                celda.fill = COLOR_VEX

            else:

                celda.value = servicio + " REG"



            ws.cell(
                fila_reg,
                37
            ).value = registro["hora"]




            ws.cell(
                fila_reg,
                51
            ).value = registro["fin"]


            # ==============================================
            # ANEXO 5 - REGRESO
            # Solo expediciones comerciales desde Anexo 4
            # AZ = IP / IE
            # BA = Punto de control teorico
            # ==============================================

            if (
                indice_anexo5 is not None
                and registro.get("tipo") == "EXP"
                and registro.get("codigo_ts_a5")
            ):

                medicion_ip = buscar_medicion_ip_ts(
                    indice_anexo5,
                    unidad,
                    registro.get("codigo_ts_a5"),
                    registro.get("tipo_dia"),
                    sentido,
                    registro.get("hora"),
                )

                if medicion_ip:

                    celda_indicador = ws.cell(
                        fila_reg,
                        52
                    )

                    celda_indicador.value = "IP"
                    celda_indicador.fill = COLOR_IP
                    celda_indicador.font = FUENTE_BLANCA

                    celda_control = ws.cell(
                        fila_reg,
                        53
                    )

                    celda_control.value = hora_excel(
                        medicion_ip.get(
                            "hora_teorica_texto"
                        )
                    )

                    celda_control.number_format = "hh:mm"

                else:

                    celda_indicador = ws.cell(
                        fila_reg,
                        52
                    )

                    celda_indicador.value = "IE"
                    celda_indicador.fill = COLOR_IE
                    celda_indicador.font = FUENTE_BLANCA

                    ws.cell(
                        fila_reg,
                        53
                    ).value = None


            fila_reg += 1






    # ======================================================
    # VALIDAR DATOS
    # ======================================================

    if fila_ida == 13 and fila_reg == 13:


        wb.close()

        return None

    # ======================================================
    # LIMPIAR IDA
    # A:I
    # ======================================================

    limpiar_bloque(
        ws,
        fila_ida,
        1,
        9
    )

        # U:V
    limpiar_bloque(
        ws,
        fila_ida,
        21,
        23
    )

    # X:Y
    limpiar_bloque(
        ws,
        fila_ida,
        24,
        25
    )

    # Z:AB
    # Eliminar datos/formato sobrante despues
    # de las expediciones IDA generadas
    limpiar_bloque(
        ws,
        fila_ida,
        26,
        28
    )

    # ======================================================
    # LIMPIAR REG
    # ======================================================

    # AE:AM
    limpiar_bloque(
        ws,
        fila_reg,
        31,
        39
    )

    # AY:BA
    limpiar_bloque(
        ws,
        fila_reg,
        51,
        53
    )

    # BB:BF
    # Limpiar espacio sobrante despues
    # de la ultima expedicion REG.
    #
    # BB se conserva dentro de las expediciones
    # porque sera Punto Control real.
    limpiar_bloque(
        ws,
        fila_reg,
        54,
        58
    )


    # ======================================================
    # GUARDAR
    # ======================================================

    salida.parent.mkdir(

        parents=True,

        exist_ok=True

    )




    wb.save(
        salida
    )



    wb.close()

    print(
        f"{servicio} {tipo_dia}: "
        f"{round(time.time() - inicio_planilla, 2)} segundos"
    )


    return salida
# ==========================================================
# GENERADOR GENERAL
#
# Maneja:
# - Servicio individual
# - Servicio TODO
# - Tipo Día individual
# - Tipo Día TODO
#
# Retorna resumen
# ==========================================================

def generar_planillas(

    archivos_fus,

    plantilla,

    info_excel,

    salida,

    unidad,

    servicio,

    tipo_dia,

    usar_anexo4=False,

    incluir_vacios=True

):


    # ======================================================
    # LEER ORIGEN DE DATOS
    # ======================================================

    indice_anexo5 = None

    # ======================================================
    # LEER ORIGEN
    # ======================================================

    if usar_anexo4:

        registros = leer_anexo4(
            archivos_fus[0]
        )

    else:

        registros = leer_fus(
            archivos_fus
        )

    # ======================================================
    # CONFIGURACION DE SALIDAS
    # ======================================================

    cantidad_antes_filtro = len(registros)

    if not incluir_vacios:

        registros = [
            registro
            for registro in registros
            if str(
                registro.get("tipo", "")
            ).strip().upper() == "EXP"
        ]

    print("==============================")
    print("CONFIGURACION PLANILLAS")
    print("Incluir vacios:", incluir_vacios)
    print("Registros antes:", cantidad_antes_filtro)
    print("Registros despues:", len(registros))
    print("==============================")

    # ======================================================
    # CARGAR CATALOGO ANEXO 5 VIGENTE
    # PARA ANEXO 4 Y FUS
    # ======================================================

    ruta_a5 = ruta_catalogo_anexo5(
        unidad
    )

    if not ruta_a5.exists():

        raise FileNotFoundError(
            "No existe catalogo Anexo 5 vigente "
            f"para {unidad}: {ruta_a5}"
        )

    registros_a5 = leer_anexo5(
        ruta_a5
    )

    indice_anexo5 = crear_indice_anexo5_ts(
        registros_a5
    )

    print(
        "INDICE ANEXO 5 CARGADO:",
        len(indice_anexo5),
        "claves"
    )


    info = cargar_info()


    # ======================================================
    # DICCIONARIO DE TERMINALES
    # ======================================================

    terminales = {}
    codigos_ts = {}

    for dato in info:

        clave = (
            dato["unidad"],
            servicio_puro(dato["servicio"])
        )

        terminales[clave] = dato["terminal"]

        codigo_ts = dato.get("codigo_ts")

        if codigo_ts:

            codigos_ts[clave] = str(
                codigo_ts
            ).strip()

    print("TERMINALES CARGADOS:")
    print(terminales)

    # ======================================================
    # ASOCIAR CODIGO TS A REGISTROS
    # FUS: servicio cliente -> INFO -> codigo_ts
    # A4: conserva su codigo TS original
    # ======================================================

    unidad_info = unidad_empresa(
        unidad
    )

    sin_codigo_ts = set()

    for registro in registros:

        if usar_anexo4:

            registro["codigo_ts_a5"] = registro.get(
                "linea"
            )

        else:

            clave_info = (
                unidad_info,
                servicio_puro(
                    registro.get("servicio")
                )
            )

            codigo_ts = codigos_ts.get(
                clave_info
            )

            registro["codigo_ts_a5"] = codigo_ts

            if (
                registro.get("tipo") == "EXP"
                and not codigo_ts
            ):

                sin_codigo_ts.add(
                    servicio_puro(
                        registro.get("servicio")
                    )
                )

    if sin_codigo_ts:

        print(
            "ADVERTENCIA - SERVICIOS SIN CODIGO TS:",
            sorted(sin_codigo_ts)
        )

    # ======================================================
    # OBTENER SERVICIOS
    # ======================================================

    if servicio == "TODO":


        servicios = sorted(

            list(

                set(

                    servicio_puro(
                        x["servicio"]
                    )

                    for x in registros

                    if x["servicio"]

                )

            )

        )


    else:


        servicios = [

            servicio_puro(
                servicio
            )

        ]






    # ======================================================
    # OBTENER TIPOS DIA
    # ======================================================

    if tipo_dia == "TODO":


        tipos_dia = [

            "Laboral",

            "Sabado",

            "Domingo"

        ]


    else:


        tipos_dia = [

            formato_tipo_dia(
                tipo_dia
            )

        ]






    contador = {


        "total":0,

        "laboral":0,

        "sabado":0,

        "domingo":0

    }

    # ======================================================
    # INDICE DE REGISTROS
    # ======================================================

    indice_registros = {}

    for registro in registros:

        clave = (
            servicio_puro(registro["servicio"]),
            formato_tipo_dia(registro["tipo_dia"])
        )

        if clave not in indice_registros:
            indice_registros[clave] = []

        indice_registros[clave].append(registro)

    print(f"Indice creado: {len(indice_registros)} combinaciones")



    # ======================================================
    # GENERAR
    # ======================================================

    for dia in tipos_dia:

        for serv in servicios:

            # ----------------------------------------------
            # BUSCAR TERMINAL
            # ----------------------------------------------

            unidad_terminal = unidad_empresa(
                unidad
            )

            terminal = terminales.get(
                (
                    unidad_terminal,
                    servicio_puro(serv)
                ),
                ""
            )

            if not terminal:
                print("SIN TERMINAL:", serv)

            # ----------------------------------------------
            # CARPETA DESTINO
            # ----------------------------------------------

            carpeta = (
                salida
                / unidad
                / dia
            )

            archivo_salida = carpeta / (
                serv + " " + dia + ".xlsx"
            )

            print(
                f"Generando: Servicio={serv} | Tipo={dia}"
            )

            resultado = crear_planilla(

                indice_registros.get(
                    (
                        servicio_puro(serv),
                        formato_tipo_dia(dia)
                    ),
                    []
                ),

                plantilla,
                archivo_salida,
                unidad,
                serv,
                dia,
                terminal,
                indice_anexo5,
                usar_anexo4

            )

            if resultado:

                print("GENERADA:", resultado)

                contador["total"] += 1

                if dia == "Laboral":
                    contador["laboral"] += 1

                elif dia == "Sabado":
                    contador["sabado"] += 1

                elif dia == "Domingo":
                    contador["domingo"] += 1

    return contador