from pathlib import Path

import pandas as pd

from .config import INFO_EXCEL

from .info import obtener_unidad_por_servicio


# ==========================================================
# NORMALIZAR
# ==========================================================

def normaliza(valor):

    if pd.isna(valor):
        return ""

    return str(valor).strip()


# ==========================================================
# DICCIONARIO
# CODIGO TS -> SERVICIO CLIENTE
# ==========================================================

def cargar_servicios():

    df = pd.read_excel(INFO_EXCEL)

    df.columns = df.columns.str.strip()

    for col in [
        "servicio",
        "CODIGO TS SERVICIO"
    ]:

        df[col] = (
            df[col]
            .fillna("")
            .astype(str)
            .str.strip()
        )

    dic = {}

    for _, fila in df.iterrows():

        codigo = normaliza(
            fila["CODIGO TS SERVICIO"]
        )

        servicio = normaliza(
            fila["servicio"]
        )

        if codigo:
            dic[codigo] = servicio

    return dic


# ==========================================================
# FORMATO TIPO DIA
# ==========================================================

def formato_tipo_dia(valor):

    texto = normaliza(valor).upper()

    if texto == "LABORAL":
        return "Laboral"

    if texto in ("SABADO", "SÁBADO"):
        return "Sabado"

    if texto == "DOMINGO":
        return "Domingo"

    return texto.title()


# ==========================================================
# LEER ANEXO 4
# ==========================================================

def leer_anexo4(ruta_excel):

    ruta = Path(ruta_excel)

    df = pd.read_excel(
        ruta,
        sheet_name="Tabla Horaria",
        header=6
    )

    if "CODIGO TS SERVICIO" not in df.columns:

        return (
            False,
            "El archivo no contiene la columna CODIGO TS SERVICIO."
        )

    columnas_obligatorias = [

        "UNIDAD DE SERVICIO",
        "BUS_LOGICO",
        "CODIGO TS SERVICIO",
        "SENTIDO",
        "TIPO_DIA",
        "TIPO_EVENTO",
        "HORA_INICIO",
        "HORA_FIN",
        "PUNTO_INICIO",
        "PUNTO_FIN",
        "DISTANCIA (KM)",
        "TIPO_BUS",
    ]

    if not all(c in df.columns for c in columnas_obligatorias):

        raise ValueError(
            "El archivo seleccionado no corresponde a un Anexo 4 válido."
        )

    # ======================================================
    # SOLO EXPEDICIONES (C01)
    # ======================================================

    df = df[
        df["TIPO_EVENTO"] == "C01"
    ].copy()

    # ======================================================
    # ELIMINAR FS
    # ======================================================

    df = df[
        df["SENTIDO"] != "FS"
    ].copy()

    # ======================================================
    # DICCIONARIO TS -> SERVICIO
    # ======================================================

    dic_servicios = cargar_servicios()

    registros = []

    for _, fila in df.iterrows():

        codigo_ts = normaliza(
            fila["CODIGO TS SERVICIO"]
        )

        servicio = dic_servicios.get(
            codigo_ts,
            codigo_ts
        )

        sentido = normaliza(
            fila["SENTIDO"]
        ).upper()

        if sentido == "IDA":
            sentido = "1"

        elif sentido == "RET":
            sentido = "2"

        else:
            continue

        tipo_bus = normaliza(
            fila["TIPO_BUS"]
        )

        if tipo_bus:
            tipo_bus = tipo_bus[0]

        registro = {

            "tipo": "EXP",

            "evento": "EXP",

            "hora":
                fila["HORA_INICIO"],

            "fin":
                fila["HORA_FIN"],

            "tipo_bus":
                tipo_bus,

            "servicio":
                servicio,

            "linea":
                codigo_ts,

            "tipo_dia":
                formato_tipo_dia(
                    fila["TIPO_DIA"]
                ),

            "sentido":
                sentido,

            # Información adicional
            "bus":
                fila["BUS_LOGICO"],

            "desde":
                fila["PUNTO_INICIO"],

            "hasta":
                fila["PUNTO_FIN"],

            "km":
                fila["DISTANCIA (KM)"],

            "fila":
                fila.tolist()

        }

        registros.append(
            registro
        )

    registros.sort(
        key=lambda x: (
            x["servicio"] or "",
            x["sentido"] or "",
            str(x["hora"])
        )
    )

    print("==============================")
    print("TOTAL EXP:", len(registros))
    print("==============================")

    return registros
# ==========================================================
# VALIDAR UNIDAD ANEXO 4
# ==========================================================
def validar_unidad_anexo4(
    ruta_excel,
    unidad_seleccionada
):
    df = pd.read_excel(
    ruta_excel,
    sheet_name="Tabla Horaria",
        header=6
    )

    if "UNIDAD DE SERVICIO" not in df.columns:

        return (
            False,
            "El archivo no contiene la columna UNIDAD DE SERVICIO."
        )

    servicios = (
        df["CODIGO TS SERVICIO"]
        .dropna()
        .astype(str)
        .str.strip()
        .unique()
    )

    print("SERVICIOS ENCONTRADOS EN ANEXO 4:")
    print(servicios)

    unidades_detectadas = set()

    for servicio in servicios:

        unidad = obtener_unidad_por_servicio(servicio)

        if unidad:

            unidades_detectadas.add(unidad)

    

    if not unidades_detectadas:

        return (
            False,
            "No fue posible determinar la unidad del archivo."
        )

    if unidad_seleccionada not in unidades_detectadas:

        return (
            False,
            f"El Anexo 4 corresponde a {', '.join(unidades_detectadas)} y no a {unidad_seleccionada}."
        )

    return True, None