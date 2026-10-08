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

    from openpyxl import load_workbook

    ruta = Path(ruta_excel)

    if not ruta.exists():
        raise FileNotFoundError(
            f"No existe Anexo 4: {ruta}"
        )

    print("Abriendo Anexo 4...")

    wb = load_workbook(
        ruta,
        read_only=True,
        data_only=True,
    )

    if "Tabla Horaria" not in wb.sheetnames:
        wb.close()

        raise ValueError(
            "El archivo no contiene la hoja Tabla Horaria."
        )

    ws = wb["Tabla Horaria"]

    print("Hoja:", ws.title)
    print("Leyendo encabezados...")

    fila_encabezado = 7

    encabezados = {}

    for celda in next(
        ws.iter_rows(
            min_row=fila_encabezado,
            max_row=fila_encabezado,
            values_only=False,
        )
    ):
        if celda.value is None:
            continue

        nombre = str(
            celda.value
        ).strip()

        encabezados[nombre] = (
            celda.column - 1
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

    faltantes = [
        c
        for c in columnas_obligatorias
        if c not in encabezados
    ]

    if faltantes:
        wb.close()

        raise ValueError(
            "Faltan columnas obligatorias Anexo 4: "
            + ", ".join(faltantes)
        )

    dic_servicios = cargar_servicios()

    registros = []

    total_filas = 0
    total_c01 = 0
    total_fs = 0

    print("Procesando registros...")

    for numero_fila, valores in enumerate(
        ws.iter_rows(
            min_row=fila_encabezado + 1,
            values_only=True,
        ),
        start=fila_encabezado + 1,
    ):

        total_filas += 1

        if total_filas % 10000 == 0:
            print(
                f"Filas procesadas: {total_filas:,}"
            )

        def valor(nombre):
            pos = encabezados[nombre]

            if pos >= len(valores):
                return None

            return valores[pos]

        tipo_evento = normaliza(
            valor("TIPO_EVENTO")
        ).upper()

        # SOLO EXPEDICIONES COMERCIALES C01
        if tipo_evento != "C01":
            continue

        total_c01 += 1

        sentido = normaliza(
            valor("SENTIDO")
        ).upper()

        # EXCLUIR FS
        if sentido == "FS":
            total_fs += 1
            continue

        if sentido == "IDA":
            sentido_final = "1"

        elif sentido == "RET":
            sentido_final = "2"

        else:
            continue

        codigo_ts = normaliza(
            valor("CODIGO TS SERVICIO")
        )

        servicio = dic_servicios.get(
            codigo_ts,
            codigo_ts,
        )

        tipo_bus = normaliza(
            valor("TIPO_BUS")
        )

        if tipo_bus:
            tipo_bus = tipo_bus[0]

        registro = {

            "tipo":
                "EXP",

            "evento":
                "EXP",

            "hora":
                valor("HORA_INICIO"),

            "fin":
                valor("HORA_FIN"),

            "tipo_bus":
                tipo_bus,

            "servicio":
                servicio,

            "linea":
                codigo_ts,

            "tipo_dia":
                formato_tipo_dia(
                    valor("TIPO_DIA")
                ),

            "sentido":
                sentido_final,

            "bus":
                valor("BUS_LOGICO"),

            "desde":
                valor("PUNTO_INICIO"),

            "hasta":
                valor("PUNTO_FIN"),

            "km":
                valor("DISTANCIA (KM)"),

            "fila":
                numero_fila,
        }

        registros.append(
            registro
        )

    wb.close()

    registros.sort(
        key=lambda x: (
            x["servicio"] or "",
            x["sentido"] or "",
            str(x["hora"]),
        )
    )

    print()
    print("=" * 60)
    print("ANEXO 4 PROCESADO")
    print("=" * 60)
    print("FILAS REVISADAS:", total_filas)
    print("C01 ENCONTRADAS:", total_c01)
    print("FS EXCLUIDAS:", total_fs)
    print("TOTAL EXP:", len(registros))
    print("=" * 60)

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