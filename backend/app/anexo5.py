from pathlib import Path
from datetime import datetime, time, timedelta
import unicodedata

from openpyxl import load_workbook


# ==========================================================
# NORMALIZACION
# ==========================================================

def normalizar_texto(valor):
    if valor is None:
        return ""

    texto = str(valor).strip().upper()

    texto = unicodedata.normalize("NFD", texto)
    texto = "".join(
        c for c in texto
        if unicodedata.category(c) != "Mn"
    )

    return " ".join(texto.split())


def normalizar_servicio(valor):
    texto = normalizar_texto(valor)

    if "-" in texto:
        texto = texto.split("-")[0].strip()

    return texto


def normalizar_tipo_dia(valor):
    texto = normalizar_texto(valor)

    if texto == "LABORAL":
        return "Laboral"

    if texto == "SABADO":
        return "Sabado"

    if texto == "DOMINGO":
        return "Domingo"

    return str(valor).strip() if valor is not None else ""


def normalizar_sentido(valor):
    texto = normalizar_texto(valor)

    if texto in ("IDA", "I", "1"):
        return "1"

    if texto in (
        "REG",
        "RET",
        "REGRESO",
        "RETORNO",
        "R",
        "2",
    ):
        return "2"

    return texto


# ==========================================================
# MOVIMIENTOS COMERCIALES
# ==========================================================

def es_movimiento_comercial(*valores):

    texto = " ".join(
        normalizar_texto(v)
        for v in valores
        if v is not None
    )

    exclusiones = (
        "VACIO A PATIO",
        "VACIO PATIO",
        "VACIO A CABEZAL",
        "VACIO CABEZAL",
    )

    return not any(
        exclusion in texto
        for exclusion in exclusiones
    )


# ==========================================================
# HORA
# ==========================================================

def normalizar_hora(valor):

    if valor is None:
        return None

    if isinstance(valor, datetime):
        return valor.time().replace(microsecond=0)

    if isinstance(valor, time):
        return valor.replace(microsecond=0)

    if isinstance(valor, timedelta):
        segundos = int(valor.total_seconds()) % 86400
        horas = segundos // 3600
        minutos = (segundos % 3600) // 60
        segundos = segundos % 60

        return time(
            horas,
            minutos,
            segundos,
        )

    if isinstance(valor, (int, float)):
        segundos = int(
            round(float(valor) * 24 * 60 * 60)
        ) % 86400

        horas = segundos // 3600
        minutos = (segundos % 3600) // 60
        segundos = segundos % 60

        return time(
            horas,
            minutos,
            segundos,
        )

    texto = str(valor).strip()

    formatos = (
        "%H:%M:%S",
        "%H:%M",
    )

    for formato in formatos:
        try:
            return datetime.strptime(
                texto,
                formato
            ).time()
        except ValueError:
            pass

    return None


def hora_texto(valor):

    hora = normalizar_hora(valor)

    if hora is None:
        return ""

    return hora.strftime("%H:%M:%S")


def ponderador_es_uno(valor):

    if valor is None:
        return False

    try:
        return abs(
            float(valor) - 1.0
        ) < 0.000001
    except Exception:
        pass

    texto = normalizar_texto(valor)

    return texto in (
        "1",
        "1.0",
        "1,0",
    )


# ==========================================================
# BUSQUEDA DE COLUMNAS
# ==========================================================


def ponderador_es_uno_o_dos(valor):

    if valor is None:
        return False

    try:

        numero = float(valor)

        return (
            abs(numero - 1.0) < 0.000001
            or
            abs(numero - 2.0) < 0.000001
        )

    except Exception:
        pass

    texto = normalizar_texto(valor)

    return texto in (
        "1",
        "1.0",
        "2",
        "2.0",
    )



def buscar_encabezado(
    ws,
    aliases,
    filas_maximas=15,
):

    aliases_normalizados = {
        normalizar_texto(x)
        for x in aliases
    }

    for fila in range(
        1,
        min(ws.max_row, filas_maximas) + 1
    ):

        for columna in range(
            1,
            ws.max_column + 1
        ):

            valor = normalizar_texto(
                ws.cell(
                    fila,
                    columna
                ).value
            )

            if valor in aliases_normalizados:
                return fila, columna

    return None, None


def detectar_estructura(ws):

    import unicodedata

    def normalizar_encabezado(valor):

        if valor is None:
            return ""

        texto = str(valor).strip().upper()

        texto = "".join(
            c
            for c in unicodedata.normalize(
                "NFD",
                texto,
            )
            if unicodedata.category(c) != "Mn"
        )

        texto = " ".join(
            texto.split()
        )

        return texto

    campos = {

        "unidad": (
            "UNIDAD DE SERVICIO",
        ),

        "codigo_ts": (
            "CODIGO TS SERVICIO",
        ),

        "servicio": (
            "SERVICIO",
            "SERVICIO USUARIO",
            "SERVICIO CLIENTE",
            "CODIGO USUARIO SERVICIO",
        ),

        "sentido": (
            "SENTIDO",
        ),

        "tipo_dia": (
            "TIPO DIA",
            "TIPO_DIA",
            "TIPO DE DIA",
        ),

        "expedicion": (
            "EXPEDICION",
        ),

        "codigo_paradero": (
            "CODIGO PARADERO USUARIO",
            "CODIGO PARADERO",
        ),

        "nombre_paradero": (
            "NOMBRE PARADERO",
        ),

        "correlativo": (
            "CORRELATIVO PARADERO",
            "CORRELATIVO",
        ),

        "hora_pasada": (
            "HORARIO DE PASADA PARADERO",
            "HORARIO PASADA PARADERO",
        ),

        "hora_salida": (
            "HORARIO DE SALIDA DESDE CABEZAL",
            "HORARIO SALIDA DESDE CABEZAL",
        ),

        "ponderador_salida": (
            "PONDERADOR DE SALIDA",
            "PONDERADOR SALIDA",
        ),

        "tipo_medicion": (
            "TIPO DE MEDICION",
            "TIPO MEDICION",
        ),

        "medicion_ip": (
            "MEDICION IP",
        ),
    }

    aliases = {
        campo: {
            normalizar_encabezado(x)
            for x in opciones
        }
        for campo, opciones in campos.items()
    }

    # ======================================================
    # BUSCAR UNA FILA QUE CONTENGA LA ESTRUCTURA COMPLETA
    # DE LA TABLA, NO ENCABEZADOS GENERALES DEL DOCUMENTO.
    # ======================================================

    campos_clave = {
        "codigo_ts",
        "servicio",
        "sentido",
        "tipo_dia",
        "expedicion",
        "hora_pasada",
        "hora_salida",
        "ponderador_salida",
        "tipo_medicion",
        "medicion_ip",
    }

    mejor_fila = None
    mejor_columnas = None
    mejor_puntaje = -1

    limite_filas = min(
        ws.max_row or 30,
        30,
    )

    for numero_fila, valores in enumerate(
        ws.iter_rows(
            min_row=1,
            max_row=limite_filas,
            values_only=True,
        ),
        start=1,
    ):

        encontrados = {}

        for numero_columna, valor in enumerate(
            valores,
            start=1,
        ):

            encabezado = normalizar_encabezado(
                valor
            )

            if not encabezado:
                continue

            for campo, opciones in aliases.items():

                if (
                    campo not in encontrados
                    and encabezado in opciones
                ):
                    encontrados[campo] = numero_columna

        puntaje = len(
            campos_clave.intersection(
                encontrados.keys()
            )
        )

        if puntaje > mejor_puntaje:
            mejor_puntaje = puntaje
            mejor_fila = numero_fila
            mejor_columnas = encontrados

    if mejor_fila is None:
        raise ValueError(
            "No fue posible detectar la fila "
            "de encabezados del Anexo 5."
        )

    faltantes = [
        campo
        for campo in campos.keys()
        if campo not in mejor_columnas
    ]

    if faltantes:
        raise ValueError(
            "Faltan columnas obligatorias "
            "en Anexo 5: "
            + ", ".join(faltantes)
        )

    return (
        mejor_fila,
        mejor_columnas,
    )


def leer_anexo5(ruta_excel):

    ruta = Path(ruta_excel)

    if not ruta.exists():
        raise FileNotFoundError(
            f"No existe Anexo 5: {ruta}"
        )

    print("Abriendo Anexo 5...")

    wb = load_workbook(
        ruta,
        data_only=True,
        read_only=True,
    )

    nombre_preferido = "Horarios de pasada_completo"

    if nombre_preferido in wb.sheetnames:
        ws = wb[nombre_preferido]
    else:
        ws = wb[wb.sheetnames[0]]

    print("Hoja:", ws.title)
    print("Buscando encabezados...")

    fila_encabezado, col = detectar_estructura(ws)

    print("Fila encabezado:", fila_encabezado)
    print("Procesando registros...")

    candidatos = []

    # ------------------------------------------------------
    # Convertimos numero Excel de columna a indice 0
    # para trabajar directamente con iter_rows(values_only=True)
    # ------------------------------------------------------

    idx = {
        nombre: numero - 1
        for nombre, numero in col.items()
    }

    total_leidas = 0

    for numero_fila, valores in enumerate(
        ws.iter_rows(
            min_row=fila_encabezado + 1,
            values_only=True,
        ),
        start=fila_encabezado + 1,
    ):

        total_leidas += 1

        # Mostrar avance cada 10.000 filas
        if total_leidas % 10000 == 0:
            print(
                f"Filas procesadas: {total_leidas:,}"
            )

        def valor(campo):
            posicion = idx[campo]

            if posicion >= len(valores):
                return None

            return valores[posicion]

        # ==================================================
        # CRITERIO 1:
        # TIPO DE MEDICION = VIGENTE
        # ==================================================

        tipo_medicion = normalizar_texto(
            valor("tipo_medicion")
        )

        if tipo_medicion != "VIGENTE":
            continue

        # ==================================================
        # CRITERIO 2:
        # MEDICION IP = SI
        # ==================================================

        medicion_ip = normalizar_texto(
            valor("medicion_ip")
        )

        if medicion_ip not in (
            "SI",
            "S",
            "YES",
            "1",
        ):
            continue

        # ==================================================
        # CRITERIO 3:
        # PONDERADOR DE SALIDA = 1 O 2
        # ==================================================

        ponderador_salida = valor(
            "ponderador_salida"
        )

        if not ponderador_es_uno_o_dos(
            ponderador_salida
        ):
            continue

        unidad_valor = valor(
            "unidad"
        )

        if unidad_valor is None:
            unidad = ""
        else:
            unidad = str(
                unidad_valor
            ).strip()

        codigo_ts_valor = valor(
            "codigo_ts"
        )

        if codigo_ts_valor is None:
            codigo_ts = ""
        else:
            codigo_ts = str(
                codigo_ts_valor
            ).strip()

            # Excel puede entregar 801.0
            if codigo_ts.endswith(".0"):
                codigo_ts = codigo_ts[:-2]

        servicio = normalizar_servicio(
            valor("servicio")
        )

        sentido = normalizar_sentido(
            valor("sentido")
        )

        tipo_dia = normalizar_tipo_dia(
            valor("tipo_dia")
        )

        expedicion = valor(
            "expedicion"
        )

        codigo_paradero = valor(
            "codigo_paradero"
        )

        nombre_paradero = valor(
            "nombre_paradero"
        )

        correlativo = valor(
            "correlativo"
        )

        hora_pasada = normalizar_hora(
            valor("hora_pasada")
        )

        hora_salida = normalizar_hora(
            valor("hora_salida")
        )

        if not servicio:
            continue

        if sentido not in ("1", "2"):
            continue

        if hora_salida is None:
            continue

        if hora_pasada is None:
            continue

        # ==================================================
        # SOLO SALIDAS COMERCIALES
        # ==================================================

        if not es_movimiento_comercial(
            servicio,
            nombre_paradero,
        ):
            continue

        candidatos.append({

            "unidad":
                unidad,

            "codigo_ts":
                codigo_ts,

            "servicio":
                servicio,

            "sentido":
                sentido,

            "tipo_dia":
                tipo_dia,

            "expedicion":
                expedicion,

            "codigo_paradero":
                codigo_paradero,

            "nombre_paradero":
                nombre_paradero,

            "correlativo":
                correlativo,

            "hora_salida":
                hora_salida,

            "hora_salida_texto":
                hora_texto(hora_salida),

            "hora_teorica":
                hora_pasada,

            "hora_teorica_texto":
                hora_texto(hora_pasada),

            "tipo_medicion":
                tipo_medicion,

            "medicion_ip":
                medicion_ip,

            "ponderador_salida":
                ponderador_salida,

            "fila_anexo5":
                numero_fila,
        })

    wb.close()

    # ======================================================
    # Con PONDERADOR DE SALIDA = 1 ya estamos seleccionando
    # el punto IP definido para la salida.
    #
    # Si existiera accidentalmente mas de uno para la misma
    # expedicion, conservamos uno solo.
    # ======================================================

    # ======================================================
    # REGLA:
    # Para una misma salida pueden existir varias expediciones.
    # Se conserva SIEMPRE la primera expedicion.
    #
    # NO se modifica:
    # - Tipo de medicion
    # - Medicion IP
    # - Ponderador
    # - Servicio / TS
    # - Tipo de dia
    # - Sentido
    # - Hora de salida desde cabezal
    # - Hora teorica de pasada
    # ======================================================

    # ======================================================
    # REGLA OFICIAL DE SELECCION ANEXO 5
    #
    # Los candidatos ya cumplen:
    # - Tipo de Medicion = VIGENTE
    # - Medicion IP = SI
    # - Ponderador de Salida = 1 o 2
    #
    # Para CADA EXPEDICION se conserva la PRIMERA FILA
    # valida que aparece en el Anexo 5.
    #
    # Esa misma fila aporta:
    # - HORARIO DE SALIDA DESDE CABEZAL
    # - HORARIO DE PASADA PARADERO
    #
    # La comparacion posterior de la hora NO se modifica.
    # ======================================================

    primeros = {}

    for registro in candidatos:

        clave = (
            registro["servicio"],
            registro.get("codigo_ts"),
            normalizar_tipo_dia(
                registro["tipo_dia"]
            ),
            normalizar_sentido(
                registro["sentido"]
            ),
            registro["expedicion"],
        )

        actual = primeros.get(clave)

        # --------------------------------------------------
        # La primera fila valida encontrada de la expedicion
        # queda seleccionada.
        # --------------------------------------------------

        if actual is None:

            primeros[clave] = registro
            continue

        # --------------------------------------------------
        # Seguridad:
        # si por orden del archivo apareciera un correlativo
        # menor posteriormente, conservar el menor.
        # --------------------------------------------------

        try:
            correlativo_nuevo = int(
                float(
                    registro["correlativo"]
                )
            )
        except Exception:
            correlativo_nuevo = 999999

        try:
            correlativo_actual = int(
                float(
                    actual["correlativo"]
                )
            )
        except Exception:
            correlativo_actual = 999999

        if correlativo_nuevo < correlativo_actual:

            primeros[clave] = registro


    registros = list(
        primeros.values()
    )

    registros.sort(
        key=lambda x: (
            x["servicio"],
            x["tipo_dia"],
            x["sentido"],
            x["hora_salida_texto"],
        )
    )

    print()
    print("=" * 60)
    print("ANEXO 5 PROCESADO")
    print("=" * 60)
    print("ARCHIVO:", ruta.name)
    print("HOJA:", ws.title)
    print("FILAS REVISADAS:", total_leidas)
    print(
        "FILAS QUE CUMPLEN "
        "VIGENTE + IP SI + POND 1:",
        len(candidatos)
    )
    print(
        "REGISTROS IP FINALES:",
        len(registros)
    )
    print("=" * 60)

    return registros


# ==========================================================
# CREAR INDICE DE BUSQUEDA
# ==========================================================

def crear_indice_anexo5(registros):

    indice = {}

    for registro in registros:

        clave = (
            normalizar_servicio(
                registro["servicio"]
            ),
            normalizar_sentido(
                registro["sentido"]
            ),
            normalizar_tipo_dia(
                registro["tipo_dia"]
            ),
            registro["hora_salida_texto"],
        )

        indice[clave] = registro

    return indice


# ==========================================================
# BUSCAR UNA SALIDA
# ==========================================================

def buscar_medicion_ip(
    indice,
    servicio,
    sentido,
    tipo_dia,
    hora_salida,
):

    clave = (
        normalizar_servicio(servicio),
        normalizar_sentido(sentido),
        normalizar_tipo_dia(tipo_dia),
        hora_texto(hora_salida),
    )

    return indice.get(clave)


# ==========================================================
# INDICE TECNICO ANEXO 5
# UNIDAD + CODIGO TS + TIPO DIA + SENTIDO + HORA SALIDA
# ==========================================================

def normalizar_unidad_anexo5(valor):

    texto = str(
        valor or ""
    ).strip().upper()

    texto = (
        texto
        .replace(" ", "")
    )

    equivalencias = {
        "U8": "8",
        "8": "8",
        "ALFA": "8",
        "ALFAU8": "8",

        "U9": "9",
        "9": "9",
        "OMEGA": "9",
        "OMEGAU9": "9",
    }

    if texto.endswith(".0"):
        texto = texto[:-2]

    return equivalencias.get(
        texto,
        texto
    )



def normalizar_codigo_ts(valor):

    texto = str(
        valor or ""
    ).strip().upper()

    if texto.startswith("T"):
        texto = texto[1:]

    if texto.endswith(".0"):
        texto = texto[:-2]

    return texto


def crear_indice_anexo5_ts(registros):

    indice = {}

    for registro in registros:

        clave = (
            normalizar_unidad_anexo5(
                registro.get("unidad")
            ),
            normalizar_codigo_ts(
                registro.get("codigo_ts")
            ),
            normalizar_tipo_dia(
                registro.get("tipo_dia")
            ),
            normalizar_sentido(
                registro.get("sentido")
            ),
            hora_texto(
                registro.get("hora_salida")
            ),
        )

        indice[clave] = registro

    return indice


def buscar_medicion_ip_ts(
    indice,
    unidad,
    codigo_ts,
    tipo_dia,
    sentido,
    hora_salida,
):

    clave = (
        normalizar_unidad_anexo5(
            unidad
        ),
        normalizar_codigo_ts(
            codigo_ts
        ),
        normalizar_tipo_dia(
            tipo_dia
        ),
        normalizar_sentido(
            sentido
        ),
        hora_texto(
            hora_salida
        ),
    )

    return indice.get(
        clave
    )

