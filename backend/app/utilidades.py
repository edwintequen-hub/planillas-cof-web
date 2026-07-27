# =====================================================
# NORMALIZAR TEXTOS
# =====================================================

def normaliza(valor):

    """
    Equivalente a Normaliza() de VBA.

    Convierte:
    - None -> ""
    - elimina espacios
    - pasa a mayúsculas
    """

    if valor is None:
        return ""


    return str(valor).strip().upper()





# =====================================================
# SERVICIO PURO
# =====================================================

def servicio_puro(valor):

    """
    Obtiene el código de servicio.

    Ejemplos:

    120-IDA  -> 120

    B18-REG -> B18

    """

    texto = normaliza(valor)


    if "-" in texto:

        return texto.split("-")[0]


    return texto





# =====================================================
# CONVERTIR HORA
# =====================================================

def hora_segura(valor):

    """
    Conversión segura de horas.

    Soporta:

    08:30
    24:00
    25:30

    Retorna formato Excel.
    """

    if valor is None:

        return None



    texto = str(valor).strip()


    if texto == "":

        return None



    texto = texto.replace(
        ".",
        ":"
    )


    partes = texto.split(":")



    try:


        if len(partes) >= 2:


            horas = int(partes[0])

            minutos = int(partes[1])


            return (
                horas / 24
                +
                minutos / 1440
            )


    except Exception:

        return None



    return None