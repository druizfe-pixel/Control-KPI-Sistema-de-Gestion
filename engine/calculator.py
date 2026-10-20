"""
TBG Calidad - Motor de Cálculo
Paso 10

Responsabilidad:
    Calcular el cumplimiento y resultado ponderado
    de cada indicador.

Reglas:

    P = Mayor es mejor
        Cumplimiento = Real / Meta

    N = Menor es mejor
        Cumplimiento = Meta / Real

    El cumplimiento tiene un tope de 100%.

    Resultado ponderado =
        Cumplimiento × Peso
"""
from datetime import datetime

def calcular_indicador(indicador):
    """
    Calcula un indicador previamente validado.

    Parámetro
    ---------
    indicador : dict

    Retorna
    -------
    dict
        Resultado del cálculo.
    """

    meta = float(indicador["meta"])
    real = float(indicador["real"])
    tipo = str(indicador["tipo"]).strip().upper()
    peso = float(indicador["peso"])

    # ---------------------------------------------------------
    # 1. CALCULAR CUMPLIMIENTO BRUTO
    # ---------------------------------------------------------


    if tipo == "P":

        cumplimiento_bruto = real / meta

    elif tipo == "N":

        cumplimiento_bruto = meta / real

    elif tipo == "N0":

        if real == 0:
            # Cero real representa el mejor resultado posible.
            # Se evita la división por cero.
            cumplimiento_bruto = 1.0
        else:
            cumplimiento_bruto = meta / real

    else:

        raise ValueError(
            f"Tipo de indicador inválido: {tipo}"
        )
    # if tipo == "P":

    #     cumplimiento_bruto = real / meta

    # elif tipo == "N":

    #     cumplimiento_bruto = meta / real

    # else:

    #     raise ValueError(
    #         f"Tipo de indicador inválido: {tipo}"
    #     )

    # ---------------------------------------------------------
    # 2. APLICAR TOPE
    # ---------------------------------------------------------

    cumplimiento = min(
        cumplimiento_bruto,
        1.0
    )

    # ---------------------------------------------------------
    # 3. CALCULAR RESULTADO PONDERADO
    # ---------------------------------------------------------

    resultado_ponderado = (
        cumplimiento * peso
    )

    # ---------------------------------------------------------
    # 4. CONSTRUIR RESULTADO
    # ---------------------------------------------------------

    return {

        "codigo": indicador["codigo"],

        "indicador": indicador["indicador"],

        "meta": meta,

        "real": real,

        "tipo": tipo,

        "peso": peso,

        "cumplimiento_bruto": cumplimiento_bruto,

        "cumplimiento": cumplimiento,

        "resultado_ponderado": resultado_ponderado
    }

def preparar_calculo_periodo(indicadores):
    """
    Preparar los indicadores según su estado operativo
    antes de ejecutar el cálculo.

    Estados de indicador:

        OK
            El indicador puede calcularse.

        PENDIENTE
            El indicador no tiene datos todavía.
            No es un error y no se calcula.

        DATO_FALTANTE
            Existe información incompleta.
            No se calcula y queda pendiente.

        REVISION REAL
            Existe un valor que requiere revisión (por ejemplo,
            Real = 0 en un indicador tipo N).
            No se calcula y no bloquea los demás indicadores.

        ERROR
            Existe un error estructural/de validación.
            Bloquea el período.

    Estados del período:

        COMPLETO
            Todos los indicadores están en OK.

        PARCIAL
            Existen indicadores OK y PENDIENTE,
            sin errores.

        SIN_DATOS
            Todos los indicadores están PENDIENTE.

        BLOQUEADO
            Existe al menos un error.

    Retorna
    -------
    dict
        {
            "calculables": [...],
            "pendientes": [...],
            "errores": [...],
            "estado_periodo": "...",
            "puede_calcular": True/False
        }

    Importante
    ----------
    Esta función NO realiza cálculos matemáticos.
    Solamente determina qué indicadores pueden pasar
    al motor de cálculo.
    """

    calculables = []
    pendientes = []
    revisiones = []
    errores = []

    for indicador in indicadores:

        estado = str(
            indicador.get("estado", "")
        ).strip().upper()

        # -----------------------------------------------------
        # INDICADOR OK
        # -----------------------------------------------------

        if estado == "OK":

            calculables.append(indicador)

        # -----------------------------------------------------
        # INDICADOR PENDIENTE
        # -----------------------------------------------------

        elif estado == "PENDIENTE":

            pendientes.append(indicador)

        # -----------------------------------------------------
        # DATO FALTANTE
        # -----------------------------------------------------

        elif estado == "DATO_FALTANTE":

            pendientes.append(indicador)

        # -----------------------------------------------------
        # REVISIÓN REAL
        # -----------------------------------------------------

        elif estado.startswith("REVISION REAL"):

            revisiones.append(indicador)

        # -----------------------------------------------------
        # ERROR ESTRUCTURAL / VALIDACIÓN BLOQUEANTE
        # -----------------------------------------------------

        elif estado == "ERROR":

            errores.append(indicador)

        # -----------------------------------------------------
        # ESTADO NO RECONOCIDO
        # -----------------------------------------------------

        else:

            errores.append({
                **indicador,
                "estado": "ERROR",
                "error": "ESTADO_INVALIDO",
            })

    # ---------------------------------------------------------
    # DETERMINAR ESTADO DEL PERÍODO
    # ---------------------------------------------------------

    if errores:

        estado_periodo = "BLOQUEADO"

    elif not calculables and (pendientes or revisiones):

        estado_periodo = "SIN_DATOS"

    elif pendientes or revisiones:

        estado_periodo = "PARCIAL"

    else:

        estado_periodo = "COMPLETO"

    # ---------------------------------------------------------
    # DETERMINAR SI EL MOTOR PUEDE CALCULAR
    # ---------------------------------------------------------

    puede_calcular = estado_periodo in (
        "COMPLETO",
        "PARCIAL",
    )

    return {
        "calculables": calculables,
        "pendientes": pendientes,
        "revisiones": revisiones,
        "errores": errores,
        "estado_periodo": estado_periodo,
        "puede_calcular": puede_calcular,
    }




def calcular_indicadores(indicadores, callback=None):
    """
    Calcula una lista de indicadores.

    Parámetro
    ---------
    indicadores : list[dict]

    Retorna
    -------
    list[dict]
        Resultados calculados.
    """

    resultados = []

    # TBG_PROGRESS_CALLBACK
    total_indicadores = len(indicadores)

    for posicion, indicador in enumerate(indicadores, start=1):

        if callback is not None:
            try:
                callback(posicion, total_indicadores)
            except Exception:
                pass

        resultado = calcular_indicador(
            indicador
        )

        resultados.append(resultado)

    return resultados


def total_resultado_ponderado(resultados):
    """
    Calcula el resultado ponderado total.

    Parámetro
    ---------
    resultados : list[dict]

    Retorna
    -------
    float
        Resultado total.
    """

    return sum(
        resultado["resultado_ponderado"]
        for resultado in resultados
    )

def calcular_acumulado(cumplimientos):
    """
    Calcular el acumulado de un indicador.

    El acumulado corresponde al promedio simple de los
    cumplimientos mensuales válidos.

    Los meses sin datos válidos no participan en el cálculo.

    Parámetros:
        cumplimientos: lista de valores numéricos de cumplimiento.

    Retorna:
        Acumulado como valor decimal.

    Ejemplo:
        [0.95, 1.00, 0.97] → 0.973333...
    """

    if not cumplimientos:
        return None

    valores_validos = []

    for cumplimiento in cumplimientos:

        if cumplimiento is None:
            continue

        if not isinstance(cumplimiento, (int, float)):
            continue

        valores_validos.append(float(cumplimiento))

    if not valores_validos:
        return None

    return sum(valores_validos) / len(valores_validos)


def calcular_acumulado_desde_historico(registros_historicos, codigo):
    """
    Calcular el acumulado de un indicador a partir de registros históricos.

    Parámetros:
        registros_historicos: lista de registros históricos.
        codigo: código del indicador.

    Retorna:
        Acumulado como valor decimal.

    Reglas:
        - Solo considera el código solicitado.
        - Solo considera registros con Estado = OK.
        - Solo considera cumplimientos numéricos.
        - El acumulado es el promedio simple de los meses válidos.
        - Si no existen meses válidos, retorna None.
    """

    cumplimientos = []

    for registro in registros_historicos:

        if int(registro["codigo"]) != int(codigo):
            continue

        estado = registro.get("estado")

        if estado is None:
            continue

        if str(estado).strip().upper() != "OK":
            continue

        cumplimiento = registro.get("cumplimiento")

        if cumplimiento is None:
            continue

        if not isinstance(cumplimiento, (int, float)):
            continue

        cumplimientos.append(float(cumplimiento))

    return calcular_acumulado(cumplimientos)

def agregar_acumulado(resultado, registros_historicos):
    """
    Agregar el acumulado histórico a un resultado calculado.

    Parámetros:
        resultado: diccionario generado por calcular_indicador().
        registros_historicos: registros históricos disponibles.

    Retorna:
        Nuevo diccionario con el campo 'acumulado'.

    Esta función no modifica el resultado original.
    """

    resultado_final = resultado.copy()

    codigo = resultado["codigo"]

    acumulado = calcular_acumulado_desde_historico(
        registros_historicos,
        codigo
    )

    resultado_final["acumulado"] = acumulado

    return resultado_final


def preparar_registros_historico(resultados, indicadores, anio, mes):
    """
    Preparar los resultados mensuales para ser almacenados
    en la hoja HISTORICO.

    Esta función NO escribe en Excel.

    Parámetros:
        resultados: resultados generados por calcular_indicadores().
        indicadores: indicadores integrados y validados.
        anio: año del período procesado.
        mes: mes del período procesado.

    Retorna:
        Lista de registros con la estructura requerida por HISTORICO.
    """

    indicadores_por_codigo = {
        int(indicador["codigo"]): indicador
        for indicador in indicadores
    }

    registros = []

    for resultado in resultados:

        codigo = int(resultado["codigo"])

        if codigo not in indicadores_por_codigo:
            raise ValueError(
                f"El código {codigo} del resultado "
                "no existe en los indicadores."
            )

        indicador = indicadores_por_codigo[codigo]

        registro = {
            "anio": int(anio),
            "mes": int(mes),
            "codigo": codigo,
            "indicador": resultado["indicador"],
            "objetivo": int(indicador["objetivo"]),
            "tipo": resultado["tipo"],
            "meta": float(resultado["meta"]),
            "real": float(resultado["real"]),
            "cumplimiento": float(resultado["cumplimiento"]),
            "peso": float(resultado["peso"]),
            "resultado_ponderado": float(
                resultado["resultado_ponderado"]
            ),
            "fecha_proceso": datetime.now(),
            "estado": "OK",
        }

        registros.append(registro)

    return registros
