"""
TBG Calidad - Motor de Validaciones
Paso 9

Responsabilidad:
    Validar los datos antes de realizar cualquier cálculo.

Regla fundamental:
    Un dato inválido NO se convierte automáticamente en 0.
    Se genera un error explícito para que pueda ser auditado.
"""

import math

def validar_entrada(datos_entrada):
    errores = []

    for dato in datos_entrada:

        codigo = dato.get("codigo")

        if codigo is None or str(codigo).strip() == "":
            errores.append({
                "codigo": "DATO_FALTANTE",
                "campo": "codigo",
                "mensaje": "El código es obligatorio."
            })
            continue

        if isinstance(codigo, bool) or not isinstance(codigo, (int, float)):
            errores.append({
                "codigo": "CODIGO_NO_NUMERICO",
                "campo": "codigo",
                "mensaje": "El código debe ser numérico.",
                "valor": codigo
            })
            continue

        try:
            codigo_numerico = float(codigo)
        except (TypeError, ValueError):
            errores.append({
                "codigo": "CODIGO_NO_NUMERICO",
                "campo": "codigo",
                "mensaje": "El código debe ser numérico.",
                "valor": codigo
            })
            continue

        if not math.isfinite(codigo_numerico) or not codigo_numerico.is_integer():
            errores.append({
                "codigo": "CODIGO_INVALIDO",
                "campo": "codigo",
                "mensaje": "El código debe ser un número entero finito.",
                "valor": codigo
            })
            continue

        codigo_normalizado = int(codigo_numerico)

        if codigo_normalizado <= 0:
            errores.append({
                "codigo": "CODIGO_INVALIDO",
                "campo": "codigo",
                "mensaje": "El código debe ser mayor que cero.",
                "valor": codigo
            })
            continue

        dato["codigo"] = codigo_normalizado

    return errores


def validar_codigos_duplicados(datos_entrada):
    errores = []
    codigos_vistos = {}

    for dato in datos_entrada:
        codigo = dato.get("codigo")

        if codigo is None:
            continue

        codigos_vistos.setdefault(codigo, 0)
        codigos_vistos[codigo] += 1

    for codigo, cantidad in codigos_vistos.items():
        if cantidad > 1:
            errores.append({
                "codigo": "CODIGO_DUPLICADO",
                "campo": "codigo",
                "mensaje": (
                    f"El código {codigo} aparece "
                    f"{cantidad} veces en 02_ENTRADA."
                ),
                "valor": codigo
            })

    return errores


def validar_coherencia_entrada_maestro(datos_entrada, maestro):
    """
    Valida que los datos descriptivos de 02_ENTRADA correspondan
    al indicador definido en el maestro.

    Campos validados:
        - objetivo
        - indicador
        - unidad

    El código funciona como llave maestra.

    Retorna:
        Lista de errores encontrados.
    """

    errores = []

    # Crear índice del maestro por código
    maestro_por_codigo = {
        int(item["codigo"]): item
        for item in maestro
    }

    for dato in datos_entrada:
        codigo = dato.get("codigo")

        # La existencia y validez del código
        # se valida en pasos anteriores.
        if codigo is None:
            continue

        if codigo not in maestro_por_codigo:
            continue

        maestro_indicador = maestro_por_codigo[codigo]

        # =====================================================
        # 1. OBJETIVO
        # =====================================================

        objetivo_entrada = dato.get("objetivo")
        objetivo_maestro = maestro_indicador.get("objetivo")

        if objetivo_entrada != objetivo_maestro:
            errores.append({
                "codigo": "OBJETIVO_NO_COINCIDE_MAESTRO",
                "campo": "objetivo",
                "mensaje": (
                    f"El objetivo del código {codigo} no coincide "
                    f"con el maestro. "
                    f"Entrada: {objetivo_entrada}. "
                    f"Maestro: {objetivo_maestro}."
                ),
                "valor": objetivo_entrada,
                "valor_maestro": objetivo_maestro,
                "indicador_codigo": codigo,
                "indicador_nombre": maestro_indicador.get("indicador")
            })

        # =====================================================
        # 2. INDICADOR
        # =====================================================

        indicador_entrada = dato.get("indicador")
        indicador_maestro = maestro_indicador.get("indicador")

        if indicador_entrada != indicador_maestro:
            errores.append({
                "codigo": "INDICADOR_NO_COINCIDE_MAESTRO",
                "campo": "indicador",
                "mensaje": (
                    f"El indicador del código {codigo} no coincide "
                    f"con el maestro. "
                    f"Entrada: {indicador_entrada}. "
                    f"Maestro: {indicador_maestro}."
                ),
                "valor": indicador_entrada,
                "valor_maestro": indicador_maestro,
                "indicador_codigo": codigo,
                "indicador_nombre": indicador_maestro
            })

        # =====================================================
        # 3. UNIDAD
        # =====================================================

        unidad_entrada = dato.get("unidad")
        unidad_maestro = maestro_indicador.get("unidad")

        if unidad_entrada != unidad_maestro:
            errores.append({
                "codigo": "UNIDAD_NO_COINCIDE_MAESTRO",
                "campo": "unidad",
                "mensaje": (
                    f"La unidad del código {codigo} no coincide "
                    f"con el maestro. "
                    f"Entrada: {unidad_entrada}. "
                    f"Maestro: {unidad_maestro}."
                ),
                "valor": unidad_entrada,
                "valor_maestro": unidad_maestro,
                "indicador_codigo": codigo,
                "indicador_nombre": indicador_maestro
            })

    return errores


def validar_codigos_entrada(datos_entrada, maestro, configuracion):
    errores = []

    codigos_maestro = {
        int(item["codigo"])
        for item in maestro
    }

    codigos_configuracion = set(configuracion.keys())

    for dato in datos_entrada:
        codigo = dato.get("codigo")

        if codigo not in codigos_maestro:
            errores.append({
                "codigo": "CODIGO_NO_EXISTE_MAESTRO",
                "campo": "codigo",
                "mensaje": (
                    f"El código {codigo} no existe en el "
                    "maestro de indicadores."
                ),
                "valor": codigo
            })

        if codigo not in codigos_configuracion:
            errores.append({
                "codigo": "CODIGO_NO_EXISTE_CONFIGURACION",
                "campo": "codigo",
                "mensaje": (
                    f"El código {codigo} no existe en "
                    "05_CONFIGURACION."
                ),
                "valor": codigo
            })

    return errores




def clasificar_estado_entrada(meta, real, tipo=None):
    """
    Clasifica el estado operativo de una entrada Meta/Real.

    Reglas:
        - Meta vacía + Real vacía -> PENDIENTE.
        - Una sola de las dos vacía -> DATO_FALTANTE.
        - Meta/Real presentes y válidas -> OK.
        - Meta o Real inválida -> ERROR con el código específico.

    Esta función es una capa previa a ``validar_indicador`` y se limita
    a clasificar los campos operativos Meta/Real. No calcula cumplimiento
    ni modifica la entrada.

    Retorna
    -------
    dict
        {
            "estado": "PENDIENTE" | "OK" | "DATO_FALTANTE" | "REVISION REAL - DIVISOR CERO" | "ERROR",
            "errores": list[dict]
        }
    """

    meta_vacia = meta is None or meta == ""
    real_vacio = real is None or real == ""

    # ---------------------------------------------------------
    # 1. Sin datos todavía
    # ---------------------------------------------------------
    if meta_vacia and real_vacio:
        return {
            "estado": "PENDIENTE",
            "errores": []
        }

    # ---------------------------------------------------------
    # 2. Solo uno de los dos datos está diligenciado
    # ---------------------------------------------------------
    if meta_vacia or real_vacio:
        campo = "meta" if meta_vacia else "real"
        mensaje = (
            "La meta está vacía."
            if meta_vacia
            else "El valor real está vacío."
        )
        return {
            "estado": "DATO_FALTANTE",
            "errores": [{
                "codigo": "DATO_FALTANTE",
                "campo": campo,
                "mensaje": mensaje,
            }]
        }

    errores = []

    # ---------------------------------------------------------
    # 3. Validación de Meta
    # ---------------------------------------------------------
    try:
        meta_num = float(meta)
        if not math.isfinite(meta_num) or meta_num <= 0:
            errores.append({
                "codigo": "META_INVALIDA",
                "campo": "meta",
                "mensaje": "La meta debe ser mayor que cero.",
            })
    except (TypeError, ValueError):
        errores.append({
            "codigo": "META_NO_NUMERICA",
            "campo": "meta",
            "mensaje": "La meta debe ser un valor numérico.",
        })

    # ---------------------------------------------------------
    # 4. Validación de Real
    # ---------------------------------------------------------
    try:
        real_num = float(real)
        if not math.isfinite(real_num):
            errores.append({
                "codigo": "REAL_NO_NUMERICO",
                "campo": "real",
                "mensaje": "El valor real debe ser un número finito.",
            })
        elif real_num < 0:
            errores.append({
                "codigo": "REAL_NEGATIVO",
                "campo": "real",
                "mensaje": "El valor real no puede ser negativo.",
            })
        #elif real_num == 0 and str(tipo).strip().upper() == "N":
        elif (
            real_num == 0
            and str(tipo).strip().upper() == "N"
        ):
            errores.append({
                "codigo": "DIVISOR_CERO_N",
                "campo": "real",
                "mensaje": "El valor real no puede ser cero para un indicador tipo N.",
            })
    except (TypeError, ValueError):
        errores.append({
            "codigo": "REAL_NO_NUMERICO",
            "campo": "real",
            "mensaje": "El valor real debe ser un valor numérico.",
        })

    # DIVISOR_CERO_N es una revisión operativa, no un bloqueo.
    # El indicador no se calcula y su estado visible será
    # "REVISION REAL - DIVISOR CERO". Otros errores sí siguen siendo bloqueantes.
    codigos_error = {error.get("codigo") for error in errores}

    if "DIVISOR_CERO_N" in codigos_error:
        errores_bloqueantes = [
            error
            for error in errores
            if error.get("codigo") != "DIVISOR_CERO_N"
        ]
        if not errores_bloqueantes:
            return {
                "estado": "REVISION REAL - DIVISOR CERO",
                "errores": errores,
            }
        return {
            "estado": "ERROR",
            "errores": errores,
        }

    return {
        "estado": "ERROR" if errores else "OK",
        "errores": errores,
    }


def validar_indicador(indicador):
    """
    Valida la información de un indicador.

    Parámetro
    ---------
    indicador : dict
        Diccionario con la información integrada del indicador.

    Retorna
    -------
    list
        Lista de errores encontrados.
        Si no existen errores, retorna [].
    """

    errores = []

    # ---------------------------------------------------------
    # 1. Código
    # ---------------------------------------------------------
    codigo = indicador.get("codigo")

    if codigo is None or codigo == "":
        errores.append({
            "codigo": "DATO_FALTANTE",
            "campo": "codigo",
            "mensaje": "El código del indicador está vacío."
        })

    # ---------------------------------------------------------
    # 2. Indicador
    # ---------------------------------------------------------
    nombre = indicador.get("indicador")

    if nombre is None or str(nombre).strip() == "":
        errores.append({
            "codigo": "DATO_FALTANTE",
            "campo": "indicador",
            "mensaje": "El nombre del indicador está vacío."
        })


    # ---------------------------------------------------------
    # 3. Unidad
    # ---------------------------------------------------------
    unidad = indicador.get("unidad")

    if unidad is None or str(unidad).strip() == "":
        errores.append({
            "codigo": "DATO_FALTANTE",
            "campo": "unidad",
            "mensaje": "La unidad del indicador está vacía."
        })

    # ---------------------------------------------------------
    # 4. Objetivo
    # ---------------------------------------------------------

    objetivo = indicador.get("objetivo")

    if objetivo is None or str(objetivo).strip() == "":
        errores.append({
            "codigo": "DATO_FALTANTE",
            "campo": "objetivo",
            "mensaje": "El objetivo es obligatorio."
        })
    elif not isinstance(objetivo, (int, float)):
        errores.append({
            "codigo": "OBJETIVO_NO_NUMERICO",
            "campo": "objetivo",
            "mensaje": "El objetivo debe ser numérico."
        })


    # ---------------------------------------------------------
    # 5. Meta
    # ---------------------------------------------------------
    meta = indicador.get("meta")

    if meta is None or meta == "":
        errores.append({
            "codigo": "DATO_FALTANTE",
            "campo": "meta",
            "mensaje": "La meta está vacía."
        })
    else:
        try:
            meta_num = float(meta)

            if not math.isfinite(meta_num) or meta_num <= 0:
                errores.append({
                    "codigo": "META_INVALIDA",
                    "campo": "meta",
                    "mensaje": "La meta debe ser mayor que cero."
                })

        except (TypeError, ValueError):
            errores.append({
                "codigo": "META_NO_NUMERICA",
                "campo": "meta",
                "mensaje": "La meta debe ser un valor numérico."
            })
    tipo = indicador.get("tipo")

    # ---------------------------------------------------------
    # 6. Real
    # ---------------------------------------------------------
    real = indicador.get("real")

    if real is None or real == "":
        errores.append({
            "codigo": "DATO_FALTANTE",
            "campo": "real",
            "mensaje": "El valor real está vacío."
        })
    else:
        try:
            real_num = float(real)

            if not math.isfinite(real_num):
                errores.append({
                    "codigo": "REAL_NO_NUMERICO",
                    "campo": "real",
                    "mensaje": "El valor real debe ser un número finito."
                })

            elif real_num < 0:
                errores.append({
                    "codigo": "REAL_NEGATIVO",
                    "campo": "real",
                    "mensaje": "El valor real no puede ser negativo."
                })

            tipo_normalizado = str(tipo).strip().upper()

            if real_num == 0 and tipo_normalizado == "N":
                errores.append({
                    "codigo": "DIVISOR_CERO_N",
                    "campo": "real",
                    "mensaje": "El valor real no puede ser cero para un indicador tipo N."
                })

        except (TypeError, ValueError):
            errores.append({
                "codigo": "REAL_NO_NUMERICO",
                "campo": "real",
                "mensaje": "El valor real debe ser un valor numérico."
            })

    # ---------------------------------------------------------
    # 7. Tipo P / N / N0
    # ---------------------------------------------------------


    if tipo is None or str(tipo).strip() == "":
        errores.append({
            "codigo": "TIPO_INVALIDO",
            "campo": "tipo",
            "mensaje": "El tipo del indicador debe ser P, N o N0."
        })
    else:
        tipo_normalizado = str(tipo).strip().upper()

        if tipo_normalizado not in ("P", "N", "N0"):
            errores.append({
                "codigo": "TIPO_INVALIDO",
                "campo": "tipo",
                "mensaje": "El tipo del indicador debe ser P, N o N0."
            })

    # ---------------------------------------------------------
    # 8. Peso
    # ---------------------------------------------------------
    peso = indicador.get("peso")

    if peso is None or peso == "":
        errores.append({
            "codigo": "DATO_FALTANTE",
            "campo": "peso",
            "mensaje": "El peso del indicador está vacío."
        })
    else:
        try:
            peso_num = float(peso)

            if not math.isfinite(peso_num) or peso_num <= 0:
                errores.append({
                    "codigo": "PESO_INVALIDO",
                    "campo": "peso",
                    "mensaje": "El peso debe ser mayor que cero."
                })

        except (TypeError, ValueError):
            errores.append({
                "codigo": "PESO_INVALIDO",
                "campo": "peso",
                "mensaje": "El peso debe ser un valor numérico."
            })

    # ---------------------------------------------------------
    # 9. Dirección / criterio
    # ---------------------------------------------------------
    criterio = indicador.get("criterio")

    if criterio is None or str(criterio).strip() == "":
        errores.append({
            "codigo": "DATO_FALTANTE",
            "campo": "criterio",
            "mensaje": "El criterio del indicador está vacío."
        })

    # ---------------------------------------------------------
    # 10. Validación de coherencia Tipo vs Criterio
    # ---------------------------------------------------------
    if tipo is not None and criterio is not None:

        tipo_normalizado = str(tipo).strip().upper()
        criterio_normalizado = str(criterio).strip().lower()

        if tipo_normalizado == "P":
            if criterio_normalizado != "mayor es mejor":
                errores.append({
                    "codigo": "TIPO_CRITERIO_INCONSISTENTE",
                    "campo": "tipo",
                    "mensaje": (
                        "El tipo P debe corresponder al criterio "
                        "'Mayor es mejor'."
                    )
                })

        elif tipo_normalizado in ("N", "N0"):
            if criterio_normalizado != "menor es mejor":
                errores.append({
                    "codigo": "TIPO_CRITERIO_INCONSISTENTE",
                    "campo": "tipo",
                    "mensaje": (
                        "El tipo N debe corresponder al criterio "
                        "'Menor es mejor'."
                    )
                })

    return errores


def validar_indicadores(indicadores, objetivos_validos=None):
    errores = []

    for indicador in indicadores:

        # =====================================================
        # 1. CLASIFICAR ESTADO OPERATIVO META / REAL
        # =====================================================

        estado_entrada = clasificar_estado_entrada(
            indicador.get("meta"),
            indicador.get("real"),
            indicador.get("tipo"),
        )

        estado = estado_entrada["estado"]

        # -----------------------------------------------------
        # 1.1 PENDIENTE
        # -----------------------------------------------------
        # Meta y Real están vacías.
        # No es un error y no debe generar errores de
        # DATO_FALTANTE para Meta/Real.
        #
        # Sin embargo, las demás validaciones estructurales
        # del indicador sí deben ejecutarse.
        # -----------------------------------------------------

        if estado == "PENDIENTE":

            errores_indicador = validar_indicador(
                {
                    **indicador,
                    "meta": 1,
                    "real": 0,
                }
            )

            # Eliminar únicamente los errores artificiales
            # asociados a Meta/Real que fueron generados
            # por los valores temporales usados arriba.
            errores_indicador = [
                error
                for error in errores_indicador
                if error.get("campo") not in ("meta", "real")
            ]

        # -----------------------------------------------------
        # 1.2 OK
        # -----------------------------------------------------
        elif estado == "OK":

            errores_indicador = validar_indicador(
                indicador
            )

        # -----------------------------------------------------
        # 1.3 DATO_FALTANTE / REVISION REAL / ERROR
        # -----------------------------------------------------
        else:

            errores_indicador = []

            # DATO_FALTANTE se conserva como estado operativo no bloqueante.
            if estado == "DATO_FALTANTE":
                errores_indicador.extend(estado_entrada["errores"])

            # Los datos realmente inválidos sí deben reportarse como errores.
            elif estado == "ERROR":
                errores_indicador.extend(estado_entrada["errores"])

            # REVISION REAL no agrega DIVISOR_CERO_N como error bloqueante;
            # el estado queda visible en 02_ENTRADA.
            real_seguro = 1 if estado.startswith("REVISION REAL") else 0

            errores_estructura = validar_indicador(
                {
                    **indicador,
                    "meta": 1,
                    "real": real_seguro,
                }
            )

            errores_estructura = [
                error
                for error in errores_estructura
                if error.get("campo") not in ("meta", "real")
            ]

            errores_indicador.extend(
                errores_estructura
            )

        # Los campos Meta/Real incompletos son estados operativos
        # (DATO_FALTANTE) y no deben bloquear los demás indicadores.
        # Se conservan los errores estructurales de DATO_FALTANTE
        # (por ejemplo código, indicador, unidad, peso, etc.).
        errores_indicador = [
            error
            for error in errores_indicador
            if not (
                error.get("codigo") == "DATO_FALTANTE"
                and error.get("campo") in ("meta", "real")
            )
        ]

        # =====================================================
        # 2. VALIDAR OBJETIVO
        # =====================================================

        objetivo = indicador.get("objetivo")

        if (
            objetivos_validos is not None
            and objetivo is not None
            and isinstance(objetivo, (int, float))
            and objetivo not in objetivos_validos
        ):
            errores_indicador.append({
                "codigo": "OBJETIVO_INVALIDO",
                "campo": "objetivo",
                "mensaje": "El objetivo no existe en la configuración."
            })

        # =====================================================
        # 3. AGREGAR CONTEXTO
        # =====================================================

        for error in errores_indicador:

            error["indicador_codigo"] = indicador.get(
                "codigo"
            )

            error["indicador_nombre"] = indicador.get(
                "indicador"
            )

            errores.append(error)

    return errores

def mostrar_errores(errores):
    """
    Muestra los errores de validación en consola.
    """

    if not errores:
        print("VALIDACIÓN OK - No se encontraron errores.")
        return

    print()
    print("=" * 70)
    print("ERRORES DE VALIDACIÓN")
    print("=" * 70)

    for error in errores:
        print(
            f"[{error['codigo']}] "
            f"Código {error.get('indicador_codigo')} - "
            f"{error.get('campo')}: "
            f"{error.get('mensaje')}"
        )

    print("=" * 70)
    print(f"Total de errores: {len(errores)}")


def validar_pesos_objetivos(
    indicadores,
    pesos_indicadores,
    pesos_objetivos,
):
    """
    Valida la consistencia de los pesos de los indicadores
    frente al peso definido para cada objetivo.

    Reglas:
    1. Cada indicador debe tener peso.
    2. Cada peso debe ser numérico y mayor que cero.
    3. Cada indicador debe pertenecer a un objetivo válido.
    4. La suma de pesos de indicadores por objetivo debe
       coincidir con el peso del objetivo.
    5. La suma de los pesos de objetivos debe ser 100%.
    """

    errores = []

    tolerancia = 0.000001

    objetivos_validos = set(pesos_objetivos.keys())

    suma_por_objetivo = {
        objetivo: 0.0
        for objetivo in objetivos_validos
    }

    # ---------------------------------------------------------
    # 1. Validar pesos individuales
    # ---------------------------------------------------------

    for codigo, datos in indicadores.items():

        if codigo not in pesos_indicadores:
            errores.append({
                "codigo": codigo,
                "error": "PESO_FALTANTE",
            })
            continue

        peso = pesos_indicadores[codigo]

        if peso is None or str(peso).strip() == "":
            errores.append({
                "codigo": codigo,
                "error": "PESO_FALTANTE",
            })
            continue

        if not isinstance(peso, (int, float)):
            errores.append({
                "codigo": codigo,
                "error": "PESO_NO_NUMERICO",
            })
            continue

        if not isinstance(peso, (int, float)):
            errores.append({
                "codigo": codigo,
                "error": "PESO_NO_NUMERICO",
            })
            continue

        if peso <= 0:
            errores.append({
                "codigo": codigo,
                "error": "PESO_INVALIDO",
            })
            continue

        objetivo = datos.get("objetivo")

        if objetivo not in objetivos_validos:
            errores.append({
                "codigo": codigo,
                "error": "OBJETIVO_INVALIDO",
            })
            continue

        suma_por_objetivo[objetivo] += peso

    # ---------------------------------------------------------
    # 2. Validar suma de pesos por objetivo
    # ---------------------------------------------------------

    for objetivo, peso_objetivo in pesos_objetivos.items():

        suma_indicadores = suma_por_objetivo.get(
            objetivo,
            0.0
        )

        diferencia = (
            suma_indicadores - peso_objetivo
        )

        if abs(diferencia) > tolerancia:

            errores.append({
                "objetivo": objetivo,
                "error": "PESO_OBJETIVO_INCONSISTENTE",
                "suma_indicadores": suma_indicadores,
                "peso_objetivo": peso_objetivo,
            })

    # ---------------------------------------------------------
    # 3. Validar total de pesos de objetivos
    # ---------------------------------------------------------

    total_objetivos = sum(
        pesos_objetivos.values()
    )

    if abs(total_objetivos - 1.0) > tolerancia:

        errores.append({
            "error": "PESO_OBJETIVO_TOTAL_INVALIDO",
            "total": total_objetivos,
        })

    return errores
