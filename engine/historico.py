"""
Módulo para gestión del histórico mensual de indicadores.

Responsabilidades:
    - Leer registros históricos.
    - Validar existencia de períodos.
    - Detectar duplicados Año + Mes + Código.
    - Guardar resultados mensuales.
    - Obtener cumplimientos históricos para acumulados.

Este módulo NO realiza:
    - Validaciones de indicadores.
    - Cálculos de cumplimiento.
    - Cálculos de pesos.
    - Cálculo del acumulado.

Esas responsabilidades pertenecen a otros módulos.
"""

import win32com.client as win32
from datetime import datetime

from config import (
    RUTA_EXCEL,
    HOJA_HISTORICO,
)

def leer_historico(libro=None):
    """Leer el histórico en bloque para reducir llamadas COM a Excel."""
    excel = None
    libro_propio = False
    try:
        if libro is None:
            excel = win32.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            libro = excel.Workbooks.Open(str(RUTA_EXCEL), ReadOnly=True)
            libro_propio = True

        hoja = libro.Worksheets(HOJA_HISTORICO)
        ultima_fila = hoja.Cells(hoja.Rows.Count, 3).End(-4162).Row
        if ultima_fila < 2:
            return []

        datos = hoja.Range(f"A2:M{ultima_fila}").Value
        if not datos:
            return []
        # Excel COM devuelve normalmente una tupla de tuplas.
        if len(datos) == 13 and not isinstance(datos[0], (tuple, list)):
            datos = (datos,)

        nombres = (
            "anio", "mes", "codigo", "indicador", "objetivo", "tipo",
            "meta", "real", "cumplimiento", "peso",
            "resultado_ponderado", "fecha_proceso", "estado",
        )
        registros = []
        for fila in datos:
            if not fila or fila[2] is None or str(fila[2]).strip() == "":
                continue
            registros.append(dict(zip(nombres, fila)))
        return registros
    finally:
        if libro_propio and libro is not None:
            libro.Close(SaveChanges=False)
        if libro_propio and excel is not None:
            excel.Quit()


def periodo_tiene_registros(anio, mes, libro=None):
    """
    Determinar si un período ya tiene registros en HISTORICO.

    Un período se considera existente si tiene al menos
    un registro para el Año y Mes consultados.

    Parámetros:
        anio: Año del período.
        mes: Mes del período.

    Retorna:
        True  -> el período tiene uno o más registros.
        False -> el período no tiene registros.
    """

    registros = leer_historico(libro)

    anio = int(anio)
    mes = int(mes)

    for registro in registros:

        registro_anio = int(registro["anio"])
        registro_mes = int(registro["mes"])

        if registro_anio == anio and registro_mes == mes:
            return True

    return False


def verificar_periodo_existente(anio, mes, codigo):
    """
    Verificar si ya existe un registro para la combinación
    Año + Mes + Código en la hoja HISTORICO.

    Retorna:
        True  -> si el período ya existe.
        False -> si el período no existe.
    """

    registros = leer_historico()

    for registro in registros:
        registro_anio = int(registro["anio"])
        registro_mes = int(registro["mes"])
        registro_codigo = int(registro["codigo"])

        if (
            registro_anio == int(anio)
            and registro_mes == int(mes)
            and registro_codigo == int(codigo)
        ):
            return True

    return False

def _eliminar_registros_periodo(hoja, anio, mes, codigos=None):
    """Eliminar todos los registros del período o solo los códigos indicados.

    Si ``codigos`` se entrega, conserva intactos los demás indicadores del mes.
    La lectura de A:C se realiza en bloque; solo se borran filas coincidentes.
    """
    anio, mes = int(anio), int(mes)
    codigos_objetivo = None if codigos is None else {int(c) for c in codigos}
    ultima_fila = hoja.Cells(hoja.Rows.Count, 1).End(-4162).Row
    if ultima_fila < 2:
        return 0

    datos = hoja.Range(f"A2:C{ultima_fila}").Value
    if not datos:
        return 0
    if len(datos) == 3 and not isinstance(datos[0], (tuple, list)):
        datos = (datos,)

    filas_borrar = []
    for indice, fila in enumerate(datos, start=2):
        try:
            fila_anio, fila_mes, fila_codigo = int(fila[0]), int(fila[1]), int(fila[2])
        except (TypeError, ValueError, IndexError):
            continue
        if fila_anio == anio and fila_mes == mes and (
            codigos_objetivo is None or fila_codigo in codigos_objetivo
        ):
            filas_borrar.append(indice)

    # Borrar desde abajo para que no cambien los índices pendientes.
    for fila in reversed(filas_borrar):
        hoja.Rows(fila).Delete()
    return len(filas_borrar)


def reemplazar_periodo(registros_nuevos, anio, mes):
    """
    Reemplazar completamente los registros de un período
    en HISTORICO.

    El proceso se ejecuta en una única sesión de Excel:

        1. Abrir el libro.
        2. Eliminar todos los registros del período.
        3. Insertar los nuevos registros.
        4. Guardar el libro.
        5. Cerrar Excel.

    Si ocurre un error antes de guardar, el libro se cierra
    sin guardar cambios para evitar una actualización parcial.

    Parámetros:
        registros_nuevos: lista de registros a insertar.
        anio: año del período.
        mes: mes del período.

    Retorna:
        Diccionario con:
            eliminados: cantidad de registros eliminados.
            insertados: cantidad de registros insertados.
    """

    if not registros_nuevos:
        raise ValueError(
            "No existen registros nuevos para reemplazar "
            "el período."
        )

    anio = int(anio)
    mes = int(mes)

    # ---------------------------------------------------------
    # Validar que todos los registros pertenecen al período
    # que se pretende reemplazar.
    # ---------------------------------------------------------

    for registro in registros_nuevos:

        registro_anio = int(registro["anio"])
        registro_mes = int(registro["mes"])

        if (
            registro_anio != anio
            or registro_mes != mes
        ):
            raise ValueError(
                "REGISTRO_PERIODO_INCONSISTENTE: "
                f"Se esperaba {anio}-{mes:02d}, "
                f"pero se encontró "
                f"{registro_anio}-{registro_mes:02d}."
            )

    # ---------------------------------------------------------
    # Validar códigos duplicados dentro del nuevo período.
    # ---------------------------------------------------------

    claves = set()

    for registro in registros_nuevos:

        codigo = int(registro["codigo"])

        clave = (
            anio,
            mes,
            codigo
        )

        if clave in claves:
            raise ValueError(
                "CODIGO_DUPLICADO_EN_REEMPLAZO: "
                f"Año={anio}, Mes={mes}, Código={codigo}"
            )

        claves.add(clave)

    excel = None
    libro = None

    try:

        # -----------------------------------------------------
        # Abrir Excel
        # -----------------------------------------------------

        excel = win32.DispatchEx(
            "Excel.Application"
        )

        excel.Visible = False
        excel.DisplayAlerts = False

        libro = excel.Workbooks.Open(
            str(RUTA_EXCEL),
            UpdateLinks=0,
            ReadOnly=False
        )

        hoja = libro.Worksheets(
            HOJA_HISTORICO
        )

        # -----------------------------------------------------
        # Eliminar período existente
        # -----------------------------------------------------

        eliminados = _eliminar_registros_periodo(
            hoja,
            anio,
            mes
        )

        # -----------------------------------------------------
        # Determinar siguiente fila disponible
        # -----------------------------------------------------

        ultima_fila = hoja.Cells(
            hoja.Rows.Count,
            1
        ).End(-4162).Row

        if ultima_fila < 2:
            siguiente_fila = 2
        else:
            siguiente_fila = ultima_fila + 1

        # -----------------------------------------------------
        # Insertar nuevos registros
        # -----------------------------------------------------

        fila = siguiente_fila

        for registro in registros_nuevos:

            hoja.Cells(
                fila,
                1
            ).Value = int(registro["anio"])

            hoja.Cells(
                fila,
                2
            ).Value = int(registro["mes"])

            hoja.Cells(
                fila,
                3
            ).Value = int(registro["codigo"])

            hoja.Cells(
                fila,
                4
            ).Value = registro["indicador"]

            hoja.Cells(
                fila,
                5
            ).Value = int(registro["objetivo"])

            hoja.Cells(
                fila,
                6
            ).Value = registro["tipo"]

            hoja.Cells(
                fila,
                7
            ).Value = float(registro["meta"])

            hoja.Cells(
                fila,
                8
            ).Value = float(registro["real"])

            hoja.Cells(
                fila,
                9
            ).Value = float(
                registro["cumplimiento"]
            )

            hoja.Cells(
                fila,
                10
            ).Value = float(
                registro["peso"]
            )

            hoja.Cells(
                fila,
                11
            ).Value = float(
                registro["resultado_ponderado"]
            )

            # Fecha de proceso
            fecha_proceso = registro.get(
                "fecha_proceso"
            )

            if isinstance(
                fecha_proceso,
                datetime
            ):
                hoja.Cells(
                    fila,
                    12
                ).Value = fecha_proceso

            else:
                hoja.Cells(
                    fila,
                    12
                ).Value = datetime.now()

            hoja.Cells(
                fila,
                13
            ).Value = registro.get(
                "estado",
                "OK"
            )

            fila += 1

        # -----------------------------------------------------
        # Guardar únicamente después de completar
        # eliminación + inserción.
        # -----------------------------------------------------

        libro.Save()

        return {
            "eliminados": eliminados,
            "insertados": len(
                registros_nuevos
            ),
        }

    except Exception:

        # Si algo falla antes de guardar, no conservar
        # cambios parciales.
        if libro is not None:

            try:
                libro.Close(
                    SaveChanges=False
                )
            except Exception:
                pass

            libro = None

        raise

    finally:

        if libro is not None:

            try:
                libro.Close(
                    SaveChanges=False
                )
            except Exception:
                pass

        if excel is not None:

            try:
                excel.Quit()
            except Exception:
                pass


def guardar_historico(registros_nuevos):
    """
    Guardar registros mensuales en la hoja HISTORICO.

    Cada registro debe contener:
        anio
        mes
        codigo
        indicador
        objetivo
        tipo
        meta
        real
        cumplimiento
        peso
        resultado_ponderado
        fecha_proceso
        estado

    La clave única del histórico es:
        Año + Mes + Código

    Si alguno de los registros ya existe, el proceso se bloquea
    y no se escribe ningún registro.
    """

    if not registros_nuevos:
        raise ValueError("No existen registros para guardar.")

    # ---------------------------------------------------------
    # 1. Validar duplicados dentro de los registros nuevos
    # ---------------------------------------------------------

    claves_nuevas = set()

    for registro in registros_nuevos:

        anio = int(registro["anio"])
        mes = int(registro["mes"])
        codigo = int(registro["codigo"])

        clave = (anio, mes, codigo)

        if clave in claves_nuevas:
            raise ValueError(
                f"PERIODO_DUPLICADO_EN_ENTRADA: "
                f"Año={anio}, Mes={mes}, Código={codigo}"
            )

        claves_nuevas.add(clave)

    # ---------------------------------------------------------
    # 2. Validar que ningún período ya exista
    # ---------------------------------------------------------

    for anio, mes, codigo in claves_nuevas:

        if verificar_periodo_existente(anio, mes, codigo):
            raise ValueError(
                f"PERIODO_YA_PROCESADO: "
                f"Año={anio}, Mes={mes}, Código={codigo}"
            )

    # ---------------------------------------------------------
    # 3. Abrir Excel
    # ---------------------------------------------------------

    excel = None
    libro = None

    try:

        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        libro = excel.Workbooks.Open(
            str(RUTA_EXCEL)
        )

        hoja = libro.Worksheets(HOJA_HISTORICO)

        # -----------------------------------------------------
        # 4. Determinar primera fila disponible
        # -----------------------------------------------------

        ultima_fila = hoja.Cells(
            hoja.Rows.Count,
            1
        ).End(-4162).Row  # xlUp

        if ultima_fila < 2:
            siguiente_fila = 2
        else:
            siguiente_fila = ultima_fila + 1

        # -----------------------------------------------------
        # 5. Escribir registros
        # -----------------------------------------------------

        for registro in registros_nuevos:

            fila = siguiente_fila

            hoja.Cells(fila, 1).Value = registro["anio"]
            hoja.Cells(fila, 2).Value = registro["mes"]
            hoja.Cells(fila, 3).Value = registro["codigo"]
            hoja.Cells(fila, 4).Value = registro["indicador"]
            hoja.Cells(fila, 5).Value = registro["objetivo"]
            hoja.Cells(fila, 6).Value = registro["tipo"]
            hoja.Cells(fila, 7).Value = registro["meta"]
            hoja.Cells(fila, 8).Value = registro["real"]
            hoja.Cells(fila, 9).Value = registro["cumplimiento"]
            hoja.Cells(fila, 10).Value = registro["peso"]
            hoja.Cells(fila, 11).Value = registro["resultado_ponderado"]
            fecha_proceso = registro["fecha_proceso"]

            if isinstance(fecha_proceso, datetime):
                fecha_excel = (
                    fecha_proceso - datetime(1899, 12, 30)
                ).total_seconds() / 86400

                hoja.Cells(fila, 12).Value = fecha_excel
            else:
                hoja.Cells(fila, 12).Value = fecha_proceso

            hoja.Cells(fila, 13).Value = registro["estado"]

            siguiente_fila += 1

        # -----------------------------------------------------
        # 6. Guardar
        # -----------------------------------------------------

        libro.Save()

    except Exception:
        if libro is not None:
            libro.Close(SaveChanges=False)
        raise

    finally:

        if libro is not None:
            try:
                libro.Close(SaveChanges=False)
            except Exception:
                pass

        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass

def obtener_cumplimientos_historicos(codigo):
    """
    Obtener los cumplimientos históricos válidos de un indicador.

    Parámetros:
        codigo: código del indicador.

    Retorna:
        Lista de cumplimientos válidos, ordenados por Año y Mes.

    Solo se consideran:
        - Registros del código solicitado.
        - Estado = "OK".
        - Cumplimiento numérico.
    """

    registros = leer_historico()

    cumplimientos = []

    for registro in registros:

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

        cumplimientos.append({
            "anio": int(registro["anio"]),
            "mes": int(registro["mes"]),
            "cumplimiento": float(cumplimiento),
        })

    cumplimientos.sort(
        key=lambda x: (x["anio"], x["mes"])
    )

    return cumplimientos
