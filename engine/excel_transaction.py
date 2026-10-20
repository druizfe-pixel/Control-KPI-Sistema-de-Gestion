"""Actualización atómica del histórico y de los resultados del libro."""

from datetime import datetime

import win32com.client as win32

from config import (
    FILA_INICIO_RESULTADOS,
    HOJA_AUDITORIA,
    HOJA_HISTORICO,
    HOJA_RESULTADOS,
    RUTA_EXCEL,
)
from historico import _eliminar_registros_periodo


def _filas_resultados(hoja):
    filas = {}
    fila = FILA_INICIO_RESULTADOS

    while True:
        valor = hoja.Range(f"A{fila}").Value
        if valor is None or str(valor).strip() == "":
            return filas

        try:
            codigo = int(valor)
        except (TypeError, ValueError) as error:
            raise ValueError(
                f"El código de {HOJA_RESULTADOS}!A{fila} no es válido: {valor}"
            ) from error

        if codigo in filas:
            raise ValueError(
                f"El código {codigo} aparece más de una vez en {HOJA_RESULTADOS}."
            )

        filas[codigo] = fila
        fila += 1

def sincronizar_resultados(hoja_resultados, indicadores_activos, fila_inicio=FILA_INICIO_RESULTADOS):
    """Sincronizar A:B por código preservando C:G del mismo indicador."""
    catalogo = {}
    for indicador in indicadores_activos:
        try:
            codigo = int(indicador["codigo"])
        except (KeyError, TypeError, ValueError) as error:
            raise ValueError(f"Código de indicador inválido: {indicador.get('codigo')}") from error
        if codigo in catalogo:
            raise ValueError(f"El código {codigo} aparece más de una vez en los indicadores activos.")
        catalogo[codigo] = indicador

    anteriores = {}
    fila = fila_inicio
    while True:
        valor = hoja_resultados.Cells(fila, 1).Value
        if valor is None or str(valor).strip() == "":
            break
        try:
            codigo = int(valor)
        except (TypeError, ValueError) as error:
            raise ValueError(f"El código de {HOJA_RESULTADOS}!A{fila} no es válido: {valor}") from error
        valores = hoja_resultados.Range(f"C{fila}:G{fila}").Value
        if valores and isinstance(valores[0], (tuple, list)):
            valores = valores[0]
        anteriores[codigo] = tuple(valores or (None, None, None, None, None))
        fila += 1

    ultima_fila = max(hoja_resultados.Cells(hoja_resultados.Rows.Count, 1).End(-4162).Row, fila_inicio)
    hoja_resultados.Range(f"A{fila_inicio}:B{ultima_fila}").ClearContents()
    codigos = sorted(catalogo)
    for desplazamiento, codigo in enumerate(codigos):
        destino = fila_inicio + desplazamiento
        indicador = catalogo[codigo]
        hoja_resultados.Cells(destino, 1).Value = codigo
        hoja_resultados.Cells(destino, 2).Value = indicador.get("indicador")
        if codigo in anteriores:
            hoja_resultados.Range(f"C{destino}:G{destino}").Value = (anteriores[codigo],)

    sobrante = fila_inicio + len(codigos)
    if sobrante <= ultima_fila:
        hoja_resultados.Range(f"A{sobrante}:G{ultima_fila}").ClearContents()
    return {"activos": len(codigos), "codigos": codigos,
            "primera_fila": fila_inicio if codigos else None,
            "ultima_fila": fila_inicio + len(codigos) - 1 if codigos else None}


def _periodo_existe(hoja, anio, mes):
    ultima_fila = hoja.Cells(hoja.Rows.Count, 1).End(-4162).Row
    if ultima_fila < 2:
        return False
    datos = hoja.Range(f"A2:B{ultima_fila}").Value
    if not datos:
        return False
    if len(datos) == 2 and not isinstance(datos[0], (tuple, list)):
        datos = (datos,)
    return any(
        fila[0] is not None and fila[1] is not None
        and int(fila[0]) == int(anio) and int(fila[1]) == int(mes)
        for fila in datos
    )


def _codigos_periodo(hoja, anio, mes):
    """Obtener códigos guardados para un período leyendo A:C en bloque."""
    ultima_fila = hoja.Cells(hoja.Rows.Count, 1).End(-4162).Row
    if ultima_fila < 2:
        return set()
    datos = hoja.Range(f"A2:C{ultima_fila}").Value
    if not datos:
        return set()
    if len(datos) == 3 and not isinstance(datos[0], (tuple, list)):
        datos = (datos,)
    codigos = set()
    for fila in datos:
        try:
            if int(fila[0]) == int(anio) and int(fila[1]) == int(mes):
                codigos.add(int(fila[2]))
        except (TypeError, ValueError, IndexError):
            continue
    return codigos


def _validar_accion(hoja_historico, anio, mes, accion_periodo):
    existe = _periodo_existe(hoja_historico, anio, mes)
    if accion_periodo == "PROCESAR" and existe:
        raise RuntimeError(f"El período {anio}-{mes:02d} ya fue procesado.")
    if accion_periodo in {"COMPLETAR", "CORREGIR"} and not existe:
        raise RuntimeError(f"El período {anio}-{mes:02d} no existe para completar/corregir.")
    if accion_periodo not in {"PROCESAR", "COMPLETAR", "CORREGIR"}:
        raise ValueError(f"Acción de período no válida: {accion_periodo}")


def _insertar_historico(hoja, registros):
    """Insertar los registros históricos en un único bloque COM."""
    if not registros:
        return
    ultima_fila = hoja.Cells(hoja.Rows.Count, 1).End(-4162).Row
    fila_inicio = max(ultima_fila + 1, 2)
    base_fecha = datetime(1899, 12, 30)
    matriz = []
    for registro in registros:
        fecha = registro.get("fecha_proceso", datetime.now())
        if hasattr(fecha, "year") and hasattr(fecha, "hour"):
            fecha = fecha.replace(tzinfo=None)
            fecha = (fecha - base_fecha).total_seconds() / 86400
        matriz.append((
            int(registro["anio"]), int(registro["mes"]), int(registro["codigo"]),
            registro["indicador"], int(registro["objetivo"]), registro["tipo"],
            float(registro["meta"]), float(registro["real"]),
            float(registro["cumplimiento"]), float(registro["peso"]),
            float(registro["resultado_ponderado"]), float(fecha),
            registro.get("estado", "OK"),
        ))
    destino = hoja.Range(
        f"A{fila_inicio}:M{fila_inicio + len(matriz) - 1}"
    )
    destino.Value2 = tuple(matriz)
    hoja.Range(
        f"L{fila_inicio}:L{fila_inicio + len(matriz) - 1}"
    ).NumberFormat = "dd/mm/yyyy hh:mm:ss"


def _escribir_resultados(hoja, resultados, filas, anio, mes, accion_periodo):
    """
    Actualizar la hoja 03_RESULTADOS para el período procesado.

    A:B = catálogo fijo de indicadores.
    C:G = valores del período actual.

    C = Meta
    D = Real
    E = Cumplimiento
    F = Resultado Ponderado
    G = Acumulado
    """

    # ============================================================
    # 1. VALIDAR QUE TODOS LOS RESULTADOS TENGAN FILA
    # ============================================================

    for resultado in resultados:
        codigo = int(resultado["codigo"])

        if codigo not in filas:
            raise ValueError(
                f"No existe una fila configurada para el código "
                f"{codigo} en {HOJA_RESULTADOS}."
            )

    # ============================================================
    # 2. LIMPIAR LOS VALORES DEL PERÍODO ANTERIOR
    # ============================================================

    # Los 28 indicadores ocupan las filas 6:33.
    ultima_fila_resultados = max(filas.values())

    if accion_periodo == "PROCESAR" and ultima_fila_resultados >= FILA_INICIO_RESULTADOS:
        hoja.Range(
            f"C{FILA_INICIO_RESULTADOS}:G{ultima_fila_resultados}"
        ).ClearContents()

    # ============================================================
    # 3. ESCRIBIR PERÍODO
    # ============================================================

    hoja.Range("B3").Value = f"{anio}-{mes:02d}"

    # ============================================================
    # 4. ESCRIBIR RESULTADOS
    # ============================================================

    total_resultado_ponderado = 0.0

    for resultado in resultados:
        codigo = int(resultado["codigo"])
        fila = filas[codigo]

        meta = float(resultado["meta"])
        real = float(resultado["real"])
        cumplimiento = float(resultado["cumplimiento"])
        resultado_ponderado = float(
            resultado["resultado_ponderado"]
        )
        acumulado = resultado.get("acumulado")

        # Meta
        hoja.Range(
            f"C{fila}"
        ).Value = meta

        # Real
        hoja.Range(
            f"D{fila}"
        ).Value = real

        # Cumplimiento
        hoja.Range(
            f"E{fila}"
        ).Value = cumplimiento

        # Resultado ponderado
        hoja.Range(
            f"F{fila}"
        ).Value = resultado_ponderado

        # Acumulado
        if acumulado is None:
            hoja.Range(
                f"G{fila}"
            ).ClearContents()
        else:
            hoja.Range(
                f"G{fila}"
            ).Value = float(acumulado)

        total_resultado_ponderado += resultado_ponderado

    # ============================================================
    # 5. ESCRIBIR RESULTADO TOTAL
    # ============================================================

    # En corrección parcial, E3 debe incluir los resultados que ya estaban
    # guardados y que no se modificaron en esta ejecución.
    if ultima_fila_resultados >= FILA_INICIO_RESULTADOS:
        valores_pesados = hoja.Range(
            f"F{FILA_INICIO_RESULTADOS}:F{ultima_fila_resultados}"
        ).Value
        if valores_pesados:
            if len(valores_pesados) == 1 and not isinstance(valores_pesados[0], (tuple, list)):
                valores_pesados = (valores_pesados,)
            total_resultado_ponderado = sum(
                float(fila[0]) for fila in valores_pesados
                if fila and isinstance(fila[0], (int, float))
            )
    hoja.Range("E3").Value = total_resultado_ponderado


def _registrar_auditoria(hoja, anio, mes, modo, accion_periodo, cantidad):
    """Agregar un evento exitoso al registro visible de auditoría."""
    ultima_fila = hoja.Cells(hoja.Rows.Count, 1).End(-4162).Row
    fila = max(ultima_fila + 1, 6)
    celda_fecha = hoja.Cells(fila, 1)
    fecha_local = datetime.now()
    serial = (
        fecha_local - datetime(1899, 12, 30)
    ).total_seconds() / 86400
    celda_fecha.Value2 = serial
    celda_fecha.NumberFormat = "dd/mm/yyyy hh:mm:ss"
    hoja.Cells(fila, 2).Value = f"{anio}-{mes:02d}"
    hoja.Cells(fila, 3).Value = modo
    hoja.Cells(fila, 4).Value = "OK"
    hoja.Cells(fila, 5).Value = (
        f"{accion_periodo}: {cantidad} indicadores procesados."
    )


def guardar_y_escribir_resultados(
    registros,
    resultados,
    anio,
    mes,
    accion_periodo,
    modo,
    indicadores_activos=None,
    libro=None
):
    """Guardar histórico y salida en una única operación confirmada por Excel."""
    if not registros:
        raise ValueError("No existen registros históricos para guardar.")
    if not resultados:
        raise ValueError("No existen resultados para escribir.")

    excel = None
    libro_propio = False

    try:
        if libro is None:
            excel = win32.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False

            libro = excel.Workbooks.Open(
                str(RUTA_EXCEL),
                UpdateLinks=0,
                ReadOnly=False
            )

            libro_propio = True

        hoja_historico = libro.Worksheets(HOJA_HISTORICO)
        hoja_resultados = libro.Worksheets(HOJA_RESULTADOS)
        hoja_auditoria = libro.Worksheets(HOJA_AUDITORIA)
        _validar_accion(hoja_historico, anio, mes, accion_periodo)
        if accion_periodo == "COMPLETAR":
            existentes = _codigos_periodo(hoja_historico, anio, mes)
            duplicados = existentes.intersection(int(registro["codigo"]) for registro in registros)
            if duplicados:
                raise RuntimeError(
                    "No se puede completar porque estos códigos ya tienen histórico: "
                    + ", ".join(map(str, sorted(duplicados)))
                )
        if indicadores_activos is not None:
            sincronizar_resultados(
                hoja_resultados,
                indicadores_activos,
            )

        filas = _filas_resultados(hoja_resultados)
        _escribir_resultados(
            hoja_resultados,
            resultados,
            filas,
            anio,
            mes,
            accion_periodo,
        )

        eliminados = 0
        if accion_periodo == "CORREGIR":
            codigos_corregidos = {int(registro["codigo"]) for registro in registros}
            eliminados = _eliminar_registros_periodo(
                hoja_historico, anio, mes, codigos=codigos_corregidos
            )
        _insertar_historico(hoja_historico, registros)
        _registrar_auditoria(
            hoja_auditoria,
            anio,
            mes,
            modo,
            accion_periodo,
            len(resultados),
        )

        libro.Save()
        return {
            "eliminados": eliminados,
            "insertados": len(registros),
            "resultados_escritos": len(resultados),
        }
    except Exception:
        if libro_propio and libro is not None:
            libro.Close(SaveChanges=False)
            libro = None
        raise

    finally:
        if libro_propio and libro is not None:
            libro.Close(SaveChanges=False)

        if libro_propio and excel is not None:
            excel.Quit()
