import win32com.client as win32

from config import (
    RUTA_EXCEL,
    HOJA_RESULTADOS,
    FILA_INICIO_RESULTADOS,
)


# ============================================================
# OBTENER MAPA DINÁMICO CÓDIGO → FILA
# ============================================================

def obtener_filas_resultados(hoja_resultados):

    filas_resultado = {}

    fila = FILA_INICIO_RESULTADOS

    while True:

        valor_codigo = hoja_resultados.Range(f"A{fila}").Value

        # Fin de la tabla
        if valor_codigo is None or str(valor_codigo).strip() == "":
            break

        try:
            codigo = int(valor_codigo)

        except (ValueError, TypeError):

            raise ValueError(
                f"El código de la fila {fila} de "
                f"{HOJA_RESULTADOS} no es válido: {valor_codigo}"
            )

        # Código duplicado
        if codigo in filas_resultado:

            raise ValueError(
                f"El código {codigo} aparece más de una vez "
                f"en {HOJA_RESULTADOS}."
            )

        filas_resultado[codigo] = fila

        fila += 1

    return filas_resultado


# ============================================================
# ESCRIBIR RESULTADOS EN EXCEL
# ============================================================

def escribir_resultados(resultados):

    excel = None
    libro = None

    cantidad_escrita = 0

    try:

        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False

        libro = excel.Workbooks.Open(
            Filename=str(RUTA_EXCEL),
            ReadOnly=False
        )

        hoja_resultados = libro.Worksheets(HOJA_RESULTADOS)

        # Crear mapa dinámico código → fila
        filas_resultado = obtener_filas_resultados(
            hoja_resultados
        )

        # ----------------------------------------------------
        # ESCRIBIR CADA RESULTADO
        # ----------------------------------------------------

        for resultado in resultados:

            codigo = int(resultado["codigo"])

            if codigo not in filas_resultado:

                raise ValueError(
                    f"No existe una fila configurada para el "
                    f"código {codigo} en {HOJA_RESULTADOS}."
                )

            fila = filas_resultado[codigo]

            cumplimiento = resultado["cumplimiento"]
            resultado_ponderado = resultado["resultado_ponderado"]
            acumulado = resultado["acumulado"]

            hoja_resultados.Range(
                f"E{fila}"
            ).Value = cumplimiento

            hoja_resultados.Range(
                f"F{fila}"
            ).Value = resultado_ponderado

            hoja_resultados.Range(
                f"G{fila}"
            ).Value = acumulado

            cantidad_escrita += 1

        # ----------------------------------------------------
        # GUARDAR
        # ----------------------------------------------------

        libro.Save()

        return cantidad_escrita

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