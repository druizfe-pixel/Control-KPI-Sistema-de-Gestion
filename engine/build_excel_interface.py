"""Construye la interfaz operativa del libro TBG Calidad mediante Excel COM.

Use siempre TBG_RUTA_EXCEL para apuntar a una copia de prueba.
"""

from __future__ import annotations

import win32com.client as win32

from config import (
    HOJA_AUDITORIA,
    HOJA_ENTRADA,
    HOJA_INICIO,
    HOJA_RESULTADOS,
    RUTA_EXCEL,
)


AZUL = 0x7F6000
AZUL_CLARO = 0xEAF2F8
AMARILLO = 0xCCFFFF
VERDE = 0xE2F0D9
GRIS = 0xF2F2F2
BLANCO = 0xFFFFFF


def color(rgb):
    rojo, verde, azul = rgb
    return rojo + (verde << 8) + (azul << 16)


def aplicar_titulo(hoja, rango, texto):
    hoja.Range(rango).Merge()
    celda = hoja.Range(rango.split(":")[0])
    celda.Value = texto
    celda.Font.Name = "Arial"
    celda.Font.Size = 14
    celda.Font.Bold = True
    celda.Font.Color = BLANCO
    celda.Interior.Color = AZUL
    celda.HorizontalAlignment = -4131  # xlLeft
    celda.VerticalAlignment = -4108  # xlCenter
    hoja.Range(rango).RowHeight = 26


def aplicar_encabezado(hoja, rango):
    bloque = hoja.Range(rango)
    bloque.Font.Name = "Arial"
    bloque.Font.Size = 10
    bloque.Font.Bold = True
    bloque.Font.Color = BLANCO
    bloque.Interior.Color = AZUL
    bloque.HorizontalAlignment = -4108  # xlCenter
    bloque.VerticalAlignment = -4108


def construir_inicio(hoja):
    hoja.Range("B2:O12").Font.Name = "Arial"
    aplicar_titulo(hoja, "B2:O2", "TBG Calidad - Control de indicadores")
    hoja.Range("B4:O4").Merge()
    hoja.Range("B4").Value = (
        "Defina el período, complete la entrada y ejecute el motor Python sobre este libro."
    )
    hoja.Range("B4").Font.Italic = True
    hoja.Range("B4").Font.Color = 0x666666

    hoja.Range("D6:D8").Font.Bold = True
    hoja.Range("E6:E8").Interior.Color = AMARILLO
    hoja.Range("E6:E8").Font.Bold = True
    hoja.Range("E6:E8").HorizontalAlignment = -4108
    hoja.Range("E6:E8").Borders.LineStyle = 1

    hoja.Range("E6").Validation.Delete()
    hoja.Range("E6").Validation.Add(1, 1, 1, "2026", "2099")
    hoja.Range("E7").Validation.Delete()
    hoja.Range("E7").Validation.Add(1, 1, 1, "1", "12")
    hoja.Range("E8").Validation.Delete()
    hoja.Range("E8").Validation.Add(3, 1, 1, '"NORMAL,CORRECCIÓN"')

    hoja.Range("H6:O6").Merge()
    hoja.Range("H6").Value = "Guía de uso"
    aplicar_encabezado(hoja, "H6:O6")
    guia = [
        "1. Seleccione Año, Mes y Modo.",
        "2. Complete Meta y Real en 02_ENTRADA.",
        "3. Ejecute el motor desde PowerShell.",
        "4. Consulte resultados y auditoría.",
    ]
    for fila, texto in enumerate(guia, start=7):
        hoja.Range(f"H{fila}:O{fila}").Merge()
        hoja.Range(f"H{fila}").Value = texto
        hoja.Range(f"H{fila}").Interior.Color = GRIS

    hoja.Range("B11:O11").Merge()
    hoja.Range("B11").Value = (
        "Campos amarillos: edición requerida. Modo NORMAL bloquea períodos ya procesados."
    )
    hoja.Range("B11").Font.Italic = True
    hoja.Range("B11").Font.Color = 0x666666
    hoja.Columns("B:O").ColumnWidth = 12
    hoja.Columns("B").ColumnWidth = 3
    hoja.Columns("H").ColumnWidth = 18
    hoja.Columns("O").ColumnWidth = 12
    hoja.Range("B2:O12").VerticalAlignment = -4108
    hoja.Range("B2:O12").HorizontalAlignment = -4131
    hoja.Range("E6:E8").HorizontalAlignment = -4108


def construir_entrada(hoja):
    aplicar_titulo(hoja, "A2:G2", "Entrada de indicadores")
    hoja.Range("A3:G3").Merge()
    hoja.Range("A3").Value = "Edite solo Meta y Real. Los demás campos se validan contra el maestro."
    hoja.Range("A3").Font.Italic = True
    hoja.Range("A3").Font.Color = 0x666666
    aplicar_encabezado(hoja, "A5:G5")
    hoja.Range("E6:F200").Interior.Color = AMARILLO
    hoja.Range("A6:D200").Interior.Color = GRIS
    hoja.Range("G6:G200").Interior.Color = VERDE
    hoja.Range("A5:G200").Font.Name = "Arial"
    hoja.Range("E6:F200").NumberFormat = "0.00"
    hoja.Columns("A").ColumnWidth = 10
    hoja.Columns("B").ColumnWidth = 10
    hoja.Columns("C").ColumnWidth = 42
    hoja.Columns("D").ColumnWidth = 12
    hoja.Columns("E:F").ColumnWidth = 14
    hoja.Columns("G").ColumnWidth = 16
    hoja.Range("A5:G200").VerticalAlignment = -4108


def construir_resultados(hoja):
    aplicar_titulo(hoja, "A2:G2", "Resultados del período")
    hoja.Range("A3").Value = "Período"
    hoja.Range("B3").Formula = '=TEXT(\'01_INICIO\'!E6,"0")&"-"&TEXT(\'01_INICIO\'!E7,"00")'
    hoja.Range("D3").Value = "Resultado total"
    hoja.Range("E3").Formula = "=SUM(F6:F241)"
    hoja.Range("G3").Value = "Valores calculados por el motor Python"
    hoja.Range("A3:E3").Font.Bold = True
    hoja.Range("B3").Interior.Color = AZUL_CLARO
    hoja.Range("E3").Interior.Color = VERDE
    hoja.Range("E3").NumberFormat = "0.00%"
    hoja.Range("G5").Value = "Acumulado"
    aplicar_encabezado(hoja, "A5:G5")
    hoja.Range("E6:G241").NumberFormat = "0.00%"
    hoja.Range("A5:G241").Font.Name = "Arial"
    hoja.Columns("A").ColumnWidth = 10
    hoja.Columns("B").ColumnWidth = 42
    hoja.Columns("C:D").ColumnWidth = 14
    hoja.Columns("E:G").ColumnWidth = 18
    hoja.Range("A5:G241").VerticalAlignment = -4108


def construir_auditoria(hoja):
    aplicar_titulo(hoja, "A2:E2", "Auditoría de ejecuciones")
    hoja.Range("A3:E3").Merge()
    hoja.Range("A3").Value = "El motor registra aquí cada ejecución completada correctamente."
    hoja.Range("A3").Font.Italic = True
    hoja.Range("A3").Font.Color = 0x666666
    hoja.Range("A5:E5").Value = [[
        "Fecha de proceso", "Período", "Modo", "Estado", "Detalle",
    ]]
    aplicar_encabezado(hoja, "A5:E5")
    hoja.Range("A6:A1000").NumberFormat = "yyyy-mm-dd hh:mm"
    hoja.Columns("A").ColumnWidth = 20
    hoja.Columns("B:D").ColumnWidth = 15
    hoja.Columns("E").ColumnWidth = 52
    hoja.Range("A5:E1000").Font.Name = "Arial"
    hoja.Range("A5:E1000").VerticalAlignment = -4108


def construir_interfaz():
    excel = None
    libro = None
    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        libro = excel.Workbooks.Open(str(RUTA_EXCEL), UpdateLinks=0, ReadOnly=False)
        construir_inicio(libro.Worksheets(HOJA_INICIO))
        construir_entrada(libro.Worksheets(HOJA_ENTRADA))
        construir_resultados(libro.Worksheets(HOJA_RESULTADOS))
        construir_auditoria(libro.Worksheets(HOJA_AUDITORIA))
        libro.Save()
    finally:
        if libro is not None:
            libro.Close(SaveChanges=False)
        if excel is not None:
            excel.Quit()


if __name__ == "__main__":
    construir_interfaz()
