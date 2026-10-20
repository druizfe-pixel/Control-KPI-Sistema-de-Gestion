from pathlib import Path
import os
import sys


# ============================================================
# CONFIGURACIÃ“N GENERAL DEL PROYECTO
# ============================================================

if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys.executable).resolve().parent
else:
    BASE_DIR = Path(__file__).resolve().parent.parent

# Por defecto se usa el libro operativo. Para pruebas, definir
# TBG_RUTA_EXCEL con la ruta absoluta de una copia .xlsm.
RUTA_EXCEL = Path(
    os.environ.get(
        "TBG_RUTA_EXCEL",
        BASE_DIR / "TBG_Calidad.xlsm",
    )
).resolve()


# ============================================================
# HOJAS DE EXCEL
# ============================================================

HOJA_INICIO = "01_INICIO"
HOJA_ENTRADA = "02_ENTRADA"
HOJA_RESULTADOS = "03_RESULTADOS"
HOJA_INDICADORES = "04_INDICADORES"
HOJA_CONFIGURACION = "05_CONFIGURACION"
HOJA_AUDITORIA = "06_AUDITORIA"
HOJA_HISTORICO = "HISTORICO"

# ============================================================
# ARCHIVOS DE CONFIGURACION
# ============================================================



# ============================================================
# FILAS INICIALES DE LAS TABLAS EN EXCEL
# ============================================================

FILA_INICIO_ENTRADA = 6
FILA_INICIO_CONFIGURACION = 6
FILA_INICIO_RESULTADOS = 6
CELDA_ANIO = "E6"
CELDA_MES = "E7"
CELDA_MODO = "E8"
