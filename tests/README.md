# Suite de pruebas

`tests/run_tests.py` separa los controles por nivel de riesgo.

## Unitarias

No abren Excel ni modifican archivos. Ejecútelas desde la raíz de `TGB_Calidad_v2`:

```powershell
.\.venv\Scripts\python.exe .\tests\run_tests.py
```

## Integración de solo lectura

Requieren Microsoft Excel y una copia `.xlsm`. Nunca use el libro operativo:

```powershell
$env:TBG_RUTA_EXCEL = (Resolve-Path .\TBG_Calidad_PRUEBA_52_25_12G.xlsm).Path
$env:RUN_EXCEL_TESTS = "1"
.\.venv\Scripts\python.exe .\tests\run_tests.py --integration
```

## Pruebas con escritura

No forman parte de la ejecución automática. Los scripts históricos dentro de
`engine/test_*.py` deben migrarse gradualmente: cada uno debe crear su propia
copia temporal del XLSM, ejecutar el escenario y eliminar solo esa copia.
