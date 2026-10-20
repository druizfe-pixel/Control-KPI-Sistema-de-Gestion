# TBG Calidad v2

Motor Python para validar, calcular y registrar indicadores de un sistema de gestión de calidad mediante el libro de Excel `TBG_Calidad.xlsm`.

## Requisitos

- Windows.
- Microsoft Excel de escritorio.
- Python compatible con las dependencias del proyecto.
- Dependencias instaladas desde `requirements.txt`.

Instalación de dependencias:

```powershell
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

## Componentes

- `TBG_Calidad.xlsm`: interfaz y libro de trabajo de Excel.
- `engine/`: validación, cálculos, transacciones, escritura e histórico.
- `tests/`: pruebas unitarias automatizadas.

## Flujo del motor

1. Lee el período y el modo desde `01_INICIO`.
2. Lee los indicadores de `02_ENTRADA` y su configuración.
3. Valida los datos y su coherencia con el maestro.
4. Calcula el cumplimiento y las ponderaciones.
5. Actualiza `HISTORICO` y `03_RESULTADOS` en una operación de guardado.

Los modos de operación son `NORMAL` para un período nuevo y `CORRECCIÓN` para reemplazar un período registrado.

## Ejecución

Para ejecutar el motor, desde la raíz del proyecto:

```powershell
Set-Location .\engine
..\.venv\Scripts\python.exe .\main.py
```

Por seguridad, realiza las pruebas funcionales con una copia del libro, no con un archivo operativo.

Ejemplo desde la raíz del proyecto:

```powershell
Copy-Item .\TBG_Calidad.xlsm .\TBG_Calidad_prueba_local.xlsm
$env:TBG_RUTA_EXCEL = (Resolve-Path .\TBG_Calidad_prueba_local.xlsm).Path
Set-Location .\engine
..\.venv\Scripts\python.exe .\main.py
```

La variable `TBG_RUTA_EXCEL` permite indicar la ruta del libro de prueba para la sesión actual de PowerShell.

## Pruebas unitarias

Desde la raíz del proyecto:

```powershell
.\.venv\Scripts\python.exe .\tests\run_tests.py
```

Las pruebas unitarias verifican los cálculos, las validaciones, el histórico y el comportamiento transaccional.

## Construcción del ejecutable

El archivo `TBG_Calidad.spec` contiene la configuración de PyInstaller para construir la aplicación.

## Alcance

Este repositorio contiene el código fuente y los archivos necesarios para el proyecto. Antes de distribuirlo, verifica que el libro de Excel, sus macros y sus metadatos no incluyan información confidencial.
