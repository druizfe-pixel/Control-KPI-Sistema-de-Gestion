"""Sincronización dinámica de la hoja 02_ENTRADA.

22.13.9.2.4
- Fuente estructural: 04_INDICADORES, Activo=SI.
- Identificación por código.
- Conserva Meta/Real existentes por código.
- Nuevos indicadores quedan vacíos.
- Indicadores inactivos salen de 02_ENTRADA.
- No modifica HISTORICO, RESULTADOS ni CONFIGURACION.
"""

from __future__ import annotations

from typing import Any, Iterable


COLUMNAS_ENTRADA = 7
FILA_INICIO_ENTRADA = 6


def _normalizar_activo(valor: Any) -> str:
    """Normaliza SI/NO para comparación."""
    if valor is None:
        return ""
    return str(valor).strip().upper().replace("Í", "I")


def _es_vacio(valor: Any) -> bool:
    return valor is None or str(valor).strip() == ""


def _clave_codigo(valor: Any) -> int:
    """Convierte un código a entero y rechaza códigos inválidos."""
    if _es_vacio(valor):
        raise ValueError("El código del indicador está vacío.")
    try:
        codigo = int(valor)
    except (TypeError, ValueError) as exc:
        raise ValueError(f"Código de indicador inválido: {valor!r}") from exc
    if codigo <= 0:
        raise ValueError(f"Código de indicador inválido: {codigo}")
    return codigo


def _leer_valor_fila(fila: Any, columna: int) -> Any:
    """Permite trabajar con dicts de prueba o con secuencias."""
    if isinstance(fila, dict):
        nombres = {
            1: "codigo",
            2: "objetivo",
            3: "indicador",
            4: "unidad",
            5: "meta",
            6: "real",
            7: "estado",
        }
        return fila.get(nombres[columna])
    return fila[columna - 1]


def extraer_datos_existentes(filas: Iterable[Any]) -> dict[int, dict[str, Any]]:
    """Extrae Meta/Real actuales de 02_ENTRADA, indexados por código."""
    existentes: dict[int, dict[str, Any]] = {}

    for fila in filas:
        codigo_valor = _leer_valor_fila(fila, 1)

        if _es_vacio(codigo_valor):
            continue

        codigo = _clave_codigo(codigo_valor)

        if codigo in existentes:
            raise ValueError(
                f"El código {codigo} aparece más de una vez en 02_ENTRADA."
            )

        existentes[codigo] = {
            "meta": _leer_valor_fila(fila, 5),
            "real": _leer_valor_fila(fila, 6),
        }

    return existentes


def construir_bloque_entrada(
    indicadores: Iterable[dict[str, Any]],
    existentes: dict[int, dict[str, Any]],
) -> tuple[list[dict[str, Any]], dict[str, int]]:
    """Construye el bloque activo de 02_ENTRADA sin tocar Excel."""
    activos: dict[int, dict[str, Any]] = {}

    for indicador in indicadores:
        codigo = _clave_codigo(indicador.get("codigo"))

        if _normalizar_activo(indicador.get("activo")) != "SI":
            continue

        if codigo in activos:
            raise ValueError(
                f"El código {codigo} aparece más de una vez en 04_INDICADORES."
            )

        activos[codigo] = {
            "codigo": codigo,
            "objetivo": indicador.get("objetivo"),
            "indicador": indicador.get("indicador"),
            "unidad": indicador.get("unidad"),
            "tipo": indicador.get("tipo"),
        }

    bloque: list[dict[str, Any]] = []
    nuevos = 0
    conservados = 0

    for codigo in sorted(activos):
        estructura = activos[codigo]
        previo = existentes.get(codigo)

        if previo is None:
            meta = None
            real = None
            nuevos += 1
        else:
            meta = previo.get("meta")
            real = previo.get("real")
            conservados += 1

        tipo = str(estructura.get("tipo") or "").strip().upper()

        if _es_vacio(meta) and _es_vacio(real):
            estado = "PENDIENTE"
        elif _es_vacio(meta) or _es_vacio(real):
            estado = "DATO_FALTANTE"
        elif tipo == "N":
            try:
                real_num = float(real)
            except (TypeError, ValueError):
                real_num = None

            if real_num == 0:
                estado = "REVISION REAL - DIVISOR CERO"
            else:
                estado = "OK"
        else:
            estado = "OK"

        bloque.append(
            {
                **estructura,
                "meta": meta,
                "real": real,
                "estado": estado,
            }
        )

    codigos_activos = set(activos)
    codigos_existentes = set(existentes)
    eliminados = len(codigos_existentes - codigos_activos)

    resumen = {
        "activos": len(bloque),
        "nuevos": nuevos,
        "eliminados": eliminados,
        "conservados": conservados,
    }

    return bloque, resumen


def sincronizar_entrada(
    hoja_entrada: Any,
    indicadores: Iterable[dict[str, Any]],
    fila_inicio: int = FILA_INICIO_ENTRADA,
) -> dict[str, int]:
    """Sincroniza 02_ENTRADA en una hoja COM de Excel.

    La función:
    1. Lee A:G existentes.
    2. Construye el bloque activo por código.
    3. Limpia A:G desde la fila inicial hasta el último bloque anterior.
    4. Escribe el nuevo bloque.
    5. Devuelve un resumen.

    No guarda el libro; la decisión de Save/rollback pertenece al
    nivel transaccional que invoque esta función.
    """
    # Determinar el último registro actual mediante la columna A.
    ultima_fila = hoja_entrada.Cells(
        hoja_entrada.Rows.Count, 1
    ).End(-4162).Row  # xlUp

    filas_existentes = []
    if ultima_fila >= fila_inicio:
        for fila in range(fila_inicio, ultima_fila + 1):
            valores = [
                hoja_entrada.Cells(fila, columna).Value
                for columna in range(1, COLUMNAS_ENTRADA + 1)
            ]
            if all(_es_vacio(valor) for valor in valores):
                continue
            filas_existentes.append(valores)

    existentes = extraer_datos_existentes(filas_existentes)
    bloque, resumen = construir_bloque_entrada(indicadores, existentes)

    # Limpiar exclusivamente el bloque operativo A:G.
    if ultima_fila >= fila_inicio:
        hoja_entrada.Range(
            f"A{fila_inicio}:G{max(ultima_fila, fila_inicio)}"
        ).ClearContents

    # Escribir el bloque nuevo.
    for desplazamiento, registro in enumerate(bloque):
        fila = fila_inicio + desplazamiento
        valores = (
            registro["codigo"],
            registro["objetivo"],
            registro["indicador"],
            registro["unidad"],
            registro["meta"],
            registro["real"],
            registro["estado"],
        )
        for columna, valor in enumerate(valores, start=1):
            hoja_entrada.Cells(fila, columna).Value = valor

    return resumen
