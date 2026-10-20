from pathlib import Path
import ctypes
import json
import win32com.client as win32

from config import (
    RUTA_EXCEL,
    HOJA_INICIO,
    HOJA_ENTRADA,
    HOJA_INDICADORES,
    HOJA_CONFIGURACION,
    FILA_INICIO_ENTRADA,
    FILA_INICIO_CONFIGURACION,
    CELDA_MODO,
)

from validator import (
    validar_indicadores,
    validar_pesos_objetivos,
    validar_entrada,
    validar_codigos_entrada,
    validar_codigos_duplicados,
    validar_coherencia_entrada_maestro
)

from calculator import (
    calcular_indicadores,
    preparar_calculo_periodo,
    total_resultado_ponderado,
    preparar_registros_historico,
    agregar_acumulado,
)

from historico import (
    leer_historico,
    periodo_tiene_registros,
)

from excel_transaction import guardar_y_escribir_resultados
from entrada import sincronizar_entrada
from progress_ui import ProgressWindow

# ============================================================
# CARGAR MAESTRO DE INDICADORES
# ============================================================

# def cargar_maestro():
#     with open(RUTA_MAESTRO, "r", encoding="utf-8") as archivo:
#         return json.load(archivo)


# ============================================================
# LEER PERÃODO
# ============================================================

def leer_periodo(hoja_inicio):
    """
    Leer el AÃ±o y Mes seleccionados en 01_INICIO.

    Retorna:
        Diccionario con:
            anio
            mes

    El perÃ­odo es definido explÃ­citamente por el usuario
    y no se deriva de la fecha del computador.
    """

    anio = hoja_inicio.Range("E6").Value
    mes = hoja_inicio.Range("E7").Value

    if anio is None or str(anio).strip() == "":
        raise ValueError(
            "El AÃ±o en 01_INICIO!E6 estÃ¡ vacÃ­o."
        )

    if mes is None or str(mes).strip() == "":
        raise ValueError(
            "El Mes en 01_INICIO!E7 estÃ¡ vacÃ­o."
        )

    try:
        anio = int(anio)
    except (TypeError, ValueError):
        raise ValueError(
            "El AÃ±o en 01_INICIO!E6 debe ser numÃ©rico."
        )

    try:
        mes = int(mes)
    except (TypeError, ValueError):
        raise ValueError(
            "El Mes en 01_INICIO!E7 debe ser numÃ©rico."
        )

    if anio < 2026:
        raise ValueError(
            f"El AÃ±o {anio} no es vÃ¡lido."
        )

    if mes < 1 or mes > 12:
        raise ValueError(
            f"El Mes {mes} no es vÃ¡lido. Debe estar entre 1 y 12."
        )

    return {
        "anio": anio,
        "mes": mes,
    }


def leer_modo(hoja_inicio):
    """
    Leer y validar el modo de ejecuciÃ³n seleccionado en Excel.

    Modos permitidos:
        NORMAL
        CORRECCIÃ“N
    """

    valor_modo = hoja_inicio.Range(CELDA_MODO).Value

    if valor_modo is None or str(valor_modo).strip() == "":
        raise ValueError(
            "El modo de ejecuciÃ³n es obligatorio."
        )

    modo = str(valor_modo).strip().upper()

    if modo == "NORMAL":
        return "NORMAL"

    if modo in {"CORRECCION", "CORRECCIÓN", "CORRECCIÃ“N"}:
        # Usar un valor interno sin tilde evita diferencias de codificación.
        return "CORRECCION"

    raise ValueError(
        f"Modo de ejecuciÃ³n invÃ¡lido: {modo}. "
        f"Valores permitidos: NORMAL, CORRECCIÃ“N."
    )


def determinar_accion_periodo(modo, periodo_existe):
    """
    Determinar la acciÃ³n que debe realizar el proceso
    segÃºn el modo seleccionado y la existencia del perÃ­odo.

    Retorna:
        PROCESAR
        BLOQUEAR_EXISTENTE
        CORREGIR
        BLOQUEAR_NO_EXISTENTE
    """

    if modo == "NORMAL":
        # La existencia del período se valida por código más adelante.
        if periodo_existe:
            return "COMPLETAR"
        return "PROCESAR"

    if modo in {"CORRECCION", "CORRECCIÓN", "CORRECCIÃ“N"}:

        if periodo_existe:
            return "CORREGIR"

        return "BLOQUEAR_NO_EXISTENTE"

    raise ValueError(
        f"Modo invÃ¡lido para determinar acciÃ³n: {modo}"
    )



# ============================================================
# ABRIR EXCEL
# ============================================================

def abrir_excel(read_only=False):

    excel = None
    libro = None

    # ------------------------------------------------------------
    # 1. Intentar conectarse a una instancia de Excel ya abierta
    # ------------------------------------------------------------
    try:
        libro = win32.GetObject(str(RUTA_EXCEL))
        excel = libro.Application

        return excel, libro, False

    except Exception:
        excel = None
        libro = None

    # ------------------------------------------------------------
    # 2. Si no estaba abierto, crear una instancia propia
    # ------------------------------------------------------------
    excel = win32.DispatchEx(
        "Excel.Application"
    )

    excel.Visible = False
    excel.DisplayAlerts = False

    libro = excel.Workbooks.Open(
        Filename=str(RUTA_EXCEL),
        ReadOnly=read_only,
    )

    return excel, libro, True


def leer_indicadores_configurados(hoja_indicadores):
    """
    Lee la configuraciÃ³n estructural desde 04_INDICADORES.

    04_INDICADORES es la fuente maestra operativa del sistema.

    No utiliza diccionario_datos_v1.json.
    """

    indicadores = []
    codigos_vistos = set()

    fila = 6

    while True:

        codigo = hoja_indicadores.Range(
            f"A{fila}"
        ).Value

        if codigo is None or str(codigo).strip() == "":
            break

        try:
            codigo = int(codigo)
        except (TypeError, ValueError) as error:
            raise ValueError(
                f"04_INDICADORES!A{fila} contiene "
                f"un cÃ³digo invÃ¡lido: {codigo!r}"
            ) from error

        if codigo <= 0:
            raise ValueError(
                f"El cÃ³digo {codigo} de 04_INDICADORES "
                "debe ser mayor que cero."
            )

        if codigo in codigos_vistos:
            raise ValueError(
                f"El cÃ³digo {codigo} aparece duplicado "
                "en 04_INDICADORES."
            )

        objetivo = hoja_indicadores.Range(
            f"B{fila}"
        ).Value

        indicador = hoja_indicadores.Range(
            f"C{fila}"
        ).Value

        unidad = hoja_indicadores.Range(
            f"D{fila}"
        ).Value

        tipo = hoja_indicadores.Range(
            f"E{fila}"
        ).Value

        direccion = hoja_indicadores.Range(
            f"F{fila}"
        ).Value

        definicion = hoja_indicadores.Range(
            f"G{fila}"
        ).Value

        formula = hoja_indicadores.Range(
            f"J{fila}"
        ).Value

        tope = hoja_indicadores.Range(
            f"K{fila}"
        ).Value

        frecuencia = hoja_indicadores.Range(
            f"L{fila}"
        ).Value

        fuente_actual = hoja_indicadores.Range(
            f"M{fila}"
        ).Value

        acumulado = hoja_indicadores.Range(
            f"N{fila}"
        ).Value

        activo = hoja_indicadores.Range(
            f"O{fila}"
        ).Value

        activo_normalizado = (
            str(activo).strip().upper()
            if activo is not None
            else ""
        )

        if activo_normalizado not in {"SI", "NO"}:
            raise ValueError(
                f"El cÃ³digo {codigo} tiene un valor "
                f"Activo invÃ¡lido: {activo!r}. "
                "Solo se permite SI o NO."
            )

        indicadores.append({
            "codigo": codigo,
            "objetivo": objetivo,
            "indicador": indicador,
            "unidad": unidad,
            "tipo": tipo,
            "direccion": direccion,
            "definicion": definicion,
            "formula_cumplimiento": formula,
            "tope": tope,
            "frecuencia": frecuencia,
            "fuente_actual": fuente_actual,
            "acumulado": acumulado,
            "activo": activo_normalizado,
        })

        codigos_vistos.add(codigo)

        fila += 1

    if not indicadores:
        raise ValueError(
            "04_INDICADORES no contiene indicadores."
        )

    return indicadores


# ============================================================
# LEER ENTRADA
# ============================================================

def leer_entrada(hoja_entrada):

    datos = []
    ultima_fila = (
        hoja_entrada.UsedRange.Row
        + hoja_entrada.UsedRange.Rows.Count
        - 1
    )

    for fila in range(FILA_INICIO_ENTRADA, ultima_fila + 1):

        codigo = hoja_entrada.Range(f"A{fila}").Value
        objetivo = hoja_entrada.Range(f"B{fila}").Value
        indicador = hoja_entrada.Range(f"C{fila}").Value
        unidad = hoja_entrada.Range(f"D{fila}").Value
        meta = hoja_entrada.Range(f"E{fila}").Value
        real = hoja_entrada.Range(f"F{fila}").Value
        estado = hoja_entrada.Range(f"G{fila}").Value

        valores_fila = [
            codigo,
            objetivo,
            indicador,
            unidad,
            meta,
            real,
        ]

        fila_vacia = all(
            valor is None or str(valor).strip() == ""
            for valor in valores_fila
        )

        if fila_vacia:
            continue

        if codigo is None or str(codigo).strip() == "":
            raise ValueError(
                f"La fila {fila} de {HOJA_ENTRADA} tiene datos "
                "pero el cÃ³digo estÃ¡ vacÃ­o."
            )

        datos.append({
            "codigo": codigo,
            "objetivo": objetivo,
            "indicador": indicador,
            "unidad": unidad,
            "meta": meta,
            "real": real,
            "estado": estado,
        })

    return datos


# ============================================================
# LEER CONFIGURACIÃ“N
# ============================================================

def leer_configuracion(hoja_config):

    configuracion = {}

    fila = FILA_INICIO_CONFIGURACION

    while True:

        codigo = hoja_config.Range(f"A{fila}").Value

        if codigo is None or str(codigo).strip() == "":
            break

        tipo = hoja_config.Range(f"B{fila}").Value
        peso = hoja_config.Range(f"C{fila}").Value
        criterio = hoja_config.Range(f"D{fila}").Value

        configuracion[int(codigo)] = {
            "tipo": tipo,
            "peso": peso,
            "criterio": criterio,
        }

        fila += 1

    return configuracion

# ============================================================
# LEER PESOS DE OBJETIVOS
# ============================================================

def leer_pesos_objetivos(hoja_config):

    pesos_objetivos = {}

    for fila in range(6, 11):

        objetivo = hoja_config.Range(f"F{fila}").Value
        peso = hoja_config.Range(f"G{fila}").Value

        if objetivo is None or str(objetivo).strip() == "":
            continue

        if peso is None or str(peso).strip() == "":
            continue

        pesos_objetivos[int(objetivo)] = float(peso)

    return pesos_objetivos


def integrar_indicadores(
    datos_entrada,
    configuracion,
    maestro,
):
    """
    Integra los datos de entrada con la configuraciÃ³n
    y el catÃ¡logo dinÃ¡mico de 04_INDICADORES.
    """

    maestro_por_codigo = {
        int(indicador["codigo"]): indicador
        for indicador in maestro
        if indicador["activo"] == "SI"
    }

    indicadores = []

    for dato in datos_entrada:
        codigo = int(dato["codigo"])

        if codigo not in maestro_por_codigo:
            raise ValueError(
                f"El cÃ³digo {codigo} de 02_ENTRADA "
                "no existe entre los indicadores activos."
            )

        if codigo not in configuracion:
            raise ValueError(
                f"El cÃ³digo {codigo} no tiene configuraciÃ³n "
                "en 05_CONFIGURACION."
            )

        maestro_indicador = maestro_por_codigo[codigo]
        config_indicador = configuracion[codigo]

        indicadores.append({
            **maestro_indicador,
            "meta": dato["meta"],
            "real": dato["real"],
            "estado": dato.get("estado"),
            "tipo": config_indicador["tipo"],
            "peso": config_indicador["peso"],
            "criterio": config_indicador["criterio"],
        })

    return indicadores


# ============================================================
# LEER DATOS COMPLETOS DE EXCEL
# ============================================================

def leer_datos_excel():

    excel = None
    libro = None

    try:

        excel, libro = abrir_excel()

        hoja_entrada = libro.Worksheets(HOJA_ENTRADA)
        hoja_config = libro.Worksheets(HOJA_CONFIGURACION)

        datos_entrada = leer_entrada(hoja_entrada)

        configuracion = leer_configuracion(hoja_config)

        indicadores = integrar_indicadores(
            datos_entrada,
            configuracion,
            maestro,
        )

        return indicadores

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


# ============================================================
# PROCESO PRINCIPAL
# ============================================================


def mostrar_msgbox(mensaje, titulo="TBG CALIDAD", tipo="info", hwnd_excel=0):
    """Muestra el mensaje asociado a la ventana principal de Excel."""
    if not hwnd_excel:
        hwnd_excel = globals().get("HWND_EXCEL_FINAL", 0)

    iconos = {"info": 0x40, "warning": 0x30, "error": 0x10}
    flags = 0x40000 | iconos.get(tipo, 0x40)

    try:
        ctypes.windll.user32.MessageBoxW(
            int(hwnd_excel or 0),
            str(mensaje),
            str(titulo),
            flags,
        )
    except Exception as exc:
        print(f"No fue posible mostrar el mensaje: {exc}")


RESUMEN_FINAL = {}
HWND_EXCEL_FINAL = 0


def ejecutar_proceso():
    global HWND_EXCEL_FINAL
    print("=" * 70)
    print("TBG CALIDAD - MOTOR PYTHON")
    print("=" * 70)

    excel = None
    libro = None
    excel_propio = False
    progreso = None

    try:
        progreso = ProgressWindow()
        progreso.actualizar("Iniciando proceso...", 2)
    except Exception as exc:
        print(f"No se pudo abrir la ventana de progreso: {exc}")
        progreso = None

    try:

        # ========================================================
        # 1. ABRIR Y LEER EXCEL
        # ========================================================

        print("\n1. Leyendo Excel...")

        excel, libro, excel_propio = abrir_excel()

        try:
            excel.StatusBar = (
                "TBG CALIDAD: leyendo configuracion e indicadores..."
            )
        except Exception:
            pass

        if progreso is not None:
            progreso.actualizar("Leyendo configuracion e indicadores...", 10)

        hoja_inicio = libro.Worksheets(HOJA_INICIO)
        hoja_entrada = libro.Worksheets(HOJA_ENTRADA)
        hoja_indicadores = libro.Worksheets(HOJA_INDICADORES)
        hoja_config = libro.Worksheets(HOJA_CONFIGURACION)

        # --------------------------------------------------------
        # 1.1 PerÃ­odo
        # --------------------------------------------------------

        periodo = leer_periodo(hoja_inicio)

        anio = periodo["anio"]
        mes = periodo["mes"]

        print(
            f"   PerÃ­odo seleccionado: "
            f"{anio}-{mes:02d}"
        )

        # --------------------------------------------------------
        # 1.2 Modo
        # --------------------------------------------------------

        modo = leer_modo(hoja_inicio)

        print(
            f"   Modo de ejecuciÃ³n: {modo}"
        )

        # --------------------------------------------------------
        # 1.3 Verificar existencia del perÃ­odo
        # --------------------------------------------------------

        # Leer el histórico una sola vez: se reutiliza para validar el período
        # y calcular los acumulados, evitando una segunda lectura COM completa.
        registros_historicos_previos = leer_historico(libro)
        periodo_existe = any(
            int(registro["anio"]) == anio
            and int(registro["mes"]) == mes
            for registro in registros_historicos_previos
            if registro.get("anio") is not None and registro.get("mes") is not None
        )

        print(
            f"   PerÃ­odo ya procesado: "
            f"{'SÃ' if periodo_existe else 'NO'}"
        )

        accion_periodo = determinar_accion_periodo(
            modo,
            periodo_existe
        )

        print(
            f"   AcciÃ³n del perÃ­odo: {accion_periodo}"
        )

        # --------------------------------------------------------
        # 1.4 Bloqueos por modo/perÃ­odo
        # --------------------------------------------------------

        if accion_periodo == "BLOQUEAR_EXISTENTE":

            print()
            print("[BLOQUEADO] PROCESO BLOQUEADO")
            print(
                f"   El perÃ­odo {anio}-{mes:02d} "
                "ya fue procesado."
            )
            print(
                "   Para reprocesarlo debe utilizar "
                "el modo CORRECCIÃ“N."
            )

            raise RuntimeError(
                f"El perÃ­odo {anio}-{mes:02d} "
                "ya fue procesado. "
                "Use el modo CORRECCIÃ“N para modificarlo."
            )

        if accion_periodo == "BLOQUEAR_NO_EXISTENTE":

            print()
            print("[BLOQUEADO] PROCESO BLOQUEADO")
            print(
                f"   El perÃ­odo {anio}-{mes:02d} "
                "no existe en HISTORICO."
            )
            print(
                "   El modo CORRECCIÃ“N solo puede utilizarse "
                "sobre perÃ­odos existentes."
            )

            raise RuntimeError(
                f"El perÃ­odo {anio}-{mes:02d} "
                "no existe. "
                "El modo CORRECCIÃ“N requiere un perÃ­odo existente."
            )

        # ========================================================
        # 1.5 LEER CATÃLOGO DINÃMICO
        # ========================================================

        print(
            "\n1.1 Leyendo configuraciÃ³n de 04_INDICADORES..."
        )

        maestro = leer_indicadores_configurados(
            hoja_indicadores
        )

        print(
            f"   Indicadores configurados: "
            f"{len(maestro)}"
        )

        # --------------------------------------------------------
        # Indicadores activos
        # --------------------------------------------------------

        activos = [
            indicador
            for indicador in maestro
            if indicador["activo"] == "SI"
        ]

        print(
            f"   Indicadores activos: "
            f"{len(activos)}"
        )

        # Identificar códigos ya guardados para el período seleccionado.
        historico_periodo = {
            int(registro["codigo"]): registro
            for registro in registros_historicos_previos
            if registro.get("anio") is not None and registro.get("mes") is not None
            and registro.get("codigo") is not None
            and int(registro["anio"]) == anio and int(registro["mes"]) == mes
        }
        codigos_activos = {int(item["codigo"]) for item in activos}
        codigos_faltantes = codigos_activos - set(historico_periodo)
        if accion_periodo == "COMPLETAR" and not codigos_faltantes:
            raise RuntimeError(
                f"El período {anio}-{mes:02d} ya tiene histórico para todos los indicadores activos. "
                "Use CORRECCIÓN si necesita modificar datos."
            )
        print(f"   Indicadores sin histórico del período: {len(codigos_faltantes)}")

        # ========================================================
        # 1.6 SINCRONIZAR 02_ENTRADA
        # ========================================================

        print(
            "\n1.2 Sincronizando 02_ENTRADA..."
        )

        resumen_entrada = sincronizar_entrada(
            hoja_entrada,
            activos,
        )

        print(
            f"   Activos     : "
            f"{resumen_entrada['activos']}"
        )

        print(
            f"   Conservados : "
            f"{resumen_entrada['conservados']}"
        )

        print(
            f"   Nuevos      : "
            f"{resumen_entrada['nuevos']}"
        )

        print(
            f"   Eliminados  : "
            f"{resumen_entrada['eliminados']}"
        )

        # --------------------------------------------------------
        # IMPORTANTE:
        # Volver a leer 02_ENTRADA despuÃ©s de sincronizarla.
        # --------------------------------------------------------

        datos_entrada = leer_entrada(
            hoja_entrada
        )

        print(
            f"   Registros leÃ­dos de 02_ENTRADA: "
            f"{len(datos_entrada)}"
        )

        # ========================================================
        # 1.7 VALIDAR ENTRADA
        # ========================================================

        errores_duplicados = validar_codigos_duplicados(
            datos_entrada
        )

        if errores_duplicados:

            print(
                "\n[ERROR] CÃ“DIGOS DUPLICADOS"
            )

            for error in errores_duplicados:
                print(
                    f"[ERROR] {error}"
                )

            print(
                "\n[BLOQUEADO] PROCESO BLOQUEADO"
            )

            raise ValueError(
                f"Se encontraron "
                f"{len(errores_duplicados)} "
                "errores de cÃ³digos duplicados."
            )

        errores_entrada = validar_entrada(
            datos_entrada
        )

        if errores_entrada:

            print(
                "\n[ERROR] VALIDACIÃ“N DE ENTRADA FALLIDA"
            )

            for error in errores_entrada:
                print(
                    error
                )

            raise ValueError(
                f"Se encontraron "
                f"{len(errores_entrada)} "
                "errores en la entrada."
            )

        # ========================================================
        # 1.8 LEER CONFIGURACIÃ“N
        # ========================================================

        configuracion = leer_configuracion(
            hoja_config
        )

        pesos_objetivos = leer_pesos_objetivos(
            hoja_config
        )

        # ========================================================
        # 1.9 VALIDAR EXISTENCIA DE CÃ“DIGOS
        # ========================================================

        print(
            "\n1.5 Validando existencia de cÃ³digos..."
        )

        errores_codigos = validar_codigos_entrada(
            datos_entrada,
            maestro,
            configuracion
        )

        if errores_codigos:

            print(
                "\n[ERROR] CÃ“DIGOS INVÃLIDOS"
            )

            for error in errores_codigos:
                print(
                    f"[ERROR] {error}"
                )

            print(
                "\n[BLOQUEADO] PROCESO BLOQUEADO"
            )

            raise ValueError(
                f"Se encontraron "
                f"{len(errores_codigos)} "
                "errores de cÃ³digos."
            )

        print(
            "   [OK] CÃ³digos OK"
        )

        # ========================================================
        # 1.10 VALIDAR COHERENCIA CON MAESTRO
        # ========================================================

        print(
            "\n1.6 Validando coherencia CÃ³digo â†” Maestro..."
        )

        errores_coherencia = (
            validar_coherencia_entrada_maestro(
                datos_entrada,
                maestro
            )
        )

        if errores_coherencia:

            print(
                "\n[ERROR] INCONSISTENCIAS "
                "CÃ“DIGO â†” MAESTRO"
            )

            for error in errores_coherencia:
                print(
                    f"[ERROR] {error}"
                )

            print(
                "\n[BLOQUEADO] PROCESO BLOQUEADO"
            )

            raise ValueError(
                f"Se encontraron "
                f"{len(errores_coherencia)} "
                "inconsistencias entre "
                "entrada y maestro."
            )

        print(
            "   [OK] Coherencia CÃ³digo â†” Maestro OK"
        )

        # ========================================================
        # 1.11 INTEGRAR DATOS
        # ========================================================

        indicadores = integrar_indicadores(
            datos_entrada,
            configuracion,
            maestro,
        )

        print(
            f"   Indicadores integrados: "
            f"{len(indicadores)}"
        )

        # ========================================================
        # 2. VALIDAR CONFIGURACIÃ“N DE PESOS
        # ========================================================

        print(
            "\n2. Validando configuraciÃ³n de pesos..."
        )

        indicadores_maestro = {
            int(item["codigo"]): {
                "objetivo": int(item["objetivo"])
            }
            for item in maestro
        }

        pesos_indicadores = {
            codigo: datos["peso"]
            for codigo, datos in configuracion.items()
        }

        errores_pesos = validar_pesos_objetivos(
            indicadores_maestro,
            pesos_indicadores,
            pesos_objetivos,
        )

        if errores_pesos:

            print(
                "\n[ERROR] CONFIGURACIÃ“N DE PESOS INVÃLIDA"
            )

            for error in errores_pesos:
                print(
                    f"[ERROR] {error}"
                )

            print(
                "\n[BLOQUEADO] PROCESO BLOQUEADO"
            )

            raise ValueError(
                "La configuraciÃ³n de pesos es invÃ¡lida. "
                "El proceso no puede continuar."
            )

        print(
            "   [OK] ConfiguraciÃ³n de pesos OK"
        )

        # ========================================================
        # 3. VALIDAR INDICADORES
        # ========================================================

        print(
            "\n3. Validando indicadores..."
        )

        objetivos_validos = set(
            pesos_objetivos.keys()
        )

        errores = validar_indicadores(
            indicadores,
            objetivos_validos
        )

        # ========================================================
        # 3.1 PREPARAR ESTADO DEL PERÃODO
        # ========================================================

        preparacion = preparar_calculo_periodo(
            indicadores
        )

        estado_periodo = preparacion[
            "estado_periodo"
        ]

        if accion_periodo == "COMPLETAR":
            preparacion["calculables"] = [
                item for item in preparacion["calculables"]
                if int(item["codigo"]) in codigos_faltantes
            ]
            preparacion["pendientes"] = [
                item for item in preparacion["pendientes"]
                if int(item["codigo"]) in codigos_faltantes
            ]
        elif accion_periodo == "CORREGIR":
            def valor_distinto(actual, anterior):
                if actual is None or anterior is None:
                    return actual != anterior
                try:
                    return abs(float(actual) - float(anterior)) > 1e-9
                except (TypeError, ValueError):
                    return str(actual).strip() != str(anterior).strip()

            campos = ("meta", "real", "peso", "objetivo", "tipo", "indicador")
            codigos_a_corregir = set()
            for item in indicadores:
                codigo = int(item["codigo"])
                anterior = historico_periodo.get(codigo)
                if anterior is None or any(
                    valor_distinto(item.get(campo), anterior.get(campo))
                    for campo in campos
                ):
                    codigos_a_corregir.add(codigo)
            preparacion["calculables"] = [
                item for item in preparacion["calculables"]
                if int(item["codigo"]) in codigos_a_corregir
            ]
            preparacion["pendientes"] = [
                item for item in preparacion["pendientes"]
                if int(item["codigo"]) in codigos_a_corregir
            ]
            print(f"   Indicadores identificados para corrección: {len(codigos_a_corregir)}")

        print(
            f"\n   Estado del perÃ­odo: "
            f"{estado_periodo}"
        )

        print(
            f"   Calculables: "
            f"{len(preparacion['calculables'])}"
        )

        print(
            f"   Pendientes: "
            f"{len(preparacion['pendientes'])}"
        )

        # --------------------------------------------------------
        # Errores de validaciÃ³n
        # --------------------------------------------------------

        if errores:

            print(
                "\n[ERROR] VALIDACIÃ“N FALLIDA"
            )

            for error in errores:
                print(
                    error
                )

            raise ValueError(
                f"Se encontraron "
                f"{len(errores)} "
                "errores de validaciÃ³n."
            )

        print(
            "   [OK] ValidaciÃ³n OK"
        )

        if accion_periodo in {"COMPLETAR", "CORREGIR"} and not preparacion["calculables"]:
            mensaje = (
                "No hay indicadores faltantes con datos válidos para completar."
                if accion_periodo == "COMPLETAR"
                else "No se detectaron cambios en los indicadores respecto al histórico."
            )
            print(f"\n[INFO] {mensaje}")
            RESUMEN_FINAL.clear()
            RESUMEN_FINAL.update({
                "anio": anio, "mes": mes, "estado_periodo": "SIN_CAMBIOS",
                "calculados": 0, "pendientes": len(preparacion["pendientes"]),
                "historicos": 0, "resultados_escritos": 0,
            })
            return []

        # ========================================================
        # 4. ESTADOS SIN DATOS / BLOQUEADO
        # ========================================================

        if estado_periodo == "SIN_DATOS":

            print(
                "\n[SIN_DATOS] No existen datos "
                "vÃ¡lidos para calcular el perÃ­odo."
            )

            return {
                "estado_periodo": "SIN_DATOS",
                "resultados": [],
            }

        if estado_periodo == "BLOQUEADO":

            print(
                "\n[BLOQUEADO] El perÃ­odo tiene "
                "errores y no puede calcularse."
            )

            raise ValueError(
                "El perÃ­odo estÃ¡ BLOQUEADO. "
                "No se puede calcular."
            )

        # ========================================================
        # 5. CALCULAR INDICADORES
        # ========================================================

        print(
            "\n4. Calculando indicadores..."
        )

        # TBG_PROGRESS_CALLBACK
        if progreso is not None:
            progreso.actualizar("Calculando indicadores...", 20)

        try:
            excel.StatusBar = "TBG CALIDAD: calculando indicadores..."
        except Exception:
            pass

        def actualizar_progreso_indicadores(actual, total):
            if progreso is not None:
                progreso.indicador(actual, total)

            try:
                excel.StatusBar = (
                    f"TBG CALIDAD: calculando indicador "
                    f"{actual} de {total}..."
                )
            except Exception:
                pass

        resultados = calcular_indicadores(
            preparacion["calculables"],
            callback=actualizar_progreso_indicadores,
        )

        total = total_resultado_ponderado(
            resultados
        )

        print(
            f"   Indicadores calculados: "
            f"{len(resultados)}"
        )

        print(
            f"   Resultado ponderado total: "
            f"{total:.2%}"
        )

        # ========================================================
        # 5.1 PREPARAR HISTÃ“RICO
        # ========================================================

        registros_historico = (
            preparar_registros_historico(
                resultados,
                indicadores,
                anio,
                mes
            )
        )

        print(
            f"   Registros preparados para "
            f"HISTORICO: "
            f"{len(registros_historico)}"
        )

        # # ========================================================
        # 5.2 MANTENER EXCEL ABIERTO
        # ========================================================
        #
        # El libro permanece abierto para:
        # - calcular acumulados desde el mismo HISTORICO
        # - escribir resultados en el mismo TBG_Calidad.xlsm
        # - evitar abrir una segunda instancia de Excel



        # ========================================================
        # 6. CALCULAR ACUMULADOS
        # ========================================================

        print(
            "\n6. Calculando acumulados..."
        )

        # registros_historicos_previos ya se leyó al validar el período.
        if accion_periodo in {"PROCESAR", "COMPLETAR"}:

            registros_historicos_actualizados = (
                registros_historicos_previos
                + registros_historico
            )

        elif accion_periodo == "CORREGIR":

            # Corrección incremental: reemplazar únicamente los códigos
            # calculados ahora; conservar los otros del mismo período.
            codigos_corregidos = {
                int(registro["codigo"]) for registro in registros_historico
            }
            registros_historicos_actualizados = [
                registro
                for registro in registros_historicos_previos
                if not (
                    int(registro["anio"]) == anio
                    and int(registro["mes"]) == mes
                    and int(registro["codigo"]) in codigos_corregidos
                )
            ] + registros_historico

        else:

            raise RuntimeError(
                f"AcciÃ³n de perÃ­odo no vÃ¡lida: "
                f"{accion_periodo}"
            )

        try:
            excel.StatusBar = "TBG CALIDAD: calculando acumulados..."
        except Exception:
            pass

        if progreso is not None:
            progreso.actualizar("Calculando acumulados...", 80)

        try:
            excel.StatusBar = "TBG CALIDAD: calculando acumulados..."
        except Exception:
            pass

        resultados_con_acumulado = []

        for resultado in resultados:

            resultado_con_acumulado = (
                agregar_acumulado(
                    resultado,
                    registros_historicos_actualizados
                )
            )

            resultados_con_acumulado.append(
                resultado_con_acumulado
            )

        resultados = resultados_con_acumulado

        print(
            "   Acumulados calculados: OK"
        )

        for resultado in resultados:

            print(
                f"   CÃ³digo {resultado['codigo']}: "
                f"Acumulado = "
                f"{resultado['acumulado']:.2%}"
            )

        # ========================================================
        # 7. ESCRITURA TRANSACCIONAL
        # ========================================================

        print(
            "\n7. Guardando resultados..."
        )

        try:
            excel.StatusBar = (
                "TBG CALIDAD: guardando resultados e hist?rico..."
            )
        except Exception:
            pass

        if progreso is not None:
            progreso.actualizar("Guardando resultados e historico...", 90)

        try:
            excel.StatusBar = (
                "TBG CALIDAD: guardando resultados e historico..."
            )
        except Exception:
            pass

        resumen_escritura = (
            guardar_y_escribir_resultados(
                registros_historico,
                resultados,
                anio,
                mes,
                accion_periodo,
                modo,
                indicadores_activos=activos,
                libro=libro,
            )
        )

        print(
            "   Libro actualizado en una Ãºnica operaciÃ³n: "
            f"{resumen_escritura['insertados']} "
            "histÃ³ricos y "
            f"{resumen_escritura['resultados_escritos']} "
            "resultados."
        )

        # ========================================================
        # 8. PROCESO COMPLETADO
        # ========================================================

        print(
            "\n" + "=" * 70
        )

        print(
            "PROCESO COMPLETADO"
        )

        print(
            "=" * 70
        )

        RESUMEN_FINAL.clear()
        RESUMEN_FINAL.update({
            "anio": anio,
            "mes": mes,
            "estado_periodo": estado_periodo,
            "calculados": len(resultados),
            "pendientes": len(preparacion["pendientes"]),
            "historicos": resumen_escritura["insertados"],
            "resultados_escritos": resumen_escritura["resultados_escritos"],
        })

        return resultados

    finally:

        if progreso is not None:
            try:
                progreso.cerrar()
            except Exception:
                pass

        try:
            if excel is not None:
                excel.StatusBar = False
        except Exception:
            pass

        # ========================================================
        # MOSTRAR EXCEL Y DEJAR EL LIBRO DISPONIBLE AL FINALIZAR
        # ========================================================

        if excel is not None:

            try:
                excel.Visible = True
            except Exception:
                pass

        if libro is not None:

            try:
                if libro.Windows.Count > 0:
                    libro.Windows(1).Visible = True
            except Exception:
                pass

            try:
                libro.Activate()
            except Exception:
                pass

            try:
                libro.Worksheets(HOJA_INICIO).Activate()
            except Exception:
                pass

        if excel is not None:
            try:
                HWND_EXCEL_FINAL = int(excel.Hwnd)
            except Exception:
                HWND_EXCEL_FINAL = 0
# ============================================================
# EJECUCION DIRECTA
# ============================================================

# TBG_CALIDAD_LOG_REDIRECT
# Guardar la salida del programa aunque el ejecutable no tenga consola.
import os
import sys
from pathlib import Path

try:
    _log_root = Path(
        os.environ.get("LOCALAPPDATA", str(Path.home()))
    ) / "TBG_Calidad" / "logs"
    _log_root.mkdir(parents=True, exist_ok=True)
    _log_handle = open(
        _log_root / "TBG_Calidad.log",
        "a",
        encoding="utf-8",
        buffering=1,
        errors="replace",
    )
    sys.stdout = _log_handle
    sys.stderr = _log_handle
except Exception:
    # Evita fallos de print() si no se puede crear el registro.
    _log_handle = open(os.devnull, "w", encoding="utf-8")
    sys.stdout = _log_handle
    sys.stderr = _log_handle

if __name__ == "__main__":

    try:
        resultado_proceso = ejecutar_proceso()

        resumen = RESUMEN_FINAL

        if isinstance(resultado_proceso, dict) and resultado_proceso.get("estado_periodo") == "SIN_DATOS":
            mostrar_msgbox(
                "PROCESO SIN DATOS\n\n"
                "No existen datos validos para calcular el periodo.",
                "TBG CALIDAD - SIN DATOS",
                "warning",
            )
        elif resumen:
            pendientes = resumen["pendientes"]
            estado = (
                "COMPLETADO CON PENDIENTES"
                if pendientes > 0
                else "COMPLETADO CORRECTAMENTE"
            )

            mensaje = (
                f"ESTADO: {estado}\n\n"
                f"Periodo: {resumen['mes']:02d}/{resumen['anio']}\n"
                f"Indicadores calculados: {resumen['calculados']}\n"
                f"Indicadores pendientes: {pendientes}\n"
                f"Registros guardados en HISTORICO: {resumen['historicos']}\n"
                f"Resultados escritos en Excel: {resumen['resultados_escritos']}"
            )

            mostrar_msgbox(
                mensaje,
                f"TBG CALIDAD - {estado}",
                "warning" if pendientes > 0 else "info",
            )
        else:
            mostrar_msgbox(
                "El proceso finalizo, pero no se pudo construir el resumen.\n\n"
                "Verifique el libro de Excel.",
                "TBG CALIDAD - RESUMEN NO DISPONIBLE",
                "warning",
            )
    except RuntimeError as exc:
        print("\n" + "=" * 70)
        print("PROCESO NO EJECUTADO")
        print("=" * 70)
        print(f"\n   {exc}")
        print("=" * 70)

        mostrar_msgbox(
            f"El proceso no se ejecuto:\n\n{exc}",
            "TBG CALIDAD - PROCESO NO EJECUTADO",
            "warning",
        )

    except Exception as exc:
        print("\n" + "=" * 70)
        print("ERROR DURANTE EL PROCESO")
        print("=" * 70)
        print(f"\n   {exc}")
        print("\nConsulte la configuracion del archivo Excel.")
        print("=" * 70)

        mostrar_msgbox(
            f"Ocurrio un error durante el proceso:\n\n{exc}\n\n"
            "Consulte la configuracion del archivo Excel.",
            "TBG CALIDAD - ERROR",
            "error",
        )
