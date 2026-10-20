import math
import unittest

from engine.validator import (
    validar_entrada,
    validar_indicadores,
    validar_pesos_objetivos,
    validar_coherencia_entrada_maestro,
    validar_codigos_entrada,
    validar_codigos_duplicados,
    clasificar_estado_entrada,
)


class ValidarEntradaTests(unittest.TestCase):
    def test_acepta_codigo_entero(self):
        datos = [{"codigo": 2.0}]

        self.assertEqual(validar_entrada(datos), [])
        self.assertEqual(datos[0]["codigo"], 2)

    def test_rechaza_codigo_decimal_y_no_finito(self):
        for codigo in (2.5, math.nan, math.inf, True):
            with self.subTest(codigo=codigo):
                self.assertTrue(validar_entrada([{"codigo": codigo}]))


class ValidarIndicadoresTests(unittest.TestCase):

    def indicador_base(self):
        return {
            "codigo": 2,
            "indicador": "Indicador de prueba",
            "unidad": "%",
            "objetivo": 1,
            "meta": 0.90,
            "real": 0.81,
            "tipo": "P",
            "peso": 0.05,
            "criterio": "Mayor es mejor",
        }

    def test_rechaza_valores_no_finitos(self):
        indicador = self.indicador_base()

        indicador["meta"] = math.nan
        indicador["real"] = math.inf
        indicador["peso"] = math.nan

        codigos_error = {
            error["codigo"]
            for error in validar_indicadores([indicador])
        }

        self.assertIn("META_INVALIDA", codigos_error)
        self.assertIn("REAL_NO_NUMERICO", codigos_error)
        self.assertIn("PESO_INVALIDO", codigos_error)

    def test_meta_vacia_con_real_presente_se_clasifica_como_dato_faltante(self):
        estado = clasificar_estado_entrada(None, 0.81, "P")

        self.assertEqual(estado["estado"], "DATO_FALTANTE")
        self.assertTrue(
            any(error["codigo"] == "DATO_FALTANTE"
                and error.get("campo") == "meta"
                for error in estado["errores"])
        )

    def test_real_vacio_con_meta_presente_se_clasifica_como_dato_faltante(self):
        estado = clasificar_estado_entrada(0.90, None, "P")

        self.assertEqual(estado["estado"], "DATO_FALTANTE")
        self.assertTrue(
            any(error["codigo"] == "DATO_FALTANTE"
                and error.get("campo") == "real"
                for error in estado["errores"])
        )

    def test_rechaza_real_no_numerico(self):
        indicador = self.indicador_base()
        indicador["real"] = "ABC"

        errores = validar_indicadores([indicador])

        self.assertTrue(
            any(
                error["codigo"] == "REAL_NO_NUMERICO"
                and error.get("campo") == "real"
                for error in errores
            )
        )


    def test_rechaza_real_negativo(self):
        indicador = self.indicador_base()
        indicador["real"] = -0.81

        errores = validar_indicadores([indicador])

        self.assertTrue(
            any(
                error["codigo"] == "REAL_NEGATIVO"
                and error.get("campo") == "real"
                for error in errores
            )
        )

    def test_rechaza_meta_no_numerica(self):
        indicador = self.indicador_base()
        indicador["meta"] = "ABC"

        errores = validar_indicadores([indicador])

        self.assertTrue(
            any(
                error["codigo"] == "META_NO_NUMERICA"
                and error.get("campo") == "meta"
                for error in errores
            )
        )

    def test_rechaza_meta_negativa(self):
        indicador = self.indicador_base()
        indicador["meta"] = -0.90

        errores = validar_indicadores([indicador])

        self.assertTrue(
            any(
                error["codigo"] == "META_INVALIDA"
                and error.get("campo") == "meta"
                for error in errores
            )
        )

    def test_real_cero_en_tipo_n_se_marca_para_revision_sin_bloquear(self):
        estado = clasificar_estado_entrada(0.90, 0, "N")

        self.assertEqual(estado["estado"], "REVISION REAL - DIVISOR CERO")
        self.assertTrue(
            any(error["codigo"] == "DIVISOR_CERO_N"
                and error.get("campo") == "real"
                for error in estado["errores"])
        )

    def test_rechaza_tipo_invalido(self):
        indicador = self.indicador_base()
        indicador["tipo"] = "X"

        errores = validar_indicadores([indicador])

        self.assertTrue(
            any(
                error["codigo"] == "TIPO_INVALIDO"
                and error.get("campo") == "tipo"
                for error in errores
            )
        )

    def test_rechaza_peso_faltante(self):
        indicadores = {
            2: {"objetivo": 1}
        }

        pesos_indicadores = {}

        pesos_objetivos = {
            1: 0.05
        }

        errores = validar_pesos_objetivos(
            indicadores,
            pesos_indicadores,
            pesos_objetivos,
        )

        self.assertTrue(
            any(
                error["error"] == "PESO_FALTANTE"
                and error["codigo"] == 2
                for error in errores
            )
        )

    def test_rechaza_peso_no_numerico(self):
        indicadores = {
            2: {"objetivo": 1}
        }

        pesos_indicadores = {
            2: "ABC"
        }

        pesos_objetivos = {
            1: 0.05
        }

        errores = validar_pesos_objetivos(
            indicadores,
            pesos_indicadores,
            pesos_objetivos,
        )

        self.assertTrue(
            any(
                error["error"] == "PESO_NO_NUMERICO"
                and error["codigo"] == 2
                for error in errores
            )
        )

    def test_rechaza_peso_invalido(self):
        indicador = self.indicador_base()
        indicador["peso"] = -0.05

        errores = validar_indicadores([indicador])

        self.assertTrue(
            any(
                error["codigo"] == "PESO_INVALIDO"
                and error.get("campo") == "peso"
                for error in errores
            )
        )

    def test_rechaza_tipo_criterio_inconsistente(self):
        indicador = self.indicador_base()
        indicador["tipo"] = "P"
        indicador["criterio"] = "Menor es mejor"

        errores = validar_indicadores([indicador])

        self.assertTrue(
            any(
                error["codigo"] == "TIPO_CRITERIO_INCONSISTENTE"
                for error in errores
            )
        )
    def test_rechaza_objetivo_no_coincide_maestro(self):
        entrada = [
            {
                "codigo": 2,
                "objetivo": 2,
                "indicador": "Indicador de prueba",
                "unidad": "%",
            }
        ]

        maestro = [
            {
                "codigo": 2,
                "objetivo": 1,
                "indicador": "Indicador de prueba",
                "unidad": "%",
            }
        ]

        errores = validar_coherencia_entrada_maestro(entrada, maestro)

        self.assertTrue(
            any(
                error["codigo"] == "OBJETIVO_NO_COINCIDE_MAESTRO"
                for error in errores
            )
        )

    def test_rechaza_indicador_no_coincide_maestro(self):
        entrada = [
            {
                "codigo": 2,
                "objetivo": 1,
                "indicador": "Otro indicador",
                "unidad": "%",
            }
        ]

        maestro = [
            {
                "codigo": 2,
                "objetivo": 1,
                "indicador": "Indicador de prueba",
                "unidad": "%",
            }
        ]

        errores = validar_coherencia_entrada_maestro(entrada, maestro)

        self.assertTrue(
            any(
                error["codigo"] == "INDICADOR_NO_COINCIDE_MAESTRO"
                for error in errores
            )
        )

    def test_rechaza_unidad_no_coincide_maestro(self):
        entrada = [
            {
                "codigo": 2,
                "objetivo": 1,
                "indicador": "Indicador de prueba",
                "unidad": "Unidades",
            }
        ]

        maestro = [
            {
                "codigo": 2,
                "objetivo": 1,
                "indicador": "Indicador de prueba",
                "unidad": "%",
            }
        ]

        errores = validar_coherencia_entrada_maestro(entrada, maestro)

        self.assertTrue(
            any(
                error["codigo"] == "UNIDAD_NO_COINCIDE_MAESTRO"
                for error in errores
            )
        )

    def test_acepta_codigo_existente_en_maestro_y_configuracion(self):
        entrada = [
            {
                "codigo": 2,
                "objetivo": 1,
                "indicador": "Indicador de prueba",
                "unidad": "%",
            }
        ]

        maestro = [
            {
                "codigo": 2,
                "objetivo": 1,
                "indicador": "Indicador de prueba",
                "unidad": "%",
            }
        ]

        configuracion = {
            2: {
                "tipo": "P",
                "peso": 0.05,
                "criterio": "Mayor es mejor",
            }
        }

        errores = validar_codigos_entrada(
            entrada,
            maestro,
            configuracion,
        )

        self.assertEqual(errores, [])

    def test_rechaza_codigo_no_existe_maestro(self):
        entrada = [
            {
                "codigo": 999,
                "objetivo": 1,
                "indicador": "Indicador de prueba",
                "unidad": "%",
            }
        ]

        maestro = [
            {
                "codigo": 2,
                "objetivo": 1,
                "indicador": "Indicador de prueba",
                "unidad": "%",
            }
        ]

        configuracion = {
            999: {
                "tipo": "P",
                "peso": 0.05,
                "criterio": "Mayor es mejor",
            }
        }

        errores = validar_codigos_entrada(
            entrada,
            maestro,
            configuracion,
        )

        self.assertTrue(
            any(
                error["codigo"] == "CODIGO_NO_EXISTE_MAESTRO"
                for error in errores
            )
        )


    def test_rechaza_codigo_no_existe_configuracion(self):
        entrada = [
            {
                "codigo": 999,
                "objetivo": 1,
                "indicador": "Indicador de prueba",
                "unidad": "%",
            }
        ]

        maestro = [
            {
                "codigo": 999,
                "objetivo": 1,
                "indicador": "Indicador de prueba",
                "unidad": "%",
            }
        ]

        configuracion = {
            2: {
                "tipo": "P",
                "peso": 0.05,
                "criterio": "Mayor es mejor",
            }
        }

        errores = validar_codigos_entrada(
            entrada,
            maestro,
            configuracion,
        )

        self.assertTrue(
            any(
                error["codigo"] == "CODIGO_NO_EXISTE_CONFIGURACION"
                for error in errores
            )
        )

    def test_acepta_codigos_sin_duplicados(self):
        entrada = [
            {"codigo": 2},
            {"codigo": 3},
            {"codigo": 4},
        ]

        errores = validar_codigos_duplicados(entrada)

        self.assertEqual(errores, [])

    def test_rechaza_codigo_duplicado(self):
        entrada = [
            {"codigo": 2},
            {"codigo": 3},
            {"codigo": 2},
        ]

        errores = validar_codigos_duplicados(entrada)

        self.assertTrue(
            any(
                error["codigo"] == "CODIGO_DUPLICADO"
                and error.get("valor") == 2
                for error in errores
            )
        )

    def test_acepta_pesos_objetivos_consistentes(self):
        indicadores = {
            2: {"objetivo": 1},
            3: {"objetivo": 1},
            4: {"objetivo": 2},
            5: {"objetivo": 2},
        }

        pesos_indicadores = {
            2: 0.05,
            3: 0.05,
            4: 0.45,
            5: 0.45,
        }

        pesos_objetivos = {
            1: 0.10,
            2: 0.90,
        }

        errores = validar_pesos_objetivos(
            indicadores,
            pesos_indicadores,
            pesos_objetivos,
        )

        self.assertEqual(errores, [])


    def test_rechaza_peso_objetivo_inconsistente(self):
        indicadores = {
            2: {"objetivo": 1},
            3: {"objetivo": 1},
        }

        pesos_indicadores = {
            2: 0.05,
            3: 0.03,
        }

        pesos_objetivos = {
            1: 0.10,
        }

        errores = validar_pesos_objetivos(
            indicadores,
            pesos_indicadores,
            pesos_objetivos,
        )

        self.assertTrue(
            any(
                error["error"] == "PESO_OBJETIVO_INCONSISTENTE"
                and error["objetivo"] == 1
                for error in errores
            )
        )

    def test_rechaza_total_pesos_objetivos_distinto_100(self):
        indicadores = {
            2: {"objetivo": 1},
        }

        pesos_indicadores = {
            2: 0.05,
        }

        pesos_objetivos = {
            1: 0.05,
        }

        errores = validar_pesos_objetivos(
            indicadores,
            pesos_indicadores,
            pesos_objetivos,
        )

        self.assertTrue(
            any(
                error["error"] == "PESO_OBJETIVO_TOTAL_INVALIDO"
                for error in errores
            )
        )


if __name__ == "__main__":
    unittest.main()