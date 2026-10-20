import unittest

from calculator import (
    calcular_indicador,
    total_resultado_ponderado,
    calcular_acumulado,
)


class CalculatorTests(unittest.TestCase):
    def test_indicador_positivo_aplica_tope(self):
        resultado = calcular_indicador({
            "codigo": 2, "indicador": "Prueba", "meta": 0.9,
            "real": 1.0, "tipo": "P", "peso": 0.05,
        })
        self.assertAlmostEqual(resultado["cumplimiento_bruto"], 1 / 0.9)
        self.assertEqual(resultado["cumplimiento"], 1.0)
        self.assertEqual(resultado["resultado_ponderado"], 0.05)

    def test_indicador_negativo_y_total(self):
        resultado = calcular_indicador({
            "codigo": 5, "indicador": "Prueba", "meta": 26,
            "real": 13, "tipo": "N", "peso": 0.03,
        })
        self.assertEqual(resultado["cumplimiento"], 1.0)
        self.assertEqual(total_resultado_ponderado([resultado]), 0.03)

    def test_indicador_positivo_por_debajo_del_100(self):
        resultado = calcular_indicador({
            "codigo": 2,
            "indicador": "Prueba",
            "meta": 0.90,
            "real": 0.81,
            "tipo": "P",
            "peso": 0.05,
        })

        self.assertAlmostEqual(
            resultado["cumplimiento_bruto"],
            0.90,
        )

        self.assertAlmostEqual(
            resultado["cumplimiento"],
            0.90,
        )

        self.assertAlmostEqual(
            resultado["resultado_ponderado"],
            0.045,
        )

    def test_indicador_negativo_por_debajo_del_100(self):
        resultado = calcular_indicador({
            "codigo": 5,
            "indicador": "Prueba",
            "meta": 10,
            "real": 12,
            "tipo": "N",
            "peso": 0.03,
        })

        self.assertAlmostEqual(
            resultado["cumplimiento_bruto"],
            10 / 12,
        )

        self.assertAlmostEqual(
            resultado["cumplimiento"],
            10 / 12,
        )

        self.assertAlmostEqual(
            resultado["resultado_ponderado"],
            (10 / 12) * 0.03,
        )


    def test_indicador_rechaza_tipo_invalido(self):
        with self.assertRaises(ValueError):
            calcular_indicador({
                "codigo": 2,
                "indicador": "Prueba",
                "meta": 0.90,
                "real": 0.81,
                "tipo": "X",
                "peso": 0.05,
            })

    def test_calcular_acumulado_promedia_meses_validos(self):
        acumulado = calcular_acumulado([
            0.90,
            0.80,
            1.00,
        ])

        self.assertAlmostEqual(
            acumulado,
            0.90,
        )

    def test_calcular_acumulado_excluye_valores_invalidos(self):
        acumulado = calcular_acumulado([
            0.90,
            None,
            1.00,
            "ABC",
        ])

        self.assertAlmostEqual(
            acumulado,
            0.95,
        )