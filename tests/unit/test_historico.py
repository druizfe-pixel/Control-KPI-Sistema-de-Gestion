import unittest
from unittest.mock import patch

from historico import (
    obtener_cumplimientos_historicos,
    periodo_tiene_registros,
)


class HistoricoTests(unittest.TestCase):

    @patch("historico.leer_historico")
    def test_obtener_cumplimientos_historicos_filtra_y_ordena(
        self,
        mock_leer_historico,
    ):
        mock_leer_historico.return_value = [
            {
                "anio": 2026,
                "mes": 12,
                "codigo": 2,
                "cumplimiento": 0.90,
                "estado": "OK",
            },
            {
                "anio": 2026,
                "mes": 10,
                "codigo": 2,
                "cumplimiento": 0.95,
                "estado": "OK",
            },
            {
                "anio": 2026,
                "mes": 11,
                "codigo": 2,
                "cumplimiento": 0.90,
                "estado": "OK",
            },
            {
                "anio": 2026,
                "mes": 9,
                "codigo": 3,
                "cumplimiento": 1.00,
                "estado": "OK",
            },
            {
                "anio": 2026,
                "mes": 8,
                "codigo": 2,
                "cumplimiento": 0.80,
                "estado": "CORREGIDO",
            },
            {
                "anio": 2026,
                "mes": 7,
                "codigo": 2,
                "cumplimiento": "ABC",
                "estado": "OK",
            },
            {
                "anio": 2026,
                "mes": 6,
                "codigo": 2,
                "cumplimiento": None,
                "estado": "OK",
            },
        ]

        resultado = obtener_cumplimientos_historicos(2)

        self.assertEqual(
            resultado,
            [
                {
                    "anio": 2026,
                    "mes": 10,
                    "cumplimiento": 0.95,
                },
                {
                    "anio": 2026,
                    "mes": 11,
                    "cumplimiento": 0.90,
                },
                {
                    "anio": 2026,
                    "mes": 12,
                    "cumplimiento": 0.90,
                },
            ],
        )

        mock_leer_historico.assert_called_once()

    @patch("historico.leer_historico")
    def test_periodo_tiene_registros(
        self,
        mock_leer_historico,
    ):
        mock_leer_historico.return_value = [
            {
                "anio": 2026,
                "mes": 10,
                "codigo": 2,
            },
            {
                "anio": 2026,
                "mes": 10,
                "codigo": 3,
            },
            {
                "anio": 2026,
                "mes": 11,
                "codigo": 2,
            },
        ]

        self.assertTrue(
            periodo_tiene_registros(2026, 10)
        )

        self.assertTrue(
            periodo_tiene_registros(2026, 11)
        )

        self.assertFalse(
            periodo_tiene_registros(2026, 12)
        )

        self.assertFalse(
            periodo_tiene_registros(2025, 10)
        )


if __name__ == "__main__":
    unittest.main()