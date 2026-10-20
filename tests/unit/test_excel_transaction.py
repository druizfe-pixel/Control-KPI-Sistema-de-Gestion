import unittest
from unittest.mock import MagicMock, patch

from excel_transaction import (
    _escribir_resultados,
    _filas_resultados,
    guardar_y_escribir_resultados,
)


class Celda:
    def __init__(self, valor=None):
        self.Value = valor


class HojaFalsa:
    def __init__(self, codigos):
        self.celdas = {
            f"A{fila}": Celda(codigo)
            for fila, codigo in enumerate(codigos, start=6)
        }

    def Range(self, referencia):
        return self.celdas.setdefault(referencia, Celda())


class TransactionTests(unittest.TestCase):
    def test_mapea_filas_por_codigo(self):
        self.assertEqual(_filas_resultados(HojaFalsa([2, 3, None])), {2: 6, 3: 7})

    def test_valida_todos_los_destinos_antes_de_escribir(self):
        hoja = HojaFalsa([2, None])
        resultados = [
            {"codigo": 2, "cumplimiento": 1, "resultado_ponderado": 0.05, "acumulado": 1},
            {"codigo": 99, "cumplimiento": 1, "resultado_ponderado": 0.05, "acumulado": 1},
        ]
        with self.assertRaises(ValueError):
            _escribir_resultados(
                hoja,
                resultados,
                _filas_resultados(hoja),
                2027,
                3,
                "PROCESAR",
            )
        self.assertIsNone(hoja.Range("E6").Value)


    def test_fallo_durante_historico_no_guarda(self):
        excel = MagicMock()
        libro = MagicMock()

        excel.Workbooks.Open.return_value = libro

        with patch(
            "excel_transaction.win32.DispatchEx",
            return_value=excel,
        ), patch(
            "excel_transaction._validar_accion",
        ), patch(
            "excel_transaction._filas_resultados",
            return_value={2: 6},
        ), patch(
            "excel_transaction._escribir_resultados",
        ), patch(
            "excel_transaction._insertar_historico",
            side_effect=RuntimeError(
                "FALLO_CONTROLADO_21_2_HISTORICO"
            ),
        ):

            with self.assertRaisesRegex(
                RuntimeError,
                "FALLO_CONTROLADO_21_2_HISTORICO",
            ):
                guardar_y_escribir_resultados(
                    registros=[
                        {
                            "anio": 2027,
                            "mes": 3,
                            "codigo": 2,
                            "indicador": "Prueba",
                            "objetivo": 1,
                            "tipo": "P",
                            "meta": 0.90,
                            "real": 0.81,
                            "cumplimiento": 0.90,
                            "peso": 0.05,
                            "resultado_ponderado": 0.045,
                        }
                    ],
                    resultados=[
                        {
                            "codigo": 2,
                            "meta": 0.90,
                            "real": 0.81,
                            "cumplimiento": 0.90,
                            "resultado_ponderado": 0.045,
                            "acumulado": 0.90,
                        }
                    ],
                    anio=2027,
                    mes=3,
                    accion_periodo="PROCESAR",
                    modo="NORMAL",
                )

        excel.Save.assert_not_called()
        libro.Save.assert_not_called()

        libro.Close.assert_called_once_with(
            SaveChanges=False
        )

        excel.Quit.assert_called_once()

    def test_fallo_durante_auditoria_no_guarda(self):
        excel = MagicMock()
        libro = MagicMock()

        excel.Workbooks.Open.return_value = libro

        with patch(
            "excel_transaction.win32.DispatchEx",
            return_value=excel,
        ), patch(
            "excel_transaction._validar_accion",
        ), patch(
            "excel_transaction._filas_resultados",
            return_value={2: 6},
        ), patch(
            "excel_transaction._escribir_resultados",
        ), patch(
            "excel_transaction._insertar_historico",
        ) as mock_insertar_historico, patch(
            "excel_transaction._registrar_auditoria",
            side_effect=RuntimeError(
                "FALLO_CONTROLADO_21_3_AUDITORIA"
            ),
        ):

            with self.assertRaisesRegex(
                RuntimeError,
                "FALLO_CONTROLADO_21_3_AUDITORIA",
            ):
                guardar_y_escribir_resultados(
                    registros=[
                        {
                            "anio": 2027,
                            "mes": 4,
                            "codigo": 2,
                            "indicador": "Prueba",
                            "objetivo": 1,
                            "tipo": "P",
                            "meta": 0.90,
                            "real": 0.81,
                            "cumplimiento": 0.90,
                            "peso": 0.05,
                            "resultado_ponderado": 0.045,
                        }
                    ],
                    resultados=[
                        {
                            "codigo": 2,
                            "meta": 0.90,
                            "real": 0.81,
                            "cumplimiento": 0.90,
                            "resultado_ponderado": 0.045,
                            "acumulado": 0.90,
                        }
                    ],
                    anio=2027,
                    mes=4,
                    accion_periodo="PROCESAR",
                    modo="NORMAL",
                )

        mock_insertar_historico.assert_called_once()

        excel.Save.assert_not_called()
        libro.Save.assert_not_called()

        libro.Close.assert_called_once_with(
            SaveChanges=False
        )

        excel.Quit.assert_called_once()