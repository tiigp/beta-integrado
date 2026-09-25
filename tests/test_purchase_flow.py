import sys
import types
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
APP_PATH = ROOT / "Control de combustible"
sys.path.insert(0, str(APP_PATH))

import app as compras_app


class PurchaseFlowTests(unittest.TestCase):
    def test_gerencia_threshold_uses_estimated_total_not_zero_amount(self):
        pedido = types.SimpleNamespace(
            observacion="",
            requiere_gerencia=False,
            rango_monto="mayor",
            total_estimado=2_500_000,
            presupuestos=[],
        )
        self.assertTrue(compras_app.evaluar_requisito_gerencia(pedido, 2_500_000))

    def test_special_case_allows_exception_for_urgent_or_single_supplier(self):
        pedido = types.SimpleNamespace(
            observacion="Compra urgente con proveedor único",
            requiere_gerencia=False,
            rango_monto="menor",
            total_estimado=150_000,
            presupuestos=[],
        )
        self.assertTrue(compras_app.solicitud_puede_omitir_presupuestos(pedido))

    def test_horno_laminacion_summary_aggregates_palanquillas_and_tons(self):
        registros = [
            {"colada": "127ABC", "cantidad": 18, "peso": "600"},
            {"colada": "127ABC", "cantidad": 12, "peso": "400"},
            {"colada": "999XYZ", "cantidad": 8, "peso": "240"},
        ]
        resumen = compras_app.resumen_colada_horno(registros)
        self.assertEqual(resumen["colada"], "127ABC")
        self.assertEqual(resumen["palanquillas"], 30)
        self.assertAlmostEqual(resumen["tonelada"], 1.0)

    def test_horno_laminacion_summary_exposes_length_and_carbon_for_active_colada(self):
        etiquetas = [
            types.SimpleNamespace(colada="127ABC", cantidad=10, peso="320", longitud="6", carbono="0.72"),
            types.SimpleNamespace(colada="127ABC", cantidad=8, peso="260", longitud="6", carbono="0.72"),
            types.SimpleNamespace(colada="999XYZ", cantidad=5, peso="200", longitud="8", carbono="0.90"),
        ]

        class FakeQuery:
            def __init__(self, rows):
                self.rows = rows

            def filter_by(self, **filters):
                return FakeQuery([row for row in self.rows if row.colada == filters.get("colada")])

            def all(self):
                return self.rows

        original = compras_app.EtiquetaCalidad
        try:
            compras_app.EtiquetaCalidad = types.SimpleNamespace(query=FakeQuery(etiquetas))
            resumen = compras_app._get_horno_colada_summary("127ABC")
            self.assertEqual(resumen["palanquillas"], 18)
            self.assertEqual(resumen["longitud"], "6")
            self.assertEqual(resumen["grado_carbono"], "0.72")
            self.assertAlmostEqual(resumen["tonelada"], 0.58)
        finally:
            compras_app.EtiquetaCalidad = original


if __name__ == "__main__":
    unittest.main()
