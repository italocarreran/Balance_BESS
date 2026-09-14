import tempfile
import unittest
from pathlib import Path

import pandas as pd

from Script import nucleo


class TestProrrataRetiros(unittest.TestCase):
    def test_lectura_valida_y_asignacion_conserva_totales(self):
        with tempfile.TemporaryDirectory() as tmp:
            ruta = Path(tmp) / "Prorrata_Retiros_2607_def.xlsx"
            fuente = pd.DataFrame({
                "Cuarto de Hora": [1, 1, 2, 2],
                "Suministrador": ["A", "B", "A", "B"],
                "Prorrata": [0.6, 0.4, 0.25, 0.75],
            })
            with pd.ExcelWriter(ruta, engine="openpyxl") as writer:
                fuente.to_excel(writer, sheet_name="Prorrata 15min", index=False)

            cargada = nucleo.leer_prorrata_retiros(ruta)
            ecostos = pd.DataFrame({
                "Bloque Mes Descarga": [1, 2],
                "Monto a compensar": [100.0, 40.0],
            })
            re545 = pd.DataFrame({
                "Bloque horario": [1, 2],
                "Monto a compensar": [20.0, 60.0],
            })
            detalle, por_cuarto, pagos = nucleo.construir_prorrata_retiros(
                cargada, ecostos, re545
            )

            self.assertAlmostEqual(detalle["Pago"].sum(), 220.0)
            self.assertEqual(por_cuarto["Total Compensacion [$]"].tolist(), [120.0, 100.0])
            por_empresa = dict(zip(pagos["Suministrador"], pagos["Pago"]))
            self.assertAlmostEqual(por_empresa["A"], 97.0)
            self.assertAlmostEqual(por_empresa["B"], 123.0)

    def test_rechaza_suma_distinta_de_uno(self):
        with tempfile.TemporaryDirectory() as tmp:
            ruta = Path(tmp) / "Prorrata_Retiros_2607_pre.xlsx"
            pd.DataFrame({
                "Cuarto de Hora": [1, 1],
                "Suministrador": ["A", "B"],
                "Prorrata": [0.4, 0.4],
            }).to_excel(ruta, sheet_name="Prorrata 15min", index=False)
            with self.assertRaisesRegex(nucleo.ErrorEntrada, "RET-001"):
                nucleo.leer_prorrata_retiros(ruta)

    def test_resumen_une_quien_recibe_y_quien_paga(self):
        compensacion = pd.DataFrame({
            "Empresa": ["A", "C"],
            "Compensación Total [$]": [150.0, 20.0],
        })
        pagos = pd.DataFrame({
            "Suministrador": ["A", "B"], "Pago": [100.0, 70.0]
        })
        resumen = nucleo.construir_resumen(compensacion, pagos).set_index("NOMBRE")
        self.assertEqual(set(resumen.index), {"A", "B", "C"})
        self.assertAlmostEqual(resumen.loc["A", "NETO"], 50.0)
        self.assertAlmostEqual(resumen.loc["B", "NETO"], -70.0)
        self.assertAlmostEqual(resumen.loc["C", "NETO"], 20.0)

    def test_busqueda_exige_un_archivo_del_periodo(self):
        with tempfile.TemporaryDirectory() as tmp:
            carpeta = Path(tmp)
            (carpeta / "Prorrata_Retiros_2606_def.xlsx").touch()
            esperado = carpeta / "Prorrata_Retiros_2607_pre.xlsx"
            esperado.touch()
            self.assertEqual(
                nucleo.buscar_archivo_prorrata(carpeta, "2607"), esperado
            )


if __name__ == "__main__":
    unittest.main()
