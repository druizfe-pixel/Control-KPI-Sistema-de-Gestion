"""Ejecutor de la suite segura del motor TBG Calidad.

Uso:
    python tests/run_tests.py
    python tests/run_tests.py --integration
"""

from __future__ import annotations

import argparse
import sys
import unittest
from pathlib import Path


PROJECT_DIR = Path(__file__).resolve().parent.parent
ENGINE_DIR = PROJECT_DIR / "engine"
sys.path.insert(0, str(ENGINE_DIR))


def descubrir(directorio: Path) -> unittest.TestSuite:
    return unittest.defaultTestLoader.discover(
        start_dir=str(directorio),
        pattern="test_*.py",
        top_level_dir=str(PROJECT_DIR),
    )


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--integration",
        action="store_true",
        help="Incluye comprobaciones de solo lectura contra un XLSM configurado.",
    )
    args = parser.parse_args()

    suite = unittest.TestSuite()
    suite.addTests(descubrir(PROJECT_DIR / "tests" / "unit"))
    if args.integration:
        suite.addTests(descubrir(PROJECT_DIR / "tests" / "integration"))

    resultado = unittest.TextTestRunner(verbosity=2).run(suite)
    return 0 if resultado.wasSuccessful() else 1


if __name__ == "__main__":
    raise SystemExit(main())
