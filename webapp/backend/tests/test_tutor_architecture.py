"""Architecture guards for the Tutor Scheduling module's solver layer and shared helpers."""

import ast
import unittest
from pathlib import Path

APP = Path(__file__).resolve().parents[1] / "app"
SERVICES = APP / "services"

PURE_SOLVER_MODULES = [
    SERVICES / "tutor_scheduling_models.py",
    SERVICES / "tutor_candidates.py",
    SERVICES / "tutor_readiness.py",
    SERVICES / "tutor_optimizer.py",
    SERVICES / "tutor_placement_review.py",
    APP / "tutor_domain.py",
]
FORBIDDEN_FOR_SOLVER = ("sqlalchemy", "tutor_models", "module_models", "models", "database", "fastapi")


def imported_modules(path: Path) -> set[str]:
    names: set[str] = set()
    for node in ast.walk(ast.parse(path.read_text(encoding="utf-8"))):
        if isinstance(node, ast.Import):
            names.update(alias.name for alias in node.names)
        elif isinstance(node, ast.ImportFrom):
            base = node.module or ""
            names.add(base)
            names.update(f"{base}.{alias.name}".strip(".") for alias in node.names)
    return names


def leaf_names(modules: set[str]) -> set[str]:
    return {part for name in modules for part in name.split(".") if part}


class TutorArchitectureTests(unittest.TestCase):
    def test_solver_layer_does_not_import_persistence(self):
        for path in PURE_SOLVER_MODULES:
            leaves = leaf_names(imported_modules(path))
            for forbidden in FORBIDDEN_FOR_SOLVER:
                self.assertNotIn(forbidden, leaves, f"{path.name} must not import {forbidden}")

    def test_tutor_modules_do_not_depend_on_clinical_or_other_modules(self):
        tutor_files = [*SERVICES.glob("tutor_*.py"), APP / "tutor_models.py", APP / "tutor_schemas.py"]
        for path in tutor_files:
            for name in imported_modules(path):
                for part in name.split("."):
                    self.assertFalse(part.startswith("clinical"), f"{path.name} imports {name}")
                    self.assertFalse(part.startswith("lesson"), f"{path.name} imports {name}")
                    self.assertFalse(part.startswith("accompanist"), f"{path.name} imports {name}")
                    self.assertNotIn(part, {"jury", "accompanist_results"}, f"{path.name} imports {name}")


if __name__ == "__main__":
    unittest.main()
