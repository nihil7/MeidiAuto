from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from pipeline.models import PipelineStep
from pipeline.validators import validate_step_support_inputs


class ValidateStepSupportInputsTests(unittest.TestCase):
    def test_accepts_support_file_from_primary_data_dir(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            data_dir = Path(temp_dir)
            (data_dir / "list.xlsx").write_text("placeholder", encoding="utf-8")
            step = PipelineStep("033 list insertion.py", support_patterns=("list.xlsx",))

            ok, message = validate_step_support_inputs(step, (data_dir, data_dir / "legacy"))

            self.assertTrue(ok)
            self.assertEqual(message, "")

    def test_accepts_support_file_from_legacy_script_data_dir(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            data_dir = Path(temp_dir) / "data"
            legacy_dir = Path(temp_dir) / "script-data"
            data_dir.mkdir(parents=True)
            legacy_dir.mkdir(parents=True)
            (legacy_dir / "list.xlsx").write_text("placeholder", encoding="utf-8")
            step = PipelineStep("033 list insertion.py", support_patterns=("list.xlsx",))

            ok, message = validate_step_support_inputs(step, (data_dir, legacy_dir))

            self.assertTrue(ok)
            self.assertEqual(message, "")

    def test_reports_missing_support_file(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            data_dir = Path(temp_dir) / "data"
            legacy_dir = Path(temp_dir) / "legacy"
            data_dir.mkdir(parents=True)
            legacy_dir.mkdir(parents=True)
            step = PipelineStep("033 list insertion.py", support_patterns=("list.xlsx",))

            ok, message = validate_step_support_inputs(step, (data_dir, legacy_dir))

            self.assertFalse(ok)
            self.assertIn("list.xlsx", message)


if __name__ == "__main__":
    unittest.main()
