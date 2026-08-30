"""Opt-in integration validation through the installed Microsoft Excel COM server."""

import os
import shutil
import subprocess
from pathlib import Path

import pytest

from xlspy import XlsbUpdater, XlsbWriter, XlsxUpdater, XlsxWriter


@pytest.mark.skipif(
    os.environ.get("XLSPY_RUN_EXCEL_COM") != "1",
    reason="Set XLSPY_RUN_EXCEL_COM=1 to run Microsoft Excel COM validation",
)
def test_updated_files_open_in_excel_com(tmp_path):
    powershell = shutil.which("powershell.exe") or shutil.which("powershell")
    if powershell is None:
        pytest.skip("Windows PowerShell is not available")

    xlsx_path = tmp_path / "updated.xlsx"
    xlsb_path = tmp_path / "updated.xlsb"
    with XlsxWriter(xlsx_path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet([["old", 1], ["old2", 2]])
    with XlsbWriter(xlsb_path) as writer:
        writer.add_sheet("Data")
        writer.write_sheet([["old", 1], ["old2", 2]])

    xlsx_updater = XlsxUpdater(xlsx_path)
    xlsx_updater.replace_sheet_data("Data", [["new", 10], ["newer", 20]], headers=["Name", "Value"])
    xlsx_updater.save()

    xlsb_updater = XlsbUpdater(xlsb_path)
    xlsb_updater.replace_sheet_data("Data", [["new", 10], ["newer", 20]], headers=["Name", "Value"])
    xlsb_updater.save()

    script = Path(__file__).parents[1] / "tools" / "validate_excel_com.ps1"
    completed = subprocess.run(
        [powershell, "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(script), str(xlsx_path), str(xlsb_path)],
        capture_output=True,
        text=True,
        check=False,
    )
    assert completed.returncode == 0, completed.stdout + "\n" + completed.stderr
