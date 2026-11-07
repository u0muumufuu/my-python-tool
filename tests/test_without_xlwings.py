"""xlwingsなしの実行経路: ファイル入力→計算→別.xlsx出力を実ファイルで検証。"""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook

PROJECT = Path(__file__).resolve().parents[1]
SCRIPT = PROJECT / "src" / "run_without_xlwings.py"


def make_input_excel(path: Path) -> None:
    book = Workbook()
    control = book.active
    control.title = "00_計算管理"
    airframe = book.create_sheet("01_機体諸元")
    airframe["B2"], airframe["B3"] = 40.0, 2.0
    flight = book.create_sheet("03_飛行条件")
    for cell, value in {"B2": 1.225, "B3": 340.0, "B4": 0.0,
                        "B5": 20.0, "B6": 50.0, "B7": 4}.items():
        flight[cell] = value
    aero = book.create_sheet("10_空力DB")
    aero.append(["alpha", "Mach", "CL", "CD", "Cm"])
    for alpha in [0.0, 5.0, 10.0]:
        for mach in [0.05, 0.10, 0.20]:
            aero.append([alpha, mach, 0.2 + 0.1 * alpha, 0.03 + 0.001 * alpha + 0.01 * mach, 0.0])
    prop = book.create_sheet("11_推進DB")
    prop.append(["V", "Altitude", "Thrust"])
    for speed in [20.0, 30.0, 40.0, 50.0]:
        for altitude in [0.0, 1000.0]:
            prop.append([speed, altitude, 120.0 - 0.2 * speed - 0.01 * altitude])
    book.save(path)


def test_no_xlwings_cli_roundtrip_and_preserve_original(tmp_path):
    input_path = tmp_path / "AircraftPerformance.xlsm"
    output_path = tmp_path / "results.xlsx"
    make_input_excel(input_path)
    before = input_path.read_bytes()

    cp = subprocess.run([sys.executable, str(SCRIPT), str(input_path), str(output_path)],
                        capture_output=True, text=True, check=False)
    assert cp.returncode == 0, cp.stderr
    assert input_path.read_bytes() == before  # 入力ブックを決して変更しない
    assert output_path.is_file()
    results = load_workbook(output_path, data_only=False)
    tv = results["30_TV"]
    ps = results["31_Ps"]
    assert tv["A5"].value == "V"
    assert ps["A5"].value == "V"
    assert tv["A6"].value == pytest.approx(20.0)
    assert tv["C6"].value == pytest.approx(116.0)
    assert tv["E9"].value == "NO_LIFT_BALANCE"
    assert ps["B9"].value == "=NA()"
    assert results["00_計算管理"]["B2"].value == "完了"
    results.close()


def test_no_xlwings_cli_error_when_missing_input(tmp_path):
    path = tmp_path / "missing.xlsm"
    output = tmp_path / "result.xlsx"
    cp = subprocess.run([sys.executable, str(SCRIPT), str(path), str(output)],
                        capture_output=True, text=True, check=False)
    assert cp.returncode != 0
    assert not output.exists()
    error_file = output.with_suffix(".error.txt")
    assert "FileNotFoundError" in error_file.read_text(encoding="utf-8")


def test_no_xlwings_will_not_overwrite_source(tmp_path):
    path = tmp_path / "source.xlsm"
    make_input_excel(path)
    sys.path.insert(0, str(PROJECT / "src"))
    from run_without_xlwings import run
    with pytest.raises(ValueError, match="直接書き戻すことは禁止"):
        run(path, path)
