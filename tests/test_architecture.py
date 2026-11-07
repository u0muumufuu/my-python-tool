"""共通補間モデルと性能マップ、および main の Excel 出力の単体検証。"""

from __future__ import annotations

from pathlib import Path
import sys

import numpy as np
import pandas as pd
import pytest
from openpyxl import Workbook

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from models.aerodynamics import AerodynamicsModel
from models.propulsion import PropulsionModel
from analysis.performance_maps import calculate_tv_ps
from main import read_conditions, _write_result_table


@pytest.fixture
def sample_workbook(tmp_path):
    filename = tmp_path / "aircraft.xlsx"
    book = Workbook()
    aero = book.active
    aero.title = "10_空力DB"
    aero.append(["alpha", "Mach", "CL", "CD", "Cm"])
    for a in [0.0, 5.0, 10.0]:
        for mach in [0.05, 0.1, 0.2]:
            cl = 0.2 + 0.1 * a
            aero.append([a, mach, cl, 0.03 + 0.001 * a + 0.01 * mach, 0.0])
    prop = book.create_sheet("11_推進DB")
    prop.append(["V", "Altitude", "Thrust"])
    for v in [20, 30, 40, 50]:
        for h in [0, 1000]:
            prop.append([v, h, 120 - 0.2 * v - 0.01 * h])
    book.save(filename)
    return filename


def test_models_built_from_independent_excel_file(sample_workbook):
    aero = AerodynamicsModel(excel_path=sample_workbook)
    prop = PropulsionModel(excel_path=sample_workbook)
    assert aero.alpha_bounds == (0.0, 10.0)
    assert aero.supports_mach(0.1)
    assert not aero.supports_mach(0.3)
    assert aero.coefficients(alpha=5, mach=0.1)["CL"] == pytest.approx(0.7)
    assert prop.thrust(velocity=30, altitude=0) == pytest.approx(114.0)


def test_maps_success_and_no_solution(sample_workbook):
    aero = AerodynamicsModel(excel_path=sample_workbook)
    prop = PropulsionModel(excel_path=sample_workbook)
    df = calculate_tv_ps(aero, prop, mass_kg=40, wing_area_m2=2,
        rho_kg_m3=1.225, speed_of_sound_m_s=340, altitude_m=0,
        velocities_m_s=[20, 30, 40, 50])
    assert df.loc[0, "status"] == "OK"
    assert df.loc[1, "status"] == "OK"
    assert df.loc[2, "status"] == "OK"
    assert df.loc[3, "status"] == "NO_LIFT_BALANCE"
    assert np.isnan(df.loc[3, "Ps"])
    assert np.isfinite(df.loc[0, "Ps"])
    assert df.loc[1, "CL_required"] == pytest.approx(40*9.80665/(0.5*1.225*30**2*2))


def test_outside_map_returns_nan(sample_workbook):
    aero = AerodynamicsModel(excel_path=sample_workbook)
    prop = PropulsionModel(excel_path=sample_workbook)
    df = calculate_tv_ps(aero, prop, mass_kg=40, wing_area_m2=2,
        rho_kg_m3=1.225, speed_of_sound_m_s=340, altitude_m=0,
        velocities_m_s=[60])
    assert df.loc[0, "status"] == "OUT_OF_MAP"
    assert np.isnan(df.loc[0, "T_required"])


def test_bad_conditions_fail(sample_workbook):
    aero = AerodynamicsModel(excel_path=sample_workbook)
    prop = PropulsionModel(excel_path=sample_workbook)
    with pytest.raises(ValueError, match="質量"):
        calculate_tv_ps(aero, prop, mass_kg=0, wing_area_m2=2,
            rho_kg_m3=1.2, speed_of_sound_m_s=340, altitude_m=0,
            velocities_m_s=[30])


class FakeRange:
    def __init__(self, sheet, cell):
        self.sheet, self.cell = sheet, cell
    @property
    def value(self):
        return self.sheet.cells.get(self.cell)
    @value.setter
    def value(self, value):
        self.sheet.cells[self.cell] = value
    @property
    def formula(self):
        return self.value
    @formula.setter
    def formula(self, formula):
        self.value = formula
    def clear_contents(self):
        self.sheet.cleared.append(self.cell)
    def offset(self, row, col):
        # This test only uses B2 as starting point with 2D values.
        from openpyxl.utils.cell import coordinate_from_string, column_index_from_string, get_column_letter
        c, r = coordinate_from_string(self.cell)
        return FakeRange(self.sheet, f"{get_column_letter(column_index_from_string(c) + col)}{r+row}")

class FakeSheet:
    def __init__(self, cells=None):
        self.cells = dict(cells or {})
        self.cleared = []
    def range(self, cell):
        return FakeRange(self, cell)

class FakeBook:
    def __init__(self):
        self.sheets = {
            "01_機体諸元": FakeSheet({"B2": 40, "B3": 2}),
            "03_飛行条件": FakeSheet({"B2": 1.225, "B3": 340, "B4": 0,
                "B5": 20, "B6": 40, "B7": 5}),
        }


def test_excel_inputs_with_fake_book():
    values = read_conditions(FakeBook())
    assert np.allclose(values["velocities_m_s"], [20, 25, 30, 35, 40])
    assert values["wing_area_m2"] == 2


def test_excel_table_nan_to_na_formula():
    ws = FakeSheet()
    df = pd.DataFrame({"V": [30, 50], "Ps": [2.5, np.nan], "status": ["OK", "NO_LIFT_BALANCE"]})
    _write_result_table(ws, ["V", "Ps", "status"], df)
    assert "A5:C10005" in ws.cleared
    assert ws.cells["A5"] == [["V", "Ps", "status"]]
    assert ws.cells["A6"] == [[30, 2.5, "OK"], [50, None, "NO_LIFT_BALANCE"]]
    assert ws.cells["B7"] == "=NA()"


def test_run_maps_from_excel_button_simulated(sample_workbook, monkeypatch):
    """ボタン相当のBook.caller→DB読込→計算→Excel書き込み→保存を模擬。"""
    import types
    from openpyxl import load_workbook
    import main

    class DataRange:
        def __init__(self, frame): self.frame = frame
        def options(self, *a, **kw): return self
        @property
        def value(self): return self.frame.copy()

    class DataSheet:
        def __init__(self, frame): self.frame = frame
        def range(self, cell): return DataRange(self.frame)

    class App:
        def __init__(self): self.calls = 0
        def calculate(self): self.calls += 1

    class Book:
        fullname = str(sample_workbook)
        def __init__(self):
            self.sheets = {
                "10_空力DB": DataSheet(pd.read_excel(sample_workbook, sheet_name="10_空力DB")),
                "11_推進DB": DataSheet(pd.read_excel(sample_workbook, sheet_name="11_推進DB")),
                "00_計算管理": FakeSheet(),
                "01_機体諸元": FakeSheet({"B2": 40.0, "B3": 2.0}),
                "03_飛行条件": FakeSheet({"B2": 1.225, "B3": 340.0, "B4": 0.0,
                    "B5": 20.0, "B6": 50.0, "B7": 4}),
                "30_TV": FakeSheet(),
                "31_Ps": FakeSheet(),
            }
            self.app = App()
            self.saves = 0
        def save(self): self.saves += 1

    wb = Book()
    fake_xw = types.ModuleType("xlwings")
    fake_xw.Book = types.SimpleNamespace(caller=lambda: wb)
    monkeypatch.setitem(sys.modules, "xlwings", fake_xw)
    main.run_maps()

    assert wb.saves == 1
    assert wb.app.calls == 2
    assert wb.sheets["00_計算管理"].cells["B2"] == "完了"
    tv = wb.sheets["30_TV"]
    ps = wb.sheets["31_Ps"]
    assert tv.cells["A5"] == [["V", "T_required", "T_available", "alpha", "status"]]
    assert ps.cells["A5"] == [["V", "Ps", "alpha", "status"]]
    assert ps.cells["B9"] == "=NA()"  # 最後の速度は CL 条件を満たす迎角なし
