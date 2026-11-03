"""Excel 実機なしで共通出力関数のセル操作を検証する。"""

import sys
from pathlib import Path

import numpy as np
import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))
from utils.excel_io import write_excel_result


class FakeRange:
    def __init__(self, sheet, row, col):
        self.sheet, self.row, self.col = sheet, row, col

    @property
    def value(self):
        return self.sheet.data.get((self.row, self.col))

    @value.setter
    def value(self, val):
        if isinstance(val, list):
            for r, line in enumerate(val):
                for c, item in enumerate(line):
                    self.sheet.data[(self.row + r, self.col + c)] = item
        else:
            self.sheet.data[(self.row, self.col)] = val

    @property
    def formula(self):
        return self.sheet.data.get((self.row, self.col))

    @formula.setter
    def formula(self, formula):
        self.sheet.data[(self.row, self.col)] = formula

    def offset(self, row, col):
        return FakeRange(self.sheet, self.row + row, self.col + col)


class FakeSheet:
    def __init__(self):
        self.data = {}

    def range(self, cell):
        import re
        m = re.fullmatch(r"([A-Z]+)([1-9][0-9]*)", cell)
        assert m
        col = 0
        for character in m[1]:
            col = col * 26 + ord(character) - ord("A") + 1
        return FakeRange(self, int(m[2]), col)


def test_single_nan_is_formula():
    sheet = FakeSheet()
    write_excel_result(sheet, "B2", np.nan)
    assert sheet.data[(2, 2)] == "=NA()"


def test_scalar_nonmissing_values():
    sheet = FakeSheet()
    write_excel_result(sheet, "B2", 0.0)
    write_excel_result(sheet, "B3", "OK")
    assert sheet.data[(2, 2)] == 0.0
    assert sheet.data[(3, 2)] == "OK"


def test_infinity_is_na():
    sheet = FakeSheet()
    write_excel_result(sheet, "A1", np.inf)
    write_excel_result(sheet, "A2", -np.inf)
    assert sheet.data[(1, 1)] == "=NA()"
    assert sheet.data[(2, 1)] == "=NA()"


def test_vector_column_default_and_original_unmodified():
    sheet = FakeSheet()
    original = np.array([1.0, np.nan, 2.0])
    write_excel_result(sheet, "D5", original)
    assert sheet.data[(5, 4)] == 1.0
    assert sheet.data[(6, 4)] == "=NA()"
    assert sheet.data[(7, 4)] == 2.0
    assert np.isnan(original[1])


def test_vector_row():
    sheet = FakeSheet()
    write_excel_result(sheet, "A2", [1, np.nan, "x"], orientation="row")
    assert [sheet.data[(2, c)] for c in (1, 2, 3)] == [1, "=NA()", "x"]


def test_matrix_missing_values():
    sheet = FakeSheet()
    write_excel_result(sheet, "B3", [[1, np.nan], [np.inf, -1]])
    assert sheet.data[(3, 2)] == 1
    assert sheet.data[(3, 3)] == "=NA()"
    assert sheet.data[(4, 2)] == "=NA()"
    assert sheet.data[(4, 3)] == -1


def test_none_not_turned_into_na():
    sheet = FakeSheet()
    write_excel_result(sheet, "A1", [[None, np.nan]])
    assert sheet.data[(1, 1)] is None
    assert sheet.data[(1, 2)] == "=NA()"


def test_reject_bad_shape_or_orientation():
    sheet = FakeSheet()
    with pytest.raises(ValueError, match="orientation"):
        write_excel_result(sheet, "A1", [1], orientation="invalid")
    with pytest.raises(ValueError, match="2次元"):
        write_excel_result(sheet, "A1", np.zeros((2, 2, 2)))
    with pytest.raises(ValueError, match="空"):
        write_excel_result(sheet, "A1", [])
