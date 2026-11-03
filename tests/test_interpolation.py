"""ExcelInterpolator の数値検証テスト（Excel 実機は使わない）。"""

import sys
from pathlib import Path

import numpy as np
import pandas as pd
import pytest
from openpyxl import Workbook, load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))
from utils.interpolation import ExcelInterpolator, ExcelRBFInterpolator


def make_excel(tmp_path, data, sheet="DB", start="A1", filename="input.xlsx"):
    from openpyxl.utils.cell import coordinate_from_string, column_index_from_string
    col, row = coordinate_from_string(start)
    col_idx = column_index_from_string(col)
    wb = Workbook()
    ws = wb.active
    ws.title = sheet
    # 表外の飾り要素 (D5 開始でも A1 に値を持たせる)
    if start != "A1":
        ws["A1"] = "Other information"
    for c, header in enumerate(data.columns):
        ws.cell(row=row, column=col_idx + c, value=header)
    for r, values in enumerate(data.itertuples(index=False, name=None), start=row + 1):
        for c, value in enumerate(values):
            ws.cell(row=r, column=col_idx + c, value=value)
    path = tmp_path / filename
    wb.save(path)
    return path


def grid_data():
    rng = np.random.default_rng(123)
    data = pd.DataFrame([
        dict(alpha=alpha, Mach=mach,
             CL=0.3 + 0.08 * alpha + 2 * mach,
             CD=0.02 + 0.001 * alpha + 0.03 * mach,
             Cm=0.1 - 0.02 * alpha + 0.01 * mach)
        for alpha in [-10., 0., 10.] for mach in [0.1, 0.2, 0.5]
    ])
    return data.iloc[rng.permutation(len(data))].reset_index(drop=True)


def test_regular_grid_unsorted_start_cell_and_multi_output(tmp_path):
    path = make_excel(tmp_path, grid_data(), start="D5")
    interp = ExcelInterpolator(path, "DB", ["alpha", "Mach"], ["CL", "CD", "Cm"],
                               start_cell="D5", method="regular_grid")
    result = interp.predict(alpha=5., Mach=0.3)
    assert result == pytest.approx({"CL": 0.3 + .08 * 5 + 2 * .3,
                                     "CD": .02 + .001 * 5 + .03 * .3,
                                     "Cm": .1 - .02 * 5 + .01 * .3})
    assert interp.n_points == 9


def test_regular_grid_batch_and_broadcast(tmp_path):
    path = make_excel(tmp_path, grid_data())
    interp = ExcelInterpolator(path, "DB", ["alpha", "Mach"], "CL", method="regular_grid")
    values = interp.predict(alpha=[-5., 0., 5.], Mach=.2)["CL"]
    np.testing.assert_allclose(values, .3 + .08 * np.array([-5., 0., 5.]) + .4)
    assert interp.predict(alpha=np.array([[-5.], [5.]]),
                          Mach=np.array([[.1, .2, .5]]))["CL"].shape == (2, 3)
    np.testing.assert_array_equal(interp.predict(alpha=[], Mach=.2)["CL"], np.empty(0))


def test_grid_nearest(tmp_path):
    path = make_excel(tmp_path, grid_data())
    interp = ExcelInterpolator(path, "DB", ["alpha", "Mach"], "CL", method="regular_grid",
                               grid_method="nearest")
    assert interp.predict(alpha=0.1, Mach=.21)["CL"] == pytest.approx(.3 + .4)


def test_regular_grid_missing_point(tmp_path):
    data = grid_data().iloc[:-1]
    path = make_excel(tmp_path, data)
    with pytest.raises(ValueError, match="全格子点"):
        ExcelInterpolator(path, "DB", ["alpha", "Mach"], "CL", method="regular_grid")


def test_regular_grid_duplicate_point(tmp_path):
    data = pd.concat([grid_data(), grid_data().iloc[[0]]], ignore_index=True)
    path = make_excel(tmp_path, data)
    with pytest.raises(ValueError, match="重複"):
        ExcelInterpolator(path, "DB", ["alpha", "Mach"], "CL", method="regular_grid")


def test_rbf_normalization_and_exact_linear_recovery(tmp_path):
    # 速度 1e3, Mach 0.1 のように非常に異なるスケール
    a = np.array([0., 10., 20.])
    h = np.array([0., 4000., 10000.])
    data = pd.DataFrame([
        dict(alpha=x, Altitude=z, CL=2 + .01*x + .00005*z,
             Cm=-.1 + .004*x - .00001*z)
        for x in a for z in h
    ])
    path = make_excel(tmp_path, data)
    it = ExcelInterpolator(path, "DB", ["alpha", "Altitude"], ["CL", "Cm"], method="rbf")
    assert it.x_scale.tolist() == [20., 10000.]
    assert it.predict(alpha=5., Altitude=2000.) == pytest.approx({"CL": 2.15, "Cm": -.1})
    for row in data.itertuples(index=False):
        assert it.predict(alpha=row.alpha, Altitude=row.Altitude) == pytest.approx(
            {"CL": row.CL, "Cm": row.Cm}, rel=1e-10, abs=1e-10
        )


def test_rbf_alias_compatibility(tmp_path):
    data = grid_data()
    path = make_excel(tmp_path, data)
    old = ExcelRBFInterpolator(path, "DB", ["alpha", "Mach"], ["CL", "CD"])
    new = ExcelInterpolator(path, "DB", ["alpha", "Mach"], ["CL", "CD"], method="rbf")
    assert old.predict(alpha=5, Mach=.3) == pytest.approx(new.predict(alpha=5, Mach=.3))


def test_linear_nd_and_nearest_nd(tmp_path):
    data = pd.DataFrame({
        "x": [0., 1., 0., 1., .2],
        "y": [0., 0., 1., 1., .4],
    })
    data["F"] = 1 + 2 * data.x - data.y
    data["G"] = -2 * data.x + 3 * data.y
    path = make_excel(tmp_path, data)
    lin = ExcelInterpolator(path, "DB", ["x", "y"], ["F", "G"], method="linear_nd")
    assert lin.predict(x=.2, y=.3) == pytest.approx({"F": 1.1, "G": .5})
    near = ExcelInterpolator(path, "DB", ["x", "y"], "F", method="nearest_nd")
    assert near.predict(x=.0, y=.0)["F"] == pytest.approx(1.0)


def test_linear_nd_outside_convex_hull(tmp_path):
    # (0.9,0.9) は各軸 [0,1] 内だが三角形 x+y<=1 の外。
    data = pd.DataFrame({"x": [0.,1.,0.], "y": [0.,0.,1.], "F": [0.,1.,1.]})
    path = make_excel(tmp_path, data)
    lin = ExcelInterpolator(path, "DB", ["x", "y"], "F", method="linear_nd")
    with pytest.raises(ValueError, match="凸包外"):
        lin.predict(x=.9, y=.9)


def test_pchip_1d(tmp_path):
    data = pd.DataFrame({"rpm": [1000., 1800., 2500., 3200.],
                         "power": [10., 20., 32., 39.],
                         "fuel": [1., 2., 3., 3.5]})
    path = make_excel(tmp_path, data)
    p = ExcelInterpolator(path, "DB", ["rpm"], ["power", "fuel"], method="pchip")
    assert p.predict(rpm=2500) == pytest.approx({"power": 32., "fuel": 3.})
    assert p.predict(rpm=np.array([1000., 2500.]))["power"].tolist() == pytest.approx([10., 32.])
    assert 20. < p.predict(rpm=2000.)["power"] < 32.


def test_reject_bounds_and_permit_extrapolation_grid(tmp_path):
    path = make_excel(tmp_path, grid_data())
    args = dict(excel_path=path, sheet_name="DB", input_columns=["alpha", "Mach"],
                output_columns="CL", method="regular_grid")
    strict = ExcelInterpolator(**args)
    with pytest.raises(ValueError, match="範囲"):
        strict.predict(alpha=11., Mach=.2)
    loose = ExcelInterpolator(**args, check_bounds=False)
    assert loose.predict(alpha=11., Mach=.2)["CL"] == pytest.approx(.3 + .08 * 11 + .4)


def test_reload_saved_xlsx(tmp_path):
    path = make_excel(tmp_path, grid_data())
    obj = ExcelInterpolator(path, "DB", ["alpha", "Mach"], "CL", method="regular_grid")
    before = obj.predict(alpha=0., Mach=.2)["CL"]
    wb = load_workbook(path)
    ws = wb["DB"]
    # 元データのうち CL 列に一律+1
    for row in range(2, 11):
        ws.cell(row=row, column=3).value += 1
    wb.save(path)
    assert obj.predict(alpha=0., Mach=.2)["CL"] == pytest.approx(before)
    obj.reload()
    assert obj.predict(alpha=0., Mach=.2)["CL"] == pytest.approx(before + 1)


def test_empty_column_and_invalid_method(tmp_path):
    data = grid_data()
    data.loc[3, "CL"] = np.nan
    path = make_excel(tmp_path, data)
    model = ExcelInterpolator(path, "DB", ["alpha", "Mach"], "CL", method="rbf")
    assert model.n_points_by_output["CL"] == 8
    assert np.isfinite(model.predict(alpha=0., Mach=.2)["CL"])
    with pytest.raises(ValueError, match="未対応"):
        ExcelInterpolator(path, "DB", ["alpha", "Mach"], "CL", method="lanczos")
    with pytest.raises(ValueError, match="start_cell"):
        ExcelInterpolator(path, "DB", ["alpha", "Mach"], "CL", start_cell="D5:E7")


def test_missing_header(tmp_path):
    path = make_excel(tmp_path, grid_data())
    with pytest.raises(ValueError, match="列がありません"):
        ExcelInterpolator(path, "DB", ["alpha", "Mach"], ["CL", "Cn"])


def test_one_dim_method_and_all_constant_input(tmp_path):
    data = pd.DataFrame({"x": [1, 1, 1], "y": [2, 3, 4], "z": [1, 3, 6]})
    path = make_excel(tmp_path, data)
    with pytest.raises(ValueError, match="一定"):
        ExcelInterpolator(path, "DB", ["x"], "z", method="pchip")
    with pytest.raises(ValueError, match="1個"):
        ExcelInterpolator(path, "DB", ["x", "y"], "z", method="pchip")


class DummyRange:
    def __init__(self, frame):
        self.frame = frame
        self.kwargs = None

    def options(self, converter, **kwargs):
        assert converter is pd.DataFrame
        self.kwargs = kwargs
        return self

    @property
    def value(self):
        return self.frame.copy()


class DummySheet:
    def __init__(self, frame):
        self.frame = frame
        self.cell = None
        self.range_obj = DummyRange(frame)

    def range(self, cell):
        self.cell = cell
        return self.range_obj


class DummyBook:
    def __init__(self, frame):
        self.sheets = {"DB": DummySheet(frame)}


def test_workbook_unsaved_and_reload(tmp_path):
    path = tmp_path / "nonexistent.xlsm"  # workbook 指定時はディスクの実ファイルがなくても動作
    frame = grid_data()
    wb = DummyBook(frame)
    interp = ExcelInterpolator(path, "DB", ["alpha", "Mach"], "CL",
                               method="regular_grid", start_cell="F10", workbook=wb)
    assert wb.sheets["DB"].cell == "F10"
    assert wb.sheets["DB"].range_obj.kwargs == {"index": False, "expand": "table"}
    before = interp.predict(alpha=0, Mach=.2)["CL"]
    wb.sheets["DB"].frame["CL"] += 3
    interp.reload()  # DummyRange は元frame を保持しているため差し替えではなく mutate
    assert interp.predict(alpha=0, Mach=.2)["CL"] == pytest.approx(before + 3)


def test_regular_grid_one_dim_and_pchip_grid_method(tmp_path):
    d = pd.DataFrame({"rpm": [3000., 1000., 2000., 4000.], "power": [35., 10., 20., 40.]})
    path = make_excel(tmp_path, d)
    lin = ExcelInterpolator(path, "DB", ["rpm"], "power", method="regular_grid")
    assert lin.predict(rpm=1500)["power"] == pytest.approx(15.)
    shape_preserving = ExcelInterpolator(path, "DB", ["rpm"], "power",
                                         method="regular_grid", grid_method="pchip")
    assert shape_preserving.predict(rpm=2000)["power"] == pytest.approx(20.)


def test_scattered_distances_normalized_for_nearest(tmp_path):
    # 2番目の軸は数値レンジ10000だが、正規化後は x,y のレンジがともに1。
    d = pd.DataFrame({"x": [0., 10., 10.], "Altitude": [0., 0., 10000.],
                      "F": [1., 2., 3.]})
    p = make_excel(tmp_path, d)
    nn = ExcelInterpolator(p, "DB", ["x", "Altitude"], "F", method="nearest_nd")
    # normalized (0.1, 0.6) に最も近い (0,0) => 距離 sqrt(.01+.36),
    # (1,1) => sqrt(.81+.16) ; (1,0) => sqrt(.81+.36)
    assert nn.predict(x=1, Altitude=6000)["F"] == pytest.approx(1.)


def test_rbf_bounds_scatter_axis_only_not_convex_hull(tmp_path):
    # 三角形の凸包外でも、RBFは補間値を返す（軸範囲チェックしか実施しない）。
    d = pd.DataFrame({"x": [0., 1., 0., .3],
                      "y": [0., 0., 1., .3],
                      "F": [0., 1., 2., .9]})
    p = make_excel(tmp_path, d)
    rb = ExcelInterpolator(p, "DB", ["x", "y"], "F", method="rbf")
    assert np.isfinite(rb.predict(x=.9, y=.9)["F"])
    with pytest.raises(ValueError, match="範囲"):
        rb.predict(x=1.1, y=.2)


def test_xlsm_loading_and_start_address_lowercase(tmp_path):
    path = make_excel(tmp_path, grid_data(), filename="input.xlsm", start="D5")
    interp = ExcelInterpolator(path, "DB", ["alpha", "Mach"], "CL",
                               method="regular_grid", start_cell="d5")
    assert interp.predict(alpha=0., Mach=.2)["CL"] == pytest.approx(.7)


def test_input_coordinate_validation(tmp_path):
    path = make_excel(tmp_path, grid_data())
    obj = ExcelInterpolator(path, "DB", ["alpha", "Mach"], "CL", method="regular_grid")
    with pytest.raises(ValueError, match="不足"):
        obj.predict(alpha=0.)
    with pytest.raises(ValueError, match="未定義"):
        obj.predict(alpha=0., Mach=.2, beta=0.)
    with pytest.raises(ValueError, match="ブロードキャスト"):
        obj.predict(alpha=[0., 5.], Mach=[.1, .2, .5])
    with pytest.raises(ValueError, match="NaN"):
        obj.predict(alpha=np.nan, Mach=.2)


def test_grid_reload_failed_keeps_old_model(tmp_path):
    path = make_excel(tmp_path, grid_data())
    obj = ExcelInterpolator(path, "DB", ["alpha", "Mach"], "CL", method="regular_grid")
    before = obj.predict(alpha=0., Mach=.2)["CL"]
    wb = load_workbook(path)
    ws = wb["DB"]
    ws.cell(row=2, column=3).value = "invalid numeric text"
    wb.save(path)
    with pytest.raises(ValueError, match="数値でも NA"):
        obj.reload()
    assert obj.predict(alpha=0., Mach=.2)["CL"] == pytest.approx(before)


def test_main_example_with_simulated_excel_book(monkeypatch, tmp_path):
    """main.run の入力→補間→Excel書込→再計算/保存を模擬ブックで確認。"""
    import importlib.util
    import types

    class ScalarRange:
        def __init__(self, sheet, cell):
            self.sheet, self.cell = sheet, cell

        @property
        def value(self):
            return self.sheet.cells.get(self.cell)

        @value.setter
        def value(self, new_value):
            if isinstance(new_value, list) and new_value and isinstance(new_value[0], list):
                from openpyxl.utils.cell import coordinate_from_string
                col, start_row = coordinate_from_string(self.cell)
                for index, row in enumerate(new_value):
                    assert len(row) == 1
                    self.sheet.cells[f"{col}{start_row + index}"] = row[0]
            else:
                self.sheet.cells[self.cell] = new_value

    class TableSheet:
        def __init__(self, frame):
            self.frame = frame
        def range(self, address):
            return DummyRange(self.frame)

    class CellSheet:
        def __init__(self, cells=None):
            self.cells = {} if cells is None else cells
        def range(self, address):
            return ScalarRange(self, address)

    class FakeApp:
        def __init__(self):
            self.calls = 0
        def calculate(self):
            self.calls += 1

    class FakeBook:
        fullname = str(tmp_path / "example.xlsm")
        def __init__(self):
            self.sheets = {
                "空力DB": TableSheet(grid_data()),
                "エンジンDB": TableSheet(pd.DataFrame({
                    "RPM": [1000., 2000., 1000., 2000.],
                    "Altitude": [0., 0., 1000., 1000.],
                    "Power": [12., 22., 10., 20.],
                    "FuelFlow": [1., 2., 1.2, 2.2],
                })),
                "計算条件": CellSheet({"B2": 5., "B3": .2,
                                    "B4": 1500., "B5": 500.}),
                "計算結果": CellSheet(),
            }
            self.app = FakeApp()
            self.saves = 0
        def save(self):
            self.saves += 1

    fake_book = FakeBook()
    fake_xw = types.ModuleType("xlwings")
    fake_xw.Book = types.SimpleNamespace(caller=lambda: fake_book)
    monkeypatch.setitem(sys.modules, "xlwings", fake_xw)
    main_path = Path(__file__).resolve().parents[1] / "src" / "main.py"
    spec = importlib.util.spec_from_file_location("aircraft_main_for_test", main_path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    module.run()

    cells = fake_book.sheets["計算結果"].cells
    assert cells["B2"] == pytest.approx(.3 + .08 * 5 + 2 * .2)
    assert cells["B3"] == pytest.approx(.02 + .001 * 5 + .03 * .2)
    assert cells["B4"] == pytest.approx(.1 - .02 * 5 + .01 * .2)
    assert cells["B6"] == pytest.approx(16.)
    assert cells["B7"] == pytest.approx(1.6)
    assert fake_book.app.calls == 2
    assert fake_book.saves == 1

# -- v5: Excel #N/A の部分欠損を出力列ごとに除外 --

@pytest.mark.parametrize("missing_value", ["#N/A", "=NA()", " NA() ", None, np.nan])
def test_rbf_ignore_na_per_output_column(tmp_path, missing_value):
    data = grid_data()
    selected = data.index[(data.alpha == 0) & (data.Mach == .2)][0]
    data["Cm"] = data["Cm"].astype(object)
    data.loc[selected, "Cm"] = missing_value
    path = make_excel(tmp_path, data, start="D5")
    obj = ExcelInterpolator(path, "DB", ["alpha", "Mach"], ["CL", "CD", "Cm"],
                            start_cell="D5", method="rbf")
    assert obj.n_points_by_output == {"CL": 9, "CD": 9, "Cm": 8}
    assert obj.output_status == {"CL": "complete", "CD": "complete", "Cm": "partial"}
    result = obj.predict(alpha=0, Mach=.2)
    assert result["CL"] == pytest.approx(.3 + 2 * .2)
    assert result["CD"] == pytest.approx(.02 + .03 * .2)
    assert result["Cm"] == pytest.approx(.1 + .01 * .2, abs=1e-8)


def test_regular_grid_one_output_missing_uses_scattered_fallback(tmp_path):
    data = grid_data()
    idx = data.index[(data.alpha == 0) & (data.Mach == .2)][0]
    data["Cm"] = data["Cm"].astype(object)
    data.loc[idx, "Cm"] = "#N/A"
    path = make_excel(tmp_path, data)
    it = ExcelInterpolator(path, "DB", ["alpha", "Mach"], ["CL", "CD", "Cm"],
                           method="regular_grid", grid_method="linear")
    assert it.n_points_by_output == {"CL": 9, "CD": 9, "Cm": 8}
    result = it.predict(alpha=0, Mach=.2)
    assert result == pytest.approx({"CL": .3 + .4, "CD": .02 + .006,
                                    "Cm": .1 + .002}, abs=1e-10)


def test_regular_grid_missing_corner_returns_nan_only_for_that_output(tmp_path):
    d = pd.DataFrame({"x": [0., 0., 1., 1.], "y": [0., 1., 0., 1.],
                      "CL": [1., 2., 3., 4.], "Cm": [5., 6., 7., "#N/A"]})
    path = make_excel(tmp_path, d)
    obj = ExcelInterpolator(path, "DB", ["x", "y"], ["CL", "Cm"],
                            method="regular_grid")
    res = obj.predict(x=1., y=1.)
    assert res["CL"] == pytest.approx(4.)
    assert np.isnan(res["Cm"])
    res_other = obj.predict(x=0., y=1.)
    assert res_other["Cm"] == pytest.approx(6.)


@pytest.mark.parametrize("method", ["linear_nd", "nearest_nd"])
def test_scattered_methods_ignore_missing_outputs(tmp_path, method):
    d = pd.DataFrame({"x": [0, 1, 0, 1, .5], "y": [0, 0, 1, 1, .5],
                      "CL": [0, 1, 1, 2, 1], "Cm": [0, 1, 1, "#N/A", 1]})
    path = make_excel(tmp_path, d)
    obj = ExcelInterpolator(path, "DB", ["x", "y"], ["CL", "Cm"], method=method)
    assert obj.n_points_by_output == {"CL": 5, "Cm": 4}
    res = obj.predict(x=0., y=0.)
    assert res["CL"] == pytest.approx(0.)
    assert res["Cm"] == pytest.approx(0.)


def test_pchip_ignore_one_missing_output(tmp_path):
    d = pd.DataFrame({"x": [0., 1., 2., 3., 4.],
                      "CL": [0., 2., 4., 6., 8.],
                      "Cm": [5., 4., "#N/A", 2., 1.]})
    path = make_excel(tmp_path, d)
    obj = ExcelInterpolator(path, "DB", ["x"], ["CL", "Cm"], method="pchip")
    assert obj.n_points_by_output == {"CL": 5, "Cm": 4}
    assert obj.predict(x=2.) == pytest.approx({"CL": 4., "Cm": 3.})


def test_regular_grid_1d_missing_value_and_nearest(tmp_path):
    d = pd.DataFrame({"x": [0., 1., 2., 3.], "CL": [0., 1., 2., 3.],
                      "Cm": [5., None, 3., 2.]})
    p = make_excel(tmp_path, d)
    linear = ExcelInterpolator(p, "DB", ["x"], ["CL", "Cm"], method="regular_grid")
    nearest = ExcelInterpolator(p, "DB", ["x"], ["CL", "Cm"], method="regular_grid",
                                grid_method="nearest")
    assert linear.predict(x=1.) == pytest.approx({"CL": 1., "Cm": 4.})
    assert nearest.predict(x=.1)["Cm"] == pytest.approx(5.)


def test_all_na_output_does_not_block_other_output(tmp_path):
    d = pd.DataFrame({"x": [0., 1., 2., 3.],
                      "CL": [1., 2., 3., 4.], "Cm": ["#N/A"] * 4})
    p = make_excel(tmp_path, d)
    for method in ["pchip", "rbf", "regular_grid", "nearest_nd"]:
        obj = ExcelInterpolator(p, "DB", ["x"], ["CL", "Cm"], method=method)
        assert obj.output_status["Cm"] == "unavailable"
        res = obj.predict(x=[.5, 1.5])
        if method == "nearest_nd":
            np.testing.assert_allclose(res["CL"], [1., 2.])
        else:
            np.testing.assert_allclose(res["CL"], [1.5, 2.5], atol=1e-8)
        assert np.isnan(res["Cm"]).all()


def test_insufficient_points_in_one_output_does_not_block_other(tmp_path):
    d = pd.DataFrame({"x": [0., 1., 2., 3.], "CL": [1., 2., 3., 4.],
                      "Cm": [1., "#N/A", "#N/A", "#N/A"]})
    p = make_excel(tmp_path, d)
    for method in ["pchip", "rbf", "regular_grid"]:
        it = ExcelInterpolator(p, "DB", ["x"], ["CL", "Cm"], method=method)
        assert it.output_status["Cm"] == "unavailable"
        assert np.isnan(it.predict(x=2.)["Cm"])
        assert np.isfinite(it.predict(x=2.)["CL"])


def test_excel_na_formula_in_saved_file(tmp_path):
    path = make_excel(tmp_path, pd.DataFrame({
        "x": [0., 1., 2., 3.], "CL": [0., 1., 2., 3.],
        "Cm": [2., 3., 4., 5.]}, dtype=float), filename="aero.xlsm")
    wb = load_workbook(path)
    wb["DB"]["C3"] = "=NA()"  # data_only=True ではキャッシュ未設定→None
    wb.save(path)
    it = ExcelInterpolator(path, "DB", ["x"], ["CL", "Cm"], method="pchip")
    assert it.n_points_by_output == {"CL": 4, "Cm": 3}
    assert it.predict(x=1.) == pytest.approx({"CL": 1., "Cm": 3.})


def test_missing_input_skips_whole_row(tmp_path):
    d = pd.DataFrame({"x": [0., 1., "#N/A", 2., 3.],
                      "CL": [0., 1., 99., 2., 3.],
                      "Cm": [1., 2., 99., 3., 4.]})
    p = make_excel(tmp_path, d)
    it = ExcelInterpolator(p, "DB", ["x"], ["CL", "Cm"], method="pchip")
    assert it.n_points == 4
    assert it.n_dropped_input_rows == 1
    assert it.n_points_by_output == {"CL": 4, "Cm": 4}
    assert it.predict(x=1.5) == pytest.approx({"CL": 1.5, "Cm": 2.5})


def test_invalid_text_is_not_silently_ignored(tmp_path):
    d = pd.DataFrame({"x": [0., 1., 2.], "CL": [0., 1., 2.], "Cm": [1., "oops", 3.]})
    p = make_excel(tmp_path, d)
    with pytest.raises(ValueError, match="数値でも NA"):
        ExcelInterpolator(p, "DB", ["x"], ["CL", "Cm"], method="pchip")


def test_workbook_read_na_per_output_and_reload(tmp_path):
    frame = pd.DataFrame({"x": [0., 1., 2., 3.], "CL": [0., 1., 2., 3.],
                          "Cm": [1., 2., None, 4.]})
    wb = DummyBook(frame)
    obj = ExcelInterpolator(tmp_path / "unsaved.xlsm", "DB", ["x"], ["CL", "Cm"],
                            method="pchip", workbook=wb)
    assert obj.n_points_by_output["Cm"] == 3
    assert obj.predict(x=2.) == pytest.approx({"CL": 2., "Cm": 3.})
    wb.sheets["DB"].frame.loc[2, "Cm"] = 3.
    obj.reload()
    assert obj.output_status["Cm"] == "complete"
    assert obj.n_points_by_output["Cm"] == 4


def test_existing_input_grid_holes_still_rejected(tmp_path):
    # 入力座標の組合せそのものが欠けているなら regular_grid は使用不可。
    d = pd.DataFrame({"x": [0., 0., 1.], "y": [0., 1., 0.],
                      "CL": [0., 1., 1.], "Cm": [2., None, 3.]})
    p = make_excel(tmp_path, d)
    with pytest.raises(ValueError, match="全格子点"):
        ExcelInterpolator(p, "DB", ["x", "y"], ["CL", "Cm"], method="regular_grid")
