"""SciPy ベースの求根クラスの単一/連立/補間連携テスト。"""

import sys
from pathlib import Path

import numpy as np
import pandas as pd
import pytest
from openpyxl import Workbook

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))
from utils.interpolation import ExcelInterpolator
from utils.solvers import ConvergenceError, EquationSolver


def test_single_with_bracket_brentq():
    solver = EquationSolver()
    result = solver.solve(lambda x: x**2 - 4, bracket=(0, 3))
    assert result.converged
    assert result.method == "brentq"
    assert result.x == pytest.approx(2.)
    assert result.residual_norm < 1e-8
    assert result.function_calls >= 2


def test_single_without_bracket_root():
    result = EquationSolver().solve(lambda x: x**2 - 4, x0=1.5)
    assert result.converged
    assert result.method == "root/hybr"
    assert isinstance(result.x, float)
    assert result.x == pytest.approx(2.)


def test_scalar_with_extra_args():
    result = EquationSolver().solve(
        lambda x, a, b: a*x - b, bracket=(0, 10), args=(2., 6.)
    )
    assert result.x == pytest.approx(3.)


def test_scalar_x0_singleton_list():
    result = EquationSolver().solve(lambda x: x - 3., x0=[0.])
    assert result.x == pytest.approx(3.)


def test_system_two_equations_root():
    def equations(x):
        return [x[0] + x[1] - 3, x[0] - x[1] - 1]
    result = EquationSolver().solve(equations, x0=[0, 0])
    assert result.converged
    assert result.method == "root/hybr"
    assert isinstance(result.x, np.ndarray)
    np.testing.assert_allclose(result.x, [2., 1.], atol=1e-8)
    np.testing.assert_allclose(result.residual, [0., 0.], atol=1e-8)


def test_system_three_nonlinear_equations():
    # x=1, y=2, z=3
    result = EquationSolver().solve(
        lambda v: [v[0]**2 - 1, v[1] - 2, v[2] - v[0] - v[1]],
        x0=[0.8, 1.6, 2.5],
    )
    np.testing.assert_allclose(result.x, [1, 2, 3], atol=1e-8)


def test_failure_no_root_returns_nan_by_default_and_keeps_diagnostics():
    solver = EquationSolver(root_method="lm")
    result = solver.solve(lambda x: x*x + 1, x0=1.)
    assert not result.converged
    assert np.isnan(result.x)
    assert result.residual_norm >= 1.
    assert np.isfinite(result.last_x)
    with pytest.raises(ConvergenceError) as exc:
        solver.solve(lambda x: x*x + 1, x0=1., raise_on_failure=True)
    assert np.isnan(exc.value.result.x)
    assert exc.value.result.residual_norm >= 1.


def test_reject_invalid_bracket():
    solver = EquationSolver()
    for bracket in ((3, 0), (1, 1), (0, np.nan), (1, 2, 3)):
        with pytest.raises(ValueError, match="bracket"):
            solver.solve(lambda x: x - 1, bracket=bracket)
    result = solver.solve(lambda x: x*x + 1, bracket=(-1., 1.))
    assert not result.converged
    assert np.isnan(result.x)
    assert result.method == "brentq"
    assert "同符号" in result.message
    with pytest.raises(ConvergenceError):
        solver.solve(lambda x: x*x + 1, bracket=(-1., 1.), raise_on_failure=True)


def test_validation_of_start_and_bracket():
    solver = EquationSolver()
    with pytest.raises(ValueError, match="初期値"):
        solver.solve(lambda x: x - 1)
    with pytest.raises(ValueError, match="同時"):
        solver.solve(lambda x: x - 1, x0=0., bracket=(-1, 3))
    for x0 in ([], [1, float("inf")], [[1., 2.]]):
        with pytest.raises(ValueError, match="x0"):
            solver.solve(lambda x: x-1, x0=x0)
    with pytest.raises(ValueError, match="tol"):
        EquationSolver(tol=-1)
    with pytest.raises(ValueError, match="residual_tol"):
        EquationSolver(residual_tol=-1)
    with pytest.raises(ValueError, match="maxiter"):
        EquationSolver(maxiter=0)


def test_validate_function_signature_and_value():
    solver = EquationSolver()
    with pytest.raises(ValueError, match="スカラー"):
        solver.solve(lambda x: [x-1, x-2], x0=0.)
    with pytest.raises(ValueError, match="長さ"):
        solver.solve(lambda x: [x[0]-1], x0=[1., 2.])
    with pytest.raises(ValueError, match="NaN"):
        solver.solve(lambda x: float("nan"), bracket=(-1, 1))
    with pytest.raises(TypeError, match="args"):
        solver.solve(lambda x: x-1, x0=0, args=[1])


def test_residual_tolerance_enforced():
    # 解法が正常終了しても残差が閾値より大きければ不合格
    solver = EquationSolver(root_method="lm", residual_tol=1e-12)
    result = solver.solve(lambda x: x*x + 1, x0=1, raise_on_failure=False)
    assert not result.converged
    assert np.isnan(result.x)


def test_system_failure_returns_nan_array_with_shape_preserved():
    solver = EquationSolver(root_method="hybr")
    result = solver.solve(
        lambda v: [v[0]**2 + 1., v[1]**2 + 2.], x0=[0.2, 0.3]
    )
    assert not result.converged
    assert isinstance(result.x, np.ndarray)
    assert result.x.shape == (2,)
    assert np.isnan(result.x).all()
    assert result.last_x.shape == (2,)
    assert result.residual_norm >= 1.


def test_brentq_maxiter_failure_returns_nan_and_last_guess():
    result = EquationSolver(maxiter=1, residual_tol=1e-14).solve(
        lambda x: x**3 - 2., bracket=(0., 2.)
    )
    assert not result.converged
    assert np.isnan(result.x)
    assert np.isfinite(result.last_x)
    assert result.iterations == 1


def test_brentq_endpoint_is_exact_root():
    result = EquationSolver().solve(lambda x: x - 2., bracket=(2., 5.))
    assert result.converged
    assert result.x == 2.0
    assert result.iterations == 0
    assert result.residual == 0.


def test_interpolation_failure_nan_handling_in_main_when_outside_or_no_crossing(monkeypatch, tmp_path):
    """mainの失敗時Excel書き込みは別の小さいモックで確認する。"""
    import importlib.util
    import types

    class Range:
        def __init__(self, frame=None, values=None, cell=None):
            self.frame, self.values, self.cell = frame, values, cell
        def options(self, *_a, **_kw):
            return self
        @property
        def value(self):
            return self.frame.copy() if self.frame is not None else self.values.get(self.cell)
        @value.setter
        def value(self, val):
            self.values[self.cell] = val
        @property
        def formula(self):
            return self.values.get(self.cell)
        @formula.setter
        def formula(self, val):
            self.values[self.cell] = val

    class Sheet:
        def __init__(self, frame=None):
            self.frame = frame
            self.values = {}
        def range(self, cell):
            return Range(frame=self.frame, values=self.values, cell=cell)

    class FakeBook:
        fullname = str(tmp_path / "dummy.xlsm")
        def __init__(self):
            frame = pd.DataFrame({
                "x": [0., 1., 2., 3., 4.],
                "f": [1., 2., 3., 4., 5.],
                "g": [10., 10., 10., 10., 10.],
            })
            self.sheets = {"交点DB": Sheet(frame), "交点結果": Sheet()}
            self.app = type("App", (), {"calculate": lambda self: None})()
        def save(self):
            pass

    book = FakeBook()
    mock_xw = types.ModuleType("xlwings")
    mock_xw.Book = types.SimpleNamespace(caller=lambda: book)
    monkeypatch.setitem(sys.modules, "xlwings", mock_xw)
    src = Path(__file__).resolve().parents[1] / "src" / "main.py"
    spec = importlib.util.spec_from_file_location("example_main_failure", src)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    module.run_intersection_example()
    assert book.sheets["交点結果"].values["B2"] == "=NA()"
    assert book.sheets["交点結果"].values["B3"] == "=NA()"
    assert "同符号" in book.sheets["交点結果"].values["B5"]


def test_interpolation_f_minus_g_root(tmp_path):
    """Excel表の f(x), g(x) を二つの補間器として構築し、f-g=0 を求解。"""
    wb = Workbook()
    ws = wb.active
    ws.title = "交点DB"
    ws.append(["x", "f", "g"])
    for x in [0., 1., 2., 3., 4.]:
        ws.append([x, 2*x+1, 0.5*x+4])
    path = tmp_path / "aircraft.xlsm"
    wb.save(path)

    f_interp = ExcelInterpolator(path, "交点DB", ["x"], "f", method="pchip")
    g_interp = ExcelInterpolator(path, "交点DB", ["x"], "g", method="pchip")

    def difference(x):
        return f_interp.predict(x=x)["f"] - g_interp.predict(x=x)["g"]

    bounded = EquationSolver().solve(difference, bracket=(0., 4.))
    unbounded = EquationSolver().solve(difference, x0=1.)
    assert bounded.x == pytest.approx(2., abs=1e-8)
    assert unbounded.x == pytest.approx(2., abs=1e-8)
    assert abs(bounded.residual) < 1e-8
    assert abs(unbounded.residual) < 1e-8


def test_main_intersection_example_with_simulated_excel_book(monkeypatch, tmp_path):
    """Excel を開かずに RunPython の受け渡し・結果保存までを模擬検証。"""
    import importlib.util
    import types

    class TableRange:
        def __init__(self, frame):
            self.frame = frame
        def options(self, _converter, **_kwargs):
            return self
        @property
        def value(self):
            return self.frame.copy()

    class TableSheet:
        def __init__(self, frame):
            self.frame = frame
        def range(self, _cell):
            return TableRange(self.frame)

    class ResultRange:
        def __init__(self, data, cell):
            self.data, self.cell = data, cell
        @property
        def value(self):
            return self.data.get(self.cell)
        @value.setter
        def value(self, value):
            self.data[self.cell] = value

    class ResultSheet:
        def __init__(self):
            self.data = {}
        def range(self, cell):
            return ResultRange(self.data, cell)

    class App:
        def __init__(self):
            self.calculations = 0
        def calculate(self):
            self.calculations += 1

    class FakeBook:
        fullname = str(tmp_path / "dummy.xlsm")
        def __init__(self):
            self.sheets = {
                "交点DB": TableSheet(pd.DataFrame({
                    "x": [0., 1., 2., 3., 4.],
                    "f": [1., 3., 5., 7., 9.],
                    "g": [4., 4.5, 5., 5.5, 6.],
                })),
                "交点結果": ResultSheet(),
            }
            self.app = App()
            self.saves = 0
        def save(self):
            self.saves += 1

    book = FakeBook()
    mock_xw = types.ModuleType("xlwings")
    mock_xw.Book = types.SimpleNamespace(caller=lambda: book)
    monkeypatch.setitem(sys.modules, "xlwings", mock_xw)
    src = Path(__file__).resolve().parents[1] / "src" / "main.py"
    spec = importlib.util.spec_from_file_location("example_main_intersection", src)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    module.run_intersection_example()

    result = book.sheets["交点結果"].data
    assert result["B2"] == pytest.approx(2.)
    assert result["B3"] == pytest.approx(0., abs=1e-8)
    assert result["B4"] == "brentq"
    assert book.saves == 1
    assert book.app.calculations == 2
