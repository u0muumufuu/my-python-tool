"""SciPy を用いた単一・連立非線形方程式の求根ユーティリティ。

- 単一方程式、探索区間あり: scipy.optimize.brentq
- 単一方程式、探索区間なし: scipy.optimize.root
- 連立方程式 (n 本・n 未知数): scipy.optimize.root

数値的な計算不能（初期値・探索区間・残差の NaN/±Inf）と未収束は
``converged=False, x=np.nan`` として返す（連立では全成分 NaN）。
不正な引数構造や関数出力の次元違いは、プログラミングエラーなので例外とする。
``raise_on_failure=True`` を指定した場合だけ、数値的な失敗も ConvergenceError。
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Callable, Mapping

import numpy as np
from scipy.optimize import brentq, root


@dataclass(frozen=True)
class EquationResult:
    """解と収束情報。

    未収束時の x は NaN、連立では全要素 NaN。
    residual は最後の残差（評価不能成分は NaN）。
    last_x は最後に評価を試みた有限の推定値。初期値自体が無効なら None。
    """

    x: float | np.ndarray
    residual: float | np.ndarray
    converged: bool
    method: str
    message: str
    function_calls: int | None
    iterations: int | None
    last_x: float | np.ndarray | None = None

    @property
    def residual_norm(self) -> float:
        """残差の最大絶対値。1成分でも NaN なら NaN。"""
        values = np.asarray(self.residual, dtype=float)
        return float(np.max(np.abs(values)))


class ConvergenceError(RuntimeError):
    """明示的に raise_on_failure=True にした際の収束失敗例外。"""

    def __init__(self, result: EquationResult) -> None:
        self.result = result
        super().__init__(
            f"求根に失敗しました ({result.method}): {result.message} "
            f"[最大絶対残差={result.residual_norm:.6g}]"
        )


class _NonFiniteResidual(Exception):
    """残差 NaN/±Inf だけを solver 内で捕捉する内部用例外。

    元の Python 例外（例: TypeError / ValueError）や残差の shape 違いは
    捕捉せず呼び出し側に伝える。
    """

    def __init__(self, x: float | np.ndarray, residual: float | np.ndarray):
        self.x = float(x) if np.ndim(x) == 0 else np.asarray(x, dtype=float).copy()
        raw = np.asarray(residual, dtype=float)
        normalized = np.where(np.isfinite(raw), raw, np.nan)
        self.residual = float(normalized) if normalized.ndim == 0 else normalized.copy()


class EquationSolver:
    """方程式 f(x)=0 または F(x)=0 を求解するクラス。

    Parameters
    ----------
    root_method : str, default 'hybr'
        scipy.optimize.root の方式。
    tol : float, default 1e-10
        SciPy の収束許容値。brentq では xtol。
    residual_tol : float, default 1e-8
        解として採用するための各残差の最大絶対値の上限。
    maxiter : int, default 100
        brentq の最大反復回数。root は root_options で制御する。
    root_options : dict | None
        scipy.optimize.root に渡すオプション。

    Examples
    --------
    >>> solver = EquationSolver()
    >>> abs(solver.solve(lambda x: x**2 - 4, bracket=(0, 3)).x - 2) < 1e-8
    True
    >>> abs(solver.solve(lambda x: x**2 - 4, x0=1.5).x - 2.0) < 1e-8
    True
    >>> result = solver.solve(lambda v: [v[0]+v[1]-3, v[0]-v[1]-1], x0=[1, 1])
    >>> np.allclose(result.x, [2, 1])
    True
    >>> np.isnan(solver.solve(lambda x: x - 1, x0=np.nan).x)
    True
    """

    def __init__(
        self,
        *,
        root_method: str = "hybr",
        tol: float = 1e-10,
        residual_tol: float = 1e-8,
        maxiter: int = 100,
        root_options: Mapping[str, Any] | None = None,
    ) -> None:
        if not isinstance(root_method, str) or not root_method.strip():
            raise ValueError("root_method は空でない文字列を指定してください。")
        if not np.isfinite(tol) or tol <= 0:
            raise ValueError("tol は正の有限数を指定してください。")
        if not np.isfinite(residual_tol) or residual_tol < 0:
            raise ValueError("residual_tol は0以上の有限数を指定してください。")
        if isinstance(maxiter, bool) or not isinstance(maxiter, int) or maxiter < 1:
            raise ValueError("maxiter は正の整数を指定してください。")
        self.root_method = root_method
        self.tol = float(tol)
        self.residual_tol = float(residual_tol)
        self.maxiter = maxiter
        self.root_options = dict(root_options or {})

    @staticmethod
    def _scalar_output(func: Callable[..., Any], x: float, args: tuple[Any, ...]) -> float:
        values = np.asarray(func(float(x), *args), dtype=float)
        if values.ndim != 0:
            raise ValueError("単一方程式の関数は1個のスカラーを返してください。")
        y = float(values)
        if not np.isfinite(y):
            raise _NonFiniteResidual(float(x), y)
        return y

    @staticmethod
    def _vector_output(
        func: Callable[..., Any], x: np.ndarray, args: tuple[Any, ...], n: int
    ) -> np.ndarray:
        values = np.asarray(func(x, *args), dtype=float)
        if values.shape != (n,):
            raise ValueError(
                f"連立方程式は未知数{n}個に対し残差を長さ{n}の配列で返してください。"
            )
        if not np.isfinite(values).all():
            raise _NonFiniteResidual(x, values)
        return values

    @staticmethod
    def _failure(
        *,
        method: str,
        n: int,
        message: str,
        residual: float | np.ndarray | None = None,
        function_calls: int | None = 0,
        iterations: int | None = 0,
        last_x: float | np.ndarray | None = None,
    ) -> EquationResult:
        nan_x = float("nan") if n == 1 else np.full(n, np.nan, dtype=float)
        if residual is None:
            nan_residual = float("nan") if n == 1 else np.full(n, np.nan, dtype=float)
        else:
            raw = np.asarray(residual, dtype=float)
            finite_only = np.where(np.isfinite(raw), raw, np.nan)
            nan_residual = float(finite_only) if n == 1 else finite_only.copy()
        if last_x is not None:
            last_x = float(last_x) if n == 1 else np.asarray(last_x, dtype=float).copy()
            if not np.isfinite(np.asarray(last_x, dtype=float)).all():
                last_x = None
        return EquationResult(
            x=nan_x, residual=nan_residual, converged=False,
            method=method, message=message,
            function_calls=function_calls, iterations=iterations, last_x=last_x,
        )

    @staticmethod
    def _finish(result: EquationResult, raise_on_failure: bool) -> EquationResult:
        if not result.converged and raise_on_failure:
            raise ConvergenceError(result)
        return result

    def solve(
        self,
        func: Callable[..., Any],
        *,
        x0: float | list[float] | np.ndarray | None = None,
        bracket: tuple[float, float] | None = None,
        args: tuple[Any, ...] = (),
        raise_on_failure: bool = False,
    ) -> EquationResult:
        """方程式を求解し、数値的な失敗は NaN の結果として返す。

        - ``bracket=(lower, upper)``: 単一変数の brentq。
        - ``x0=...``: root。x0 が配列の場合は連立方程式。
        - NaN/±Inf の x0・bracket・残差は未収束として返す。
        - 不正な形状、指定不足、両方式の同時指定は例外。
        - ``raise_on_failure=True`` のときは失敗時に ConvergenceError。
        """
        if not callable(func):
            raise TypeError("func には呼び出し可能な関数を指定してください。")
        if not isinstance(args, tuple):
            raise TypeError("args はタプルで指定してください。")

        if bracket is not None:
            if x0 is not None:
                raise ValueError("bracket と x0 は同時に指定できません。")
            bound = np.asarray(bracket, dtype=float)
            if bound.shape != (2,):
                raise ValueError("bracket は下限 < 上限の数値2個で指定してください。")
            if not np.isfinite(bound).all():
                return self._finish(self._failure(
                    method="brentq", n=1,
                    message="bracket の端点に NaN または無限大が含まれます。",
                ), raise_on_failure)
            if bound[0] >= bound[1]:
                raise ValueError("bracket は下限 < 上限で指定してください。")
            lower, upper = map(float, bound)
            calls = 0

            def scalar_func(x: float) -> float:
                nonlocal calls
                calls += 1
                return self._scalar_output(func, x, args)

            try:
                # 端点が同符号なら、解を推測せず未収束として返す。
                f_lower, f_upper = scalar_func(lower), scalar_func(upper)
                if f_lower != 0 and f_upper != 0 and np.sign(f_lower) == np.sign(f_upper):
                    result = self._failure(
                        method="brentq", n=1,
                        message="bracket の端点の残差が同符号です。解を挟む区間を指定してください。",
                        function_calls=calls,
                    )
                else:
                    if f_lower == 0 or f_upper == 0:
                        x = lower if f_lower == 0 else upper
                        residual = 0.0
                        info_converged = True
                        iterations = 0
                    else:
                        x, info = brentq(
                            scalar_func, lower, upper,
                            xtol=self.tol, maxiter=self.maxiter,
                            full_output=True, disp=False,
                        )
                        x = float(x)
                        residual = scalar_func(x)
                        info_converged = bool(info.converged)
                        iterations = int(info.iterations)
                    ok = info_converged and abs(residual) <= self.residual_tol
                    result = EquationResult(
                        x=float(x) if ok else float("nan"),
                        residual=float(residual), converged=ok, method="brentq",
                        message="収束" if ok else "brentqの終了条件未達、または残差許容値超過",
                        function_calls=calls, iterations=iterations,
                        last_x=None if ok else float(x),
                    )
            except _NonFiniteResidual as bad:
                result = self._failure(
                    method="brentq", n=1,
                    message="方程式の残差が NaN または無限大になりました。",
                    residual=bad.residual, function_calls=calls,
                    iterations=None, last_x=bad.x,
                )
            return self._finish(result, raise_on_failure)

        if x0 is None:
            raise ValueError("探索区間がない場合は初期値 x0 が必要です。")
        initial = np.asarray(x0, dtype=float)
        if initial.ndim > 1 or initial.size == 0:
            raise ValueError("x0 はスカラーまたは空でない1次元配列を指定してください。")
        initial = np.atleast_1d(initial).astype(float)
        n = len(initial)
        scalar = n == 1
        method = f"root/{self.root_method}"
        if not np.isfinite(initial).all():
            return self._finish(self._failure(
                method=method, n=n,
                message="x0 に NaN または無限大が含まれます。",
            ), raise_on_failure)

        calls = 0
        if scalar:
            def vector_func(x: np.ndarray) -> np.ndarray:
                nonlocal calls
                calls += 1
                return np.array([self._scalar_output(func, float(x[0]), args)])
        else:
            def vector_func(x: np.ndarray) -> np.ndarray:
                nonlocal calls
                calls += 1
                return self._vector_output(func, x, args, n)

        try:
            info = root(
                vector_func, initial,
                method=self.root_method,
                tol=self.tol,
                options=self.root_options,
            )
            solution = np.asarray(info.x, dtype=float)
            if not np.isfinite(solution).all():
                result = self._failure(
                    method=method, n=n,
                    message="求根処理が NaN または無限大の推定値を返しました。",
                    function_calls=calls,
                    iterations=int(info.nit) if hasattr(info, "nit") else None,
                )
            else:
                final_residual = vector_func(solution)
                residual_norm = float(np.max(np.abs(final_residual)))
                ok = bool(info.success) and residual_norm <= self.residual_tol
                result = EquationResult(
                    x=(float(solution[0]) if scalar else solution.copy()) if ok else (
                        float("nan") if scalar else np.full(n, np.nan, dtype=float)
                    ),
                    residual=float(final_residual[0]) if scalar else final_residual.copy(),
                    converged=ok, method=method,
                    message=(
                        "収束" if ok else
                        "求根アルゴリズムは終了しましたが残差許容値を超過しました。"
                        if info.success else str(info.message)
                    ),
                    function_calls=int(info.nfev) if hasattr(info, "nfev") else calls,
                    iterations=int(info.nit) if hasattr(info, "nit") else None,
                    last_x=None if ok else (float(solution[0]) if scalar else solution.copy()),
                )
        except _NonFiniteResidual as bad:
            result = self._failure(
                method=method, n=n,
                message="方程式の残差に NaN または無限大が発生しました。",
                residual=bad.residual, function_calls=calls,
                iterations=None, last_x=bad.x,
            )
        return self._finish(result, raise_on_failure)
