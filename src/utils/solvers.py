"""SciPy を用いた単一・連立非線形方程式の求根ユーティリティ。

* 単一方程式、探索区間あり: scipy.optimize.brentq
* 単一方程式、探索区間なし: scipy.optimize.root
* 連立方程式 (n本・n未知数): scipy.optimize.root

「収束」とみなすために SciPy の終了フラグに加えて残差も検査する。
未収束時は result.x を NaN にする（連立方程式では全成分 NaN）。
`residual_tol` は対象方程式の出力単位に合わせて指定すること。
連立方程式では、各式の残差を無次元化・スケーリングすることを推奨する。
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Callable, Mapping

import numpy as np
from scipy.optimize import brentq, root


@dataclass(frozen=True)
class EquationResult:
    """解と収束情報。未収束時の x は NaN（連立では NaN 配列）。

    residual には最後に評価された残差を残す（符号条件違反では NaN）。
    last_x は未収束時の最後の推定値。解が存在しない区間では None。
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
        """残差の最大絶対値。"""
        values = np.asarray(self.residual, dtype=float)
        return float(np.max(np.abs(values)))


class ConvergenceError(RuntimeError):
    """収束失敗時の例外。`result.last_x` に最終推定値を保持する。"""

    def __init__(self, result: EquationResult) -> None:
        self.result = result
        super().__init__(
            f"求根に失敗しました ({result.method}): {result.message} "
            f"[最大絶対残差={result.residual_norm:.6g}]"
        )


class EquationSolver:
    """方程式 f(x)=0 または F(x)=0 を求解するクラス。

    Parameters
    ----------
    root_method : str, default 'hybr'
        区間を指定しない場合の scipy.optimize.root の方式。
    tol : float, default 1e-10
        SciPy に渡す求根の収束許容値。brentq では xtol、root では tol。
    residual_tol : float, default 1e-8
        解として採用するための各残差の最大絶対値の上限。
        連立の場合には各式の残差の尺度を揃えることが望ましい。
    maxiter : int, default 100
        brentq の最大反復回数。root の制限は root_options で指定する
        (例: hybr なら {'maxfev': 300})。
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
            raise ValueError("方程式の残差が NaN または無限大です。")
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
            raise ValueError("方程式の残差に NaN または無限大があります。")
        return values

    def solve(
        self,
        func: Callable[..., Any],
        *,
        x0: float | list[float] | np.ndarray | None = None,
        bracket: tuple[float, float] | None = None,
        args: tuple[Any, ...] = (),
        raise_on_failure: bool = False,
    ) -> EquationResult:
        """方程式を求解する。

        `bracket=(lower, upper)` 指定時は単一変数の brentq を選択。
        指定しない場合は `x0` を初期値として root を選択。
        連立の場合は `x0=[初期値1, 初期値2, ...]` と渡し、
        `func(x)` は同数の残差をリストまたは1次元配列で返す。
        
        未収束では `converged=False` とし、`x` は NaN を返す（既定）。
        連立方程式では x の全要素が NaN となる。
        `raise_on_failure=True` では ConvergenceError を送出する。
        不正な入力や方程式の評価エラーは通常どおり例外を送出する。
        """
        if not callable(func):
            raise TypeError("func には呼び出し可能な関数を指定してください。")
        if not isinstance(args, tuple):
            raise TypeError("args はタプルで指定してください。")
        if bracket is not None:
            if x0 is not None:
                raise ValueError("bracket と x0 は同時に指定できません。")
            bound = np.asarray(bracket, dtype=float)
            if bound.shape != (2,) or not np.isfinite(bound).all() or bound[0] >= bound[1]:
                raise ValueError("bracket は下限 < 上限の有限数2個で指定してください。")
            lower, upper = map(float, bound)

            def scalar_func(x: float) -> float:
                return self._scalar_output(func, x, args)

            # 端点が同符号なら、解を推測せず「未収束」として返す。
            # 下限 >= 上限など、指定方法自体の誤りは上の検証で ValueError。
            f_lower, f_upper = scalar_func(lower), scalar_func(upper)
            if f_lower != 0 and f_upper != 0 and np.sign(f_lower) == np.sign(f_upper):
                result = EquationResult(
                    x=float("nan"), residual=float("nan"), converged=False,
                    method="brentq",
                    message="bracket の端点の残差が同符号です。解を挟む区間を指定してください。",
                    function_calls=2, iterations=0,
                )
            else:
                # 端点が解のケースは brentq を起動せず終了する。
                # SciPy の一部バージョンでは端点解の iterations が不定になるため。
                if f_lower == 0 or f_upper == 0:
                    x = lower if f_lower == 0 else upper
                    residual = 0.0
                    info_converged = True
                    function_calls = 2
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
                    function_calls = int(info.function_calls)
                    iterations = int(info.iterations)
                ok = info_converged and abs(residual) <= self.residual_tol
                result = EquationResult(
                    x=float(x) if ok else float("nan"),
                    residual=float(residual), converged=ok, method="brentq",
                    message="収束" if ok else "brentqの終了条件未達、または残差許容値超過",
                    function_calls=function_calls, iterations=iterations,
                    last_x=None if ok else float(x),
                )
        else:
            if x0 is None:
                raise ValueError("探索区間がない場合は初期値 x0 が必要です。")
            initial = np.asarray(x0, dtype=float)
            if initial.ndim > 1 or initial.size == 0 or not np.isfinite(initial).all():
                raise ValueError("x0 は有限なスカラーまたは1次元配列を指定してください。")
            initial = np.atleast_1d(initial).astype(float)
            n = len(initial)
            scalar = n == 1

            if scalar:
                def vector_func(x: np.ndarray) -> np.ndarray:
                    return np.array([self._scalar_output(func, x[0], args)])
            else:
                def vector_func(x: np.ndarray) -> np.ndarray:
                    return self._vector_output(func, x, args, n)

            info = root(
                vector_func, initial,
                method=self.root_method,
                tol=self.tol,
                options=self.root_options,
            )
            solution = np.asarray(info.x, dtype=float)
            final_residual = vector_func(solution)
            residual_norm = float(np.max(np.abs(final_residual)))
            ok = bool(info.success) and residual_norm <= self.residual_tol
            result = EquationResult(
                x=(float(solution[0]) if scalar else solution.copy()) if ok else (
                    float("nan") if scalar else np.full(n, np.nan, dtype=float)
                ),
                residual=float(final_residual[0]) if scalar else final_residual.copy(),
                converged=ok,
                method=f"root/{self.root_method}",
                message=(
                    "収束" if ok else
                    "求根アルゴリズムは終了しましたが残差許容値を超過しました。"
                    if info.success else str(info.message)
                ),
                function_calls=int(info.nfev) if hasattr(info, "nfev") else None,
                iterations=int(info.nit) if hasattr(info, "nit") else None,
                last_x=None if ok else (float(solution[0]) if scalar else solution.copy()),
            )

        if not result.converged and raise_on_failure:
            raise ConvergenceError(result)
        return result
