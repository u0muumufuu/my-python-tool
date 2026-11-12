"""SciPy を用いた単一・連立非線形方程式の求根ユーティリティ。

- 単一方程式、探索区間あり: scipy.optimize.brentq
- 単一方程式、探索区間なし: scipy.optimize.root
- 連立方程式 (n 本・n 未知数): scipy.optimize.root
- 複数解探索: solve_multiple_roots (1変数の区間走査、または多点初期値)

数値的な計算不能（初期値・探索区間・残差の NaN/±Inf）と未収束は
``converged=False, x=np.nan`` として返す（連立では全成分 NaN）。
不正な引数構造や関数出力の次元違いは、プログラミングエラーなので例外とする。
``raise_on_failure=True`` を指定した場合だけ、数値的な失敗も ConvergenceError。
"""

from __future__ import annotations

from dataclasses import dataclass, replace
from typing import Any, Callable, Mapping

import numpy as np
from scipy.optimize import brentq, root
from scipy.stats import qmc


@dataclass(frozen=True)
class EquationResult:
    """解と収束情報。

    未収束時の x は NaN、連立では全要素 NaN。
    residual は最後の残差（評価不能成分は NaN）。
    last_x は最後に評価を試みた有限の推定値。初期値自体が無効なら None。
    initial_x / search_bracket は solve_multiple_roots で発見した解の探索元。
    """

    x: float | np.ndarray
    residual: float | np.ndarray
    converged: bool
    method: str
    message: str
    function_calls: int | None
    iterations: int | None
    last_x: float | np.ndarray | None = None
    initial_x: float | np.ndarray | None = None
    search_bracket: tuple[float, float] | None = None

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

    @staticmethod
    def _multi_bounds(bounds: Any) -> np.ndarray:
        """変数ごとの (lower, upper) を (n, 2) の配列に変換する。"""
        try:
            values = np.asarray(bounds, dtype=float)
        except (TypeError, ValueError) as exc:
            raise ValueError("bounds は数値の [(lower, upper), ...] で指定してください。") from exc
        if values.shape == (2,):
            values = values.reshape(1, 2)
        if values.ndim != 2 or values.shape[1] != 2 or values.shape[0] < 1:
            raise ValueError("bounds は [(lower, upper), ...] の形式で指定してください。")
        # NaN/Inf は solve(x0=np.nan) と同様に数値的失敗とする。
        # 呼び出し側で空リストを返し、残差関数は一度も評価しない。
        if np.isfinite(values).all() and np.any(values[:, 0] >= values[:, 1]):
            raise ValueError("bounds の各組は下限 < 上限で指定してください。")
        return values

    @staticmethod
    def _multi_starts(
        bounds: np.ndarray, *, n_starts: int, strategy: str, seed: int | None
    ) -> np.ndarray:
        """root に渡す初期値を (n_starts, n_unknowns) で生成する。"""
        dimensions = len(bounds)
        if strategy == "sobol":
            # random_base2 は 2 の冪個生成しており、SciPy の警告も出ない。
            m = (n_starts - 1).bit_length()
            unit = qmc.Sobol(d=dimensions, scramble=True, seed=seed).random_base2(m)[:n_starts]
        elif strategy == "random":
            unit = np.random.default_rng(seed).random((n_starts, dimensions))
        else:  # grid
            n_axis = int(np.ceil(n_starts ** (1.0 / dimensions)))
            # 超高次元の直積でメモリを使い尽くさないための保護。
            if n_axis ** dimensions > 1_000_000:
                raise ValueError("grid の格子点が多すぎます。sobol または random を使用してください。")
            axes = [np.array([0.5]) if n_axis == 1 else np.linspace(0.0, 1.0, n_axis)
                    for _ in range(dimensions)]
            mesh = np.meshgrid(*axes, indexing="ij")
            full = np.stack([a.ravel() for a in mesh], axis=-1)
            # 先頭だけを切り出すと特定の変数に偏るので全体から等間隔抽出。
            indices = np.linspace(0, len(full) - 1, n_starts).astype(int)
            unit = full[indices]
        return bounds[:, 0] + unit * (bounds[:, 1] - bounds[:, 0])

    @staticmethod
    def _unique_multi_root(
        found: list[EquationResult], candidate: EquationResult, limits: np.ndarray
    ) -> None:
        """各変数の絶対誤差が limits 以下なら同一解としてまとめる。"""
        x = np.atleast_1d(np.asarray(candidate.x, dtype=float))
        for i, prev in enumerate(found):
            old = np.atleast_1d(np.asarray(prev.x, dtype=float))
            if np.all(np.abs(x - old) <= limits):
                if candidate.residual_norm < prev.residual_norm:
                    found[i] = candidate
                return
        found.append(candidate)

    def solve_multiple_roots(
        self,
        func: Callable[..., Any],
        bounds: Any,
        *,
        method: str = "auto",
        scan_points: int = 100,
        n_starts: int = 100,
        initial_strategy: str = "sobol",
        unique_tol: float | list[float] | np.ndarray = 1e-6,
        args: tuple[Any, ...] = (),
        seed: int | None = 0,
    ) -> list[EquationResult]:
        """指定した探索範囲から、重複しない複数の求根候補を探索する。

        Parameters
        ----------
        func : callable
            1変数では ``func(x, *args) -> float``、n変数では
            ``func(x, *args) -> 長さnの1次元配列``。
        bounds : sequence[(float, float)]
            変数の順に並べた探索範囲。1変数のみ ``(low, high)`` も可。
            ``root`` の反復中の制約ではなく、初期値生成・最終解判定に使用。
        method : {'auto', 'brentq', 'root'}
            auto は 1変数で brentq、2変数以上で root を使用。
            brentq は1変数のみ使用可能。
        scan_points : int
            brentq で探索区間を分割する個数（格子点は +1）。
        n_starts : int
            root を呼び出す初期値の数。
        initial_strategy : {'sobol', 'random', 'grid'}
            root の初期値生成方式。
        unique_tol : float or (n,) sequence
            個々の変数について重複とみなす絶対距離。単位は各変数の単位。
        args : tuple
            方程式に追加する引数。
        seed : int | None
            Sobol/乱数の生成 seed。デフォルトは再現性のため 0。

        Returns
        -------
        list[EquationResult]
            有効で異なる解のみを昇順（多変数は辞書順）で返す。
            1変数は ``result.x: float``、連立は ``result.x: np.ndarray(n,)``。
            各解は ``converged=True`` で、残差・計算回数・探索元も保持。
            発見できない場合、または bounds に NaN/Inf がある場合は ``[]``。

        Notes
        -----
        - 全解が発見できる保証はない。
        - brentq は隣接する有限値の符号変化と、評価値が厳密に0の格子点を探索。
          接するだけの解や、狭い区間内の偶数個の解を見逃す場合がある。
        - root は複数初期値から solve を繰り返す。範囲外に収束した解は除外。
          root 内部で NaN が生じた試行は、その試行だけ破棄する。
        - プログラムの指定ミス・残差ベクトルの次元不一致・例外は隠蔽しない。
        - 各解の function_calls は探索全体ではなく、その解の求根実行に対応。
        """
        if not callable(func):
            raise TypeError("func には呼び出し可能な関数を指定してください。")
        if not isinstance(args, tuple):
            raise TypeError("args はタプルで指定してください。")
        if not isinstance(method, str) or method not in ("auto", "brentq", "root"):
            raise ValueError("method は 'auto', 'brentq', 'root' から指定してください。")
        if not isinstance(initial_strategy, str) or initial_strategy not in (
            "sobol", "random", "grid"
        ):
            raise ValueError("initial_strategy は 'sobol', 'random', 'grid' から指定してください。")
        for name, val, lower in (("scan_points", scan_points, 2), ("n_starts", n_starts, 1)):
            if isinstance(val, bool) or not isinstance(val, (int, np.integer)) or val < lower:
                raise ValueError(f"{name} は {lower} 以上の整数で指定してください。")
        if seed is not None and (
            isinstance(seed, bool) or not isinstance(seed, (int, np.integer)) or seed < 0
        ):
            raise ValueError("seed は0以上の整数または None にしてください。")
        scan_points, n_starts = int(scan_points), int(n_starts)
        seed = int(seed) if seed is not None else None

        limits_bounds = self._multi_bounds(bounds)
        n = len(limits_bounds)
        use_method = ("brentq" if n == 1 else "root") if method == "auto" else method
        if use_method == "brentq" and n != 1:
            raise ValueError("brentq は1方程式（1変数）でのみ利用できます。")
        try:
            limits = np.asarray(unique_tol, dtype=float)
        except (TypeError, ValueError) as exc:
            raise ValueError("unique_tol は数値で指定してください。") from exc
        if limits.ndim == 0:
            limits = np.full(n, float(limits))
        if limits.shape != (n,) or not np.isfinite(limits).all() or np.any(limits <= 0):
            raise ValueError("unique_tol は正の有限数、または未知数と同数の正の数で指定してください。")
        if not np.isfinite(limits_bounds).all():
            return []

        found: list[EquationResult] = []
        if use_method == "brentq":
            lower, upper = limits_bounds[0]
            grid = np.linspace(lower, upper, scan_points + 1)
            sampled = np.full(len(grid), np.nan, dtype=float)
            for i, x in enumerate(grid):
                try:
                    sampled[i] = self._scalar_output(func, float(x), args)
                except _NonFiniteResidual:
                    # 欠損点を挟んだ符号変化は連続性を仮定できないため使用しない。
                    continue

            for i, (x, y) in enumerate(zip(grid, sampled)):
                if np.isfinite(y) and y == 0.0:
                    direct = EquationResult(
                        x=float(x), residual=0.0, converged=True,
                        method="brentq", message="走査格子点で残差が厳密に0です。",
                        function_calls=1, iterations=0,
                    )
                    self._unique_multi_root(found, direct, limits)
                if i + 1 >= len(grid):
                    continue
                next_y = sampled[i + 1]
                if not (np.isfinite(y) and np.isfinite(next_y)):
                    continue
                if (y < 0 < next_y) or (next_y < 0 < y):
                    bracket = (float(x), float(grid[i + 1]))
                    candidate = self.solve(func, bracket=bracket, args=args)
                    if candidate.converged:
                        candidate = replace(candidate, search_bracket=bracket)
                        self._unique_multi_root(found, candidate, limits)
        else:
            starts = self._multi_starts(
                limits_bounds, n_starts=n_starts,
                strategy=initial_strategy, seed=seed,
            )
            for start in starts:
                start_x = float(start[0]) if n == 1 else start.copy()
                candidate = self.solve(func, x0=start_x, args=args)
                if not candidate.converged:
                    continue
                x = np.atleast_1d(np.asarray(candidate.x, dtype=float))
                # 丸め誤差の範囲だけ境界の超過を許す（実質的に範囲内判定）。
                boundary_eps = 8 * np.finfo(float).eps * np.maximum(
                    1, np.max(np.abs(limits_bounds), axis=1)
                )
                if np.any(x < limits_bounds[:, 0] - boundary_eps) or np.any(
                    x > limits_bounds[:, 1] + boundary_eps
                ):
                    continue
                candidate = replace(candidate, initial_x=start_x)
                self._unique_multi_root(found, candidate, limits)
        found.sort(key=lambda result: tuple(np.atleast_1d(np.asarray(result.x, dtype=float))))
        return found

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
