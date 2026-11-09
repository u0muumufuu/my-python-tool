"""Excel データから補間モデルを作成するユーティリティ。

出力列ごとに欠損値を除外して学習する。Excel の =NA() / #N/A や
Python の np.nan は欠損値として扱い、他の出力列の有効値を捨てない。

Excel 表の前提:
- ``start_cell`` (初期値 A1) にヘッダー行の左端セルがある。
- ヘッダーは右方向に連続し、その下に数値行が連続する。
- 指定データ範囲内には完全空白行・空白列を設けない。
- ディスク上のファイルは .xlsx / .xlsm。xls は未対応。

保存済みファイルは openpyxl から読み込む。workbook を渡した場合は
xlwings のブックから未保存のセル値も読み込める。どちらの場合も
補間器の作成時と reload() 時だけ Excel を読む。
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any

import numpy as np
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils.cell import coordinate_from_string, column_index_from_string
from scipy.interpolate import (
    LinearNDInterpolator,
    NearestNDInterpolator,
    PchipInterpolator,
    RBFInterpolator,
    RegularGridInterpolator,
)


class ExcelInterpolator:
    """Excel の指定テーブルから多入力・多出力補間器を作成する。

    Parameters
    ----------
    excel_path : str | Path
        Excel ファイルパス。workbook がある場合も管理情報として指定。
    sheet_name : str
        数値テーブルがあるシート。
    input_columns : list[str]
        補間関数の入力列名 (指定順が内部の次元順になる)。
    output_columns : str | list[str]
        出力列名。複数列をまとめて補間できる。
    method : str, default 'rbf'
        'rbf', 'regular_grid', 'linear_nd', 'nearest_nd', 'pchip'。
        'pchip' は 1 入力のみ。
    start_cell : str, default 'A1'
        見出し行の左端セル。例: 'D5'。
    workbook : xlwings.Book | None
        Excel で開いている Book オブジェクト。未保存の編集も読み込む。
        指定しない場合は openpyxl で保存済みファイルを読む。
    check_bounds : bool, default True
        各入力列の最小値～最大値の範囲外を禁止する。
        ※RBF/散在点方式では凸包外判定まではしない。
    smoothing : float, default 0
        RBF の平滑化パラメータ。
    neighbors : int | None, default None
        RBF 評価時に使う近傍点数。None は全点。
    rbf_kernel : str, default 'thin_plate_spline'
        RBF のカーネル名。
    rbf_epsilon : float | None, default None
        RBF の形状パラメータ。一部のカーネルで必須。
    grid_method : str, default 'linear'
        regular_grid の補間方式: linear/nearest/slinear/cubic/quintic/pchip。

    Notes
    -----
    RBF と散在点 (linear_nd/nearest_nd) では、各入力を
    ``(x - min) / (max - min)`` に正規化してから構築・評価する。
    特に RBF の距離尺度が各入力単位に依存しないようにするため。
    regular_grid は元の軸を使い、入力座標の全組み合わせの存在を確認する。
    特定の出力に欠損がある場合、その出力だけ有効点で補間する。
    有効点が不足する出力は np.nan を返し、他の出力は補間できる。
    """

    METHODS = frozenset({"rbf", "regular_grid", "linear_nd", "nearest_nd", "pchip"})
    GRID_METHODS = frozenset({"linear", "nearest", "slinear", "cubic", "quintic", "pchip"})
    _CELL_REGEX = re.compile(r"^\$?[A-Za-z]{1,3}\$?[1-9][0-9]*$")

    def __init__(
        self,
        excel_path: str | Path,
        sheet_name: str,
        input_columns: list[str] | tuple[str, ...],
        output_columns: str | list[str] | tuple[str, ...],
        *,
        method: str = "rbf",
        start_cell: str = "A1",
        workbook: Any = None,
        check_bounds: bool = True,
        smoothing: float = 0.0,
        neighbors: int | None = None,
        rbf_kernel: str = "thin_plate_spline",
        rbf_epsilon: float | None = None,
        grid_method: str = "linear",
    ) -> None:
        self.excel_path = Path(excel_path).expanduser().resolve()
        self.sheet_name = sheet_name
        self.input_columns = list(input_columns)
        self.output_columns = [output_columns] if isinstance(output_columns, str) else list(output_columns)
        self.method = method.lower()
        self.start_cell = start_cell.upper().replace("$", "")
        self.workbook = workbook
        self.check_bounds = bool(check_bounds)
        self.smoothing = float(smoothing)
        self.neighbors = neighbors
        self.rbf_kernel = rbf_kernel
        self.rbf_epsilon = rbf_epsilon
        self.grid_method = grid_method.lower()

        if not self._CELL_REGEX.fullmatch(start_cell):
            raise ValueError(f"start_cell は D5 のような単一セルで指定してください: {start_cell!r}")
        column, row = coordinate_from_string(self.start_cell)
        self._start_row = row
        self._start_col = column_index_from_string(column)
        if not (1 <= row <= 1_048_576 and 1 <= self._start_col <= 16_384):
            raise ValueError("start_cell が Excel の有効な行・列範囲外です。")
        if self.method not in self.METHODS:
            raise ValueError(f"未対応 method={method!r}。使用可能: {sorted(self.METHODS)}")
        if self.grid_method not in self.GRID_METHODS:
            raise ValueError(f"未対応 grid_method={grid_method!r}。使用可能: {sorted(self.GRID_METHODS)}")
        if not self.input_columns or not self.output_columns:
            raise ValueError("入力列と出力列はそれぞれ1つ以上指定してください。")
        columns = self.input_columns + self.output_columns
        if any(not isinstance(c, str) or not c.strip() for c in columns):
            raise ValueError("列名には空でない文字列を指定してください。")
        if len(set(columns)) != len(columns):
            raise ValueError("入力列または出力列の名前が重複しています。")
        if self.smoothing < 0 or not np.isfinite(self.smoothing):
            raise ValueError("smoothing は有限の0以上の値を指定してください。")
        if self.neighbors is not None and (
            isinstance(self.neighbors, bool)
            or not isinstance(self.neighbors, int)
            or self.neighbors < 1
        ):
            raise ValueError("neighbors は正の整数または None を指定してください。")
        if self.method == "pchip" and len(self.input_columns) != 1:
            raise ValueError("method='pchip' は入力列が1個の場合のみ使用できます。")
        if self.method == "linear_nd" and len(self.input_columns) < 2:
            raise ValueError("1次元の線形補間は method='regular_grid' を使ってください。")
        if self.method != "rbf" and (self.smoothing != 0 or self.neighbors is not None):
            raise ValueError("smoothing / neighbors は method='rbf' 専用です。")
        self.reload()

    def _read_table(self) -> pd.DataFrame:
        """指定したヘッダー開始セルから連続表を DataFrame にする。"""
        if self.workbook is not None:
            # Excel 上の未保存値も取得可能。xlwings が返す DataFrame。
            df = (
                self.workbook.sheets[self.sheet_name]
                .range(self.start_cell)
                .options(pd.DataFrame, index=False, expand="table")
                .value
            )
            if not isinstance(df, pd.DataFrame):
                raise ValueError("Excel シートから表データを取得できませんでした。")
            return df

        if self.excel_path.suffix.lower() not in {".xlsx", ".xlsm"}:
            raise ValueError("保存済みファイルの読み込みは .xlsx / .xlsm に対応します。")
        # read_only + data_only: Excel 数式がある場合は最終保存時の計算済み値を読む。
        wb = load_workbook(self.excel_path, read_only=True, data_only=True)
        try:
            if self.sheet_name not in wb.sheetnames:
                raise ValueError(f"シート {self.sheet_name!r} が存在しません。")
            ws = wb[self.sheet_name]
            rows = ws.iter_rows(
                min_row=self._start_row, min_col=self._start_col, values_only=True
            )
            header_row = next(rows, None)
            if header_row is None or not header_row or header_row[0] is None:
                raise ValueError(f"{self.sheet_name}!{self.start_cell} に見出しがありません。")
            headers = []
            for value in header_row:
                if value is None or (isinstance(value, str) and not value.strip()):
                    break
                headers.append(value)
            if any(not isinstance(h, str) for h in headers):
                raise ValueError("ヘッダーには文字列を指定してください。")
            if len(headers) != len(set(headers)):
                raise ValueError("Excel のヘッダー名が重複しています。")
            values = []
            for row in rows:
                data_row = list(row[:len(headers)])
                if all(v is None or (isinstance(v, str) and not v.strip()) for v in data_row):
                    break
                values.append(data_row)
            return pd.DataFrame(values, columns=headers)
        finally:
            wb.close()

    @staticmethod
    def _numeric_column(series: pd.Series, column: str) -> pd.Series:
        """数値または既知の欠損値を受け付ける。不正な文字列はエラー。"""
        # Excel の =NA() は、xlwings 経由では None、ファイル経由では
        # データキャッシュにより None または '#N/A' で返ることがある。
        # '=NA()' という文字列そのものにも対応する。
        tokens = series.astype("string").str.strip().str.upper().str.replace(
            r"\s+", "", regex=True
        )
        missing = series.isna() | tokens.isin({"", "#N/A", "=NA()", "NA()"})
        converted = pd.to_numeric(series.mask(missing), errors="coerce")
        invalid = converted.isna() & ~missing
        if invalid.any():
            bad_row = series.index[invalid][0]
            raise ValueError(
                f"列 {column!r} のデータ行 {bad_row + 1} に数値でも NA() でもない値があります: "
                f"{series.loc[bad_row]!r}"
            )
        numeric = converted.to_numpy(dtype=float)
        if np.isinf(numeric).any():
            raise ValueError(f"列 {column!r} に無限大があります。")
        return converted.astype(float)

    def reload(self) -> None:
        """Excel 表を再読込し、出力列ごとに有効データのみで補間器を構築する。

        読み込みまたは構築が失敗した場合、既存の補間器は維持する。
        """
        df = self._read_table()
        names = self.input_columns + self.output_columns
        if len(df.columns) != len(set(df.columns)):
            raise ValueError("Excel のヘッダー名が重複しています。")
        missing_columns = [col for col in names if col not in df.columns]
        if missing_columns:
            raise ValueError(f"シート {self.sheet_name!r} に列がありません: {missing_columns}")
        if df.empty:
            raise ValueError("補間に必要なデータ行がありません。")

        numeric = pd.DataFrame({
            col: self._numeric_column(df[col], col) for col in names
        })
        # 独立変数が未定義なら座標を決定できないため全出力に対して除外。
        valid_inputs = numeric[self.input_columns].notna().all(axis=1)
        data = numeric.loc[valid_inputs, names].copy()
        if data.empty:
            raise ValueError("有効な入力座標が1行もありません。")

        x = data[self.input_columns].to_numpy(dtype=float)
        x_min = x.min(axis=0)
        x_max = x.max(axis=0)
        x_scale = x_max - x_min
        if np.any(x_scale == 0):
            cols = [c for c, scale in zip(self.input_columns, x_scale) if scale == 0]
            raise ValueError(f"全行で値が一定の入力列: {cols}。入力列から外してください。")
        if pd.DataFrame(x).duplicated().any():
            raise ValueError("補間の入力座標が重複しています。入力列の組を一意にしてください。")

        axes = None
        grid_index = None
        if self.method == "regular_grid":
            axes = tuple(np.unique(x[:, i]) for i in range(x.shape[1]))
            expected = int(np.prod([len(axis) for axis in axes]))
            if expected != len(x):
                raise ValueError(
                    f"regular_grid には入力座標の全格子点が必要です。"
                    f"実データ={len(x)}点、必要={expected}点。"
                )
            grid_index = pd.MultiIndex.from_product(axes, names=self.input_columns)

        # 補間器は出力列ごとに独立して構築する。
        # 一部出力が欠損していても、他の出力の学習点は減らさない。
        fitted = {}
        count_by_output = {}
        status_by_output = {}
        for name in self.output_columns:
            valid_y = data[name].notna().to_numpy()
            count = int(np.count_nonzero(valid_y))
            count_by_output[name] = count
            is_partial = count != len(x)
            x_good = x[valid_y]
            y_good = data.loc[valid_y, name].to_numpy(dtype=float)
            scaled_good = (x_good - x_min) / x_scale
            model = None
            model_mode = "missing"

            if self.method == "rbf":
                # thin_plate_spline の既定の多項式次数は1。入力点が
                # d+1 個に満たないと補間モデルは作れない。
                dim = len(self.input_columns)
                min_count = dim + 1 if self.rbf_kernel in {
                    "thin_plate_spline", "cubic", "quintic"
                } else 1
                if count >= min_count:
                    if self.rbf_kernel in {"thin_plate_spline", "cubic", "quintic"}:
                        affine = np.column_stack([np.ones(count), scaled_good])
                        if np.linalg.matrix_rank(affine) < dim + 1:
                            if not is_partial:
                                raise ValueError("RBF入力点が指定次元を張っていません。")
                        else:
                            model = RBFInterpolator(
                                scaled_good, y_good,
                                kernel=self.rbf_kernel,
                                epsilon=self.rbf_epsilon,
                                smoothing=self.smoothing,
                                neighbors=self.neighbors,
                            )
                            model_mode = "scaled"
                    else:
                        model = RBFInterpolator(
                            scaled_good, y_good,
                            kernel=self.rbf_kernel,
                            epsilon=self.rbf_epsilon,
                            smoothing=self.smoothing,
                            neighbors=self.neighbors,
                        )
                        model_mode = "scaled"
            elif self.method == "regular_grid":
                if not is_partial:
                    indexed = pd.Series(
                        y_good, index=pd.MultiIndex.from_frame(data[self.input_columns])
                    )
                    # 先に入力格子の全点数を検査しているので reindex で欠損しない。
                    values = indexed.reindex(grid_index).to_numpy(dtype=float)
                    if np.isnan(values).any():
                        raise ValueError("regular_grid で入力座標の全格子点が揃っていません。")
                    model = RegularGridInterpolator(
                        axes, values.reshape(tuple(len(axis) for axis in axes)),
                        method=self.grid_method,
                        bounds_error=self.check_bounds,
                        fill_value=np.nan if self.check_bounds else None,
                    )
                    model_mode = "grid"
                elif len(self.input_columns) == 1:
                    # 1次元なら有効点だけの格子として補間できる。
                    mode = "nearest" if self.grid_method == "nearest" else "linear"
                    min_count = 1 if mode == "nearest" else 2
                    if count >= min_count:
                        order = np.argsort(x_good[:, 0])
                        model = RegularGridInterpolator(
                            (x_good[order, 0],), y_good[order],
                            method=mode, bounds_error=False, fill_value=np.nan,
                        )
                        model_mode = "grid"
                elif count >= 1 and self.grid_method == "nearest":
                    model = NearestNDInterpolator(scaled_good, y_good)
                    model_mode = "scaled"
                elif count >= len(self.input_columns) + 1:
                    # RGI は穴の空いた多次元格子を扱えない。
                    # 欠損のある出力列だけ、散在点の区分線形に切り替える。
                    if np.linalg.matrix_rank(
                        np.column_stack([np.ones(count), scaled_good])
                    ) >= len(self.input_columns) + 1:
                        model = LinearNDInterpolator(scaled_good, y_good, fill_value=np.nan)
                        model_mode = "scaled"
            elif self.method == "linear_nd":
                dim = len(self.input_columns)
                if count >= dim + 1 and np.linalg.matrix_rank(
                    np.column_stack([np.ones(count), scaled_good])
                ) >= dim + 1:
                    model = LinearNDInterpolator(scaled_good, y_good, fill_value=np.nan)
                    model_mode = "scaled"
            elif self.method == "nearest_nd":
                if count >= 1:
                    if len(self.input_columns) == 1:
                        # SciPy の NearestNDInterpolator は1次元を受け付けない。
                        order = np.argsort(x_good[:, 0])
                        model = RegularGridInterpolator(
                            (x_good[order, 0],), y_good[order],
                            method="nearest", bounds_error=False,
                            fill_value=None,
                        )
                        model_mode = "grid"
                    else:
                        model = NearestNDInterpolator(scaled_good, y_good)
                        model_mode = "scaled"
            else:  # pchip は1入力
                if count >= 2:
                    order = np.argsort(x_good[:, 0])
                    model = PchipInterpolator(
                        x_good[order, 0], y_good[order],
                        extrapolate=not self.check_bounds,
                    )
                    model_mode = "one_dim"

            if model is None and not is_partial:
                raise ValueError(f"出力列 {name!r} の補間器を構築できませんでした。")
            fitted[name] = (model, model_mode, is_partial)
            status_by_output[name] = (
                "unavailable" if model is None else "partial" if is_partial else "complete"
            )

        self._models = fitted
        self.x_min = x_min
        self.x_max = x_max
        self.x_scale = x_scale
        self.n_points = len(x)
        self.n_points_by_output = count_by_output
        self.output_status = status_by_output
        self.n_dropped_input_rows = int((~valid_inputs).sum())

    def predict(self, **coordinates: Any) -> dict[str, float | np.ndarray]:
        """入力座標から出力列を補間する（スカラー・配列・broadcast 対応）。

        入力のいずれかに np.nan / +/-np.inf / None が含まれる評価点は、
        補間器に渡さず、その評価点に対応する **全出力列** を np.nan にする。
        配列の他の評価点は通常どおり補間する。

        列の不足・余分な列・配列形状の不一致は従来どおりエラー。
        check_bounds=True の場合、有限な入力の範囲外も従来どおりエラー。
        """
        missing = [c for c in self.input_columns if c not in coordinates]
        unknown = [c for c in coordinates if c not in self.input_columns]
        if missing or unknown:
            raise ValueError(f"入力列の不足: {missing}、未定義の列: {unknown}")
        try:
            arrays = np.broadcast_arrays(
                *(np.asarray(coordinates[c], dtype=float) for c in self.input_columns)
            )
        except ValueError as exc:
            raise ValueError("入力座標の配列形状をブロードキャストできません。") from exc

        shape = arrays[0].shape
        points = np.column_stack([a.reshape(-1) for a in arrays])
        valid_input = np.isfinite(points).all(axis=1)
        valid_points = points[valid_input]

        # NaN や inf は範囲判定の対象外。有限な評価点だけ従来どおり判定する。
        if self.check_bounds and len(valid_points):
            tolerance = 1e-12 * np.maximum(self.x_scale, 1.0)
            if np.any(
                (valid_points < self.x_min - tolerance)
                | (valid_points > self.x_max + tolerance)
            ):
                raise ValueError("評価点が各入力列の最小～最大範囲を外れています。")

        # 空配列／全点が非有限の場合も、モデルを呼ばず、形状を保持して返す。
        if not len(valid_points):
            return {
                name: float("nan") if shape == () else np.full(shape, np.nan)
                for name in self.output_columns
            }

        results = {}
        for name in self.output_columns:
            model, mode, has_missing = self._models[name]
            values = np.full(len(points), np.nan, dtype=float)
            if model is not None:
                if mode == "scaled":
                    evaluation_points = (valid_points - self.x_min) / self.x_scale
                elif mode == "one_dim":
                    evaluation_points = valid_points[:, 0]
                else:
                    evaluation_points = valid_points

                evaluated = np.asarray(model(evaluation_points), dtype=float).reshape(-1)
                if len(evaluated) != len(valid_points):
                    raise RuntimeError(f"出力列 {name!r} の補間値の個数が一致しません。")
                if np.isinf(evaluated).any() or (np.isnan(evaluated).any() and not has_missing):
                    raise ValueError(
                        f"出力列 {name!r} の補間結果が非有限値です。"
                        "散在点の凸包外への評価やデータ配置を確認してください。"
                    )
                values[valid_input] = evaluated
            results[name] = float(values[0]) if shape == () else values.reshape(shape)
        return results

    __call__ = predict


class ExcelRBFInterpolator(ExcelInterpolator):
    """従来の ExcelRBFInterpolator 名を維持する互換クラス。"""

    def __init__(self, *args: Any, **kwargs: Any) -> None:
        method = kwargs.pop("method", "rbf")
        if method != "rbf":
            raise ValueError("ExcelRBFInterpolator では method='rbf' のみ使用できます。")
        super().__init__(*args, method="rbf", **kwargs)
