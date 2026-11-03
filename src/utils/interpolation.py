"""航空機の性能・空力データ解析で使う共通ユーティリティ。

現時点では Excel 表に基づく補間器のみ提供する。
今後、求根・収束計算などの独立した関数／クラスを追加する想定。

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
    regular_grid は元の軸を使い、全組み合わせの存在を確認する。
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

    def reload(self) -> None:
        """Excel の数値表を再読込し、補間器を再構築する。"""
        df = self._read_table()
        names = self.input_columns + self.output_columns
        if len(df.columns) != len(set(df.columns)):
            raise ValueError("Excel のヘッダー名が重複しています。")
        missing = [col for col in names if col not in df.columns]
        if missing:
            raise ValueError(f"シート {self.sheet_name!r} に列がありません: {missing}")
        if df.empty:
            raise ValueError("補間に必要なデータ行がありません。")

        data = df.loc[:, names].apply(pd.to_numeric, errors="raise")
        arr = data.to_numpy(dtype=float)
        if not np.isfinite(arr).all():
            raise ValueError("対象列に空欄、NaN、無限大があります。")
        x = data[self.input_columns].to_numpy(dtype=float)
        y = data[self.output_columns].to_numpy(dtype=float)
        x_min = x.min(axis=0)
        x_max = x.max(axis=0)
        x_scale = x_max - x_min
        if np.any(x_scale == 0):
            cols = [c for c, s in zip(self.input_columns, x_scale) if s == 0]
            raise ValueError(f"全行で値が一定の入力列: {cols}。入力列から外してください。")
        if pd.DataFrame(x).duplicated().any():
            raise ValueError("補間の入力座標が重複しています。入力列の組を一意にしてください。")

        # 一時的な変数で構築し、失敗した場合は古い補間器を破壊しない。
        if self.method == "rbf":
            scaled_x = (x - x_min) / x_scale
            # 標準カーネルにはアフィン項が含まれるため、通常は各次元に独立な点が必要。
            if self.rbf_kernel == "thin_plate_spline":
                affine = np.column_stack([np.ones(len(x)), scaled_x])
                if np.linalg.matrix_rank(affine) < len(self.input_columns) + 1:
                    raise ValueError("RBF入力点が指定次元を張っていません。点の配置を確認してください。")
            interpolator = RBFInterpolator(
                scaled_x, y,
                kernel=self.rbf_kernel,
                epsilon=self.rbf_epsilon,
                smoothing=self.smoothing,
                neighbors=self.neighbors,
            )
        elif self.method == "regular_grid":
            # 直交格子 (必ずしも等間隔でなくてよい) の全組み合わせをチェック。
            axes = tuple(np.unique(x[:, i]) for i in range(x.shape[1]))
            expected = int(np.prod([len(axis) for axis in axes]))
            if expected != len(x):
                raise ValueError(
                    f"regular_grid には全格子点が必要です。"
                    f"実データ={len(x)}点、必要={expected}点。"
                )
            point_index = pd.MultiIndex.from_frame(data[self.input_columns])
            product = pd.MultiIndex.from_product(axes, names=self.input_columns)
            ordered = pd.DataFrame(y, index=point_index, columns=self.output_columns).reindex(product)
            if ordered.isna().any().any():
                raise ValueError("regular_grid で欠けている格子点があります。")
            grid_shape = tuple(len(axis) for axis in axes) + (len(self.output_columns),)
            grid_values = ordered.to_numpy(dtype=float).reshape(grid_shape)
            interpolator = RegularGridInterpolator(
                axes, grid_values,
                method=self.grid_method,
                bounds_error=self.check_bounds,
                fill_value=np.nan if self.check_bounds else None,
            )
        elif self.method in {"linear_nd", "nearest_nd"}:
            scaled_x = (x - x_min) / x_scale
            if self.method == "linear_nd":
                interpolator = LinearNDInterpolator(scaled_x, y, fill_value=np.nan)
            else:
                interpolator = NearestNDInterpolator(scaled_x, y)
        else:  # pchip (1次元)
            order = np.argsort(x[:, 0])
            interpolator = PchipInterpolator(
                x[order, 0], y[order], axis=0, extrapolate=not self.check_bounds
            )

        self._interpolator = interpolator
        self.x_min = x_min
        self.x_max = x_max
        self.x_scale = x_scale
        self.n_points = len(x)

    def predict(self, **coordinates: Any) -> dict[str, float | np.ndarray]:
        """入力列名に対応する値から補間する。スカラー/配列の混合も可能。

        例: predict(alpha=5, Mach=0.2) -> {'CL': 0.5, 'CD': 0.04}
        リスト/ndarray を渡すとブロードキャストして同形状の配列を返す。
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
        if not np.isfinite(points).all():
            raise ValueError("評価点に NaN または無限大があります。")
        if self.check_bounds:
            tolerance = 1e-12 * np.maximum(self.x_scale, 1.0)
            if np.any((points < self.x_min - tolerance) | (points > self.x_max + tolerance)):
                raise ValueError("評価点が各入力列の最小～最大範囲を外れています。")
        if len(points) == 0:
            return {name: np.empty(shape) for name in self.output_columns}

        if self.method in {"rbf", "linear_nd", "nearest_nd"}:
            evaluation_points = (points - self.x_min) / self.x_scale
        elif self.method == "pchip":
            evaluation_points = points[:, 0]
        else:
            evaluation_points = points
        values = np.asarray(self._interpolator(evaluation_points), dtype=float)
        values = values.reshape(len(points), len(self.output_columns))
        if not np.isfinite(values).all():
            raise ValueError(
                "補間結果が非有限値です。散在点の凸包外への評価や欠損格子を確認してください。"
            )
        return {
            name: (float(values[0, i]) if shape == () else values[:, i].reshape(shape))
            for i, name in enumerate(self.output_columns)
        }

    __call__ = predict


class ExcelRBFInterpolator(ExcelInterpolator):
    """従来の ExcelRBFInterpolator 名を維持する互換クラス。"""

    def __init__(self, *args: Any, **kwargs: Any) -> None:
        method = kwargs.pop("method", "rbf")
        if method != "rbf":
            raise ValueError("ExcelRBFInterpolator では method='rbf' のみ使用できます。")
        super().__init__(*args, method="rbf", **kwargs)
