"""xlwings/Windows Excel のXY散布図を更新する共通ユーティリティ。

ポイント
--------
- 計算済みの Excel セルをグラフ系列に直接結び付ける（重複データ不要）。
- 行数や系列数が変わっても、既存グラフの位置・軸などを基本的に維持。
- 複数シートの点群（生データ）と補間曲線の重ね描きに対応。
- ExcelのネイティブCOM APIを使用するので **WindowsデスクトップExcel専用**。

Excel実機の検証は各PC上で実行してください。Linux/macOSではCOM操作できません。
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Literal, Sequence, Mapping

import pandas as pd
from openpyxl.utils.cell import coordinate_from_string, column_index_from_string, get_column_letter

from .excel_io import write_excel_result

PlotMode = Literal["lines", "markers", "lines+markers"]

# Excel XlChartType values: xlXYScatterLinesNoMarkers, xlXYScatter, xlXYScatterLines
_XY_TYPES = {"lines": 75, "markers": -4169, "lines+markers": 74}


@dataclass(frozen=True)
class ExcelSeriesSpec:
    """グラフ内の1つのXY系列（別シート・別行数でもよい）。

    Parameters
    ----------
    name: 凡例名
    sheet_name: 元データがあるシート名
    x_column, y_column: ヘッダー行に記された列名
    header_cell: ヘッダー行の左端。例 ``'A5'``
    row_count: データ行数。None なら ``expand('table')`` で自動検出。
               A列などに空白行がある場合は明示的に指定する。
    mode: ``'lines'`` / ``'markers'`` / ``'lines+markers'``
    """

    name: str
    sheet_name: str
    x_column: str
    y_column: str
    header_cell: str = "A1"
    row_count: int | None = None
    mode: PlotMode = "lines"


def _column_ranges(wb, spec: ExcelSeriesSpec):
    if spec.mode not in _XY_TYPES:
        raise ValueError(f"対応しない mode: {spec.mode!r}")
    sheet = wb.sheets[spec.sheet_name]
    col_letter, header_row = coordinate_from_string(spec.header_cell.upper())
    start_col = column_index_from_string(col_letter)
    header_values = sheet.range(spec.header_cell).expand("right").value
    if not isinstance(header_values, (list, tuple)):
        header_values = [header_values]
    headers = [str(x) if x is not None else "" for x in header_values]
    if len(headers) != len(set(headers)):
        raise ValueError(f"重複したヘッダーがあります: {spec.sheet_name}")
    for column in (spec.x_column, spec.y_column):
        if column not in headers:
            raise KeyError(f"{spec.sheet_name}!{spec.header_cell} に列 {column!r} がありません。列: {headers}")

    if spec.row_count is None:
        row_count = int(sheet.range(spec.header_cell).expand("table").rows.count) - 1
    else:
        row_count = spec.row_count
    if not isinstance(row_count, int) or row_count < 1:
        raise ValueError(f"系列 {spec.name!r}: データ行数は1以上にしてください: {row_count}")
    if header_row + row_count > 1048576:
        raise ValueError("Excelシートの最大行数を超えています")

    def make_range(name: str):
        c = get_column_letter(start_col + headers.index(name))
        return sheet.range(f"{c}{header_row + 1}:{c}{header_row + row_count}")

    return make_range(spec.x_column), make_range(spec.y_column)


def _native_chart(chart):
    """xlwings Chart.api のExcel COM Chartオブジェクトを得る。"""
    api = chart.api
    if hasattr(api, "SeriesCollection"):
        return api
    if hasattr(api, "Chart") and hasattr(api.Chart, "SeriesCollection"):
        return api.Chart
    raise RuntimeError("ネイティブExcelグラフAPIにアクセスできません。Windows版Excelが必要です")


def update_xy_chart(
    wb,
    *,
    chart_sheet: str,
    chart_name: str,
    series_specs: Sequence[ExcelSeriesSpec],
    create_if_missing: bool = True,
    title: str | None = None,
    position: tuple[float, float, float, float] = (650, 85, 620, 365),
):
    """Excelグラフの系列と参照範囲を現在のデータ行数に合わせて更新する。

    COM依存のため、Windowsデスクトップ版Excelからxlwings経由で実行する。
    グラフが既存なら再利用し、不要になった系列だけ削除する。
    ``position=(left, top, width, height)`` は新規作成時だけ使う（単位pt）。
    """
    specs = list(series_specs)
    if not specs:
        raise ValueError("series_specs を1件以上指定してください")
    if not chart_name:
        raise ValueError("chart_name は空文字列にできません")

    # 出力対象をまず検証。データ/列の不備で既存グラフを書き換えない。
    ranges = [_column_ranges(wb, spec) for spec in specs]

    dest = wb.sheets[chart_sheet]
    chart_names = [chart.name for chart in dest.charts]
    is_new = chart_name not in chart_names
    if is_new:
        if not create_if_missing:
            raise KeyError(f"グラフ {chart_name!r} は {chart_sheet!r} にありません")
        left, top, width, height = position
        chart = dest.charts.add(left=left, top=top, width=width, height=height)
        chart.name = chart_name
    else:
        chart = dest.charts[chart_name]

    native = _native_chart(chart)
    if is_new:
        native.ChartType = 75  # XY（直線、マーカーなし）。Excelの平滑線ではない。
    if title is not None:
        native.HasTitle = True
        native.ChartTitle.Text = title

    collection = native.SeriesCollection()
    target_count = len(specs)
    # 再実行で曲線が減った場合、余分な系列を削除する。
    while collection.Count > target_count:
        collection(collection.Count).Delete()

    for index, (spec, (x_range, y_range)) in enumerate(zip(specs, ranges), start=1):
        is_new_series = index > collection.Count
        ser = collection.NewSeries() if is_new_series else collection(index)
        ser.Name = spec.name
        # .api で参照範囲を渡し、画像や値のコピーではなくセルとリンクする。
        ser.XValues = x_range.api
        ser.Values = y_range.api
        # 軸やグラフ配置は維持し、系列の表示形式だけ明示されたモードに統一。
        ser.ChartType = _XY_TYPES[spec.mode]
        if spec.mode in ("markers", "lines+markers"):
            ser.MarkerStyle = 8  # 丸
        else:
            ser.MarkerStyle = -4142  # none
    return chart


def update_chart_from_table(
    wb,
    *,
    chart_sheet: str,
    chart_name: str,
    data_sheet: str,
    x_column: str,
    y_columns: Sequence[str],
    header_cell: str = "A1",
    row_count: int | None = None,
    create_if_missing: bool = True,
    title: str | None = None,
):
    """同じ表にある複数のY列を一括グラフ化する便利ラッパー。"""
    names = list(y_columns)
    if not names:
        raise ValueError("y_columns が空です")
    specs = [
        ExcelSeriesSpec(
            name=name, sheet_name=data_sheet,
            x_column=x_column, y_column=name,
            header_cell=header_cell, row_count=row_count, mode="lines",
        )
        for name in names
    ]
    return update_xy_chart(
        wb,
        chart_sheet=chart_sheet, chart_name=chart_name,
        series_specs=specs,
        create_if_missing=create_if_missing, title=title,
    )


def write_chart_table(sheet, df: pd.DataFrame, *, start_cell: str = "A1") -> None:
    """補間曲線などの **専用出力表** を更新する（前回の残存値も消す）。

    ``start_cell`` 直下に見出し・データが連続する専用領域にのみ使用すること。
    共存する別の数式や入力セルを消さないよう配置する。
    """
    if not isinstance(df, pd.DataFrame) or df.empty or df.shape[1] == 0:
        raise ValueError("空でないDataFrameを指定してください")
    if df.columns.has_duplicates:
        raise ValueError("DataFrameの列名が重複しています")
    sheet.range(start_cell).expand("table").clear_contents()
    sheet.range(start_cell).value = [list(map(str, df.columns))]
    # np.nanはExcelの#N/A数式に変換する既存関数を共用。
    column, row = coordinate_from_string(start_cell.upper())
    write_excel_result(sheet, f"{column}{row + 1}", df.to_numpy(dtype=object))


def update_contour_chart_from_dataframe(
    wb,
    data: pd.DataFrame,
    *,
    fixed_value: float = 1.0,
    levels: Sequence[float],
    output_data_sheet: str = "32_Ps等値線データ",
    chart_sheet: str = "40_グラフ",
    chart_name: str = "Ps_contour_chart",
    fixed_column: str = "G",
    fixed: Mapping[str, float] | None = None,
    fixed_columns: Sequence[str] | None = None,
    fixed_values: Sequence[float] | None = None,
    level_column: str | None = None,
    x_column: str = "Mach",
    y_column: str = "Altitude",
    z_column: str = "Ps",
    fixed_mode: str = "exact",
    resample: str = "none",
    title: str | None = None,
    **grid_options,
) -> pd.DataFrame:
    """長形式Ps表から等値線を抽出し、ExcelネイティブXY図を作成/更新する。

    fixed_columns=['G'], fixed_values=[1.0] で固定列名を指定できる。
    level_column='Ps' と levels=[0, 10, 20] は別々の指定。
    fixed={'G': 1.0, 'Weight': 120.0} の辞書指定も可能。
    従来の fixed_column='G', fixed_value=1.0, z_column='Ps' も使用可能。
    元データは変更せず、*専用の* output_data_sheet だけを更新する。
    配列長や等値線の数・分離枝の数が変わっても系列を作成/削除する。
    戻り値は縦長の等値線座標表（検証・CSV保存などに再利用可）。

    1系列/分離枝につき2列を使うため、極端に多数の枝のときは
    Excelの最大列数を超えることがある。その場合はPlotly HTMLを推奨。
    """
    from .contour_maps import (
        prepare_contour_grid, extract_contour_lines, contours_to_wide_table
    )
    grid = prepare_contour_grid(
        data,
        fixed_column=fixed_column, fixed_value=fixed_value,
        fixed=fixed, fixed_columns=fixed_columns, fixed_values=fixed_values,
        level_column=level_column,
        x_column=x_column, y_column=y_column, z_column=z_column,
        fixed_mode=fixed_mode, resample=resample, **grid_options,
    )
    lines = extract_contour_lines(grid, levels=levels)
    wide, definitions = contours_to_wide_table(
        lines, x_column=grid.x_column, y_column=grid.y_column, z_column=grid.z_column,
    )
    if not definitions:
        raise ValueError("指定した等値線は計算領域内に存在しません。既存グラフは変更していません")
    if len(wide.columns) > 16384:
        raise ValueError("等値線の枝数が多すぎます。Plotly HTMLを使用してください")

    # 以前の可視化表を消去。専用シートに値や別の表を置かないこと。
    sheet = wb.sheets[output_data_sheet]
    # 異なる等値線の長さを揃えるために生じたNaNは、補間の失敗値とは
    # 異なり単なる末尾の余白。空欄として出力し、Excel COMの大量の
    # セル単位 =NA() 書き込みを回避する（分離枝同士は別系列）。
    wide = wide.astype(object).where(pd.notna(wide), None)
    write_chart_table(sheet, wide, start_cell="A1")
    specs = [
        ExcelSeriesSpec(name=legend, sheet_name=output_data_sheet,
                        x_column=xname, y_column=yname,
                        header_cell="A1", row_count=len(wide), mode="lines")
        for legend, xname, yname in definitions
    ]
    update_xy_chart(
        wb, chart_sheet=chart_sheet, chart_name=chart_name,
        series_specs=specs,
        title=title or (
            f"{grid.z_column} contours: "
            + ", ".join(f"{name}={value:g}" for name, value in grid.fixed_conditions)
        ),
    )
    return lines
