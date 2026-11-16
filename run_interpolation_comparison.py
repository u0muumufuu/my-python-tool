"""補間方式・生値をPlotly HTMLで比較する独立実行ファイル（main.pyは不使用）。

Excel VBA:
    RunPython "import run_interpolation_comparison; run_interpolation_comparison.run()"

Excelを開かない実行:
    python src/run_interpolation_comparison.py --excel AircraftPerformance.xlsx

編集する場所は下の REPORTS です。横並び個数 ncols と縦個数 nrows も指定できます。
"""

from __future__ import annotations

import argparse
from pathlib import Path
from typing import Any, Mapping, Sequence

import pandas as pd
from openpyxl.utils.cell import column_index_from_string, coordinate_from_string

from utils.interpolation import ExcelInterpolator
from utils.plotly_charts import InterpolationPlotSpec, export_interpolation_report_html


# ---------------------------------------------------------------------------
# ユーザー設定。各グラフを複数方式で比較する。表示範囲Noneは実データの範囲。
# 線色=系列値、線種=補間方式。Plotly.jsをHTML内に埋め込む。
# ---------------------------------------------------------------------------
REPORTS = [
    {
        "name": "aerodynamics",
        "sheet_name": "10_空力DB",
        "start_cell": "A1",
        "input_columns": ["alpha", "Mach"],
        "output_columns": ["CL", "CD", "Cm"],
        "methods": {
            "RegularGrid (linear)": {"method": "regular_grid", "grid_method": "linear"},
            "RBF": {"method": "rbf"},
            "LinearND": {"method": "linear_nd"},
        },
        "ncols": 2,  # 横に2個
        "nrows": 2,  # 縦に2個。3つのグラフなので右下は空白
        "plots": [
            InterpolationPlotSpec(
                title="CL - alpha (Mach別)",
                x_column="alpha", y_column="CL",
                series_columns=["Mach"], series_values=None,
                x_range=None, x_samples=201,
            ),
            InterpolationPlotSpec(
                title="CD - alpha (Mach別)",
                x_column="alpha", y_column="CD",
                series_columns=["Mach"], series_values=None,
                x_range=None, x_samples=201,
            ),
            InterpolationPlotSpec(
                title="Cm - alpha (Mach別)",
                x_column="alpha", y_column="Cm",
                series_columns=["Mach"], series_values=None,
                x_range=None, x_samples=201,
            ),
        ],
        "output_html": "aerodynamics_interpolation_comparison.html",
    },
    {
        "name": "propulsion",
        "sheet_name": "11_推進DB",
        "start_cell": "A1",
        "input_columns": ["V", "Altitude"],
        "output_columns": ["Thrust"],
        "methods": {
            "RegularGrid (linear)": {"method": "regular_grid", "grid_method": "linear"},
            "RBF": {"method": "rbf"},
            "LinearND": {"method": "linear_nd"},
        },
        "ncols": 1,
        "nrows": 1,
        "plots": [
            InterpolationPlotSpec(
                title="Thrust - Altitude (V別)",
                x_column="Altitude", y_column="Thrust",
                series_columns=["V"], series_values=None,
                x_range=None, x_samples=201,
            ),
        ],
        "output_html": "propulsion_interpolation_comparison.html",
    },
]


def _read_raw_table(
    excel_path: Path, sheet_name: str, start_cell: str, *, workbook=None
) -> pd.DataFrame:
    """保存済みxlsx/xlsmまたはxlwingsの開いているExcelから元データを読む。"""
    if workbook is not None:
        data = workbook.sheets[sheet_name].range(start_cell).options(
            pd.DataFrame, index=False, expand="table"
        ).value
    else:
        col_letters, row = coordinate_from_string(start_cell)
        start_col = column_index_from_string(col_letters)
        data = pd.read_excel(excel_path, sheet_name=sheet_name, header=row - 1)
        if start_col > 1:
            data = data.iloc[:, start_col - 1 :]
    if not isinstance(data, pd.DataFrame) or data.empty:
        raise ValueError(f"シート {sheet_name!r} の元データ表が空です")
    return data


def run(
    *,
    excel_path: str | Path | None = None,
    workbook=None,
    reports: Sequence[Mapping[str, Any]] | None = None,
    output_dir: str | Path | None = None,
) -> dict[str, Path]:
    """Excelの生値と各補間方式を比較し、レポート毎にHTMLを1つ生成する。

    excel_pathが省略されworkbookもない場合はxlwings.Book.caller()を使用。
    Pythonからファイル指定する場合はExcelを開く必要がない。
    返り値: {レポート名: HTMLファイルのPath}。
    """
    if workbook is None and excel_path is None:
        import xlwings as xw
        workbook = xw.Book.caller()
    if excel_path is None:
        excel_path = workbook.fullname
    path = Path(excel_path).expanduser().resolve()
    target = Path(output_dir).expanduser().resolve() if output_dir is not None else path.parent / "reports"
    target.mkdir(parents=True, exist_ok=True)

    selections = REPORTS if reports is None else list(reports)
    results: dict[str, Path] = {}
    for config in selections:
        name = str(config["name"])
        sheet_name = str(config["sheet_name"])
        start_cell = str(config.get("start_cell", "A1"))
        raw = _read_raw_table(path, sheet_name, start_cell, workbook=workbook)
        method_configs = config["methods"]
        if not method_configs:
            raise ValueError(f"{name}: 比較対象のmethodsが空です")
        models = {}
        for label, options in method_configs.items():
            options = dict(options)
            models[str(label)] = ExcelInterpolator(
                excel_path=path,
                sheet_name=sheet_name,
                start_cell=start_cell,
                input_columns=list(config["input_columns"]),
                output_columns=list(config["output_columns"]),
                workbook=workbook,
                **options,
            )
        result = export_interpolation_report_html(
            models,
            raw,
            plot_specs=list(config["plots"]),
            output_html=target / str(config["output_html"]),
            report_title=f"{name} — raw data vs interpolation methods",
            show_range_slider=True,
            ncols=int(config.get("ncols", 1)),
            nrows=(None if config.get("nrows") is None else int(config["nrows"])),
        )
        results[name] = result
    return results


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Excel生データと補間方式のPlotly比較HTML出力")
    parser.add_argument("--excel", required=True, help="Excelファイル(.xlsx/.xlsm)")
    parser.add_argument("--output-dir", default=None, help="HTML出力先（省略時はExcel横のreports/）")
    args = parser.parse_args()
    for label, out in run(excel_path=args.excel, output_dir=args.output_dir).items():
        print(f"{label}: {out}")
