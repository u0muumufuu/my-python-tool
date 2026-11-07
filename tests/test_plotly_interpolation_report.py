import sys
from pathlib import Path

import numpy as np
import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from utils.interpolation import ExcelInterpolator
from utils.plotly_charts import (
    InterpolationPlotSpec,
    build_interpolation_report_figure,
    export_interpolation_report_html,
)


def _make_aero_excel(path: Path) -> pd.DataFrame:
    rows = []
    for mach in [0.2, 0.4]:
        for alpha in [0.0, 5.0, 10.0]:
            rows.append({
                "alpha": alpha,
                "Mach": mach,
                "CL": 0.08 * alpha + 0.2 * mach,
                "CD": 0.02 + 0.001 * alpha**2 + 0.01 * mach,
                "Cm": 0.1 - 0.01 * alpha - 0.02 * mach,
            })
    df = pd.DataFrame(rows)
    df.to_excel(path, index=False, sheet_name="10_空力DB")
    return df


def test_build_interpolation_report_with_multiple_subplots_and_rangesliders(tmp_path):
    excel = tmp_path / "aero.xlsx"
    raw = _make_aero_excel(excel)
    interp = ExcelInterpolator(
        excel_path=excel,
        sheet_name="10_空力DB",
        input_columns=["alpha", "Mach"],
        output_columns=["CL", "CD", "Cm"],
        method="regular_grid",
        check_bounds=False,
    )

    specs = [
        InterpolationPlotSpec(
            title="CL-alpha",
            x_column="alpha",
            y_column="CL",
            series_columns=["Mach"],
            series_values=[0.2, 0.4],
            x_range=(-5.0, 15.0),
            x_samples=41,
        ),
        InterpolationPlotSpec(
            title="CD-alpha",
            x_column="alpha",
            y_column="CD",
            series_columns=["Mach"],
            series_values=[0.2, 0.4],
            x_range=(-2.0, 12.0),
            x_samples=31,
        ),
    ]
    fig = build_interpolation_report_figure(
        interp, raw, plot_specs=specs, report_title="aero report"
    )
    # 2グラフ × 2系列 × (生値 + 補間)
    assert len(fig.data) == 8
    assert fig.layout.xaxis.rangeslider.visible is True
    assert fig.layout.xaxis2.rangeslider.visible is True
    # 外挿範囲を含むか（補間曲線の先頭x）
    first_fit = fig.data[1]
    assert np.isclose(first_fit.x[0], -5.0)


def test_export_interpolation_report_html_contains_titles(tmp_path):
    excel = tmp_path / "aero.xlsx"
    raw = _make_aero_excel(excel)
    interp = ExcelInterpolator(
        excel_path=excel,
        sheet_name="10_空力DB",
        input_columns=["alpha", "Mach"],
        output_columns=["CL", "CD", "Cm"],
        method="regular_grid",
        check_bounds=False,
    )
    specs = [
        InterpolationPlotSpec(
            title="CL-alpha report",
            x_column="alpha",
            y_column="CL",
            series_columns=["Mach"],
            series_values=[0.2, 0.4],
        )
    ]
    out = tmp_path / "reports" / "aero.html"
    path = export_interpolation_report_html(
        interp, raw, plot_specs=specs, output_html=out,
        report_title="Aerodynamics interpolation report"
    )
    assert path.exists()
    html = path.read_text(encoding="utf-8")
    assert "Aerodynamics interpolation report" in html
    assert "CL-alpha report" in html
    assert "Plotly.newPlot" in html


def test_series_values_can_be_auto_detected(tmp_path):
    excel = tmp_path / "aero.xlsx"
    raw = _make_aero_excel(excel)
    interp = ExcelInterpolator(
        excel_path=excel,
        sheet_name="10_空力DB",
        input_columns=["alpha", "Mach"],
        output_columns=["CL", "CD", "Cm"],
        method="regular_grid",
        check_bounds=False,
    )
    fig = build_interpolation_report_figure(
        interp, raw,
        plot_specs=[InterpolationPlotSpec(
            title="auto series",
            x_column="alpha",
            y_column="CL",
            series_columns=["Mach"],
            x_range=(0.0, 10.0),
        )],
    )
    # Mach 0.2, 0.4 の2系列で生値+補間の合計4トレース
    assert len(fig.data) == 4
