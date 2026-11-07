"""新しい列名指定APIのテスト（v8）。"""
from pathlib import Path
import sys

import numpy as np
import pandas as pd
import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))
from utils.contour_maps import prepare_contour_grid, extract_contour_lines
from utils.excel_charts import update_contour_chart_from_dataframe
from utils.plotly_charts import export_isolines_html


def sample():
    # 任意の行順、異なる重量・Gを混在させる。
    rows = [
        (alt, mach, g, weight, 100 * mach + alt / 100 + g + weight / 10)
        for weight in (120.0, 140.0)
        for g in (1.0, 2.0)
        for alt in (0.0, 1000.0, 2000.0)
        for mach in (0.1, 0.2, 0.3, 0.4)
    ]
    return pd.DataFrame(rows, columns=["Altitude", "Mach", "G", "Weight", "Ps"]).sample(frac=1, random_state=3)


def test_one_fixed_list_api():
    # 重量が2種類ある場合はG以外の条件も指定しないと重複扱い。
    grid = prepare_contour_grid(
        sample(), fixed_columns=["G", "Weight"], fixed_values=[1.0, 120.0],
        level_column="Ps", x_column="Mach", y_column="Altitude",
    )
    assert grid.z.shape == (3, 4)
    assert np.isclose(grid.z[0, 0], 23.0)
    lines = extract_contour_lines(grid, levels=[35.0])
    assert not lines.empty
    assert set(lines["G"]) == {1.0}
    assert set(lines["Weight"]) == {120.0}
    assert "Ps_level" in lines.columns
    assert np.allclose(100 * lines["Mach"] + lines["Altitude"] / 100 + 13, 35)


def test_fixed_dict_and_axes_swapped():
    grid = prepare_contour_grid(
        sample(), fixed={"Weight": 140.0, "G": 2.0},
        level_column="Ps", x_column="Altitude", y_column="Mach",
    )
    assert grid.z.shape == (4, 3)
    assert grid.fixed_conditions == (("Weight", 140.0), ("G", 2.0))
    lines = extract_contour_lines(grid, levels=[40.0])
    assert {"Altitude", "Mach", "Ps_level", "G", "Weight"} <= set(lines.columns)
    assert len(lines) > 0


def test_generic_level_column():
    df = sample().rename(columns={"Ps": "ExcessPower"})
    grid = prepare_contour_grid(df, fixed={"G": 1.0, "Weight": 120.0},
                                level_column="ExcessPower")
    lines = extract_contour_lines(grid, levels=[35.0])
    assert "ExcessPower_level" in lines.columns
    assert "Ps_level" not in lines.columns


def test_mixed_conditions_must_all_be_specified():
    with pytest.raises(ValueError, match="重複"):
        prepare_contour_grid(sample(), fixed={"G": 1.0}, level_column="Ps")


@pytest.mark.parametrize("kwargs,match", [
    ({"fixed_columns": ["G"]}, "両方"),
    ({"fixed_columns": ["G", "Weight"], "fixed_values": [1.0]}, "zip()"),
    ({"fixed": {"G": 1.0}, "fixed_columns": ["Weight"], "fixed_values": [120.0]}, "併用"),
    ({"fixed": {"Mach": 0.2}, "x_column": "Mach"}, "異なる"),
    ({"fixed": {"G": np.nan}}, "有限"),
    ({"fixed": {"G": 1.0, "Weight": 120.0}, "fixed_mode": "linear"}, "複数"),
])
def test_invalid_fixed_parameters(kwargs, match):
    with pytest.raises((ValueError, TypeError), match=match):
        prepare_contour_grid(sample(), **kwargs)


def test_excel_export_with_new_columns(monkeypatch):
    from utils import excel_charts
    calls = {}

    class Book:
        sheets = {"coordinates": object()}

    def mock_write(sheet, df, start_cell):
        calls["table"] = df

    def mock_update(wb, **kwargs):
        calls["specs"] = kwargs["series_specs"]
        calls["title"] = kwargs["title"]

    monkeypatch.setattr(excel_charts, "write_chart_table", mock_write)
    monkeypatch.setattr(excel_charts, "update_xy_chart", mock_update)
    lines = update_contour_chart_from_dataframe(
        Book(), sample(),
        fixed_columns=["G", "Weight"], fixed_values=[1.0, 120.0],
        level_column="Ps", levels=[25, 35, 45],
        x_column="Mach", y_column="Altitude",
        output_data_sheet="coordinates",
    )
    assert not lines.empty
    assert "G=1" in calls["title"] and "Weight=120" in calls["title"]
    assert len(calls["specs"]) >= 2
    assert len(calls["table"].columns) == 2 * len(calls["specs"])


def test_plotly_export_with_new_columns(tmp_path):
    output = export_isolines_html(
        sample(), fixed={"G": 1.0, "Weight": 120.0},
        level_column="Ps", levels=[25, 35],
        x_column="Mach", y_column="Altitude",
        output_html=tmp_path / "plot.html", show_heatmap=True,
    )
    text = output.read_text(encoding="utf-8")
    assert "Ps = 25" in text
    assert "Weight=120" in text
    assert "plotly" in text.lower()


def test_old_api_backward_compatibility():
    df = sample().loc[lambda d: d["Weight"].eq(120)]
    grid = prepare_contour_grid(df, fixed_column="G", fixed_value=1.0, z_column="Ps")
    assert grid.z.shape == (3, 4)
    assert grid.fixed_column == "G"
    assert grid.fixed_value == 1.0
