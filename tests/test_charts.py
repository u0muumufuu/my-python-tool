"""Excel COM のインターフェースは模擬し、Plotly HTMLは実際に生成して検証。"""

import sys
from pathlib import Path

import numpy as np
import pandas as pd
import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from utils.excel_charts import (
    ExcelSeriesSpec, update_xy_chart, update_chart_from_table, write_chart_table,
)
from utils.plotly_charts import (
    build_xy_figure, export_xy_html, export_interpolation_html, export_contour_html,
)


class MockRows:
    def __init__(self, count): self.count = count


class MockRange:
    def __init__(self, sheet, addr, value=None):
        self.sheet, self.addr = sheet, addr
        self.value = value
        self.api = f"{sheet.name}!{addr}"
    def expand(self, direction):
        if direction == "right":
            return MockRange(self.sheet, self.addr, self.sheet.headers)
        if direction == "table":
            result = MockRange(self.sheet, self.addr)
            result.rows = MockRows(self.sheet.data_rows + 1)
            return result
        raise AssertionError(direction)
    def clear_contents(self):
        self.sheet.cleared.append(self.addr)
    def offset(self, rows, cols):
        from openpyxl.utils.cell import coordinate_from_string, column_index_from_string, get_column_letter
        col, row = coordinate_from_string(self.addr)
        return self.sheet.range(f"{get_column_letter(column_index_from_string(col) + cols)}{row + rows}")
    @property
    def formula(self): return self.value
    @formula.setter
    def formula(self, val): self.value = val


class MockSheet:
    def __init__(self, name, headers=None, data_rows=0):
        self.name = name
        self.headers = headers or []
        self.data_rows = data_rows
        self.charts = MockCharts()
        self.cleared = []
        self.ranges = {}
    def range(self, addr):
        if addr not in self.ranges:
            self.ranges[addr] = MockRange(self, addr)
        return self.ranges[addr]


class MockSeries:
    def __init__(self):
        self.Name = ""
        self.Values = self.XValues = None
        self.MarkerStyle = self.ChartType = None
        self.deleted = False
    def Delete(self):
        self.deleted = True


class MockSeriesCollection:
    def __init__(self): self.items = []
    @property
    def Count(self): return len(self.items)
    def NewSeries(self):
        obj = MockSeries()
        self.items.append(obj)
        return obj
    def __call__(self, n): return self.items[n-1]
    def delete_last(self): self.items.pop()


class MockChartNative:
    def __init__(self):
        self.series_collection = MockSeriesCollection()
        self.ChartType = None
        self.HasTitle = False
        self.ChartTitle = type("Title", (), {"Text": None})()
    def SeriesCollection(self):
        return self.series_collection


# emulate COM deletion: Delete() also removes it from the native collection
_original_del = MockSeries.Delete


def _remove_parent(self):
    self.deleted = True
    self.parent.items.remove(self)


MockSeries.Delete = _remove_parent
_original_new = MockSeriesCollection.NewSeries


def _new_parent(self):
    obj = _original_new(self)
    obj.parent = self
    return obj


MockSeriesCollection.NewSeries = _new_parent


class MockChart:
    def __init__(self):
        self.name = ""
        self.api = MockChartNative()


class MockCharts:
    def __init__(self):
        self.charts = {}
        self.add_count = 0
    def __iter__(self): return iter(self.charts.values())
    def __getitem__(self, key):
        for chart in self.charts.values():
            if chart.name == key:
                return chart
        raise KeyError(key)
    def add(self, **kwargs):
        chart = MockChart()
        self.add_count += 1
        self.charts[f"temporary_{self.add_count}"] = chart
        # name is assigned after add; simple lookup resolves from names below
        return chart
    # __getitem__ by chart.name; dictionary remains keyed by temp key for new charts


class MockBook:
    def __init__(self):
        self.sheets = {
            "30_TV": MockSheet("30_TV", ["V", "T_required", "T_available"], 100),
            "40_グラフ": MockSheet("40_グラフ"),
            "10_空力DB": MockSheet("10_空力DB", ["alpha", "CL", "CD"], 4),
            "41_補間データ": MockSheet("41_補間データ", ["alpha", "CL"], 201),
        }


def _find_chart(book, chart_name):
    for chart in book.sheets["40_グラフ"].charts:
        if chart.name == chart_name:
            return chart
    raise AssertionError(chart_name)


def test_table_chart_new_and_range_dynamic():
    wb = MockBook()
    update_chart_from_table(wb, chart_sheet="40_グラフ", chart_name="TV_chart",
        data_sheet="30_TV", header_cell="A5", x_column="V",
        y_columns=["T_required", "T_available"], row_count=100,
        title="T-V")
    ch = _find_chart(wb, "TV_chart")
    assert ch.api.ChartTitle.Text == "T-V"
    assert ch.api.series_collection.Count == 2
    assert ch.api.series_collection(1).XValues == "30_TV!A6:A105"
    assert ch.api.series_collection(2).Values == "30_TV!C6:C105"
    assert wb.sheets["40_グラフ"].charts.add_count == 1

    wb.sheets["30_TV"].data_rows = 23
    update_chart_from_table(wb, chart_sheet="40_グラフ", chart_name="TV_chart",
        data_sheet="30_TV", header_cell="A5", x_column="V",
        y_columns=["T_required"], row_count=None)
    assert wb.sheets["40_グラフ"].charts.add_count == 1
    assert ch.api.series_collection.Count == 1
    assert ch.api.series_collection(1).Values == "30_TV!B6:B28"


def test_multi_sheet_raw_and_fit_with_independent_row_counts():
    wb = MockBook()
    update_xy_chart(wb, chart_sheet="40_グラフ", chart_name="interp", series_specs=[
        ExcelSeriesSpec("CL（生値）", "10_空力DB", "alpha", "CL", row_count=4, mode="markers"),
        ExcelSeriesSpec("CL（補間）", "41_補間データ", "alpha", "CL", row_count=201, mode="lines"),
    ])
    ch = _find_chart(wb, "interp")
    assert ch.api.series_collection(1).ChartType == -4169
    assert ch.api.series_collection(1).MarkerStyle == 8
    assert ch.api.series_collection(2).ChartType == 75
    assert ch.api.series_collection(2).XValues == "41_補間データ!A2:A202"


def test_fails_before_modifying_chart_on_bad_column():
    wb = MockBook()
    with pytest.raises(KeyError, match="列"):
        update_chart_from_table(wb, chart_sheet="40_グラフ", chart_name="bad",
            data_sheet="30_TV", x_column="missing", y_columns=["T_required"])
    assert wb.sheets["40_グラフ"].charts.add_count == 0


@pytest.mark.parametrize("row_count", [0, -1])
def test_reject_invalid_data_rows(row_count):
    wb = MockBook()
    with pytest.raises(ValueError, match="行数"):
        update_xy_chart(wb, chart_sheet="40_グラフ", chart_name="bad", series_specs=[
            ExcelSeriesSpec("x", "30_TV", "V", "T_required", row_count=row_count)
        ])


def test_reject_unknown_mode():
    wb = MockBook()
    with pytest.raises(ValueError, match="mode"):
        update_xy_chart(wb, chart_sheet="40_グラフ", chart_name="bad", series_specs=[
            ExcelSeriesSpec("x", "30_TV", "V", "T_required", mode="spline")
        ])


def test_no_create_when_disabled():
    wb = MockBook()
    with pytest.raises(KeyError, match="ありません"):
        update_xy_chart(wb, chart_sheet="40_グラフ", chart_name="bad", create_if_missing=False,
            series_specs=[ExcelSeriesSpec("x", "30_TV", "V", "T_required")])


def test_plotly_xy_html_offline_and_nan(tmp_path):
    df = pd.DataFrame({"V": [20, 30, 40, 50], "T_req": [120, np.nan, 90, 100],
                       "T_avail": [140, 130, "#N/A", 110]})
    output = tmp_path / "plot" / "tv.html"
    path = export_xy_html(df, x_column="V", y_columns=["T_req", "T_avail"],
                          output_html=output, title="T-V")
    assert path.exists()
    content = path.read_text(encoding="utf-8")
    assert "Plotly.newPlot" in content
    assert "plotly.js" in content.lower()
    assert len(content) > 1_000_000
    fig = build_xy_figure(df, x_column="V", y_columns=["T_avail"])
    assert len(fig.data[0].x) == 3


def test_plotly_interpolation_html(tmp_path):
    raw = pd.DataFrame({"AOA": [0, 1, 2], "CL": [0.0, 0.1, 0.2],
                        "Cm": [0.1, "#N/A", 0.0]})
    fit = pd.DataFrame({"AOA": [0, .5, 1, 1.5, 2],
                        "CL": [0, .05, .1, .15, .2], "Cm": [.1, .075, .05, .025, 0]})
    html = tmp_path / "check.html"
    assert export_interpolation_html(raw, fit, x_column="AOA",
        y_columns=["CL", "Cm"], output_html=html).exists()
    assert "（補間曲線）" in html.read_text(encoding="utf-8")


def test_contour_grid_and_duplicate_points(tmp_path):
    df = pd.DataFrame({"h": [0, 0, 1000, 1000], "V": [20, 30, 20, 30],
                        "Ps": [3, 4, 2, 3]})
    assert export_contour_html(df, x_column="V", y_column="h", z_column="Ps",
                               output_html=tmp_path / "ps_contour.html").exists()
    with pytest.raises(ValueError, match="重複"):
        export_contour_html(pd.concat([df, df.iloc[[0]]]), x_column="V",
            y_column="h", z_column="Ps", output_html=tmp_path / "bad.html")


def test_plotly_path_requires_html(tmp_path):
    with pytest.raises(ValueError, match=".html"):
        export_xy_html(pd.DataFrame({"x": [1], "y": [2]}),
            x_column="x", y_columns=["y"], output_html=tmp_path / "bad.png")


def test_write_chart_table_overwrites_and_nan():
    ws = MockSheet("41_補間データ", ["AOA", "CL"], 8)
    df = pd.DataFrame({"AOA": [0.0, 1.0], "CL": [0.1, np.nan]})
    write_chart_table(ws, df, start_cell="D5")
    assert "D5" in ws.cleared
    assert ws.range("D5").value == [["AOA", "CL"]]
    assert ws.range("D6").value == [[0.0, 0.1], [1.0, None]]
    assert ws.range("E7").formula == "=NA()"


def test_plotly_empty_y_columns_fail():
    with pytest.raises(ValueError, match="y_columns"):
        build_xy_figure(pd.DataFrame({"V": [10]}), x_column="V", y_columns=[])


def test_excel_range_uses_offset_header():
    wb = MockBook()
    wb.sheets["30_TV"].headers = ["other", "V", "Thrust"]
    update_xy_chart(wb, chart_sheet="40_グラフ", chart_name="offset", series_specs=[
        ExcelSeriesSpec("thrust", "30_TV", "V", "Thrust", header_cell="D5", row_count=7)
    ])
    ch = _find_chart(wb, "offset")
    assert ch.api.series_collection(1).XValues == "30_TV!E6:E12"
    assert ch.api.series_collection(1).Values == "30_TV!F6:F12"
