import sys
from pathlib import Path

import numpy as np
import pandas as pd
import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from utils.contour_maps import (
    prepare_contour_grid, extract_contour_lines, contours_to_wide_table,
)
from utils.plotly_charts import export_isolines_html, build_isoline_figure
from utils.excel_charts import update_contour_chart_from_dataframe


def planar_field():
    # GごとにPs=100 Mach + Altitude/100 + G とする。順序を混ぜる。
    rows = [(h, m, g, 100*m + h/100 + g)
            for g in (1, 2)
            for h in (0, 1000, 2000, 3000)
            for m in (0.1, 0.2, 0.3, 0.4)]
    return pd.DataFrame(rows, columns=['Altitude', 'Mach', 'G', 'Ps']).sample(frac=1, random_state=42)


def test_exact_slice_long_form_any_row_order():
    grid = prepare_contour_grid(planar_field(), fixed_value=1.0)
    assert grid.z.shape == (4,4)
    assert np.allclose(grid.z, [[11,21,31,41],[21,31,41,51],[31,41,51,61],[41,51,61,71]])
    assert np.allclose(grid.x, [0.1,0.2,0.3,0.4])


def test_arbitrary_ps_level_coordinates_not_exact_rows():
    grid = prepare_contour_grid(planar_field(), fixed_value=1)
    lines = extract_contour_lines(grid, levels=[25.0])
    assert len(lines) >= 2
    # 元データにPs=25がなくても 100M+h/100+1=25 の等値線が得られる。
    assert not np.isclose(planar_field()['Ps'], 25).any()
    assert np.max(np.abs(100*lines['Mach'] + lines['Altitude']/100 + 1 - 25)) < 1e-10


def test_fixed_value_linear_interpolation():
    grid = prepare_contour_grid(planar_field(), fixed_value=1.5, fixed_mode='linear')
    assert grid.fixed_value == 1.5
    assert np.isclose(grid.z[0,0], 11.5)


def test_fixed_value_linear_without_overlapping_coordinates():
    df = planar_field()
    df = df.loc[~((df['G'] == 2) & (df['Mach'] == 0.4))]
    grid = prepare_contour_grid(df, fixed_value=1.5, fixed_mode='linear')
    assert grid.z.shape == (4,3)   # 共通Machを残す


def test_exact_missing_fixed_value_error():
    with pytest.raises(ValueError, match='ありません'):
        prepare_contour_grid(planar_field(), fixed_value=1.5)


def test_fixed_outside_no_extrapolation():
    with pytest.raises(ValueError, match='範囲外'):
        prepare_contour_grid(planar_field(), fixed_value=3, fixed_mode='linear')


def test_duplicate_xy_raises_instead_of_silent_mix():
    df = planar_field()
    df = pd.concat([df, df.iloc[:1]], ignore_index=True)
    with pytest.raises(ValueError, match='重複'):
        prepare_contour_grid(df, fixed_value=float(df.iloc[0]['G']))


def test_nan_z_only_masked():
    df = planar_field()
    mask = ((df['G'] == 1) & (df['Mach'] == 0.2) & (df['Altitude'] == 1000))
    df.loc[mask, 'Ps'] = '#N/A'
    grid = prepare_contour_grid(df, fixed_value=1)
    assert np.isnan(grid.z[1,1])
    assert np.isfinite(grid.z[0,1])
    lines = extract_contour_lines(grid, levels=[25])
    assert isinstance(lines, pd.DataFrame)


def test_missing_coordinate_dropped():
    df = planar_field()
    df.loc[df.index[:1], 'Mach'] = '=NA()'
    grid = prepare_contour_grid(df, fixed_value=1)
    assert len(grid.x) >= 2


def test_resample_irregular_points_convex_hull():
    df = planar_field()
    df = df[(df['G'] == 1) & ~((df['Mach'] == 0.4)&(df['Altitude'] == 3000))]
    grid = prepare_contour_grid(df, fixed_value=1, resample='linear', nx=13, ny=11)
    assert grid.z.shape == (11,13)
    assert np.isnan(grid.z[-1,-1])  # 凸包外
    assert np.isfinite(grid.z[5,5])


def test_line_segments_are_not_accidentally_connected():
    df = pd.DataFrame([(y,x,1, (x-0.3)**2*(x-0.7)**2 + (y-0.5)**2)
                       for x in np.linspace(0,1,41)
                       for y in np.linspace(0,1,41)],
                      columns=['Altitude','Mach','G','Ps'])
    grid = prepare_contour_grid(df, fixed_value=1)
    lines = extract_contour_lines(grid, levels=[0.0005])
    assert lines['segment_id'].nunique() >= 2
    wide, defs = contours_to_wide_table(lines)
    assert len(defs) >= 2
    assert len(wide.columns) == 2*len(defs)


def test_different_contour_levels_wide_columns():
    grid = prepare_contour_grid(planar_field(), fixed_value=1)
    lines = extract_contour_lines(grid, levels=[20, 30, 40])
    wide, specs = contours_to_wide_table(lines)
    assert specs and all('branch' in label for label,_,_ in specs)
    assert len(wide.columns) == 2*len(specs)
    assert lines['Ps_level'].nunique() == 3


def test_none_found_returns_empty():
    grid = prepare_contour_grid(planar_field(), fixed_value=1)
    lines = extract_contour_lines(grid, levels=[100000])
    assert lines.empty
    wide, specs = contours_to_wide_table(lines)
    assert wide.empty and specs == []


def test_invalid_level_raises():
    grid = prepare_contour_grid(planar_field(), fixed_value=1)
    with pytest.raises(ValueError, match='levels'):
        extract_contour_lines(grid, levels=[0, 0])


def test_plotly_html_offline(tmp_path):
    outfile = export_isolines_html(planar_field(), fixed_value=1, levels=[20,30,40],
                                    output_html=tmp_path / 'iso.html',show_heatmap=True)
    content = outfile.read_text(encoding='utf-8')
    assert 'plotly.js' in content.lower()
    assert 'Ps = 20' in content
    assert 'Ps = 30' in content
    assert 'Ps = 40' in content


def test_plotly_has_separate_trace_per_branch():
    grid = prepare_contour_grid(planar_field(), fixed_value=1)
    lines = extract_contour_lines(grid, levels=[20,30])
    fig = build_isoline_figure(grid, lines)
    assert len(fig.data) == lines.groupby(['Ps_level','segment_id']).ngroups


def test_excel_chart_bridge_calls_existing_utils(monkeypatch):
    # Excel実機は不要。DataFrameからテーブル/系列指定が作られることを確認。
    from utils import excel_charts
    calls = {}
    class Sheet:
        def __init__(self): self.name='32_Ps等値線データ'
    class Book:
        sheets = {'32_Ps等値線データ':Sheet()}
    def fake_write(sheet, df, start_cell):
        calls['table'] = df.copy()
        calls['start_cell'] = start_cell
    def fake_update(book, **kwargs):
        calls['series'] = kwargs['series_specs']
        calls['chart_name'] = kwargs['chart_name']
    monkeypatch.setattr(excel_charts, 'write_chart_table', fake_write)
    monkeypatch.setattr(excel_charts, 'update_xy_chart', fake_update)
    lines = update_contour_chart_from_dataframe(Book(), planar_field(),
                       fixed_value=1, levels=[20,30,40])
    assert not lines.empty
    assert calls['start_cell'] == 'A1'
    assert calls['chart_name'] == 'Ps_contour_chart'
    assert len(calls['table'].columns) == 2*len(calls['series'])
    assert all(s.mode == 'lines' for s in calls['series'])
    assert all(s.row_count == len(calls['table']) for s in calls['series'])


def test_excel_chart_dynamic_series_count_when_levels_change(monkeypatch):
    # 既存のExcel COM模擬モデルに接続し、1回目の系列数が減る際に
    # グラフを新規作成せず余剰系列を消すことまで確認する。
    from test_charts import MockBook, MockSheet
    from utils import excel_charts
    book = MockBook()
    book.sheets['32_Ps等値線データ'] = MockSheet('32_Ps等値線データ')
    def fake_write(sheet, frame, *, start_cell):
        assert start_cell == 'A1'
        sheet.headers = list(frame.columns)
        sheet.data_rows = len(frame)
    monkeypatch.setattr(excel_charts, 'write_chart_table', fake_write)

    update_contour_chart_from_dataframe(book, planar_field(), fixed_value=1,
                                        levels=[20,30,40], chart_sheet='40_グラフ')
    ch = book.sheets['40_グラフ'].charts['Ps_contour_chart']
    first_count = ch.api.series_collection.Count
    assert first_count >= 3
    assert book.sheets['40_グラフ'].charts.add_count == 1

    update_contour_chart_from_dataframe(book, planar_field(), fixed_value=1,
                                        levels=[30], chart_sheet='40_グラフ')
    assert 1 <= ch.api.series_collection.Count < first_count
    assert book.sheets['40_グラフ'].charts.add_count == 1
