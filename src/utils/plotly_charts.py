"""Plotly による科学技術グラフのオフラインHTML出力。

- PythonのDataFrameから直接作成。Excelへグラフ用セルを出力する必要なし。
- HTMLにPlotly.jsを埋め込むので、閲覧者にPythonやネット接続は不要。
- 元CSV等へのリンクではなく、生成時点のデータを埋め込むスナップショット。
  グラフの元数値もHTMLへ埋め込まれるため機密データの配布には注意。
"""

from __future__ import annotations

from pathlib import Path
from typing import Sequence

import numpy as np
import pandas as pd
import plotly.graph_objects as go
from plotly.colors import qualitative


def _make_path(output_html: str | Path) -> Path:
    path = Path(output_html).expanduser()
    if path.suffix.lower() != ".html":
        raise ValueError("出力ファイルは .html 形式としてください")
    path.parent.mkdir(parents=True, exist_ok=True)
    return path


def save_html(fig: go.Figure, output_html: str | Path, *, title: str | None = None) -> Path:
    """Plotly figure を完全独立型のオフラインHTMLへ出力する。"""
    path = _make_path(output_html)
    if title is not None:
        fig.update_layout(title=title)
    fig.write_html(str(path), include_plotlyjs=True, full_html=True, auto_open=False)
    return path


def _as_numeric(df: pd.DataFrame, name: str) -> pd.Series:
    if name not in df.columns:
        raise KeyError(f"DataFrameに列 {name!r} がありません")
    # Excel #N/A、=NA()、空欄などはNaNとして扱う。元データは変更しない。
    return pd.to_numeric(df[name], errors="coerce").replace([np.inf, -np.inf], np.nan)


def build_xy_figure(
    data: pd.DataFrame,
    *,
    x_column: str,
    y_columns: Sequence[str],
    title: str | None = None,
    x_label: str | None = None,
    y_label: str | None = None,
    sort_x: bool = True,
    mode: str = "lines",
) -> go.Figure:
    """T–VやPsなど、1つのDataFrameの複数Y列をXYグラフ化する。"""
    if mode not in ("lines", "markers", "lines+markers"):
        raise ValueError("mode は 'lines', 'markers', 'lines+markers' のいずれか")
    names = list(y_columns)
    if not names:
        raise ValueError("y_columnsが空です")
    x = _as_numeric(data, x_column)
    fig = go.Figure()
    for name in names:
        y = _as_numeric(data, name)
        valid = x.notna() & y.notna()
        frame = pd.DataFrame({"x": x[valid], "y": y[valid]})
        if sort_x:
            frame = frame.sort_values("x", kind="stable")
        fig.add_trace(go.Scatter(
            x=frame["x"].to_numpy(), y=frame["y"].to_numpy(),
            mode=mode, name=name,
        ))
    fig.update_layout(
        title=title,
        xaxis_title=x_label or x_column,
        yaxis_title=y_label or (names[0] if len(names) == 1 else "Value"),
        template="plotly_white",
        hovermode="closest",
    )
    return fig


def export_xy_html(
    data: pd.DataFrame,
    *,
    x_column: str,
    y_columns: Sequence[str],
    output_html: str | Path,
    **figure_options,
) -> Path:
    """DataFrameの複数系列をオフラインHTMLへ直接保存する。"""
    fig = build_xy_figure(data, x_column=x_column, y_columns=y_columns, **figure_options)
    return save_html(fig, output_html)


def export_interpolation_html(
    raw_data: pd.DataFrame,
    curve_data: pd.DataFrame,
    *,
    x_column: str,
    y_columns: Sequence[str],
    output_html: str | Path,
    title: str = "補間確認：生データと補間曲線",
) -> Path:
    """生値マーカーと補間ラインを同じ図に重ねる。

    両DataFrameに同じx_column/y_columns名を用意する。多次元データでは、
    呼び出し側で固定Machなどの断面を抽出してから渡すこと。
    """
    names = list(y_columns)
    if not names:
        raise ValueError("y_columnsが空です")
    fig = go.Figure()
    for i, name in enumerate(names):
        color = qualitative.Plotly[i % len(qualitative.Plotly)]
        x_raw = _as_numeric(raw_data, x_column)
        y_raw = _as_numeric(raw_data, name)
        valid_raw = x_raw.notna() & y_raw.notna()
        fig.add_trace(go.Scatter(
            x=x_raw[valid_raw].to_numpy(), y=y_raw[valid_raw].to_numpy(),
            mode="markers", name=f"{name}（生データ）",
            marker=dict(size=8, symbol="circle", color=color),
        ))

        x_fit = _as_numeric(curve_data, x_column)
        y_fit = _as_numeric(curve_data, name)
        valid_fit = x_fit.notna() & y_fit.notna()
        xy = pd.DataFrame({"x": x_fit[valid_fit], "y": y_fit[valid_fit]})
        xy = xy.sort_values("x", kind="stable")
        fig.add_trace(go.Scatter(
            x=xy["x"].to_numpy(), y=xy["y"].to_numpy(),
            mode="lines", name=f"{name}（補間曲線）",
            line=dict(color=color),
        ))
    fig.update_layout(title=title, xaxis_title=x_column,
                      yaxis_title=names[0] if len(names) == 1 else "Value",
                      template="plotly_white", hovermode="closest")
    return save_html(fig, output_html)


def export_contour_html(
    data: pd.DataFrame,
    *,
    x_column: str,
    y_column: str,
    z_column: str,
    output_html: str | Path,
    title: str | None = None,
) -> Path:
    """規則格子の長形式データ(x,y,z)から2Dコンターを出力する。

    同じ(x,y)の重複はエラー。欠損点は穴として残し、勝手に補間しない。
    """
    x = _as_numeric(data, x_column)
    y = _as_numeric(data, y_column)
    z = _as_numeric(data, z_column)
    frame = pd.DataFrame({"x": x, "y": y, "z": z})
    frame = frame[frame["x"].notna() & frame["y"].notna()]
    if frame.empty:
        raise ValueError("有効な(x, y)座標がありません")
    if frame.duplicated(["x", "y"]).any():
        raise ValueError("重複(x, y)格子点があります")
    grid = frame.pivot(index="y", columns="x", values="z").sort_index().sort_index(axis=1)
    fig = go.Figure(go.Contour(
        x=grid.columns.to_numpy(dtype=float),
        y=grid.index.to_numpy(dtype=float),
        z=grid.to_numpy(dtype=float),
        connectgaps=False,
        colorbar=dict(title=z_column),
    ))
    fig.update_layout(title=title or z_column, xaxis_title=x_column,
                      yaxis_title=y_column, template="plotly_white")
    return save_html(fig, output_html)


def build_isoline_figure(
    grid,
    lines: pd.DataFrame,
    *,
    title: str | None = None,
    show_heatmap: bool = False,
) -> go.Figure:
    """ContourPyで抽出した任意の等値線（枝分かれ含む）をPlotlyで描く。

    ``go.Contour`` の等間隔レベル指定と異なり、任意の levels を正確に
    指定できる。各levelの複数枝は別Scatterとし、接続しない。
    """
    fig = go.Figure()
    if show_heatmap:
        fig.add_trace(go.Heatmap(
            x=grid.x, y=grid.y, z=grid.z,
            colorbar=dict(title=grid.z_column),
            hoverongaps=False,
            name=f"{grid.z_column} map",
        ))
    if not lines.empty:
        lev_name = f"{grid.z_column}_level"
        for level, group in lines.groupby(lev_name, sort=False):
            legend = f"{grid.z_column} = {level:g}"
            for i, (_, branch) in enumerate(group.groupby("segment_id", sort=False)):
                branch = branch.sort_values("point_id")
                fig.add_trace(go.Scatter(
                    x=branch[grid.x_column].to_numpy(),
                    y=branch[grid.y_column].to_numpy(),
                    name=legend,
                    legendgroup=legend,
                    showlegend=(i == 0),
                    mode="lines",
                    hovertemplate=f"{legend}<br>{grid.x_column}: %{{x:.4g}}<br>"
                                  f"{grid.y_column}: %{{y:.4g}}<extra></extra>",
                ))
    fig.update_layout(
        title=title or f"{grid.z_column}等値線 ({grid.fixed_column}={grid.fixed_value:g})",
        xaxis_title=grid.x_column,
        yaxis_title=grid.y_column,
        template="plotly_white",
    )
    return fig


def export_isolines_html(
    data: pd.DataFrame,
    *,
    fixed_value: float,
    levels: Sequence[float],
    output_html: str | Path,
    fixed_column: str = "G",
    x_column: str = "Mach",
    y_column: str = "Altitude",
    z_column: str = "Ps",
    fixed_mode: str = "exact",
    resample: str = "none",
    title: str | None = None,
    show_heatmap: bool = False,
    **grid_options,
) -> Path:
    """長形式データ1枚から任意の固定条件・等値線でオフラインHTML出力。"""
    from .contour_maps import prepare_contour_grid, extract_contour_lines

    grid = prepare_contour_grid(
        data,
        fixed_column=fixed_column, fixed_value=fixed_value,
        x_column=x_column, y_column=y_column, z_column=z_column,
        fixed_mode=fixed_mode, resample=resample, **grid_options,
    )
    lines = extract_contour_lines(grid, levels=levels)
    fig = build_isoline_figure(grid, lines, title=title, show_heatmap=show_heatmap)
    return save_html(fig, output_html)
