"""Plotly による科学技術グラフのオフラインHTML出力。

主な用途:
- T–V、Ps–V などの通常の XY グラフ
- 生値と補間曲線の比較グラフ
- Ps 等値線のようなコンター／アイソライン
- 複数の補間確認グラフを 1 つの HTML にまとめたレポート

設計方針:
- HTML に Plotly.js を埋め込むため、閲覧者に Python やネット接続は不要
- Excel にグラフ描画用の細かい補間点を書き出さなくてよい
- 列名ベースで軸・固定条件・系列条件を指定する
"""

from __future__ import annotations

from dataclasses import dataclass, field
from itertools import product
from pathlib import Path
from typing import Any, Mapping, Sequence

import numpy as np
import pandas as pd
import plotly.graph_objects as go
from plotly.colors import qualitative
from plotly.subplots import make_subplots


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
    return pd.to_numeric(df[name], errors="coerce").replace([np.inf, -np.inf], np.nan)


@dataclass(slots=True)
class InterpolationPlotSpec:
    """補間確認レポート内の 1 グラフ分の仕様。

    Parameters
    ----------
    title : str
        サブプロットのタイトル。
    x_column : str
        横軸に使う入力列名。補間器の input_columns に含まれている必要がある。
    y_column : str
        縦軸に使う出力列名。補間器の output_columns に含まれている必要がある。
    series_columns : Sequence[str], default ()
        系列として振る入力列名。例: ["Mach"]。
    series_values : Sequence[Any] | None, default None
        系列値の指定。
        - series_columns が 1 個なら [0.2, 0.4, ...] のような配列
        - series_columns が複数なら [{"Mach":0.2, "Re":1e6}, ...] のような配列
        None の場合は raw_data から固定条件に合うユニーク値を自動取得する。
    fixed : Mapping[str, float], default {}
        固定する入力列と値。例: {"beta": 0.0, "de": 0.0}
    x_range : tuple[float, float] | None, default None
        補間曲線を評価する x 範囲。None なら raw_data の該当断面の最小最大、
        raw が無い場合は補間器の訓練データ範囲を使用する。
    x_samples : int, default 201
        補間曲線用の評価点数。
    show_raw : bool, default True
        生値マーカーも重ねるかどうか。
    show_interpolated : bool, default True
        補間曲線を表示するかどうか。
    raw_name : str | None, default None
        生値系列名のプレフィックス。None なら y_column を使用。
    interpolation_name : str | None, default None
        補間曲線系列名のプレフィックス。None なら y_column を使用。
    """

    title: str
    x_column: str
    y_column: str
    series_columns: Sequence[str] = field(default_factory=tuple)
    series_values: Sequence[Any] | None = None
    fixed: Mapping[str, float] = field(default_factory=dict)
    x_range: tuple[float, float] | None = None
    x_samples: int = 201
    show_raw: bool = True
    show_interpolated: bool = True
    raw_name: str | None = None
    interpolation_name: str | None = None


# ---------------------------------------------------------------------------
# 既存の基本グラフ
# ---------------------------------------------------------------------------

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
    show_range_slider: bool = False,
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
    if show_range_slider:
        fig.update_xaxes(rangeslider=dict(visible=True))
    return fig


def export_xy_html(
    data: pd.DataFrame,
    *,
    x_column: str,
    y_columns: Sequence[str],
    output_html: str | Path,
    **figure_options,
) -> Path:
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
    show_range_slider: bool = True,
) -> Path:
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
    if show_range_slider:
        fig.update_xaxes(rangeslider=dict(visible=True))
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


# ---------------------------------------------------------------------------
# 等値線（既存）
# ---------------------------------------------------------------------------

def build_isoline_figure(
    grid,
    lines: pd.DataFrame,
    *,
    title: str | None = None,
    show_heatmap: bool = False,
) -> go.Figure:
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
        title=title or (
            f"{grid.z_column}等値線 ("
            + ", ".join(f"{name}={value:g}" for name, value in (
                grid.fixed_conditions or ((grid.fixed_column, grid.fixed_value),)
            )) + ")"
        ),
        xaxis_title=grid.x_column,
        yaxis_title=grid.y_column,
        template="plotly_white",
    )
    return fig


def export_isolines_html(
    data: pd.DataFrame,
    *,
    fixed_value: float = 1.0,
    levels: Sequence[float],
    output_html: str | Path,
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
    show_heatmap: bool = False,
    **grid_options,
) -> Path:
    from .contour_maps import prepare_contour_grid, extract_contour_lines

    grid = prepare_contour_grid(
        data,
        fixed_column=fixed_column, fixed_value=fixed_value,
        fixed=fixed, fixed_columns=fixed_columns, fixed_values=fixed_values,
        level_column=level_column,
        x_column=x_column, y_column=y_column, z_column=z_column,
        fixed_mode=fixed_mode, resample=resample, **grid_options,
    )
    lines = extract_contour_lines(grid, levels=levels)
    fig = build_isoline_figure(grid, lines, title=title, show_heatmap=show_heatmap)
    return save_html(fig, output_html)


# ---------------------------------------------------------------------------
# 複数の補間確認グラフを 1 HTML にまとめる新機能
# ---------------------------------------------------------------------------

def _normalize_fixed(mapping: Mapping[str, Any] | None) -> dict[str, float]:
    if not mapping:
        return {}
    return {str(k): float(v) for k, v in mapping.items()}


def _normalize_series_values(
    raw_data: pd.DataFrame,
    *,
    series_columns: Sequence[str],
    series_values: Sequence[Any] | None,
    fixed: Mapping[str, float],
) -> list[dict[str, float]]:
    cols = list(series_columns)
    if not cols:
        return [{}]

    if series_values is None:
        work = raw_data.copy()
        for name, value in fixed.items():
            col = _as_numeric(work, name)
            work = work[np.isclose(col, value, rtol=0, atol=1e-9) | col.isna()]
        unique_sets: list[dict[str, float]] = []
        if len(cols) == 1:
            col = _as_numeric(work, cols[0]).dropna().unique()
            for value in np.sort(col.astype(float)):
                unique_sets.append({cols[0]: float(value)})
        else:
            sub = work[cols].apply(pd.to_numeric, errors="coerce").dropna().drop_duplicates()
            for _, row in sub.sort_values(cols).iterrows():
                unique_sets.append({name: float(row[name]) for name in cols})
        return unique_sets

    normalized: list[dict[str, float]] = []
    if len(cols) == 1:
        name = cols[0]
        for value in series_values:
            if isinstance(value, Mapping):
                if name not in value:
                    raise ValueError(f"series_values の辞書に {name!r} がありません")
                normalized.append({name: float(value[name])})
            else:
                normalized.append({name: float(value)})
        return normalized

    for value in series_values:
        if not isinstance(value, Mapping):
            raise ValueError("series_columns が複数のとき series_values は辞書の配列にしてください")
        missing = [name for name in cols if name not in value]
        if missing:
            raise ValueError(f"series_values の辞書に列が足りません: {missing}")
        normalized.append({name: float(value[name]) for name in cols})
    return normalized


def _format_series_label(series_case: Mapping[str, float], *, fallback: str) -> str:
    if not series_case:
        return fallback
    return ", ".join(f"{k}={v:g}" for k, v in series_case.items())


def _filtered_raw_slice(
    raw_data: pd.DataFrame,
    *,
    x_column: str,
    y_column: str,
    conditions: Mapping[str, float],
    tolerance: float = 1e-9,
) -> pd.DataFrame:
    if x_column not in raw_data.columns or y_column not in raw_data.columns:
        raise KeyError(f"raw_data に x/y 列 {x_column!r}, {y_column!r} がありません")
    mask = pd.Series(True, index=raw_data.index)
    for name, value in conditions.items():
        if name == x_column:
            continue
        if name not in raw_data.columns:
            raise KeyError(f"raw_data に条件列 {name!r} がありません")
        col = _as_numeric(raw_data, name)
        mask &= np.isclose(col, value, rtol=0, atol=tolerance)
    x = _as_numeric(raw_data, x_column)
    y = _as_numeric(raw_data, y_column)
    frame = pd.DataFrame({x_column: x, y_column: y})
    frame = frame[mask & x.notna() & y.notna()]
    return frame.sort_values(x_column, kind="stable")


def _default_x_range(raw_slice: pd.DataFrame, *, x_column: str, interpolator) -> tuple[float, float]:
    if not raw_slice.empty:
        x = _as_numeric(raw_slice, x_column).dropna()
        return float(x.min()), float(x.max())
    try:
        idx = interpolator.input_columns.index(x_column)
    except ValueError as exc:
        raise ValueError(f"x_column={x_column!r} は補間器の入力列に含まれていません") from exc
    return float(interpolator.x_min[idx]), float(interpolator.x_max[idx])


def _evaluate_interpolation_curve(
    interpolator,
    *,
    x_column: str,
    y_column: str,
    conditions: Mapping[str, float],
    x_range: tuple[float, float],
    x_samples: int,
) -> pd.DataFrame:
    if x_samples < 2:
        raise ValueError("x_samples は 2 以上としてください")
    x0, x1 = map(float, x_range)
    if not np.isfinite([x0, x1]).all() or not x0 < x1:
        raise ValueError(f"x_range が不正です: {x_range!r}")

    coords: dict[str, Any] = {}
    x = np.linspace(x0, x1, int(x_samples))
    for name in interpolator.input_columns:
        if name == x_column:
            coords[name] = x
        elif name in conditions:
            coords[name] = np.full_like(x, float(conditions[name]), dtype=float)
        else:
            raise ValueError(
                f"入力列 {name!r} の値が未指定です。"
                f"fixed または series_columns / series_values で指定してください。"
            )
    y = interpolator.predict(**coords)[y_column]
    return pd.DataFrame({x_column: x, y_column: np.asarray(y, dtype=float)})


def build_interpolation_report_figure(
    interpolator,
    raw_data: pd.DataFrame,
    *,
    plot_specs: Sequence[InterpolationPlotSpec],
    report_title: str | None = None,
    show_range_slider: bool = True,
) -> go.Figure:
    """複数の補間確認グラフを縦に並べた Plotly Figure を作る。

    Parameters
    ----------
    interpolator : ExcelInterpolator
        補間器。`predict(**coords)` と `input_columns`, `output_columns`,
        `x_min`, `x_max` を持つこと。
    raw_data : pandas.DataFrame
        生データ。通常は補間器の元テーブル。
    plot_specs : Sequence[InterpolationPlotSpec]
        各グラフの仕様。
    report_title : str | None
        HTML 全体のタイトル。
    show_range_slider : bool, default True
        すべてのサブプロットの横軸に range slider をつける。
    """
    specs = list(plot_specs)
    if not specs:
        raise ValueError("plot_specs が空です")

    fig = make_subplots(rows=len(specs), cols=1, shared_xaxes=False,
                        subplot_titles=[spec.title for spec in specs],
                        vertical_spacing=0.09)

    palette = qualitative.Plotly
    for row, spec in enumerate(specs, start=1):
        if spec.y_column not in interpolator.output_columns:
            raise ValueError(
                f"y_column={spec.y_column!r} は補間器の output_columns にありません"
            )
        if spec.x_column not in interpolator.input_columns:
            raise ValueError(
                f"x_column={spec.x_column!r} は補間器の input_columns にありません"
            )
        fixed = _normalize_fixed(spec.fixed)
        series_cases = _normalize_series_values(
            raw_data, series_columns=spec.series_columns,
            series_values=spec.series_values, fixed=fixed
        )
        if not series_cases:
            series_cases = [{}]

        for idx, series_case in enumerate(series_cases):
            color = palette[idx % len(palette)]
            conditions = {**fixed, **series_case}
            label = _format_series_label(series_case, fallback=spec.y_column)

            raw_slice = _filtered_raw_slice(
                raw_data,
                x_column=spec.x_column,
                y_column=spec.y_column,
                conditions=conditions,
            )
            x_range = spec.x_range or _default_x_range(
                raw_slice, x_column=spec.x_column, interpolator=interpolator
            )

            if spec.show_raw and not raw_slice.empty:
                fig.add_trace(go.Scatter(
                    x=raw_slice[spec.x_column].to_numpy(),
                    y=raw_slice[spec.y_column].to_numpy(),
                    mode="markers",
                    name=f"{spec.raw_name or spec.y_column}（生値）: {label}",
                    legendgroup=f"raw-{row}-{label}",
                    marker=dict(size=7, color=color),
                ), row=row, col=1)

            if spec.show_interpolated:
                curve = _evaluate_interpolation_curve(
                    interpolator,
                    x_column=spec.x_column,
                    y_column=spec.y_column,
                    conditions=conditions,
                    x_range=x_range,
                    x_samples=spec.x_samples,
                )
                fig.add_trace(go.Scatter(
                    x=curve[spec.x_column].to_numpy(),
                    y=curve[spec.y_column].to_numpy(),
                    mode="lines",
                    name=f"{spec.interpolation_name or spec.y_column}（補間）: {label}",
                    legendgroup=f"fit-{row}-{label}",
                    line=dict(color=color),
                ), row=row, col=1)

        fig.update_xaxes(title_text=spec.x_column, row=row, col=1)
        fig.update_yaxes(title_text=spec.y_column, row=row, col=1)
        if show_range_slider:
            fig.update_xaxes(rangeslider=dict(visible=True), row=row, col=1)

    height = max(420 * len(specs), 480)
    fig.update_layout(
        title=report_title,
        template="plotly_white",
        hovermode="closest",
        height=height,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0),
    )
    return fig


def export_interpolation_report_html(
    interpolator,
    raw_data: pd.DataFrame,
    *,
    plot_specs: Sequence[InterpolationPlotSpec],
    output_html: str | Path,
    report_title: str | None = None,
    show_range_slider: bool = True,
) -> Path:
    """複数の補間確認グラフを 1 つの HTML として保存する。"""
    fig = build_interpolation_report_figure(
        interpolator,
        raw_data,
        plot_specs=plot_specs,
        report_title=report_title,
        show_range_slider=show_range_slider,
    )
    return save_html(fig, output_html)
