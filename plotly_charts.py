"""Plotly による科学技術グラフのオフラインHTML出力。

主な用途:
- 汎用の XY グラフ
- 生値と補間曲線の比較グラフ
- 複数の補間確認グラフを 1 つの HTML にまとめたレポート

設計方針:
- HTML に Plotly.js を埋め込むため、閲覧者に Python やネット接続は不要
- Excel にグラフ描画用の細かい補間点を書き出さなくてよい
- 列名ベースで軸・固定条件・系列条件を指定する
"""

from __future__ import annotations

from dataclasses import dataclass, field
from itertools import product
from math import ceil
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
    """1つのDataFrameの複数Y列をXYグラフ化する。"""
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
            work = work[np.isclose(col, value, rtol=0, atol=1e-9)]
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


def _as_interpolators(interpolator) -> dict[str, Any]:
    """単一補間器（後方互換）と名称→補間器の辞書を共通化する。"""
    if isinstance(interpolator, Mapping):
        if not interpolator:
            raise ValueError("比較する補間器を1つ以上指定してください")
        models = dict(interpolator)
    else:
        models = {getattr(interpolator, "method", "interpolation"): interpolator}

    for label, model in models.items():
        if not isinstance(label, str) or not label.strip():
            raise ValueError("補間方式の表示名は空でない文字列としてください")
        for attr in ("predict", "input_columns", "output_columns", "x_min", "x_max"):
            if not hasattr(model, attr):
                raise TypeError(f"{label!r} の補間器に {attr!r} がありません")
    reference = next(iter(models.values()))
    for label, model in models.items():
        if list(model.input_columns) != list(reference.input_columns):
            raise ValueError(
                f"補間方式 {label!r} の入力列が他の方式と一致しません: "
                f"{list(model.input_columns)} != {list(reference.input_columns)}"
            )
    return models


def _subplot_shape(count: int, *, ncols: int, nrows: int | None) -> tuple[int, int]:
    if isinstance(ncols, bool) or not isinstance(ncols, int) or ncols < 1:
        raise ValueError("ncols は1以上の整数としてください")
    if nrows is None:
        nrows = ceil(count / ncols)
    if isinstance(nrows, bool) or not isinstance(nrows, int) or nrows < 1:
        raise ValueError("nrows は1以上の整数としてください")
    if nrows * ncols < count:
        raise ValueError(f"{count}個のグラフには nrows*ncols >= {count} が必要です")
    return nrows, ncols


def build_interpolation_report_figure(
    interpolator,
    raw_data: pd.DataFrame,
    *,
    plot_specs: Sequence[InterpolationPlotSpec],
    report_title: str | None = None,
    show_range_slider: bool = True,
    ncols: int = 1,
    nrows: int | None = None,
) -> go.Figure:
    """生値と複数補間方式を、指定行数・列数のPlotlyレポートにする。

    Parameters
    ----------
    interpolator : ExcelInterpolator | Mapping[str, ExcelInterpolator]
        単独補間器、または {"RegularGrid": grid, "RBF": rbf, ...}。
        方式を複数指定すると生値マーカーは1回のみ表示し、各方式を異なる線種で重ねる。
    raw_data : pandas.DataFrame
        入力列と出力列を持つ元のExcelテーブル。
    plot_specs : Sequence[InterpolationPlotSpec]
        1つの設定が1つのサブプロットに対応する。
    ncols : int, default 1
        横方向のグラフ数。既定値は従来どおり1列。
    nrows : int | None, default None
        縦方向のグラフ数。Noneなら必要行数を自動計算する。
        余った枠は空白のままにする。
    show_range_slider : bool, default True
        各グラフの横軸にレンジスライダーを表示する。
    """
    specs = list(plot_specs)
    if not specs:
        raise ValueError("plot_specs が空です")
    models = _as_interpolators(interpolator)
    rows, cols = _subplot_shape(len(specs), ncols=ncols, nrows=nrows)
    first_model = next(iter(models.values()))
    is_comparison = len(models) > 1

    titles = [spec.title for spec in specs] + [""] * (rows * cols - len(specs))
    fig = make_subplots(
        rows=rows, cols=cols, shared_xaxes=False,
        subplot_titles=titles,
        vertical_spacing=(0.16 / (rows - 1) if rows > 1 else 0.0),
        horizontal_spacing=(0.10 / (cols - 1) if cols > 1 else 0.0),
    )
    palette = qualitative.Plotly
    dashes = ["solid", "dash", "dot", "dashdot", "longdash", "longdashdot"]
    method_dash = {name: dashes[i % len(dashes)] for i, name in enumerate(models)}
    legend_seen: set[tuple[str, str]] = set()

    for i, spec in enumerate(specs):
        row, col = divmod(i, cols)
        row += 1
        col += 1
        if spec.y_column not in first_model.output_columns:
            raise ValueError(f"y_column={spec.y_column!r} は補間器の出力列にありません")
        if spec.x_column not in first_model.input_columns:
            raise ValueError(f"x_column={spec.x_column!r} は補間器の入力列にありません")
        if spec.x_column in spec.fixed or spec.x_column in spec.series_columns:
            raise ValueError("横軸列を fixed や series_columns に重複指定できません")
        if set(spec.series_columns) & set(spec.fixed):
            raise ValueError("series_columns と fixed の条件列が重複しています")
        for name, model in models.items():
            if spec.y_column not in model.output_columns:
                raise ValueError(f"補間方式 {name!r} の出力列に {spec.y_column!r} がありません")

        fixed = _normalize_fixed(spec.fixed)
        series_cases = _normalize_series_values(
            raw_data, series_columns=spec.series_columns,
            series_values=spec.series_values, fixed=fixed,
        )
        if not series_cases:
            series_cases = [{}]

        for idx, series_case in enumerate(series_cases):
            color = palette[idx % len(palette)]
            conditions = {**fixed, **series_case}
            series_label = _format_series_label(series_case, fallback=spec.y_column)
            raw_slice = _filtered_raw_slice(
                raw_data, x_column=spec.x_column, y_column=spec.y_column,
                conditions=conditions,
            )
            x_range = spec.x_range or _default_x_range(
                raw_slice, x_column=spec.x_column, interpolator=first_model,
            )

            if spec.show_raw and not raw_slice.empty:
                key = ("raw", series_label)
                fig.add_trace(go.Scatter(
                    x=raw_slice[spec.x_column].to_numpy(),
                    y=raw_slice[spec.y_column].to_numpy(),
                    mode="markers",
                    name=(f"生値: {series_label}" if is_comparison
                          else f"{spec.raw_name or spec.y_column}（生値）: {series_label}"),
                    legendgroup=f"raw:{series_label}",
                    showlegend=key not in legend_seen,
                    marker=dict(size=7, color=color, symbol="circle"),
                ), row=row, col=col)
                legend_seen.add(key)

            if spec.show_interpolated:
                for method_label, model in models.items():
                    try:
                        curve = _evaluate_interpolation_curve(
                            model,
                            x_column=spec.x_column,
                            y_column=spec.y_column,
                            conditions=conditions,
                            x_range=x_range,
                            x_samples=spec.x_samples,
                        )
                    except Exception as exc:
                        raise ValueError(
                            f"グラフ {spec.title!r} / 補間方式 {method_label!r} / "
                            f"系列 {series_label!r} の評価に失敗: {exc}"
                        ) from exc
                    key = (f"method:{method_label}", series_label)
                    fig.add_trace(go.Scatter(
                        x=curve[spec.x_column].to_numpy(),
                        y=curve[spec.y_column].to_numpy(),
                        mode="lines",
                        name=(f"{method_label}: {series_label}" if is_comparison
                              else f"{spec.interpolation_name or spec.y_column}（補間）: {series_label}"),
                        legendgroup=f"method:{method_label}",
                        showlegend=key not in legend_seen,
                        line=dict(color=color, dash=method_dash[method_label], width=2.4),
                    ), row=row, col=col)
                    legend_seen.add(key)

        fig.update_xaxes(title_text=spec.x_column, row=row, col=col)
        fig.update_yaxes(title_text=spec.y_column, row=row, col=col)
        if show_range_slider:
            fig.update_xaxes(rangeslider=dict(visible=True, thickness=0.09), row=row, col=col)

    # 1グラフあたりの幅・高さを確保。スライダーは別の行に重ならないよう余白も使う。
    fig.update_layout(
        title=report_title,
        template="plotly_white",
        hovermode="closest",
        height=max(490 * rows, 510),
        width=max(590 * cols, 850),
        margin=dict(t=145, b=75, l=90, r=60),
        legend=dict(orientation="h", yanchor="bottom", y=1.04,
                    xanchor="left", x=0, groupclick="togglegroup"),
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
    ncols: int = 1,
    nrows: int | None = None,
) -> Path:
    """生値＋単一／複数の補間方式の比較グラフをオフラインHTMLへ保存する。"""
    fig = build_interpolation_report_figure(
        interpolator, raw_data,
        plot_specs=plot_specs,
        report_title=report_title,
        show_range_slider=show_range_slider,
        ncols=ncols,
        nrows=nrows,
    )
    return save_html(fig, output_html)
