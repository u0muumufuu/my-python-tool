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
    series_columns : Sequence[str] | str | None, default ()
        系列として振る入力列名。例: ["Mach"]。
        系列なしなら [] または None。文字列 "Mach" も指定可。
    series_values : Sequence[Any] | None, default None
        系列値の指定。
        - series_columns が 1 個なら [0.2, 0.4, ...] のような配列
        - series_columns が複数なら [{"Mach":0.2, "Re":1e6}, ...] のような配列
        None の場合は raw_data から固定条件に合うユニーク値を自動取得する。
    fixed : Mapping[str, float], default {}
        固定する入力列と値。例: {"beta": 0.0, "de": 0.0}。
        入力列のうち、横軸・系列に使わない列はすべて固定する。
        例: CL(alpha, Mach, beta, de) なら x=alpha, series=Mach,
        fixed={"beta":0, "de":0}。系列なしの場合は
        fixed={"Mach":0.2, "beta":0, "de":0}。
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
    series_columns: Sequence[str] | str | None = field(default_factory=tuple)
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
    normalized = {str(k): float(v) for k, v in mapping.items()}
    for key, value in normalized.items():
        if not np.isfinite(value):
            raise ValueError(f"fixed[{key!r}] は有限の数値で指定してください")
    return normalized


def _series_columns(value: Sequence[str] | str | None) -> list[str]:
    """系列なし(None/[])、単列文字列、複数列の指定を共通化する。"""
    if value is None:
        return []
    if isinstance(value, str):
        return [value] if value else []
    return list(value)


def _validate_plot_dimensions(spec: InterpolationPlotSpec, input_columns: Sequence[str]) -> list[str]:
    """1グラフに必要な全入力次元を、x / series / fixed に分配する。"""
    inputs = list(input_columns)
    series = _series_columns(spec.series_columns)
    if spec.x_column not in inputs:
        raise ValueError(f"x_column={spec.x_column!r} は補間器の入力列にありません")
    if len(set(series)) != len(series):
        raise ValueError(f"series_columns に重複があります: {series}")
    invalid_series = sorted(set(series) - set(inputs))
    invalid_fixed = sorted(set(spec.fixed) - set(inputs))
    if invalid_series or invalid_fixed:
        raise ValueError(
            f"不明な入力列です: series_columns={invalid_series}, fixed={invalid_fixed}; "
            f"使用可能な列={inputs}"
        )
    if spec.x_column in spec.fixed or spec.x_column in series:
        raise ValueError("横軸列を fixed や series_columns に重複指定できません")
    if set(series) & set(spec.fixed):
        raise ValueError("series_columns と fixed の条件列が重複しています")
    missing = [n for n in inputs if n != spec.x_column and n not in series and n not in spec.fixed]
    if missing:
        raise ValueError(
            f"グラフ {spec.title!r} の入力列 {missing} が未指定です。"
            f"fixed={{...}} または series_columns=[...] に指定してください。"
            f"系列なしにする場合は series_columns=[] とし、残りの全列を fixed に指定します。"
        )
    if not series and spec.series_values is not None and len(spec.series_values) > 0:
        raise ValueError("系列なしのとき series_values は None または [] にしてください")
    return series


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


def _display_units(text: str) -> int:
    """ラベルの表示幅の簡易推定。日本語などの全角文字は約2文字分。"""
    from unicodedata import east_asian_width
    return sum(2 if east_asian_width(ch) in ("F", "W") else 1 for ch in text)


def _wrap_subplot_title(text: str, *, max_units: int) -> str:
    """横並び時の長いサブプロット見出しを適度な位置で折り返す。"""
    lines = []
    current = ""
    for char in str(text):
        if char == "\n":
            lines.append(current)
            current = ""
        elif current and _display_units(current + char) > max_units:
            lines.append(current)
            current = char
        else:
            current += char
    lines.append(current)
    return "<br>".join(lines)


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
    """生値と複数補間方式を任意グリッド配置したPlotly比較レポートにする。

    - 各グラフで横軸・系列列・固定列を独立指定できる（入力4次元以上も可）。
    - 系列なしの場合は series_columns=[] とし、他入力をすべて fixed にする。
    - 凡例は描画領域外の右側に配置し、長いときはPlotly標準スクロールを使う。
    - 図幅・図高とサブプロット間隔はグラフ数から算出し、軸とレンジスライダー
      に独立したスペースを確保する。小画面は縮小せず横スクロールで閲覧する。
    """
    specs = list(plot_specs)
    if not specs:
        raise ValueError("plot_specs が空です")
    models = _as_interpolators(interpolator)
    rows, cols = _subplot_shape(len(specs), ncols=ncols, nrows=nrows)
    first_model = next(iter(models.values()))
    is_comparison = len(models) > 1

    # 凡例が上部タイトルと重ならないよう、右側に専用の余白を設ける。
    # 横軸のレンジスライダーと目盛り・軸名が次の行に重ならないよう高さを確保する。
    figure_width = 310 + 585 * cols
    row_height = 760 if show_range_slider else 550
    figure_height = 135 + row_height * rows
    # 行間のピクセル数を一定に近づけ、行数に応じて必要な高さを増やす。
    gap_px = 275 if show_range_slider else 125
    usable_height = figure_height - 115 - 85
    vertical_spacing = (gap_px / usable_height) if rows > 1 else 0.0
    horizontal_spacing = (85 / (585 * cols)) if cols > 1 else 0.0

    titles = [
        _wrap_subplot_title(spec.title, max_units=(44 if cols > 1 else 72))
        for spec in specs
    ] + [""] * (rows * cols - len(specs))
    fig = make_subplots(
        rows=rows, cols=cols, shared_xaxes=False,
        subplot_titles=titles,
        vertical_spacing=vertical_spacing,
        horizontal_spacing=horizontal_spacing,
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
        series_cols = _validate_plot_dimensions(spec, first_model.input_columns)
        for name, model in models.items():
            if spec.y_column not in model.output_columns:
                raise ValueError(f"補間方式 {name!r} の出力列に {spec.y_column!r} がありません")

        fixed = _normalize_fixed(spec.fixed)
        series_cases = _normalize_series_values(
            raw_data, series_columns=series_cols,
            series_values=spec.series_values, fixed=fixed,
        )
        if not series_cases:
            raise ValueError(
                f"グラフ {spec.title!r}: 条件 {fixed} に一致する系列値がありません。"
                f"series_values を明示するか、fixed の値を確認してください"
            )

        for idx, series_case in enumerate(series_cases):
            color = palette[idx % len(palette)]
            conditions = {**fixed, **series_case}
            series_label = _format_series_label(series_case, fallback="系列なし")
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

        # automarginはPlotlyに軸名・目盛り文字列を考慮した余白調整を依頼する。
        fig.update_xaxes(
            title_text=spec.x_column, title_standoff=18,
            automargin=True, tickfont=dict(size=11),
            row=row, col=col,
        )
        fig.update_yaxes(
            title_text=spec.y_column, title_standoff=16,
            automargin=True, tickfont=dict(size=11),
            row=row, col=col,
        )
        if show_range_slider:
            fig.update_xaxes(
                rangeslider=dict(visible=True, thickness=0.09),
                row=row, col=col,
            )

    # 日本語や長い補間方式名が凡例領域からはみ出さないよう右余白を拡張する。
    longest_legend = max(
        (_display_units(str(trace.name)) for trace in fig.data if trace.showlegend),
        default=0,
    )
    legend_width = min(620, max(285, int(longest_legend * 7.2 + 85)))
    figure_width = 585 * cols + 25 + legend_width

    fig.update_layout(
        title=dict(text=report_title or "", x=0.015, xanchor="left", y=0.987, yanchor="top"),
        template="plotly_white",
        hovermode="closest",
        autosize=False,
        height=figure_height,
        width=figure_width,
        margin=dict(t=115, b=85, l=95, r=legend_width),
        legend=dict(
            title=dict(text="系列 / 補間方式"),
            orientation="v", yanchor="top", y=1.0,
            xanchor="left", x=1.015,
            groupclick="togglegroup",
            font=dict(size=11),
            tracegroupgap=4,
        ),
    )
    # Plotlyの独立した凡例は多数の系列があれば自動スクロール表示になる。
    # HTMLのウィンドウ幅が狭い場合でもFigureを圧縮せず、横スクロールで閲覧する。
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
