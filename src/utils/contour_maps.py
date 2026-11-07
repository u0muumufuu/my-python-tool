"""多次元の長形式データから断面を抽出し、2D等値線を計算する共通処理。

1行 = 1評価条件（例: Altitude, Mach, G, Ps）。Excel上でレベルごとに
シート分割する必要はない。等値線は ``Ps == 10`` の行を探すのではなく、
2D格子上の数値場から ContourPy により幾何学的に抽出する。

インポート時にはxlwings不要。表DataFrameは、Excel、CSV、Parquetなどの
任意のソースから与えられる。
"""
from __future__ import annotations

from dataclasses import dataclass
from collections.abc import Sequence

import numpy as np
import pandas as pd
from contourpy import contour_generator
from scipy.interpolate import griddata


@dataclass(frozen=True)
class ContourGrid:
    x: np.ndarray
    y: np.ndarray
    z: np.ndarray                 # z.shape == (len(y), len(x))
    x_column: str
    y_column: str
    z_column: str
    fixed_column: str
    fixed_value: float


def _numbers(frame: pd.DataFrame, names: Sequence[str]) -> pd.DataFrame:
    missing = [n for n in names if n not in frame.columns]
    if missing:
        raise KeyError(f"データの列がありません: {missing}")
    data = frame.loc[:, list(names)].copy()
    for name in names:
        data[name] = pd.to_numeric(data[name], errors="coerce").replace(
            [np.inf, -np.inf], np.nan
        )
    return data


def _one_slice(
    data: pd.DataFrame,
    *, x_column: str, y_column: str, z_column: str,
) -> pd.DataFrame:
    """各 (x,y) を1点として抽出。Z欠損は残す（穴として扱う）。"""
    data = data.loc[data[x_column].notna() & data[y_column].notna()].copy()
    if data.duplicated([x_column, y_column]).any():
        raise ValueError("同じ固定条件で (x, y) が重複しています。重複を整理してください")
    return data.loc[:, [x_column, y_column, z_column]]


def prepare_contour_grid(
    data: pd.DataFrame,
    *,
    fixed_column: str = "G",
    fixed_value: float = 1.0,
    x_column: str = "Mach",
    y_column: str = "Altitude",
    z_column: str = "Ps",
    fixed_mode: str = "exact",
    resample: str = "none",
    nx: int = 101,
    ny: int = 81,
    tolerance: float = 1e-9,
) -> ContourGrid:
    """長形式データから1枚の x-y-z 格子断面を作る。

    fixed_mode='exact': 指定Gが元データに存在する必要がある。
    fixed_mode='linear': 指定Gがなくても、上下2つのG断面を補間する。
          ただし上下断面で同じ (x, y) 点が存在するところだけ評価できる。
          G方向の外挿はしない。
    resample='none': 既存 (x,y) 座標の表へピボット。欠損格子は NaN。
    resample='linear': 散在 (x,y) をSciPyの線形補間で等間隔格子に変換。
          データの凸包の外はNaN。穴を埋める処理をするため、欠損格子の
          意味が計算失敗/物理的に存在しない領域なら使用しないこと。
    """
    if fixed_mode not in {"exact", "linear"}:
        raise ValueError("fixed_modeは'exact'または'linear'を指定してください")
    if resample not in {"none", "linear"}:
        raise ValueError("resampleは'none'または'linear'を指定してください")
    if not (np.isfinite(fixed_value) and np.isfinite(tolerance) and tolerance >= 0):
        raise ValueError("fixed_valueおよびtoleranceが不正です")
    if resample == "linear" and (nx < 2 or ny < 2):
        raise ValueError("nx, nyは2以上が必要です")
    names = [fixed_column, x_column, y_column, z_column]
    if len(set(names)) != len(names):
        raise ValueError("fixed/x/y/zの列名はすべて異なる必要があります")

    frame = _numbers(data, names)
    frame = frame[frame[fixed_column].notna()].copy()
    distinct = np.sort(frame[fixed_column].unique())
    if len(distinct) == 0:
        raise ValueError(f"{fixed_column}に有効な値がありません")

    near = distinct[np.isclose(distinct, fixed_value, rtol=0.0, atol=tolerance)]
    if len(near) > 1:
        raise ValueError("指定値付近に複数の固定座標があります。toleranceを小さくしてください")
    if len(near) == 1:
        k = float(near[0])
        section = _one_slice(
            frame.loc[frame[fixed_column] == k],
            x_column=x_column, y_column=y_column, z_column=z_column,
        )
    else:
        if fixed_mode == "exact":
            raise ValueError(f"{fixed_column}={fixed_value} はありません。選択可能: {distinct.tolist()}")
        low = distinct[distinct < fixed_value]
        high = distinct[distinct > fixed_value]
        if len(low) == 0 or len(high) == 0:
            raise ValueError(f"{fixed_column}={fixed_value} は元データ範囲外です（外挿不可）")
        lo = float(low[-1]); hi = float(high[0]); t = (fixed_value - lo) / (hi - lo)
        low_data = _one_slice(frame.loc[frame[fixed_column] == lo],
                              x_column=x_column,y_column=y_column,z_column=z_column)
        high_data = _one_slice(frame.loc[frame[fixed_column] == hi],
                               x_column=x_column,y_column=y_column,z_column=z_column)
        common = low_data.merge(high_data, on=[x_column, y_column], how="inner",
                                validate="one_to_one", suffixes=("_lower", "_upper"))
        section = common.loc[:, [x_column, y_column]].copy()
        section[z_column] = ((1 - t) * common[f"{z_column}_lower"]
                             + t * common[f"{z_column}_upper"])

    if len(section) == 0:
        raise ValueError("断面を構築するための有効な座標点がありません")
    if resample == "none":
        grid = section.pivot(index=y_column, columns=x_column, values=z_column)
        grid = grid.sort_index().sort_index(axis=1)
        xs = grid.columns.to_numpy(dtype=float)
        ys = grid.index.to_numpy(dtype=float)
        zs = grid.to_numpy(dtype=float)
    else:
        valid = section.dropna(subset=[z_column])
        if len(valid) < 3:
            raise ValueError("2次元線形再格子化には3点以上の有効なデータが必要です")
        xs = np.linspace(section[x_column].min(), section[x_column].max(), nx)
        ys = np.linspace(section[y_column].min(), section[y_column].max(), ny)
        xx, yy = np.meshgrid(xs, ys)
        zs = np.asarray(griddata(
            valid[[x_column, y_column]].to_numpy(dtype=float),
            valid[z_column].to_numpy(dtype=float),
            (xx, yy), method="linear", fill_value=np.nan, rescale=True,
        ), dtype=float)
    if len(xs) < 2 or len(ys) < 2:
        raise ValueError("等値線を作成するにはx,y各2点以上が必要です")
    return ContourGrid(xs, ys, zs, x_column, y_column, z_column,
                       fixed_column, float(fixed_value))


def extract_contour_lines(grid: ContourGrid, *, levels: Sequence[float]) -> pd.DataFrame:
    """各Ps等値線を座標点に変換。複数に分離した枝はsegment_idで区別する。

    戻り値: [fixed_column, <z_column>_level, segment_id, point_id,
             <x_column>, <y_column>]
    """
    names = [grid.fixed_column, f"{grid.z_column}_level", "segment_id", "point_id",
             grid.x_column, grid.y_column]
    lev = [float(level) for level in levels]
    if len(lev) != len(set(lev)) or not lev or not np.isfinite(lev).all():
        raise ValueError("levelsは有限で重複のない1個以上の値を指定してください")
    if np.asarray(grid.z).shape != (len(grid.y),len(grid.x)):
        raise ValueError("zの形状とx/yが一致しません")
    # 欠損点をまたぐ線を作らないよう、corner_mask=Falseとする。
    gen = contour_generator(x=grid.x, y=grid.y,
                            z=np.ma.masked_invalid(grid.z),
                            name="serial", line_type="Separate", corner_mask=False)
    frames = []
    for level in lev:
        for segment_id, xy in enumerate(gen.lines(level), start=1):
            if len(xy) < 2:
                continue
            frames.append(pd.DataFrame({
                grid.fixed_column: grid.fixed_value,
                f"{grid.z_column}_level": level,
                "segment_id": segment_id,
                "point_id": np.arange(len(xy)),
                grid.x_column: xy[:, 0],
                grid.y_column: xy[:, 1],
            }))
    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame(columns=names)


def contours_to_wide_table(lines: pd.DataFrame, *,
                           x_column="Mach", y_column="Altitude", z_column="Ps") -> tuple[pd.DataFrame, list[tuple[str, str, str]]]:
    """Excel XY散布図用。各等値線の各枝を独立したX/Y列に展開。

    returns (wide_df, [(legend, x_column_name, y_column_name), ...])
    """
    if lines.empty:
        return pd.DataFrame(), []
    for key in [f"{z_column}_level", "segment_id", x_column, y_column, "point_id"]:
        if key not in lines.columns:
            raise KeyError(f"等値線データに列 {key!r} がありません")
    cols: dict[str, pd.Series] = {}
    defs: list[tuple[str, str, str]] = []
    groups = lines.groupby([f"{z_column}_level", "segment_id"], sort=False)
    for i, ((level, seg), subset) in enumerate(groups, start=1):
        subset = subset.sort_values("point_id")
        colx = f"x_line{i:03d}"
        coly = f"y_line{i:03d}"
        cols[colx] = pd.Series(subset[x_column].to_numpy(dtype=float))
        cols[coly] = pd.Series(subset[y_column].to_numpy(dtype=float))
        defs.append((f"{z_column}={level:g} (branch {int(seg)})", colx, coly))
    return pd.DataFrame(cols), defs
