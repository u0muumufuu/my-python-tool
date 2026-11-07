"""空力係数を取得する共通モデル。Excel の読み込みは生成時の一度だけ。"""

from __future__ import annotations

from pathlib import Path
from typing import Any
import numpy as np

from utils.interpolation import ExcelInterpolator


class AerodynamicsModel:
    """α[deg], Mach -> CL, CD, Cm の簡易空力モデル。

    本例は横滑り角・舵角・重心位置依存性などを省略している。
    実機に合わせて input_columns を拡張できる。
    """

    def __init__(
        self,
        *,
        excel_path: str | Path,
        workbook: Any = None,
        sheet_name: str = "10_空力DB",
        start_cell: str = "A1",
        method: str = "regular_grid",
    ) -> None:
        self.interpolator = ExcelInterpolator(
            excel_path=excel_path,
            workbook=workbook,
            sheet_name=sheet_name,
            start_cell=start_cell,
            input_columns=["alpha", "Mach"],
            output_columns=["CL", "CD", "Cm"],
            method=method,
            grid_method="linear",
        )

    def coefficients(self, *, alpha: float, mach: float) -> dict[str, float]:
        return self.interpolator.predict(alpha=alpha, Mach=mach)

    @property
    def alpha_bounds(self) -> tuple[float, float]:
        i = self.interpolator.input_columns.index("alpha")
        return (float(self.interpolator.x_min[i]), float(self.interpolator.x_max[i]))

    def supports_mach(self, mach: float) -> bool:
        i = self.interpolator.input_columns.index("Mach")
        return bool(np.isfinite(mach) and self.interpolator.x_min[i] <= mach <= self.interpolator.x_max[i])
