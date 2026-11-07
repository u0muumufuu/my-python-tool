"""推力を取得する共通モデル（例：高度・真対気速度から利用可能推力を取得）。"""

from __future__ import annotations

from pathlib import Path
from typing import Any
import numpy as np

from utils.interpolation import ExcelInterpolator


class PropulsionModel:
    """速度 V[m/s]・高度 Altitude[m] -> Thrust[N] の簡易モデル。

    本例は固定出力設定の推力表を仮定する。プロペラ機等では
    RPM/スロットル/燃料流量などの入力・出力を適切に拡張する。
    """

    def __init__(
        self,
        *,
        excel_path: str | Path,
        workbook: Any = None,
        sheet_name: str = "11_推進DB",
        start_cell: str = "A1",
        method: str = "regular_grid",
    ) -> None:
        self.interpolator = ExcelInterpolator(
            excel_path=excel_path,
            workbook=workbook,
            sheet_name=sheet_name,
            start_cell=start_cell,
            input_columns=["V", "Altitude"],
            output_columns=["Thrust"],
            method=method,
            grid_method="linear",
        )

    def thrust(self, *, velocity: float, altitude: float) -> float:
        return float(self.interpolator.predict(V=velocity, Altitude=altitude)["Thrust"])

    def supports(self, *, velocity: float, altitude: float) -> bool:
        x = np.array([velocity, altitude], dtype=float)
        model = self.interpolator
        return bool(np.all(np.isfinite(x)) and np.all(x >= model.x_min) and np.all(x <= model.x_max))
