"""説明用の仮データからPlotlyのオフラインHTMLを生成する例。"""
from __future__ import annotations

import sys
from pathlib import Path

import numpy as np
import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))
from utils.plotly_charts import export_xy_html, export_interpolation_html

out = Path(__file__).resolve().parent
v = np.linspace(20, 70, 90)
t_req = 0.15 * (v - 42)**2 + 90
t_avail = 250 - 1.65 * v
export_xy_html(
    pd.DataFrame({"V": v, "T_required": t_req, "T_available": t_avail}),
    x_column="V", y_columns=["T_required", "T_available"],
    x_label="V [m/s]", y_label="Thrust [N]", title="T-V (demo only)",
    output_html=out / "TV_demo.html",
)

alpha = np.array([-5, 0, 5, 10, 15], dtype=float)
raw = pd.DataFrame({"alpha": alpha, "CL": 0.15 + 0.095*alpha + 0.04*np.sin(alpha / 4)})
dense = np.linspace(alpha.min(), alpha.max(), 201)
curve = pd.DataFrame({"alpha": dense, "CL": np.interp(dense, raw["alpha"], raw["CL"])})
export_interpolation_html(
    raw, curve, x_column="alpha", y_columns=["CL"],
    output_html=out / "interpolation_demo.html", title="CL-alpha demo (not RBF)",
)
print(out)
