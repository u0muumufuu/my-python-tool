"""Excel不要で動作するサンプル。Gを固定しPs=0,10,20の等値線を表示。"""
from pathlib import Path
import sys

import numpy as np
import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))
from utils.contour_maps import prepare_contour_grid, extract_contour_lines
from utils.plotly_charts import export_isolines_html

mach = np.linspace(0.1, 0.9, 25)
altitude = np.linspace(0, 5000, 21)
rows = []
for g in (1.0, 2.0):
    for h in altitude:
        for m in mach:
            ps = 35 - 0.006 * h - 180 * (m - 0.5) ** 2 - 12 * (g - 1) ** 2
            rows.append((h, m, g, ps))

# 並び順は不問。完全に混ぜても同じ結果になる。
data = pd.DataFrame(rows, columns=["Altitude", "Mach", "G", "Ps"])
data = data.sample(frac=1, random_state=2026).reset_index(drop=True)

out = ROOT / "examples"
data.to_csv(out / "ps_data_demo.csv", index=False)

grid = prepare_contour_grid(data, fixed_value=1.0)
lines = extract_contour_lines(grid, levels=[0, 10, 20])
lines.to_csv(out / "ps_contour_points_demo.csv", index=False)
export_isolines_html(
    data,
    fixed_value=1.0,
    levels=[0, 10, 20],
    output_html=out / "Ps_contours_demo.html",
    show_heatmap=True,
    title="Ps contour demonstration (G=1)",
)
print("Data rows:", len(data), "Contour points:", len(lines))
print("HTML:", out / "Ps_contours_demo.html")
