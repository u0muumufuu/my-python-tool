# AircraftPerformance

Excel をフロントエンド、Python を計算・可視化バックエンドとして使うための、航空機性能推算サンプルです。

- **Excel**: 入力値管理、単純計算、結果保存、Excel 標準グラフ
- **Python**: 補間、収束計算、性能計算、Plotly HTML 可視化
- **主な対象**: 空力 DB、推進 DB、T–V 線図、Ps 線図、補間確認

---

## 1. ディレクトリ構成

```text
AircraftPerformance/
├─ README.md
├─ requirements.txt
├─ src/
│  ├─ main.py
│  ├─ analysis/
│  │  └─ performance_maps.py
│  ├─ models/
│  │  ├─ aerodynamics.py
│  │  └─ propulsion.py
│  └─ utils/
│     ├─ interpolation.py
│     ├─ solvers.py
│     ├─ excel_io.py
│     ├─ excel_charts.py
│     ├─ contour_maps.py
│     └─ plotly_charts.py
├─ excel_vba/
└─ tests/
```

---

## 2. セットアップ

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

Windows で Excel から実行する場合は、Excel 側から参照できる Python を使用してください。

---

## 3. Excel と Python の役割分担

### Excel で行うこと
- 機体諸元・飛行条件の入力
- 単純な四則演算や参照計算
- 結果の保存
- Excel 標準グラフの表示

### Python で行うこと
- 補間関数の作成
- 収束計算・方程式求解
- T–V / Ps の計算
- Plotly による HTML 可視化

---

## 4. Excel から Python を実行する

### 4.1 xlwings を使う方法

VBA:

```vb
Sub RunMapsWithCharts()
    RunPython "import main; main.run_maps_with_charts()"
End Sub
```

### 4.2 Plotly レポートを出力する例

VBA:

```vb
Sub ExportPlotlyReports()
    RunPython "import main; main.export_demo_plotly_reports()"
End Sub
```

この例では次の 2 つを出力します。

- `reports/aerodynamics_report.html`
- `reports/propulsion_report.html`

---

## 5. 補間関数 (`utils/interpolation.py`)

### 5.1 対応メソッド
- `rbf`
- `regular_grid`
- `linear_nd`
- `nearest_nd`
- `pchip`（1入力のみ）

### 5.2 ExcelInterpolator の基本例

```python
from utils.interpolation import ExcelInterpolator

aero = ExcelInterpolator(
    excel_path=wb.fullname,
    workbook=wb,
    sheet_name="10_空力DB",
    input_columns=["alpha", "Mach"],
    output_columns=["CL", "CD", "Cm"],
    method="regular_grid",
)

value = aero.predict(alpha=5.0, Mach=0.3)
# {'CL': ..., 'CD': ..., 'Cm': ...}
```

### 5.3 欠損値の扱い
- Excel の `#N/A` や `=NA()` は欠損値として扱います。
- 出力列ごとに独立に補間器を作るので、`Cm` が欠損していても `CL`, `CD` は利用できます。

---

## 6. 収束計算 (`utils/solvers.py`)

単一方程式、複数方程式に対応しています。収束しない場合や `np.nan` は、Excel 側へ `=NA()` として書き込む設計を想定しています。

---

## 7. Excel 標準グラフ (`utils/excel_charts.py`)

### 7.1 通常の XY グラフ

```python
from utils.excel_charts import update_chart_from_table

update_chart_from_table(
    wb,
    chart_sheet="30_TV",
    chart_name="TV_chart",
    data_sheet="30_TV",
    header_cell="A5",
    x_column="V",
    y_columns=["T_required", "T_available"],
    title="T-V Diagram",
)
```

### 7.2 等値線グラフ

```python
from utils.excel_charts import update_contour_chart_from_dataframe

update_contour_chart_from_dataframe(
    wb,
    df,
    fixed_columns=["G"],
    fixed_values=[1.0],
    level_column="Ps",
    levels=[0, 10, 20],
    x_column="Mach",
    y_column="Altitude",
    output_data_sheet="32_Ps等値線データ",
    chart_sheet="40_グラフ",
    chart_name="Ps_contour_chart",
    title="Ps contour (G=1)",
)
```

> 注意: Excel の等値線グラフは、内部的に **中間の座標データ** を `output_data_sheet` へ書き出してから XY グラフ化しています。複数グラフを同時に保持したい場合は、`output_data_sheet` も別にしてください。

---

## 8. Plotly HTML 可視化 (`utils/plotly_charts.py`)

### 8.1 単純な XY グラフ

```python
from utils.plotly_charts import export_xy_html

export_xy_html(
    df,
    x_column="V",
    y_columns=["T_required", "T_available"],
    output_html="reports/tv.html",
    title="T-V",
    show_range_slider=True,
)
```

### 8.2 等値線 HTML

```python
from utils.plotly_charts import export_isolines_html

export_isolines_html(
    df,
    fixed_columns=["G"],
    fixed_values=[1.0],
    level_column="Ps",
    levels=[0, 10, 20],
    x_column="Mach",
    y_column="Altitude",
    output_html="reports/ps_contour.html",
    show_heatmap=True,
)
```

---

## 9. 複数グラフを 1 HTML にまとめる補間確認レポート

今回の更新で、**生値と補間曲線の比較グラフを複数まとめて 1 つの HTML に保存**できるようにしました。

### 9.1 主な機能
- 1 HTML に複数グラフ（サブプロット）を縦に並べる
- 各グラフごとに
  - 横軸列名
  - 縦軸列名
  - 系列として振る列名
  - 系列の離散値
  - 固定する列名と固定値
  - 補間曲線の評価範囲 (`x_range`)
  - 評価点数 (`x_samples`)
  を指定できる
- 各グラフの横軸に **range slider** を付与できる
- Excel へ高密度な補間点を書き戻す必要がない

### 9.2 使うクラス

```python
from utils.plotly_charts import InterpolationPlotSpec
```

### 9.3 空力レポートの例

```python
from utils.plotly_charts import (
    InterpolationPlotSpec,
    export_interpolation_report_html,
)
from utils.interpolation import ExcelInterpolator
import pandas as pd

raw = pd.read_excel("AircraftPerformance.xlsm", sheet_name="10_空力DB")

interp = ExcelInterpolator(
    excel_path="AircraftPerformance.xlsm",
    sheet_name="10_空力DB",
    input_columns=["alpha", "Mach"],
    output_columns=["CL", "CD", "Cm"],
    method="regular_grid",
    check_bounds=False,   # 外挿確認もしたいとき
)

specs = [
    InterpolationPlotSpec(
        title="CL-α（Mach別）",
        x_column="alpha",
        y_column="CL",
        series_columns=["Mach"],
        series_values=[0.2, 0.4, 0.6],
        x_range=(-5.0, 18.0),
        x_samples=301,
    ),
    InterpolationPlotSpec(
        title="CD-α（Mach別）",
        x_column="alpha",
        y_column="CD",
        series_columns=["Mach"],
        series_values=[0.2, 0.4, 0.6],
        x_range=(-5.0, 18.0),
        x_samples=301,
    ),
    InterpolationPlotSpec(
        title="Cm-α（Mach別）",
        x_column="alpha",
        y_column="Cm",
        series_columns=["Mach"],
        series_values=[0.2, 0.4, 0.6],
        x_range=(-5.0, 18.0),
        x_samples=301,
    ),
]

export_interpolation_report_html(
    interp,
    raw,
    plot_specs=specs,
    output_html="reports/aerodynamics_report.html",
    report_title="Aerodynamics interpolation report",
)
```

### 9.4 推進系レポートの例

```python
raw = pd.read_excel("AircraftPerformance.xlsm", sheet_name="11_推進DB")

interp = ExcelInterpolator(
    excel_path="AircraftPerformance.xlsm",
    sheet_name="11_推進DB",
    input_columns=["V", "Altitude"],
    output_columns=["Thrust"],
    method="regular_grid",
    check_bounds=False,
)

specs = [
    InterpolationPlotSpec(
        title="Thrust-V（高度別）",
        x_column="V",
        y_column="Thrust",
        series_columns=["Altitude"],
        series_values=[0, 2000, 4000, 6000],
        x_range=(0.0, 120.0),
        x_samples=301,
    )
]

export_interpolation_report_html(
    interp,
    raw,
    plot_specs=specs,
    output_html="reports/propulsion_report.html",
    report_title="Propulsion interpolation report",
)
```

### 9.5 複数条件を固定する例

```python
InterpolationPlotSpec(
    title="CL-alpha (beta=0, de=0)",
    x_column="alpha",
    y_column="CL",
    series_columns=["Mach"],
    series_values=[0.2, 0.4],
    fixed={"beta": 0.0, "de": 0.0},
    x_range=(-5.0, 18.0),
)
```

### 9.6 系列値を自動取得する例

```python
InterpolationPlotSpec(
    title="CL-alpha",
    x_column="alpha",
    y_column="CL",
    series_columns=["Mach"],
    x_range=(-5.0, 18.0),
)
```

`series_values` を省略すると、生データに含まれる `Mach` のユニーク値を自動で使います。

### 9.7 複数列で系列を振る例

```python
InterpolationPlotSpec(
    title="Thrust-V",
    x_column="V",
    y_column="Thrust",
    series_columns=["Altitude", "Throttle"],
    series_values=[
        {"Altitude": 0.0, "Throttle": 1.0},
        {"Altitude": 3000.0, "Throttle": 1.0},
        {"Altitude": 3000.0, "Throttle": 0.7},
    ],
    x_range=(0.0, 120.0),
)
```

---

## 10. `main.py` の主な入口関数

- `run_maps()` : T–V / Ps の数値計算
- `run_maps_with_charts()` : 上記 + Excel グラフ更新
- `run_ps_contours()` : 長形式データから Ps 等値線グラフ更新
- `export_demo_plotly_reports()` : 空力・推進の補間確認 HTML を出力

---

## 11. テスト

```bash
pytest -q
```

この版では次を検証しています。
- 補間器の作成と欠損値対応
- 収束計算
- Excel グラフ更新
- Ps 等値線
- Plotly HTML 出力
- **複数グラフ補間レポート出力**

---

## 12. 運用上のおすすめ

### Excel グラフ向き
- T–V 線図
- Ps–V 線図
- 日常確認用の定型グラフ

### Plotly HTML 向き
- 補間関数と生値の比較
- 多系列比較
- 外挿挙動の確認
- 等値線の詳細確認

今回のように **補間確認は Plotly HTML に寄せる** のは非常に合理的です。補間点数が多くても Excel を重くせず、HTML 単体でオフライン配布・保存できます。
