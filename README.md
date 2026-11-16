# AircraftPerformance v15 — Excel補間・方程式ソルバー・Plotly比較

この配布版は**補間モデルの構築、生データと複数補間方式の比較、方程式の求根**に特化しています。
専用の飛行性能マップ計算・コンター等値線生成・Excelグラフ自動操作は含めていません。

## 1. プロジェクト構成

| パス | 機能 |
|---|---|
| `src/run_interpolation_comparison.py` | Excel生データと各補間方式を比較する独立した実行・設定ファイル |
| `src/utils/interpolation.py` | `ExcelInterpolator`：Excelデータの補間 |
| `src/utils/plotly_charts.py` | `InterpolationPlotSpec`、比較Figureの生成、HTML出力、汎用XYプロット |
| `src/utils/solvers.py` | `EquationSolver.solve()`／`solve_multiple_roots()` |
| `src/utils/excel_io.py` | Excelへの数値・NaN出力を支援 |
| `src/models/aerodynamics.py` | 空力DB補間ラッパー（CL/CD/Cm） |
| `src/models/propulsion.py` | 推進DB補間ラッパー（Thrust） |
| `src/main.py`、`src/legacy_examples.py` | 既存の汎用補間・交点計算のExcel起動サンプル |
| `excel_vba/RunInterpolationComparison_xlwings.bas` | ExcelボタンからPlotly比較を起動 |
| `docs/INTERPOLATION_COMPARISON.md` | すべての入力形式・実行例・注意点 |

## 2. インストール

```bash
python -m pip install -r requirements.txt
```

Excelを起動せず保存済み `.xlsx` / `.xlsm` を読む場合は、xlwingsが不要な依存ファイルもあります。

```bash
python -m pip install -r requirements_without_xlwings.txt
```

## 3. 入力Excelの基本形式

例：`10_空力DB` シート、A1からヘッダー付きの縦長データ。

| alpha | Mach | CL | CD | Cm |
|---:|---:|---:|---:|---:|
| -5 | 0.2 | … | … | … |
| 0 | 0.2 | … | … | … |
| 5 | 0.2 | … | … | … |
| -5 | 0.4 | … | … | … |

入力列名（`input_columns`）と出力列名（`output_columns`）は自由に変更できます。
規則格子補間（`regular_grid`）には入力列の直積格子が必要です。

## 4. 補間比較レポートを実行

`src/run_interpolation_comparison.py` の **`REPORTS`** を編集してください。
デフォルトでは、空力DBの CL/CD/Cm を2列×2行に配置し、推進DBの推力–高度を1枚別HTMLに出力します。
どちらも複数の補間方式と生データを重ねます。

```bash
python src/run_interpolation_comparison.py --excel AircraftPerformance.xlsx
```

出力先はデフォルトでExcelファイルと同じディレクトリの `reports/` です。
`--output-dir ./reports` でも指定できます。

```python
# REPORTS内の設定例
'methods': {
    'RegularGrid': {'method': 'regular_grid', 'grid_method': 'linear'},
    'RBF': {'method': 'rbf'},
    'LinearND': {'method': 'linear_nd'},
},
'ncols': 2,       # 横方向のグラフ数
'nrows': 2,       # 縦方向のグラフ数。Noneなら必要行数を自動計算
```

グラフ1枚の仕様：

```python
from utils.plotly_charts import InterpolationPlotSpec

InterpolationPlotSpec(
    title='CL-alpha', x_column='alpha', y_column='CL',
    series_columns=['Mach'], series_values=[0.2, 0.4],
    fixed={}, x_range=(-5, 15), x_samples=301,
)
```

保存形式はオフラインで閲覧できる単体HTMLです（Plotly.jsを埋込）。
色で系列、線種で補間方式を識別し、生データはマーカーで表示します。
横軸レンジスライダー、ホバー、表示系列切替が利用できます。

実行後の返り値は `{レポート名: pathlib.Path}` の辞書です。

### Excelから起動する場合（xlwings）

VBAモジュール `excel_vba/RunInterpolationComparison_xlwings.bas` に対応します。

```vb
RunPython "import run_interpolation_comparison; run_interpolation_comparison.run()"
```

## 5. 補間器単体の使い方

```python
from utils.interpolation import ExcelInterpolator

model = ExcelInterpolator(
    excel_path='AircraftPerformance.xlsx', sheet_name='10_空力DB',
    input_columns=['alpha', 'Mach'],
    output_columns=['CL', 'CD', 'Cm'],
    method='regular_grid', grid_method='linear',
)
result = model.predict(alpha=3, Mach=0.2)
print(result)  # {'CL': float, 'CD': float, 'Cm': float}
```

利用可能なメソッド：`regular_grid`、`rbf`、`linear_nd`、`nearest_nd`、
1入力列の場合は`pchip`。NaN/±inf/Noneの入力値は対象の出力をNaNとして返します。
一方、有限の範囲外入力は`check_bounds=True`（既定）の場合エラーになります。

## 6. ソルバーの使い方

```python
from utils.solvers import EquationSolver

solver = EquationSolver(tol=1e-10, residual_tol=1e-8)
one = solver.solve(lambda x: x**2 - 4, bracket=(0, 4))
roots = solver.solve_multiple_roots(
    lambda x: (x-1)*(x-3)*(x-5),
    bounds=[(0, 6)], method='brentq', scan_points=100,
)
print(one.x, [r.x for r in roots])
```

`solve()` は1解、`solve_multiple_roots()` は多点探索で発見した複数解を返します。
数値的な失敗はNaN、引数・方程式形式の指定ミスは例外として扱います。

## 7. サンプルとテスト

- `examples/generate_interpolation_comparison_demo.py`：架空データから比較HTMLを生成
- `examples/interpolation_comparison_demo.html`：ブラウザ閲覧用のサンプル
- `examples/multiple_roots_demo.py`：単一・連立方程式の複数解探索
- `tests/`：補間、求根、Excel模擬入出力、Plotly比較のテスト

```bash
PYTHONPATH=src pytest -q
```

**関連詳細：** [`docs/INTERPOLATION_COMPARISON.md`](docs/INTERPOLATION_COMPARISON.md)
