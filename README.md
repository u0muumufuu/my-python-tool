# AircraftPerformance v19 — Excel補間・方程式ソルバー（複数関数・制約付き）・Plotly比較

この配布版は**補間モデルの構築、生データと複数補間方式の比較、方程式の求根**に特化しています。
専用の飛行性能マップ計算・コンター等値線生成・Excelグラフ自動操作は含めていません。

## 1. プロジェクト構成

| パス | 機能 |
|---|---|
| `src/run_interpolation_comparison.py` | Excel生データと各補間方式を比較する独立した実行・設定ファイル |
| `src/utils/interpolation.py` | `ExcelInterpolator`：Excelデータの補間 |
| `src/utils/plotly_charts.py` | `InterpolationPlotSpec`、比較Figureの生成、HTML出力、汎用XYプロット。余白・凡例を自動調整 |
| `src/utils/solvers.py` | `EquationSolver.solve()`／`solve_multiple_roots()`。関数リストによる連立方程式指定、省略可能な不等式制約の事後判定付き |
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

グラフ1枚の仕様（3入力以上でも同様に設定可能）：

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
**v16の改善：** 凡例をグラフの右側に分離し、軸・タイトル・レンジスライダー間の余白を自動調整します。
入力が3変数以上の場合も、各グラフで横軸・系列・固定する列を独立して選べます。例えば
`input_columns=["alpha", "Mach", "beta", "de"]` なら、
`x_column="alpha", series_columns=["Mach"], fixed={"beta":0, "de":0}` と指定します。
**系列なし**なら `series_columns=[]`、`fixed={"Mach":0.2, "beta":0, "de":0}` とします。
4入力の複数パターンは `docs/INTERPOLATION_COMPARISON.md` 第3～4節とサンプルを参照。
`plots` は `InterpolationPlotSpec(...)` のほか、通常の辞書形式 `{...}` でも指定できます。
レポート設定 `show_range_slider=False` でスライダーを非表示にできます。

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

### v19：`func` も関数リスト／lambda式リストで直接指定可能

`solve()` と `solve_multiple_roots()` の両方で、従来の `func=equations` に加え、
`func=[f, g]`、`func=(f, g)`、`func=[lambda x: ..., lambda x: ...]` に対応しました。
**`constraints` の指定方法と同様に、各関数はスカラーまたは1次元の数値列を返せます。**
方程式では出力した残差の**総数が未知数の個数と同じ**である必要があります。
`func` の省略／空配列は不可ですが、`constraints` は省略できます。

```python
from utils.solvers import EquationSolver

solver = EquationSolver()

def f(x): return x[0] + x[1] - 3

def g(x): return x[0] - x[1] - 1

result = solver.solve(
    func=[f, g], x0=[0, 0],
    constraints=[lambda x: 3 - x[0], lambda x: 2 - abs(x[1])],
)
print(result.x)  # [2. 1.]

# 制約なしは従来どおり
result2 = solver.solve(func=[f, g], x0=[0, 0])
```

`args=(...)` を指定した場合は、リスト中の**各方程式の関数**に共通して渡します。
`constraints` は従来どおり、最終解の採否判定にのみ使用されます。
詳しくは **[`docs/SOLVER_EQUATIONS.md`](docs/SOLVER_EQUATIONS.md)** を参照してください。

### v18：求根に不等式制約を付ける（省略可能・関数／lambda両対応）

`solve()` と `solve_multiple_roots()` に、次の2引数を追加しました。

- `constraints=None`：制約関数、または制約関数のリスト。省略／`None`／`[]` では **従来どおりの動作**。
- `constraint_tol=1e-6`：各制約関数値に許容する小さな負値（単位は制約式に依存）。

各制約関数は、数値スカラーまたは1次元の制約値の列を返し、各値が **0以上なら適合** するよう書きます。
実際の許容判定は `c(x) >= -constraint_tol` です。

```python
# 1変数：方程式 x-2=0。ただし x<=1.5 とする → 不適合
r = solver.solve(lambda x: x-2, x0=0,
                 constraints=[lambda x: 1.5-x])
print(r.converged)          # False
print(r.x)                  # nan
print(r.last_x)             # 2.0（収束したが棄却された解）
print(r.constraint_values)  # (-0.5,) に近い

# 複数解：x=1,3,5 から x<=3 の根だけ残す
rs = solver.solve_multiple_roots(
    lambda x: (x-1)*(x-3)*(x-5), [(0,6)],
    constraints=[lambda x: 3-x],
)
print([r.x for r in rs])    # [1.0, 3.0]
```

### v18：1つの関数で複数の制約を指定する

```python
def constraints_eqs(x):
    return [10 - x[0], 25 - abs(x[1])]

result = solver.solve(
    lambda x: [x[0] - 4, x[1] + 2],
    x0=[0, 0], constraints=constraints_eqs,
)
print(result.x, result.constraint_values)  # [4, -2], (6, 23)
```

`constraints=[lambda x: 10-x[0], lambda x: 25-abs(x[1])]` の従来形式や、`constraints=lambda x: [...]` という1つのlambda形式も引き続き対応。`constraints` に複数の関数を並べ、個々の関数がリストを返す混在形式にも対応しています。`constraint_values` では制約値が指定した順に平坦化されます。

`constraints` は**収束後だけ**評価し、探索中の変数は制限しません。
連立方程式では `x` は配列なので `lambda x: alpha_stall-x[0]` のように書きます。
`CL<=CLmax` だけでは失速後の解を排除できない点にも注意してください。

**入力形式・返り値・トリム計算例の詳細**：[`docs/SOLVER_CONSTRAINTS.md`](docs/SOLVER_CONSTRAINTS.md)

## 7. サンプルとテスト

- `examples/generate_interpolation_comparison_demo.py`：架空データから比較HTMLを生成
- `examples/generate_multidimensional_comparison_demo.py`：4入力変数の系列あり／なしの比較HTMLを生成
- `examples/multidimensional_comparison_demo.html`、`.xlsx`：4入力の結果入りサンプル
- `examples/interpolation_comparison_demo.html`：ブラウザ閲覧用のサンプル
- `examples/multiple_roots_demo.py`：単一・連立方程式の複数解探索
- `examples/solver_constraints_demo.py`：省略可能な不等式制約・トリム模擬例
- `examples/solver_equation_forms_demo.py`：`func=[f,g]`、lambdaリスト、制約あり／なしの例
- `tests/test_solver_equation_forms.py`：`func` の指定形式の追加テスト
- `tests/test_solver_constraints.py`：制約あり／なし・棄却／NaN・入力形式の検証
- `tests/`：補間、求根、Excel模擬入出力、Plotly比較のテスト

```bash
PYTHONPATH=src pytest -q
```

**関連詳細：** [`docs/INTERPOLATION_COMPARISON.md`](docs/INTERPOLATION_COMPARISON.md)
