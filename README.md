# my-python-tool

Excel のシートに空力・エンジン特性などの数値テーブルを保存し、Python / SciPy で補間・非線形方程式の求根を行い、結果を Excel のセル・グラフに返すためのコードです。

- `src/utils/interpolation.py`：Excel 表から補間関数を生成する `ExcelInterpolator`。
- `src/utils/solvers.py`：単一・連立の非線形方程式を求解する `EquationSolver`。
- `src/utils/excel_io.py`：計算結果をExcelへ出力し、`np.nan`/`np.inf` を `=NA()` に変換する `write_excel_result`。
- `src/main.py`：Excel のボタンから呼び出す実装例。2つの補間関数の交点 `f(x) - g(x) = 0` も扱います。

**この版（v4）の変更点**：`write_excel_result()` を新設し、求根だけでなく補間・性能計算など、すべての出力に対して非有限値を Excel の `#N/A` として書き込めます。`src/main.py` でもこの関数を使用します。`EquationSolver.solve()` は**未収束時に例外を投げる代わりに、既定で `result.x = np.nan` を返します**。連立の場合は要素がすべて `np.nan` の配列になります。`result.converged` で成否を必ず確認してください。例外が必要なら `raise_on_failure=True` を指定します。

## 1. ファイル構成とインストール

```text
AircraftAnalysis/
├── AircraftAnalysis.xlsm          # ご自身で用意するExcelブック（配布物には含まれません）
├── src/
│   ├── main.py                    # Excel連携の例
│   └── utils/
│       ├── __init__.py
│       ├── interpolation.py       # ExcelInterpolator / ExcelRBFInterpolator
│       ├── excel_io.py            # write_excel_result (NaN→Excel #N/A)
│       └── solvers.py             # EquationSolver / EquationResult / ConvergenceError
├── tests/
│   ├── test_interpolation.py
│   ├── test_solvers.py
│   └── test_excel_io.py
├── requirements.txt
└── TEST_RESULTS.md
```

Python の利用環境に必要なライブラリを入れます。

```bash
python -m pip install -r requirements.txt
```

Excel のボタンから呼び出すには、xlwings のアドインまたは Standalone 方式のVBA連携が必要です。`PYTHONPATH` に `src` フォルダを指定し、Excel ブックはVBAマクロを保存できる `.xlsm` にしてください。

```vb
Sub RunCalculation()
    RunPython "import main; main.run()"
End Sub
```

> 補間・求根の自動テストは Excel アプリを起動せずに行えます。配布物に Excel ブックは含まれません。

---

# 2. `interpolation.py`：ExcelInterpolator の使い方

## 2.1 Excel 表の準備

例：「空力DB」シートの `D5` に見出し行を置き、その下に数値を配置します。

| alpha | Mach | CL | CD | Cm |
|---:|---:|---:|---:|---:|
| 0 | 0.1 | 0.20 | 0.030 | 0.06 |
| 5 | 0.1 | 0.70 | 0.045 | 0.01 |
| 10 | 0.1 | 1.10 | 0.080 | -0.04 |
| 0 | 0.2 | 0.22 | 0.032 | 0.06 |
| 5 | 0.2 | 0.72 | 0.048 | 0.01 |
| 10 | 0.2 | 1.12 | 0.085 | -0.04 |

この例は `alpha` × `Mach` の **3 × 2 = 6個の全組合せ**が存在するため、`regular_grid` で読み込めます。格子点は等間隔でなくてもよく、Excel内で行が並び替わっていてもかまいません。

前提・注意点：

- ヘッダーは開始セルから右に連続し、データはヘッダーの次の行から下に連続すること。表の途中に完全空白行を入れないこと。
- 列名は重複不可。`input_columns` / `output_columns` は Excel のヘッダー名と**完全一致**させること。
- 対象となる列の値は数値のみ。欠損・文字列・NaN・無限大は許可しません。
- 入力座標の組合せが重複しているとエラーになります。変化しない入力列も指定できません。
- 保存済みファイルを読み込む場合は `.xlsx` / `.xlsm` に対応します。`.xls` は未対応です。

## 2.2 基本的な利用例

```python
from utils.interpolation import ExcelInterpolator

aero = ExcelInterpolator(
    excel_path="AircraftAnalysis.xlsm",  # Excelファイルの保存場所
    sheet_name="空力DB",
    start_cell="D5",                     # デフォルト: "A1"
    input_columns=["alpha", "Mach"],
    output_columns=["CL", "CD", "Cm"],
    method="regular_grid",
)

result = aero.predict(alpha=5.0, Mach=0.2)
print(result)         # {'CL': 0.72, 'CD': 0.048, 'Cm': 0.01} に相当
print(result["CL"])   # float
```

`excel_path` は文字列または `pathlib.Path` で指定できます。**補間器の構築時に表を一度だけ読み込みます**。その後の `predict()` で Excel にアクセスしません。

### 開いているExcel（未保存の変更を含む）を利用

Excel の `RunPython` から呼び出す場合は、**`workbook=wb` を指定**するのがおすすめです。

```python
import xlwings as xw
from utils.interpolation import ExcelInterpolator

wb = xw.Book.caller()  # Excelのボタンから呼び出されたブック

aero = ExcelInterpolator(
    excel_path=wb.fullname,
    sheet_name="空力DB",
    input_columns=["alpha", "Mach"],
    output_columns=["CL", "CD", "Cm"],
    method="regular_grid",
    start_cell="D5",
    workbook=wb,   # Excel上の、未保存のセル変更も反映できる
)
```

`workbook` 省略時は、ディスクに**保存済み**のファイルから `openpyxl` で読み込みます。Excel数式のセルは、保存時に保持されているキャッシュ値を読みます（Python側でExcel数式を計算しません）。`workbook=wb` の場合は起動中のExcelから値を取得するため、必要なら読み込み前に `wb.app.calculate()` を実行してください。

## 2.3 主なコンストラクタ引数

| 引数 | 型・既定値 | 用途 |
|---|---|---|
| `excel_path` | `str \| Path` 必須 | Excelファイルパス。`workbook` 指定時も管理情報として渡す |
| `sheet_name` | `str` 必須 | 表のあるシート名 |
| `input_columns` | 文字列のリスト／タプル | 独立変数の列名（複数可）。順序は内部の次元順 |
| `output_columns` | 文字列または文字列のリスト／タプル | 出力する列名（複数可） |
| `method` | `"rbf"` | 補間方式。下表参照 |
| `start_cell` | `"A1"` | 見出し行の左上セル。例：`"D5"` |
| `workbook` | `None` | xlwings の `Book`。未保存セル変更も読み込みたいときに指定 |
| `check_bounds` | `True` | 各入力の最小・最大値を越えた評価を拒否する |
| `smoothing` | `0.0` | RBFの平滑化パラメータ |
| `neighbors` | `None` | RBFの評価時に使う近傍点数。`None` は全点 |
| `rbf_kernel` | `"thin_plate_spline"` | RBFのカーネル |
| `rbf_epsilon` | `None` | カーネルによって必要になる形状パラメータ |
| `grid_method` | `"linear"` | 規則格子補間の内部方式 |

`method` の文字列は小文字で扱われます。`smoothing` / `neighbors` はRBF専用です。

## 2.4 選べる補間方式

| `method` | SciPy | 想定するデータ | 注意点 |
|---|---|---|---|
| `"regular_grid"` | `RegularGridInterpolator` | 直交格子（全組合せがある） | 欠損格子があるとエラー。格子間隔は不等間隔でも可 |
| `"rbf"` | `RBFInterpolator` | 不規則に分布する散在点 | 距離を使うため内部で min–max 正規化。多数点では計算負荷に注意 |
| `"linear_nd"` | `LinearNDInterpolator` | 2次元以上の散在点 | 凸包外は評価不可（NaNとなり、クラスからValueError） |
| `"nearest_nd"` | `NearestNDInterpolator` | 散在点の最近傍 | 値は段階的に変化する。入力の正規化あり |
| `"pchip"` | `PchipInterpolator` | 1次元の曲線 | 局所的な単調性・形状の保持を意図する |

**空力・エンジン特性の選び方**：

- 全格子点が存在する空力6分力テーブルやエンジン性能マップ → `regular_grid` を第一候補にする。
- CFD、風洞試験などで不規則に配置された複数変数の点 → `rbf` または `linear_nd`。
- 1次元の特性曲線 → `pchip`。
- 不連続な最近傍選択でも問題ないデータ → `nearest_nd`。

### `regular_grid` の補間次数を選ぶ

```python
interpolator = ExcelInterpolator(
    excel_path="AircraftAnalysis.xlsm",
    sheet_name="空力DB",
    input_columns=["alpha", "Mach"],
    output_columns=["CL", "CD"],
    method="regular_grid",
    grid_method="linear",  # nearest / slinear / cubic / quintic / pchip も指定可能
)
```

高次補間では、各軸に必要な点数などの条件があります。格子データの密度が十分でない場合は `linear` を使ってください。

### `rbf` の自動正規化

`rbf`（および `linear_nd` / `nearest_nd`）では、各独立変数について

\[
 x_{i,\mathrm{norm}} = \frac{x_i-x_{i,\min}}{x_{i,\max}-x_{i,\min}}
\]

を**クラス内部で**実施します。補間器の作成時と評価時の両方で同じ変換を使うため、呼び出し側で正規化する必要はありません。なお、各変数の物理的な重要度が異なる場合は、0–1正規化だけで適切な距離になるとは限りません。

```python
engine = ExcelInterpolator(
    excel_path="AircraftAnalysis.xlsm",
    sheet_name="エンジンDB",
    input_columns=["RPM", "Altitude"],
    output_columns=["Power", "FuelFlow"],
    method="rbf",
    smoothing=0.0,
    neighbors=None,
)

result = engine.predict(RPM=3000, Altitude=1000)
print(result["Power"], result["FuelFlow"])
```

RBFは全点を使用する場合、点数増加に伴う計算・メモリ負荷が大きくなります。点数が多い場合には `neighbors` の指定や `regular_grid` の採用を検討してください。

## 2.5 複数条件を一括評価する／`reload()`

```python
import numpy as np

alpha = np.array([0.0, 5.0, 10.0])
result = aero.predict(alpha=alpha, Mach=0.2)
print(result["CL"])  # shape=(3,) の NumPy 配列

# Excelシートの値を編集した後に補間モデルを作り直す
# workbook=wb を使っていれば未保存の変更も対象
# workbook を省略していればExcel側で保存してから reload()
aero.reload()
```

スカラー値と配列値は NumPy のブロードキャスト規則に従って組み合わせられます。`predict()` は**常に列名をキーとする辞書**を返します（スカラー入力なら値は `float`、配列入力なら `ndarray`）。`aero(...)` と直接呼び出す方法も `predict()` と同等です。

範囲外について：`check_bounds=True` は**各入力変数の最小～最大の箱型範囲**だけを検査します。散在点で定義される多次元の凸包外を完全に判定する機能ではありません。`linear_nd` は凸包外がNaNになればValueErrorを発生させます。`check_bounds=False` では補間方式に応じて外挿が行われますが、物理的妥当性は保証しません。

エラー例：列名違い、入力座標の重複、数値以外・欠損、規則格子の欠損、評価点の範囲外、必要点数不足、データが一次元部分空間に偏ることによるSciPyの失敗など。

---

# 3. `solvers.py`：EquationSolver の使い方

`EquationSolver` は**数値関数を `func=...` として渡す**設計です。対象関数はExcelから作った `ExcelInterpolator` でも、通常のPython関数でも構いません。

## 3.1 方程式数・区間に応じた解法の選択

| 入力条件 | 解法 | 必須指定 | 戻り値の `result.x` |
|---|---|---|---|
| 単一方程式・区間あり | `scipy.optimize.brentq` | `bracket=(lower, upper)` | `float` |
| 単一方程式・区間なし | `scipy.optimize.root` | `x0=初期値` | `float` |
| 連立方程式（同数の未知数と式） | `scipy.optimize.root` | `x0=[初期値1, ...]` | `ndarray` |

`bracket` と `x0` は**同時に指定できません**。`brentq` では、通常、区間内で関数が連続し、両端の関数値が異符号（または一方が0）である必要があります。符号条件を満たさない場合、このクラスでは**未収束結果 `x=NaN` を返す**のが既定です。下限が上限以上など、入力そのものの誤りは `ValueError` です。

## 3.2 単一方程式・区間あり（brentq）

\(f(x)=x^2-4=0\) の正の解を求める例：

```python
from utils.solvers import EquationSolver

solver = EquationSolver()
result = solver.solve(lambda x: x**2 - 4, bracket=(0.0, 3.0))

print(result.x)          # 約 2.0
print(result.converged)  # True
print(result.method)     # "brentq"
print(result.residual)   # 0 に近い値
```

## 3.3 単一方程式・区間なし（root）

初期値から探索し、探索範囲は限定しません。

```python
result = solver.solve(lambda x: x**2 - 4, x0=1.5)
print(result.x)       # 約 2.0
print(result.method)  # "root/hybr"
```

初期値によって別の根に収束する、根が見つからない、途中で補間器の範囲外を評価するなどの可能性があります。

## 3.4 連立非線形方程式（root）

\[
\begin{cases}
x+y-3 = 0 \\
x-y-1 = 0
\end{cases}
\]

```python
import numpy as np
from utils.solvers import EquationSolver

def equations(values):
    x, y = values
    return [x + y - 3, x - y - 1]

result = EquationSolver().solve(equations, x0=[0.0, 0.0])

if result.converged:
    print(result.x)  # np.ndarray: [2., 1.]
else:
    print(result.x)  # np.ndarray: [nan, nan]
```

連立の場合、`func` が**未知数の個数と同じ長さの1次元の残差ベクトル**を返す必要があります。

## 3.5 追加引数を渡す

```python
def force_balance(x, k, force):
    return k*x - force

result = solver.solve(force_balance, bracket=(0, 10), args=(2.0, 6.0))
print(result.x)  # 3.0
```

`args` はタプルで指定してください。

## 3.6 未収束時の扱い：NaN と Excel の `#N/A`

**未収束時に `result.x` へ最後の推定値を入れると、誤って有効な解として利用される危険があります。** この版では既定で次のように返します。

```python
import numpy as np
from utils.solvers import EquationSolver

solver = EquationSolver(root_method="lm")
result = solver.solve(lambda x: x**2 + 1.0, x0=1.0)

print(result.converged)  # False
print(np.isnan(result.x))  # True
print(result.residual_norm)  # 最終評価点での最大絶対残差（診断用）
print(result.last_x)  # 最終推定値（診断用。有効な解ではない）
print(result.message)  # 未収束の理由
```

- **単一方程式**：`result.x` は Python の `float('nan')`（`np.isnan()` で判定可能）。
- **連立方程式**：`result.x` は未知数の数と同じ長さの、すべてが `np.nan` の配列。
- `result.residual`：最終評価点での残差を**診断用として保持**。区間の両端が同符号の場合は `NaN`。
- `result.last_x`：アルゴリズムが最後に推定した値（存在する場合）。収束解として使用しないこと。
- `result.converged=False` の場合、解が存在しないと証明できたとは限りません。**解を見つけられなかった**ことを意味します。

**重要：Python の `np.nan` と Excel の `=NA()` は別物です。** `np.nan` をそのまま Excel セルに代入しても、必ず `#N/A` と表示されるわけではありません。Excel 上で `#N/A` を表示したい場合は `=NA()` を**数式として**設定してください。

```python
# wb は xw.Book.caller() で得たブック
result = EquationSolver().solve(lambda x: x**2 + 1, x0=1.0)
ws = wb.sheets["計算結果"]

if result.converged:
    ws.range("B2").value = result.x
else:
    ws.range("B2").formula = "=NA()"  # Excelの #N/A

ws.range("B3").value = "OK" if result.converged else result.message
```

複数未知数なら、それぞれの出力セルについて `=NA()` を入れてください。`result.x` をExcelに書くだけでは、PythonとExcelのNaN表現が一致しない場合があります。

### 未収束を例外にしたい場合

```python
from utils.solvers import ConvergenceError

try:
    solver.solve(lambda x: x**2 + 1, x0=1.0, raise_on_failure=True)
except ConvergenceError as exc:
    print(exc.result.converged)  # False
    print(exc.result.x)          # nan
    print(exc.result.last_x)     # 最終推定値があれば表示
```

ただし、**誤った型・初期値不足・不正な区間指定・方程式からNaNが返る・補間器が範囲外を拒否した**などの入力・評価エラーは、`raise_on_failure=False` でも通常の例外になります。これは「収束できなかった」と「入力が不正」を区別するためです。

## 3.7 引数一覧

### `EquationSolver(...)`

| 引数 | 既定 | 意味 |
|---|---|---|
| `root_method` | `"hybr"` | `scipy.optimize.root` で使う方式 |
| `tol` | `1e-10` | SciPyに渡す解探索の収束許容値（brentqの `xtol`、rootの `tol`） |
| `residual_tol` | `1e-8` | 残差の最大絶対値がこの値以下であることを追加で要求 |
| `maxiter` | `100` | `brentq` の最大反復回数 |
| `root_options` | `None` | `root` に渡す方式依存のオプション（辞書） |

`root_options` の例（`hybr` の最大評価回数を指定）：

```python
solver = EquationSolver(
    root_method="hybr",
    tol=1e-10,
    residual_tol=1e-7,
    root_options={"maxfev": 300},
)
```

**注意**：`maxiter` は `brentq` 側の制限であり、`root` の反復・評価回数を指定するには `root_options` を使います。また、`tol` は残差の許容値と同じではありません。残差の合格基準には `residual_tol` を使ってください。

### `solver.solve(func, ...)`

| 引数 | 既定 | 意味 |
|---|---|---|
| `func` | 必須 | 残差関数。1変数ではスカラーを、連立では残差ベクトルを返す |
| `x0` | `None` | `root` の初期値。単一なら数値、連立なら長さnのリストなど |
| `bracket` | `None` | `brentq` の上下限 `(lower, upper)` |
| `args` | `()` | 関数に追加で渡す引数のタプル |
| `raise_on_failure` | `False` | `True` なら未収束時に `ConvergenceError` |

### `EquationResult` の属性

| 属性 | 内容 |
|---|---|
| `x` | 採用した解。未収束時は `nan` または `nan` 配列 |
| `converged` | SciPy成功判定と残差許容値の両方が成立したか |
| `residual` | 評価した残差（収束しなかった場合も診断情報を可能な限り保持） |
| `residual_norm` | 残差の最大絶対値（性質上 NaN となる場合もある） |
| `last_x` | 未収束時の最後の推定値。なければ `None` |
| `method` | 使用した解法（例：`"brentq"` / `"root/hybr"`） |
| `message` | 終了理由・失敗理由 |
| `function_calls` | 関数評価回数の参考値 |
| `iterations` | 反復回数（取得可能な場合） |

飛行性能やトリム計算の連立式では、揚力[N]・モーメント[N·m]など、式ごとの残差の単位と桁が異なることがあります。適切な基準値で**残差を無次元化**してから `root` に渡すと、判定が明確になります。

---

# 4. 補間関数2つの交点を求める例（Excelと連携）

`src/main.py` の `run_intersection_example()` は、Excelから読み取った2つの補間関数を生成し、\(f(x)-g(x)=0\) を解きます。

Excelシート **「交点DB」** の `A1` から次の表を配置します。

| x | f | g |
|---:|---:|---:|
| 0 | 1 | 4 |
| 1 | 3 | 4.5 |
| 2 | 5 | 5 |
| 3 | 7 | 5.5 |
| 4 | 9 | 6 |

この表は `f(x)=2x+1`、`g(x)=0.5x+4` に相当するため、交点は `x=2` です。

```python
from utils.interpolation import ExcelInterpolator
from utils.solvers import EquationSolver

# ここでは wb = xw.Book.caller() により呼び出し元ブックを取得済みとする
f_interp = ExcelInterpolator(
    wb.fullname, "交点DB", ["x"], "f",
    method="pchip", workbook=wb,
)
g_interp = ExcelInterpolator(
    wb.fullname, "交点DB", ["x"], "g",
    method="pchip", workbook=wb,
)

def difference(x):
    return f_interp.predict(x=x)["f"] - g_interp.predict(x=x)["g"]

result = EquationSolver().solve(difference, bracket=(0.0, 4.0))

ws = wb.sheets["交点結果"]
if result.converged:
    ws.range("B2").value = result.x
else:
    ws.range("B2").formula = "=NA()"
import math
if math.isfinite(result.residual):
    ws.range("B3").value = result.residual
else:
    ws.range("B3").formula = "=NA()"
ws.range("B4").value = result.method
ws.range("B5").value = "OK" if result.converged else result.message
wb.save()
```

Excel の「交点結果」シートを用意し、VBAボタンに以下のマクロを割り当てます。

```vb
Sub RunIntersectionExample()
    RunPython "import main; main.run_intersection_example()"
End Sub
```

`src/main.py` の `run()` は、別例として空力DBの `regular_grid` 補間とエンジンDBの `rbf` 補間を扱います。具体的なシート名・セル位置は `main.py` 冒頭のコメントを参照してください。

---

# 5. テストと検証

```bash
python -m pip install pytest
python -m pytest -q tests/
python -m doctest src/utils/solvers.py
python -m compileall -q src tests
```

テストは実際の `.xlsx` / `.xlsm` ファイルを使った補間器検証と、Excelブックを模擬した読み書き確認を含みます。**Windows / macOS の Excel 実機上でのVBAボタン実行は、別途利用環境で確認してください。**

詳しい結果は `TEST_RESULTS.md` に記載します。


---

# 5. `excel_io.py`：計算結果をExcelに書き込む共通関数

計算結果をExcelに書き込むときは、**補間結果か求根結果かを区別せず**、
`write_excel_result()` を使用してください。これでPythonの `np.nan` をExcelの
エラー値 `#N/A` として表示できます。具体的には、該当セルの `.formula`
に `=NA()` を設定します（文字列の `"NA()"` を保存するわけではありません）。
`+np.inf` / `-np.inf` も同様に `#N/A` に変換します。

```python
import numpy as np
from utils.excel_io import write_excel_result

# wb = xw.Book.caller()
ws = wb.sheets["計算結果"]

write_excel_result(ws, "B2", 0.32)           # 数値をそのまま保存
write_excel_result(ws, "B3", np.nan)        # Excel B3に =NA() → #N/A
write_excel_result(ws, "B4", np.inf)        # Excel B4に =NA() → #N/A

# 列方向に一括出力（1次元配列は既定で縦方向）
write_excel_result(ws, "D2", [0.30, np.nan, 0.51])

# 横方向に出力したい場合
write_excel_result(ws, "F2", [0.30, np.nan, 0.51], orientation="row")

# 2次元配列は指定された形のまま出力
write_excel_result(ws, "H2", np.array([[1.0, np.nan], [3.0, 4.0]]))

wb.app.calculate()
wb.save()
```

## 関数の引数

| 引数 | 意味 |
|---|---|
| `sheet` | xlwingsのSheet（`wb.sheets["計算結果"]` など） |
| `start_cell` | 出力開始セル。例：`"B2"` |
| `values` | スカラー、1次元リスト/NumPy配列、2次元リスト/NumPy配列 |
| `orientation` | 1次元の場合の方向。既定 `"column"`、横方向なら `"row"` |

空欄 `None` は `#N/A` にせず空欄のまま書き込みます。文字列はそのまま書き込みます。
入力が3次元以上の場合や、空の配列の場合は `ValueError` となります。
変換はExcelへの出力時にのみ行い、元のNumPy配列は変更しません。

内部では通常の数値をまず範囲単位で書き込み、非有限値のセルにだけ
`=NA()` 数式を書き込みます。このため非有限値のセルが非常に多い大きな
配列ではExcelへの個別操作が増え、処理時間が長くなる可能性があります。

**グラフでの注意**：Excelのグラフは通常 `#N/A` を欠損値として扱えますが、
表示方法はグラフ種別・Excelの設定に依存するため、実機で確認してください。

## `main.py` での利用

```python
from utils.excel_io import write_excel_result

# 補間結果にも求根結果にも同じ書き込み関数を使用
write_excel_result(wb.sheets["計算結果"], "B2", aero_result["CL"])
write_excel_result(wb.sheets["交点結果"], "B2", solution.x)
write_excel_result(wb.sheets["交点結果"], "B3", solution.residual)
```

本実装は `xlwings` Classic のデスクトップExcel操作を想定しています。
テストはExcelオブジェクトを模したモデルで行い、Excel実機での表示は未検証です。
