# 中立点算出プログラム（マッハ数別CSV対応）

`neutral_point_from_csv.py` は、**ファイル名に `Mach0.3` のようにマッハ数が記されたCSV**から、空力係数を読み込んで中立点を計算します。設定変更はPythonファイル冒頭の「1. 設定」だけで行います。

## フォルダ構成

```text
neutral_point_from_csv.py
neutral_point_README.md
aero_csv/
    aircraft_Mach0.2_aero.csv
    aircraft_Mach0.3_aero.csv
    aircraft_Mach0.4_aero.csv
```

`INPUT_DIR` に指定したフォルダ内にある、ファイル名に `Mach` と数値を含むCSVを一括処理します。マッハ数は**ファイル名から抽出**し、CSV内に `Mach` 列は不要です。`Mach0.3` / `Mach1.2` / `mach0.85` などが使えます。

## 入力CSV形式

例えば `aircraft_Mach0.3_aero.csv` は次のようにします。

```csv
この1行目は不要なので読み飛ばす
alpha,CA,CN,Cm
3,0.026,0.395,0.0242
4,0.024,0.480,0.0156
5,0.022,0.565,0.0070
6,0.020,0.650,-0.0016
7,0.018,0.735,-0.0102
```

- **1行目を必ず読み飛ばし、2行目を列名として読み込みます**（`pd.read_csv(..., skiprows=1)`）。
- 迎角列は **`alpha` のみ**です。単位は**度**。`alpha_deg` や `AoA` は使用できません。
- 必須の空力係数は `Cm` と `CA,CN`（または `CX,CZ`）。列名は `Cm`、`CA`、`CN`（または `CX`、`CZ`）としてください。
- `CA=-CX`（機尾方向正）、`CN=-CZ`（上方向正）。`Cm` は機首上げ正。
- `Re` 列があれば、同じMach・Reの各グループごとに計算します。使用する場合は全CSVに`Re`列が必要です。
- 同じMach・Re・迎角の組合せが重複している場合はエラーにします。
- 各Mach・Reに**迎角3点以上**が必要です。
- 入力はマッハ数ごとに別ファイルなので、同一フォルダ内に配置してください。

## 設定項目

Pythonファイル冒頭の「1. 設定」を編集します。

```python
INPUT_DIR = Path(__file__).resolve().parent / "aero_csv"

X_REF = 0.25
Y_REF = 0.0
Z_REF = 0.0

Z_NP = 0.05
Y_NP = Y_REF

X_MAC_LE = 0.0
C_MAC = 1.0
C_CM = C_MAC

FIT_ALPHA_MIN_DEG = -2.0
FIT_ALPHA_MAX_DEG = 8.0
```

- **位置座標** [m]：機首原点、x=機尾正、y=右翼正、z=上正。
- `X_REF,Y_REF,Z_REF`：入力`Cm`のモーメント基準点。
- `Z_NP`：中立点を評価する**固定高さ**。`Y_NP`：出力する左右位置（`Cm`のみからは決まらない）。
- `X_MAC_LE`：MAC前縁のx座標。`C_MAC`：MAC長。`C_CM`：`Cm`を無次元化した基準長。
- `FIT_ALPHA_MIN_DEG`～`FIT_ALPHA_MAX_DEG`：代表中立点の直線近似に使う迎角範囲。**主翼・尾翼がほぼ線形に応答する範囲**を機体に合わせて設定してください。各Machで最低3点必要です。

## 出力ファイル

1. **`neutral_point_result.csv`（局所中立点）**：各Mach・迎角の点で、`CA,CN,Cm`を迎角で微分し中立点を計算します。内部点は中心差分、両端は3点を使う2次精度の前進・後退差分（`np.gradient(edge_order=2)`）。
2. **`neutral_point_representative.csv`（代表中立点）**：設定した迎角範囲の`CA,CN,Cm`をそれぞれ一次回帰（`np.polyfit`）し、各Machに1つの中立点を計算します。

どちらにも、`CA_alpha_per_rad,CN_alpha_per_rad,Cm_alpha_per_rad,x_np_m,y_np_m,z_np_m,h_np,h_np_percent` を記録します。局所結果には元のファイル名`source_file`も入ります。

算出式は

\[
x_{NP}=x_{ref}+\frac{(z_{NP}-z_{ref})C_{A\alpha}-C_{CM}C_{m\alpha}}{C_{N\alpha}}
\]

\[
h_{NP}=\frac{x_{NP}-x_{MAC,LE}}{C_{MAC}},\qquad \%MAC=100h_{NP}
\]

です。`C_Nα`がゼロに近い場合、計算結果はNaNになります。

## 実行

```bash
pip install numpy pandas
python neutral_point_from_csv.py
```

単体テストは `python test_neutral_point.py` で実行できます。

## 注意事項

- 通常の設計用の「中立点」は、多くの場合、**あるマッハ数・形態の線形迎角範囲に対して1つ**設定します。本プログラムの代表中立点がそれに相当します。
- 高迎角・失速付近では微係数が変化するため、局所中立点を併せて確認してください。
- `Cm` の基準点、`Cm` の基準長、`CA/CN` の符号、舵角・形態・推進状態の固定条件を揃えてください。
- 代表中立点を求める際に、失速によって空力係数の勾配が変わる領域を不用意に含めないでください。

**参考文献**：Etkin & Reid, *Dynamics of Flight: Stability and Control*, 3rd ed., Chapter 2, §2.3（書籍pp.29–31、添付PDF pp.41–43）。中立点の定義と線形近似の考え方を参照しています。
