# 中立点計算プログラム（Excel出力版）

## 実行環境
WindowsまたはmacOSのデスクトップ版Microsoft ExcelとPythonが必要です。

```bash
pip install numpy pandas xlwings
python neutral_point_from_csv.py
```

設定はPythonファイル冒頭の「1. 設定」に集約しています。

添付の共通関数 `uitls/excel_io.py` を使用します。フォルダ名は指定どおり `uitls` としています。実際のフォルダ名が `utils` の場合は、Python本体のimport文を `from utils.excel_io import write_excel_result` に変更してください。共通関数の本体は添付されたものから変更していません。

## 入力
`aero_csv` フォルダにMach別のCSVを保存します。
ファイル名の例：`aircraft_Mach0.3_aero.csv`。
1行目を読み飛ばし、2行目をヘッダーとして読み込みます。

```csv
任意の説明文
alpha,CA,CN,Cm
3,0.026,0.395,0.0242
4,0.024,0.480,0.0156
5,0.022,0.565,0.0070
```

- `alpha` は度。`CA,CN` の代わりに `CX,CZ` を入力した場合、符号反転します。
- 座標は機首原点、x機尾方向正、y右翼方向正、z上方向正。CAは機尾方向正、CNは上方向正、Cmは機首上げ正。
- 同じMach・迎角の重複はエラー。各Machに最低3迎角点が必要です。
- Reynolds数によるグループ化は行いません。同一形態・舵角・その他の解析条件のデータを入力してください。

## 設定
`INPUT_DIR`, `OUTPUT_EXCEL` に入力フォルダと出力ブックを指定します。
`SHEET_NAMES` は既存ブックの出力先シート名に合わせて変更可能です。
`X_REF, Y_REF, Z_REF` はモーメント基準点、`Z_NP` は評価高さ、`X_MAC_25` は25% MAC位置、`C_MAC` はMAC長、`C_CM` はピッチモーメント係数の基準長、`L_FUSELAGE` は胴体長です（長さの単位はm）。初期値は架空のサンプル用です。
`FIT_ALPHA_MIN_DEG`, `FIT_ALPHA_MAX_DEG` は代表中立点の直線近似範囲（度）です。

## 計算
```
x_NP = X_REF + ((Z_NP-Z_REF)*CA_alpha-C_CM*Cm_alpha)/CN_alpha
h_NP = 0.25 + (x_NP-X_MAC_25)/C_MAC
x_NP_over_L = x_NP/L_FUSELAGE
```
局所微係数は `np.gradient(..., edge_order=2)`、代表微係数は指定範囲での `np.polyfit(..., 1)` による最小二乗直線の傾きです。CN_alphaがしきい値未満の絶対値の場合は算出不能とし、Excelでは数式 `=NA()` を設定して `#N/A` と表示します。

## Excel出力
| シート | 内容 |
|---|---|
| x_NP | 前後座標 [m] |
| h_NP | MAC基準の無次元位置 [-] |
| x_NP_over_L | 胴体長基準の無次元位置 [-] |
| NP_representative | Machごとの代表中立点3種類、%MAC、近似条件、微係数 |

最初の3シートはA列がalpha（度）、B列以降がMach別データです。迎角やMachは昇順。欠測は補間せず `=NA()` を設定します。NaNと±infは4シートすべてで同じ扱いです。入力CSV自体の欠損・無限大は従来どおりエラーとします。ExcelでXY散布図を作成してください。

同名ブックがあればxlwingsで開くか接続し、同名シートを更新します。存在しないシートは末尾に追加します。既存シートの順序は変更しません。対象シートのA1を含む連続データ領域（current_region）の内容を消去してから新データを書き込みます。この範囲は出力専用とし、手入力の表や数式を隣接させないでください。シート・書式・グラフ自体は削除しませんが、グラフの参照範囲は自動変更しません。行数・列数が変化する場合は確認してください。既存ファイルは上書き保存し、Excelは開いたままにします。開いているブックの他の未保存変更も保存されるため注意してください。大切なブックはバックアップしてから実行してください。

## テスト
`python test_neutral_point.py` で、数値計算・CSV読み込み・表変換・新規/既存Excelの呼び出しを確認できます。9項目が合格済みです。ただし、作成環境にはExcelとxlwingsがないため、Excel操作は模擬オブジェクトによる検証であり、実際のExcelでの保存・更新は未検証です。

xlwingsの参照： https://docs.xlwings.org/en/stable/api/book.html

## 共通Excel書き込み関数の注意
欠損セルごとに数式を設定するため、大量の欠損があると遅くなることがあります。また、共通関数単体ではNoneやpandas.NAはNA()になりません。今回の呼び出しではDataFrameの欠損をnp.nanに統一して渡しています。
