# 検証結果（v5）

検証環境：Python 3.13 / NumPy 2.3.5 / SciPy 1.17.0 / pandas 2.2.3 / openpyxl 3.1.5 / pytest 9.0.2

実行コマンド：

```bash
python -m pytest -q tests/
python -m doctest src/utils/solvers.py
python -m compileall -q src tests
```

結果：**66 passed**。doctest および構文チェックも正常終了。

v5 で追加・変更した主な検証項目：
- 2つ以上の出力列で1列のみが Excel の `=NA()` / `#N/A` / 空欄 / `np.nan` になった場合、他の列のデータは保持される。
- RBF、規則格子（欠損列のみ散在点に切替）、線形散在点、最近傍、1次元PCHIPの欠損値処理。
- `.xlsm` に `=NA()` を保存したセルを `data_only=True` で読み込む場合の対応。
- 有効点のない出力／点数が不足する出力は `np.nan` を返し、他の出力は正常に評価できる。
- 入力座標に欠損がある行は全出力で除外する。
- 不正な文字列や入力格子点そのものが欠けた場合のエラー。
- Excelブックを模擬した未保存データと `reload()` の動作。
- v4 の収束計算、Excel 出力、`main.py` とその他の回帰テスト。

注意：**実際のデスクトップ版Excelでのxlwings/VBA呼び出しは未検証**です（Excel実機を使用できない環境のため）。
