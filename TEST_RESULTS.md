# テスト結果（列名指定可能な等値線機能 v8）

- 実行環境：Python 3.13 / NumPy 2.3.5 / pandas 2.2.3 / SciPy 1.17.0 / Plotly 6.5.2 / ContourPy 1.3.3 / openpyxl 3.1.5
- 実行：`PYTHONPATH=src python -m pytest -q --disable-warnings`
- 結果：**121 passed, 2 warnings**
- `python -m compileall -q src` ：成功

## 今回追加したテスト

- `fixed_columns=["G", "Weight"]` + `fixed_values=[1.0, 120.0]` で2条件を固定
- `fixed={"Weight":140.0, "G":2.0}` で辞書形式の固定指定
- X/Y軸列名の交換、等値線の出力列名の変更 (`level_column="ExcessPower"`)
- 固定条件を指定しきらない場合の重複格子点エラー
- 無効な列指定、固定値のNaN、引数の組み合わせ、複数固定条件での線形補間拒否
- Excelグラフ系列の構築を模擬した検証
- 実際のオフラインPlotly HTMLの生成
- 従来の `fixed_column`, `fixed_value`, `z_column` による呼び出しとの互換性
- 既存の補間・収束計算・Excel入出力・性能計算機能の回帰テスト

**未実施：** Windows版Excel実機でのVBAボタン実行・グラフCOM操作のE2E試験。

**注記：** デモのPs値は説明用の人工数値で、実機の性能データではありません。
