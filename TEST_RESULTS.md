# 検証結果（v4）

- SciPy補間器・収束計算：従来テスト継続
- Excel数値出力共通関数：NaN・Inf・単一値・1次元・2次元・文字列・None・入力不備の模擬テスト
- `main.py` の交点計算・未収束時の `=NA()` 書き込み・保存呼び出しの模擬テスト

実行コマンド:

```bash
python -m pytest -q tests/
python -m doctest src/utils/solvers.py
python -m compileall -q src tests
```

Excel実機でのVBAボタン操作・グラフへの `#N/A` の反映は未検証です。
