# Test Results

実行日時: 2026-11-08

## Command

```bash
pytest -q
```

## Result

```text
124 passed, 2 warnings in 5.47s
```

## Notes

- 追加した Plotly 補間レポート機能について、以下を検証済みです。
  - 1 HTML に複数サブプロットが生成されること
  - 各サブプロットの x 軸に range slider が付与されること
  - `x_range` による外挿範囲の可視化が反映されること
  - `series_values` の明示指定と自動取得の両方が動作すること
- 既存機能（補間、収束計算、Excel グラフ、等値線、HTML 出力）の回帰テストも通過しています。
