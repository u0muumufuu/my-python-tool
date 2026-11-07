# テスト結果（Ps等値線機能 v7）

- 環境：Python 3.13, NumPy 2.3.5, pandas 2.2.3, SciPy 1.17.0, Plotly 6.5.2, ContourPy 1.3.3, openpyxl 3.1.5
- 実行コマンド：`pytest -q --disable-warnings`
- 結果：**108 passed, 2 warnings**
- 構文チェック：`python -m compileall -q src tests examples` 成功

主な追加検証：
- 順序が混在した長形式 `Altitude, Mach, G, Ps` の断面抽出
- Ps=固定値に等しい生データがなくても、等値線座標を抽出
- G値の線形補間（ただしG方向の外挿は拒否）
- 同一点の重複検出と欠損のマスク
- 不規則格子の線形再格子化と凸包外のNaN
- 分離枝の検出と、Excelの枝ごとXY系列作成
- Excelグラフの更新時に、系列数が減る場合の余剰系列削除（模擬COM）
- オフラインPlotly HTMLの実生成・レベル表示
- 以前の補間・収束計算・Excel入出力の回帰テスト

**未実施：** Windows版Excel実機とVBAボタンのE2E試験。グラフ作成・更新は模擬COMで検証。
**留意：** デモ数値は航空機実測・推算データではなく説明用の人工データ。
