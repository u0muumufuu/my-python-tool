"""xlwings を使う Excel 出力用の共通処理。

NaN または無限大を Excel のエラー値 ``#N/A`` として書くため、
該当セルには文字列ではなく数式 ``=NA()`` を設定する。

一般的な単一セル、および NumPy / Python の1次元・2次元配列に対応する。
Excel の Worksheet（``wb.sheets[... ]``）を ``sheet`` に渡す。
"""

from __future__ import annotations

import numpy as np


def _is_nonfinite_number(value: object) -> bool:
    """NaN / ±inf の数値かどうかを判定（文字列や None は対象外）。"""
    if isinstance(value, (str, bytes)) or value is None:
        return False
    try:
        result = np.isfinite(value)
        return bool(not result)  # scalar number only; arrays are handled outside
    except (TypeError, ValueError):
        return False


def write_excel_result(sheet, start_cell: str, values, *, orientation: str = "column") -> None:
    """計算結果をExcelへ出力し、NaN / ±inf を ``=NA()`` にする。

    Parameters
    ----------
    sheet:
        xlwings の Sheet オブジェクト。
    start_cell:
        書き込み先の左上セル（例: ``'B2'``）。
    values:
        数値・文字列等の単一値、1次元/2次元のリスト、NumPy配列。
        NaN / 無限大以外の値はそのまま書き込む。
    orientation:
        1次元配列の出力方向。``'column'``（既定）または ``'row'``。
        2次元配列には影響しない。

    Notes
    -----
    複数セルでは、まず非有限値を空白としてまとめて書き込み、該当セルの
    ``.formula`` に ``=NA()`` を設定する。これにより実際のExcelエラー値となる。
    置換はExcel出力時にだけ行い、入力データや元のNumPy配列は変更しない。
    """
    if orientation not in ("row", "column"):
        raise ValueError("orientation は 'row' または 'column' を指定してください。")

    try:
        data = np.asarray(values, dtype=object)
    except ValueError as exc:
        raise ValueError("values は長方形の1次元または2次元配列にしてください。") from exc

    if data.ndim > 2:
        raise ValueError("values は単一値、1次元、または2次元のデータのみ対応します。")

    target = sheet.range(start_cell)
    if data.ndim == 0:
        item = data.item()
        if _is_nonfinite_number(item):
            target.formula = "=NA()"
        else:
            target.value = item
        return

    if not data.size:
        raise ValueError("values が空です。")

    if data.ndim == 1:
        data = data.reshape((-1, 1) if orientation == "column" else (1, -1))

    clean = data.tolist()
    missing_positions = []
    for r, row in enumerate(clean):
        for c, item in enumerate(row):
            if _is_nonfinite_number(item):
                clean[r][c] = None
                missing_positions.append((r, c))

    target.value = clean
    for r, c in missing_positions:
        target.offset(r, c).formula = "=NA()"
