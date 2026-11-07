"""xlwingsなしでExcelからPython計算を実行するための別エントリーポイント。

Windows VBA -> WScript.Shell.Run() -> このスクリプト

  python src/run_without_xlwings.py INPUT.xlsm OUTPUT.xlsx

入力は保存済みINPUT.xlsmをopenpyxlで*読み込みのみ*とし、別ファイルに
計算結果を書き出す。開いたままの元ブックをPythonから上書きしない。
VBAがOUTPUT.xlsxの結果を開いて元ブックへ書き戻す。
"""

from __future__ import annotations

import os
import sys
import tempfile
import traceback
from datetime import datetime
from pathlib import Path

import numpy as np
from openpyxl import Workbook, load_workbook

from models.aerodynamics import AerodynamicsModel
from models.propulsion import PropulsionModel
from analysis.performance_maps import calculate_tv_ps


def read_saved_conditions(path: str | Path) -> dict:
    """Excelが保存・再計算した数式キャッシュを含む、保存済み入力を取得。"""
    book = load_workbook(path, data_only=True, read_only=True)
    try:
        airframe = book["01_機体諸元"]
        flight = book["03_飛行条件"]
        n = int(flight["B7"].value)
        if not 2 <= n <= 10000:
            raise ValueError("速度の分割点数は2～10000としてください")
        v_min = float(flight["B5"].value)
        v_max = float(flight["B6"].value)
        if not 0 < v_min < v_max:
            raise ValueError("速度範囲は0 < 最小速度 < 最大速度を満たす必要があります")
        return dict(
            mass_kg=float(airframe["B2"].value),
            wing_area_m2=float(airframe["B3"].value),
            rho_kg_m3=float(flight["B2"].value),
            speed_of_sound_m_s=float(flight["B3"].value),
            altitude_m=float(flight["B4"].value),
            velocities_m_s=np.linspace(v_min, v_max, n),
        )
    finally:
        book.close()


def _excel_cell_value(value):
    """有限数以外の数値はExcelの #N/A 数式にする。"""
    if value is None:
        return None
    if isinstance(value, (int, float, np.number)) and not isinstance(value, (bool, np.bool_)):
        return float(value) if np.isfinite(value) else "=NA()"
    return value


def _write_table(book: Workbook, sheet: str, columns: list[str], df) -> None:
    ws = book.create_sheet(sheet)
    # ヘッダーを元のmain.pyと同じA5に配置
    for col_idx, name in enumerate(columns, 1):
        ws.cell(5, col_idx, name)
    for row_idx, record in enumerate(df[columns].itertuples(index=False, name=None), start=6):
        for col_idx, value in enumerate(record, 1):
            ws.cell(row_idx, col_idx, _excel_cell_value(value))


def run(input_excel: str | Path, output_excel: str | Path) -> None:
    input_excel = Path(input_excel).expanduser().resolve()
    output_excel = Path(output_excel).expanduser().resolve()
    if not input_excel.is_file():
        raise FileNotFoundError(input_excel)
    if input_excel == output_excel:
        raise ValueError("元のExcelブックに直接書き戻すことは禁止しています")
    if input_excel.suffix.lower() not in {".xlsx", ".xlsm"}:
        raise ValueError("入力は .xlsx または .xlsm を指定してください")
    if output_excel.suffix.lower() != ".xlsx":
        raise ValueError("出力は .xlsx を指定してください")

    conditions = read_saved_conditions(input_excel)
    aero = AerodynamicsModel(excel_path=input_excel)
    propulsion = PropulsionModel(excel_path=input_excel)
    df = calculate_tv_ps(aero, propulsion, **conditions)

    book = Workbook()
    book.remove(book.active)
    _write_table(book, "30_TV", ["V", "T_required", "T_available", "alpha", "status"], df)
    _write_table(book, "31_Ps", ["V", "Ps", "alpha", "status"], df)
    ws = book.create_sheet("00_計算管理")
    ws["B2"] = "完了"
    ws["B3"] = datetime.now()
    ws["B4"] = f"正常 {int((df['status'] == 'OK').sum())}/{len(df)} 点"

    output_excel.parent.mkdir(parents=True, exist_ok=True)
    # 完全に書き出せてから最終ファイルへ置き換える
    fd, filename = tempfile.mkstemp(dir=output_excel.parent, suffix=".xlsx")
    os.close(fd)
    try:
        book.save(filename)
        os.replace(filename, output_excel)
    finally:
        if os.path.exists(filename):
            os.unlink(filename)


def main(argv: list[str] | None = None) -> int:
    argv = sys.argv[1:] if argv is None else argv
    if len(argv) != 2:
        print("usage: python run_without_xlwings.py INPUT.xlsm OUTPUT.xlsx", file=sys.stderr)
        return 2
    input_path, output_path = argv
    error_path = Path(output_path).with_suffix(".error.txt")
    try:
        run(input_path, output_path)
        error_path.unlink(missing_ok=True)
        return 0
    except Exception:
        error_path.write_text(traceback.format_exc(), encoding="utf-8")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
