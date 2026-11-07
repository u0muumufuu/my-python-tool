"""Excelのボタンから実行する性能計算の入口。

VBA: RunPython "import main; main.run_maps()"
xlwingsのPYTHONPATH: このsrcフォルダを指定。

利用するExcelシート（詳細はREADME参照）:
  00_計算管理
  01_機体諸元 (B2: 質量[kg], B3: 主翼面積[m2])
  03_飛行条件 (B2: 空気密度, B3: 音速, B4: 高度,
               B5: 最小速度, B6: 最大速度, B7: 速度点数)
  10_空力DB   (A1開始、alpha, Mach, CL, CD, Cm)
  11_推進DB   (A1開始、V, Altitude, Thrust)
  30_TV       (A5から結果出力)
  31_Ps       (A5から結果出力)
"""

from __future__ import annotations

from datetime import datetime

import numpy as np

from models.aerodynamics import AerodynamicsModel
from models.propulsion import PropulsionModel
from analysis.performance_maps import calculate_tv_ps
from utils.excel_io import write_excel_result


def read_conditions(wb) -> dict:
    """Excelの単純計算はExcel側で終え、確定したセル値をPythonに渡す。"""
    airframe = wb.sheets["01_機体諸元"]
    flight = wb.sheets["03_飛行条件"]
    n = int(flight.range("B7").value)
    if not 2 <= n <= 10000:
        raise ValueError("速度の分割点数は2～10000としてください")
    v_min = float(flight.range("B5").value)
    v_max = float(flight.range("B6").value)
    if not (0 < v_min < v_max):
        raise ValueError("速度範囲は0 < 最小速度 < 最大速度を満たす必要があります")

    return dict(
        mass_kg=float(airframe.range("B2").value),
        wing_area_m2=float(airframe.range("B3").value),
        rho_kg_m3=float(flight.range("B2").value),
        speed_of_sound_m_s=float(flight.range("B3").value),
        altitude_m=float(flight.range("B4").value),
        velocities_m_s=np.linspace(v_min, v_max, n),
    )


def _write_result_table(ws, columns: list[str], df) -> None:
    """計算結果だけをA5以降に出力。既存のグラフは変更しない。"""
    # この例は最大10000点を想定。古い結果を残さないため範囲を消去。
    # グラフや入力欄はこの範囲の外に配置する。
    ws.range(f"A5:{chr(ord('A') + len(columns) - 1)}10005").clear_contents()
    ws.range("A5").value = [columns]
    write_excel_result(ws, "A6", df[columns].to_numpy(dtype=object))


def run_maps(*, update_charts: bool = False) -> None:
    """VBAからT–V・Psを計算。update_charts=Trueならグラフも同期する。"""
    import xlwings as xw  # Excelのないテスト環境でもモデルのテストが可能

    wb = xw.Book.caller()
    control = wb.sheets["00_計算管理"]
    control.range("B2").value = "計算中"
    control.range("B4").value = ""

    try:
        wb.app.calculate()  # Excel数式の更新後に読み取る
        conditions = read_conditions(wb)

        # DBから一度だけ補間器を作成し、両性能マップで共有。
        aero = AerodynamicsModel(excel_path=wb.fullname, workbook=wb)
        propulsion = PropulsionModel(excel_path=wb.fullname, workbook=wb)
        df = calculate_tv_ps(aero, propulsion, **conditions)

        # Pythonで計算を終えてから、Excelへ結果をまとめて出力。
        _write_result_table(
            wb.sheets["30_TV"],
            ["V", "T_required", "T_available", "alpha", "status"], df,
        )
        _write_result_table(
            wb.sheets["31_Ps"],
            ["V", "Ps", "alpha", "status"], df,
        )

        if update_charts:
            # ExcelネイティブXYグラフを再利用。行数の増減にも対応。
            # 使用者がグラフを書式変更しても基本の配置・軸設定を保持。
            from utils.excel_charts import update_chart_from_table
            update_chart_from_table(
                wb, chart_sheet="30_TV", chart_name="TV_chart",
                data_sheet="30_TV", header_cell="A5", row_count=len(df),
                x_column="V", y_columns=["T_required", "T_available"],
                title="T–V",
            )
            update_chart_from_table(
                wb, chart_sheet="31_Ps", chart_name="Ps_chart",
                data_sheet="31_Ps", header_cell="A5", row_count=len(df),
                x_column="V", y_columns=["Ps"], title="Ps–V",
            )

        wb.app.calculate()  # 書き込まれた値に依存するExcel数式とグラフを更新
        control.range("B2").value = "完了"
        control.range("B3").value = datetime.now()
        control.range("B4").value = f"正常 {int((df['status'] == 'OK').sum())}/{len(df)} 点"
        wb.save()
    except Exception as exc:
        # 途中で失敗した場合は、結果が前回のものか一部更新された可能性を明示。
        control.range("B2").value = "エラー（結果要確認）"
        control.range("B4").value = f"{type(exc).__name__}: {exc}"[:300]
        wb.save()
        raise


def run_maps_with_charts() -> None:
    """Excelボタン用: 数値を保存した後にグラフの系列範囲も更新する。"""
    run_maps(update_charts=True)


def run_all() -> None:
    """現段階はマップ計算のみ。今後、各プロファイルを順次追加する。"""
    run_maps()


def run() -> None:
    """v5 の main.run() を維持する互換入口（既存シート構成向け）。"""
    import importlib
    import legacy_examples
    legacy_examples.xw = importlib.import_module("xlwings")
    legacy_examples.run()


def run_intersection_example() -> None:
    """v5 の交点計算例を維持する互換入口。"""
    import importlib
    import legacy_examples
    legacy_examples.xw = importlib.import_module("xlwings")
    legacy_examples.run_intersection_example()


def run_ps_contours() -> None:
    """Excelのボタン用: 31_PsMapの長形式データからG固定のPs等値線を更新。

    31_PsMap: A1から [Altitude, Mach, G, Ps] （行順自由）
    33_Ps等値線設定: B2 = 固定値CSV（例: 1.0 / 1.0,120）
                     B3 = 等値線の数値（例: 0,10,20）
                     B5 = fixed_mode, B6 = resample
                     B7 = 固定列名CSV（例: G / G,Weight）
                     B8 = 等値線にする列名（例: Ps）
                     B9 = 横軸列名（例: Mach）
                     B10 = 縦軸列名（例: Altitude）
    32_Ps等値線データ: 等値線XY座標の自動出力専用シート
    40_グラフ: Ps_contour_chart が生成/更新される

    計算結果そのもののPsデータを変更せず可視化のみ更新する。
    """
    import pandas as pd
    import xlwings as xw
    from utils.excel_charts import update_contour_chart_from_dataframe

    wb = xw.Book.caller()
    wb.app.calculate()

    data = wb.sheets["31_PsMap"].range("A1").options(
        pd.DataFrame, index=False, expand="table"
    ).value
    config = wb.sheets["33_Ps等値線設定"]
    fixed_columns = [s.strip() for s in str(config.range("B7").value or "G").split(",")]
    fixed_values = [float(s.strip()) for s in str(config.range("B2").value).split(",")]
    level_column = str(config.range("B8").value or "Ps").strip()
    x_column = str(config.range("B9").value or "Mach").strip()
    y_column = str(config.range("B10").value or "Altitude").strip()
    levels = [float(s.strip()) for s in str(config.range("B3").value).split(",")]
    fixed_mode = str(config.range("B5").value or "exact").strip().lower()
    resample = str(config.range("B6").value or "none").strip().lower()
    config.range("B4").value = "描画処理中"
    try:
        lines = update_contour_chart_from_dataframe(
            wb, data,
            fixed_columns=fixed_columns, fixed_values=fixed_values,
            level_column=level_column, levels=levels,
            x_column=x_column, y_column=y_column,
            fixed_mode=fixed_mode,
            resample=resample,
            output_data_sheet="32_Ps等値線データ",
            chart_sheet="40_グラフ",
            chart_name="Ps_contour_chart",
        )
        n_branches = lines.groupby([f"{level_column}_level", "segment_id"]).ngroups
        config.range("B4").value = f"完了: {len(levels)}レベル / {n_branches}曲線"
        wb.app.calculate()
    except Exception as exc:
        config.range("B4").value = f"ERROR（グラフ要確認）: {type(exc).__name__}: {exc}"[:255]
        wb.save()
        raise
    wb.save()
