"""Excel のボタンから呼ぶサンプル。

構成:
  src/main.py
  src/utils/interpolation.py
  src/utils/solvers.py
  src/utils/excel_io.py

Excel の xlwings 設定 PYTHONPATH をこの src フォルダに設定する。
VBA: RunPython "import main; main.run()"

必要シート (例):
- 空力DB: D5 に alpha, Mach, CL, CD, Cm の見出し
           → alpha, Mach の全組み合わせを含む直交格子
- エンジンDB: A1 に RPM, Altitude, Power, FuelFlow の見出し
- 計算条件: B2=alpha, B3=Mach, B4=RPM, B5=Altitude
- 計算結果: B2:B4=CL,CD,Cm, B6:B7=Power,FuelFlow
"""

import xlwings as xw

from utils.interpolation import ExcelInterpolator
from utils.solvers import EquationSolver
from utils.excel_io import write_excel_result


def run() -> None:
    # 呼出し元ブック (未保存の Excel セル変更も読み取る)
    wb = xw.Book.caller()
    wb.app.calculate()  # 入力セルが Excel 数式なら、読み取り前に再計算

    condition = wb.sheets["計算条件"]
    alpha = float(condition.range("B2").value)
    mach = float(condition.range("B3").value)
    rpm = float(condition.range("B4").value)
    altitude = float(condition.range("B5").value)

    aero = ExcelInterpolator(
        excel_path=wb.fullname,
        sheet_name="空力DB",
        start_cell="D5",           # 省略すると A1
        input_columns=["alpha", "Mach"],
        output_columns=["CL", "CD", "Cm"],
        method="regular_grid",
        grid_method="linear",
        workbook=wb,
    )
    engine = ExcelInterpolator(
        excel_path=wb.fullname,
        sheet_name="エンジンDB",
        input_columns=["RPM", "Altitude"],
        output_columns=["Power", "FuelFlow"],
        method="rbf",
        workbook=wb,
    )

    # 補間器は一度構築し、何度 predict しても Excel の再読み込みは発生しない
    aero_result = aero.predict(alpha=alpha, Mach=mach)
    engine_result = engine.predict(RPM=rpm, Altitude=altitude)

    result = wb.sheets["計算結果"]
    write_excel_result(result, "B2", [
        aero_result["CL"], aero_result["CD"], aero_result["Cm"]
    ])
    write_excel_result(result, "B6", [
        engine_result["Power"], engine_result["FuelFlow"]
    ])

    wb.app.calculate()  # これらの出力値を参照する Excel 数式を更新
    wb.save()           # グラフや値を含めてブックに保存



def run_intersection_example() -> None:
    """同じExcel表に記録した2曲線 f(x), g(x) の交点 f-g=0 を求める例。

    必要シート:
    - 「交点DB」: A1から x, f, g のヘッダーとデータ
      例: x=0,1,2,3,4 / f=2*x+1 / g=0.5*x+4
    - 「交点結果」: B2に解 x、B3に f-g の残差、B4に手法を出力

    VBA から: RunPython "import main; main.run_intersection_example()"
    """
    wb = xw.Book.caller()
    wb.app.calculate()

    f_interp = ExcelInterpolator(
        excel_path=wb.fullname,
        sheet_name="交点DB",
        input_columns=["x"],
        output_columns="f",
        method="pchip",
        workbook=wb,
    )
    g_interp = ExcelInterpolator(
        excel_path=wb.fullname,
        sheet_name="交点DB",
        input_columns=["x"],
        output_columns="g",
        method="pchip",
        workbook=wb,
    )

    def difference(x: float) -> float:
        return f_interp.predict(x=x)["f"] - g_interp.predict(x=x)["g"]

    # 両補間器の定義域内で符号が反転する区間を指定
    solution = EquationSolver(residual_tol=1e-8).solve(
        difference, bracket=(0.0, 4.0)
    )

    output = wb.sheets["交点結果"]
    # 求根以外の計算結果でも、NaN / ±inf は一律 Excel の #N/A にする。
    write_excel_result(output, "B2", solution.x)
    write_excel_result(output, "B3", solution.residual)
    output.range("B4").value = solution.method
    output.range("B5").value = "OK" if solution.converged else solution.message
    wb.app.calculate()
    wb.save()
