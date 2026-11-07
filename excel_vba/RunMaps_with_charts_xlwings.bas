Attribute VB_Name = "RunMaps_with_charts_xlwings"
Option Explicit

' xlwings Standalone VBAモジュールまたはアドインが使えるExcelで実行。
' データ計算に加えて、T-V と Ps のXY散布図の系列参照を更新します。
Public Sub RunMapsWithCharts()
    On Error GoTo Failure
    RunPython "import main; main.run_maps_with_charts()"
    MsgBox "計算とExcelグラフの更新が完了しました。", vbInformation
    Exit Sub
Failure:
    MsgBox "実行に失敗しました: " & Err.Description, vbExclamation
End Sub
