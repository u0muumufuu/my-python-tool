Attribute VB_Name = "RunMaps_xlwings"
Option Explicit

' xlwingsアドインまたはStandalone VBAモジュールが設定されている場合のみ使用。
Public Sub RunMapsWithXlwings()
    On Error GoTo Failure
    RunPython "import main; main.run_maps()"
    MsgBox "計算が完了しました。", vbInformation
    Exit Sub
Failure:
    MsgBox "実行に失敗しました: " & Err.Description, vbExclamation
End Sub
