Attribute VB_Name = "RunMaps_without_xlwings"
Option Explicit

' Windows版Excel向け。
' Excelブックのあるフォルダに .venv と src があることを想定。
' 計算は保存済みブックを読む。結果は一時.xlsxに出力され、
' このVBAが開いているExcelブックにインポートする。
Public Sub RunMapsWithoutXlwings()
    Dim wb As Workbook
    Dim pyExe As String, scriptPath As String
    Dim resultPath As String, errorPath As String, command As String
    Dim ret As Long, message As String
    Dim outBook As Workbook
    Dim fso As Object, stream As Object
    Dim previousEvents As Boolean

    Set wb = ThisWorkbook
    If Len(wb.Path) = 0 Then
        MsgBox "先にこのExcelブックを保存してください。", vbExclamation
        Exit Sub
    End If

    pyExe = wb.Path & "\.venv\Scripts\python.exe"
    scriptPath = wb.Path & "\src\run_without_xlwings.py"
    If Len(Dir(pyExe)) = 0 Or Len(Dir(scriptPath)) = 0 Then
        MsgBox "Python環境またはスクリプトが見つかりません。" & vbCrLf & _
            pyExe & vbCrLf & scriptPath, vbExclamation
        Exit Sub
    End If

    Set fso = CreateObject("Scripting.FileSystemObject")
    resultPath = Environ$("TEMP") & "\" & fso.GetTempName & ".xlsx"
    errorPath = Left$(resultPath, Len(resultPath) - 5) & ".error.txt"
    previousEvents = Application.EnableEvents

    On Error GoTo Failure
    wb.Worksheets("00_計算管理").Range("B2").Value = "計算中"
    Application.CalculateFull
    wb.Save '未保存データおよびExcel数式の計算結果を保存済みファイルに反映する

    command = """" & pyExe & """ """ & scriptPath & """ """ & wb.FullName & """ """ & resultPath & """"
    ret = CreateObject("WScript.Shell").Run(command, 1, True) '完了まで待機
    If ret <> 0 Then
        message = "Pythonの実行に失敗しました (終了コード " & ret & ")。" & vbCrLf & _
                  "詳細ログ: " & errorPath
        Err.Raise vbObjectError + 1000, , message
    End If
    If Not fso.FileExists(resultPath) Then
        Err.Raise vbObjectError + 1001, , "Pythonの出力ファイルがありません。"
    End If

    Set outBook = Workbooks.Open(resultPath, ReadOnly:=True, UpdateLinks:=False)
    Application.EnableEvents = False
    CopyResultTable outBook.Worksheets("30_TV"), wb.Worksheets("30_TV"), "E"
    CopyResultTable outBook.Worksheets("31_Ps"), wb.Worksheets("31_Ps"), "D"
    wb.Worksheets("00_計算管理").Range("B2:B4").Formula = _
        outBook.Worksheets("00_計算管理").Range("B2:B4").Formula
    outBook.Close SaveChanges:=False
    Set outBook = Nothing
    Application.EnableEvents = previousEvents
    Application.Calculate
    wb.Save

    If fso.FileExists(resultPath) Then Kill resultPath
    If fso.FileExists(errorPath) Then Kill errorPath
    MsgBox "Python計算と結果の取り込みが完了しました。", vbInformation
    Exit Sub

Failure:
    message = Err.Description
    On Error Resume Next
    If Not outBook Is Nothing Then outBook.Close SaveChanges:=False
    Application.EnableEvents = previousEvents
    wb.Worksheets("00_計算管理").Range("B2").Value = "エラー（結果要確認）"
    wb.Worksheets("00_計算管理").Range("B4").Value = Left$(message, 300)
    wb.Save
    MsgBox "計算を完了できませんでした。" & vbCrLf & message, vbExclamation
    If Not fso Is Nothing Then
        If fso.FileExists(resultPath) Then Kill resultPath
        ' error.txt は調査用に残す。機密データがある場合は処理後に削除する。
    End If
End Sub

Private Sub CopyResultTable(ByVal src As Worksheet, ByVal dst As Worksheet, ByVal lastCol As String)
    Dim finalRow As Long
    finalRow = src.Cells(src.Rows.Count, "A").End(xlUp).Row
    dst.Range("A5:" & lastCol & "10005").ClearContents
    ' Value2でなくFormulaを転記することで、=NA() も数式として保持する。
    dst.Range("A5").Resize(finalRow - 4, src.Range("A5:" & lastCol & "5").Columns.Count).Formula = _
        src.Range("A5:" & lastCol & finalRow).Formula
End Sub
