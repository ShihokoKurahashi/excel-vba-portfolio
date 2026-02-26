Attribute VB_Name = "Module1"
Option Explicit

'=========================
' 設定保持用
'=========================
Private Type AppConfig
    ClientCode As String
    ClientName As String
    RecordCount As Long
    OutputColStart As Long
    OutputColEnd As Long
End Type

Private gConfig As AppConfig
Private gRunDateTime As String

'=========================
' メイン処理
'=========================
Sub Main()

    If GetErrorFlag <> 0 Then
        MsgBox "入力値エラーのため処理を中断します。", vbCritical
        Exit Sub
    End If

    gRunDateTime = Format(Now, "yyyyMMddHHmmss")

    WriteProcessTime 1
    LoadConfig
    GenerateTemplateRows
    ExportCSV
    WriteProcessTime 2

    MsgBox "処理完了"

End Sub

'=========================
' 設定読込
'=========================
Private Sub LoadConfig()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Config")

    With gConfig
        .RecordCount = ws.Range("A2").Value
        .ClientName = ws.Range("B2").Value
        .ClientCode = ws.Range("C2").Value
        .OutputColStart = ws.Range("E2").Value
        .OutputColEnd = ws.Range("F2").Value
    End With

End Sub

Private Function GetErrorFlag() As Long
    GetErrorFlag = ThisWorkbook.Worksheets("Config").Range("A9").Value
End Function

Private Sub WriteProcessTime(ByVal ptn As Long)

    With ThisWorkbook.Worksheets("Config")
        If ptn = 1 Then
            .Range("A6").Value = Now
        Else
            .Range("B6").Value = Now
        End If
    End With

End Sub

'=========================
' データ削除
'=========================
Private Sub ClearDataFast(ws As Worksheet, checkCol As Long)

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, checkCol).End(xlUp).Row

    If lastRow >= 3 Then
        ws.Rows("3:" & lastRow).Delete
    End If

End Sub

'=========================
' テンプレート行複製
'=========================
Private Sub GenerateTemplateRows()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Data")

    ClearDataFast ws, 1

    If gConfig.RecordCount <= 1 Then Exit Sub

    ws.Rows(2).Copy
    ws.Rows(3).Resize(gConfig.RecordCount - 1).Insert Shift:=xlDown

End Sub

'=========================
' CSV出力
'=========================
Private Sub ExportCSV()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Data")

    Dim basePath As String
    Dim outputFolder As String
    Dim csvFile As String

    basePath = ThisWorkbook.path & "\" & _
               Format(Date, "yymmdd") & "_Output_" & gConfig.ClientName

    CreateFolderIfNotExists basePath

    outputFolder = basePath & "\CSV"
    CreateFolderIfNotExists outputFolder

    csvFile = outputFolder & "\" & _
              gConfig.ClientCode & "_DATA_" & gRunDateTime & ".csv"

    Dim fno As Integer
    fno = FreeFile

    Open csvFile For Output As #fno

    Dim i As Long
    Dim iCol As Long
    Dim lineText As String

    i = 2

    Do While ws.Cells(i, 1).Value <> ""

        lineText = ""

        For iCol = gConfig.OutputColStart To gConfig.OutputColEnd
            lineText = lineText & ws.Cells(i, iCol).Value & ","
        Next iCol

        Print #fno, Left(lineText, Len(lineText) - 1)

        i = i + 1

    Loop

    Close #fno

End Sub

'=========================
' フォルダ作成共通
'=========================
Private Sub CreateFolderIfNotExists(path As String)

    If Dir(path, vbDirectory) = "" Then
        MkDir path
    End If

End Sub

