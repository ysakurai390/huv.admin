Attribute VB_Name = "NanoActDashboardVBA"
Option Explicit

Private Const SHEET_DASHBOARD As String = "Dashboard"
Private Const SHEET_INPUT As String = "Input"
Private Const SHEET_DATA As String = "Data"
Private Const SHEET_MASTER As String = "Master"

Private Const DASHBOARD_TABLE_ROW As Long = 8
Private Const DATA_HEADER_ROW As Long = 1

Public Sub SetupWorkbook()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    EnsureSheet SHEET_DASHBOARD
    EnsureSheet SHEET_INPUT
    EnsureSheet SHEET_DATA
    EnsureSheet SHEET_MASTER

    BuildMasterSheet
    BuildDataSheet
    BuildInputSheet
    BuildDashboardSheet
    RefreshDashboard

    Worksheets(SHEET_INPUT).Activate
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox "セットアップが完了しました。Input シートから入力を始めてください。", vbInformation
End Sub

Public Sub SaveRecord()
    Dim wsInput As Worksheet
    Dim wsData As Worksheet
    Dim targetRow As Long
    Dim recordId As String

    Set wsInput = Worksheets(SHEET_INPUT)
    Set wsData = Worksheets(SHEET_DATA)

    If Trim$(wsInput.Range("B4").Value) = "" Then
        MsgBox "施設名は必須です。", vbExclamation
        Exit Sub
    End If

    recordId = Trim$(wsInput.Range("B2").Value)

    If recordId = "" Then
        recordId = CreateRecordId()
        targetRow = NextDataRow(wsData)
    Else
        targetRow = FindDataRowById(recordId)
        If targetRow = 0 Then targetRow = NextDataRow(wsData)
    End If

    wsData.Cells(targetRow, 1).Value = recordId
    wsData.Cells(targetRow, 2).Value = Trim$(wsInput.Range("B4").Value)
    wsData.Cells(targetRow, 3).Value = wsInput.Range("D4").Value
    wsData.Cells(targetRow, 4).Value = Trim$(wsInput.Range("F4").Value)
    wsData.Cells(targetRow, 5).Value = Trim$(wsInput.Range("B6").Value)
    wsData.Cells(targetRow, 6).Value = wsInput.Range("D6").Value
    wsData.Cells(targetRow, 7).Value = Trim$(wsInput.Range("F6").Value)
    wsData.Cells(targetRow, 8).Value = Trim$(wsInput.Range("B8").Value)
    wsData.Cells(targetRow, 9).Value = Trim$(wsInput.Range("D8").Value)
    wsData.Cells(targetRow, 10).Value = Trim$(wsInput.Range("F8").Value)
    wsData.Cells(targetRow, 11).Value = wsInput.Range("B10").Value
    wsData.Cells(targetRow, 12).Value = wsInput.Range("D10").Value
    wsData.Cells(targetRow, 13).Value = wsInput.Range("F10").Value
    wsData.Cells(targetRow, 14).Value = Trim$(wsInput.Range("B12").Value)
    wsData.Cells(targetRow, 15).Value = wsInput.Range("D12").Value
    wsData.Cells(targetRow, 16).Value = wsInput.Range("F12").Value
    wsData.Cells(targetRow, 17).Value = wsInput.Range("B14").Value
    wsData.Cells(targetRow, 18).Value = wsInput.Range("D14").Value
    wsData.Cells(targetRow, 19).Value = Trim$(wsInput.Range("B16").Value)
    wsData.Cells(targetRow, 20).Value = Now

    wsInput.Range("B2").Value = recordId

    RefreshDashboard
    MsgBox "保存しました。", vbInformation
End Sub

Public Sub ClearInputForm()
    Dim ws As Worksheet
    Set ws = Worksheets(SHEET_INPUT)

    ws.Range("B2").ClearContents
    ws.Range("B4,D4,F4,B6,D6,F6,B8,D8,F8,B10,D10,F10,B12,D12,F12,B14,D14,B16").ClearContents
    ws.Range("D14").ClearContents
    ws.Range("B4").Select
End Sub

Public Sub DeleteCurrentRecord()
    Dim wsInput As Worksheet
    Dim wsData As Worksheet
    Dim rowNum As Long
    Dim recordId As String

    Set wsInput = Worksheets(SHEET_INPUT)
    Set wsData = Worksheets(SHEET_DATA)
    recordId = Trim$(wsInput.Range("B2").Value)

    If recordId = "" Then
        MsgBox "削除対象がありません。", vbExclamation
        Exit Sub
    End If

    If MsgBox("このレコードを削除します。よろしいですか？", vbYesNo + vbQuestion) <> vbYes Then Exit Sub

    rowNum = FindDataRowById(recordId)
    If rowNum > 0 Then
        wsData.Rows(rowNum).Delete
    End If

    ClearInputForm
    RefreshDashboard
    MsgBox "削除しました。", vbInformation
End Sub

Public Sub RefreshDashboard()
    Dim wsDash As Worksheet
    Dim wsData As Worksheet
    Dim lastRow As Long
    Dim outRow As Long
    Dim i As Long
    Dim filterType As String
    Dim filterStage As String
    Dim filterTemp As String
    Dim stageValue As String
    Dim matchRow As Boolean
    Dim contactCount As Long
    Dim appointmentCount As Long
    Dim siteCheckCount As Long
    Dim pilotAgreeCount As Long

    Set wsDash = Worksheets(SHEET_DASHBOARD)
    Set wsData = Worksheets(SHEET_DATA)

    filterType = Trim$(wsDash.Range("B2").Value)
    filterStage = Trim$(wsDash.Range("D2").Value)
    filterTemp = Trim$(wsDash.Range("F2").Value)

    wsDash.Range("A9:K10000").ClearContents
    outRow = 9
    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row

    contactCount = 0
    appointmentCount = 0
    siteCheckCount = 0
    pilotAgreeCount = 0

    If lastRow < 2 Then GoTo UpdateKPI

    For i = 2 To lastRow
        matchRow = True

        If filterType <> "" And wsData.Cells(i, 3).Value <> filterType Then matchRow = False
        If filterStage <> "" And wsData.Cells(i, 11).Value <> filterStage Then matchRow = False
        If filterTemp <> "" And wsData.Cells(i, 12).Value <> filterTemp Then matchRow = False

        If matchRow Then
            wsDash.Cells(outRow, 1).Value = wsData.Cells(i, 1).Value
            wsDash.Cells(outRow, 2).Value = wsData.Cells(i, 2).Value
            wsDash.Cells(outRow, 3).Value = wsData.Cells(i, 3).Value
            wsDash.Cells(outRow, 4).Value = wsData.Cells(i, 11).Value
            wsDash.Cells(outRow, 5).Value = wsData.Cells(i, 12).Value
            wsDash.Cells(outRow, 6).Value = wsData.Cells(i, 5).Value
            wsDash.Cells(outRow, 7).Value = wsData.Cells(i, 13).Value
            wsDash.Cells(outRow, 8).Value = wsData.Cells(i, 15).Value
            wsDash.Cells(outRow, 9).Value = wsData.Cells(i, 16).Value
            wsDash.Cells(outRow, 10).Value = wsData.Cells(i, 17).Value
            wsDash.Cells(outRow, 11).Value = wsData.Cells(i, 18).Value
            outRow = outRow + 1
        End If

        stageValue = wsData.Cells(i, 11).Value
        If stageValue = "接触" Then contactCount = contactCount + 1
        If stageValue = "アポ" Then appointmentCount = appointmentCount + 1
        If stageValue = "現地確認" Then siteCheckCount = siteCheckCount + 1
        If stageValue = "実証合意" Then pilotAgreeCount = pilotAgreeCount + 1
    Next i

UpdateKPI:
    wsDash.Range("B5").Value = contactCount
    wsDash.Range("D5").Value = appointmentCount
    wsDash.Range("F5").Value = siteCheckCount
    wsDash.Range("H5").Value = pilotAgreeCount

    wsDash.Columns("A").Hidden = True
    wsDash.Columns("G:H").NumberFormatLocal = "yyyy/mm/dd"
    wsDash.Columns.AutoFit
End Sub

Public Sub OpenSelectedRecord()
    Dim wsDash As Worksheet
    Dim rowNum As Long
    Dim recordId As String

    Set wsDash = Worksheets(SHEET_DASHBOARD)
    rowNum = ActiveCell.Row

    If ActiveSheet.Name <> SHEET_DASHBOARD Or rowNum < 9 Then
        MsgBox "Dashboard シートの一覧行を選択してください。", vbExclamation
        Exit Sub
    End If

    recordId = Trim$(wsDash.Cells(rowNum, 1).Value)
    If recordId = "" Then
        MsgBox "有効な行を選択してください。", vbExclamation
        Exit Sub
    End If

    LoadRecordById recordId
End Sub

Public Sub OpenMapFromInput()
    Dim url As String
    url = Trim$(Worksheets(SHEET_INPUT).Range("B6").Value)

    If url = "" Then
        MsgBox "GoogleMapリンクが空です。", vbExclamation
        Exit Sub
    End If

    ThisWorkbook.FollowHyperlink url
End Sub

Public Sub ExportDataToCsv()
    Dim wsData As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim i As Long, j As Long
    Dim filePath As Variant
    Dim lineText As String
    Dim valueText As String
    Dim fileNo As Integer

    Set wsData = Worksheets(SHEET_DATA)
    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row
    lastCol = wsData.Cells(1, wsData.Columns.Count).End(xlToLeft).Column

    filePath = Application.GetSaveAsFilename(InitialFileName:="hotel_partnership_dashboard.csv", FileFilter:="CSV Files (*.csv), *.csv")
    If filePath = False Then Exit Sub

    fileNo = FreeFile
    Open CStr(filePath) For Output As #fileNo

    For i = 1 To lastRow
        lineText = ""
        For j = 1 To lastCol
            valueText = Replace(CStr(wsData.Cells(i, j).Value), Chr(34), Chr(34) & Chr(34))
            If j > 1 Then lineText = lineText & ","
            lineText = lineText & Chr(34) & valueText & Chr(34)
        Next j
        Print #fileNo, lineText
    Next i

    Close #fileNo
    MsgBox "CSVを出力しました。", vbInformation
End Sub

Public Sub LoadRecordById(ByVal recordId As String)
    Dim wsInput As Worksheet
    Dim wsData As Worksheet
    Dim rowNum As Long

    Set wsInput = Worksheets(SHEET_INPUT)
    Set wsData = Worksheets(SHEET_DATA)

    rowNum = FindDataRowById(recordId)
    If rowNum = 0 Then
        MsgBox "対象レコードが見つかりません。", vbExclamation
        Exit Sub
    End If

    wsInput.Range("B2").Value = wsData.Cells(rowNum, 1).Value
    wsInput.Range("B4").Value = wsData.Cells(rowNum, 2).Value
    wsInput.Range("D4").Value = wsData.Cells(rowNum, 3).Value
    wsInput.Range("F4").Value = wsData.Cells(rowNum, 4).Value
    wsInput.Range("B6").Value = wsData.Cells(rowNum, 5).Value
    wsInput.Range("D6").Value = wsData.Cells(rowNum, 6).Value
    wsInput.Range("F6").Value = wsData.Cells(rowNum, 7).Value
    wsInput.Range("B8").Value = wsData.Cells(rowNum, 8).Value
    wsInput.Range("D8").Value = wsData.Cells(rowNum, 9).Value
    wsInput.Range("F8").Value = wsData.Cells(rowNum, 10).Value
    wsInput.Range("B10").Value = wsData.Cells(rowNum, 11).Value
    wsInput.Range("D10").Value = wsData.Cells(rowNum, 12).Value
    wsInput.Range("F10").Value = wsData.Cells(rowNum, 13).Value
    wsInput.Range("B12").Value = wsData.Cells(rowNum, 14).Value
    wsInput.Range("D12").Value = wsData.Cells(rowNum, 15).Value
    wsInput.Range("F12").Value = wsData.Cells(rowNum, 16).Value
    wsInput.Range("B14").Value = wsData.Cells(rowNum, 17).Value
    wsInput.Range("D14").Value = wsData.Cells(rowNum, 18).Value
    wsInput.Range("B16").Value = wsData.Cells(rowNum, 19).Value

    wsInput.Activate
    wsInput.Range("B4").Select
End Sub

Private Sub BuildMasterSheet()
    Dim ws As Worksheet
    Set ws = Worksheets(SHEET_MASTER)
    ws.Cells.Clear

    ws.Range("A1").Value = "種別"
    ws.Range("A2").Resize(3, 1).Value = Application.Transpose(Array("ホテル", "旅館", "民泊"))

    ws.Range("B1").Value = "連絡手段"
    ws.Range("B2").Resize(3, 1).Value = Application.Transpose(Array("電話", "メール", "LINE"))

    ws.Range("C1").Value = "ステージ"
    ws.Range("C2").Resize(9, 1).Value = Application.Transpose(Array("未接触", "接触", "資料送付", "アポ", "現地確認", "条件調整", "実証合意", "導入", "保留・見送り"))

    ws.Range("D1").Value = "温度感"
    ws.Range("D2").Resize(3, 1).Value = Application.Transpose(Array("A", "B", "C"))

    ws.Range("E1").Value = "主要懸念"
    ws.Range("E2").Resize(6, 1).Value = Application.Transpose(Array("置き場", "充電", "安全", "運用", "料金", "その他"))

    ws.Range("F1").Value = "想定台数"
    ws.Range("F2").Value = "未定"
    ws.Range("F3:F21").Formula = "=ROW()-1"

    ws.Range("G1").Value = "置き場状況"
    ws.Range("G2").Resize(4, 1).Value = Application.Transpose(Array("屋内OK", "軒下OK", "未確認", "難しい"))

    ws.Columns.AutoFit
End Sub

Private Sub BuildDataSheet()
    Dim ws As Worksheet
    Dim headers As Variant

    Set ws = Worksheets(SHEET_DATA)
    ws.Cells.Clear

    headers = Array("RecordID", "施設名", "種別", "住所", "GoogleMapリンク", "連絡手段", "担当者名役職", "電話番号", "メールアドレス", "紹介元", "ステージ", "温度感", "最終接触日", "次アクション", "次回期限", "主要懸念", "想定台数", "置き場状況", "最新メモ", "更新日時")
    ws.Range("A1").Resize(1, UBound(headers) + 1).Value = headers
    ws.Rows(1).Font.Bold = True
    ws.Columns.AutoFit
End Sub

Private Sub BuildInputSheet()
    Dim ws As Worksheet
    Set ws = Worksheets(SHEET_INPUT)
    ws.Cells.Clear

    ws.Range("A1").Value = "宿泊施設提携進捗入力"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 16

    ws.Range("A2").Value = "RecordID"
    ws.Range("B2").Interior.Color = RGB(242, 242, 242)

    ws.Range("A4").Value = "施設名"
    ws.Range("C4").Value = "種別"
    ws.Range("E4").Value = "住所"

    ws.Range("A6").Value = "GoogleMapリンク"
    ws.Range("C6").Value = "連絡手段"
    ws.Range("E6").Value = "担当者名 / 役職"

    ws.Range("A8").Value = "電話番号"
    ws.Range("C8").Value = "メールアドレス"
    ws.Range("E8").Value = "紹介元"

    ws.Range("A10").Value = "ステージ"
    ws.Range("C10").Value = "温度感"
    ws.Range("E10").Value = "最終接触日"

    ws.Range("A12").Value = "次アクション"
    ws.Range("C12").Value = "次回期限"
    ws.Range("E12").Value = "主要懸念"

    ws.Range("A14").Value = "想定台数"
    ws.Range("C14").Value = "置き場状況"
    ws.Range("A16").Value = "最新メモ"

    ws.Range("B16:F20").Merge
    ws.Range("B16").WrapText = True
    ws.Range("B16").VerticalAlignment = xlTop

    ws.Range("A2:F20").Columns.ColumnWidth = 18
    ws.Range("B16:F20").RowHeight = 80

    SetValidation ws.Range("D4"), "='Master'!$A$2:$A$4"
    SetValidation ws.Range("D6"), "='Master'!$B$2:$B$4"
    SetValidation ws.Range("B10"), "='Master'!$C$2:$C$10"
    SetValidation ws.Range("D10"), "='Master'!$D$2:$D$4"
    SetValidation ws.Range("F12"), "='Master'!$E$2:$E$7"
    SetValidation ws.Range("B14"), "='Master'!$F$2:$F$21"
    SetValidation ws.Range("D14"), "='Master'!$G$2:$G$5"

    AddButton ws, "保存", "SaveRecord", 420, 20, 90, 28
    AddButton ws, "入力クリア", "ClearInputForm", 520, 20, 90, 28
    AddButton ws, "削除", "DeleteCurrentRecord", 620, 20, 90, 28
    AddButton ws, "GoogleMapを開く", "OpenMapFromInput", 420, 56, 140, 28
    AddButton ws, "Dashboard更新", "RefreshDashboard", 570, 56, 140, 28

    ws.Columns.AutoFit
End Sub

Private Sub BuildDashboardSheet()
    Dim ws As Worksheet
    Set ws = Worksheets(SHEET_DASHBOARD)
    ws.Cells.Clear

    ws.Range("A1").Value = "宿泊施設提携進捗ダッシュボード"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 16

    ws.Range("A2").Value = "種別フィルタ"
    ws.Range("C2").Value = "ステージフィルタ"
    ws.Range("E2").Value = "温度感フィルタ"

    SetValidation ws.Range("B2"), "='Master'!$A$2:$A$4", True
    SetValidation ws.Range("D2"), "='Master'!$C$2:$C$10", True
    SetValidation ws.Range("F2"), "='Master'!$D$2:$D$4", True

    ws.Range("A5").Value = "接触数"
    ws.Range("C5").Value = "アポ数"
    ws.Range("E5").Value = "現地確認数"
    ws.Range("G5").Value = "実証合意数"
    ws.Range("B5,D5,F5,H5").Font.Bold = True

    ws.Range("A8:K8").Value = Array("RecordID", "施設名", "種別", "ステージ", "温度感", "GoogleMapリンク", "最終接触日", "次回期限", "主要懸念", "想定台数", "置き場状況")
    ws.Rows(8).Font.Bold = True

    AddButton ws, "更新", "RefreshDashboard", 520, 8, 90, 28
    AddButton ws, "選択行を開く", "OpenSelectedRecord", 620, 8, 110, 28
    AddButton ws, "CSV出力", "ExportDataToCsv", 740, 8, 100, 28

    ws.Columns("A").Hidden = True
    ws.Columns.AutoFit
End Sub

Private Sub EnsureSheet(ByVal sheetName As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Worksheets.Add(After:=Worksheets(Worksheets.Count)).Name = sheetName
    End If
End Sub

Private Sub SetValidation(ByVal target As Range, ByVal formula1 As String, Optional ByVal allowBlankChoice As Boolean = False)
    target.Validation.Delete
    target.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=formula1
    target.Validation.IgnoreBlank = True
    target.Validation.InCellDropdown = True
    If allowBlankChoice Then target.ClearContents
End Sub

Private Sub AddButton(ByVal ws As Worksheet, ByVal caption As String, ByVal macroName As String, ByVal leftPos As Double, ByVal topPos As Double, ByVal widthVal As Double, ByVal heightVal As Double)
    Dim shp As Shape
    On Error Resume Next
    ws.Shapes(caption).Delete
    On Error GoTo 0

    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, leftPos, topPos, widthVal, heightVal)
    shp.Name = caption
    shp.TextFrame.Characters.Text = caption
    shp.OnAction = macroName
    shp.Fill.ForeColor.RGB = RGB(240, 240, 240)
    shp.Line.ForeColor.RGB = RGB(180, 180, 180)
    shp.TextFrame.HorizontalAlignment = xlHAlignCenter
    shp.TextFrame.VerticalAlignment = xlVAlignCenter
End Sub

Private Function NextDataRow(ByVal ws As Worksheet) As Long
    NextDataRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If NextDataRow < 2 Then NextDataRow = 2
End Function

Private Function FindDataRowById(ByVal recordId As String) As Long
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long

    Set ws = Worksheets(SHEET_DATA)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastRow
        If Trim$(ws.Cells(i, 1).Value) = recordId Then
            FindDataRowById = i
            Exit Function
        End If
    Next i

    FindDataRowById = 0
End Function

Private Function CreateRecordId() As String
    Randomize
    CreateRecordId = Format(Now, "yyyymmddhhnnss") & "-" & Format(Int(Rnd() * 10000), "0000")
End Function
