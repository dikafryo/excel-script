Attribute VB_Name = "modMergeSheets"
Sub MergeSheetsToCollection(control As IRibbonControl)
    Dim sources As Collection, ws As Worksheet, item As Object, Target As Worksheet
    Dim wb As Workbook, answer As VbMsgBoxResult
    Set wb = ActiveWorkbook
    Set sources = New Collection
    If ActiveWindow.selectedSheets.Count = 1 Then
        answer = UI_ChoiceToMsg(UI_Choose("시트 병합", "시트를 1개만 선택했습니다. 이 통합 문서의 모든 시트를 합칠까요?" & vbCrLf & "(기존 시트모음·숨김 백업 시트는 제외)", Array("모든 시트 합치기", "취소")))
        If answer <> vbYes Then Exit Sub
        For Each ws In wb.Worksheets
            If Left$(ws.name, 4) <> "시트모음" And ws.Visible = xlSheetVisible Then sources.Add ws
        Next ws
    Else
        For Each item In ActiveWindow.selectedSheets
            If TypeName(item) = "Worksheet" Then sources.Add item
        Next item
    End If
    If sources.Count = 0 Then
        MsgBox "합칠 시트가 없습니다.", vbExclamation, "시트 병합"
        Exit Sub
    End If
    sources(1).Select   ' 여러 시트 그룹 선택 해제 후 새 시트 추가
    Set Target = MergeCollectionCore(wb, sources)
    Application.CutCopyMode = False
    Target.Activate
    MsgBox "시트 병합이 완료되었습니다: " & Target.name, vbInformation
End Sub

Public Function MergeCollectionCore(ByVal wb As Workbook, ByVal sources As Collection) As Worksheet
    Dim Target As Worksheet, item As Worksheet, existing As Worksheet
    Dim title As String, n As Long, targetRow As Long
    title = "시트모음"
    Do
        Set existing = Nothing
        On Error Resume Next
        Set existing = wb.Worksheets(title)
        On Error GoTo 0
        If existing Is Nothing Then Exit Do
        n = n + 1: title = "시트모음_" & n
    Loop
    Set Target = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    Target.name = title: targetRow = 1
    For Each item In sources
        CopySheetContent item, Target, targetRow
    Next item
    Set MergeCollectionCore = Target
End Function

'=================================================================
' 개별 시트에서 값을 복사하는 보조 프로시저
Sub CopySheetContent(wsSource As Worksheet, wsTarget As Worksheet, ByRef targetRow As Long)
    Dim found As Range, rng As Range, lastRow As Long, lastCol As Long
    Set found = wsSource.Cells.Find(what:="*", After:=wsSource.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False, SearchFormat:=False)
    If found Is Nothing Then Exit Sub
    lastRow = found.row
    Set found = wsSource.Cells.Find(what:="*", After:=wsSource.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False, SearchFormat:=False)
    lastCol = found.Column
    If targetRow + lastRow > wsTarget.rows.Count Then Err.Raise vbObjectError + 220, , "합친 결과가 시트 행 한도를 넘습니다."
    Set rng = wsSource.Range(wsSource.Cells(1, 1), wsSource.Cells(lastRow, lastCol))
    wsTarget.Cells(targetRow, 1).value = "시트명: " & wsSource.name
    wsTarget.Cells(targetRow, 1).Font.Bold = True
    targetRow = targetRow + 1
    rng.Copy wsTarget.Cells(targetRow, 1)
    targetRow = targetRow + rng.rows.Count + 2
End Sub

