Attribute VB_Name = "modSchoolSheet"
Option Explicit
' =====================================================================================
'  modSchoolSheet : 초·중·고 현황시트 생성 (개별 업데이트용, 애플 스타일 표)
'   - 시트명: 초등학교현황 / 중학교현황 / 고등학교현황
'   - 열: 연번 | 학교명 | 교육지원청 | 설립구분 | 약칭 | 비고
'   - 이 시트가 있으면 학교명 정리·지원청 표시 때 내장 데이터보다 우선 사용된다(학교명·교육지원청·설립구분 열을 읽음).
' =====================================================================================
Private Const CLR_TEXT As Long = &H1E1C1C      ' (28,28,30)
Private Const CLR_SUB As Long = &H8E8E8E       ' (142,142,142)
Private Const CLR_LINE As Long = &HDCD6D6      ' (214,214,220)
Private Const CLR_HEAD As Long = &HF7F2F2      ' (242,242,247) systemGray6
Private Const CLR_ACCENT As Long = &HFF7A00    ' (0,122,255)
Private Const HDR_ROW As Long = 4
Public SS_LastStep As String

' 급별 현황시트 생성. 반환: 만들어진(또는 유지된) 시트, 취소면 Nothing
Public Function SS_Create(ByVal lv As Long, Optional ByVal askIfExists As Boolean = True) As Worksheet
    Dim wb As Workbook, ws As Worksheet, nm As String, rows As Variant, n As Long, i As Long, choice As Long
    Set wb = ActiveWorkbook
    nm = SD_SheetName(lv)
    On Error Resume Next
    Set ws = wb.Worksheets(nm)
    On Error GoTo 0
    If Not ws Is Nothing Then
        If askIfExists Then
            choice = UI_Choose(nm & " 시트가 이미 있습니다", "기존 시트를 어떻게 할까요?" & vbCrLf & "(새로 만들면 내장 데이터로 다시 채워지고, 직접 수정한 내용은 사라집니다)", _
                               Array("새로 만들기 (내장 데이터로 덮어쓰기)", "기존 시트 그대로 두기", "취소"))
        Else
            choice = 2
        End If
        Select Case choice
            Case 1
                Application.DisplayAlerts = False
                ws.Delete
                Application.DisplayAlerts = True
                Set ws = Nothing
            Case 2
                Set SS_Create = ws
                Exit Function
            Case Else
                Exit Function
        End Select
    End If

    SS_LastStep = "rows"
    rows = SD_EmbeddedRows(lv)
    n = UBound(rows, 1)
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.name = nm

    SS_LastStep = "fill"
    ' 제목 / 안내
    With ws
        .Range("A1").value = SD_LevelName(lv) & " 현황"
        .Range("A2").value = "기준 " & SD_DATA_DATE & " · " & n & "교 · 학교명 정리(리본 ▶ 학교명·지원청) 실행 시 이 시트가 내장 데이터보다 우선 적용됩니다. 학교명·교육지원청·설립구분 열을 직접 수정해 두세요."
        .Range("A" & HDR_ROW & ":F" & HDR_ROW).value = Array("연번", "학교명", "교육지원청", "설립구분", "약칭", "비고")
        For i = 1 To n
            .Cells(HDR_ROW + i, 1).value = i
            .Cells(HDR_ROW + i, 2).value = rows(i, 1)
            .Cells(HDR_ROW + i, 3).value = rows(i, 2)
            .Cells(HDR_ROW + i, 4).value = rows(i, 3)
            .Cells(HDR_ROW + i, 5).value = SM_Abbrev(CStr(rows(i, 1)))
            If InStr(rows(i, 2), "·") > 0 Then .Cells(HDR_ROW + i, 6).value = "동명 학교 2곳(지원청 확인)"
            If InStr(rows(i, 1), "(폐교)") > 0 Then .Cells(HDR_ROW + i, 6).value = "폐교"
            If InStr(rows(i, 1), "가칭") > 0 Then .Cells(HDR_ROW + i, 6).value = "개교 예정(가칭)"
            If InStr(rows(i, 1), "분교장") > 0 Then .Cells(HDR_ROW + i, 6).value = "분교장"
        Next i
    End With
    SS_LastStep = "style"
    SS_ApplyStyle ws, n
    SS_LastStep = "done"
    ws.Activate
    ws.Range("B" & HDR_ROW + 1).Select
    Set SS_Create = ws
End Function

' 세 시트 모두 생성
Public Sub SS_CreateAll()
    Dim lv As Long, ws As Worksheet, made As String
    For lv = 1 To 3
        Set ws = SS_Create(lv, True)
        If Not ws Is Nothing Then made = made & IIf(made = "", "", ", ") & ws.name
    Next lv
    If made <> "" Then MsgBox "현황시트 준비 완료: " & made, vbInformation, "학교명·지원청"
End Sub

' ---------------------------------------------------------------- 애플 스타일 표
Private Sub SS_ApplyStyle(ByVal ws As Worksheet, ByVal n As Long)
    Dim lastR As Long, body As Range, hdr As Range, r As Long
    lastR = HDR_ROW + n
    With ws
        .Cells.Font.name = "맑은 고딕"
        .Cells.Font.Size = 10.5
        .Cells.Font.color = CLR_TEXT
        .Cells.Interior.Pattern = xlNone
        .Cells.Borders.LineStyle = xlNone

        ' 제목
        With .Range("A1")
            .Font.Size = 17: .Font.Bold = True
        End With
        .rows(1).RowHeight = 30
        With .Range("A2")
            .Font.Size = 9.5: .Font.color = CLR_SUB
        End With
        .rows(2).RowHeight = 18
        .rows(3).RowHeight = 8

        SS_LastStep = "header"
        ' 머리글
        Set hdr = .Range(.Cells(HDR_ROW, 1), .Cells(HDR_ROW, 6))
        With hdr
            .Font.Bold = True: .Font.Size = 10.5
            .Font.color = CLR_SUB
            .Interior.color = CLR_HEAD
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
            .IndentLevel = 1
            With .Borders(xlEdgeBottom)
                .LineStyle = xlContinuous: .weight = xlThin: .color = CLR_LINE
            End With
        End With
        .rows(HDR_ROW).RowHeight = 26

        SS_LastStep = "body"
        ' 본문
        Set body = .Range(.Cells(HDR_ROW + 1, 1), .Cells(lastR, 6))
        With body
            .VerticalAlignment = xlCenter
            .HorizontalAlignment = xlLeft
            .IndentLevel = 1
            .Borders(xlInsideHorizontal).LineStyle = xlContinuous
            .Borders(xlInsideHorizontal).weight = xlHairline
            .Borders(xlInsideHorizontal).color = CLR_LINE
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).weight = xlThin
            .Borders(xlEdgeBottom).color = CLR_LINE
        End With
        .Range(.Cells(HDR_ROW + 1, 1), .Cells(lastR, 1)).HorizontalAlignment = xlCenter
        .Range(.Cells(HDR_ROW + 1, 1), .Cells(lastR, 1)).IndentLevel = 0
        .Range(.Cells(HDR_ROW + 1, 1), .Cells(lastR, 1)).Font.color = CLR_SUB
        .Range(.Cells(HDR_ROW + 1, 5), .Cells(lastR, 6)).Font.color = CLR_SUB
        .Range(.Cells(HDR_ROW, 1), .Cells(lastR, 1)).HorizontalAlignment = xlCenter
        For r = HDR_ROW + 1 To lastR
            .rows(r).RowHeight = 21
        Next r

        SS_LastStep = "width"
        ' 열 너비
        .Columns(1).ColumnWidth = 6
        .Columns(2).ColumnWidth = 26
        .Columns(3).ColumnWidth = 18
        .Columns(4).ColumnWidth = 10
        .Columns(5).ColumnWidth = 14
        .Columns(6).ColumnWidth = 22

        ' 화면 (창이 없는 환경에서도 오류 없이)
        SS_LastStep = "window"
        On Error Resume Next
        .Activate
        ActiveWindow.DisplayGridlines = False
        ActiveWindow.Zoom = 100
        .Range("A" & HDR_ROW + 1).Select
        ActiveWindow.FreezePanes = False
        ActiveWindow.SplitColumn = 0: ActiveWindow.SplitRow = HDR_ROW
        ActiveWindow.FreezePanes = True
        If .AutoFilterMode Then .AutoFilterMode = False
        .Range(.Cells(HDR_ROW, 1), .Cells(lastR, 6)).AutoFilter
        .Tab.color = CLR_ACCENT
        On Error GoTo 0
    End With
End Sub


