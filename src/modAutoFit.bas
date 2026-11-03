Attribute VB_Name = "modAutoFit"
Option Explicit
' =====================================================================================
'  modAutoFit : 셀간격정리 (v2.3 고도화)
'   메뉴 1. 모든 글자 1행으로   : 줄바꿈 해제, 모든 내용이 한 줄로 보이도록 열 너비 자동(#### 방지)
'   메뉴 2. A4 가로 인쇄용      : 인쇄 가능 폭에 맞춰 열 너비 배분, 글자=자동 줄바꿈, 숫자=셀에 맞춤
'   메뉴 3. A4 세로 인쇄용      : 위와 같음(세로 방향)
'   메뉴 4. 되돌리기(원래대로)  : 1~3 실행 전 상태(열 너비·행 높이·줄바꿈·셀에 맞춤·인쇄 방향) 복원
'   - 되돌리기 정보는 통합 문서의 숨김 시트(_BX_셀간격백업)에 시트별로 1개 보관(첫 실행 시점 기준). 이름이 '_셀간격백업'으로 끝나는 예전 백업 시트도 자동 인식
'   - 대상 범위: 2셀 이상 선택 시 그 범위, 아니면 시트 사용 영역 전체.
' =====================================================================================
Private Const SNAP_SHEET As String = "_BX_셀간격백업"
Private Const SNAP_SUFFIX As String = "_셀간격백업"   '예전 버전 백업 시트 인식용(이름 끝 글자로 찾음)
Private Const CHUNK As Long = 30000            ' 셀 하나에 담는 최대 글자 수
Private Const A4_W As Single = 595.3           ' A4 폭(pt)
Private Const A4_H As Single = 841.9           ' A4 높이(pt)
Private Const MIN_TEXT_W As Single = 6         ' 글자 열 최소 너비(문자 단위)
Private Const MAX_COL_W As Single = 250        ' 열 너비 상한
Private AF_Silent As Boolean                   ' True 면 결과 MsgBox 생략(자동 검증용)

Public Sub AF_SetSilent(ByVal v As Boolean)
    AF_Silent = v
End Sub

Private Sub AF_Info(ByVal msg As String)
    If AF_Silent Or Not Application.Visible Then Exit Sub
    MsgBox msg, vbInformation, "셀간격정리"
End Sub

' ---------------------------------------------------------------- 진입점(리본)
Sub 셀자동맞춤2(control As IRibbonControl)
    Dim ws As Worksheet, rng As Range, choice As Long
    Set ws = ActiveSheet
    Set rng = AF_TargetRange(ws)
    If rng Is Nothing Then
        MsgBox "정리할 데이터가 없습니다.", vbExclamation, "셀간격정리"
        Exit Sub
    End If
    choice = UI_Choose("셀간격정리", "대상 범위: " & rng.Address(False, False) & vbCrLf & "원하는 정리 방식을 선택하세요.", _
             Array("모든 글자 1행으로  (줄바꿈 없이, 숫자 #### 방지)", _
                   "A4 가로 인쇄용 정리  (글자 자동 줄바꿈, 숫자 셀에 맞춤)", _
                   "A4 세로 인쇄용 정리  (글자 자동 줄바꿈, 숫자 셀에 맞춤)", _
                   "되돌리기  (정리 전 원래 상태로)"))
    Select Case choice
        Case 1: AF_OneLine ws, rng
        Case 2: AF_FitPage ws, rng, xlLandscape
        Case 3: AF_FitPage ws, rng, xlPortrait
        Case 4: AF_Restore ws
    End Select
End Sub

' ---------------------------------------------------------------- 대상 범위
Private Function AF_TargetRange(ByVal ws As Worksheet) As Range
    Dim rng As Range
    If TypeName(Selection) = "Range" Then
        If Selection.Cells.CountLarge > 1 Then Set rng = Selection
    End If
    If rng Is Nothing Then Set rng = ws.UsedRange
    If rng Is Nothing Then Exit Function
    If rng.Areas.Count > 1 Then Set rng = rng.Areas(1)
    If rng.rows.Count = ws.rows.Count Or rng.Columns.Count = ws.Columns.Count Then
        If Intersect(rng, ws.UsedRange) Is Nothing Then Exit Function
        Set rng = Intersect(rng, ws.UsedRange)
    End If
    If Application.WorksheetFunction.CountA(rng) = 0 Then Exit Function
    Set AF_TargetRange = rng
End Function

' ---------------------------------------------------------------- 1. 모든 글자 1행으로
Public Sub AF_OneLine(ByVal ws As Worksheet, ByVal rng As Range)
    Dim c As Range, capped As Long
    AF_Snapshot ws, rng
    Application.ScreenUpdating = False
    With rng
        .WrapText = False
        .ShrinkToFit = False
    End With
    rng.Columns.AutoFit
    For Each c In rng.Columns
        If Not c.hidden Then
            If c.ColumnWidth > MAX_COL_W Then c.ColumnWidth = MAX_COL_W: capped = capped + 1
        End If
    Next c
    rng.rows.AutoFit
    Application.ScreenUpdating = True
    AF_Info "모든 글자를 한 줄로 정리했습니다." & vbCrLf & vbCrLf & _
           "대상 범위: " & rng.Address(False, False) & _
           IIf(capped > 0, vbCrLf & "너비 상한(" & MAX_COL_W & ")에 걸린 열: " & capped & "개", "") & vbCrLf & vbCrLf & _
           "되돌리려면 [셀간격정리 → 되돌리기]를 실행하세요."
End Sub

' ---------------------------------------------------------------- 2·3. A4 인쇄용
Public Sub AF_FitPage(ByVal ws As Worksheet, ByVal rng As Range, ByVal orient As Long)
    Dim c As Range, col As Range, i As Long, n As Long
    Dim natW() As Single, minW() As Single, hidden() As Boolean
    Dim u As Single, p As Single, availPt As Single, availUnits As Single
    Dim sumNat As Single, sumMin As Single, sumFlex As Single, remain As Single
    Dim W As Single, orientName As String

    AF_Snapshot ws, rng
    Application.ScreenUpdating = False

    ' (1) 페이지 설정: A4 + 방향
    On Error Resume Next
    With ws.PageSetup
        .PaperSize = xlPaperA4
        .Orientation = orient
    End With
    On Error GoTo 0
    availPt = IIf(orient = xlLandscape, A4_H, A4_W) - ws.PageSetup.LeftMargin - ws.PageSetup.RightMargin

    ' (2) 셀 속성: 글자=자동 줄바꿈, 숫자(날짜·수식 결과 포함)=셀에 맞춤
    AF_ApplyCellFormat rng

    ' (3) 각 열의 자연 너비(한 줄 기준)와 최소 너비(숫자가 깨지지 않는 폭) 측정
    n = rng.Columns.Count
    ReDim natW(1 To n): ReDim minW(1 To n): ReDim hidden(1 To n)
    rng.WrapText = False                       ' 측정을 위해 잠시 해제
    rng.Columns.AutoFit
    For i = 1 To n
        Set col = rng.Columns(i)
        hidden(i) = col.hidden
        If hidden(i) Then
            natW(i) = 0: minW(i) = 0
        Else
            natW(i) = col.ColumnWidth
            If natW(i) > MAX_COL_W Then natW(i) = MAX_COL_W
            minW(i) = AF_MinWidth(col)
            If minW(i) > natW(i) Then minW(i) = natW(i)
            sumNat = sumNat + natW(i)
            sumMin = sumMin + minW(i)
        End If
    Next i
    AF_ApplyCellFormat rng                     ' 줄바꿈 다시 적용

    ' (4) 열 너비 배분: 너비(문자단위)→pt 환산계수 측정 후, 인쇄 폭 안에 들어가도록 배분
    AF_MeasureUnit ws, u, p
    availUnits = (availPt - p * AF_VisibleCount(hidden)) / u
    If availUnits < 1 Then availUnits = 1

    If sumNat <= availUnits Then
        For i = 1 To n
            If Not hidden(i) Then rng.Columns(i).ColumnWidth = natW(i)
        Next i
    ElseIf sumMin >= availUnits Then
        ' 최소 너비 합조차 페이지보다 크면 비율 축소(+ 인쇄 시 가로 1쪽 맞춤이 보완)
        For i = 1 To n
            If Not hidden(i) Then rng.Columns(i).ColumnWidth = AF_Max(minW(i) * availUnits / sumMin, 2)
        Next i
    Else
        ' 최소 너비는 보장하고 남는 폭을 '더 필요한 만큼'에 비례해 나눠준다
        remain = availUnits - sumMin
        For i = 1 To n
            If Not hidden(i) Then sumFlex = sumFlex + (natW(i) - minW(i))
        Next i
        For i = 1 To n
            If Not hidden(i) Then
                W = minW(i)
                If sumFlex > 0 Then W = W + (natW(i) - minW(i)) * remain / sumFlex
                rng.Columns(i).ColumnWidth = W
            End If
        Next i
    End If

    ' (5) 행 높이 자동 + 인쇄 시 가로 1쪽 안전장치
    rng.rows.AutoFit
    On Error Resume Next
    With ws.PageSetup
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
    End With
    On Error GoTo 0
    Application.ScreenUpdating = True

    orientName = IIf(orient = xlLandscape, "가로", "세로")
    AF_Info "A4 " & orientName & " 인쇄에 맞춰 정리했습니다." & vbCrLf & vbCrLf & _
           "대상 범위: " & rng.Address(False, False) & vbCrLf & _
           "글자 셀: 자동 줄바꿈 / 숫자 셀: 셀에 맞춤" & vbCrLf & _
           "페이지 설정: A4 " & orientName & ", 가로 1쪽 맞춤" & vbCrLf & vbCrLf & _
           "되돌리려면 [셀간격정리 → 되돌리기]를 실행하세요."
End Sub

' 글자 셀 = 자동 줄바꿈, 숫자 셀 = 셀에 맞춤
Private Sub AF_ApplyCellFormat(ByVal rng As Range)
    Dim c As Range, v As Variant
    For Each c In rng.Cells
        If c.MergeCells Then
            If c.Address <> c.MergeArea.Cells(1, 1).Address Then GoTo NextCell
        End If
        v = c.Value2
        If IsEmpty(v) Then
            ' 빈 셀은 그대로
        ElseIf AF_IsNumberLike(v) Then
            c.WrapText = False
            c.ShrinkToFit = True
        Else
            c.ShrinkToFit = False
            c.WrapText = True
        End If
NextCell:
    Next c
End Sub

Private Function AF_IsNumberLike(ByVal v As Variant) As Boolean
    Select Case VarType(v)
        Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDate, vbDecimal, vbBoolean, vbError
            AF_IsNumberLike = True
        Case Else
            AF_IsNumberLike = False
    End Select
End Function

' 열 안의 숫자 셀이 깨지지 않을 최소 너비(문자 단위) - 표시 문자열 길이 × 글꼴 배율
Private Function AF_MinWidth(ByVal col As Range) As Single
    Dim c As Range, L As Single, best As Single, base As Single, T As String
    base = ActiveWorkbook.Styles("Normal").Font.Size
    If base <= 0 Then base = 11
    best = MIN_TEXT_W
    For Each c In col.Cells
        If Not IsEmpty(c.Value2) Then
            If AF_IsNumberLike(c.Value2) Then
                T = c.text
                If Left$(T, 1) = "#" Then T = CStr(c.Value2)
                L = Len(T) * (c.Font.Size / base) * 1.05 + 1.5
                If c.Font.Bold Then L = L * 1.08
                If L > best Then best = L
            End If
        End If
    Next c
    AF_MinWidth = best
End Function

' 열 너비 1문자 단위가 몇 pt인지(u)와 열당 고정 여백(p)을 실측
Private Sub AF_MeasureUnit(ByVal ws As Worksheet, ByRef u As Single, ByRef p As Single)
    Dim c As Range, w1 As Single, pt1 As Single, pt2 As Single
    On Error GoTo Fallback
    Set c = ws.Columns(ws.Columns.Count)
    w1 = c.ColumnWidth
    If w1 <= 0 Then w1 = ws.StandardWidth: c.ColumnWidth = w1
    pt1 = c.Width
    c.ColumnWidth = w1 + 10
    pt2 = c.Width
    c.ColumnWidth = w1
    u = (pt2 - pt1) / 10
    p = pt1 - w1 * u
    If u <= 0 Then GoTo Fallback
    Exit Sub
Fallback:
    u = 5.25: p = 3.75
End Sub

Private Function AF_VisibleCount(ByRef hidden() As Boolean) As Long
    Dim i As Long
    For i = LBound(hidden) To UBound(hidden)
        If Not hidden(i) Then AF_VisibleCount = AF_VisibleCount + 1
    Next i
End Function

Private Function AF_Max(ByVal a As Single, ByVal b As Single) As Single
    If a > b Then AF_Max = a Else AF_Max = b
End Function

' ---------------------------------------------------------------- 되돌리기 정보 저장/복원
' 숨김 시트 한 행 = 시트 하나. A:시트명 B:저장시각 C:범위 D:방향 E:용지 F:Zoom G:가로쪽수 H:세로쪽수
'   I~: 열너비 "열:너비;..." / 행높이 "행:높이;..." / 줄바꿈 셀 / 셀에맞춤 셀  (각각 CHUNK 단위로 이어서 저장, 구분자 셀 "|")
Private Function AF_SnapSheet(ByVal wb As Workbook, ByVal createIfMissing As Boolean) As Worksheet
    Dim ws As Worksheet, W As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(SNAP_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        '예전 버전이 만든 백업 시트(이름이 '_셀간격백업'으로 끝나는 시트)가 있으면 새 이름으로 바꿔 계속 사용
        For Each W In wb.Worksheets
            If W.name Like "*" & SNAP_SUFFIX Then Set ws = W: Exit For
        Next W
        If Not ws Is Nothing Then
            On Error Resume Next
            ws.name = SNAP_SHEET
            On Error GoTo 0
        End If
    End If
    If ws Is Nothing And createIfMissing Then
        On Error Resume Next
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        If Not ws Is Nothing Then
            ws.name = SNAP_SHEET
            ws.Visible = xlSheetVeryHidden
        End If
        On Error GoTo 0
    End If
    Set AF_SnapSheet = ws
End Function

Private Function AF_FindRow(ByVal snap As Worksheet, ByVal sheetName As String) As Long
    Dim r As Long
    For r = 1 To snap.Cells(snap.rows.Count, 1).End(xlUp).row
        If snap.Cells(r, 1).value = sheetName Then AF_FindRow = r: Exit Function
    Next r
End Function

Private Sub AF_Snapshot(ByVal ws As Worksheet, ByVal rng As Range)
    Dim snap As Worksheet, r As Long, c As Range, col As Range, rw As Range
    Dim sW As String, Sh As String, sWrap As String, sShr As String, colIdx As Long
    Dim orig As Worksheet
    Set orig = ActiveSheet
    Set snap = AF_SnapSheet(ws.parent, True)
    If snap Is Nothing Then
        orig.Activate
        MsgBox "되돌리기 정보를 저장할 수 없어(통합 문서 구조 보호 등) 되돌리기 없이 진행합니다.", vbExclamation, "셀간격정리"
        Exit Sub
    End If
    r = AF_FindRow(snap, ws.name)
    If r > 0 Then Exit Sub                     ' 첫 실행 상태를 유지(되돌리기 전까지 덮어쓰지 않음)
    r = snap.Cells(snap.rows.Count, 1).End(xlUp).row + 1
    If snap.Cells(1, 1).value = "" Then r = 1

    For Each col In rng.Columns
        sW = sW & col.Column & ":" & col.ColumnWidth & IIf(col.hidden, "h", "") & ";"
    Next col
    For Each rw In rng.rows
        Sh = Sh & rw.row & ":" & rw.RowHeight & IIf(rw.hidden, "h", "") & ";"
    Next rw
    For Each c In rng.Cells
        If c.WrapText Then sWrap = sWrap & c.Address(False, False) & ";"
        If c.ShrinkToFit Then sShr = sShr & c.Address(False, False) & ";"
    Next c

    With snap
        .Cells(r, 1).value = ws.name
        .Cells(r, 2).value = Format$(Now, "yyyy-mm-dd hh:nn:ss")
        .Cells(r, 3).value = rng.Address(False, False)
        On Error Resume Next
        .Cells(r, 4).value = ws.PageSetup.Orientation
        .Cells(r, 5).value = ws.PageSetup.PaperSize
        .Cells(r, 6).value = ws.PageSetup.Zoom
        .Cells(r, 7).value = ws.PageSetup.FitToPagesWide
        .Cells(r, 8).value = ws.PageSetup.FitToPagesTall
        On Error GoTo 0
        colIdx = 9
        AF_PutLong snap, r, colIdx, sW
        AF_PutLong snap, r, colIdx, Sh
        AF_PutLong snap, r, colIdx, sWrap
        AF_PutLong snap, r, colIdx, sShr
    End With
    orig.Activate
End Sub

' 긴 문자열을 CHUNK 단위로 여러 셀에 나눠 쓰고 마지막에 "|" 구분 셀을 둔다
Private Sub AF_PutLong(ByVal snap As Worksheet, ByVal r As Long, ByRef colIdx As Long, ByVal s As String)
    Dim pos As Long
    pos = 1
    Do While pos <= Len(s)
        snap.Cells(r, colIdx).NumberFormat = "@"
        snap.Cells(r, colIdx).value = Mid$(s, pos, CHUNK)
        colIdx = colIdx + 1
        pos = pos + CHUNK
    Loop
    snap.Cells(r, colIdx).value = "|"
    colIdx = colIdx + 1
End Sub

Private Function AF_GetLong(ByVal snap As Worksheet, ByVal r As Long, ByRef colIdx As Long) As String
    Dim s As String, v As String
    Do
        v = CStr(snap.Cells(r, colIdx).value)
        colIdx = colIdx + 1
        If v = "|" Then Exit Do
        If Len(v) = 0 Then Exit Do
        s = s & v
    Loop
    AF_GetLong = s
End Function

Public Sub AF_Restore(ByVal ws As Worksheet)
    Dim snap As Worksheet, r As Long, colIdx As Long, rng As Range
    Dim sW As String, Sh As String, sWrap As String, sShr As String
    Dim parts() As String, kv() As String, i As Long, isHidden As Boolean, num As String
    Set snap = AF_SnapSheet(ws.parent, False)
    If Not snap Is Nothing Then r = AF_FindRow(snap, ws.name)
    If r = 0 Then
        MsgBox "'" & ws.name & "' 시트에는 되돌릴 셀간격정리 기록이 없습니다." & vbCrLf & _
               "(1~3번 정리를 실행하면 실행 전 상태가 자동 저장됩니다.)"
        Exit Sub
    End If
    Application.ScreenUpdating = False
    On Error Resume Next
    Set rng = ws.Range(CStr(snap.Cells(r, 3).value))
    On Error GoTo 0
    colIdx = 9
    sW = AF_GetLong(snap, r, colIdx)
    Sh = AF_GetLong(snap, r, colIdx)
    sWrap = AF_GetLong(snap, r, colIdx)
    sShr = AF_GetLong(snap, r, colIdx)

    ' 줄바꿈·셀에 맞춤: 범위 전체 해제 후 기록된 셀만 복원
    If Not rng Is Nothing Then
        rng.WrapText = False
        rng.ShrinkToFit = False
    End If
    parts = Split(sShr, ";")
    For i = 0 To UBound(parts)
        If Len(parts(i)) > 0 Then ws.Range(parts(i)).ShrinkToFit = True
    Next i
    parts = Split(sWrap, ";")
    For i = 0 To UBound(parts)
        If Len(parts(i)) > 0 Then ws.Range(parts(i)).WrapText = True
    Next i

    ' 열 너비
    parts = Split(sW, ";")
    For i = 0 To UBound(parts)
        If InStr(parts(i), ":") > 0 Then
            kv = Split(parts(i), ":")
            num = kv(1): isHidden = (Right$(num, 1) = "h")
            If isHidden Then num = Left$(num, Len(num) - 1)
            ws.Columns(CLng(kv(0))).ColumnWidth = CSng(num)
            ws.Columns(CLng(kv(0))).hidden = isHidden
        End If
    Next i
    ' 행 높이
    parts = Split(Sh, ";")
    For i = 0 To UBound(parts)
        If InStr(parts(i), ":") > 0 Then
            kv = Split(parts(i), ":")
            num = kv(1): isHidden = (Right$(num, 1) = "h")
            If isHidden Then num = Left$(num, Len(num) - 1)
            ws.rows(CLng(kv(0))).RowHeight = CSng(num)
            ws.rows(CLng(kv(0))).hidden = isHidden
        End If
    Next i
    ' 페이지 설정
    On Error Resume Next
    With ws.PageSetup
        .Orientation = snap.Cells(r, 4).value
        .PaperSize = snap.Cells(r, 5).value
        If snap.Cells(r, 6).value = False Or snap.Cells(r, 6).value = 0 Then
            .Zoom = False
            .FitToPagesWide = snap.Cells(r, 7).value
            .FitToPagesTall = snap.Cells(r, 8).value
        Else
            .Zoom = snap.Cells(r, 6).value
        End If
    End With
    On Error GoTo 0

    snap.rows(r).Delete                       ' 기록 소거(다음 정리 때 새로 저장)
    If snap.Cells(1, 1).value = "" Then
        On Error Resume Next
        Application.DisplayAlerts = False
        snap.Visible = xlSheetVisible
        snap.Delete
        Application.DisplayAlerts = True
        If Err.Number <> 0 Then Err.Clear: snap.Visible = xlSheetVeryHidden
        On Error GoTo 0
    End If
    ws.Activate
    Application.ScreenUpdating = True
    AF_Info "'" & ws.name & "' 시트를 셀간격정리 이전 상태로 되돌렸습니다." & vbCrLf & _
           "(열 너비·행 높이·줄바꿈·셀에 맞춤·인쇄 방향)"
End Sub


