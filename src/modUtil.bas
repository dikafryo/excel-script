Attribute VB_Name = "modUtil"
Option Explicit
'---------------------------------------------------------------- 보드 통합문서 (xlsm 단독 / xlam 추가기능 공용)
'  xlsm 으로 쓰면 ThisWorkbook, 추가기능(xlam)으로 쓰면 활성 통합문서(보드) 또는 새 통합문서
Public gBoardWB As Workbook


Public gPrintLight As Boolean   '인쇄용 밝은 모드 (진한 배경 → 흰색)
'=====================================================================
' modUtil : 공통 상수 / 유틸리티
'=====================================================================
Public Const SH_MAIN As String = "메인"
Public Const SH_REQ As String = "원본_요구서"
Public Const SH_A As String = "원본_102A"
Public Const SH_B As String = "원본_102B"
Public Const SH_CARD As String = "단위사업카드"
Public Const SH_STAFF As String = "부서담당자"
Public Const SH_DASH As String = "사업비현황"
Public Const SH_DETAIL As String = "사업상세"
Public Const SH_ITEM As String = "세목현황"
Public Const SH_CFG As String = "설정"

Public Const FONT_NM As String = "맑은 고딕"

'---------------------------------------------------------------- 색상
Public Function BW() As Workbook
    Dim nm As String
    If Not ThisWorkbook.IsAddin Then Set BW = ThisWorkbook: Exit Function
    On Error Resume Next
    If Not ActiveWorkbook Is Nothing Then
        If Not (ActiveWorkbook Is ThisWorkbook) Then
            If IsBoardWB(ActiveWorkbook) Then Set gBoardWB = ActiveWorkbook
        End If
    End If
    nm = ""
    If Not gBoardWB Is Nothing Then nm = gBoardWB.name
    If Err.Number <> 0 Or nm = "" Then Err.Clear: Set gBoardWB = Nothing
    On Error GoTo 0
    If gBoardWB Is Nothing Then
        Set gBoardWB = Workbooks.Add
    End If
    Set BW = gBoardWB
End Function

'예산집행보드 통합문서인지 (메인 또는 설정 시트 존재)
Public Function IsBoardWB(wb As Workbook) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(SH_MAIN)
    If ws Is Nothing Then Set ws = wb.Worksheets(SH_CFG)
    On Error GoTo 0
    IsBoardWB = Not ws Is Nothing
End Function

Public Function ClrNavy() As Long
    ClrNavy = RGB(28, 28, 30)        'Apple Label (거의 검정, 텍스트용)
End Function
Public Function ClrBlue() As Long
    ClrBlue = RGB(0, 122, 255)       'Apple Blue
End Function
Public Function ClrGreen() As Long
    ClrGreen = RGB(52, 199, 89)      'Apple Green
End Function
Public Function ClrOrange() As Long
    ClrOrange = RGB(255, 149, 0)     'Apple Orange
End Function
Public Function ClrGray() As Long
    ClrGray = RGB(142, 142, 147)     'Apple Gray
End Function
Public Function ClrLight() As Long
    ClrLight = RGB(242, 242, 247)    'Apple Gray6 (배경)
End Function
Public Function ClrRed() As Long
    ClrRed = RGB(255, 59, 48)        'Apple Red
End Function
Public Function ClrIndigo() As Long
    ClrIndigo = RGB(88, 86, 214)     'Apple Indigo
End Function
Public Function ClrHead() As Long
    ClrHead = RGB(229, 229, 234)     'Apple Gray5 (헤더/구분선)
End Function

'---------------------------------------------------------------- 시트
Public Function SheetExists(ByVal nm As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = BW().Worksheets(nm)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function

Public Function GetSheet(ByVal nm As String, Optional ByVal clearAll As Boolean = False) As Worksheet
    Dim ws As Worksheet
    If SheetExists(nm) Then
        Set ws = BW().Worksheets(nm)
        If clearAll Then ClearSheet ws
    Else
        Set ws = BW().Worksheets.Add(After:=BW().Worksheets(BW().Worksheets.Count))
        ws.name = nm
    End If
    Set GetSheet = ws
End Function

Public Sub ClearSheet(ws As Worksheet)
    ClearRowsKeepTables ws, 1
End Sub

'행 fromRow 이후를 비우되, 엑셀 표(머리글 행)와 슬라이서는 유지 → 슬라이서 캐시가 살아 있어 재생성 불필요
'  표 본문 행은 삭제(머리글만 남김), 차트·일반 도형 삭제, 나머지 셀 Clear
Public Sub ClearRowsKeepTables(ws As Worksheet, ByVal fromRow As Long)
    Dim shp As Shape, co As ChartObject, lo As ListObject, i As Long, keep As Range, r As Long, lastR As Long, lastC As Long, c As Long
    Dim r1 As Long
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    For i = ws.ListObjects.Count To 1 Step -1
        Set lo = ws.ListObjects(i)
        On Error Resume Next
        lo.AutoFilter.ShowAllData
        lo.ShowTotals = False
        If lo.ListRows.Count > 0 Then lo.DataBodyRange.Delete
        If Err.Number <> 0 Then Err.Clear: lo.Unlist
        On Error GoTo 0
    Next i
    For i = 1 To ws.ListObjects.Count
        If keep Is Nothing Then Set keep = ws.ListObjects(i).HeaderRowRange Else Set keep = Union(keep, ws.ListObjects(i).HeaderRowRange)
    Next i
    If keep Is Nothing Then
        If fromRow <= 1 Then ws.Cells.Clear Else ws.Range(ws.rows(fromRow), ws.rows(ws.rows.Count)).Clear
    Else
        lastR = ws.UsedRange.row + ws.UsedRange.rows.Count - 1
        lastC = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1
        r1 = fromRow
        For r = fromRow To lastR
            If Not Intersect(ws.rows(r), keep) Is Nothing Then
                If r > r1 Then ws.Range(ws.rows(r1), ws.rows(r - 1)).Clear
                For c = 1 To lastC
                    If Intersect(ws.Cells(r, c), keep) Is Nothing Then ws.Cells(r, c).Clear
                Next c
                r1 = r + 1
            End If
        Next r
        If lastR >= r1 Then ws.Range(ws.rows(r1), ws.rows(ws.rows.Count)).Clear Else ws.Range(ws.rows(r1), ws.rows(ws.rows.Count)).Clear
    End If
    ws.Range(ws.rows(fromRow), ws.rows(ws.rows.Count)).RowHeight = ws.StandardHeight
    For Each co In ws.ChartObjects: co.Delete: Next co
    For i = ws.Shapes.Count To 1 Step -1
        Set shp = ws.Shapes(i)
        If shp.Type <> msoSlicer Then
            If fromRow <= 1 Or shp.top >= ws.rows(fromRow).top - 1 Then shp.Delete
        End If
    Next i
End Sub

'엑셀 표 생성/재사용: 같은 이름의 표가 있으면 크기만 조정(슬라이서 캐시 유지), 없으면 새로 생성
Public Function MakeTable(ws As Worksheet, rng As Range, ByVal nm As String) As ListObject
    Dim lo As ListObject
    On Error Resume Next
    Set lo = ws.ListObjects(nm)
    On Error GoTo 0
    If Not lo Is Nothing Then
        If lo.HeaderRowRange.row = rng.row And lo.Range.Column = rng.Column Then
            lo.Resize rng
        Else
            lo.Unlist: Set lo = Nothing
        End If
    End If
    If lo Is Nothing Then
        Set lo = ws.ListObjects.Add(xlSrcRange, rng, , xlYes)
        lo.name = nm
    End If
    lo.TableStyle = "TableStyleLight1"
    lo.ShowTableStyleRowStripes = False
    Set MakeTable = lo
End Function

Public Function lastRow(ws As Worksheet, Optional ByVal col As Long = 1) As Long
    lastRow = ws.Cells(ws.rows.Count, col).End(xlUp).row
End Function

'헤더 텍스트가 있는 행 찾기 (완전일치)
Public Function FindHeaderRow(ws As Worksheet, ByVal mustText As String, Optional ByVal maxScan As Long = 20) As Long
    Dim r As Long, f As Range
    For r = 1 To maxScan
        Set f = ws.rows(r).Find(what:=mustText, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
        If Not f Is Nothing Then FindHeaderRow = r: Exit Function
    Next r
    FindHeaderRow = 0
End Function

'헤더명 -> 열번호 (첫 번째 등장 기준)
Public Function HeaderMap(ws As Worksheet, ByVal hdrRow As Long) As Object
    Dim d As Object, c As Long, lastC As Long, T As String
    Set d = CreateObject("Scripting.Dictionary")
    lastC = ws.Cells(hdrRow, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastC
        T = CleanText(ws.Cells(hdrRow, c).Value2)
        If T <> "" Then If Not d.Exists(T) Then d.Add T, c
    Next c
    Set HeaderMap = d
End Function

Public Function HCol(hm As Object, ByVal nm As String, Optional ByVal required As Boolean = True) As Long
    If hm.Exists(nm) Then
        HCol = hm(nm)
    ElseIf required Then
        Err.Raise vbObjectError + 100, "HCol", "헤더 '" & nm & "' 열을 찾지 못했습니다."
    Else
        HCol = 0
    End If
End Function

'---------------------------------------------------------------- 문자열
Public Function CleanText(ByVal v As Variant) As String
    Dim T As String
    If IsError(v) Then CleanText = "": Exit Function
    T = CStr(v)
    T = Replace(T, vbLf, ""): T = Replace(T, vbCr, ""): T = Replace(T, vbTab, "")
    T = Replace(T, ChrW(160), " "): T = Replace(T, ChrW(12288), " ")
    CleanText = Trim(T)
End Function

'매칭용 정규화 키
Public Function NormKey(ByVal s As String) As String
    Dim T As String
    T = CleanText(s)
    T = Replace(T, " ", "")
    T = Replace(T, "(가칭)", ""): T = Replace(T, "(신)", ""): T = Replace(T, "(구)", "")
    T = Replace(T, ChrW(65288), "("): T = Replace(T, ChrW(65289), ")")
    NormKey = LCase(T)
End Function

'사업항목 키
Public Function MakeKey(ByVal 세부사업 As String, ByVal 사업항목 As String) As String
    MakeKey = NormKey(세부사업) & "|" & NormKey(사업항목)
End Function

'날짜 표시 형식 "2026.09.24."  (문자열/일련번호/날짜 모두 허용)
Public Function FmtDate(ByVal v As Variant) As String
    Dim T As String, d As Date, y As String, m As String, dd As String, i As Long, digits As String
    On Error Resume Next
    If IsDate(v) And Not IsNumeric(v) Then
        d = CDate(v): FmtDate = Format(d, "yyyy.mm.dd."): Exit Function
    End If
    If IsNumeric(v) Then
        If CDbl(v) > 30000 And CDbl(v) < 80000 Then
            d = CDate(CDbl(v)): FmtDate = Format(d, "yyyy.mm.dd."): Exit Function
        End If
    End If
    T = CStr(v)
    '숫자만 추출 → 8자리(yyyymmdd) 이면 변환
    For i = 1 To Len(T)
        If Mid(T, i, 1) Like "#" Then digits = digits & Mid(T, i, 1)
    Next i
    If Len(digits) >= 8 Then
        y = Left(digits, 4): m = Mid(digits, 5, 2): dd = Mid(digits, 7, 2)
        If IsDate(y & "-" & m & "-" & dd) Then FmtDate = y & "." & m & "." & dd & ".": Exit Function
    End If
    FmtDate = Trim(T)
End Function

Public Function NowText() As String
    NowText = Format(Now, "yyyy.mm.dd. hh:nn")
End Function

'안전한 나눗셈 (0/0 은 VBA에서 '오버플로' 오류 → 반드시 이 함수 사용)
Public Function rate(ByVal num As Double, ByVal den As Double) As Double
    If den = 0 Then rate = 0 Else rate = num / den
End Function

Public Function NzD(ByVal v As Variant) As Double
    If IsError(v) Then NzD = 0: Exit Function
    If IsNumeric(v) Then NzD = CDbl(v) Else NzD = 0
End Function

'중복 없이 ", " 로 연결
Public Function JoinUnique(ByVal base As String, ByVal addText As String) As String
    Dim arr As Variant, i As Long, T As String, r As String
    r = base
    arr = Split(addText, ",")
    For i = LBound(arr) To UBound(arr)
        T = CleanText(arr(i))
        If T <> "" Then
            If InStr(1, ", " & r & ", ", ", " & T & ", ") = 0 Then
                If r = "" Then r = T Else r = r & ", " & T
            End If
        End If
    Next i
    JoinUnique = r
End Function

'---------------------------------------------------------------- 속도
Public Sub SpeedOn(Optional ByVal msg As String = "")
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
        .DisplayAlerts = False
        If msg <> "" Then .StatusBar = msg
    End With
End Sub

Public Sub SpeedOff()
    With Application
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
        .ScreenUpdating = True
        .DisplayAlerts = True
        .StatusBar = False
    End With
End Sub

'---------------------------------------------------------------- 설정
Public Function CfgSheet() As Worksheet
    Dim ws As Worksheet
    Set ws = GetSheet(SH_CFG)
    If ws.Range("A1").Value2 = "" Then
        ws.Range("A1:B1").Value2 = Array("키", "값")
        ws.Range("D1").Value2 = "담당자목록"
    End If
    ws.Visible = xlSheetVeryHidden
    Set CfgSheet = ws
End Function

Public Function CfgGet(ByVal key As String) As String
    Dim ws As Worksheet, f As Range
    Set ws = CfgSheet()
    Set f = ws.Columns(1).Find(what:=key, LookIn:=xlValues, LookAt:=xlWhole)
    If f Is Nothing Then CfgGet = "" Else CfgGet = CStr(f.Offset(0, 1).Value2)
End Function

Public Sub CfgSet(ByVal key As String, ByVal val As String)
    Dim ws As Worksheet, f As Range, r As Long
    Set ws = CfgSheet()
    Set f = ws.Columns(1).Find(what:=key, LookIn:=xlValues, LookAt:=xlWhole)
    If f Is Nothing Then
        r = lastRow(ws, 1) + 1
        ws.Cells(r, 1).Value2 = key
        ws.Cells(r, 2).Value2 = val
    Else
        f.Offset(0, 1).Value2 = val
    End If
End Sub

'---------------------------------------------------------------- 버튼
Public Sub AddButton(ws As Worksheet, ByVal nm As String, ByVal cap As String, ByVal macro As String, _
                     ByVal L As Single, ByVal T As Single, ByVal W As Single, ByVal h As Single, ByVal fillColor As Long, _
                     Optional ByVal fontPt As Single = 9)
    Dim shp As Shape
    On Error Resume Next
    ws.Shapes(nm).Delete
    On Error GoTo 0
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, L, T, W, h)
    With shp
        .name = nm
        .Fill.ForeColor.RGB = fillColor
        .line.Visible = msoFalse
        .Shadow.Visible = msoFalse
        With .TextFrame2
            .TextRange.text = cap
            .TextRange.Font.name = FONT_NM
            .TextRange.Font.Size = fontPt
            .TextRange.Font.Bold = msoTrue
            .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
            .VerticalAnchor = msoAnchorMiddle
            .MarginLeft = 2: .MarginRight = 2: .MarginTop = 1: .MarginBottom = 1
            .WordWrap = msoTrue
        End With
        .OnAction = "'" & ThisWorkbook.name & "'!" & macro    '추가기능(xlam)에서도 매크로를 찾도록 파일명 지정
        .Placement = xlFreeFloating
    End With
End Sub

'"메인으로" 버튼
Public Sub AddHomeButton(ws As Worksheet, Optional ByVal L As Single = 5, Optional ByVal T As Single = 5)
    AddButton ws, "btnHome", "◀ 메인", "메인으로", L, T, 60, 20, ClrGray()
End Sub

'---------------------------------------------------------------- 서식
Public Sub StyleHeader(rng As Range)
    With rng
        .Font.name = FONT_NM
        .Font.Bold = True
        .Font.Size = 10
        .Interior.color = ClrLight()
        .Font.color = ClrNavy()
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .RowHeight = 28
        .Borders.LineStyle = xlContinuous
        .Borders.color = ClrHead()
        .Borders(xlEdgeBottom).weight = xlMedium
        .Borders(xlEdgeBottom).color = ClrGray()
    End With
End Sub

Public Sub StyleBody(rng As Range)
    With rng
        .Font.name = FONT_NM
        .Font.Size = 10
        .RowHeight = 19
        .VerticalAlignment = xlCenter
        .Font.color = ClrNavy()
        .Borders.LineStyle = xlContinuous
        .Borders.color = ClrHead()
    End With
End Sub

'집행률 열 색상 (3색 스케일 + 데이터막대)
Public Sub StylePercent(rng As Range)
    Dim cs As ColorScale
    rng.NumberFormat = "0.0%"
    rng.HorizontalAlignment = xlRight
    rng.FormatConditions.Delete
    Set cs = rng.FormatConditions.AddColorScale(ColorScaleType:=3)
    cs.ColorScaleCriteria(1).Type = xlConditionValueNumber
    cs.ColorScaleCriteria(1).value = 0
    cs.ColorScaleCriteria(1).FormatColor.color = RGB(255, 180, 176)
    cs.ColorScaleCriteria(2).Type = xlConditionValueNumber
    cs.ColorScaleCriteria(2).value = 0.5
    cs.ColorScaleCriteria(2).FormatColor.color = RGB(255, 236, 153)
    cs.ColorScaleCriteria(3).Type = xlConditionValueNumber
    cs.ColorScaleCriteria(3).value = 1
    cs.ColorScaleCriteria(3).FormatColor.color = RGB(178, 236, 194)
End Sub

'데이터막대 (진행률 시각화)
Public Sub StyleBar(rng As Range)
    Dim db As Databar
    rng.NumberFormat = "0.0%"
    rng.FormatConditions.Delete
    Set db = rng.FormatConditions.AddDatabar
    db.MinPoint.Modify newtype:=xlConditionValueNumber, newValue:=0
    db.MaxPoint.Modify newtype:=xlConditionValueNumber, newValue:=1
    db.BarColor.color = ClrBlue()
    db.BarFillType = xlDataBarFillSolid
    db.ShowValue = False
    db.BarBorder.Type = xlDataBarBorderNone
End Sub


'행마다 색이 다른 데이터막대 (Apple 팔레트 순환)
Public Sub StyleBarMulti(rng As Range)
    Dim db As Databar, i As Long, cell As Range, pal As Variant
    pal = Array(ClrBlue(), ClrGreen(), ClrOrange(), ClrIndigo(), ClrRed(), RGB(90, 200, 250), RGB(175, 82, 222), RGB(255, 204, 0))
    rng.NumberFormat = "0.0%"
    rng.FormatConditions.Delete
    i = 0
    For Each cell In rng.rows
        Set db = cell.Cells(1, 1).FormatConditions.AddDatabar
        db.MinPoint.Modify newtype:=xlConditionValueNumber, newValue:=0
        db.MaxPoint.Modify newtype:=xlConditionValueNumber, newValue:=1
        db.BarColor.color = pal(i Mod (UBound(pal) + 1))
        db.BarFillType = xlDataBarFillSolid
        db.ShowValue = False
        db.BarBorder.Type = xlDataBarBorderNone
        i = i + 1
    Next cell
End Sub

'보드 인쇄용 밝은 모드 토글: 제목 띠(bandRow1~2)·KPI 값 셀(kpiRow) 진한 배경 ↔ 흰 배경 (다시 그리지 않음)
Public Sub BoardLight(ws As Worksheet, ByVal lastCol As Long, ByVal bandRow As Long, ByVal KpiRow As Long, ByVal lightOn As Boolean)
    Dim c As Long, ma As Range, clr As Long, Done As Range
    If lightOn Then
        ws.Range(ws.Cells(bandRow, 1), ws.Cells(bandRow + 1, lastCol)).Interior.color = RGB(255, 255, 255)
        ws.Cells(bandRow, 1).Font.color = ClrNavy(): ws.Cells(bandRow + 1, 1).Font.color = ClrGray()
    Else
        ws.Range(ws.Cells(bandRow, 1), ws.Cells(bandRow + 1, lastCol)).Interior.color = ClrNavy()
        ws.Cells(bandRow, 1).Font.color = RGB(255, 255, 255): ws.Cells(bandRow + 1, 1).Font.color = RGB(220, 220, 230)
    End If
    c = 1
    Do While c <= lastCol
        Set ma = ws.Cells(KpiRow, c).MergeArea
        If lightOn Then
            clr = ma.Interior.color
            If clr <> RGB(255, 255, 255) Then ma.Font.color = clr: ma.Interior.color = Tint(clr, 0.92)
        Else
            clr = ma.Font.color
            If clr <> RGB(255, 255, 255) Then ma.Interior.color = clr: ma.Font.color = RGB(255, 255, 255)
        End If
        c = ma.Column + ma.Columns.Count
    Loop
End Sub

Public Function StatusText(ByVal 예산현액 As Double, ByVal rate As Double, ByVal 매칭 As String) As String
    If 예산현액 = 0 Then
        If InStr(매칭, "장표없음") > 0 Then StatusText = "장표없음" Else StatusText = "예산없음"
    ElseIf rate >= 0.95 Then
        StatusText = "■ 완료"
    ElseIf rate >= 0.5 Then
        StatusText = "● 정상"
    ElseIf rate > 0 Then
        StatusText = "▲ 저조"
    Else
        StatusText = "△ 미집행"
    End If
End Function

'KPI 카드 그리기
Public Sub DrawKpi(ws As Worksheet, ByVal r As Long, ByVal c As Long, ByVal title As String, ByVal val As Variant, ByVal fmt As String, ByVal clr As Long, Optional ByVal span As Long = 2)
    'KPI 카드: 제목(연한 톤) + 값(진한 색 배경·흰 글씨) - 대시보드 카드 스타일
    With ws.Range(ws.Cells(r, c), ws.Cells(r, c + span - 1))
        .Merge
        .Value2 = title
        .Font.name = FONT_NM: .Font.Size = 9: .Font.color = clr: .Font.Bold = True
        .Interior.color = Tint(clr, 0.85)
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlNone
        .Borders(xlEdgeLeft).LineStyle = xlContinuous: .Borders(xlEdgeLeft).color = RGB(255, 255, 255): .Borders(xlEdgeLeft).weight = xlMedium
        .Borders(xlEdgeRight).LineStyle = xlContinuous: .Borders(xlEdgeRight).color = RGB(255, 255, 255): .Borders(xlEdgeRight).weight = xlMedium
        .RowHeight = 18
    End With
    With ws.Range(ws.Cells(r + 1, c), ws.Cells(r + 1, c + span - 1))
        .Merge
        If VarType(val) = vbString Then
            If Left(val, 1) = "=" Then .Formula = val Else .Value2 = val
        Else
            .Value2 = val
        End If
        .NumberFormat = fmt
        .Font.name = FONT_NM: .Font.Size = KpiFontSize(ws, c, span, fmt): .Font.Bold = True
        If gPrintLight Then
            .Font.color = clr: .Interior.color = Tint(clr, 0.92)
        Else
            .Font.color = RGB(255, 255, 255): .Interior.color = clr
        End If
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .RowHeight = 38
        .Borders.LineStyle = xlNone
        .Borders(xlEdgeLeft).LineStyle = xlContinuous: .Borders(xlEdgeLeft).color = RGB(255, 255, 255): .Borders(xlEdgeLeft).weight = xlMedium
        .Borders(xlEdgeRight).LineStyle = xlContinuous: .Borders(xlEdgeRight).color = RGB(255, 255, 255): .Borders(xlEdgeRight).weight = xlMedium
    End With
End Sub

'색을 흰색 쪽으로 섞기 (pct=0.85 → 85% 흰색)
Public Function Tint(ByVal clr As Long, ByVal pct As Double) As Long
    Dim r As Long, g As Long, b As Long
    r = clr Mod 256: g = (clr \ 256) Mod 256: b = (clr \ 65536) Mod 256
    r = r + (255 - r) * pct: g = g + (255 - g) * pct: b = b + (255 - b) * pct
    Tint = RGB(r, g, b)
End Function

'병합 셀은 '셀에 맞춤'이 안 되므로 열 너비에 맞춰 글꼴 크기를 계산
Public Function KpiFontSize(ws As Worksheet, ByVal c As Long, ByVal span As Long, ByVal fmt As String) As Single
    Dim W As Double, i As Long, L As Double, f As Double
    For i = c To c + span - 1: W = W + ws.Columns(i).ColumnWidth: Next i
    If InStr(fmt, "%") > 0 Then
        L = 6
    ElseIf InStr(fmt, "#,##0") > 0 Then
        L = 14                       '80,198,048,000 = 14자
    Else
        L = 8
    End If
    f = Int(11 * W / (L * 1.15))
    If f > 18 Then f = 18
    If f < 9 Then f = 9
    KpiFontSize = f
End Function

'KPI 카드를 왼쪽부터 자동 배치: 금액 카드는 열 너비 합이 28자 이상, 비율·건수는 12자 이상이 되도록 열을 병합
Public Sub KpiRow(ws As Worksheet, ByVal r As Long, titles As Variant, vals As Variant, fmts As Variant, clrs As Variant)
    Dim i As Long, c As Long, span As Long, W As Double, need As Double
    c = 1
    For i = LBound(titles) To UBound(titles)
        If InStr(fmts(i), "%") > 0 Then
            need = 12
        ElseIf InStr(fmts(i), "#,##0") > 0 Then
            need = 28
        Else
            need = 12
        End If
        span = 0: W = 0
        Do
            W = W + ws.Columns(c + span).ColumnWidth
            span = span + 1
        Loop While W < need And c + span <= 60
        DrawKpi ws, r, c, CStr(titles(i)), vals(i), CStr(fmts(i)), CLng(clrs(i)), span
        c = c + span
    Next i
End Sub

Public Sub 메인으로()
    If SheetExists(SH_MAIN) Then BW().Worksheets(SH_MAIN).Activate: Range("A1").Select
End Sub

'---------------------------------------------------------------- 슬라이서 (표/피벗 공용)
Public Sub AddSlicerTo(src As Object, ws As Worksheet, ByVal fld As String, ByVal scName As String, _
                       ByVal L As Single, ByVal T As Single, ByVal W As Single, ByVal h As Single, ByVal nCols As Long)
    Dim sc As SlicerCache, sl As Slicer
    On Error Resume Next
    Set sl = BW().SlicerCaches(scName).Slicers(scName & "_s")
    If Not sl Is Nothing Then                                   '기존 슬라이서 유지: 위치·크기만 재지정
        If sl.parent.name = ws.name Then
            sl.NumberOfColumns = nCols
            sl.top = T: sl.Left = L: sl.Width = W: sl.Height = h
            sl.Shape.Placement = xlFreeFloating
            sl.DisableMoveResizeUi = True
            Exit Sub
        End If
        sl.Delete: Set sl = Nothing
    End If
    Err.Clear
    Set sc = BW().SlicerCaches.Add2(src, fld, scName)
    If sc Is Nothing Then
        Err.Clear
        Set sc = BW().SlicerCaches.Add2(src, fld)     '이름 충돌 시 자동 이름으로 생성 후 변경
        If sc Is Nothing Then Exit Sub
        sc.name = scName
    End If
    Err.Clear
    On Error GoTo EH
    Set sl = sc.Slicers.Add(ws, , scName & "_s", fld, T, L, W, h)
    sl.NumberOfColumns = nCols
    sl.Style = "SlicerStyleLight1"
    sl.RowHeight = 18
    sl.top = T: sl.Left = L: sl.Width = W: sl.Height = h     '위치·크기 강제 지정
    sl.Shape.Placement = xlFreeFloating
    sl.DisableMoveResizeUi = True
    Exit Sub
EH:
End Sub

'prefix 로 시작하는 슬라이서 캐시 삭제(delIt=True) 또는 필터 해제(False)
Public Sub DeleteOrClearSlicers(ByVal prefix As String, ByVal delIt As Boolean)
    Dim i As Long, sc As SlicerCache
    On Error Resume Next
    For i = BW().SlicerCaches.Count To 1 Step -1
        Set sc = BW().SlicerCaches(i)
        If Left(sc.name, Len(prefix)) = prefix Then
            If delIt Then sc.Delete Else sc.ClearManualFilter
        End If
    Next i
    On Error GoTo 0
End Sub

Public Sub DeleteSheetIfExists(ByVal nm As String)
    If SheetExists(nm) Then
        Application.DisplayAlerts = False
        BW().Worksheets(nm).Delete
        Application.DisplayAlerts = True
    End If
End Sub

'인쇄: 인쇄영역(표 부분만, 슬라이서 제외) 설정 후 미리보기
Public Sub PrintSheetArea(ws As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long, ByVal lastCol As Long, ByVal titleRows As String)
    With ws.PageSetup
        .PrintArea = ws.Range(ws.Cells(firstRow, 1), ws.Cells(lastRow, lastCol)).Address
        .PrintTitleRows = titleRows
        .Orientation = xlLandscape
        .PaperSize = xlPaperA4
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .LeftMargin = Application.CentimetersToPoints(1): .RightMargin = Application.CentimetersToPoints(1)
        .TopMargin = Application.CentimetersToPoints(1.2): .BottomMargin = Application.CentimetersToPoints(1.2)
        .CenterHorizontally = True
        .CenterFooter = "&P / &N"
        .BlackAndWhite = False
    End With
    ws.PrintPreview
End Sub


