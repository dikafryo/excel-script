Attribute VB_Name = "modBoard"
Option Explicit
'=====================================================================
' modBoard : 메인 시트 (버튼 + KPI + 사업별 집행 표 + 오른쪽 슬라이서)
'   - 표(ListObject) → 슬라이서(팀/담당자/세부사업/단위과제카드). 슬라이서용 열은 표 끝에 숨김
'   - 요구서 금액 <> 장표 예산금액 이면 행 색 + 예산현액 빨간 글자 + 메모(요구서 금액/차이)  (V8에서 실제 구현)
'=====================================================================
Public Const MAIN_HDR As Long = 9
Public Const TBL_MAIN As String = "tbl메인"
Private Const SC_MAIN As String = "슬라이서_메인_"
Private Const BTN_TOP As Single = 4
Private Const BTN_H As Single = 22
Private Const BTN_BIG_H As Single = 27      'V8 시트 버튼 높이 (12pt 굵게 · 녹화 1080p 에서 16px 이상)

'메인 열
Private Const M_NO As Long = 1
Private Const M_SEB As Long = 2
Private Const M_NM As Long = 3
Private Const M_HY As Long = 4
Private Const M_EX As Long = 5
Private Const M_REM As Long = 6
Private Const M_RATE As Long = 7
Private Const M_BAR As Long = 8
Private Const M_ST As Long = 9
Private Const M_TEAM As Long = 10    '숨김 (슬라이서용)
Private Const M_OWN As Long = 11     '숨김
Private Const M_CARD As Long = 12    '숨김
Private Const M_MATCH As Long = 13   '숨김
Private Const M_LAST As Long = 13

'메인 시트 생성 + 버튼 배치
Public Sub SetupMainSheet()
    Dim ws As Worksheet, x0 As Single, x As Single, W As Single, g As Single, t1 As Single, t2 As Single, h As Single
    Dim nw As Single, fp As Single, fs As Single
    Set ws = GetSheet(SH_MAIN)
    ws.Tab.color = ClrBlue()
    If ws.Index <> 1 Then ws.Move before:=BW().Worksheets(1)
    'V8: 리본은 V7 그대로 두고, V8 기능은 이 시트 버튼으로 (녹화용 큰 글자 · 2줄 격자)
    ws.rows(1).RowHeight = 3
    ws.rows("2:3").RowHeight = 31
    ws.rows(4).RowHeight = 6
    RemoveBoardButtons ws                 '이전 버튼(V7 · 이전 V8 · 다른 추가기능 이름) 모두 지우고 다시 만듦
    h = BTN_BIG_H: g = 5: t1 = 4: t2 = t1 + h + 5
    W = 125: fp = 12: x0 = 4
    '1줄: 자료 불러오기 · 새로고침 · 산출내역 / 2줄: 추경 · 버전 정보
    x = x0
    AddButton ws, "bx8_req", "① 세출예산요구서", "요구서선택", x, t1, W, h, ClrBlue(), fp: x = x + W + g
    AddButton ws, "bx8_jang", "② 장표(102-B)", "장표선택", x, t1, W, h, ClrBlue(), fp: x = x + W + g
    AddButton ws, "bx8_refresh", "새로고침", "새로고침버튼", x, t1, W, h, ClrGreen(), fp: x = x + W + g
    AddButton ws, "bx8_calc", "세부산출내역", "세부산출내역", x, t1, W, h, ClrIndigo(), fp: x = x + W + g
    x = x0
    AddButton ws, "bx8_supp", "추경예산 추가", "추경예산추가", x, t2, W, h, ClrOrange(), fp: x = x + W + g
    AddButton ws, "bx8_undo", "마지막 추경 취소", "마지막추경취소", x, t2, W, h, ClrOrange(), fp: x = x + W + g
    AddButton ws, "bx8_check", "추경 점검표", "추경점검표", x, t2, W, h, ClrOrange(), fp: x = x + W + g
    AddButton ws, "bx8_about", "버전 정보", "버전정보", x, t2, W, h, ClrGray(), fp: x = x + W + g
    'V7 화면 이동 버튼은 오른쪽 작은 격자 (같은 높이 · 11pt)
    x0 = x + 11: nw = 106: fs = 11
    x = x0
    AddButton ws, "bx8_board", "예산보드", "예산보드", x, t1, nw, h, ClrNavy(), fs: x = x + nw + g
    AddButton ws, "bx8_detail", "세부현황", "세목별현황", x, t1, nw, h, ClrNavy(), fs: x = x + nw + g
    AddButton ws, "bx8_print", "인쇄", "인쇄", x, t1, nw, h, ClrNavy(), fs
    x = x0
    AddButton ws, "bx8_staff", "부서·업무담당자", "부서담당자", x, t2, nw, h, ClrNavy(), fs: x = x + nw + g
    AddButton ws, "bx8_card", "단위사업카드", "단위사업카드", x, t2, nw, h, ClrNavy(), fs
    FitMainSlicers ws                     '옛 판 보드: 슬라이서를 버튼 아래로
End Sub

'메인 슬라이서를 버튼 격자 아래 · 표 오른쪽으로 (BuildMain 과 같은 자리). 옛 판 보드에 V8 버튼만 넣을 때 겹침 방지
Private Sub FitMainSlicers(ws As Worksheet)
    Dim nm As Variant, sl As Slicer, sx As Single, sy As Single
    sx = ws.Cells(MAIN_HDR, M_LAST + 2).Left: sy = 68
    For Each nm In Array("상태", "팀", "담당자", "단위과제카드", "세부사업")
        Set sl = Nothing
        On Error Resume Next
        Set sl = BW().SlicerCaches(SC_MAIN & nm).Slicers(SC_MAIN & nm & "_s")
        On Error GoTo 0
        If Not sl Is Nothing Then
            If sl.parent.name = ws.name Then
                sl.top = sy: sl.Left = sx: sy = sy + sl.Height + 4
            End If
        End If
    Next nm
End Sub

'메인 시트의 보드 버튼 지우기: V7 이름(btn1~btn9, btnS*, btnC, btnT, btnD) · V8 이름(bx8_*) · 다른 추가기능 파일을 가리키는 버튼
'  슬라이서 · 표 · 메모는 건드리지 않음
Public Sub RemoveBoardButtons(ws As Worksheet)
    Dim i As Long, shp As Shape, nm As String, oa As String, kill As Boolean
    For i = ws.Shapes.Count To 1 Step -1
        Set shp = ws.Shapes(i)
        nm = shp.name: oa = "": kill = False
        On Error Resume Next
        oa = shp.OnAction
        On Error GoTo 0
        If shp.Type = msoAutoShape Or shp.Type = msoFormControl Then
            If nm Like "btn*" Or nm Like "bx8_*" Then kill = True
            If oa <> "" And InStr(oa, "!") > 0 Then kill = True
        End If
        If kill Then shp.Delete
    Next i
End Sub

'보드 버튼이 이 추가기능(V8)을 가리키는지 (V7에서 만든 보드면 False)
Public Function BoardButtonsCurrent(wb As Workbook) As Boolean
    Dim ws As Worksheet, oa As String
    On Error Resume Next
    Set ws = wb.Worksheets(SH_MAIN)
    If ws Is Nothing Then Exit Function
    oa = ws.Shapes("bx8_supp").OnAction
    On Error GoTo 0
    BoardButtonsCurrent = (OnActionBook(oa) = LCase$(ThisWorkbook.name)) And (oa Like "*!추경예산추가")
End Function

'버튼 OnAction 의 통합문서 이름 (Excel 은 '(폴더)\BX-v8.xlam'!매크로 처럼 전체 경로나 따옴표 없이 돌려줄 수 있음) → 파일 이름만, 소문자
Public Function OnActionBook(ByVal oa As String) As String
    Dim p As Long
    p = InStrRev(oa, "!")
    If p = 0 Then Exit Function
    oa = Replace(Left$(oa, p - 1), "'", "")
    p = InStrRev(oa, "\")
    If p > 0 Then oa = Mid$(oa, p + 1)
    OnActionBook = LCase$(oa)
End Function

Public Sub 메인슬라이서리셋()
    DeleteOrClearSlicers SC_MAIN, False
End Sub

'메인 보드 작성
Public Sub BuildMain()
    Dim ws As Worksheet, keys As Collection, i As Long, r As Long, it As clsProject
    Dim dept As String, arr As Variant, n As Long, lo As ListObject, rng As Range, nNoOwn As Long
    Dim sx As Single, sy As Single, lo0 As ListObject, nBad As Long, nTot As Long, fA As Long, fB As Long, fC As Long

    Set ws = GetSheet(SH_MAIN)
    SetupMainSheet
    gStep = "메인 초기화"
    DeleteOrClearSlicers SC_MAIN, False
    ws.Range(ws.rows(5), ws.rows(ws.rows.Count)).ClearComments
    ClearRowsKeepTables ws, 5                    '표 머리글·슬라이서 유지
    ws.Columns.hidden = False
    ws.Cells.Interior.color = RGB(255, 255, 255)

    '제목 (5행)
    dept = CfgGet("부서명")
    ws.Range("A5").Value2 = IIf(dept <> "", dept & " ", "") & "예산집행 보드"
    ws.Range("A5").Font.Size = 20: ws.Range("A5").Font.Bold = True: ws.Range("A5").Font.name = FONT_NM
    ws.Range("A5").Font.color = ClrNavy()
    ws.rows(5).RowHeight = 34
    ws.Range("D5").Value2 = "기준일 " & CfgGet("기준일") & "   |   갱신 " & NowText() & "   |   (단위: 원)"
    ws.Range("D5").Font.Size = 9: ws.Range("D5").Font.color = ClrGray(): ws.Range("D5").VerticalAlignment = xlBottom

    If Not ModelReady Then Exit Sub
    Set keys = SortedKeys()
    n = keys.Count
    If n = 0 Then Exit Sub

    '표 데이터
    gStep = "메인 표"
    ws.Cells(MAIN_HDR, 1).Resize(1, M_LAST).Value2 = Array("No", "세부사업", "사업항목", "예산현액", "집행액", "집행잔액", "집행률", "진행", "상태", _
        "팀/부서", "담당자", "단위과제카드", "매칭")
    ReDim arr(1 To n, 1 To M_LAST)
    For i = 1 To n
        Set it = gItems(keys(i))
        arr(i, M_NO) = i
        arr(i, M_SEB) = it.세부사업
        arr(i, M_NM) = it.사업항목
        arr(i, M_HY) = it.예산현액
        arr(i, M_EX) = it.집행액
        arr(i, M_REM) = it.잔액
        arr(i, M_RATE) = it.집행률
        arr(i, M_BAR) = it.집행률
        arr(i, M_ST) = it.상태
        arr(i, M_TEAM) = IIf(it.팀 = "", "(팀 미지정)", it.팀)
        arr(i, M_OWN) = IIf(it.담당자 = "", "(담당자 미지정)", it.담당자)
        arr(i, M_CARD) = IIf(it.과제카드 = "", "-", it.과제카드)
        arr(i, M_MATCH) = it.매칭
        If it.담당자 = "" Then nNoOwn = nNoOwn + 1
    Next i
    r = MAIN_HDR + 1
    ws.Cells(r, 1).Resize(n, M_LAST).Value2 = arr

    '엑셀 표
    Set rng = ws.Range(ws.Cells(MAIN_HDR, 1), ws.Cells(MAIN_HDR + n, M_LAST))
    Set lo = MakeTable(ws, rng, TBL_MAIN)
    StyleHeader lo.HeaderRowRange
    StyleBody lo.DataBodyRange
    With ws.Range(ws.Cells(r, M_HY), ws.Cells(r + n - 1, M_REM))
        .NumberFormat = "#,##0;[Red]-#,##0;""-""": .HorizontalAlignment = xlRight: .ShrinkToFit = True
    End With
    '집행률·진행 한 셀: % 숫자 + 그라데이션 막대 (80%↑ 파랑 / 60~80% 초록 / 60%↓ 빨강), 진행 열은 숨김
    With ws.Range(ws.Cells(r, M_RATE), ws.Cells(r + n - 1, M_RATE))
        .NumberFormat = "0.0%": .HorizontalAlignment = xlRight: .Font.Bold = True: .IndentLevel = 1
    End With
    For i = 1 To n
        BarByRate ws.Cells(r + i - 1, M_RATE), NzD(arr(i, M_RATE))
    Next i
    ws.Cells(MAIN_HDR, M_RATE).Value2 = "집행률·진행"
    ws.Range(ws.Cells(r, M_ST), ws.Cells(r + n - 1, M_ST)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(r, M_NO), ws.Cells(r + n - 1, M_NO)).HorizontalAlignment = xlCenter
    ColorStatus ws.Range(ws.Cells(r, M_ST), ws.Cells(r + n - 1, M_ST))
    'V8: 세출(요구서) 금액 <> 장표 예산금액 → 행 색 + 예산현액 주황 굵은 글자 + 메모 ('확인 필요' - 추경 등으로 정상일 수 있음) (설명에만 있던 기능 구현)
    nBad = 0
    For i = 1 To n
        Set it = gItems(keys(i))
        If it.매칭 = "일치" And IsMismatch(it.요구액, it.장표예산) Then
            nBad = nBad + 1
            ws.Range(ws.Cells(r + i - 1, M_NO), ws.Cells(r + i - 1, M_REM)).Interior.color = MismatchColor()
            With ws.Cells(r + i - 1, M_HY)
                .Font.color = ClrOrange(): .Font.Bold = True
                .AddComment "장표와 " & Replace(IIf(it.차수 = "", "본예산", it.차수), " ", "") & " 금액 확인 필요" & vbLf & _
                            "세출 파일(" & IIf(it.차수 = "", "본예산", it.차수) & "): " & Format(it.요구액, "#,##0") & "원" & vbLf & _
                            "장표 예산금액: " & Format(it.장표예산, "#,##0") & "원" & vbLf & _
                            "차이(세출-장표): " & Format(it.요구액 - it.장표예산, "#,##0") & "원"
                .Comment.Shape.Width = 230: .Comment.Shape.Height = 70
            End With
        End If
    Next i
    If JangLoaded() And SheetExists(SH_REQ) Then
        '맨 위 숫자 = 추경점검·세부산출내역과 같은 최종 상태 기준 (행 색은 장표 금액이 다른 사업만)
        nTot = FinalNoteCounts(fA, fB, fC)
        ws.Range("D5").Value2 = ws.Range("D5").Value2 & "   |   확인 필요 " & nTot & "건 (장표 대조 " & fA & "건 · 최신년도 산출내역 미반영 가능성 " & fB & "건 · 예산현황-산출내역 확인 " & fC & "건 · 최종 상태 기준 · 사유는 추경점검 비고)"
        If nTot > 0 Then ws.Range("D5").Characters(InStr(ws.Range("D5").Value2, "확인 필요"), Len("확인 필요 " & nTot & "건")).Font.color = ClrOrange()
    End If

    '합계 행
    lo.ShowTotals = True
    lo.ListColumns(M_NM).total.Value2 = "합계"
    lo.ListColumns(M_NO).TotalsCalculation = xlTotalsCalculationNone
    lo.ListColumns(M_HY).TotalsCalculation = xlTotalsCalculationSum
    lo.ListColumns(M_EX).TotalsCalculation = xlTotalsCalculationSum
    lo.ListColumns(M_REM).TotalsCalculation = xlTotalsCalculationSum
    lo.ListColumns(M_RATE).total.Formula = "=IFERROR(SUBTOTAL(109,[집행액])/SUBTOTAL(109,[예산현액]),0)"
    With lo.TotalsRowRange
        .Font.name = FONT_NM: .Font.Bold = True: .Font.color = ClrNavy()
        .Interior.color = ClrLight(): .ShrinkToFit = True
        .Borders(xlEdgeTop).LineStyle = xlContinuous: .Borders(xlEdgeTop).weight = xlMedium: .Borders(xlEdgeTop).color = ClrGray()
    End With
    ws.Range(lo.TotalsRowRange.Cells(1, M_HY), lo.TotalsRowRange.Cells(1, M_REM)).NumberFormat = "#,##0;[Red]-#,##0"
    lo.TotalsRowRange.Cells(1, M_RATE).NumberFormat = "0.0%"

    'KPI (7~8행)
    gStep = "메인 KPI"
    '열너비 / 숨김 (KPI 배치 전에 확정)
    ws.Columns(M_NO).ColumnWidth = 5: ws.Columns(M_SEB).ColumnWidth = 18: ws.Columns(M_NM).ColumnWidth = 36
    ws.Range(ws.Columns(M_HY), ws.Columns(M_REM)).ColumnWidth = 18
    ws.Columns(M_RATE).ColumnWidth = 24: ws.Columns(M_ST).ColumnWidth = 11
    ws.Columns(M_BAR).hidden = True
    ws.Range(ws.Columns(M_TEAM), ws.Columns(M_MATCH)).hidden = True
    ws.Columns(M_LAST + 1).ColumnWidth = 2

    KpiRow ws, 7, Array("예산현액", "집행액", "집행잔액", "집행률", "사업 수"), _
        Array("=SUBTOTAL(109," & TBL_MAIN & "[예산현액])", "=SUBTOTAL(109," & TBL_MAIN & "[집행액])", "=SUBTOTAL(109," & TBL_MAIN & "[집행잔액])", _
              "=IFERROR(SUBTOTAL(109," & TBL_MAIN & "[집행액])/SUBTOTAL(109," & TBL_MAIN & "[예산현액]),0)", "=SUBTOTAL(103," & TBL_MAIN & "[사업항목])"), _
        Array("#,##0", "#,##0", "#,##0", "0.0%", "0"), Array(ClrNavy(), ClrBlue(), ClrOrange(), ClrGreen(), ClrIndigo())
    ws.rows(6).RowHeight = 6

    '슬라이서: 표 오른쪽 세로 고정
    gStep = "메인 슬라이서"
    sx = ws.Cells(MAIN_HDR, M_LAST + 2).Left
    sy = 68                                      'V8: 버튼 격자(1~4행) 아래부터
    '상태 슬라이서: 1열(항목 세로 나열) - 오른쪽 세로 배치 맨 위 (버튼·KPI와 겹치지 않음)
    AddSlicerTo lo, ws, "상태", SC_MAIN & "상태", sx, sy, 220, 128, 1: sy = sy + 132
    AddSlicerTo lo, ws, "팀/부서", SC_MAIN & "팀", sx, sy, 220, 128, 1: sy = sy + 132
    AddSlicerTo lo, ws, "담당자", SC_MAIN & "담당자", sx, sy, 220, 150, 2: sy = sy + 154
    AddSlicerTo lo, ws, "단위과제카드", SC_MAIN & "단위과제카드", sx, sy, 220, 118, 1: sy = sy + 122
    AddSlicerTo lo, ws, "세부사업", SC_MAIN & "세부사업", sx, sy, 220, 130, 1

    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Range("A1").Select
    ws.Cells(MAIN_HDR + 1, M_HY).Select
    ActiveWindow.FreezePanes = True
    ActiveWindow.DisplayGridlines = False
    ws.Range("A1").Select
End Sub

'인쇄 (활성 시트의 표/보드 영역만)
Public Sub 인쇄()
    Dim ws As Worksheet, lo As ListObject, lastR As Long, lastC As Long
    Set ws = ActiveSheet
    On Error GoTo EH
    If ws.name = SH_CHART Then
        PrintChartBoard
        Exit Sub
    End If
    If ws.ListObjects.Count = 0 Then
        MsgBox "인쇄할 표가 없습니다.", vbExclamation
        Exit Sub
    End If
    Set lo = ws.ListObjects(1)
    lastR = lo.Range.row + lo.Range.rows.Count - 1
    lastC = lo.Range.Columns.Count
    PrintSheetArea ws, 5, lastR, lastC, "$" & lo.HeaderRowRange.row & ":$" & lo.HeaderRowRange.row
    Exit Sub
EH:
    MsgBox "인쇄 설정 오류: " & Err.Description, vbCritical
End Sub

Public Sub ColorStatus(rng As Range)
    Dim c As Range
    For Each c In rng.Cells
        Select Case Left(CStr(c.Value2), 1)
            Case "■": c.Font.color = ClrGreen(): c.Font.Bold = True
            Case "●": c.Font.color = ClrBlue()
            Case "▲": c.Font.color = ClrRed(): c.Font.Bold = True
            Case "△": c.Font.color = ClrRed()
            Case Else: c.Font.color = ClrGray()
        End Select
    Next c
End Sub

Public Sub ColorMatch(rng As Range)
    Dim c As Range
    For Each c In rng.Cells
        Select Case CStr(c.Value2)
            Case "일치", "지정", "요구서": c.Font.color = ClrGray()
            Case Else: c.Interior.color = RGB(255, 228, 226): c.Font.color = ClrRed()
        End Select
    Next c
End Sub


