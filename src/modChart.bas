Attribute VB_Name = "modChart"
Option Explicit
'=====================================================================
' modChart : 예산보드 (그래프형 대시보드, 슬라이서 연동)  ※ 팀 목록(TeamList) 포함
'   원본 표 tbl차트원본(차트자료 시트 - 표시 시트, 세목 행 단위)에 슬라이서(팀/부서·세부사업·단위과제카드) 연결
'   차트·KPI·세목 표는 모두 "보이는 행만" 집계하는 수식 → 슬라이서 선택 시 즉시 갱신
'   - KPI 카드
'   - 팀(부서)별 집행률 도넛 (팀당 1개)   - 세출세목별 지출현황 막대
'   - 사업항목 TOP10                      - 담당자(팀원)별 집행률  ← 팀 슬라이서 선택 시 팀원만 표시
'   - 세출세목(코드)별 지출현황 표
'   인쇄: 진한 배경(제목 띠·KPI)을 흰색으로 바꿔 인쇄 후 원상 복구
'=====================================================================
Public Const SH_CHART As String = "예산보드"
Public Const SH_CDATA As String = "차트자료"
Public Const TBL_CSRC As String = "tbl차트원본"
Private Const SC_CH As String = "슬라이서_차트_"
Private Const TOPN As Long = 10          '사업항목 TOP N
Private Const SEM_N As Long = 10         '세목 최대 표시 수
Private Const OWN_N As Long = 15         '담당자 최대 표시 수
'원본 표 열
Private Const S_TEAM As Long = 1, S_OWN As Long = 2, S_SEB As Long = 3, S_NM As Long = 4, S_CARD As Long = 5
Private Const S_CODE As Long = 6, S_SEM As Long = 7, S_HY As Long = 8, S_WON As Long = 9, S_EX As Long = 10, S_FIRST As Long = 11
Private Const S_NC As Long = 11
'집계 블록 시작 열 (_보드데이터)
Public Const B_TEAM As Long = 14   'N: 팀, 예산현액, 집행액, 잔액, 집행률, 제목
Public Const B_SEM As Long = 21    'U: 코드, 세목, 예산현액, 집행액, 잔액, 집행률
Public Const B_TOP As Long = 28    'AB: 사업항목, 예산현액, 집행액, 정렬키
Public Const B_OWN As Long = 32    'AF: 담당자, 예산현액, 집행액, 집행률, 정렬키
Public Const B_KPI As Long = 37    'AK: 예산현액, 원인행위액, 집행액, 잔액, 집행률, 사업수 (2행)
Public Const B_SEB As Long = 44    'AR: 세부사업, 예산현액, 집행액, 집행률
Public Const B_TOPS As Long = 62   'BJ: 사업 TOP 정렬 결과 (키, 사업항목, 예산현액, 집행액, 집행률) - 슬라이서 연동
Public Const B_OWNS As Long = 68   'BP: 담당자 정렬 결과 (키, 담당자, 예산현액, 집행액, 집행률)
Public gBoardEv As Object   ' clsBoardEvents (2.7m 빈 행 자동 숨김)
Public Const B_TEAMS As Long = 76  'BX: 팀 압축 블록 (키, 팀, 예산현액, 집행액, 집행잔액, 집행률) - 예산현액>0 인 팀만 위로 (2.7l)


Public Sub 예산보드()
    SpeedOn "예산보드 생성 중..."
    On Error GoTo EH
    If Not ModelReady Then BuildModel
    BuildChartBoard
    BW().Worksheets(SH_CHART).Activate
    SpeedOff
    Exit Sub
EH:
    SpeedOff
    On Error Resume Next
    BW().Worksheets(SH_CDATA).Visible = xlSheetHidden
    MsgBox "예산보드 오류 [" & gStep & "]: " & Err.Description & " (" & Err.Number & ")", vbCritical
End Sub

Public Sub 차트슬라이서리셋()
    DeleteOrClearSlicers SC_CH, False
End Sub

'인쇄: 밝은 모드로 다시 그려 인쇄 → 화면 모드로 복구 (슬라이서 선택은 초기화됨)
Public Sub PrintChartBoard()
    Dim ws As Worksheet, lr As Long
    On Error GoTo EH
    Set ws = BW().Worksheets(SH_CHART)
    lr = lastRow(ws, 1)
    BoardLight ws, 26, 3, 7, True                   '제목 띠·KPI 진한 배경 → 흰색 (다시 그리지 않음)
    With ws.PageSetup
        .PrintArea = ws.Range(ws.Cells(3, 1), ws.Cells(lr, 26)).Address
        .PrintTitleRows = ""
        .Orientation = xlLandscape: .PaperSize = xlPaperA4
        .Zoom = False: .FitToPagesWide = 1: .FitToPagesTall = False
        .LeftMargin = Application.CentimetersToPoints(0.8): .RightMargin = Application.CentimetersToPoints(0.8)
        .TopMargin = Application.CentimetersToPoints(1): .BottomMargin = Application.CentimetersToPoints(1)
        .CenterHorizontally = True: .CenterFooter = "&P / &N": .BlackAndWhite = False
    End With
    ws.PrintPreview
    BoardLight ws, 26, 3, 7, False
    Exit Sub
EH:
    BoardLight ws, 26, 3, 7, False
    MsgBox "인쇄 오류: " & Err.Description, vbCritical
End Sub

'보이는 행만 합계 (슬라이서 필터 반영) - 조건: 원본표[condCol] = 셀 참조
Private Function VisF(ByVal sumCol As String, Optional ByVal condCol As String = "", Optional ByVal condRef As String = "") As String
    Dim c As String
    c = "SUBTOTAL(109,OFFSET(" & TBL_CSRC & "[" & sumCol & "],ROW(" & TBL_CSRC & "[" & sumCol & "])-MIN(ROW(" & TBL_CSRC & "[" & sumCol & "])),0,1))"
    If condCol <> "" Then
        VisF = "=SUMPRODUCT((" & TBL_CSRC & "[" & condCol & "]=" & condRef & ")*" & c & ")"
    Else
        VisF = "=SUMPRODUCT(" & c & ")"
    End If
End Function

'상대 주소 문자열
Private Function A1(wd As Worksheet, ByVal r As Long, ByVal c As Long) As String
    A1 = wd.Cells(r, c).Address(False, False)
End Function

'_보드데이터 셀 참조 수식 문자열
Private Function DRef(wd As Worksheet, ByVal r As Long, ByVal c As Long) As String
    DRef = "='" & wd.name & "'!" & wd.Cells(r, c).Address(True, True)
End Function

'---------------------------------------------------------------- 원본 표 + 집계 블록 → _보드데이터
Private Function BuildChartData(ByRef nTeam As Long, ByRef nSem As Long, ByRef nTop As Long, ByRef nOwn As Long, Optional ByRef nSeb As Long) As Worksheet
    Dim wd As Worksheet, agg As Object, cnt As Long, teams As Collection, ti As Long
    Dim k As Variant, it As clsProject, sub1 As Object, v As Variant, T As Variant, t2 As Variant, keys As Collection
    Dim out() As Variant, x As Long, i As Long, sk As String, p As Long, code As String, nm As String, first As Long
    Dim dSem As Object, dOwn As Object, lo As ListObject, r As Long, a As String

    gStep = "예산보드 데이터"
    Set agg = ItemSubtotals(cnt)
    Set keys = SortedKeys()
    DeleteOrClearSlicers SC_CH, False           '슬라이서 유지 (필터만 해제)
    DeleteOrClearSlicers "슬라이서_보드_", False
    If SheetExists(SH_CDATA) Then BW().Worksheets(SH_CDATA).Visible = xlSheetVisible
    Set wd = GetSheet(SH_CDATA, True)
    wd.Tab.color = ClrGray()

    '--- 원본 표 (세목 행)
    ReDim out(1 To cnt + gItems.Count + 1, 1 To S_NC)
    Set dSem = CreateObject("Scripting.Dictionary"): Set dOwn = CreateObject("Scripting.Dictionary")
    For i = 1 To keys.Count
        Set it = gItems(keys(i))
        nm = IIf(it.담당자 = "", "(미지정)", it.담당자)
        If dOwn.Exists(nm) Then dOwn(nm) = dOwn(nm) + it.예산현액 Else dOwn.Add nm, it.예산현액
        Set sub1 = Nothing
        If Not agg Is Nothing Then
            If agg.Exists(it.장표키) Then Set sub1 = agg(it.장표키)
        End If
        If sub1 Is Nothing Then
            x = x + 1
            FillRow out, x, it, nm, "", "(세목 없음)", it.예산현액, it.원인행위, it.집행액, 1
        Else
            first = 1
            For Each v In sub1.keys
                t2 = sub1(v): sk = CStr(t2(0))
                p = InStr(sk, " ")
                If p > 0 Then code = Left(sk, p - 1): sk = Mid(sk, p + 1) Else code = sk
                x = x + 1
                FillRow out, x, it, nm, code, sk, t2(1), t2(3), t2(2), first
                first = 0
                If dSem.Exists(code) Then
                    T = dSem(code): T(1) = T(1) + t2(1): dSem(code) = T
                Else
                    dSem.Add code, Array(sk, t2(1))
                End If
            Next v
        End If
    Next i
    wd.Cells(1, 1).Resize(1, S_NC).Value2 = Array("팀/부서", "담당자", "세부사업", "사업항목", "단위과제카드", "세목코드", "세출세목", "예산현액", "원인행위액", "집행액", "사업첫행")
    If x = 0 Then x = 1
    wd.Cells(2, 1).Resize(x, S_NC).Value2 = out
    Set lo = MakeTable(wd, wd.Range(wd.Cells(1, 1), wd.Cells(x + 1, S_NC)), TBL_CSRC)

    '--- 팀 블록
    Set teams = TeamList()
    wd.Cells(1, B_TEAM).Resize(1, 6).Value2 = Array("팀", "예산현액", "집행액", "집행잔액", "집행률", "제목")
    nTeam = 0
    For ti = 1 To teams.Count
        If TeamHasItems(CStr(teams(ti))) Then
            nTeam = nTeam + 1: r = nTeam + 1
            wd.Cells(r, B_TEAM).Value2 = teams(ti)
            a = A1(wd, r, B_TEAM)
            wd.Cells(r, B_TEAM + 1).Formula = VisF("예산현액", "팀/부서", a)
            wd.Cells(r, B_TEAM + 2).Formula = VisF("집행액", "팀/부서", a)
            wd.Cells(r, B_TEAM + 3).Formula = "=MAX(" & A1(wd, r, B_TEAM + 1) & "-" & A1(wd, r, B_TEAM + 2) & ",0)"
            wd.Cells(r, B_TEAM + 4).Formula = "=IFERROR(" & A1(wd, r, B_TEAM + 2) & "/" & A1(wd, r, B_TEAM + 1) & ",0)"
            wd.Cells(r, B_TEAM + 5).Formula = "=" & a & "&CHAR(10)&TEXT(" & A1(wd, r, B_TEAM + 4) & ",""0.0%"")"
        End If
    Next ti

    '--- 세목 블록 (전체 예산현액 내림차순 고정)
    wd.Cells(1, B_SEM).Resize(1, 6).Value2 = Array("세목코드", "세출세목", "예산현액", "집행액", "집행잔액", "집행률")
    i = 0
    For Each k In dSem.keys
        i = i + 1: T = dSem(k)
        wd.Cells(i + 1, B_SEM).Value2 = k: wd.Cells(i + 1, B_SEM + 1).Value2 = T(0): wd.Cells(i + 1, B_SEM + 2).Value2 = T(1)
    Next k
    If i > 1 Then wd.Range(wd.Cells(2, B_SEM), wd.Cells(i + 1, B_SEM + 2)).Sort Key1:=wd.Cells(2, B_SEM + 2), Order1:=xlDescending, Header:=xlNo
    nSem = IIf(i > SEM_N, SEM_N, i)
    If i > nSem Then wd.Range(wd.Cells(nSem + 2, B_SEM), wd.Cells(i + 1, B_SEM + 2)).ClearContents
    For r = 2 To nSem + 1
        a = A1(wd, r, B_SEM)
        wd.Cells(r, B_SEM + 2).Formula = VisF("예산현액", "세목코드", a)
        wd.Cells(r, B_SEM + 3).Formula = VisF("집행액", "세목코드", a)
        wd.Cells(r, B_SEM + 4).Formula = "=" & A1(wd, r, B_SEM + 2) & "-" & A1(wd, r, B_SEM + 3)
        wd.Cells(r, B_SEM + 5).Formula = "=IFERROR(" & A1(wd, r, B_SEM + 3) & "/" & A1(wd, r, B_SEM + 2) & ",0)"
    Next r

    '--- 사업항목 TOP N (전체 예산현액 내림차순 고정)
    wd.Cells(1, B_TOP).Resize(1, 3).Value2 = Array("사업항목", "예산현액", "집행액")
    i = 0
    For ti = 1 To keys.Count
        Set it = gItems(keys(ti))
        If it.예산현액 > 0 Then
            i = i + 1
            wd.Cells(i + 1, B_TOP).Value2 = it.사업항목: wd.Cells(i + 1, B_TOP + 1).Value2 = it.예산현액
        End If
    Next ti
    If i > 1 Then wd.Range(wd.Cells(2, B_TOP), wd.Cells(i + 1, B_TOP + 1)).Sort Key1:=wd.Cells(2, B_TOP + 1), Order1:=xlDescending, Header:=xlNo
    nTop = IIf(i > TOPN, TOPN, i)
    wd.Cells(1, B_TOP + 3).Value2 = "정렬키"
    For r = 2 To i + 1
        a = A1(wd, r, B_TOP)
        wd.Cells(r, B_TOP + 1).Formula = VisF("예산현액", "사업항목", a)
        wd.Cells(r, B_TOP + 2).Formula = VisF("집행액", "사업항목", a)
        wd.Cells(r, B_TOP + 3).Formula = "=IF(" & A1(wd, r, B_TOP + 1) & "<=0,-1," & A1(wd, r, B_TOP + 1) & "+(5000-ROW())/10000)"
    Next r
    SortedBlock wd, B_TOP, i, B_TOPS, nTop, "예산현액"

    '--- 담당자 블록 (전체 예산현액 내림차순 고정, 필터로 예산 0이면 #N/A → 막대 미표시)
    wd.Cells(1, B_OWN).Resize(1, 4).Value2 = Array("담당자", "예산현액", "집행액", "집행률")
    i = 0
    For Each k In dOwn.keys
        i = i + 1
        wd.Cells(i + 1, B_OWN).Value2 = k: wd.Cells(i + 1, B_OWN + 1).Value2 = dOwn(k)
    Next k
    If i > 1 Then wd.Range(wd.Cells(2, B_OWN), wd.Cells(i + 1, B_OWN + 1)).Sort Key1:=wd.Cells(2, B_OWN + 1), Order1:=xlDescending, Header:=xlNo
    nOwn = IIf(i > OWN_N, OWN_N, i)
    wd.Cells(1, B_OWN + 4).Value2 = "정렬키"
    For r = 2 To i + 1
        a = A1(wd, r, B_OWN)
        wd.Cells(r, B_OWN + 1).Formula = VisF("예산현액", "담당자", a)
        wd.Cells(r, B_OWN + 2).Formula = VisF("집행액", "담당자", a)
        wd.Cells(r, B_OWN + 3).Formula = "=IF(" & A1(wd, r, B_OWN + 1) & "=0,NA()," & A1(wd, r, B_OWN + 2) & "/" & A1(wd, r, B_OWN + 1) & ")"
        wd.Cells(r, B_OWN + 4).Formula = "=IF(" & A1(wd, r, B_OWN + 1) & "<=0,-1,IFERROR(" & A1(wd, r, B_OWN + 3) & ",0)+(5000-ROW())/1000000)"
    Next r
    SortedBlock wd, B_OWN, i, B_OWNS, nOwn, "집행률"

    '--- KPI 블록
    wd.Cells(1, B_KPI).Resize(1, 6).Value2 = Array("예산현액", "원인행위액", "집행액", "집행잔액", "집행률", "사업수")
    wd.Cells(2, B_KPI).Formula = VisF("예산현액")
    wd.Cells(2, B_KPI + 1).Formula = VisF("원인행위액")
    wd.Cells(2, B_KPI + 2).Formula = VisF("집행액")
    wd.Cells(2, B_KPI + 3).Formula = "=" & A1(wd, 2, B_KPI) & "-" & A1(wd, 2, B_KPI + 2)
    wd.Cells(2, B_KPI + 4).Formula = "=IFERROR(" & A1(wd, 2, B_KPI + 2) & "/" & A1(wd, 2, B_KPI) & ",0)"
    wd.Cells(2, B_KPI + 5).Formula = VisF("사업첫행")

    '--- 세부사업 블록 (전체 예산현액 내림차순 고정)
    wd.Cells(1, B_SEB).Resize(1, 4).Value2 = Array("세부사업", "예산현액", "집행액", "집행률")
    Set dOwn = CreateObject("Scripting.Dictionary")
    For Each k In gItems.keys
        Set it = gItems(k)
        If dOwn.Exists(it.세부사업) Then dOwn(it.세부사업) = dOwn(it.세부사업) + it.예산현액 Else dOwn.Add it.세부사업, it.예산현액
    Next k
    i = 0
    For Each k In dOwn.keys
        i = i + 1
        wd.Cells(i + 1, B_SEB).Value2 = k: wd.Cells(i + 1, B_SEB + 1).Value2 = dOwn(k)
    Next k
    If i > 1 Then wd.Range(wd.Cells(2, B_SEB), wd.Cells(i + 1, B_SEB + 1)).Sort Key1:=wd.Cells(2, B_SEB + 1), Order1:=xlDescending, Header:=xlNo
    nSeb = i
    For r = 2 To nSeb + 1
        a = A1(wd, r, B_SEB)
        wd.Cells(r, B_SEB + 1).Formula = VisF("예산현액", "세부사업", a)
        wd.Cells(r, B_SEB + 2).Formula = VisF("집행액", "세부사업", a)
        wd.Cells(r, B_SEB + 3).Formula = "=IF(" & A1(wd, r, B_SEB + 1) & "=0,NA()," & A1(wd, r, B_SEB + 2) & "/" & A1(wd, r, B_SEB + 1) & ")"
    Next r

    wd.Calculate                        '표시용 값(막대 색 결정) 1회 계산
    wd.Columns("A:AV").ColumnWidth = 12
    wd.Range("A1").Resize(1, S_NC).Font.Bold = True
    On Error Resume Next
    If SheetExists(SH_CHART) Then wd.Move After:=BW().Worksheets(SH_CHART)
    On Error GoTo 0
    wd.Visible = xlSheetHidden          '차트자료 시트 숨김 (연결 그림·수식은 숨김 상태에서도 갱신)
    Set BuildChartData = wd
End Function



'집행률(%) 가로 막대 서식
Private Sub PctBarStyle(ch As Chart, ByVal clr As Long)
    ch.HasLegend = False: ch.HasTitle = False
    ch.SeriesCollection(1).Format.Fill.ForeColor.RGB = clr
    ch.SeriesCollection(1).HasDataLabels = True
    With ch.SeriesCollection(1).DataLabels
        .NumberFormat = "0.0%": .Font.Size = 8: .Font.name = FONT_NM: .Position = xlLabelPositionOutsideEnd
    End With
    ch.ChartGroups(1).GapWidth = 60
    ch.DisplayBlanksAs = xlNotPlotted
    ch.Axes(xlCategory).ReversePlotOrder = True
    ch.Axes(xlCategory).TickLabels.Font.Size = 8: ch.Axes(xlCategory).TickLabels.Font.name = FONT_NM
    ch.Axes(xlValue).MinimumScale = 0: ch.Axes(xlValue).MaximumScale = 1
    ch.Axes(xlValue).TickLabels.NumberFormat = "0%": ch.Axes(xlValue).TickLabels.Font.Size = 8
    ch.Axes(xlValue).MajorGridlines.Format.line.ForeColor.RGB = ClrHead(): ch.Axes(xlValue).Format.line.Visible = msoFalse
    ch.Axes(xlCategory).Crosses = xlMaximum
    ch.Axes(xlCategory).Format.line.ForeColor.RGB = ClrHead()
End Sub




'정렬 결과 블록: 원본 블록(이름, 예산현액, 집행액, [집행률], 정렬키)에서 LARGE 로 상위 n행 추출 (슬라이서 필터 반영, 0이면 빈칸)
'  결과 열: 키 / 이름 / 예산현액 / 집행액 / 집행률
Private Sub SortedBlock(wd As Worksheet, ByVal srcCol As Long, ByVal nSrc As Long, ByVal dstCol As Long, ByVal n As Long, ByVal sortBy As String)
    Dim r As Long, keyCol As Long, kRng As String, nmRng As String, hyRng As String, exRng As String, kc As String
    If sortBy = "집행률" Then keyCol = srcCol + 4 Else keyCol = srcCol + 3
    kRng = wd.Range(wd.Cells(2, keyCol), wd.Cells(nSrc + 1, keyCol)).Address(True, True)
    nmRng = wd.Range(wd.Cells(2, srcCol), wd.Cells(nSrc + 1, srcCol)).Address(True, True)
    hyRng = wd.Range(wd.Cells(2, srcCol + 1), wd.Cells(nSrc + 1, srcCol + 1)).Address(True, True)
    exRng = wd.Range(wd.Cells(2, srcCol + 2), wd.Cells(nSrc + 1, srcCol + 2)).Address(True, True)
    wd.Cells(1, dstCol).Resize(1, 5).Value2 = Array("키", "이름", "예산현액", "집행액", "집행률")
    For r = 2 To n + 1
        kc = A1(wd, r, dstCol)
        wd.Cells(r, dstCol).Formula = "=IFERROR(LARGE(" & kRng & "," & (r - 1) & "),-1)"
        wd.Cells(r, dstCol + 1).Formula = "=IF(" & kc & "<0,"""",INDEX(" & nmRng & ",MATCH(" & kc & "," & kRng & ",0)))"
        wd.Cells(r, dstCol + 2).Formula = "=IF(" & kc & "<0,"""",INDEX(" & hyRng & ",MATCH(" & kc & "," & kRng & ",0)))"
        wd.Cells(r, dstCol + 3).Formula = "=IF(" & kc & "<0,"""",INDEX(" & exRng & ",MATCH(" & kc & "," & kRng & ",0)))"
        wd.Cells(r, dstCol + 4).Formula = "=IF(" & kc & "<0,"""",IFERROR(" & A1(wd, r, dstCol + 3) & "/" & A1(wd, r, dstCol + 2) & ",0))"
    Next r
End Sub

'팀 목록: 부서담당자 시트의 팀 순서 → 나머지 → (팀 미지정) 마지막
Public Function TeamList() As Collection
    Dim col As New Collection, d As Object, k As Variant, it As clsProject, T As String, ws As Worksheet, r As Long
    Set d = CreateObject("Scripting.Dictionary")
    If SheetExists(SH_STAFF) Then
        Set ws = BW().Worksheets(SH_STAFF)
        For r = 5 To lastRow(ws, 1)
            T = CleanText(ws.Cells(r, 2).Value2)
            If T <> "" Then If Not d.Exists(T) Then d.Add T, 1: col.Add T
        Next r
    End If
    For Each k In gItems.keys
        Set it = gItems(k)
        T = IIf(it.팀 = "", "(팀 미지정)", it.팀)
        If T <> "(팀 미지정)" Then If Not d.Exists(T) Then d.Add T, 1: col.Add T
    Next k
    col.Add "(팀 미지정)"
    Set TeamList = col
End Function

Private Sub FillRow(out() As Variant, ByVal x As Long, it As clsProject, ByVal own As String, ByVal code As String, ByVal sem As String, _
                    ByVal hy As Double, ByVal won As Double, ByVal ex As Double, ByVal first As Long)
    out(x, S_TEAM) = IIf(it.팀 = "", "(팀 미지정)", it.팀)
    out(x, S_OWN) = own
    out(x, S_SEB) = it.세부사업: out(x, S_NM) = it.사업항목
    out(x, S_CARD) = IIf(it.과제카드 = "", "(미지정)", it.과제카드)
    out(x, S_CODE) = code: out(x, S_SEM) = sem
    out(x, S_HY) = hy: out(x, S_WON) = won: out(x, S_EX) = ex
    out(x, S_FIRST) = first
End Sub

Private Function TeamHasItems(ByVal team As String) As Boolean
    Dim k As Variant, it As clsProject
    For Each k In gItems.keys
        Set it = gItems(k)
        If IIf(it.팀 = "", "(팀 미지정)", it.팀) = team Then TeamHasItems = True: Exit Function
    Next k
End Function

'---------------------------------------------------------------- 보드 그리기
Public Sub BuildChartBoard()
    Dim ws As Worksheet, wd As Worksheet, nTeam As Long, nSem As Long, nTop As Long, nOwn As Long, nSeb As Long, lo As ListObject
    Dim dept As String, L As Single, T As Single, W As Single, h As Single, i As Long, perRow As Long, dw As Single
    Dim rTop As Long, r As Long, co As ChartObject, ch As Chart, secTop As Single, sx As Single

    If Not ModelReady Then Exit Sub
    Set wd = BuildChartData(nTeam, nSem, nTop, nOwn, nSeb)
    Set lo = wd.ListObjects(TBL_CSRC)

    gStep = "예산보드 시트"
    Set ws = GetSheet(SH_CHART, True)
    ws.Tab.color = ClrIndigo()
    ws.Cells.Interior.color = RGB(255, 255, 255)
    ws.Cells.Font.name = FONT_NM
    ws.Columns("A:Z").ColumnWidth = 4.5            '26열 격자 (A4 가로 1페이지)
    '1행: 버튼 / 3행~: 내용 / 슬라이서: 표 오른쪽(AB열~) 세로 배치
    ws.rows(1).RowHeight = 26: ws.rows(2).RowHeight = 4
    AddHomeButton ws, 4, 3
    AddButton ws, "btnPrint", "인쇄", "인쇄", 68, 3, 50, 20, ClrNavy()
    AddButton ws, "btnRefresh", "새로고침(리셋)", "새로고침버튼", 122, 3, 90, 20, ClrGreen()
    AddButton ws, "btnDetail", "세부현황", "세목별현황", 216, 3, 66, 20, ClrIndigo()

    gStep = "예산보드 슬라이서"
    ws.Columns("AA").ColumnWidth = 2
    sx = ws.Columns(28).Left + 2
    AddSlicerTo lo, ws, "팀/부서", SC_CH & "팀", sx, ws.rows(3).top, 230, 120, 2
    AddSlicerTo lo, ws, "세부사업", SC_CH & "세부사업", sx, ws.rows(3).top + 126, 230, 230, 2
    AddSlicerTo lo, ws, "단위과제카드", SC_CH & "단위과제카드", sx, ws.rows(3).top + 362, 230, 230, 2
    AddSlicerTo lo, ws, "사업항목", SC_CH & "사업명", sx, ws.rows(3).top + 598, 230, 320, 1   '09-27 사업명 슬라이서
    On Error Resume Next
    Dim scx As SlicerCache, slx As Slicer
    For Each scx In BW().SlicerCaches
        For Each slx In scx.Slicers
            If slx.name = SC_CH & "사업명_s" Then slx.caption = "사업명"
        Next slx
    Next scx
    Err.Clear
    On Error GoTo 0

    '제목 띠 (화면: 진한 남색 / 인쇄: 흰색)
    dept = CfgGet("부서명")
    ws.Range("A3:Z4").Interior.color = IIf(gPrintLight, RGB(255, 255, 255), ClrNavy())
    With ws.Range("A3")
        .Value2 = IIf(dept <> "", dept & " ", "") & "예산 집행 대시보드"
        .Font.Size = 18: .Font.Bold = True: .Font.color = IIf(gPrintLight, ClrNavy(), RGB(255, 255, 255)): .IndentLevel = 1
    End With
    With ws.Range("A4")
        .Value2 = "기준일 " & CfgGet("기준일") & "   |   갱신 " & NowText() & "   |   단위: 원   |   슬라이서(팀/부서·세부사업·단위과제카드·사업명) 선택 시 모든 차트·표 연동"
        .Font.Size = 9: .Font.color = IIf(gPrintLight, ClrGray(), RGB(220, 220, 230)): .IndentLevel = 1
    End With
    ws.rows(3).RowHeight = 30: ws.rows(4).RowHeight = 16

    'KPI (수식 → 슬라이서 연동)
    gStep = "예산보드 KPI"
    ws.rows(5).RowHeight = 6
    DrawKpi ws, 6, 1, "예산현액", DRef(wd, 2, B_KPI), "#,##0", ClrNavy(), 5
    DrawKpi ws, 6, 6, "원인행위액", DRef(wd, 2, B_KPI + 1), "#,##0", ClrGray(), 5
    DrawKpi ws, 6, 11, "집행액", DRef(wd, 2, B_KPI + 2), "#,##0", ClrBlue(), 5
    DrawKpi ws, 6, 16, "집행잔액", DRef(wd, 2, B_KPI + 3), "#,##0", ClrOrange(), 5
    DrawKpi ws, 6, 21, "집행률", DRef(wd, 2, B_KPI + 4), "0.0%", ClrGreen(), 3
    DrawKpi ws, 6, 24, "사업 수", DRef(wd, 2, B_KPI + 5), "0", ClrIndigo(), 3
    ws.rows(8).RowHeight = 8

    '--- 1구역: 팀(부서)별 집행 현황 표 (2.7l: 그래프 대신 표 ? 팀 슬라이서로 고르면 그 팀만 남고 빈 줄이 생기지 않음)
    gStep = "예산보드 팀별 표"
    rTop = 9
    SectionTitle ws, rTop, 1, 26, "팀(부서)별 집행 현황  (팀/부서 슬라이서를 누르면 그 팀만 표시)"
    TeamCompactBlock wd, nTeam
    r = rTop + 1
    GridRow ws, r, 1, Array(6, 4, 4, 4, 2, 6), Array("팀/부서", "예산현액", "집행액", "집행잔액", "집행률", "진행"), "hhhhhh"
    For i = 1 To nTeam
        r = r + 1
        GridRow ws, r, 1, Array(6, 4, 4, 4, 2, 6), Array(DRef(wd, i + 1, B_TEAMS + 1), DRef(wd, i + 1, B_TEAMS + 2), DRef(wd, i + 1, B_TEAMS + 3), _
                DRef(wd, i + 1, B_TEAMS + 4), DRef(wd, i + 1, B_TEAMS + 5), DRef(wd, i + 1, B_TEAMS + 5)), "tnnnpb"
        DynBar ws.Cells(r, 21)
    Next i
    BX_SetHideName ws, "BX_HIDE_TEAM", ws.Range(ws.Cells(rTop + 2, 1), ws.Cells(r, 1))
    rTop = r + 2

    '--- 2구역: 세출세목(코드)별 지출현황 표 (수식 → 슬라이서 연동)
    gStep = "예산보드 세목 표"
    SectionTitle ws, rTop, 1, 26, "세출세목(코드)별 지출현황 표"
    r = rTop + 1
    GridRow ws, r, 1, Array(2, 5, 4, 4, 4, 7), Array("세목코드", "세출세목", "예산현액", "집행액", "집행잔액", "집행률·진행"), "hhhhhh"
    For i = 1 To nSem
        r = r + 1
        GridRow ws, r, 1, Array(2, 5, 4, 4, 4, 7), Array(DRef(wd, i + 1, B_SEM), DRef(wd, i + 1, B_SEM + 1), DRef(wd, i + 1, B_SEM + 2), _
                DRef(wd, i + 1, B_SEM + 3), DRef(wd, i + 1, B_SEM + 4), DRef(wd, i + 1, B_SEM + 5)), "ctnnng"
        BarByRate ws.Range(ws.Cells(r, 20), ws.Cells(r, 26)), NzD(wd.Cells(i + 1, B_SEM + 5).Value2)
    Next i
    r = r + 1
    GridRow ws, r, 1, Array(2, 5, 4, 4, 4, 7), Array("", "합계", DRef(wd, 2, B_KPI), DRef(wd, 2, B_KPI + 2), DRef(wd, 2, B_KPI + 3), DRef(wd, 2, B_KPI + 4)), "ctnnng"
    BarByRate ws.Range(ws.Cells(r, 20), ws.Cells(r, 26)), NzD(wd.Cells(2, B_KPI + 4).Value2)
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, 26))
        .Font.Bold = True: .Interior.color = ClrLight()
        .Borders(xlEdgeTop).LineStyle = xlContinuous: .Borders(xlEdgeTop).weight = xlMedium: .Borders(xlEdgeTop).color = ClrGray()
    End With
    rTop = r + 2

    '--- 3구역: 사업항목 TOP 표 (왼쪽) + 담당자(팀원)별 집행률 표 (오른쪽) - 슬라이서 연동, 보이는 값 기준 정렬
    gStep = "예산보드 사업 TOP 표"
    SectionTitle ws, rTop, 1, 12, "사업항목 TOP" & nTop & "  (예산현액 순)"
    SectionTitle ws, rTop, 14, 13, "담당자(팀원)별 집행률  (팀 선택 시 팀원만, 집행률 순)"
    r = rTop + 1
    GridRow ws, r, 1, Array(1, 4, 3, 4), Array("순위", "사업항목", "예산현액", "집행률·진행"), "hhhh"
    GridRow ws, r, 14, Array(3, 3, 3, 4), Array("담당자", "예산현액", "집행액", "집행률·진행"), "hhhh"
    Dim nRows As Long
    nRows = IIf(nTop > nOwn, nTop, nOwn)
    For i = 1 To nRows
        r = r + 1
        If i <= nTop Then
            GridRow ws, r, 1, Array(1, 4, 3, 4), Array(i, DRef(wd, i + 1, B_TOPS + 1), DRef(wd, i + 1, B_TOPS + 2), DRef(wd, i + 1, B_TOPS + 4)), "ctng"
            BarByRate ws.Range(ws.Cells(r, 9), ws.Cells(r, 12)), NzD(wd.Cells(i + 1, B_TOPS + 4).Value2)
        End If
        If i <= nOwn Then
            GridRow ws, r, 14, Array(3, 3, 3, 4), Array(DRef(wd, i + 1, B_OWNS + 1), DRef(wd, i + 1, B_OWNS + 2), DRef(wd, i + 1, B_OWNS + 3), DRef(wd, i + 1, B_OWNS + 4)), "tnng"
            BarByRate ws.Range(ws.Cells(r, 23), ws.Cells(r, 26)), NzD(wd.Cells(i + 1, B_OWNS + 4).Value2)
        End If
    Next i
    BX_SetHideName ws, "BX_HIDE_TOP", ws.Range(ws.Cells(r - nRows + 1, 2), ws.Cells(r, 2))
    BX_SetHideName ws, "BX_HIDE_OWN", ws.Range(ws.Cells(r - nRows + 1, 14), ws.Cells(r, 14))
    BX_EnsureBoardEvents
    ws.Cells(r + 1, 1).Value2 = " "        'LastRow 기준점

    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.FreezePanes = False
    ws.Range("A1").Select
End Sub

'가로 막대(예산 vs 집행 겹침) 공통 서식
Private Sub BarStyle(ch As Chart, ByVal c1 As Long, ByVal c2 As Long, ByVal numFmt As String, ByVal legend As Boolean)
    ch.HasTitle = False
    ch.HasLegend = legend
    If legend Then ch.legend.Position = xlLegendPositionTop: ch.legend.Font.Size = 8: ch.legend.Font.name = FONT_NM
    ch.SeriesCollection(1).Format.Fill.ForeColor.RGB = c1
    If ch.SeriesCollection.Count >= 2 Then ch.SeriesCollection(2).Format.Fill.ForeColor.RGB = c2
    ch.ChartGroups(1).Overlap = 100: ch.ChartGroups(1).GapWidth = 45
    ch.Axes(xlCategory).ReversePlotOrder = True
    ch.Axes(xlCategory).TickLabels.Font.Size = 8: ch.Axes(xlCategory).TickLabels.Font.name = FONT_NM
    ch.Axes(xlValue).TickLabels.NumberFormat = numFmt: ch.Axes(xlValue).TickLabels.Font.Size = 8
    ch.Axes(xlValue).MajorGridlines.Format.line.ForeColor.RGB = ClrHead(): ch.Axes(xlValue).Format.line.Visible = msoFalse
    ch.Axes(xlCategory).Crosses = xlMaximum
    ch.Axes(xlCategory).Format.line.ForeColor.RGB = ClrHead()
End Sub

'빈 행 자동 숨김용 이름 등록 (2.7m)
Private Sub BX_SetHideName(ws As Worksheet, ByVal nm As String, rg As Range)
    On Error Resume Next
    ws.parent.names(nm).Delete
    ws.parent.names.Add name:=nm, RefersTo:="='" & ws.name & "'!" & rg.Address(True, True)
    On Error GoTo 0
End Sub

'이벤트 클래스 연결 (2.7m): 리본 버튼 실행 때마다 확인해 두면 슬라이서 클릭 → 재계산 → 빈 행 숨김이 동작
Public Sub BX_EnsureBoardEvents()
    On Error Resume Next
    If gBoardEv Is Nothing Then
        Set gBoardEv = New clsBoardEvents
        Set gBoardEv.App = Application
    End If
    On Error GoTo 0
End Sub

'팀 압축 블록 (2.7l): 예산현액>0 인 팀의 행 번호를 키로 SMALL 로 위에서부터 모은다 → 슬라이서로 고른 팀만 표에 남음
Private Sub TeamCompactBlock(wd As Worksheet, ByVal nTeam As Long)
    Dim r As Long, c As Long, kRng As String, kc As String, colRng As String
    wd.Cells(1, B_TEAM + 6).Value2 = "키"
    For r = 2 To nTeam + 1
        wd.Cells(r, B_TEAM + 6).Formula = "=IF(" & A1(wd, r, B_TEAM + 1) & ">0," & r & ","""")"
    Next r
    kRng = wd.Range(wd.Cells(2, B_TEAM + 6), wd.Cells(nTeam + 1, B_TEAM + 6)).Address(True, True)
    wd.Cells(1, B_TEAMS).Resize(1, 6).Value2 = Array("키", "팀", "예산현액", "집행액", "집행잔액", "집행률")
    For r = 2 To nTeam + 1
        kc = A1(wd, r, B_TEAMS)
        wd.Cells(r, B_TEAMS).Formula = "=IFERROR(SMALL(" & kRng & "," & (r - 1) & "),-1)"
        For c = 0 To 4
            colRng = wd.Range(wd.Cells(2, B_TEAM + c), wd.Cells(nTeam + 1, B_TEAM + c)).Address(True, True)
            wd.Cells(r, B_TEAMS + 1 + c).Formula = "=IF(" & kc & "<0,"""",INDEX(" & colRng & "," & kc & "-1))"
        Next c
    Next r
End Sub

'진행 막대 (2.7l): 수식 값에 따라 움직이는 데이터 막대(값 숨김) ? 슬라이서 연동
Private Sub DynBar(rg As Range)
    Dim db As Databar
    rg.FormatConditions.Delete
    Set db = rg.FormatConditions.AddDatabar
    db.MinPoint.Modify newtype:=xlConditionValueNumber, newValue:=0
    db.MaxPoint.Modify newtype:=xlConditionValueNumber, newValue:=1
    db.BarColor.color = ClrBlue()
    db.BarFillType = xlDataBarFillGradient
    db.BarBorder.Type = xlDataBarBorderSolid
    db.BarBorder.color.color = ClrBlue()
    db.ShowValue = False
End Sub

'구역 제목 (열 c부터 span열 병합, 왼쪽 인디고 바)
Private Sub SectionTitle(ws As Worksheet, ByVal r As Long, ByVal c As Long, ByVal span As Long, ByVal txt As String)
    With ws.Range(ws.Cells(r, c), ws.Cells(r, c + span - 1))
        .Merge
        .Value2 = "  " & txt
        .Font.Size = 11: .Font.Bold = True: .Font.color = ClrNavy(): .VerticalAlignment = xlCenter
        .Borders(xlEdgeLeft).LineStyle = xlContinuous: .Borders(xlEdgeLeft).weight = xlThick: .Borders(xlEdgeLeft).color = ClrIndigo()
        .Borders(xlEdgeBottom).LineStyle = xlContinuous: .Borders(xlEdgeBottom).color = ClrHead()
    End With
    ws.rows(r).RowHeight = 22
End Sub


'격자 표 한 행: c0 열부터 spans 만큼 병합, kinds 는 칸별 종류 문자열
'  h=머리글 c=가운데 텍스트 t=왼쪽 텍스트 n=금액 p=퍼센트(색조) b=진행막대(값 숨김) g=퍼센트+그라데이션 막대(한 셀, BarByRate 로 색 지정)
'  값이 "=" 로 시작하면 수식. 서식은 행 단위로만 적용(병합 열에 세로 범위 서식을 걸면 병합이 깨지는 문제 방지)
Public Sub GridRow(ws As Worksheet, ByVal r As Long, ByVal c0 As Long, spans As Variant, vals As Variant, ByVal kinds As String)
    Dim i As Long, c As Long, rg As Range, k As String
    c = c0
    For i = LBound(spans) To UBound(spans)
        Set rg = ws.Range(ws.Cells(r, c), ws.Cells(r, c + spans(i) - 1))
        rg.UnMerge
        rg.Merge
        If VarType(vals(i)) = vbString Then
            If Left(vals(i), 1) = "=" Then rg.Formula = vals(i) Else rg.Value2 = vals(i)
        Else
            rg.Value2 = vals(i)
        End If
        k = Mid(kinds, i - LBound(spans) + 1, 1)
        rg.Font.name = FONT_NM: rg.Font.Size = 9: rg.VerticalAlignment = xlCenter
        If k = "h" Then
            StyleHeader rg
        Else
            Select Case k
                Case "c": rg.HorizontalAlignment = xlCenter
                Case "t": rg.HorizontalAlignment = xlLeft: rg.IndentLevel = 1
                Case "n": rg.NumberFormat = "#,##0;[Red]-#,##0;""-""": rg.HorizontalAlignment = xlRight: rg.ShrinkToFit = True
                Case "p": rg.NumberFormat = "0.0%": rg.HorizontalAlignment = xlCenter: rg.Font.Bold = True: PctFill rg
                Case "b": rg.NumberFormat = ";;;": rg.HorizontalAlignment = xlCenter
                Case "g": rg.NumberFormat = "0.0%": rg.HorizontalAlignment = xlRight: rg.Font.Bold = True: rg.IndentLevel = 1
            End Select
            rg.Borders(xlEdgeBottom).LineStyle = xlContinuous: rg.Borders(xlEdgeBottom).color = ClrHead()
        End If
        c = c + spans(i)
    Next i
    ws.rows(r).RowHeight = 19
End Sub

'퍼센트 셀 색조 (셀 1개 기준 3색 스케일 0 / 0.5 / 1)
Private Sub PctFill(rg As Range)
    Dim cs As ColorScale
    rg.FormatConditions.Delete
    Set cs = rg.FormatConditions.AddColorScale(ColorScaleType:=3)
    cs.ColorScaleCriteria(1).Type = xlConditionValueNumber: cs.ColorScaleCriteria(1).value = 0: cs.ColorScaleCriteria(1).FormatColor.color = RGB(255, 180, 176)
    cs.ColorScaleCriteria(2).Type = xlConditionValueNumber: cs.ColorScaleCriteria(2).value = 0.5: cs.ColorScaleCriteria(2).FormatColor.color = RGB(255, 236, 153)
    cs.ColorScaleCriteria(3).Type = xlConditionValueNumber: cs.ColorScaleCriteria(3).value = 1: cs.ColorScaleCriteria(3).FormatColor.color = RGB(178, 236, 194)
End Sub

'진행 막대: 집행률 구간별 그라데이션 색 (80% 이상 파랑 / 60~80% 초록 / 60% 미만 빨강)
Public Sub BarByRate(rg As Range, ByVal rate As Double)
    Dim db As Databar, clr As Long
    If rate >= 0.8 Then
        clr = ClrBlue()
    ElseIf rate >= 0.6 Then
        clr = ClrGreen()
    Else
        clr = ClrRed()
    End If
    rg.FormatConditions.Delete
    Set db = rg.Cells(1, 1).FormatConditions.AddDatabar
    db.MinPoint.Modify newtype:=xlConditionValueNumber, newValue:=0
    db.MaxPoint.Modify newtype:=xlConditionValueNumber, newValue:=1
    db.BarColor.color = clr
    db.BarFillType = xlDataBarFillGradient
    db.BarBorder.Type = xlDataBarBorderSolid
    db.BarBorder.color.color = clr
    db.ShowValue = (rg.NumberFormat <> ";;;")      '한 셀에 %값 + 막대
End Sub

'차트 공통 서식: 흰 배경, 테두리 없음, 인쇄 가능
Private Sub ChartFrame(co As ChartObject)
    With co.Chart
        .ChartArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .ChartArea.Format.line.Visible = msoFalse
        .PlotArea.Format.Fill.Visible = msoFalse
        .ChartArea.Font.name = FONT_NM
        .ChartArea.RoundedCorners = False
        .PlotVisibleOnly = False          '필터로 숨겨진 집계 행도 표시
    End With
    co.Placement = xlFreeFloating
    co.PrintObject = True
End Sub

'팀별 색 (Apple 팔레트 순환)
Private Function TeamColor(ByVal i As Long) As Long
    Select Case (i - 1) Mod 6
        Case 0: TeamColor = ClrBlue()
        Case 1: TeamColor = ClrGreen()
        Case 2: TeamColor = ClrOrange()
        Case 3: TeamColor = ClrIndigo()
        Case 4: TeamColor = ClrRed()
        Case Else: TeamColor = ClrGray()
    End Select
End Function

