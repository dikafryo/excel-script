Attribute VB_Name = "modDetail"
Option Explicit
'=====================================================================
' modDetail : 세부현황_세목별 (사업항목 1레벨 총액 → 세출세목별) + 오른쪽 슬라이서
'=====================================================================
Public Const SH_GROUP As String = "세부현황"
Public Const TBL_GROUP As String = "tbl세목별"
Private Const SC_GRP As String = "슬라이서_세목_"
Private Const G_HDR As Long = 9
Private Const nC As Long = 15

Public Sub 세목별현황()
    SpeedOn "세부현황(세목별) 생성 중..."
    On Error GoTo EH
    If Not ModelReady Then BuildModel
    BuildGrouped
    If SheetExists(SH_GROUP) Then BW().Worksheets(SH_GROUP).Activate
    SpeedOff
    Exit Sub
EH:
    SpeedOff
    MsgBox "세부현황(세목별) 오류 [" & gStep & "]: " & Err.Description, vbCritical
End Sub

Public Sub 세목슬라이서리셋()
    DeleteOrClearSlicers SC_GRP, False
End Sub

Public Sub BuildGrouped()
    Dim src As Worksheet, ws As Worksheet, hr As Long, hm As Object, arr As Variant, lr As Long, lc As Long
    Dim cSeb As Long, cN1 As Long, cCode As Long, cSem As Long, cHy As Long, cJi As Long, cGyo As Long, cGeup As Long, cWon As Long
    Dim useB As Boolean, r As Long, k As String, sk As String, agg As Object, sub1 As Object
    Dim v As Variant, it As clsProject, keys As Collection, i As Long, row As Long, n As Long, T As Variant
    Dim lo As ListObject, rng As Range, sx As Single, sy As Single, dept As String, out() As Variant, cnt As Long, x As Long

    If Not ModelReady Then Exit Sub
    gStep = "세목별 집계"
    Set agg = ItemSubtotals(cnt)
    If agg Is Nothing Then Exit Sub

    '출력 배열 (총액행 + 세목행)
    gStep = "세목별 표"
    ReDim out(1 To cnt + gItems.Count + 1, 1 To nC)
    Set keys = SortedKeys()
    For i = 1 To keys.Count
        Set it = gItems(keys(i))
        If it.장표키 <> "" Then
            If agg.Exists(it.장표키) Then
                Set sub1 = agg(it.장표키)
                x = x + 1
                out(x, 1) = x: out(x, 2) = "사업총액"
                out(x, 3) = it.세부사업: out(x, 4) = it.사업항목
                out(x, 5) = "(" & sub1.Count & "개 세목)"
                out(x, 6) = it.예산현액: out(x, 7) = it.원인행위: out(x, 8) = it.집행액: out(x, 9) = it.잔액
                out(x, 10) = it.집행률: out(x, 11) = it.집행률: out(x, 12) = it.상태
                out(x, 13) = IIf(it.팀 = "", "(팀 미지정)", it.팀): out(x, 14) = IIf(it.담당자 = "", "(담당자 미지정)", it.담당자)
                out(x, 15) = IIf(it.과제카드 = "", "-", it.과제카드)
                For Each v In sub1.keys
                    T = sub1(v)
                    x = x + 1
                    out(x, 1) = x: out(x, 2) = "세목"
                    out(x, 3) = it.세부사업: out(x, 4) = it.사업항목
                    out(x, 5) = "   " & T(0)
                    out(x, 6) = T(1): out(x, 7) = T(3): out(x, 8) = T(2): out(x, 9) = T(1) - T(2)
                    out(x, 10) = rate(T(2), T(1)): out(x, 11) = rate(T(2), T(1)): out(x, 12) = StatusText(T(1), rate(T(2), T(1)), "")
                    out(x, 13) = out(x - 1, 13): out(x, 14) = out(x - 1, 14): out(x, 15) = out(x - 1, 15)
                Next v
            End If
        End If
    Next i
    n = x

    '시트
    gStep = "세목별 시트"
    DeleteOrClearSlicers SC_GRP, False          '슬라이서 유지(필터만 해제)
    Set ws = GetSheet(SH_GROUP, True)
    ws.Tab.color = ClrIndigo()
    ws.Cells.Interior.color = RGB(255, 255, 255)
    ws.Columns.hidden = False
    ws.rows(1).RowHeight = 4: ws.rows("2:3").RowHeight = 14: ws.rows(4).RowHeight = 6
    AddHomeButton ws, 4, 4
    AddButton ws, "btnPrint", "인쇄", "인쇄", 68, 4, 50, 20, ClrNavy()
    AddButton ws, "btnRefresh", "새로고침(리셋)", "새로고침버튼", 122, 4, 90, 20, ClrGreen()

    dept = CfgGet("부서명")
    ws.Range("A5").Value2 = IIf(dept <> "", dept & " ", "") & "세부현황"
    ws.Range("A5").Font.Size = 20: ws.Range("A5").Font.Bold = True: ws.Range("A5").Font.name = FONT_NM: ws.Range("A5").Font.color = ClrNavy()
    ws.rows(5).RowHeight = 34
    ws.Range("A6").Value2 = "기준일 " & CfgGet("기준일") & "   |   갱신 " & NowText() & "   |   (단위: 원)   |   파란 행 = 사업항목 총액"
    ws.Range("A6").Font.Size = 9: ws.Range("A6").Font.color = ClrGray()
    ws.rows(6).RowHeight = 14

    If n = 0 Then Exit Sub
    ws.Cells(G_HDR, 1).Resize(1, nC).Value2 = Array("No", "구분", "세부사업", "사업항목", "세출세목", _
        "예산현액", "원인행위액", "집행액", "집행잔액", "집행률", "진행", "상태", "팀/부서", "담당자", "단위과제카드")
    ws.Cells(G_HDR + 1, 1).Resize(n, nC).Value2 = out

    Set rng = ws.Range(ws.Cells(G_HDR, 1), ws.Cells(G_HDR + n, nC))
    Set lo = MakeTable(ws, rng, TBL_GROUP)
    StyleHeader lo.HeaderRowRange
    StyleBody lo.DataBodyRange
    With ws.Range(ws.Cells(G_HDR + 1, 6), ws.Cells(G_HDR + n, 9))
        .NumberFormat = "#,##0;[Red]-#,##0;""-""": .HorizontalAlignment = xlRight: .ShrinkToFit = True
    End With
    '집행률·진행 한 셀 (% + 그라데이션 막대), 진행 열 숨김
    With ws.Range(ws.Cells(G_HDR + 1, 10), ws.Cells(G_HDR + n, 10))
        .NumberFormat = "0.0%": .HorizontalAlignment = xlRight: .Font.Bold = True: .IndentLevel = 1
    End With
    Dim ii As Long
    For ii = 1 To n
        BarByRate ws.Cells(G_HDR + ii, 10), NzD(out(ii, 10))
    Next ii
    ws.Cells(G_HDR, 10).Value2 = "집행률·진행"
    ColorStatus ws.Range(ws.Cells(G_HDR + 1, 12), ws.Cells(G_HDR + n, 12))
    ws.Range(ws.Cells(G_HDR + 1, 12), ws.Cells(G_HDR + n, 12)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(G_HDR + 1, 1), ws.Cells(G_HDR + n, 2)).HorizontalAlignment = xlCenter
    '총액 행 강조 (조건부 서식)
    Dim fC As FormatCondition
    Set fC = lo.DataBodyRange.FormatConditions.Add(Type:=xlExpression, formula1:="=$B" & (G_HDR + 1) & "=""사업총액""")
    fC.Interior.color = RGB(229, 240, 255)
    fC.Font.Bold = True

    '요약 행: 총액 행만 합산 (필터 반영)
    lo.ShowTotals = True
    lo.ListColumns(4).total.Value2 = "합계 (사업총액 행 기준)"
    lo.ListColumns(1).TotalsCalculation = xlTotalsCalculationNone
    lo.ListColumns(6).total.Formula = VisibleSumFormula(TBL_GROUP, "예산현액")
    lo.ListColumns(7).total.Formula = VisibleSumFormula(TBL_GROUP, "원인행위액")
    lo.ListColumns(8).total.Formula = VisibleSumFormula(TBL_GROUP, "집행액")
    lo.ListColumns(9).total.Formula = VisibleSumFormula(TBL_GROUP, "집행잔액")
    lo.ListColumns(10).total.Formula = "=IFERROR(" & Mid(VisibleSumFormula(TBL_GROUP, "집행액"), 2) & "/" & Mid(VisibleSumFormula(TBL_GROUP, "예산현액"), 2) & ",0)"
    With lo.TotalsRowRange
        .Font.name = FONT_NM: .Font.Bold = True: .Font.color = ClrNavy(): .Interior.color = ClrLight(): .ShrinkToFit = True
        .Borders(xlEdgeTop).LineStyle = xlContinuous: .Borders(xlEdgeTop).weight = xlMedium: .Borders(xlEdgeTop).color = ClrGray()
    End With
    ws.Range(lo.TotalsRowRange.Cells(1, 6), lo.TotalsRowRange.Cells(1, 9)).NumberFormat = "#,##0"
    lo.TotalsRowRange.Cells(1, 10).NumberFormat = "0.0%"

    'KPI
    ws.Columns(1).ColumnWidth = 5: ws.Columns(2).ColumnWidth = 8: ws.Columns(3).ColumnWidth = 16: ws.Columns(4).ColumnWidth = 32
    ws.Columns(5).ColumnWidth = 30
    ws.Range(ws.Columns(6), ws.Columns(9)).ColumnWidth = 18: ws.Columns(10).ColumnWidth = 24: ws.Columns(11).hidden = True: ws.Columns(12).ColumnWidth = 11
    ws.Range(ws.Columns(13), ws.Columns(15)).hidden = True
    ws.Columns(nC + 1).ColumnWidth = 2

    KpiRow ws, 7, Array("예산현액", "집행액", "집행잔액", "집행률", "사업 수"), _
        Array(VisibleSumFormula(TBL_GROUP, "예산현액"), VisibleSumFormula(TBL_GROUP, "집행액"), VisibleSumFormula(TBL_GROUP, "집행잔액"), _
              "=IFERROR(" & Mid(VisibleSumFormula(TBL_GROUP, "집행액"), 2) & "/" & Mid(VisibleSumFormula(TBL_GROUP, "예산현액"), 2) & ",0)", _
              "=SUMPRODUCT((" & TBL_GROUP & "[구분]=""사업총액"")*SUBTOTAL(103,OFFSET(" & TBL_GROUP & "[구분],ROW(" & TBL_GROUP & "[구분])-MIN(ROW(" & TBL_GROUP & "[구분])),0,1)))"), _
        Array("#,##0", "#,##0", "#,##0", "0.0%", "0"), Array(ClrNavy(), ClrBlue(), ClrOrange(), ClrGreen(), ClrIndigo())

    '슬라이서: 제목 오른쪽 상단(1~6행) 가로 배치 - 표 옆 공간 절약
    gStep = "세목별 슬라이서"
    '제목(A5, 20pt) 오른쪽부터 시작, 폭 축소 → 세부사업 슬라이서까지 화면 안에
    sx = ws.Cells(5, 4).Left: sy = 3
    If sx < 380 Then sx = 380
    '높이 76 통일 → 7~8행 KPI 카드와 겹치지 않음, 팀/부서 폭 확대
    AddSlicerTo lo, ws, "팀/부서", SC_GRP & "팀", sx, sy, 210, 76, 3: sx = sx + 214
    AddSlicerTo lo, ws, "담당자", SC_GRP & "담당자", sx, sy, 200, 76, 3: sx = sx + 204
    AddSlicerTo lo, ws, "단위과제카드", SC_GRP & "단위과제카드", sx, sy, 230, 76, 2: sx = sx + 234
    AddSlicerTo lo, ws, "세부사업", SC_GRP & "세부사업", sx, sy, 230, 76, 2

    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Range("A1").Select
    ws.Cells(G_HDR + 1, 6).Select
    ActiveWindow.FreezePanes = True
    ActiveWindow.DisplayGridlines = False
    ws.Range("A1").Select
End Sub

'장표(102-B)에서 사업항목키 -> (세목 -> Array(세목표시, 예산현액, 집행액, 원인행위)) 집계 (공용)
Public Function ItemSubtotals(ByRef cnt As Long) As Object
    Dim src As Worksheet, hr As Long, hm As Object, arr As Variant, lr As Long, lc As Long, useB As Boolean
    Dim cSeb As Long, cN1 As Long, cCode As Long, cSem As Long, cHy As Long, cJi As Long, cGyo As Long, cGeup As Long, cWon As Long
    Dim r As Long, k As String, sk As String, agg As Object, sub1 As Object, T As Variant
    Set ItemSubtotals = Nothing: cnt = 0
    If SheetExists(SH_B) Then
        Set src = BW().Worksheets(SH_B): useB = True
    ElseIf SheetExists(SH_A) Then
        Set src = BW().Worksheets(SH_A): useB = False
    Else
        Exit Function
    End If
    hr = FindHeaderRow(src, "회계연도", 10)
    If hr = 0 Then Exit Function
    Set hm = HeaderMap(src, hr)
    cSeb = HCol(hm, "세부사업"): cCode = HCol(hm, "세출세목코드"): cSem = HCol(hm, "세출세목")
    cN1 = HCol(hm, IIf(useB, "사업항목명1", "사업항목1레벨"))
    cHy = HCol(hm, "예산현액"): cJi = HCol(hm, "지출금액"): cWon = HCol(hm, "원인행위금액")
    cGyo = HCol(hm, "일상경비교부액", False): cGeup = HCol(hm, "일상경비지급액", False)
    lr = lastRow(src, 1): lc = src.Cells(hr, src.Columns.Count).End(xlToLeft).Column
    If lr <= hr Then Exit Function
    arr = src.Range(src.Cells(hr + 1, 1), src.Cells(lr, lc)).Value2
    Set agg = CreateObject("Scripting.Dictionary")
    For r = 1 To UBound(arr, 1)
        If CleanText(arr(r, cCode)) <> "" Then
            k = MakeKey(CleanText(arr(r, cSeb)), CleanText(arr(r, cN1)))
            sk = CleanText(arr(r, cCode)) & " " & CleanText(arr(r, cSem))
            If Not agg.Exists(k) Then agg.Add k, CreateObject("Scripting.Dictionary")
            Set sub1 = agg(k)
            If sub1.Exists(sk) Then T = sub1(sk) Else T = Array(sk, 0#, 0#, 0#)
            T(1) = T(1) + NzD(arr(r, cHy))
            T(2) = T(2) + ExecAmount(arr, r, cJi, cGyo, cGeup)
            T(3) = T(3) + NzD(arr(r, cWon))
            sub1(sk) = T
            cnt = cnt + 1
        End If
    Next r
    Set ItemSubtotals = agg
End Function

'필터(슬라이서) 반영 + 구분="사업총액" 행만 합산하는 수식
Private Function VisibleSumFormula(ByVal tbl As String, ByVal col As String) As String
    VisibleSumFormula = "=SUMPRODUCT((" & tbl & "[구분]=""사업총액"")*SUBTOTAL(109,OFFSET(" & tbl & "[" & col & "],ROW(" & tbl & "[" & col & "])-MIN(ROW(" & tbl & "[" & col & "])),0,1)))"
End Function
