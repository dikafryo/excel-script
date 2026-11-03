Attribute VB_Name = "modCalc"
Option Explicit
'=====================================================================
' modCalc : 세부산출내역 (세출예산요구서의 산출기초 행 단위)  - V8
'   열: No · 세부사업 · 사업항목 · 사업명 · {올해}년 산출내역 · {작년}년 산출내역 · 세목코드 · {올해}년 예산 · {작년}년 예산
'       · 장표 예산 · 차이(세출-장표) · 차이(산출합-장표) · 차수 · 메모(입력) · 비고(확인 필요 사유)
'       + 숨김: 원가통계목 · 레벨 · 팀/부서 · 담당자 · 단위과제카드 · 키 · 출처파일
'   V8: 머리글 연도는 세출 파일 머리글에서 자동 인식(modYearHdr). 열은 머리글 이름으로 찾음(위치 고정 없음)
'   V8: 추경 파일이 있으면 사업항목(세부사업|사업항목) 단위로 나중 차수 내용으로 교체, 새 사업은 추가 (차수 열·색으로 구분)
'   V8: 사업항목(1레벨) 행마다 비고 = 장표 대조 / 최신년도 산출내역 미반영 가능성 / 예산현황-산출내역 확인 → 행 색 + 맨 위 '확인 필요 n건'
'   V8: 추경에서 바뀐(또는 추가된) 사업 블록은 메모 끝에 '1추' · '2추' 자동 표시 (사용자 메모는 보존, 자동 표시는 새로고침마다 다시 계산)
'   슬라이서(오른쪽 세로): 차수 · 팀/부서 · 담당자 · 단위과제카드 · 세부사업 · 사업명 · 레벨 · 세목코드 · 원가통계목
'   찾기: E7(산출내역 키워드) / H7(메모) 입력 후 [찾기] → 해당 문구가 있는 행만 표시(둘 다 입력 시 AND), [취소] → 해제
'   메모는 새로고침해도 보존 (키 = 세부사업|사업명|산출내역|세목코드|순서)
'=====================================================================
Public Const SH_CALC As String = "세부산출내역"
Public Const TBL_CALC As String = "tbl세부산출"
Private Const SC_CAL As String = "슬라이서_산출_"
Private Const C_HDR As Long = 9
Private Const nC As Long = 22
Private Const K_NO As Long = 1, K_SEB As Long = 2, K_NM As Long = 3, K_SNM As Long = 4, K_CALC As Long = 5, K_CALCP As Long = 6
Private Const K_CODE As Long = 7, K_BUD As Long = 8, K_PRV As Long = 9, K_JANG As Long = 10, K_DIF As Long = 11, K_DIF2 As Long = 12
Private Const K_CHA As Long = 13, K_MEMO As Long = 14, K_BAD As Long = 15
Private Const K_STAT As Long = 16, K_LV As Long = 17, K_TEAM As Long = 18, K_OWN As Long = 19, K_CARD As Long = 20, K_KEY As Long = 21
Private Const K_SRC As Long = 22
Private Const MEMO_CELL As String = "H7"      '메모 찾기 입력칸
Private Const CALC_CELL As String = "E7"      '산출내역 키워드 찾기 입력칸
Private Const LAST_HDR As String = "출처파일" '표 구조 판별용 (이 머리글이 마지막 열이 아니면 V7 표 → 새로 만듦)

Public Sub 세부산출내역()
    SpeedOn "세부산출내역 생성 중..."
    On Error GoTo EH
    If Not ModelReady Then BuildModel
    BuildCalc
    If SheetExists(SH_CALC) Then BW().Worksheets(SH_CALC).Activate
    SpeedOff
    Exit Sub
EH:
    SpeedOff
    MsgBox "세부산출내역 오류 [" & gStep & "]: " & Err.Description & " (" & Err.Number & ")", vbCritical
End Sub

'찾기: 산출내역 키워드(E7) / 메모(H7) - 입력된 칸 기준으로 표 필터 (둘 다 입력 시 모두 만족하는 행)
Public Sub 메모찾기()
    Dim ws As Worksheet, lo As ListObject, tc As String, tM As String
    If Not SheetExists(SH_CALC) Then Exit Sub
    Set ws = BW().Worksheets(SH_CALC)
    If ws.ListObjects.Count = 0 Then Exit Sub
    Set lo = ws.ListObjects(TBL_CALC)
    tc = CleanText(ws.Range(CALC_CELL).Value2)
    tM = CleanText(ws.Range(MEMO_CELL).Value2)
    If tc = "" And tM = "" Then
        tc = InputBox("산출내역에서 찾을 키워드를 입력하세요. (메모 찾기는 " & MEMO_CELL & " 칸에 입력)", "산출내역 찾기")
        If tc = "" Then Exit Sub
        ws.Range(CALC_CELL).Value2 = tc
    End If
    On Error Resume Next
    lo.AutoFilter.ShowAllData
    On Error GoTo 0
    If tc <> "" Then lo.Range.AutoFilter Field:=K_CALC, Criteria1:="=*" & tc & "*"
    If tM <> "" Then lo.Range.AutoFilter Field:=K_MEMO, Criteria1:="=*" & tM & "*"
End Sub

Public Sub 메모찾기취소()
    Dim ws As Worksheet, lo As ListObject
    If Not SheetExists(SH_CALC) Then Exit Sub
    Set ws = BW().Worksheets(SH_CALC)
    If ws.ListObjects.Count = 0 Then Exit Sub
    Set lo = ws.ListObjects(TBL_CALC)
    On Error Resume Next
    lo.AutoFilter.ShowAllData
    On Error GoTo 0
    ws.Range(MEMO_CELL).Value2 = ""
    ws.Range(CALC_CELL).Value2 = ""
End Sub

'---------------------------------------------------------------- 세출 시트 한 장 읽기 → 사업항목 블록별 행
'  blocks: 블록키(세부사업|사업항목, 첫 사업항목 전 행은 "#pre") → Collection(행 배열)
'  행 배열(0~14): 세부사업, 사업항목명, 사업명, 올해 산출내역, 작년 산출내역, 세목코드, 올해 예산(천원), 작년 예산(천원),
'                 원가통계목, 레벨, 블록키, 세세부사업 원문, 산출내역(정리), 올해 금액(원), 차수 번호
Private Sub ReadCalcLines(ws As Worksheet, ByVal idx As Long, blocks As Object, order As Collection, _
                          ByRef yCur As Long, ByRef yPrv As Long)
    Dim mp As Object, hr As Long, arr As Variant, lr As Long, lc As Long, r As Long, cfgY As String
    Dim cLv As Long, cE As Long, cI As Long, cM As Long, cM2 As Long, cP As Long, cQ As Long, cCode As Long, cStat As Long
    Dim lv As String, e As String, m As String, M2 As String, nm As String, p As Long, bk As String, unit As Double
    Dim ln As Variant, col As Collection
    If idx = 0 Then cfgY = "올해연도" Else cfgY = "추경" & idx & "연도"
    Set mp = YH_Map(ws, SuppFile(idx), cfgY, Not gBX_NoUI)
    hr = mp("hr"): unit = mp("unit"): lc = mp("lastC")
    cLv = mp("lv"): cE = mp("e"): cI = mp("i"): cM = mp("calc"): cM2 = mp("calcP")
    cP = mp("amt"): cQ = mp("amtP"): cCode = mp("code"): cStat = mp("stat")   'V8: 금액은 YH_AmtAt (추경 양식: 기정 + 추경요구액)
    If idx = 0 Then
        yCur = mp("yCur"): yPrv = mp("yPrv")
    ElseIf mp("yPrv") <> yPrv Then
        cM2 = 0: cQ = 0             '추경 파일의 작년 열은 본예산 작년 연도와 같을 때만 사용 (기정액을 작년으로 쓰지 않음)
    End If
    lr = Application.WorksheetFunction.Max(lastRow(ws, cE), lastRow(ws, cM), lastRow(ws, cP))
    If mp("gd") Then lr = Application.WorksheetFunction.Max(lr, lastRow(ws, mp("delta")))
    If lr <= hr Then Exit Sub
    arr = ws.Range(ws.Cells(hr + 1, 1), ws.Cells(lr, lc)).Value2
    bk = "#pre": nm = ""
    For r = 1 To UBound(arr, 1)
        lv = CleanText(arr(r, cLv)): e = CleanText(arr(r, cE))
        m = RTrim(Replace(Replace(CStr(IIf(IsError(arr(r, cM)), "", arr(r, cM))), vbLf, ""), vbCr, ""))
        If e = "" Or m = "" Then GoTo NextRow
        If CleanText(m) = "소계" Or lv = "합계" Then GoTo NextRow
        If lv = "" And IsLevel1(e) Then
            p = InStr(e, "."): nm = Trim(Mid(e, p + 1))
            bk = MakeKey(CleanText(arr(r, cI)), nm)
            lv = "1"
        Else
            If lv <> "" Then lv = CStr(CLng(val(lv)))
        End If
        M2 = ""
        If cM2 > 0 Then M2 = RTrim(Replace(Replace(CStr(IIf(IsError(arr(r, cM2)), "", arr(r, cM2))), vbLf, ""), vbCr, ""))
        ReDim ln(0 To 14)
        ln(0) = CleanText(arr(r, cI)): ln(1) = nm: ln(2) = StripNo(e)
        If lv = "1" Then ln(3) = StripNo(CleanText(m)) Else ln(3) = m
        If lv = "1" Then ln(4) = StripNo(CleanText(M2)) Else ln(4) = M2
        If cCode > 0 Then ln(5) = CleanText(arr(r, cCode)) Else ln(5) = ""
        ln(6) = YH_AmtAt(arr, r, mp) * unit / 1000
        If cQ > 0 Then ln(7) = NzD(arr(r, cQ)) * unit / 1000 Else ln(7) = Empty
        If cStat > 0 Then ln(8) = CleanText(arr(r, cStat)) Else ln(8) = ""
        ln(9) = lv: ln(10) = bk: ln(11) = e: ln(12) = CleanText(m)
        ln(13) = YH_AmtAt(arr, r, mp) * unit
        ln(14) = idx
        If Not blocks.Exists(bk) Then
            Set col = New Collection
            blocks.Add bk, col
            order.Add bk
        End If
        Set col = blocks(bk)
        col.Add ln
NextRow:
    Next r
End Sub

'기존 표의 메모 읽기 (열 위치가 아니라 머리글 '키'·'메모'로 찾음 → V7 표에서도 보존)
Private Sub ReadMemos(lo As ListObject, memo As Object)
    Dim cK As Long, cMm As Long, i As Long, r As Long, rk As String
    For i = 1 To lo.ListColumns.Count
        Select Case CleanText(lo.ListColumns(i).name)
            Case "키": cK = i
            Case "메모": cMm = i
        End Select
    Next i
    If cK = 0 Or cMm = 0 Then Exit Sub
    If lo.ListRows.Count = 0 Then Exit Sub
    For r = 1 To lo.ListRows.Count
        rk = CleanText(lo.DataBodyRange.Cells(r, cK).Value2)
        If rk <> "" And CleanText(lo.DataBodyRange.Cells(r, cMm).Value2) <> "" Then
            If StripTag(CStr(lo.DataBodyRange.Cells(r, cMm).Value2)) <> "" Then memo(rk) = StripTag(CStr(lo.DataBodyRange.Cells(r, cMm).Value2))
        End If
    Next r
End Sub

Private Function ColL(ws As Worksheet, ByVal c As Long) As String
    ColL = Split(ws.Cells(1, c).Address(True, False), "$")(0)
End Function

'---------------------------------------------------------------- 시트 생성
Public Sub BuildCalc()
    Dim ws As Worksheet, lo As ListObject, rng As Range, fC As FormatCondition, memo As Object, cnt As Object
    Dim blocksArr() As Object, orders() As Collection, nS As Long, s As Long, i As Long
    Dim yCur As Long, yPrv As Long, keyList() As String, ownList() As Long, nk As Long, kk As Variant, pos As Long, sebOf As String
    Dim out() As Variant, x As Long, nLines As Long, ln As Variant, it As clsProject, rq As clsProject, rk As String, dept As String
    Dim bd As Object, col As Collection, jang As Double, found As Boolean, sumWon As Double, budKey As Double, hasCode As Boolean
    Dim firstLv1 As Boolean, nBad As Long, anyPrvCalc As Boolean, anyPrv As Boolean, jl As Boolean
    Dim hCalc As String, hCalcP As String, hBud As String, hPrv As String, sx As Single, sy As Single
    Dim colLv As String, colBad As String, colCha As String, prvMap As Object, occ As Object, pk As String
    Dim tagList() As String, cat As Long, note As String, nA As Long, nb As Long, nC2 As Long, sumTxt As String, nJO As Long

    If Not SheetExists(SH_REQ) Then
        MsgBox "세부산출내역은 ① 세출예산요구서가 있어야 만들 수 있습니다.", vbExclamation
        Exit Sub
    End If
    gStep = "산출내역 읽기"
    nS = SuppCount()
    ReDim blocksArr(0 To nS): ReDim orders(0 To nS)
    For s = 0 To nS
        Set blocksArr(s) = CreateObject("Scripting.Dictionary")
        Set orders(s) = New Collection
        If s = 0 Then
            ReadCalcLines BW().Worksheets(SH_REQ), 0, blocksArr(0), orders(0), yCur, yPrv
        ElseIf SheetExists(SuppSheetName(s)) Then
            ReadCalcLines BW().Worksheets(SuppSheetName(s)), s, blocksArr(s), orders(s), yCur, yPrv
        End If
    Next s

    '본예산 블록 순서에서 시작 → 추경: 같은 키는 그 자리에서 최신 차수로 교체, 새 키는 같은 세부사업 마지막 블록 뒤(없으면 맨 끝)에 추가
    gStep = "추경 반영"
    nk = 0
    ReDim keyList(1 To orders(0).Count + 1): ReDim ownList(1 To orders(0).Count + 1): ReDim tagList(1 To orders(0).Count + 1)
    For Each kk In orders(0)
        nk = nk + 1: keyList(nk) = kk: ownList(nk) = 0
    Next kk
    For s = 1 To nS
        For Each kk In orders(s)
            If kk <> "#pre" Then
                pos = 0
                For i = 1 To nk
                    If keyList(i) = kk Then pos = i: Exit For
                Next i
                If pos > 0 Then
                    '금액(사업명 행·세목 행)이 앞 차수와 같으면 출처를 바꾸지 않는다 (작년 열·차수 표시 유지)
                    If BlockSig(blocksArr(s).item(kk)) <> BlockSig(blocksArr(ownList(pos)).item(kk)) Then
                        ownList(pos) = s
                        tagList(pos) = AddTag(tagList(pos), s)          'V8: 메모 '1추' · '2추'
                    End If
                Else
                    sebOf = Split(kk, "|")(0)
                    pos = nk
                    For i = nk To 1 Step -1
                        If Split(keyList(i), "|")(0) = sebOf Then pos = i: Exit For
                    Next i
                    nk = nk + 1
                    ReDim Preserve keyList(1 To nk): ReDim Preserve ownList(1 To nk): ReDim Preserve tagList(1 To nk)
                    For i = nk To pos + 2 Step -1
                        keyList(i) = keyList(i - 1): ownList(i) = ownList(i - 1): tagList(i) = tagList(i - 1)
                    Next i
                    keyList(pos + 1) = kk: ownList(pos + 1) = s: tagList(pos + 1) = AddTag("", s)
                End If
            End If
        Next kk
    Next s
    nLines = 0
    For i = 1 To nk
        Set bd = blocksArr(ownList(i))
        Set col = bd.item(keyList(i))
        nLines = nLines + col.Count
    Next i
    If nLines = 0 Then Exit Sub

    '기존 메모 보존 + 표 구조가 다르면(V7 표) 슬라이서·시트를 지우고 새로 만든다
    Set memo = CreateObject("Scripting.Dictionary")
    If SheetExists(SH_CALC) Then
        Set ws = BW().Worksheets(SH_CALC)
        If ws.ListObjects.Count > 0 Then
            Set lo = ws.ListObjects(1)
            ReadMemos lo, memo
            If lo.ListColumns.Count <> nC Or CleanText(lo.HeaderRowRange.Cells(1, lo.ListColumns.Count).Value2) <> LAST_HDR _
               Or CleanText(lo.HeaderRowRange.Cells(1, K_BAD).Value2) <> "비고" Then
                DeleteOrClearSlicers SC_CAL, True
                DeleteSheetIfExists SH_CALC
            End If
        End If
    End If

    '행 구성
    gStep = "산출내역 행"
    jl = JangLoaded()
    ReDim out(1 To nLines, 1 To nC)
    Set cnt = CreateObject("Scripting.Dictionary")
    x = 0: nBad = 0: nA = 0: nb = 0: nC2 = 0
    For i = 1 To nk
        Set bd = blocksArr(ownList(i))
        Set col = bd.item(keyList(i))
        Set it = Nothing
        If gItems.Exists(keyList(i)) Then
            Set it = gItems(keyList(i))
        ElseIf Not gReqItems Is Nothing Then
            If gReqItems.Exists(keyList(i)) Then
                Set rq = gReqItems(keyList(i))
                If rq.장표키 <> "" Then
                    If gItems.Exists(rq.장표키) Then Set it = gItems(rq.장표키)
                End If
            End If
        End If
        '블록 합계: 사업항목(1레벨) 금액 합 / 세목코드 행 금액 합 (원)
        sumWon = 0: budKey = 0: hasCode = False
        For Each ln In col
            If ln(9) = "1" Then
                budKey = budKey + ln(13)
            ElseIf ln(5) <> "" Then
                sumWon = sumWon + ln(13): hasCode = True
            End If
        Next ln
        '추경 블록: 작년 열이 비어 있으면 본예산 같은 블록에서 (레벨|세목코드|순서) 기준으로 가져온다
        Set prvMap = Nothing
        If ownList(i) > 0 Then
            If blocksArr(0).Exists(keyList(i)) Then Set prvMap = PrevLineMap(blocksArr(0).item(keyList(i)))
        End If
        Set occ = CreateObject("Scripting.Dictionary")
        firstLv1 = True
        For Each ln In col
            If Not prvMap Is Nothing Then
                pk = ln(9) & "|" & ln(5)
                occ(pk) = occ(pk) + 1
                pk = pk & "|" & occ(pk)
                If prvMap.Exists(pk) Then
                    If ln(4) = "" Then ln(4) = prvMap(pk)(0)
                    If IsEmpty(ln(7)) Then ln(7) = prvMap(pk)(1)
                End If
            End If
            x = x + 1
            out(x, K_NO) = x
            out(x, K_SEB) = ln(0)
            out(x, K_SNM) = ln(2)
            If it Is Nothing Then out(x, K_NM) = ln(1) Else out(x, K_NM) = it.사업항목
            out(x, K_CALC) = ln(3)
            out(x, K_CALCP) = ln(4)
            If ln(4) <> "" Then anyPrvCalc = True
            out(x, K_CODE) = ln(5)
            out(x, K_BUD) = ln(6)
            out(x, K_PRV) = ln(7)
            If Not IsEmpty(ln(7)) Then anyPrv = True
            out(x, K_STAT) = ln(8)
            out(x, K_LV) = ln(9)
            out(x, K_CHA) = SuppLabel(CLng(ln(14)))
            out(x, K_SRC) = FileNameOf(SuppFile(CLng(ln(14))))
            If it Is Nothing Then
                out(x, K_TEAM) = "(팀 미지정)": out(x, K_OWN) = "(담당자 미지정)": out(x, K_CARD) = "-"
            Else
                out(x, K_TEAM) = IIf(it.팀 = "", "(팀 미지정)", it.팀)
                out(x, K_OWN) = IIf(it.담당자 = "", "(담당자 미지정)", it.담당자)
                out(x, K_CARD) = IIf(it.과제카드 = "", "-", it.과제카드)
            End If
            rk = ln(0) & "|" & ln(11) & "|" & ln(12) & "|" & ln(5)
            If cnt.Exists(rk) Then cnt(rk) = cnt(rk) + 1 Else cnt.Add rk, 1
            rk = rk & "|" & cnt(rk)
            out(x, K_KEY) = rk
            If memo.Exists(rk) Then out(x, K_MEMO) = memo(rk)
            If tagList(i) <> "" And keyList(i) <> "#pre" Then out(x, K_MEMO) = JoinMemo(CStr(out(x, K_MEMO)), tagList(i))
            '비고: 사업항목(1레벨) 첫 행에만 - 장표 대조(같은 사업항목이 여러 번 나오면 합산) + 산출내역 비교(최종 금액을 정한 차수 파일)
            If ln(9) = "1" And firstLv1 And keyList(i) <> "#pre" Then
                firstLv1 = False
                cat = 0: note = ""
                If jl Then
                    jang = JangOf(keyList(i), found)
                    If found Then out(x, K_JANG) = jang / 1000
                    out(x, K_DIF) = (budKey - jang) / 1000
                    If hasCode Then out(x, K_DIF2) = (sumWon - jang) / 1000
                    note = JangNote(budKey, IIf(hasCode, sumWon, -1), found, jang, SuppLabel(ownList(i)), cat)
                End If
                If Not gReqItems Is Nothing Then
                    If gReqItems.Exists(keyList(i)) Then note = JoinNote(note, CalcNote(gReqItems(keyList(i)), cat))
                End If
                out(x, K_BAD) = note
                If cat <> 0 Then
                    nBad = nBad + 1
                    If cat And 1 Then nA = nA + 1
                    If cat And 2 Then nb = nb + 1
                    If cat And 4 Then nC2 = nC2 + 1
                End If
            End If
        Next ln
    Next i

    '시트
    gStep = "산출내역 시트"
    DeleteOrClearSlicers SC_CAL, False
    Set ws = GetSheet(SH_CALC, True)
    ws.Tab.color = ClrOrange()
    ws.Cells.Interior.color = RGB(255, 255, 255)
    ws.Columns.hidden = False
    ws.rows(1).RowHeight = 4: ws.rows("2:3").RowHeight = 14: ws.rows(4).RowHeight = 6
    AddHomeButton ws, 4, 4
    AddButton ws, "btnPrint", "인쇄", "인쇄", 68, 4, 50, 20, ClrNavy()
    AddButton ws, "btnRefresh", "새로고침(리셋)", "새로고침버튼", 122, 4, 90, 20, ClrGreen()
    '맨 위 요약: 확인 필요 n건 (종류별 건수) - 최종 상태(본예산 + 추경 반영) 기준
    '  장표에만 있는 사업은 이 시트에 행이 없지만 추경점검·메인과 같은 숫자가 되도록 장표 대조 건수에 넣는다
    nJO = 0
    If jl Then nJO = JangOnlyCount()
    nBad = nBad + nJO: nA = nA + nJO
    sumTxt = NoteSummary(nBad, nA, nb, nC2, jl)
    With ws.Range("D2:F3")
        .Merge
        .Value2 = "확인 필요 " & nBad & "건"
        .Font.color = IIf(nBad > 0, ClrOrange(), ClrGreen())
        .Font.name = FONT_NM: .Font.Size = 13: .Font.Bold = True
        .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter
    End With
    With ws.Range("G2:L3")
        .Merge
        .Value2 = Mid$(sumTxt, InStr(sumTxt, "(") + 1, Len(sumTxt) - InStr(sumTxt, "(") - 1) & vbLf & "최종 상태 기준 · 추경 등으로 정상일 수 있음 · 비고 열에서 사유 확인" & _
                  IIf(nJO > 0, " · 장표에만 있는 사업 " & nJO & "건 포함(이 시트에 행 없음 → 추경점검)", "")
        .Font.name = FONT_NM: .Font.Size = 9: .Font.color = ClrGray(): .WrapText = True
        .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter
    End With
    dept = CfgGet("부서명")
    ws.Range("A5").Value2 = IIf(dept <> "", dept & " ", "") & "세부산출내역"
    ws.Range("A5").Font.Size = 20: ws.Range("A5").Font.Bold = True: ws.Range("A5").Font.name = FONT_NM: ws.Range("A5").Font.color = ClrNavy()
    ws.rows(5).RowHeight = 34
    ws.Range("A6").Value2 = "갱신 " & NowText() & "   |   (단위: 천원)   |   올해 " & IIf(yCur > 0, yCur & "년", "?") & " · 작년 " & IIf(yPrv > 0, yPrv & "년", "자료 없음") & _
        " (세출 파일 머리글에서 자동 인식: " & CfgGet("연도출처") & ")   |   장표 예산 = 102-B 예산금액(기준값), 차이 = 세출 - 장표   |   차수 = 본예산/추경 출처 (메모 '1추'·'2추' = 그 차수에서 바뀐 사업)   |   " & CalcRuleText() & "   |   파란 행 = 사업명(1레벨)   |   메모는 새로고침해도 보존"
    ws.Range("A6").Font.Size = 9: ws.Range("A6").Font.color = ClrGray()
    ws.rows(6).RowHeight = 14
    '찾기: D7 안내 / E7 산출내역 키워드 / G7 안내 / H7 메모 / I7·J7 [찾기][취소]
    ws.Range("D7").Value2 = "산출내역 찾기 ▶": ws.Range("G7").Value2 = "메모 찾기 ▶"
    With ws.Range("D7,G7")
        .Font.Bold = True: .Font.name = FONT_NM: .Font.Size = 9: .Font.color = ClrNavy(): .HorizontalAlignment = xlRight: .VerticalAlignment = xlCenter
    End With
    With ws.Range(CALC_CELL & "," & MEMO_CELL)
        .Interior.color = RGB(255, 255, 224): .Borders.LineStyle = xlContinuous: .Borders.color = ClrGray()
        .Font.name = FONT_NM: .VerticalAlignment = xlCenter
    End With
    ws.rows(7).RowHeight = 24
    ws.rows(8).RowHeight = 6

    '표 (머리글 연도 = 세출 파일에서 인식한 연도)
    gStep = "산출내역 표"
    hCalc = YH_Label(yCur, "산출내역", "올해 산출내역")
    hCalcP = YH_Label(yPrv, "산출내역", "작년 산출내역(없음)")
    hBud = YH_Label(yCur, "예산", "올해 예산")
    hPrv = YH_Label(yPrv, "예산", "작년 예산(없음)")
    ws.Cells(C_HDR, 1).Resize(1, nC).Value2 = Array("No", "세부사업", "사업항목", "사업명", hCalc, hCalcP, "세목코드", hBud, hPrv, _
        "장표 예산", "차이(세출-장표)", "차이(산출합-장표)", "차수", "메모", "비고", _
        "원가통계목", "레벨", "팀/부서", "담당자", "단위과제카드", "키", LAST_HDR)
    ws.Cells(C_HDR + 1, 1).Resize(x, nC).Value2 = out
    Set rng = ws.Range(ws.Cells(C_HDR, 1), ws.Cells(C_HDR + x, nC))
    Set lo = MakeTable(ws, rng, TBL_CALC)
    StyleHeader lo.HeaderRowRange
    StyleBody lo.DataBodyRange
    With ws.Range(ws.Cells(C_HDR + 1, K_BUD), ws.Cells(C_HDR + x, K_DIF2))
        .NumberFormat = "#,##0;[Red]-#,##0;""-""": .HorizontalAlignment = xlRight: .ShrinkToFit = True
    End With
    ws.Range(ws.Cells(C_HDR + 1, K_NO), ws.Cells(C_HDR + x, K_NO)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(C_HDR + 1, K_CODE), ws.Cells(C_HDR + x, K_CODE)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(C_HDR + 1, K_CHA), ws.Cells(C_HDR + x, K_CHA)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(C_HDR + 1, K_CALC), ws.Cells(C_HDR + x, K_CALCP)).WrapText = False
    ws.Range(ws.Cells(C_HDR + 1, K_MEMO), ws.Cells(C_HDR + x, K_MEMO)).Interior.color = RGB(255, 255, 224)
    '조건부 서식: 1레벨 파랑·굵게, 2레벨 굵게, 확인 필요(비고 있음) 행 분홍(우선), 차수 칸 색(최우선)
    colLv = ColL(ws, K_LV): colBad = ColL(ws, K_BAD): colCha = ColL(ws, K_CHA)
    Set fC = lo.DataBodyRange.FormatConditions.Add(Type:=xlExpression, formula1:="=$" & colLv & (C_HDR + 1) & "=""1""")
    fC.Interior.color = RGB(229, 240, 255): fC.Font.Bold = True
    Set fC = lo.DataBodyRange.FormatConditions.Add(Type:=xlExpression, formula1:="=$" & colLv & (C_HDR + 1) & "=""2""")
    fC.Font.Bold = True
    Set fC = lo.DataBodyRange.FormatConditions.Add(Type:=xlExpression, _
        formula1:="=$" & colBad & (C_HDR + 1) & "<>""""")
    fC.Interior.color = MismatchColor(): fC.Font.color = RGB(156, 0, 6): fC.Font.Bold = True
    fC.SetFirstPriority
    For s = nS To 0 Step -1
        Set fC = lo.ListColumns(K_CHA).DataBodyRange.FormatConditions.Add(Type:=xlExpression, _
            formula1:="=$" & colCha & (C_HDR + 1) & "=""" & SuppLabel(s) & """")
        fC.Interior.color = ChasuColor(s): fC.Font.color = ChasuColorDark(s): fC.Font.Bold = True
        fC.SetFirstPriority
    Next s
    '합계 (세목코드가 있는 말단 행만, 슬라이서 반영)
    lo.ShowTotals = True
    lo.ListColumns(K_CALC).total.Value2 = "합계 (세목 행 기준)"
    lo.ListColumns(K_BUD).total.Formula = LeafSum(hBud)
    lo.ListColumns(K_PRV).total.Formula = LeafSum(hPrv)
    lo.ListColumns(K_JANG).total.Formula = "=SUBTOTAL(109," & ws.Range(ws.Cells(C_HDR + 1, K_JANG), ws.Cells(C_HDR + x, K_JANG)).Address & ")"
    lo.ListColumns(K_DIF).total.Formula = "=SUBTOTAL(109," & ws.Range(ws.Cells(C_HDR + 1, K_DIF), ws.Cells(C_HDR + x, K_DIF)).Address & ")"
    lo.ListColumns(K_DIF2).total.Formula = "=SUBTOTAL(109," & ws.Range(ws.Cells(C_HDR + 1, K_DIF2), ws.Cells(C_HDR + x, K_DIF2)).Address & ")"
    With lo.TotalsRowRange
        .Font.name = FONT_NM: .Font.Bold = True: .Font.color = ClrNavy(): .Interior.color = ClrLight()
        .Borders(xlEdgeTop).LineStyle = xlContinuous: .Borders(xlEdgeTop).weight = xlMedium: .Borders(xlEdgeTop).color = ClrGray()
    End With
    ws.Range(lo.TotalsRowRange.Cells(1, K_BUD), lo.TotalsRowRange.Cells(1, K_DIF2)).NumberFormat = "#,##0;[Red]-#,##0"

    '열 너비 / 숨김
    ws.Columns(K_NO).ColumnWidth = 5: ws.Columns(K_SEB).ColumnWidth = 16: ws.Columns(K_NM).ColumnWidth = 22
    ws.Columns(K_SNM).ColumnWidth = 26: ws.Columns(K_CALC).ColumnWidth = 44: ws.Columns(K_CALCP).ColumnWidth = 40
    ws.Columns(K_CODE).ColumnWidth = 9
    ws.Columns(K_BUD).ColumnWidth = 12: ws.Columns(K_PRV).ColumnWidth = 12
    ws.Columns(K_JANG).ColumnWidth = 12: ws.Columns(K_DIF).ColumnWidth = 12: ws.Columns(K_DIF2).ColumnWidth = 12
    ws.Columns(K_CHA).ColumnWidth = 11: ws.Columns(K_MEMO).ColumnWidth = 30: ws.Columns(K_BAD).ColumnWidth = 58
    '비고가 길면(사유 2개) 칸 안에서 줄바꿈 - 오른쪽 슬라이서에 가려지지 않게
    Dim rr As Long
    ws.Range(ws.Cells(C_HDR + 1, K_BAD), ws.Cells(C_HDR + x, K_BAD)).WrapText = True
    For rr = C_HDR + 1 To C_HDR + x
        If Len(CStr(ws.Cells(rr, K_BAD).Value2)) > 26 Then ws.rows(rr).AutoFit
    Next rr
    ws.Range(ws.Columns(K_STAT), ws.Columns(K_SRC)).hidden = True
    '비고(확인 필요 사유)가 한 칸도 없으면 비고 열을 숨겨 슬라이서가 메모 옆에 오게 함
    ws.Columns(K_BAD).hidden = (Application.WorksheetFunction.CountA(ws.Range(ws.Cells(C_HDR + 1, K_BAD), ws.Cells(C_HDR + x, K_BAD))) = 0)
    If Not anyPrvCalc Then ws.Columns(K_CALCP).hidden = True       '작년 산출내역이 없는 파일(연도 하나)
    If Not anyPrv And yPrv = 0 Then ws.Columns(K_PRV).hidden = True
    If Not jl Then ws.Range(ws.Columns(K_JANG), ws.Columns(K_DIF2)).hidden = True
    ws.Columns(nC + 1).ColumnWidth = 2
    '찾기/취소 버튼: I7 · J7 셀에 고정
    With ws.Cells(7, K_PRV)
        AddButton ws, "btnFind", "찾기", "메모찾기", .Left + 2, .top + 2, 56, .Height - 4, ClrBlue()
    End With
    With ws.Cells(7, K_JANG)
        AddButton ws, "btnFindOff", "취소", "메모찾기취소", .Left + 2, .top + 2, 56, .Height - 4, ClrGray()
    End With

    '슬라이서: 표 오른쪽 세로 배치 (차수 슬라이서 추가)
    gStep = "산출내역 슬라이서"
    sx = ws.Cells(C_HDR, nC + 2).Left: sy = 4
    AddSlicerTo lo, ws, "차수", SC_CAL & "차수", sx, sy, 220, 80, 2: sy = sy + 84
    AddSlicerTo lo, ws, "팀/부서", SC_CAL & "팀", sx, sy, 220, 100, 2: sy = sy + 104
    AddSlicerTo lo, ws, "담당자", SC_CAL & "담당자", sx, sy, 220, 120, 2: sy = sy + 124
    AddSlicerTo lo, ws, "단위과제카드", SC_CAL & "단위과제카드", sx, sy, 220, 110, 1: sy = sy + 114
    AddSlicerTo lo, ws, "세부사업", SC_CAL & "세부사업", sx, sy, 220, 110, 1: sy = sy + 114
    AddSlicerTo lo, ws, "사업명", SC_CAL & "사업명", sx, sy, 220, 150, 1: sy = sy + 154
    AddSlicerTo lo, ws, "레벨", SC_CAL & "레벨", sx, sy, 220, 60, 4: sy = sy + 64
    AddSlicerTo lo, ws, "세목코드", SC_CAL & "세목코드", sx, sy, 220, 110, 3: sy = sy + 114
    AddSlicerTo lo, ws, "원가통계목", SC_CAL & "원가통계목", sx, sy, 220, 110, 2

    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Range("A1").Select
    ws.Cells(C_HDR + 1, K_CALC).Select
    ActiveWindow.FreezePanes = True
    ActiveWindow.DisplayGridlines = False
    ws.Range("A1").Select
End Sub

'블록 금액 지문: 레벨|세목코드|금액(원) 나열 (산출 문구는 비교하지 않음)
Private Function BlockSig(ByVal col As Collection) As String
    Dim ln As Variant, s As String
    For Each ln In col
        s = s & ln(9) & "|" & ln(5) & "|" & Format$(ln(13), "0") & ";"
    Next ln
    BlockSig = s
End Function

'본예산 블록의 작년 값: (레벨|세목코드|순서) → Array(작년 산출내역, 작년 예산)
Private Function PrevLineMap(ByVal col As Collection) As Object
    Dim d As Object, occ As Object, ln As Variant, k As String
    Set d = CreateObject("Scripting.Dictionary"): Set occ = CreateObject("Scripting.Dictionary")
    For Each ln In col
        k = ln(9) & "|" & ln(5)
        occ(k) = occ(k) + 1
        d(k & "|" & occ(k)) = Array(ln(4), ln(7))
    Next ln
    Set PrevLineMap = d
End Function

'세목코드가 있는 말단 행만, 보이는 행 합계
Private Function LeafSum(ByVal col As String) As String
    LeafSum = "=SUMPRODUCT((" & TBL_CALC & "[세목코드]<>"""")*SUBTOTAL(109,OFFSET(" & TBL_CALC & "[" & col & "],ROW(" & TBL_CALC & "[" & col & "])-MIN(ROW(" & TBL_CALC & "[" & col & "])),0,1)))"
End Function

'V8: 메모 자동 표시 '1추' · '1추·2추' (추경 차수 번호 = 파일 내용으로 판단한 차수)
Private Function AddTag(ByVal a As String, ByVal s As Long) As String
    Dim T As String
    T = CLng(val(CfgGet("추경" & s & "차수"))) & "추"
    If T = "0추" Then T = s & "추"
    If a = "" Then
        AddTag = T
    ElseIf InStr("·" & a & "·", "·" & T & "·") = 0 Then
        AddTag = a & "·" & T
    Else
        AddTag = a
    End If
End Function

Private Function JoinMemo(ByVal userMemo As String, ByVal tag As String) As String
    If Trim$(userMemo) = "" Then JoinMemo = tag Else JoinMemo = userMemo & " " & tag
End Function

'메모 끝의 자동 표시('1추', '1추·2추')를 떼어 사용자 메모만 남김
Private Function StripTag(ByVal m As String) As String
    Static re As Object
    If re Is Nothing Then
        Set re = CreateObject("VBScript.RegExp")
        re.Pattern = "\s*\d+추(·\d+추)*$"
    End If
    StripTag = Trim$(re.Replace(m, ""))
End Function

'"1. 사업명" → "사업명" (UBIS 입력 순서 번호일 뿐 고유번호가 아니고 업무포털에도 표시되지 않음)
Private Function StripNo(ByVal e As String) As String
    If IsLevel1(e) Then StripNo = Trim(Mid(e, InStr(e, ".") + 1)) Else StripNo = e
End Function

Private Function IsLevel1(ByVal e As String) As Boolean
    Dim p As Long
    p = InStr(e, ".")
    If p < 2 Or p > 4 Then Exit Function
    IsLevel1 = IsNumeric(Left(e, p - 1))
End Function


