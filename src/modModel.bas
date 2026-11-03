Attribute VB_Name = "modModel"
Option Explicit
'=====================================================================
' modModel : 요구서(본예산 + 추경 차수) + 장표 → 사업항목 단위 통합 모델(gItems)  - V8
'=====================================================================
Public gItems As Object     'Key -> clsProject (요구서/장표 통합)
Public gByB As Object       '장표키 -> clsProject
Public gStaff As Object     '성명 -> 팀
Public gBAgg As Object      '장표키 -> clsProject(집계용)
Public gReqItems As Object  'V8: 세출 사업항목(본예산 + 추경 반영, 키 = 세출 키, .장표키 = 짝이 된 장표 키)
Public gChasu As Object     'V8: 차수 번호(0=본예산, 1..=추경) -> Dictionary(세출 키 -> clsProject, 그 파일의 금액)
Private mJangCol As Long    'V8: 장표 기준값 열 (설정 '장표기준열' 이 있으면 그 열, 없으면 예산현액)

Public Function ModelReady() As Boolean
    ModelReady = Not (gItems Is Nothing)
End Function

'모델 전체 구축
Public Sub BuildModel()
    Dim ovr As Object
    Set gChasu = Nothing: Set gReqItems = Nothing
    Set gItems = CreateObject("Scripting.Dictionary")
    Set gByB = CreateObject("Scripting.Dictionary")
    Set gBAgg = CreateObject("Scripting.Dictionary")
    Set gStaff = ReadStaffTeams()
    Set ovr = ReadCardOverrides()

    If SheetExists(SH_REQ) Then ParseRequest ovr
    If SheetExists(SH_B) Then
        AggregateB
    ElseIf SheetExists(SH_A) Then
        AggregateA
    End If
    MatchAll ovr
    AssignTeams
End Sub

'---------------------------------------------------------------- 요구서 파싱 (V8: 본예산 + 추경 차수)
'  본예산 시트로 시작 → 추경 시트를 추가한 순서대로 반영.
'  같은 사업항목 키(세부사업|사업항목)는 나중 차수 금액으로 교체(더하지 않음), 새 키는 추가, 추경에 없는 사업은 유지.
Private Sub ParseRequest(ovr As Object)
    Dim n As Long, i As Long, d As Object, k As Variant, src As clsProject, cur As clsProject
    Set gChasu = CreateObject("Scripting.Dictionary")
    Set d = ParseReqSheet(BW().Worksheets(SH_REQ), 0)
    gChasu.Add 0, d
    For Each k In d.keys
        Set cur = CloneItem(d(k))
        ApplyOverride cur, ovr
        gItems.Add k, cur
    Next k
    n = SuppCount()
    For i = 1 To n
        If SheetExists(SuppSheetName(i)) Then
            Set d = ParseReqSheet(BW().Worksheets(SuppSheetName(i)), i)
            gChasu.Add i, d
            For Each k In d.keys
                Set src = d(k)
                If gItems.Exists(k) Then
                    Set cur = gItems(k)
                    If Abs(cur.요구액 - src.요구액) >= 1 Or Abs(cur.산출합 - src.산출합) >= 1 Then src.변경 = "변경" Else src.변경 = "동일"
                    cur.요구액 = src.요구액                '최신 차수로 교체 (더하지 않음)
                    cur.산출합 = src.산출합
                    If src.전년도액 <> 0 Then cur.전년도액 = src.전년도액
                    If src.변경 = "변경" Then
                        cur.차수 = src.차수: cur.차수번호 = src.차수번호   '금액이 같으면 앞 차수 표시 유지
                        CopyCalcEval src, cur                            'V8: 산출내역 비교도 최종 금액을 정한 차수 파일 기준
                    End If
                    cur.담당자자동 = JoinUnique(cur.담당자자동, src.담당자자동)
                Else
                    src.변경 = "추가"
                    Set cur = CloneItem(src)
                    cur.순번 = gItems.Count + 1
                    ApplyOverride cur, ovr
                    gItems.Add k, cur
                End If
            Next k
        End If
    Next i
End Sub

'세출 파일 시트 하나 → 사업항목(1레벨) 사전 (키 → clsProject: 요구액·전년도액·산출합은 원 단위)
'  열은 머리글 이름·연도로 찾는다 (modYearHdr.YH_Map). 산출합 = 세목코드가 있는 행 금액 합 (세목코드 열이 없으면 -1)
'  V8: 산출식합 = 세목 행 최신 산출내역 문구의 계산값 합(세목 행마다 금액 단위로 반올림),
'      산출식전합 = 다른 해(본예산 파일: 작년 산출내역) 또는 기정 산출내역(추경 파일) 계산값 합
Private Function ParseReqSheet(ws As Worksheet, ByVal idx As Long) As Object
    Dim d As Object, mp As Object, hr As Long, arr As Variant, lr As Long, lc As Long
    Dim cLv As Long, cE As Long, cI As Long, cG As Long, ch As Long, cM As Long, cAC As Long, cQ As Long, cCode As Long, cOld As Long
    Dim r As Long, e As String, lv As String, m As String, nm As String, p As Long, unit As Double, oldKind As String
    Dim cur As clsProject, k As String, seq As Long, cfgY As String, amtL As Double, W As Double, tOld As String
    Set d = CreateObject("Scripting.Dictionary")
    Set ParseReqSheet = d
    If idx = 0 Then cfgY = "올해연도" Else cfgY = "추경" & idx & "연도"
    Set mp = YH_Map(ws, SuppFile(idx), cfgY, Not gBX_NoUI)
    hr = mp("hr"): unit = mp("unit"): lc = mp("lastC")
    cLv = mp("lv"): cE = mp("e"): cI = mp("i"): cG = mp("g"): ch = mp("h")
    cM = mp("calc"): cQ = mp("amtP"): cAC = mp("ac"): cCode = mp("code")
    If idx = 0 Then
        CfgSet "올해연도", CStr(mp("yCur")): CfgSet "작년연도", CStr(mp("yPrv")): CfgSet "연도출처", CStr(mp("ySrc"))
    ElseIf mp("yPrv") <> CLng(val(CfgGet("작년연도"))) Then
        cQ = 0                                   '추경 파일의 작년 열은 본예산 작년 연도와 같을 때만 사용
    End If
    '비교용 다른 산출내역 열: 추경 파일은 기정 산출내역, 본예산 파일은 작년 산출내역
    If mp("calcG") > 0 Then
        cOld = mp("calcG"): oldKind = "기정"
    ElseIf mp("calcP") > 0 And (idx = 0 Or cQ > 0) Then
        cOld = mp("calcP"): oldKind = "작년"
    End If
    lr = Application.WorksheetFunction.Max(lastRow(ws, cE), lastRow(ws, cM), lastRow(ws, mp("amt")))
    If mp("gd") Then lr = Application.WorksheetFunction.Max(lr, lastRow(ws, mp("delta")))
    If lr <= hr Then Exit Function
    arr = ws.Range(ws.Cells(hr + 1, 1), ws.Cells(lr, lc)).Value2

    seq = 0
    For r = 1 To UBound(arr, 1)
        lv = CleanText(arr(r, cLv)): e = CleanText(arr(r, cE)): m = CleanText(arr(r, cM))
        If lv = "" And e <> "" And IsLevel1Name(e) Then
            '사업항목(1레벨) 시작
            p = InStr(e, ".")
            nm = Trim(Mid(e, p + 1))
            k = MakeKey(CleanText(arr(r, cI)), nm)
            If d.Exists(k) Then
                Set cur = d(k)
                cur.요구액 = cur.요구액 + YH_AmtAt(arr, r, mp) * unit
                If cQ > 0 Then cur.전년도액 = cur.전년도액 + NzD(arr(r, cQ)) * unit
            Else
                Set cur = New clsProject
                seq = seq + 1
                cur.key = k
                cur.순번 = seq
                cur.세부사업 = CleanText(arr(r, cI))
                cur.사업항목 = nm
                If cG > 0 Then cur.정책사업 = CleanText(arr(r, cG))
                If ch > 0 Then cur.단위사업 = CleanText(arr(r, ch))
                cur.요구액 = YH_AmtAt(arr, r, mp) * unit
                If cQ > 0 Then cur.전년도액 = NzD(arr(r, cQ)) * unit
                cur.매칭 = "요구서"
                cur.차수번호 = idx
                cur.차수 = SuppLabel(idx)
                If cCode = 0 Then cur.산출합 = -1: cur.산출식상태 = 2
                cur.산출식전종류 = oldKind
                cur.산출식전상태 = 2                        '다른 산출내역 문구가 하나라도 있으면 0
                d.Add k, cur
            End If
            If cAC > 0 Then cur.담당자자동 = JoinUnique(cur.담당자자동, CleanText(arr(r, cAC)))
        ElseIf Not cur Is Nothing Then
            If m = "소계" Or lv = "합계" Then
                Set cur = Nothing
            Else
                If cAC > 0 Then cur.담당자자동 = JoinUnique(cur.담당자자동, CleanText(arr(r, cAC)))
                If cCode > 0 Then
                    If CleanText(arr(r, cCode)) <> "" Then
                        amtL = YH_AmtAt(arr, r, mp) * unit
                        cur.산출합 = cur.산출합 + amtL
                        '최신 산출내역 계산
                        W = YH_EvalCalcWon(CellText(arr(r, cM)))
                        If W < 0 Then
                            If amtL <> 0 And cur.산출식상태 = 0 Then cur.산출식상태 = 1
                        Else
                            cur.산출식합 = cur.산출식합 + YH_CalcRound(W, unit): cur.산출식허용 = cur.산출식허용 + YH_CalcTol(W, unit)
                        End If
                        '다른 산출내역 계산 (작년 / 기정)
                        If cOld > 0 Then
                            tOld = CellText(arr(r, cOld))
                            If Trim$(tOld) <> "" Then
                                If cur.산출식전상태 = 2 Then cur.산출식전상태 = 0
                                W = YH_EvalCalcWon(tOld)
                                If W < 0 Then
                                    cur.산출식전상태 = 1
                                Else
                                    cur.산출식전합 = cur.산출식전합 + YH_CalcRound(W, unit): cur.산출식전허용 = cur.산출식전허용 + YH_CalcTol(W, unit)
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next r
End Function

Private Function CellText(ByVal v As Variant) As String
    If IsError(v) Or IsEmpty(v) Then CellText = "" Else CellText = CStr(v)
End Function

Private Function CloneItem(ByVal s As clsProject) As clsProject
    Dim T As clsProject
    Set T = New clsProject
    T.key = s.key: T.순번 = s.순번: T.세부사업 = s.세부사업: T.사업항목 = s.사업항목
    T.정책사업 = s.정책사업: T.단위사업 = s.단위사업: T.요구액 = s.요구액: T.전년도액 = s.전년도액
    T.담당자자동 = s.담당자자동: T.매칭 = s.매칭: T.차수 = s.차수: T.차수번호 = s.차수번호
    T.산출합 = s.산출합: T.변경 = s.변경
    CopyCalcEval s, T
    Set CloneItem = T
End Function

'V8: 산출식 계산값 복사 (src → dst)
Public Sub CopyCalcEval(ByVal s As clsProject, ByVal T As clsProject)
    T.산출식합 = s.산출식합: T.산출식허용 = s.산출식허용: T.산출식상태 = s.산출식상태
    T.산출식전합 = s.산출식전합: T.산출식전허용 = s.산출식전허용: T.산출식전상태 = s.산출식전상태: T.산출식전종류 = s.산출식전종류
End Sub

'"1. 사업명" 형태인지
Private Function IsLevel1Name(ByVal e As String) As Boolean
    Dim p As Long
    p = InStr(e, ".")
    If p < 2 Or p > 4 Then IsLevel1Name = False: Exit Function
    IsLevel1Name = IsNumeric(Left(e, p - 1))
End Function

'---------------------------------------------------------------- 장표 102-B 집계 (세목이 있는 말단행 합산)
Private Sub AggregateB()
    Set gBAgg = AggregateFromSheet(BW().Worksheets(SH_B), "")
End Sub

'102-B 형식 시트에서 사업항목(1레벨) 집계. deptFilter 가 있으면 예산소관부서명에 포함된 행만.
Public Function AggregateFromSheet(ws As Worksheet, ByVal deptFilter As String) As Object
    Dim hr As Long, hm As Object, arr As Variant, r As Long, lr As Long, lc As Long, d As Object
    Dim cGb As Long, cUnit As Long, cPol As Long, cDan As Long, cSeb As Long, cN1 As Long, cCode As Long, cDept As Long
    Dim cHy As Long, cWon As Long, cJi As Long, cGyo As Long, cGeup As Long
    Dim k As String, a As clsProject, seq As Long

    Set d = CreateObject("Scripting.Dictionary")
    Set AggregateFromSheet = d
    hr = FindHeaderRow(ws, "회계연도", 10)
    If hr = 0 Then Exit Function
    Set hm = HeaderMap(ws, hr)
    If Not hm.Exists("사업항목명1") Then Exit Function
    cGb = HCol(hm, "집행예산구분"): cUnit = HCol(hm, "단위업무"): cPol = HCol(hm, "정책사업", False)
    cDan = HCol(hm, "단위사업", False): cSeb = HCol(hm, "세부사업"): cN1 = HCol(hm, "사업항목명1")
    cCode = HCol(hm, "세출세목코드"): cHy = HCol(hm, "예산현액"): cWon = HCol(hm, "원인행위금액")
    cJi = HCol(hm, "지출금액"): cGyo = HCol(hm, "일상경비교부액", False): cGeup = HCol(hm, "일상경비지급액", False)
    cDept = HCol(hm, "예산소관부서명", False)
    mJangCol = 0
    If CfgGet("장표기준열") <> "" Then mJangCol = HCol(hm, CfgGet("장표기준열"), False)

    lr = lastRow(ws, 1): lc = ws.Cells(hr, ws.Columns.Count).End(xlToLeft).Column
    If lr <= hr Then Exit Function
    arr = ws.Range(ws.Cells(hr + 1, 1), ws.Cells(lr, lc)).Value2

    For r = 1 To UBound(arr, 1)
        If CleanText(arr(r, cCode)) <> "" Then
            If deptFilter = "" Or cDept = 0 Or InStr(1, CleanText(arr(r, cDept)), deptFilter, vbTextCompare) > 0 Then
                k = MakeKey(CleanText(arr(r, cSeb)), CleanText(arr(r, cN1)))
                If d.Exists(k) Then
                    Set a = d(k)
                Else
                    Set a = New clsProject
                    seq = seq + 1
                    a.key = k: a.장표키 = k: a.순번 = seq
                    a.세부사업 = CleanText(arr(r, cSeb))
                    a.장표명 = CleanText(arr(r, cN1))
                    If a.장표명 = "" Then a.장표명 = "(사업항목 없음)"
                    a.사업항목 = a.장표명
                    If cPol > 0 Then a.정책사업 = CleanText(arr(r, cPol))
                    If cDan > 0 Then a.단위사업 = CleanText(arr(r, cDan))
                    d.Add k, a
                End If
                AddAmounts a, arr, r, cGb, cHy, cWon, cJi, cGyo, cGeup
                a.카드자동 = JoinUnique(a.카드자동, CleanText(arr(r, cUnit)))
            End If
        End If
    Next r
End Function

'102-A만 있는 경우 집계 (사업항목1레벨 × 세목 행)
Private Sub AggregateA()
    Dim ws As Worksheet, hr As Long, hm As Object, arr As Variant, r As Long, lr As Long, lc As Long
    Dim cGb As Long, cUnit As Long, cPol As Long, cDan As Long, cSeb As Long, cN1 As Long
    Dim cHy As Long, cWon As Long, cJi As Long, cGyo As Long, cGeup As Long
    Dim k As String, a As clsProject

    Set ws = BW().Worksheets(SH_A)
    hr = FindHeaderRow(ws, "회계연도", 10)
    If hr = 0 Then Exit Sub
    Set hm = HeaderMap(ws, hr)
    cGb = HCol(hm, "집행예산구분"): cUnit = HCol(hm, "단위업무"): cPol = HCol(hm, "정책사업", False)
    cDan = HCol(hm, "단위사업", False): cSeb = HCol(hm, "세부사업"): cN1 = HCol(hm, "사업항목1레벨")
    cHy = HCol(hm, "예산현액"): cWon = HCol(hm, "원인행위금액"): cJi = HCol(hm, "지출금액")
    cGyo = HCol(hm, "일상경비교부액", False): cGeup = HCol(hm, "일상경비지급액", False)
    mJangCol = 0
    If CfgGet("장표기준열") <> "" Then mJangCol = HCol(hm, CfgGet("장표기준열"), False)

    lr = lastRow(ws, 1): lc = ws.Cells(hr, ws.Columns.Count).End(xlToLeft).Column
    If lr <= hr Then Exit Sub
    arr = ws.Range(ws.Cells(hr + 1, 1), ws.Cells(lr, lc)).Value2
    For r = 1 To UBound(arr, 1)
        k = MakeKey(CleanText(arr(r, cSeb)), CleanText(arr(r, cN1)))
        If gBAgg.Exists(k) Then
            Set a = gBAgg(k)
        Else
            Set a = New clsProject
            a.key = k: a.장표키 = k
            a.세부사업 = CleanText(arr(r, cSeb))
            a.장표명 = CleanText(arr(r, cN1))
            If cPol > 0 Then a.정책사업 = CleanText(arr(r, cPol))
            If cDan > 0 Then a.단위사업 = CleanText(arr(r, cDan))
            gBAgg.Add k, a
        End If
        AddAmounts a, arr, r, cGb, cHy, cWon, cJi, cGyo, cGeup
        a.카드자동 = JoinUnique(a.카드자동, CleanText(arr(r, cUnit)))
    Next r
End Sub

Private Sub AddAmounts(a As clsProject, arr As Variant, ByVal r As Long, ByVal cGb As Long, ByVal cHy As Long, _
                       ByVal cWon As Long, ByVal cJi As Long, ByVal cGyo As Long, ByVal cGeup As Long)
    Dim hy As Double, gb As String
    hy = NzD(arr(r, cHy))
    a.예산현액 = a.예산현액 + hy
    a.원인행위 = a.원인행위 + NzD(arr(r, cWon))
    a.집행액 = a.집행액 + ExecAmount(arr, r, cJi, cGyo, cGeup)
    gb = CleanText(arr(r, cGb))
    If InStr(gb, "예산") > 0 Then
        a.본예산현액 = a.본예산현액 + hy
        'V8: 장표 예산금액(기준값) = 집행예산구분 '예산' 행의 예산현액 (설정 '장표기준열' 이 있으면 그 열)
        If mJangCol > 0 Then a.장표예산 = a.장표예산 + NzD(arr(r, mJangCol)) Else a.장표예산 = a.장표예산 + hy
    Else
        a.이월현액 = a.이월현액 + hy
    End If
End Sub

'실집행액 = 지출금액 - 일상경비교부액 + 일상경비지급액
Public Function ExecAmount(arr As Variant, ByVal r As Long, ByVal cJi As Long, ByVal cGyo As Long, ByVal cGeup As Long) As Double
    Dim v As Double
    v = NzD(arr(r, cJi))
    If cGyo > 0 And cGeup > 0 Then v = v - NzD(arr(r, cGyo)) + NzD(arr(r, cGeup))
    ExecAmount = v
End Function

'---------------------------------------------------------------- 매칭
Private Sub MatchAll(ovr As Object)
    Dim k As Variant, it As clsProject, a As clsProject, bk As String, used As Object, req As Object
    Dim rq As clsProject
    Set used = CreateObject("Scripting.Dictionary")
    Set req = gItems                      '요구서 항목 (담당자 정보용)
    Set gReqItems = req                   'V8: 장표 대조용으로 보관
    Set gItems = CreateObject("Scripting.Dictionary")

    If gBAgg.Count = 0 Then
        '장표가 없으면 요구서 항목으로 보드 구성 (요구액 = 예산현액)
        For Each k In req.keys
            Set it = req(k)
            it.예산현액 = it.요구액: it.본예산현액 = it.요구액: it.매칭 = "요구서"
            gItems.Add k, it
        Next k
        Exit Sub
    End If

    '1) 요구서 항목의 장표 키 결정 (직접지정 > 완전일치 > 유사)
    For Each k In req.keys
        Set rq = req(k)
        bk = ""
        If Trim(rq.장표지정) <> "" Then bk = FindBKey(rq.세부사업, rq.장표지정, True)
        If bk = "" Then
            If gBAgg.Exists(rq.key) Then bk = rq.key Else bk = FindBKey(rq.세부사업, rq.사업항목, False)
        End If
        If bk <> "" Then
            If Not used.Exists(bk) Then
                used.Add bk, rq
                rq.장표키 = bk                 'V8: 세출 항목 → 장표 키 연결
            End If
        End If
    Next k

    '2) 장표 항목이 기준 (장표 순서). 요구서가 있으면 담당자·요구액·전년도 붙임
    For Each k In gBAgg.keys
        Set a = gBAgg(k)
        Set it = New clsProject
        it.key = a.key: it.순번 = a.순번
        it.세부사업 = a.세부사업: it.사업항목 = a.장표명: it.장표명 = a.장표명: it.장표키 = k
        it.정책사업 = a.정책사업: it.단위사업 = a.단위사업
        it.예산현액 = a.예산현액: it.원인행위 = a.원인행위: it.집행액 = a.집행액
        it.본예산현액 = a.본예산현액: it.이월현액 = a.이월현액: it.장표예산 = a.장표예산
        it.카드자동 = a.카드자동
        If used.Exists(k) Then
            Set rq = used(k)
            it.담당자자동 = rq.담당자자동: it.요구액 = rq.요구액: it.전년도액 = rq.전년도액
            it.차수 = rq.차수: it.차수번호 = rq.차수번호: it.산출합 = rq.산출합
            it.매칭 = "일치"
        Else
            it.매칭 = IIf(req.Count = 0, "장표", "요구서없음")
        End If
        ApplyOverride it, ovr
        gItems.Add it.key, it
        gByB.Add k, it
    Next k
End Sub

'장표 키 탐색: exact=True 는 정규화 완전일치(세부사업 무시 허용), False 는 포함관계(같은 세부사업 내)
Private Function FindBKey(ByVal 세부사업 As String, ByVal nm As String, ByVal exact As Boolean) As String
    Dim k As Variant, nk As String, sk As String, parts As Variant, best As String, bestLen As Long
    nk = NormKey(nm): sk = NormKey(세부사업)
    If nk = "" Then Exit Function
    If gBAgg.Exists(sk & "|" & nk) Then FindBKey = sk & "|" & nk: Exit Function
    For Each k In gBAgg.keys
        parts = Split(k, "|")
        If exact Then
            If parts(1) = nk Then FindBKey = k: Exit Function
        Else
            If parts(0) = sk Then
                If InStr(parts(1), nk) > 0 Or InStr(nk, parts(1)) > 0 Then
                    If Len(parts(1)) > bestLen Then best = k: bestLen = Len(parts(1))
                End If
            End If
        End If
    Next k
    FindBKey = best
End Function

Private Sub ApplyOverride(it As clsProject, ovr As Object)
    Dim v As Variant
    If ovr Is Nothing Then Exit Sub
    If ovr.Exists(it.key) Then
        v = ovr(it.key)
        it.담당자수정 = v(0): it.카드수정 = v(1): it.장표지정 = v(2)
    End If
End Sub

'담당자 → 팀 배정 (복수 담당자면 팀도 중복 없이 연결)
Private Sub AssignTeams()
    Dim k As Variant, it As clsProject, names As Variant, i As Long, T As String, nm As String
    For Each k In gItems.keys
        Set it = gItems(k)
        T = ""
        names = Split(it.담당자, ",")
        For i = LBound(names) To UBound(names)
            nm = CleanText(names(i))
            If nm <> "" Then If gStaff.Exists(nm) Then T = JoinUnique(T, gStaff(nm))
        Next i
        it.팀 = T
    Next k
End Sub

'---------------------------------------------------------------- 정렬된 키 목록 (팀 → 담당자 → 순번)
Public Function SortedKeys() As Collection
    Dim n As Long, keys() As String, sk() As String, i As Long, j As Long, tmpS As String, tmpK As String
    Dim k As Variant, it As clsProject, col As New Collection
    n = gItems.Count
    If n = 0 Then Set SortedKeys = col: Exit Function
    ReDim keys(1 To n): ReDim sk(1 To n)
    i = 0
    For Each k In gItems.keys
        i = i + 1
        Set it = gItems(k)
        keys(i) = k
        sk(i) = Format(it.순번, "00000")      '장표 순서
    Next k
    '삽입 정렬 (항목 수 수백 건 수준)
    For i = 2 To n
        tmpS = sk(i): tmpK = keys(i): j = i - 1
        Do While j >= 1
            If StrComp(sk(j), tmpS, vbTextCompare) <= 0 Then Exit Do
            sk(j + 1) = sk(j): keys(j + 1) = keys(j): j = j - 1
        Loop
        sk(j + 1) = tmpS: keys(j + 1) = tmpK
    Next i
    For i = 1 To n: col.Add keys(i): Next i
    Set SortedKeys = col
End Function

'담당자 목록(중복 제거, 정렬)
Public Function StaffNames() As Collection
    Dim d As Object, k As Variant, it As clsProject, names As Variant, i As Long, nm As String
    Dim arr() As String, n As Long, j As Long, T As String, col As New Collection
    Set d = CreateObject("Scripting.Dictionary")
    If Not gStaff Is Nothing Then
        For Each k In gStaff.keys: d(k) = 1: Next k
    End If
    If Not gItems Is Nothing Then
        For Each k In gItems.keys
            Set it = gItems(k)
            names = Split(it.담당자, ",")
            For i = LBound(names) To UBound(names)
                nm = CleanText(names(i))
                If nm <> "" Then d(nm) = 1
            Next i
        Next k
    End If
    n = d.Count
    If n = 0 Then Set StaffNames = col: Exit Function
    ReDim arr(1 To n): i = 0
    For Each k In d.keys: i = i + 1: arr(i) = k: Next k
    For i = 2 To n
        T = arr(i): j = i - 1
        Do While j >= 1
            If StrComp(arr(j), T, vbTextCompare) <= 0 Then Exit Do
            arr(j + 1) = arr(j): j = j - 1
        Loop
        arr(j + 1) = T
    Next i
    For i = 1 To n: col.Add arr(i): Next i
    Set StaffNames = col
End Function


