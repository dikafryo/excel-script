Attribute VB_Name = "modYearHdr"
Option Explicit
'=====================================================================
' modYearHdr : (V8) 세출 파일 머리글 연도 자동 인식 + 머리글 이름으로 열 찾기
'   - 열 위치(번호)를 고정하지 않는다. 머리글 글자로만 열을 찾는다.
'   - 머리글(또는 바로 위 묶음 머리글)에 '2027년 산출내역', '2026년도 예산액', '2027 산출기초' 처럼
'     연도가 있으면: 큰 연도 = 올해(최신), 작은 연도 = 작년.
'   - 연도가 없는 UBIS 기본 머리글(산출기초 / 예산요구액 / 전년도예산액)은
'     제목 행('2027년도 세출예산요구서' 등) > 파일 이름 > 설정에 저장된 연도 > 직접 입력 순으로 올해 연도를 정한다.
'   - 반환: Scripting.Dictionary
'       hr(머리글 행) lv e i g h(레벨·세세부·세부·정책·단위사업 열) calc calcP(올해·작년 산출내역 열)
'       amt amtP(올해·작년 금액 열) gij(기정액 열) code stat ac(세목코드·원가통계목·담당자 열)
'       yCur yPrv(올해·작년 연도) ySrc(연도를 찾은 곳) unit(금액 단위 배수: 천원=1000, 원=1)
'=====================================================================
Private Const HDR_SCAN As Long = 25

'문자열 안의 4자리 연도(20xx, 앞뒤가 숫자가 아님) 중 가장 큰 값. 없으면 0
Public Function YH_YearIn(ByVal s As String) As Long
    Dim i As Long, y As Long, best As Long, pre As String, nxt As String
    For i = 1 To Len(s) - 3
        If Mid$(s, i, 4) Like "20##" Then
            pre = "": nxt = ""
            If i > 1 Then pre = Mid$(s, i - 1, 1)
            If i + 4 <= Len(s) Then nxt = Mid$(s, i + 4, 1)
            If Not (pre Like "#") And Not (nxt Like "#") Then
                y = CLng(Mid$(s, i, 4))
                If y > best Then best = y
            End If
        End If
    Next i
    YH_YearIn = best
End Function

'"2027년", "2027 년도", "2027회계연도" 처럼 연도 뒤에 '년'·'회계'가 붙은 경우만 (출력일·기준일 같은 날짜는 제외)
Public Function YH_YearWithNen(ByVal s As String) As Long
    Dim i As Long, y As Long, best As Long, pre As String, rest As String
    s = Replace(s, " ", "")
    For i = 1 To Len(s) - 3
        If Mid$(s, i, 4) Like "20##" Then
            pre = ""
            If i > 1 Then pre = Mid$(s, i - 1, 1)
            rest = Mid$(s, i + 4, 2)
            If Not (pre Like "#") Then
                If Left$(rest, 1) = "년" Or rest = "회계" Then
                    y = CLng(Mid$(s, i, 4))
                    If y > best Then best = y
                End If
            End If
        End If
    Next i
    YH_YearWithNen = best
End Function

'표 머리글 글자: 2027 → "2027년 산출내역" / 연도 모름 → 대체 글자
Public Function YH_Label(ByVal y As Long, ByVal what As String, ByVal Fallback As String) As String
    If y > 0 Then YH_Label = y & "년 " & what Else YH_Label = Fallback
End Function

Private Function NonEmptyCount(ws As Worksheet, ByVal r As Long) As Long
    Dim c As Long, lastC As Long, n As Long
    If r < 1 Then Exit Function
    lastC = ws.Cells(r, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastC
        If CleanText(ws.Cells(r, c).Value2) <> "" Then n = n + 1
    Next c
    NonEmptyCount = n
End Function

'세출 파일 머리글 행 찾기: 머리글 후보 행('세부사업'(같은 행 또는 세로 병합 위 행) + 산출/금액 머리글, 또는 '예산요구액')
'  중 자료 행(숫자 칸 2개 이상)이 시작되기 바로 앞의 마지막 후보 행.
'  UBIS 실제 양식: 4~6행 병합 머리글 + 8행 기계용 머리글 → 8행 / 데모 양식: 4~5행 → 5행
Public Function YH_FindHeaderRow(ws As Worksheet) As Long
    Dim r As Long, c As Long, lastC As Long, T As String, v As Variant
    Dim hasSeb As Boolean, hasAmt As Boolean, hasReq As Boolean, nNum As Long, best As Long
    For r = 1 To HDR_SCAN
        hasSeb = False: hasAmt = False: hasReq = False: nNum = 0
        lastC = ws.Cells(r, ws.Columns.Count).End(xlToLeft).Column
        If lastC > 200 Then lastC = 200
        For c = 1 To lastC
            v = ws.Cells(r, c).Value2
            If VarType(v) = vbDouble Then nNum = nNum + 1
            T = Replace(CleanText(v), " ", "")
            If T = "세부사업" Or T = "세세부사업" Then hasSeb = True
            If T = "예산요구액" Then hasReq = True
            If InStr(T, "산출") > 0 Or InStr(T, "요구액") > 0 Or InStr(T, "예산액") > 0 Then hasAmt = True
            '세로 병합 머리글(위아래 2행)은 값 복사 후 위 행에만 글자가 남으므로, 이 칸이 비어 있으면 바로 위 행도 본다
            If Not hasSeb And r > 1 And T = "" Then
                T = Replace(CleanText(ws.Cells(r - 1, c).Value2), " ", "")
                If T = "세부사업" Or T = "세세부사업" Then hasSeb = True
            End If
        Next c
        If best > 0 And nNum >= 2 Then Exit For          '자료 행 시작 → 그 앞의 마지막 머리글 행
        If (hasSeb And hasAmt) Or hasReq Then best = r
    Next r
    YH_FindHeaderRow = best
End Function

'원본 파일의 병합 머리글(묶음 머리글·세로 병합)을 값 복사본의 각 칸에 채운다 (가져오기 직후 1회)
'  - 머리글 행과 그 위 5행까지, 폭 6칸 이하 병합만 (제목 줄처럼 표 전체에 걸친 병합은 제외)
Public Sub YH_FillMerged(wsS As Worksheet, wsD As Worksheet)
    Dim hr As Long, r As Long, c As Long, lastC As Long, ma As Range
    hr = YH_FindHeaderRow(wsD)
    If hr = 0 Then Exit Sub
    On Error Resume Next
    lastC = wsS.UsedRange.Column + wsS.UsedRange.Columns.Count - 1
    If lastC > 200 Then lastC = 200
    For r = IIf(hr > 5, hr - 5, 1) To hr
        For c = 1 To lastC
            If wsS.Cells(r, c).MergeCells Then
                Set ma = wsS.Cells(r, c).MergeArea
                If ma.Columns.Count <= 6 Then
                    If Not (ma.row = r And ma.Column = c) Then wsD.Cells(r, c).Value2 = ma.Cells(1, 1).Value2
                End If
            End If
        Next c
    Next r
    On Error GoTo 0
End Sub

Private Function IsAmtHeader(ByVal u As String) As Boolean
    If InStr(u, "증감") > 0 Or InStr(u, "차이") > 0 Or InStr(u, "비율") > 0 Or InStr(u, "%") > 0 Or InStr(u, "대비") > 0 Then Exit Function
    If InStr(u, "구분") > 0 Or InStr(u, "과목") > 0 Or InStr(u, "요구서") > 0 Or InStr(u, "코드") > 0 Then Exit Function
    IsAmtHeader = (InStr(u, "요구액") > 0 Or InStr(u, "예산액") > 0 Or InStr(u, "금액") > 0 Or InStr(u, "예산") > 0)
End Function

'오류 문구용: 어느 파일인지 (경로 없이 파일 이름만, 없으면 시트 이름)
Private Function YH_Who(ws As Worksheet, ByVal srcFile As String) As String
    Dim p As Long
    If srcFile <> "" Then
        p = InStrRev(srcFile, "\")
        YH_Who = "세출 파일 '" & Mid$(srcFile, p + 1) & "'"
    Else
        YH_Who = "세출 파일(시트 '" & ws.name & "')"
    End If
End Function

'오류 문구용: 머리글 글자 목록 (앞에서 20칸, 빈 칸 제외)
Private Function YH_HdrList(low() As String, ByVal lastC As Long) As String
    Dim c As Long, n As Long, T As String
    For c = 1 To lastC
        If low(c) <> "" Then
            T = T & IIf(T = "", "", " | ") & low(c)
            n = n + 1
            If n >= 20 Then T = T & " | …": Exit For
        End If
    Next c
    If T = "" Then T = "(비어 있음)"
    YH_HdrList = T
End Function

'세출 파일 열 지도 만들기. srcFile = 원본 파일 경로(연도 최후 보조), cfgKey = 설정에 저장된 연도 키("올해연도" 등)
'  allowAsk = True 이면 연도를 못 찾았을 때 사용자에게 묻는다 (False 면 오류)
'  연도 순서: 머리글 > 제목 행 > 저장된 설정 > 파일 이름 > 직접 입력
'  추경 양식(UBIS): 산출기초(기정예산)① · 산출기초(최종요구)② · 추경요구액(②-①) · 기정예산액 · 최종예산
'    → calc = 최종요구 산출, calcG = 기정 산출, gij = 기정예산액, delta = 추경요구액, amt = 최종예산
'    → 최종예산 열이 비어 있으면(UBIS 내려받기) 금액 = 기정예산액 + 추경요구액 (gd = True)
Public Function YH_Map(ws As Worksheet, ByVal srcFile As String, ByVal cfgKey As String, ByVal allowAsk As Boolean) As Object
    Dim d As Object, hr As Long, lastC As Long, c As Long, r As Long, k As Variant, r0 As Long
    Dim low() As String, grp() As String, yr() As Long, kind() As String, ug() As String, rowOk() As Boolean
    Dim T As String, u As String, yMax As Long, y2 As Long, ans As String, fn As String, p As Long
    Dim best As Long, bestScore As Long, sc As Long, pass As Long, want As Long, unitTxt As String
    Dim lr As Long, nVal As Long, v As Variant

    Set d = CreateObject("Scripting.Dictionary")
    hr = YH_FindHeaderRow(ws)
    If hr = 0 Then Err.Raise vbObjectError + 801, "YH_Map", YH_Who(ws, srcFile) & "에서 머리글 행을 찾지 못했습니다." & vbCrLf & _
        "찾은 글자: 같은 행에 '세부사업'과 산출내역/예산 금액 머리글이 함께 있는 행, 또는 '예산요구액' 칸 (위에서 " & HDR_SCAN & "행까지)" & vbCrLf & _
        "UBIS에서 내려받은 세출예산요구서(본예산) 또는 추경 세출 요구 자료 엑셀 원본인지 확인하세요."
    d("hr") = hr
    r0 = IIf(hr > 5, hr - 5, 1)                  '묶음 머리글은 위로 5행까지 (빈 행은 건너뜀)
    lastC = ws.Cells(hr, ws.Columns.Count).End(xlToLeft).Column
    For r = hr - 1 To r0 Step -1
        c = ws.Cells(r, ws.Columns.Count).End(xlToLeft).Column
        If c > lastC Then lastC = c
    Next r
    If lastC > 200 Then lastC = 200
    d("lastC") = lastC
    ReDim low(1 To lastC): ReDim grp(1 To lastC): ReDim yr(1 To lastC): ReDim kind(1 To lastC): ReDim ug(1 To lastC)
    ReDim rowOk(1 To hr)
    For r = r0 To hr - 1
        rowOk(r) = (NonEmptyCount(ws, r) >= 2)
    Next r

    '1) 열마다 머리글 글자(아래 행 우선) + 묶음 머리글(위 행) 모으기
    For c = 1 To lastC
        low(c) = Replace(CleanText(ws.Cells(hr, c).Value2), " ", "")
        For r = hr - 1 To r0 Step -1
            If rowOk(r) Then
                T = Replace(CleanText(ws.Cells(r, c).Value2), " ", "")
                If low(c) = "" Then
                    low(c) = T
                ElseIf T <> "" Then
                    grp(c) = T & " " & grp(c)
                End If
            End If
        Next r
        ug(c) = low(c) & " " & grp(c)
        '2) 종류: 산출내역 / 추경요구액(증감) / 기정액 / 금액
        u = low(c)
        If InStr(u, "산출") > 0 Then
            kind(c) = "calc"
        ElseIf InStr(u, "추경요구액") > 0 Or InStr(u, "②-①") > 0 Then
            kind(c) = "delta"
        ElseIf InStr(u, "기정") > 0 Then
            kind(c) = "gij"
        ElseIf IsAmtHeader(u) And InStr(ug(c), "재원") = 0 Then
            kind(c) = "amt"
        End If
        '3) 연도: 머리글 > 묶음 머리글. '전년도' 글자는 작년(-1)
        If kind(c) <> "" Then
            If InStr(u, "전년") > 0 Then
                yr(c) = -1
            Else
                yr(c) = YH_YearIn(u)
                If yr(c) = 0 Then yr(c) = YH_YearIn(grp(c))
            End If
        End If
    Next c

    '3-1) 연도를 못 찾은 산출내역 · 금액 열: 같은 열을 더 위 행까지 올라가며 '20xx년' 글자를 찾는다
    '     (UBIS 세출예산요구서: 8행 '산출기초 | 산출기초' 위 4행 M4 '2026년도 본예산' · N4 '2027년도 본예산')
    For c = 1 To lastC
        If (kind(c) = "calc" Or kind(c) = "amt") And yr(c) = 0 Then
            For r = r0 - 1 To 1 Step -1
                If NonEmptyCount(ws, r) >= 2 Then
                    p = YH_YearWithNen(CleanText(ws.Cells(r, c).Value2))
                    If p > 0 Then yr(c) = p: Exit For
                End If
            Next r
        End If
    Next c

    '4) 올해(가장 큰 연도) / 작년(그 다음 연도)
    yMax = 0: y2 = 0
    For c = 1 To lastC
        If (kind(c) = "calc" Or kind(c) = "amt") And yr(c) > yMax Then yMax = yr(c)
    Next c
    For c = 1 To lastC
        If (kind(c) = "calc" Or kind(c) = "amt") And yr(c) > 0 And yr(c) < yMax And yr(c) > y2 Then y2 = yr(c)
    Next c

    '5) 열 고르기 (머리글 연도로. 연도 없는 기본 머리글은 보조). 산출내역은 '기정'이 아닌 열(최종요구) 우선
    d("calc") = 0: d("calcP") = 0: d("calcG") = 0: d("amt") = 0: d("amtP") = 0: d("gij") = 0: d("delta") = 0: d("gd") = False
    For pass = 1 To 2
        If pass = 1 Then want = yMax Else want = 0
        If d("calc") = 0 And (pass = 2 Or yMax > 0) Then
            For c = 1 To lastC
                If kind(c) = "calc" And yr(c) = want And InStr(ug(c), "기정") = 0 Then d("calc") = c: Exit For
            Next c
            If d("calc") = 0 Then
                For c = 1 To lastC
                    If kind(c) = "calc" And yr(c) = want Then d("calc") = c: Exit For
                Next c
            End If
        End If
        If d("amt") = 0 And (pass = 2 Or yMax > 0) Then
            best = 0: bestScore = 0
            For c = 1 To lastC
                If kind(c) = "amt" And yr(c) = want Then
                    sc = 2
                    If InStr(low(c), "요구") > 0 Or InStr(low(c), "추경") > 0 Or InStr(low(c), "변경") > 0 Or InStr(low(c), "최종") > 0 Then sc = 3
                    If sc > bestScore Then best = c: bestScore = sc
                End If
            Next c
            d("amt") = best
        End If
    Next pass
    '연도를 끝내 모르는 같은 이름 산출내역 열이 둘 이상(예: '산출기초' | '산출기초')이면
    'UBIS 순서(작년 → 올해)대로 뒤 열 = 올해, 앞 열 = 작년 (앞 열을 올해로 읽어 전년도 자료가 나오지 않게)
    If d("calc") > 0 Then
        If yr(d("calc")) = 0 And InStr(ug(d("calc")), "기정") = 0 Then
            For c = lastC To d("calc") + 1 Step -1
                If kind(c) = "calc" And yr(c) = 0 And low(c) = low(d("calc")) And InStr(ug(c), "기정") = 0 Then
                    d("calcP") = d("calc"): d("calc") = c: Exit For
                End If
            Next c
        End If
    End If
    For c = 1 To lastC
        If kind(c) = "calc" And c <> d("calc") Then
            If InStr(ug(c), "기정") > 0 Then
                If d("calcG") = 0 Then d("calcG") = c
            ElseIf d("calcP") = 0 Then
                If (y2 > 0 And yr(c) = y2) Or yr(c) = -1 Then d("calcP") = c
            End If
        End If
        If kind(c) = "amt" And d("amtP") = 0 And c <> d("amt") Then
            If (y2 > 0 And yr(c) = y2) Or yr(c) = -1 Then d("amtP") = c
        End If
        If kind(c) = "gij" And d("gij") = 0 Then d("gij") = c
        If kind(c) = "delta" And d("delta") = 0 Then d("delta") = c
    Next c
    '최종예산 열이 없거나 비어 있으면 기정예산액 + 추경요구액
    If d("gij") > 0 And d("delta") > 0 Then
        If d("amt") = 0 Then
            d("gd") = True: d("amt") = d("gij")
        Else
            lr = ws.Cells(ws.rows.Count, d("amt")).End(xlUp).row
            nVal = 0
            If lr > hr Then
                For r = hr + 1 To lr
                    v = ws.Cells(r, d("amt")).Value2
                    If VarType(v) = vbDouble Then
                        If v <> 0 Then nVal = nVal + 1: Exit For
                    End If
                Next r
            End If
            If nVal = 0 Then d("gd") = True
        End If
    End If
    If d("calc") = 0 Then Err.Raise vbObjectError + 804, "YH_Map", YH_Who(ws, srcFile) & "에서 산출내역 열을 찾지 못했습니다. (머리글 " & hr & "행)" & vbCrLf & _
        "찾은 글자: 머리글에 '산출'이 들어간 열 (예: '2027년 산출내역', '산출기초', '산출기초(최종요구) ②')" & vbCrLf & _
        "이 파일의 머리글: " & YH_HdrList(low, lastC) & vbCrLf & _
        "UBIS에서 내려받은 원본 엑셀인지(머리글을 고치거나 열을 지우지 않았는지) 확인하세요."
    If d("amt") = 0 Then Err.Raise vbObjectError + 803, "YH_Map", YH_Who(ws, srcFile) & "에서 올해 예산 금액 열을 찾지 못했습니다. (머리글 " & hr & "행)" & vbCrLf & _
        "찾은 글자: 머리글에 '요구액' · '예산액' · '금액' · '예산'이 들어간 열 (예: '2027년 예산액', '예산요구액', '최종예산', 추경은 '기정예산액' + '추경요구액')" & vbCrLf & _
        "이 파일의 머리글: " & YH_HdrList(low, lastC)

    '6) 올해 연도 결정: 머리글 > 제목 행 > 저장된 설정 > 파일 이름(최후 보조) > 직접 입력
    d("ySrc") = ""
    If yMax > 0 Then
        d("yCur") = yMax: d("ySrc") = "머리글"
    Else
        d("yCur") = 0
        For r = 1 To hr - 1
            For c = 1 To lastC
                p = YH_YearWithNen(CleanText(ws.Cells(r, c).Value2))
                If p > d("yCur") Then d("yCur") = p
            Next c
        Next r
        If d("yCur") > 0 Then d("ySrc") = "제목 행"
    End If
    If d("yCur") = 0 And cfgKey <> "" Then
        p = CLng(val(CfgGet(cfgKey)))
        If p >= 2000 And p <= 2099 Then d("yCur") = p: d("ySrc") = "저장된 설정"
    End If
    If d("yCur") = 0 And srcFile <> "" Then
        fn = srcFile
        p = InStrRev(fn, "\"): If p > 0 Then fn = Mid$(fn, p + 1)
        p = YH_YearWithNen(fn)
        If p = 0 And Len(fn) > 5 Then
            If Left$(fn, 4) Like "20##" And Not (Mid$(fn, 5, 1) Like "#") Then p = CLng(Left$(fn, 4))
        End If
        If p > 0 Then d("yCur") = p: d("ySrc") = "파일 이름"
    End If
    If d("yCur") = 0 Then
        T = "세출 파일 머리글에서 연도를 찾지 못했습니다." & vbCrLf & _
            "(예: '2027년 산출내역', '2026년도 예산액' 같은 머리글이 없음)"
        If Not allowAsk Then Err.Raise vbObjectError + 802, "YH_Map", T
        MsgBox T & vbCrLf & vbCrLf & "다음 창에서 이 파일의 올해(최신) 예산 연도를 4자리로 입력하세요." & vbCrLf & _
               "취소하면 가져오기를 중단합니다.", vbExclamation, "연도 확인"
        ans = Trim$(InputBox("이 세출 파일의 올해(최신) 예산 연도 4자리 (예: 2027)", "연도 입력", ""))
        If ans Like "20##" Then
            d("yCur") = CLng(ans): d("ySrc") = "직접 입력"
        Else
            Err.Raise vbObjectError + 802, "YH_Map", T & vbCrLf & "연도를 입력하지 않아 중단했습니다."
        End If
    End If
    If y2 > 0 Then
        d("yPrv") = y2
    ElseIf d("calcP") > 0 Or d("amtP") > 0 Then
        d("yPrv") = d("yCur") - 1             ''전년도' 머리글 = 올해 - 1
    Else
        d("yPrv") = 0                         '작년 자료 없음 (연도 하나만 있음)
    End If

    '7) 나머지 열 (머리글 이름 일치, 띄어쓰기 무시)
    For Each k In Array("lv", "e", "i", "g", "h", "ac", "code", "stat")
        d(k) = 0
    Next k
    For c = 1 To lastC
        Select Case low(c)
            Case "레벨": If d("lv") = 0 Then d("lv") = c
            Case "세세부사업": If d("e") = 0 Then d("e") = c
            Case "세부사업": If d("i") = 0 Then d("i") = c
            Case "정책사업": If d("g") = 0 Then d("g") = c
            Case "단위사업": If d("h") = 0 Then d("h") = c
            Case "담당자": If d("ac") = 0 Then d("ac") = c
            Case "세목코드", "세출세목코드": If d("code") = 0 Then d("code") = c
            Case "원가통계목": If d("stat") = 0 Then d("stat") = c
        End Select
    Next c
    If d("lv") = 0 Then Err.Raise vbObjectError + 805, "YH_Map", YH_Who(ws, srcFile) & "에서 '레벨' 열을 찾지 못했습니다. (머리글 " & hr & "행)" & vbCrLf & "이 파일의 머리글: " & YH_HdrList(low, lastC)
    If d("e") = 0 Then Err.Raise vbObjectError + 805, "YH_Map", YH_Who(ws, srcFile) & "에서 '세세부사업' 열을 찾지 못했습니다. (머리글 " & hr & "행)" & vbCrLf & "이 파일의 머리글: " & YH_HdrList(low, lastC)
    If d("i") = 0 Then Err.Raise vbObjectError + 805, "YH_Map", YH_Who(ws, srcFile) & "에서 '세부사업' 열을 찾지 못했습니다. (머리글 " & hr & "행)" & vbCrLf & "이 파일의 머리글: " & YH_HdrList(low, lastC)

    '8) 금액 단위: 머리글/제목에 '(단위: 원)' 이 있으면 원, 기본은 천원 (UBIS 세출예산요구서)
    d("unit") = 1000
    unitTxt = low(d("amt"))
    For r = 1 To hr
        For c = 1 To lastC
            T = Replace(CleanText(ws.Cells(r, c).Value2), " ", "")
            If InStr(T, "단위") > 0 Then unitTxt = unitTxt & T
        Next c
    Next r
    If InStr(unitTxt, "천원") = 0 And (InStr(unitTxt, "단위:원") > 0 Or InStr(unitTxt, "(원)") > 0) Then d("unit") = 1
    Set YH_Map = d
End Function

'자료 배열(머리글 다음 행부터, 1열부터) 한 행의 올해 금액 (파일 단위 그대로). gd 모드면 기정예산액 + 추경요구액
Public Function YH_AmtAt(arr As Variant, ByVal r As Long, mp As Object) As Double
    If mp("gd") Then
        YH_AmtAt = NzD(arr(r, mp("gij"))) + NzD(arr(r, mp("delta")))
    Else
        YH_AmtAt = NzD(arr(r, mp("amt")))
    End If
End Function

'머리글 행 위(제목·안내 줄) 글자를 모두 이어 붙임 (차수·추경 판단용 - 파일 이름은 쓰지 않음)
Public Function YH_TitleText(ws As Worksheet, ByVal hr As Long, ByVal lastC As Long) As String
    Dim r As Long, c As Long, T As String, s As String
    For r = 1 To hr - 1
        For c = 1 To lastC
            T = CleanText(ws.Cells(r, c).Value2)
            If T <> "" Then s = s & " " & T
        Next c
    Next r
    YH_TitleText = s
End Function

'---------------------------------------------------------------- 산출식 계산 (V8)
'  "50,000원×10명×10회=" → 5,000,000 (원).  "300천원 × 20회" → 6,000,000.  "(5만원+2만원)" 처럼 괄호 안 + 는 더함.
'  '=' 오른쪽에 금액이 적혀 있으면 그 값을 쓴다. 여러 줄이면 줄마다 계산해 더한다. 금액(원/천원/백만원)이 하나도 없으면 -1
Public Function YH_EvalCalcWon(ByVal txt As String) As Double
    Dim lines As Variant, i As Long, v As Double, tot As Double, anyV As Boolean
    txt = Replace(txt, vbCr, "")
    lines = Split(txt, vbLf)
    For i = LBound(lines) To UBound(lines)
        If Trim$(CStr(lines(i))) <> "" Then
            v = EvalLine(CStr(lines(i)))
            If v >= 0 Then tot = tot + v: anyV = True
        End If
    Next i
    If anyV Then YH_EvalCalcWon = tot Else YH_EvalCalcWon = -1
End Function

Private Function Rx(ByVal pat As String) As Object
    Static cache As Object
    If cache Is Nothing Then Set cache = CreateObject("Scripting.Dictionary")
    If Not cache.Exists(pat) Then
        Dim re As Object
        Set re = CreateObject("VBScript.RegExp")
        re.Pattern = pat: re.Global = True
        cache.Add pat, re
    End If
    Set Rx = cache(pat)
End Function

Private Function MoneyPat() As String
    MoneyPat = "([0-9][0-9,]*(\.[0-9]+)?)\s*(백만원|천원|원)"
End Function

Private Function MoneyOf(m As Object) As Double
    Dim n As Double
    n = val(Replace(CStr(m.SubMatches(0)), ",", ""))
    Select Case CStr(m.SubMatches(2))
        Case "백만원": MoneyOf = n * 1000000#
        Case "천원": MoneyOf = n * 1000#
        Case Else: MoneyOf = n
    End Select
End Function

Private Function EvalLine(ByVal s As String) As Double
    Dim p As Long, rhs As String, ms As Object, st As Long, k As Long, expr As String
    Dim parts As Collection, T As Variant, v As Double, tot As Double, ch As String
    s = Trim$(s)
    p = InStrRev(s, "=")
    If p > 0 Then
        rhs = Mid$(s, p + 1)
        Set ms = Rx(MoneyPat()).Execute(rhs)
        If ms.Count > 0 Then EvalLine = MoneyOf(ms(0)): Exit Function
        If Rx("^[\s,0-9]+$").Test(rhs) And Rx("[0-9]").Test(rhs) Then EvalLine = val(Replace(Replace(rhs, ",", ""), " ", "")): Exit Function
        s = Left$(s, p - 1)
    End If
    Set ms = Rx(MoneyPat()).Execute(s)
    If ms.Count = 0 Then EvalLine = -1: Exit Function
    st = ms(0).FirstIndex + 1
    k = st - 1
    Do While k >= 1
        If Mid$(s, k, 1) <> " " Then Exit Do
        k = k - 1
    Loop
    If k >= 1 Then
        ch = Mid$(s, k, 1)
        If ch = "(" Or ch = ChrW(&HFF08) Then st = k
    End If
    expr = Mid$(s, st)
    Set parts = SplitTop(expr, "+")
    For Each T In parts
        If Trim$(CStr(T)) <> "" Then
            v = EvalTerm(CStr(T))
            If v < 0 Then
                Set ms = Rx(MoneyPat()).Execute(CStr(T))
                If ms.Count > 0 Then v = MoneyOf(ms(0)) Else v = 0
            End If
            tot = tot + v
        End If
    Next T
    EvalLine = tot
End Function

'괄호 밖의 구분 글자로만 나누기
Private Function SplitTop(ByVal s As String, ByVal sep As String) As Collection
    Dim col As New Collection, i As Long, ch As String, depth As Long, cur As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch = "(" Or ch = "[" Or ch = ChrW(&HFF08) Then depth = depth + 1
        If ch = ")" Or ch = "]" Or ch = ChrW(&HFF09) Then If depth > 0 Then depth = depth - 1
        If ch = sep And depth = 0 Then
            col.Add cur: cur = ""
        Else
            cur = cur & ch
        End If
    Next i
    col.Add cur
    Set SplitTop = col
End Function

Private Function IsMulChar(ByVal ch As String) As Boolean
    IsMulChar = (ch = "*" Or ch = "x" Or ch = "X" Or ch = "/" Or ch = ChrW(&HD7) Or ch = ChrW(&HF7) Or ch = ChrW(&H2715) Or ch = ChrW(&H2716))
End Function

'곱하기·나누기 항: 첫 칸은 금액(괄호 안 + 는 더함), 나머지 칸은 첫 숫자(% 는 /100)
Private Function EvalTerm(ByVal T As String) As Double
    Dim i As Long, ch As String, seg As String, acc As Double, started As Boolean, op As String, nextOp As String
    Dim ms As Object, m As Object, v As Double
    T = T & "*"
    op = "*"
    For i = 1 To Len(T)
        ch = Mid$(T, i, 1)
        If IsMulChar(ch) Then
            If ch = "/" Or ch = ChrW(&HF7) Then nextOp = "/" Else nextOp = "*"
            If Not started Then
                Set ms = Rx(MoneyPat()).Execute(seg)
                If ms.Count = 0 Then EvalTerm = -1: Exit Function
                If InStr(seg, "+") > 0 Then
                    acc = 0
                    For Each m In ms
                        acc = acc + MoneyOf(m)
                    Next m
                Else
                    acc = MoneyOf(ms(0))
                End If
                started = True
            Else
                Set ms = Rx("([0-9][0-9,]*(\.[0-9]+)?)\s*(%)?").Execute(seg)
                If ms.Count > 0 Then
                    v = val(Replace(CStr(ms(0).SubMatches(0)), ",", ""))
                    If CStr(ms(0).SubMatches(2)) = "%" Then v = v / 100
                    If op = "/" Then
                        If v = 0 Then EvalTerm = -1: Exit Function
                        acc = acc / v
                    Else
                        acc = acc * v
                    End If
                End If
            End If
            op = nextOp: seg = ""
        Else
            seg = seg & ch
        End If
    Next i
    If started Then EvalTerm = acc Else EvalTerm = -1
End Function

'산출식 결과를 금액 단위(unit: 천원 = 1000)로 반올림한 값 (원)
'  주의: 클래스 필드는 ByRef 로 넘기면 복사본만 바뀌므로 반환값으로 더한다
Public Function YH_CalcRound(ByVal won As Double, ByVal unit As Double) As Double
    If unit <= 1 Then YH_CalcRound = won: Exit Function
    YH_CalcRound = Int(won / unit + 0.5) * unit
End Function

'끝수 허용폭: 천원으로 떨어지지 않는 세목 행이면 1단위, 아니면 0 (원 단위 파일은 0 = 정확히 일치)
Public Function YH_CalcTol(ByVal won As Double, ByVal unit As Double) As Double
    Dim q As Double
    If unit <= 1 Then Exit Function
    q = won / unit
    If Abs(q - Int(q + 0.5)) > 0.000001 Then YH_CalcTol = unit
End Function



