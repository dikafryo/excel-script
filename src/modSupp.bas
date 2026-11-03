Attribute VB_Name = "modSupp"
Option Explicit
'=====================================================================
' modSupp : (V8) 추경예산 추가 · 차수별 점검표
'   - 본예산 세출 파일(① 세출예산요구서)을 불러온 뒤, 추경 세출 파일을 몇 번이든 추가 (1차, 2차, 3차 ... 제한 없음)
'   - 파일마다 숨김 시트 '원본_추경N' 에 값 복사, 설정 시트에 차수·파일·내용 지문·연도·점검 결과 기록
'   - 합치는 규칙: 사업항목 키(세부사업|사업항목)가 같으면 나중에 추가한 차수가 최신 상태로 '교체'(더하지 않음),
'                  새 키는 '추가', 추경 파일에 없는 사업은 앞 차수 값을 그대로 유지
'   - 추가할 때마다 점검: 중복 파일(같은 경로 / 같은 내용 지문), 연도 일치(본예산 올해 연도), 차수 순서, 합계(장표 기준)
'   - 금액 기준값 = 장표(102-B) 예산금액. 차이액 = 비교값(세출파일 금액 또는 산출내역 합계) - 장표
'   - 결과: '추경점검' 시트 한 화면에 본예산 · 1차 · 2차 ... · 최종 상태 구역으로 표시
'=====================================================================
Public Const SH_SUPP_PREFIX As String = "원본_추경"
Public Const SH_CHECK As String = "추경점검"
Private Const SH_TMP As String = "원본_추경_임시"
Private Const CK_NC As Long = 10
Public Const NOTE_B As String = "최신년도 산출내역 미반영 가능성"   'V8 비고 종류 (중립 표현)
Public Const NOTE_C As String = "예산현황과 산출내역 금액 확인 필요"
Public gBX_NoUI As Boolean        '자체 점검(BX_SelfTest) 중: 대화상자 없이 실행 (확인 질문은 '예'로 처리)
Public gBX_DupWarn As String      '자체 점검 중 표시하지 않은 중복 경고 (파일명 vbTab 문구, 여러 건은 vbLf+vbLf 로 구분)
Public Const DUP_TITLE As String = "B-XLAM V8 · 중복 파일"

'---------------------------------------------------------------- 등록 정보
Public Function SuppCount() As Long
    SuppCount = CLng(val(CfgGet("추경수")))
End Function

Public Function SuppSheetName(ByVal i As Long) As String
    SuppSheetName = SH_SUPP_PREFIX & i
End Function

Public Function SuppLabel(ByVal i As Long) As String
    If i <= 0 Then SuppLabel = "본예산": Exit Function
    SuppLabel = CfgGet("추경" & i & "라벨")
    If SuppLabel = "" Then SuppLabel = i & "차 추경"
End Function

Public Function SuppFile(ByVal i As Long) As String
    If i <= 0 Then SuppFile = CfgGet("요구서파일") Else SuppFile = CfgGet("추경" & i & "파일")
End Function

Public Function FileNameOf(ByVal fp As String) As String
    Dim p As Long
    p = InStrRev(fp, "\")
    If p > 0 Then FileNameOf = Mid$(fp, p + 1) Else FileNameOf = fp
End Function

'차수 색 (본예산 회색, 1차부터 6색 순환 - 차수 수 제한 없음, 글자 라벨로도 구분)
Public Function ChasuColor(ByVal i As Long) As Long
    Dim pal As Variant
    If i <= 0 Then ChasuColor = RGB(229, 229, 234): Exit Function
    pal = Array(RGB(255, 236, 179), RGB(204, 236, 214), RGB(214, 228, 255), RGB(230, 220, 255), RGB(204, 240, 240), RGB(255, 226, 199))
    ChasuColor = pal((i - 1) Mod (UBound(pal) + 1))
End Function

Public Function ChasuColorDark(ByVal i As Long) As Long
    Dim pal As Variant
    If i <= 0 Then ChasuColorDark = RGB(72, 72, 74): Exit Function
    pal = Array(RGB(176, 112, 0), RGB(30, 120, 55), RGB(0, 84, 190), RGB(104, 64, 190), RGB(0, 120, 120), RGB(190, 90, 0))
    ChasuColorDark = pal((i - 1) Mod (UBound(pal) + 1))
End Function

'확인 필요 행 색 (장표 · 산출내역과 금액이 달라 확인이 필요한 행 - 추경 등으로 정상일 수 있음)
Public Function MismatchColor() As Long
    MismatchColor = RGB(255, 199, 206)
End Function

Public Function IsMismatch(ByVal amt As Double, ByVal jang As Double) As Boolean
    IsMismatch = (Abs(amt - jang) >= 1)
End Function

'---------------------------------------------------------------- 장표 기준값
Public Function JangLoaded() As Boolean
    If gBAgg Is Nothing Then Exit Function
    JangLoaded = (gBAgg.Count > 0)
End Function

'세출 사업항목 키 → 장표 예산금액(원). found = False 면 장표에 짝이 없음
Public Function JangOf(ByVal key As String, ByRef found As Boolean) As Double
    Dim rq As clsProject, a As clsProject
    found = False
    If Not JangLoaded() Then Exit Function
    If gReqItems Is Nothing Then Exit Function
    If Not gReqItems.Exists(key) Then Exit Function
    Set rq = gReqItems(key)
    If rq.장표키 = "" Then Exit Function
    If Not gBAgg.Exists(rq.장표키) Then Exit Function
    Set a = gBAgg(rq.장표키)
    found = True
    JangOf = a.장표예산
End Function

'---------------------------------------------------------------- 내용 지문 / 차수 번호
'세출 파일 내용 지문: 머리글 행부터 끝까지의 값으로 계산 (파일 이름·출력 시각이 달라도 내용이 같으면 같은 지문)
Public Function DataFingerprint(ws As Worksheet) As String
    Dim hr As Long, arr As Variant, r As Long, c As Long, s As String, i As Long, ch As Long
    Dim h1 As Double, h2 As Double, ur As Range, lr As Long, lc As Long
    Const M1 As Double = 2147483647#
    Const M2 As Double = 2147483629#
    hr = YH_FindHeaderRow(ws)
    If hr = 0 Then hr = 1
    Set ur = ws.UsedRange
    lr = ur.row + ur.rows.Count - 1: lc = ur.Column + ur.Columns.Count - 1
    If lr <= hr Or lc < 2 Then DataFingerprint = "empty": Exit Function
    arr = ws.Range(ws.Cells(hr, 1), ws.Cells(lr, lc)).Value2
    h1 = 7: h2 = 11
    For r = 1 To UBound(arr, 1)
        For c = 1 To UBound(arr, 2)
            If IsError(arr(r, c)) Then s = "#" Else s = CStr(arr(r, c))
            For i = 1 To Len(s)
                ch = AscW(Mid$(s, i, 1)) And &HFFFF&
                h1 = h1 * 31 + ch: h1 = h1 - Int(h1 / M1) * M1
                h2 = h2 * 131 + ch: h2 = h2 - Int(h2 / M2) * M2
            Next i
            h1 = h1 * 31 + 9: h1 = h1 - Int(h1 / M1) * M1
            h2 = h2 * 131 + 9: h2 = h2 - Int(h2 / M2) * M2
        Next c
    Next r
    DataFingerprint = (lr - hr + 1) & "x" & lc & "-" & Format$(h1, "0") & "-" & Format$(h2, "0")
End Function

Private Function BaseFingerprint() As String
    BaseFingerprint = CfgGet("요구서지문")
    If BaseFingerprint = "" And SheetExists(SH_REQ) Then
        BaseFingerprint = DataFingerprint(BW().Worksheets(SH_REQ))
        CfgSet "요구서지문", BaseFingerprint
    End If
End Function

'"제1회 추경", "3회추경" (회 우선) 또는 "2차", "추경(2차)" → 차수 번호. 없으면 0
Public Function ChasuNoFromText(ByVal s As String) As Long
    Dim pass As Long, i As Long, j As Long, num As String, nx As String, mark As String
    s = Replace(s, " ", "")
    For pass = 1 To 2
        If pass = 1 Then mark = "회" Else mark = "차"
        i = 1
        Do While i <= Len(s)
            If Mid$(s, i, 1) Like "#" Then
                j = i
                Do While j <= Len(s)
                    If Not (Mid$(s, j, 1) Like "#") Then Exit Do
                    j = j + 1
                Loop
                num = Mid$(s, i, j - i)
                nx = Mid$(s, j, 1)
                If Len(num) <= 2 And nx = mark Then
                    ChasuNoFromText = CLng(num)
                    Exit Function
                End If
                i = j
            Else
                i = i + 1
            End If
        Loop
    Next pass
End Function

'V8: 파일 '내용'에서 추경 차수 찾기 (파일 이름은 쓰지 않음)
'  머리글 행 위 칸(제목 · 묶음 머리글) 중 '추경' / '추가경정' 글자가 있는 칸에서 '제N회' / 'N회' / 'N차'.
'  isSupp = 추경 글자가 있거나 기정예산액 열이 있으면 True. why = 찾은 곳 설명
Public Function ChasuFromContent(ws As Worksheet, mp As Object, ByRef isSupp As Boolean, ByRef why As String) As Long
    Dim r As Long, c As Long, T As String, n As Long
    isSupp = (mp("gij") > 0)
    why = ""
    For r = 1 To mp("hr") - 1
        For c = 1 To mp("lastC")
            T = CleanText(ws.Cells(r, c).Value2)
            If InStr(T, "추경") > 0 Or InStr(T, "추가경정") > 0 Then
                isSupp = True
                n = ChasuNoFromText(T)
                If n > 0 Then
                    why = IIf(r <= 3, "제목 행 ", "머리글 ") & ws.Cells(r, c).Address(False, False) & " '" & IIf(InStr(Replace(T, " ", ""), "제" & n & "회") > 0, "제", "") & n & IIf(InStr(Replace(T, " ", ""), n & "회") > 0, "회", "차") & " 추경'"
                    ChasuFromContent = n
                    Exit Function
                End If
            End If
        Next c
    Next r
End Function

'V8: 세출 시트의 사업항목(1레벨) 금액 합계 (원). gijSum = 기정예산액 합계 (열이 없으면 -1)
Public Sub FileTotals(ws As Worksheet, mp As Object, ByRef gijSum As Double, ByRef finSum As Double)
    Dim hr As Long, lr As Long, arr As Variant, r As Long, lv As String, e As String, p As Long, unit As Double
    hr = mp("hr"): unit = mp("unit")
    gijSum = IIf(mp("gij") > 0, 0, -1): finSum = 0
    lr = Application.WorksheetFunction.Max(lastRow(ws, mp("e")), lastRow(ws, mp("amt")))
    If mp("gd") Then lr = Application.WorksheetFunction.Max(lr, lastRow(ws, mp("delta")))
    If lr <= hr Then Exit Sub
    arr = ws.Range(ws.Cells(hr + 1, 1), ws.Cells(lr, mp("lastC"))).Value2
    For r = 1 To UBound(arr, 1)
        lv = CleanText(arr(r, mp("lv"))): e = CleanText(arr(r, mp("e")))
        If lv = "" And e <> "" Then
            p = InStr(e, ".")
            If p >= 2 And p <= 4 Then
                If IsNumeric(Left$(e, p - 1)) Then
                    finSum = finSum + YH_AmtAt(arr, r, mp) * unit
                    If mp("gij") > 0 Then gijSum = gijSum + NzD(arr(r, mp("gij"))) * unit
                End If
            End If
        End If
    Next r
End Sub

Private Function BaseTotal() As Double
    Dim mp As Object, g As Double, f As Double
    If CfgGet("요구서최종합") = "" And SheetExists(SH_REQ) Then
        Set mp = YH_Map(BW().Worksheets(SH_REQ), "", "올해연도", False)
        FileTotals BW().Worksheets(SH_REQ), mp, g, f
        CfgSet "요구서최종합", Format$(f, "0")
    End If
    BaseTotal = val(CfgGet("요구서최종합"))
End Function

Private Function MaxChasuNo() As Long
    Dim i As Long, v As Long
    For i = 1 To SuppCount()
        v = CLng(val(CfgGet("추경" & i & "차수")))
        If v > MaxChasuNo Then MaxChasuNo = v
    Next i
End Function

Private Sub LogAttempt(ByVal fn As String, ByVal why As String)
    Dim m As Long
    m = CLng(val(CfgGet("추경기록수"))) + 1
    CfgSet "추경기록수", CStr(m)
    CfgSet "추경기록" & m, NowText() & " | " & fn & " | " & why
End Sub

'---------------------------------------------------------------- V8 저장 순서 = 차수 순서 (추가 순서와 무관)
Private Function SuppKeys() As Variant
    SuppKeys = Array("파일", "지문", "연도", "차수", "라벨", "시각", "중복점검", "연도점검", "순서점검", "추가순번", "차수근거", "기정합", "최종합")
End Function

Private Function AddSeqOf(ByVal i As Long) As Long
    AddSeqOf = CLng(val(CfgGet("추경" & i & "추가순번")))
    If AddSeqOf = 0 Then AddSeqOf = i          'V8 이전 보드: 저장 순서 = 추가 순서
End Function

'저장 칸 1..n 을 (차수 번호, 추가 순번) 순서로 다시 배치 (원본_추경N 시트 + 설정 값). 이후 계산·표시는 1..n 순서 그대로
Private Sub ReorderSupp()
    Dim n As Long, i As Long, j As Long, ord() As Long, T As Long, ks As Variant, vals() As String, same As Boolean
    n = SuppCount()
    If n < 2 Then Exit Sub
    ReDim ord(1 To n)
    For i = 1 To n: ord(i) = i: Next i
    For i = 2 To n
        T = ord(i): j = i - 1
        Do While j >= 1
            If SortVal(ord(j)) <= SortVal(T) Then Exit Do
            ord(j + 1) = ord(j): j = j - 1
        Loop
        ord(j + 1) = T
    Next i
    same = True
    For i = 1 To n
        If ord(i) <> i Then same = False
    Next i
    If same Then Exit Sub
    ks = SuppKeys()
    ReDim vals(1 To n, 0 To UBound(ks))
    For i = 1 To n
        For j = 0 To UBound(ks)
            vals(i, j) = CfgGet("추경" & i & ks(j))
        Next j
        If SheetExists(SuppSheetName(i)) Then BW().Worksheets(SuppSheetName(i)).name = SH_SUPP_PREFIX & "정렬_" & i
    Next i
    For i = 1 To n
        If SheetExists(SH_SUPP_PREFIX & "정렬_" & ord(i)) Then BW().Worksheets(SH_SUPP_PREFIX & "정렬_" & ord(i)).name = SuppSheetName(i)
        For j = 0 To UBound(ks)
            CfgSet "추경" & i & ks(j), vals(ord(i), j)
        Next j
    Next i
    '시트 탭 위치도 차수 순서로 (숨김 시트라 실패해도 계산에는 영향 없음)
    On Error Resume Next
    For i = 2 To n
        If SheetExists(SuppSheetName(i)) And SheetExists(SuppSheetName(i - 1)) Then
            BW().Worksheets(SuppSheetName(i)).Move After:=BW().Worksheets(SuppSheetName(i - 1))
        End If
    Next i
    On Error GoTo 0
End Sub

'추가할 때마다 다시 보기: 앞 차수가 나중에 들어온 경우 뒤 차수 파일의 '앞 차수 없음' 안내를 지금 상태로 갱신
Private Sub RefreshSuppNotes()
    Dim n As Long, i As Long, j As Long, no As Long, W As String, g As String, k As Long, kk As Long, has As Boolean, allOk As Boolean
    n = SuppCount()
    For i = 1 To n
        no = CLng(val(CfgGet("추경" & i & "차수")))
        W = CfgGet("추경" & i & "차수근거")
        g = CfgGet("추경" & i & "기정합")
        If InStr(W, "앞 차수 파일 없음") > 0 And g <> "" Then
            For j = 1 To n
                If j <> i And CfgGet("추경" & j & "최종합") <> "" Then
                    If Abs(val(g) - val(CfgGet("추경" & j & "최종합"))) < 1 Then
                        k = CLng(val(CfgGet("추경" & j & "차수"))) + 1
                        W = Replace(W, "기정액 연결: 같은 합계의 앞 차수 파일 없음(참고)", "기정액 합계 = " & SuppLabel(j) & " 최종 합계 → " & k & "차" & _
                            IIf(k = no, " (일치 · 앞 차수 추가 후 확인)", " (제목과 다름 - 참고)"))
                        CfgSet "추경" & i & "차수근거", W
                        Exit For
                    End If
                End If
            Next j
        End If
        W = CfgGet("추경" & i & "순서점검")
        If InStr(W, "확인: 앞 차수") = 1 Then
            allOk = True
            For kk = 1 To no - 1
                has = False
                For j = 1 To n
                    If CLng(val(CfgGet("추경" & j & "차수"))) = kk Then has = True
                Next j
                If Not has Then allOk = False
            Next kk
            If allOk Then CfgSet "추경" & i & "순서점검", "통과: 앞 차수가 나중에 추가되어 차수 순서로 정렬됨 (" & (no - 1) & "차 다음 " & no & "차)"
        End If
    Next i
End Sub

Private Function SortVal(ByVal i As Long) As Double
    SortVal = CDbl(val(CfgGet("추경" & i & "차수"))) * 100000# + AddSeqOf(i)
End Function

'저장 칸 j 를 빼고 뒤 칸을 한 칸씩 앞으로
Private Sub RemoveSuppAt(ByVal j As Long)
    Dim n As Long, i As Long, k As Variant
    n = SuppCount()
    If j < 1 Or j > n Then Exit Sub
    DeleteSheetIfExists SuppSheetName(j)
    For i = j + 1 To n
        If SheetExists(SuppSheetName(i)) Then BW().Worksheets(SuppSheetName(i)).name = SuppSheetName(i - 1)
        For Each k In SuppKeys()
            CfgSet "추경" & (i - 1) & k, CfgGet("추경" & i & k)
        Next k
    Next i
    For Each k In SuppKeys()
        CfgSet "추경" & n & k, ""
    Next k
    CfgSet "추경수", CStr(n - 1)
End Sub

'본예산 연도·지문이 기록돼 있는지 확인 (V7 에서 만든 보드 호환)
Private Sub EnsureBaseRecorded()
    Dim mp As Object
    If CfgGet("올해연도") = "" Then
        Set mp = YH_Map(BW().Worksheets(SH_REQ), CfgGet("요구서파일"), "", Not gBX_NoUI)
        CfgSet "올해연도", CStr(mp("yCur")): CfgSet "작년연도", CStr(mp("yPrv")): CfgSet "연도출처", CStr(mp("ySrc"))
    End If
    BaseFingerprint
    BaseTotal
End Sub

'---------------------------------------------------------------- 버튼: 추경예산 추가 (여러 번 · 여러 파일 가능)
Public Sub 추경예산추가()
    Dim fps As Variant, i As Long, summary As String, nBefore As Long, emsg As String
    If Not SheetExists(SH_REQ) Then
        MsgBox "먼저 [① 세출예산요구서]로 본예산 세출 파일을 불러오세요." & vbCrLf & _
               "추경 파일은 파일 내용(제목 · 기정액)으로 차수를 판단해 차수 순서로 정렬합니다.", vbExclamation, "추경예산 추가"
        Exit Sub
    End If
    fps = Application.GetOpenFilename("Excel 파일 (*.xls*), *.xls*", , "추경 세출 파일 선택 (여러 개 선택 가능 - 차수는 파일 내용으로 판단)", , True)
    If VarType(fps) = vbBoolean Then Exit Sub
    nBefore = SuppCount()
    On Error GoTo EH
    EnsureBaseRecorded
    For i = LBound(fps) To UBound(fps)
        Application.StatusBar = "추경 파일 점검 중... (" & (i - LBound(fps) + 1) & "/" & (UBound(fps) - LBound(fps) + 1) & ") " & FileNameOf(CStr(fps(i)))
        AddOneSupp CStr(fps(i)), summary
    Next i
    Application.StatusBar = False
    AutoRefresh                       '모델 · 보드 · 세부산출내역 다시 만들기 (추경 반영)
    SpeedOn "추경 점검표 만드는 중..."
    BuildCheckSheet
    SpeedOff
    BW().Worksheets(SH_CHECK).Activate
    MsgBox "추경예산 추가 결과: 이번에 " & (SuppCount() - nBefore) & "건 반영 (현재 본예산 + 추경 " & SuppCount() & "개 차수, 차수 순서로 정렬)" & vbCrLf & summary & vbCrLf & vbCrLf & _
           "점검 결과(중복 · 연도 · 차수 판단 · 장표 · 산출내역)는 '" & SH_CHECK & "' 시트에 차수별로 정리했습니다.", vbInformation, "추경예산 추가"
    Exit Sub
EH:
    emsg = Err.Description
    Resume CleanFail
CleanFail:
    On Error Resume Next
    Application.StatusBar = False
    SpeedOff
    DeleteSheetIfExists SH_TMP
    If SuppCount() <> nBefore Or summary <> "" Then
        AutoRefresh
        BuildCheckSheet
    End If
    MsgBox "추경예산 추가 중 멈췄습니다: " & emsg & IIf(summary <> "", vbCrLf & vbCrLf & "처리된 파일:" & summary, ""), vbExclamation, "추경예산 추가"
End Sub

'파일 하나 점검 후 등록. summary 에 결과 한 줄 추가
'  차수 = 파일 내용(제목 · 머리글의 '제N회 추경') > 기정액 연결(이 파일 기정액 합계 = 다른 파일 최종 합계 → 그 다음 차수) > 묻기(제안값)
'  파일 이름은 차수 · 본예산/추경 판단에 쓰지 않는다. 등록 후 저장 칸을 차수 순서로 다시 정렬.
Private Sub AddOneSupp(ByVal fp As String, ByRef summary As String)
    Dim n As Long, j As Long, dupWhy As String, wbS As Workbook, wsS As Worksheet, wsT As Worksheet
    Dim fpr As String, mp As Object, y As Long, baseY As Long, yChk As String, oChk As String
    Dim no As Long, lastNo As Long, lbl As String, why As String, chain As String, chainNo As Long
    Dim fn As String, eN As Long, eD As String, isSupp As Boolean, gSum As Double, fsum As Double
    Dim prop As Long, ans As String, seq As Long, exJ As Long, cWhy As String, maxLbl As String, dupLbl As String

    On Error GoTo EH
    n = SuppCount()
    fn = FileNameOf(fp)
    '1) 중복: 같은 경로
    If StrComp(fp, CfgGet("요구서파일"), vbTextCompare) = 0 Then dupWhy = "본예산 파일과 같은 파일(경로 동일)": dupLbl = "본예산"
    For j = 1 To n
        If dupWhy = "" And StrComp(fp, SuppFile(j), vbTextCompare) = 0 Then dupWhy = SuppLabel(j) & " 파일과 같은 파일(경로 동일)": dupLbl = SuppLabel(j)
    Next j
    '2) 임시 시트로 값 복사
    DeleteSheetIfExists SH_TMP
    Set wsT = GetSheet(SH_TMP, True)
    Set wbS = Workbooks.Open(fileName:=fp, ReadOnly:=True, UpdateLinks:=0)
    Set wsS = wbS.Worksheets(1)
    CopyValues wsS, wsT, "", ""
    YH_FillMerged wsS, wsT
    wbS.Close SaveChanges:=False
    Set wbS = Nothing
    wsT.Visible = xlSheetHidden
    '3) 중복: 같은 내용 (지문)
    fpr = DataFingerprint(wsT)
    If dupWhy = "" And fpr = BaseFingerprint() Then dupWhy = "본예산 파일과 내용이 같음": dupLbl = "본예산"
    For j = 1 To n
        If dupWhy = "" And fpr = CfgGet("추경" & j & "지문") Then dupWhy = SuppLabel(j) & " 파일과 내용이 같음": dupLbl = SuppLabel(j)
    Next j
    If dupWhy <> "" Then
        DeleteSheetIfExists SH_TMP
        LogAttempt fn, "중복 - 추가하지 않음: " & dupWhy
        DupWarn fn, dupLbl                '경고 창 (자체 점검 중에는 표시하지 않고 문구만 기록)
        summary = summary & vbCrLf & "X " & fn & " : 중복 파일이라 추가하지 않음 (" & dupWhy & ")"
        Exit Sub
    End If
    '4) 연도 일치 (머리글 > 제목 > 저장된 설정 > 파일 이름 > 직접 입력)
    Set mp = YH_Map(wsT, fp, "", Not gBX_NoUI)
    y = mp("yCur")
    baseY = CLng(val(CfgGet("올해연도")))
    If baseY = 0 Then
        yChk = "확인 불가: 본예산 연도 기록 없음 (이 파일 " & y & "년)"
    ElseIf y = baseY Then
        yChk = "통과: " & y & "년 = 본예산 " & baseY & "년 (" & mp("ySrc") & "에서 인식)"
    Else
        If AskYN("연도가 다릅니다." & vbCrLf & vbCrLf & "본예산: " & baseY & "년" & vbCrLf & "추가하려는 파일(" & fn & "): " & y & "년" & vbCrLf & vbCrLf & _
                  "다른 해의 파일일 수 있습니다. 그래도 추가할까요?", vbYesNo + vbExclamation + vbDefaultButton2, "추경예산 추가 - 연도 확인") = vbNo Then
            DeleteSheetIfExists SH_TMP
            LogAttempt fn, "건너뜀 - 연도 다름(" & y & "년, 본예산 " & baseY & "년)"
            summary = summary & vbCrLf & "X " & fn & " : 연도가 달라 추가하지 않음 (" & y & "년)"
            Exit Sub
        End If
        yChk = "경고: " & y & "년 ≠ 본예산 " & baseY & "년 (사용자 확인 후 추가)"
    End If
    '5) 차수 판단 (내용만 사용)
    FileTotals wsT, mp, gSum, fsum
    lastNo = MaxChasuNo()
    no = ChasuFromContent(wsT, mp, isSupp, why)
    '기정액 연결: 이 파일 기정액 합계 = 본예산 / 다른 추경 파일 최종 합계
    chainNo = -1
    If gSum >= 0 Then
        If Abs(gSum - BaseTotal()) < 1 Then chainNo = 1: chain = "기정액 합계 = 본예산 합계 → 1차"
        For j = 1 To n
            If chainNo < 0 And CfgGet("추경" & j & "최종합") <> "" Then
                If Abs(gSum - val(CfgGet("추경" & j & "최종합"))) < 1 Then
                    chainNo = CLng(val(CfgGet("추경" & j & "차수"))) + 1
                    chain = "기정액 합계 = " & SuppLabel(j) & " 최종 합계 → " & chainNo & "차"
                End If
            End If
        Next j
    End If
    If no > 0 Then
        cWhy = "통과: 파일 내용 " & why
        If chainNo > 0 Then
            cWhy = cWhy & " · " & chain & IIf(chainNo = no, " (일치)", " (제목과 다름 - 참고)")
        ElseIf gSum >= 0 Then
            cWhy = cWhy & " · 기정액 연결: 같은 합계의 앞 차수 파일 없음(참고)"
        End If
    ElseIf chainNo > 0 Then
        no = chainNo
        cWhy = "통과: " & chain & " (제목에 차수 없음)"
    Else
        prop = lastNo + 1
        If gBX_NoUI Then
            no = prop
            cWhy = "제안: 내용에서 차수를 찾지 못해 제안값 " & prop & "차 사용 (자체 점검)"
        Else
            ans = Trim$(InputBox("이 파일의 추경 차수를 고르세요 (제안: " & prop & "차)" & vbCrLf & vbCrLf & _
                  "파일 내용(제목 · 머리글 · 기정액)에서 차수를 찾지 못했습니다." & vbCrLf & "파일: " & fn & vbCrLf & _
                  "숫자만 입력 (예: " & prop & ")  /  취소하면 추가하지 않습니다.", "추경 차수 확인", CStr(prop)))
            ans = Replace(Replace(ans, "차", ""), " ", "")
            If ans = "" Or Not IsNumeric(ans) Then
                DeleteSheetIfExists SH_TMP
                LogAttempt fn, "건너뜀 - 차수를 고르지 않음"
                summary = summary & vbCrLf & "X " & fn & " : 차수를 고르지 않아 추가하지 않음"
                Exit Sub
            End If
            no = CLng(ans)
            If no < 1 Then no = prop
            cWhy = "제안: 내용에서 차수를 찾지 못해 사용자가 " & no & "차 선택 (제안 " & prop & "차)"
        End If
    End If
    If Not isSupp Then cWhy = cWhy & " · 참고: 파일 내용에 추경 문구가 없음"
    lbl = no & "차 추경"
    '같은 차수가 이미 있으면: 교체할지 확인 (기본 = 교체). 자체 점검 중에는 기존 파일 유지
    exJ = 0
    For j = 1 To n
        If CLng(val(CfgGet("추경" & j & "차수"))) = no Then exJ = j
    Next j
    If exJ > 0 Then
        If gBX_NoUI Then
            DeleteSheetIfExists SH_TMP
            LogAttempt fn, "같은 차수 - 기존 " & lbl & " 유지 (자체 점검: 교체하지 않음)"
            summary = summary & vbCrLf & "X " & fn & " : 이미 " & lbl & "이 있어 기존 파일 유지"
            Exit Sub
        End If
        If MsgBox("이미 " & lbl & "이 있습니다." & vbCrLf & vbCrLf & "기존: " & FileNameOf(SuppFile(exJ)) & vbCrLf & "새 파일: " & fn & vbCrLf & vbCrLf & _
                  "새 파일로 교체할까요?  (예 = 교체 / 아니요 = 추가하지 않음)", vbYesNo + vbQuestion + vbDefaultButton1, "추경예산 추가 - 같은 차수") = vbNo Then
            DeleteSheetIfExists SH_TMP
            LogAttempt fn, "같은 차수 - 추가하지 않음 (기존 " & lbl & " 유지)"
            summary = summary & vbCrLf & "X " & fn & " : 기존 " & lbl & " 유지"
            Exit Sub
        End If
        LogAttempt FileNameOf(SuppFile(exJ)), "교체됨 - 같은 차수(" & lbl & ")의 새 파일로 바꿈: " & fn
        RemoveSuppAt exJ
        n = SuppCount()
        lastNo = MaxChasuNo()
    End If
    '차수 순서 안내 (중립)
    If no < lastNo Then
        oChk = "재정렬: 추가 순서와 다르게 차수 순서로 재정렬했습니다 (" & lastNo & "차 → " & no & "차 순으로 추가됨)"
    ElseIf no > lastNo + 1 Then
        oChk = "확인: 앞 차수(" & (lastNo + 1) & "차" & IIf(no - lastNo > 2, "~" & (no - 1) & "차", "") & ")가 아직 없음 - 나중에 추가하면 차수 순서로 정렬"
    Else
        oChk = "통과: 차수 순서대로 추가됨 (" & IIf(lastNo = 0, "본예산", lastNo & "차 추경") & " 다음 " & no & "차)"
    End If
    '6) 등록 (맨 뒤 칸에 넣은 뒤 차수 순서로 재배치)
    seq = CLng(val(CfgGet("추경추가순번"))) + 1
    CfgSet "추경추가순번", CStr(seq)
    n = n + 1
    DeleteSheetIfExists SuppSheetName(n)
    wsT.name = SuppSheetName(n)
    wsT.Visible = xlSheetHidden
    CfgSet "추경" & n & "파일", fp
    CfgSet "추경" & n & "지문", fpr
    CfgSet "추경" & n & "연도", CStr(y)
    CfgSet "추경" & n & "차수", CStr(no)
    CfgSet "추경" & n & "라벨", lbl
    CfgSet "추경" & n & "시각", NowText()
    CfgSet "추경" & n & "중복점검", "통과: 같은 경로·같은 내용의 파일이 없음"
    CfgSet "추경" & n & "연도점검", yChk
    CfgSet "추경" & n & "순서점검", oChk
    CfgSet "추경" & n & "추가순번", CStr(seq)
    CfgSet "추경" & n & "차수근거", cWhy
    CfgSet "추경" & n & "기정합", IIf(gSum >= 0, Format$(gSum, "0"), "")
    CfgSet "추경" & n & "최종합", Format$(fsum, "0")
    CfgSet "추경수", CStr(n)
    ReorderSupp
    RefreshSuppNotes
    summary = summary & vbCrLf & "O " & fn & " → " & lbl & " (" & Mid$(cWhy, InStr(cWhy, ":") + 2) & ")" & _
              IIf(Left$(oChk, 3) = "재정렬", " · 차수 순서로 재정렬", "") & IIf(Left$(yChk, 2) = "경고", " (연도 경고 - 점검표 확인)", "")
    Exit Sub
EH:
    eN = Err.Number: eD = Err.Description
    On Error Resume Next
    If Not wbS Is Nothing Then wbS.Close SaveChanges:=False
    DeleteSheetIfExists SH_TMP
    On Error GoTo 0
    Err.Raise eN, "AddOneSupp", fn & " : " & eD
End Sub

'---------------------------------------------------------------- 버튼: 마지막 추경 취소 (가장 최근에 '추가한' 파일) / 전체 지우기
Public Sub 마지막추경취소()
    Dim n As Long, j As Long, i As Long, fn As String, lbl As String
    n = SuppCount()
    If n = 0 Then
        MsgBox "추가된 추경 파일이 없습니다.", vbInformation, "마지막 추경 취소"
        Exit Sub
    End If
    j = 1
    For i = 2 To n
        If AddSeqOf(i) > AddSeqOf(j) Then j = i
    Next i
    fn = FileNameOf(SuppFile(j)): lbl = SuppLabel(j)
    If MsgBox("가장 최근에 추가한 파일을 뺄까요?" & vbCrLf & vbCrLf & fn & "  →  " & lbl & vbCrLf & vbCrLf & _
              "빼고 나면 남은 차수로 다시 계산합니다.", vbYesNo + vbQuestion + vbDefaultButton2, "마지막 추경 취소") = vbNo Then Exit Sub
    RemoveSuppAt j
    LogAttempt fn, "취소 - 가장 최근에 추가한 " & lbl & " 을(를) 뺌"
    AutoRefresh
    SpeedOn "추경 점검표 만드는 중..."
    BuildCheckSheet
    SpeedOff
    BW().Worksheets(SH_CHECK).Activate
    MsgBox lbl & " (" & fn & ") 을(를) 뺐습니다. 현재 본예산 + 추경 " & SuppCount() & "개 차수.", vbInformation, "마지막 추경 취소"
End Sub

'새 본예산을 불러올 때: 기존 추경 시트·기록 지우기
Public Sub ClearAllSupp()
    Dim n As Long, i As Long, k As Variant
    n = SuppCount()
    For i = 1 To n
        DeleteSheetIfExists SuppSheetName(i)
        For Each k In SuppKeys()
            CfgSet "추경" & i & k, ""
        Next k
    Next i
    If n > 0 Then CfgSet "추경수", "0"
    CfgSet "추경추가순번", "0"
    DeleteSheetIfExists SH_CHECK
End Sub

'---------------------------------------------------------------- 버튼: 추경 점검표 보기
Public Sub 추경점검표()
    On Error GoTo EH
    If Not SheetExists(SH_REQ) Then
        MsgBox "먼저 [① 세출예산요구서]로 본예산 세출 파일을 불러오세요.", vbExclamation, "추경 점검표"
        Exit Sub
    End If
    SpeedOn "추경 점검표 만드는 중..."
    If Not ModelReady Or gChasu Is Nothing Then BuildModel
    BuildCheckSheet
    SpeedOff
    BW().Worksheets(SH_CHECK).Activate
    BW().Worksheets(SH_CHECK).Range("A1").Select
    Exit Sub
EH:
    SpeedOff
    MsgBox "추경 점검표 오류: " & Err.Description, vbCritical, "추경 점검표"
End Sub

'---------------------------------------------------------------- V8 비고 (확인 필요 사유 - 중립 표현)

'장표 대조 비고. cat 비트 1 = 장표 대조
Public Function JangNote(ByVal amt As Double, ByVal calcSum As Double, ByVal found As Boolean, ByVal jang As Double, _
                         ByVal lbl As String, ByRef cat As Long) As String
    If found Then
        If IsMismatch(amt, jang) Or (calcSum >= 0 And IsMismatch(calcSum, jang)) Then
            JangNote = "장표와 " & Replace(IIf(lbl = "", "본예산", lbl), " ", "") & " 금액 확인 필요": cat = cat Or 1
        End If
    ElseIf amt <> 0 Then
        JangNote = "장표에서 찾지 못함(사업명 확인 필요)": cat = cat Or 1
    End If
End Function

'산출내역 비고: 예산현황(세출 파일 사업 금액) ≠ 최신 산출내역 계산 합이면 다른 산출내역(본예산: 작년 / 추경: 기정)과 비교
'  비교 규칙: 원 단위로 바꾼 뒤 비교. 천원 단위 파일은 세목 행마다 산출식 결과를 천원으로 반올림해 더하고,
'            천원으로 떨어지지 않는 세목 행마다 1천원까지 끝수 차이를 허용 (원 단위 파일은 정확히 일치)
'  cat 비트 2 = 최신년도 산출내역 미반영 가능성, 4 = 예산현황-산출내역 확인
Public Function CalcNote(ByVal p As clsProject, ByRef cat As Long) As String
    If p.산출식상태 <> 0 Then Exit Function
    If Abs(p.요구액 - p.산출식합) <= p.산출식허용 + 0.5 Then Exit Function
    If p.산출식전상태 = 0 Then
        If Abs(p.요구액 - p.산출식전합) <= p.산출식전허용 + 0.5 Then
            CalcNote = NOTE_B & "(" & IIf(p.산출식전종류 = "기정", "기정", "작년") & " 산출내역과 일치)"
            cat = cat Or 2
            Exit Function
        End If
    End If
    CalcNote = NOTE_C: cat = cat Or 4
End Function

'최종 상태(본예산 + 추경 반영) 확인 필요 건수 - 추경점검 최종 구역 · 세부산출내역 · 메인 보드가 같은 숫자를 쓰도록 한 곳에서 계산
'  a = 장표 대조(장표에만 있는 사업 포함), b = 최신년도 산출내역 미반영 가능성, c = 예산현황-산출내역 확인. 반환 = 행(사업) 수
Public Function FinalNoteCounts(ByRef a As Long, ByRef b As Long, ByRef c As Long) As Long
    Dim k As Variant, p As clsProject, cat As Long, jang As Double, found As Boolean, n As Long, nj As Long
    a = 0: b = 0: c = 0
    If gReqItems Is Nothing Then Exit Function
    For Each k In gReqItems.keys
        Set p = gReqItems(k)
        cat = 0
        If JangLoaded() Then
            jang = JangOf(CStr(k), found)
            JangNote p.요구액, p.산출합, found, jang, "", cat
        End If
        CalcNote p, cat
        If cat <> 0 Then
            n = n + 1
            If cat And 1 Then a = a + 1
            If cat And 2 Then b = b + 1
            If cat And 4 Then c = c + 1
        End If
    Next k
    nj = JangOnlyCount()
    FinalNoteCounts = n + nj: a = a + nj
End Function

'장표에만 있고 최종 세출 사업과 짝이 없는 사업 수 (장표 예산 0 제외)
Public Function JangOnlyCount() As Long
    Dim used As Object, k As Variant, bk As Variant, n As Long
    If gReqItems Is Nothing Or gBAgg Is Nothing Then Exit Function
    If Not JangLoaded() Then Exit Function
    Set used = CreateObject("Scripting.Dictionary")
    For Each k In gReqItems.keys
        If gReqItems(k).장표키 <> "" Then used(gReqItems(k).장표키) = 1
    Next k
    For Each bk In gBAgg.keys
        If Not used.Exists(bk) Then
            If gBAgg(bk).장표예산 <> 0 Then n = n + 1
        End If
    Next bk
    JangOnlyCount = n
End Function

Public Function JoinNote(ByVal a As String, ByVal b As String) As String
    If a = "" Then
        JoinNote = b
    ElseIf b = "" Then
        JoinNote = a
    Else
        JoinNote = a & " / " & b
    End If
End Function

'맨 위 요약 글: 확인 필요 n건 (장표 대조 a건 · 최신년도 산출내역 미반영 가능성 b건 · 예산현황-산출내역 확인 c건)
Public Function NoteSummary(ByVal nRows As Long, ByVal a As Long, ByVal b As Long, ByVal c As Long, ByVal jl As Boolean) As String
    NoteSummary = "확인 필요 " & nRows & "건 (" & IIf(jl, "장표 대조 " & a & "건", "장표 대조 - (장표 없음)") & " · " & NOTE_B & " " & b & "건 · 예산현황-산출내역 확인 " & c & "건)"
End Function

Public Function CalcRuleText() As String
    CalcRuleText = "산출내역 비교: 원 단위로 맞춘 뒤 비교, 천원 단위 파일은 세목 행마다 천원 반올림 · 천원으로 떨어지지 않는 행마다 1천원 끝수 허용"
End Function

'---------------------------------------------------------------- 점검표 시트 (한 화면, 차수별 구역)
Public Sub BuildCheckSheet()
    Dim ws As Worksheet, r As Long, s As Long, nS As Long, nBad As Long, brk As String, hist As String
    Dim prevKeys As Object, k As Variant, d As Object, m As Long, j As Long, parts As Variant, nFinal As Long
    Dim cA As Long, cb As Long, cC As Long, fA As Long, fB As Long, fC As Long, jl As Boolean, sumTxt As String
    If gChasu Is Nothing Then BuildModel
    If gChasu Is Nothing Then Exit Sub
    jl = JangLoaded()
    DeleteSheetIfExists SH_CHECK
    Set ws = GetSheet(SH_CHECK)
    ws.Tab.color = ClrOrange()
    ws.Cells.Font.name = FONT_NM: ws.Cells.Font.Size = 10
    ws.Cells.Interior.color = RGB(255, 255, 255)
    ws.Columns(1).ColumnWidth = 5: ws.Columns(2).ColumnWidth = 18: ws.Columns(3).ColumnWidth = 30
    ws.Columns(4).ColumnWidth = 10
    ws.Range(ws.Columns(5), ws.Columns(9)).ColumnWidth = 15
    ws.Columns(10).ColumnWidth = 46
    ws.rows(1).RowHeight = 4: ws.rows("2:3").RowHeight = 14: ws.rows(4).RowHeight = 6
    AddHomeButton ws, 4, 4
    AddButton ws, "btnSuppAdd", "추경예산 추가", "추경예산추가", 68, 4, 90, 20, ClrOrange()
    AddButton ws, "btnSuppUndo", "마지막 추경 취소", "마지막추경취소", 162, 4, 96, 20, ClrGray()
    AddButton ws, "btnRefresh", "새로고침", "새로고침버튼", 262, 4, 64, 20, ClrGreen()

    ws.Range("A5").Value2 = IIf(CfgGet("부서명") <> "", CfgGet("부서명") & " ", "") & "추경예산 점검표"
    ws.Range("A5").Font.Size = 20: ws.Range("A5").Font.Bold = True: ws.Range("A5").Font.color = ClrNavy()
    ws.rows(5).RowHeight = 34
    ws.Range("A6").Value2 = "올해 " & CfgGet("올해연도") & "년 · 작년 " & IIf(val(CfgGet("작년연도")) > 0, CfgGet("작년연도") & "년", "자료 없음") & _
        " (연도 인식: " & CfgGet("연도출처") & ")   |   장표 기준값 = 102-B 예산금액" & IIf(CfgGet("기준일") <> "", " (기준일 " & CfgGet("기준일") & ")", "") & _
        "   |   차이액 = 비교값 - 장표   |   단위: 원   |   차수 = 파일 내용으로 판단해 차수 순서로 정렬   |   " & CalcRuleText() & "   |   갱신 " & NowText()
    ws.Range("A6").Font.Size = 9: ws.Range("A6").Font.color = ClrGray()

    r = 9
    nS = SuppCount()
    Set prevKeys = CreateObject("Scripting.Dictionary")
    For s = 0 To nS
        If gChasu.Exists(s) Then
            Set d = gChasu(s)
            cA = 0: cb = 0: cC = 0
            nBad = WriteSection(ws, r, s, d, prevKeys, False, cA, cb, cC)
            brk = brk & IIf(brk = "", "", " · ") & SuppLabel(s) & " " & nBad & "건"
            hist = hist & IIf(hist = "", "", " · ") & Replace(SuppLabel(s), " 추경", "") & " " & nBad
            For Each k In d.keys
                prevKeys(k) = 1
            Next k
        End If
    Next s
    If Not gReqItems Is Nothing Then
        nFinal = WriteSection(ws, r, -1, gReqItems, prevKeys, True, fA, fB, fC)   '맨 위 숫자 = 최종 상태만 (앞 차수 구역은 이력)
    End If

    '추가하지 않은 파일 기록
    m = CLng(val(CfgGet("추경기록수")))
    If m > 0 Then
        With ws.Range(ws.Cells(r, 1), ws.Cells(r, CK_NC))
            .Merge
            .Value2 = "■ 추가하지 않은 파일 · 교체 · 취소 기록"
            .Font.Bold = True: .Font.color = RGB(255, 255, 255): .Interior.color = ClrGray()
        End With
        ws.rows(r).RowHeight = 22
        r = r + 1
        For j = 1 To m
            parts = Split(CfgGet("추경기록" & j) & " |  | ", " | ")
            ws.Cells(r, 2).Value2 = parts(0)
            ws.Range(ws.Cells(r, 3), ws.Cells(r, 6)).Merge          '긴 파일 이름이 결과 글자에 가려지지 않게 C:F
            ws.Cells(r, 3).Value2 = parts(1)
            ws.Range(ws.Cells(r, 7), ws.Cells(r, CK_NC)).Merge      '결과 G:J (길면 줄바꿈)
            ws.Cells(r, 7).Value2 = parts(2)
            ws.Cells(r, 7).WrapText = True
            ws.Range(ws.Cells(r, 2), ws.Cells(r, CK_NC)).VerticalAlignment = xlCenter
            If Len(CStr(parts(2))) > 60 Then ws.rows(r).RowHeight = 28
            If Left$(CStr(parts(2)), 2) = "중복" Then ws.Cells(r, 7).Font.color = ClrOrange()
            r = r + 1
        Next j
    End If

    '맨 위 요약: 확인 필요 n건 = 최종 상태(본예산 + 모든 추경 반영) 기준 + 종류별 건수. 차수별 숫자는 기록만
    With ws.Range("E2:F3")
        .Merge
        .Value2 = "확인 필요 " & nFinal & "건"
        .Font.color = IIf(nFinal > 0, ClrOrange(), ClrGreen())
        .Font.Size = 14: .Font.Bold = True: .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter
    End With
    sumTxt = NoteSummary(nFinal, fA, fB, fC, jl)
    With ws.Range("G2:J3")
        .Merge
        .Value2 = Mid$(sumTxt, InStr(sumTxt, "(") + 1, Len(sumTxt) - InStr(sumTxt, "(") - 1) & vbLf & "최종 상태 기준   |   차수별 기록: " & hist
        .Font.Size = 9: .Font.color = ClrGray(): .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter: .WrapText = True
    End With
    ws.Range("A7").Value2 = sumTxt & "  = 최종 상태 기준   |   차수별 기록: " & brk & "   (앞 차수 구역의 색칠 행은 이력이며 맨 위 숫자에 넣지 않음 · 추경 등으로 정상일 수 있음)" & _
        IIf(jl, "", "   |   장표(102-B)를 불러오면 장표 대조도 합니다.")
    ws.Range("A7").Font.Bold = True: ws.Range("A7").Font.color = IIf(nFinal > 0, ClrOrange(), ClrGreen())
    ws.rows(7).RowHeight = 20

    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Range("A9").Select
    ActiveWindow.FreezePanes = True
    ActiveWindow.DisplayGridlines = False
    ws.Range("A1").Select
End Sub

'구역 하나 쓰기. idx: 0=본예산, 1..=추경 차수, -1=최종 상태. 반환: 확인 필요(비고 있는) 행 수, cA/cB/cC = 종류별 건수
Private Function WriteSection(ws As Worksheet, ByRef r As Long, ByVal idx As Long, d As Object, prevKeys As Object, ByVal isFinal As Boolean, _
                              ByRef cA As Long, ByRef cb As Long, ByRef cC As Long) As Long
    Dim k As Variant, p As clsProject, n As Long, i As Long, out() As Variant, bad() As Boolean, hdrRow As Long
    Dim jang As Double, found As Boolean, sumAmt As Double, sumCalc As Double, sumJang As Double, calcKnown As Boolean
    Dim nAdd As Long, nChg As Long, nSame As Long, nKeep As Long, nBad As Long, used As Object, a As clsProject
    Dim lbl As String, info As String, res As String, det As String, row0 As Long, nJangOnly As Long, jTotal As Double
    Dim clrD As Long, clrL As Long, bk As Variant, cat As Long, note As String, cl As String
    Dim nEval As Long, nEvalOk As Long, nUnk As Long, nInh As Long

    If isFinal Then
        lbl = "최종 상태 (본예산 + 추경 " & SuppCount() & "개 차수 반영)"
        clrD = ClrNavy(): clrL = ClrLight()
    Else
        lbl = SuppLabel(idx)
        clrD = ChasuColorDark(idx): clrL = ChasuColor(idx)
    End If

    '행 데이터
    n = d.Count
    Set used = CreateObject("Scripting.Dictionary")
    If isFinal And JangLoaded() Then
        For Each k In d.keys
            Set p = d(k)
            If p.장표키 <> "" Then used(p.장표키) = 1
        Next k
        For Each bk In gBAgg.keys
            If Not used.Exists(bk) Then nJangOnly = nJangOnly + 1
        Next bk
    End If
    ReDim out(1 To n + nJangOnly + 1, 1 To CK_NC): ReDim bad(1 To n + nJangOnly + 1)
    calcKnown = True
    i = 0
    For Each k In d.keys
        Set p = d(k)
        i = i + 1
        cat = 0: note = ""
        out(i, 1) = i: out(i, 2) = p.세부사업: out(i, 3) = p.사업항목
        If isFinal Then
            out(i, 4) = IIf(p.차수 = "", "본예산", p.차수)
            cl = out(i, 4)
        ElseIf idx = 0 Then
            out(i, 4) = "본예산": cl = "본예산"
        Else
            out(i, 4) = p.변경: cl = SuppLabel(idx)
            Select Case p.변경
                Case "추가": nAdd = nAdd + 1
                Case "변경": nChg = nChg + 1
                Case Else: nSame = nSame + 1
            End Select
        End If
        out(i, 5) = p.요구액: sumAmt = sumAmt + p.요구액
        If p.산출합 >= 0 Then
            out(i, 6) = p.산출합: sumCalc = sumCalc + p.산출합
        Else
            calcKnown = False
        End If
        If JangLoaded() Then
            jang = JangOf(CStr(k), found)
            If found Then
                out(i, 7) = jang: sumJang = sumJang + jang
                out(i, 8) = p.요구액 - jang
                If p.산출합 >= 0 Then out(i, 9) = p.산출합 - jang
            Else
                out(i, 8) = p.요구액
                If p.산출합 >= 0 Then out(i, 9) = p.산출합
            End If
            note = JangNote(p.요구액, p.산출합, found, jang, cl, cat)
        End If
        '산출내역 계산: 추경 구역에서 앞 차수와 같은 사업(동일)은 금액을 정한 앞 차수 구역에서 계산했으므로 생략
        If Not isFinal And idx > 0 And p.변경 <> "추가" And p.변경 <> "변경" Then
            nInh = nInh + 1
        Else
            If p.산출식상태 = 0 Then
                nEval = nEval + 1
                If Abs(p.요구액 - p.산출식합) <= p.산출식허용 + 0.5 Then nEvalOk = nEvalOk + 1
            ElseIf p.산출식상태 = 1 Then
                nUnk = nUnk + 1
            End If
            note = JoinNote(note, CalcNote(p, cat))
        End If
        out(i, 10) = note
        If cat <> 0 Then
            bad(i) = True: nBad = nBad + 1
            If cat And 1 Then cA = cA + 1
            If cat And 2 Then cb = cb + 1
            If cat And 4 Then cC = cC + 1
        End If
    Next k
    '최종 상태: 장표에만 있는 사업
    If isFinal And nJangOnly > 0 Then
        For Each bk In gBAgg.keys
            If Not used.Exists(bk) Then
                Set a = gBAgg(bk)
                i = i + 1
                out(i, 1) = i: out(i, 2) = a.세부사업: out(i, 3) = a.장표명: out(i, 4) = "장표에만"
                out(i, 5) = 0: out(i, 7) = a.장표예산: sumJang = sumJang + a.장표예산
                out(i, 8) = -a.장표예산
                If a.장표예산 <> 0 Then
                    out(i, 10) = "세출 파일에 없음(확인 필요)": bad(i) = True: nBad = nBad + 1: cA = cA + 1
                End If
            End If
        Next bk
    End If
    If Not isFinal And idx > 0 Then
        For Each k In prevKeys.keys
            If Not d.Exists(k) Then nKeep = nKeep + 1
        Next k
    End If
    If JangLoaded() Then jTotal = sumJang

    '구역 제목
    info = "■ " & lbl
    If Not isFinal Then
        info = info & "   |   파일: " & FileNameOf(SuppFile(idx))
        If idx > 0 Then info = info & "   |   추가 " & CfgGet("추경" & idx & "시각") Else info = info & "   |   가져온 시각 " & CfgGet("요구서가져온시각")
        info = info & "   |   " & IIf(idx = 0, CfgGet("올해연도"), CfgGet("추경" & idx & "연도")) & "년"
    End If
    info = info & "   |   확인 필요 " & nBad & "건"
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, CK_NC))
        .Merge
        .Value2 = info
        .Font.Bold = True: .Font.Size = 11: .Font.color = RGB(255, 255, 255): .Interior.color = clrD
        .VerticalAlignment = xlCenter
    End With
    ws.rows(r).RowHeight = 24
    r = r + 1

    '점검 결과 줄
    If Not isFinal Then
        If idx = 0 Then
            CheckLine ws, r, "중복 파일", "기준", "본예산(기준 파일) - 이후 추가하는 추경 파일과 경로·내용 지문을 비교합니다."
            CheckLine ws, r, "연도", "기준", "올해 " & CfgGet("올해연도") & "년 · 작년 " & IIf(val(CfgGet("작년연도")) > 0, CfgGet("작년연도") & "년", "자료 없음") & " (" & CfgGet("연도출처") & "에서 인식)"
            CheckLine ws, r, "차수 판단", "기준", "본예산 (추경 파일은 제목 · 머리글 · 기정액으로 차수를 판단하고 차수 순서로 정렬, 파일 이름은 쓰지 않음)" & _
                IIf(CfgGet("요구서내용메모") <> "", "  " & CfgGet("요구서내용메모"), "")
        Else
            det = CfgGet("추경" & idx & "중복점검"): CheckLine ws, r, "중복 파일", ResultWord(det), AfterColon(det)
            det = CfgGet("추경" & idx & "연도점검"): CheckLine ws, r, "연도 일치", ResultWord(det), AfterColon(det)
            det = CfgGet("추경" & idx & "차수근거"): CheckLine ws, r, "차수 판단", ResultWord(det), AfterColon(det)
            det = CfgGet("추경" & idx & "순서점검"): CheckLine ws, r, "차수 순서", ResultWord(det), AfterColon(det)
        End If
    End If
    If JangLoaded() Then
        det = "세출파일 합계 " & Format$(sumAmt, "#,##0") & "원 · 산출내역 합계 " & IIf(calcKnown, Format$(sumCalc, "#,##0") & "원", "확인 불가(세목코드 열 없음)") & _
              " · 장표 합계 " & Format$(jTotal, "#,##0") & "원 · 차이액(세출-장표) " & Format$(sumAmt - jTotal, "#,##0") & "원" & _
              IIf(calcKnown, " · 차이액(산출-장표) " & Format$(sumCalc - jTotal, "#,##0") & "원", "")
        If IsMismatch(sumAmt, jTotal) Or (calcKnown And IsMismatch(sumCalc, jTotal)) Then res = "확인 필요" Else res = "일치"
        If Not isFinal And idx < SuppCount() And res = "확인 필요" Then det = det & "  (뒤 차수 추경이 있어 장표와 다른 것이 정상일 수 있음)"
    Else
        det = "세출파일 합계 " & Format$(sumAmt, "#,##0") & "원 · 장표를 불러오지 않아 대조 생략"
        res = "장표 없음"
    End If
    CheckLine ws, r, "합계 대조", res, det
    det = "산출식 계산 " & nEval & "건 중 최신 산출내역과 일치 " & nEvalOk & "건 · " & NOTE_B & " " & cb & "건 · 예산현황-산출내역 확인 " & cC & "건" & _
          IIf(nUnk > 0, " · 산출식을 읽지 못한 사업 " & nUnk & "건(비교 생략)", "") & _
          IIf(nInh > 0, " · 앞 차수와 같은 사업 " & nInh & "건은 앞 차수에서 계산", "") & "  (" & CalcRuleText() & ")"
    CheckLine ws, r, "산출내역 계산", IIf(cb + cC > 0, "확인 필요", "일치"), det
    If Not isFinal And idx > 0 Then
        CheckLine ws, r, "변경 요약", "정보", "추가 " & nAdd & "건 · 변경 " & nChg & "건 · 동일 " & nSame & "건 · 이 파일에 없어 앞 차수 값 유지 " & nKeep & "건 (같은 사업은 교체, 더하지 않음)"
    End If
    If isFinal And nJangOnly > 0 Then
        CheckLine ws, r, "장표에만 있음", "확인", "장표에는 있으나 세출 파일(본예산+추경)에 없는 사업 " & nJangOnly & "건 (아래 표 '장표에만')"
    End If

    '항목 표
    hdrRow = r
    ws.Cells(r, 1).Resize(1, CK_NC).Value2 = Array("No", "세부사업", "사업항목", IIf(isFinal, "최종 차수", IIf(idx = 0, "구분", "변경")), _
        "세출파일 금액", "산출내역 합계", "장표 예산금액", "차이액(세출-장표)", "차이액(산출-장표)", "비고")
    StyleHeader ws.Range(ws.Cells(r, 1), ws.Cells(r, CK_NC))
    ws.rows(r).RowHeight = 22
    r = r + 1
    row0 = r
    If i > 0 Then
        ws.Cells(r, 1).Resize(i, CK_NC).Value2 = out
        With ws.Range(ws.Cells(r, 1), ws.Cells(r + i - 1, CK_NC))
            .Font.name = FONT_NM: .Font.Size = 10: .Borders.LineStyle = xlContinuous: .Borders.color = ClrHead()
        End With
        ws.Range(ws.Cells(r, 5), ws.Cells(r + i - 1, 9)).NumberFormat = "#,##0;[Red]-#,##0;""-"""
        ws.Range(ws.Cells(r, 1), ws.Cells(r + i - 1, 1)).HorizontalAlignment = xlCenter
        ws.Range(ws.Cells(r, 4), ws.Cells(r + i - 1, 4)).HorizontalAlignment = xlCenter
        ws.Range(ws.Cells(r, 10), ws.Cells(r + i - 1, 10)).HorizontalAlignment = xlLeft
        ws.Range(ws.Cells(r, 10), ws.Cells(r + i - 1, 10)).ShrinkToFit = True
        If Not isFinal Then ws.Range(ws.Cells(r, 4), ws.Cells(r + i - 1, 4)).Interior.color = clrL
        For n = 1 To i
            If bad(n) Then
                With ws.Range(ws.Cells(r + n - 1, 1), ws.Cells(r + n - 1, CK_NC))
                    .Interior.color = MismatchColor(): .Font.color = RGB(156, 0, 6)
                End With
            End If
        Next n
        r = r + i
    End If
    r = r + 1
    WriteSection = nBad
End Function

Private Function ResultWord(ByVal det As String) As String
    Dim p As Long
    p = InStr(det, ":")
    If p > 0 Then ResultWord = Left$(det, p - 1) Else ResultWord = IIf(det = "", "기록 없음", "정보")
End Function

Private Function AfterColon(ByVal det As String) As String
    Dim p As Long
    p = InStr(det, ":")
    If p > 0 Then AfterColon = Trim$(Mid$(det, p + 1)) Else AfterColon = det
End Function

Private Sub CheckLine(ws As Worksheet, ByRef r As Long, ByVal item As String, ByVal res As String, ByVal det As String)
    ws.Cells(r, 2).Value2 = item
    ws.Cells(r, 2).Font.Bold = True
    ws.Cells(r, 3).Value2 = res
    ws.Cells(r, 3).Font.Bold = True
    Select Case res
        Case "통과", "일치", "기준": ws.Cells(r, 3).Font.color = ClrGreen()
        Case "경고", "장표 없음": ws.Cells(r, 3).Font.color = ClrRed()
        Case Else: ws.Cells(r, 3).Font.color = ClrOrange()          '확인 필요 · 재정렬 · 제안 · 확인 · 정보
    End Select
    ws.Range(ws.Cells(r, 4), ws.Cells(r, CK_NC)).Merge
    ws.Cells(r, 4).Value2 = det
    ws.Cells(r, 4).Font.color = ClrNavy()
    ws.Range(ws.Cells(r, 2), ws.Cells(r, CK_NC)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    ws.Range(ws.Cells(r, 2), ws.Cells(r, CK_NC)).Borders(xlEdgeBottom).color = ClrHead()
    ws.rows(r).RowHeight = 18
    r = r + 1
End Sub

'중복 파일 경고 문구 (같은 경로 또는 같은 내용 지문). matchLbl = 이미 추가된 쪽 (본예산 / n차 추경)
Public Function DupWarnText(ByVal fn As String, ByVal matchLbl As String) As String
    DupWarnText = "이미 업로드한 내용입니다." & vbCrLf & fn & "은(는) " & matchLbl & "으로 이미 추가된 파일과 내용이 같아 추가하지 않았습니다."
End Function

Private Sub DupWarn(ByVal fn As String, ByVal matchLbl As String)
    Dim T As String
    T = DupWarnText(fn, matchLbl)
    If gBX_NoUI Then
        gBX_DupWarn = gBX_DupWarn & IIf(gBX_DupWarn = "", "", vbLf & vbLf) & fn & vbTab & T
    Else
        Application.ScreenUpdating = True
        MsgBox T, vbExclamation, DUP_TITLE
    End If
End Sub

'확인 질문: 평소에는 MsgBox, 자체 점검 중에는 묻지 않고 '예' (점검 결과에는 '경고 ... 사용자 확인 후 추가'로 남음)
Private Function AskYN(ByVal prompt As String, ByVal buttons As VbMsgBoxStyle, ByVal title As String) As VbMsgBoxResult
    If gBX_NoUI Then AskYN = vbYes Else AskYN = MsgBox(prompt, buttons, title)
End Function

'---------------------------------------------------------------- V8 자체 점검 (숨김 진입점, 대화상자 없음)
'  Application.Run "'BX-v8.xlam'!BX_SelfTest", 폴더, 결과파일(.xlsx)
'  폴더의 데모 파일로 본예산 → 장표 → 2차 → 1차(차수 순서와 다르게 추가) → 1차 사본(중복 시도) 을 실행하고 결과 통합문서를 저장한다.
'  반환: 결과 요약 (줄마다 "항목=값")
Public Function BX_SelfTest(ByVal folder As String, ByVal outFile As String) As String
    If Right$(folder, 1) <> "\" Then folder = folder & "\"
    Dim r As String, want As String
    r = RunSelfTest(folder & "2027_본예산_세출예산요구서_가상교육과.xlsx", folder & "102-B_집행장표_가상교육과_2027.xlsx", _
        Array(folder & "2027_제2회추경_세출예산요구서_가상교육과.xlsx", folder & "2027_제1회추경_세출예산요구서_가상교육과.xlsx", _
              folder & "2027_제1회추경_세출예산요구서_가상교육과_사본(중복시도).xlsx"), outFile, False)
    '중복 경고 문구 확인 (표시했을 문구 = 정해진 문구)
    want = "이미 업로드한 내용입니다." & " [줄바꿈] " & "2027_제1회추경_세출예산요구서_가상교육과_사본(중복시도).xlsx은(는) 1차 추경으로 이미 추가된 파일과 내용이 같아 추가하지 않았습니다."
    If InStr(r, "dupWarnCount=1 ") > 0 And InStr(r, "dupWarnWouldShow=" & want & vbLf) > 0 Then
        r = Replace(r, "RESULT=OK", "dupWarnAssert=OK" & vbLf & "RESULT=OK")
    Else
        r = Replace(r, "RESULT=OK", "dupWarnAssert=FAIL (기대: " & want & ")" & vbLf & "RESULT=FAIL")
    End If
    BX_SelfTest = r
End Function

'  Application.Run "'BX-v8.xlam'!BX_SelfTestFiles", 본예산파일, "추경1|추경2|...", 장표파일, 결과파일, 비공개(True = 사업명 등 내용 없이 집계만 반환)
'  추경 파일은 적은 순서대로 추가 (차수는 파일 내용으로 판단 · 차수 순서로 재정렬)
Public Function BX_SelfTestFiles(ByVal baseFile As String, ByVal suppList As String, ByVal jangFile As String, ByVal outFile As String, _
                                 Optional ByVal privateMode As Boolean = True) As String
    Dim a As Variant
    If Trim$(suppList) = "" Then a = Array() Else a = Split(suppList, "|")
    BX_SelfTestFiles = RunSelfTest(baseFile, jangFile, a, outFile, privateMode)
End Function

Private Function RunSelfTest(ByVal baseFile As String, ByVal jangFile As String, supps As Variant, ByVal outFile As String, ByVal priv As Boolean) As String
    Dim res As String, summary As String, wb As Workbook, f As Variant, mp As Object, dept As String
    Dim ws As Worksheet, r As Long, lr As Long, nBad As Long, T As String, i As Long, s As Long
    On Error GoTo EH
    gBX_NoUI = True
    gBX_DupWarn = ""
    gStep = "새 보드"
    Set wb = Workbooks.Add
    Set gBoardWB = wb
    CfgSheet
    SetupMainSheet
    RemoveBlankDefaultSheets
    'V7 보드 흉내: V7 이름 · 다른 추가기능 이름의 버튼을 심어 두고 새로고침 때 V8 버튼으로 바뀌는지 확인
    Dim v7s As Shape
    Set v7s = wb.Worksheets(SH_MAIN).Shapes.AddShape(msoShapeRoundedRectangle, 4, 4, 96, 22)
    v7s.name = "btn3": v7s.TextFrame2.TextRange.text = "③ 새로고침(리셋)"
    On Error Resume Next
    v7s.OnAction = "'BX-v7.xlam'!새로고침버튼"
    On Error GoTo EH
    res = res & "v7ButtonPlanted=" & wb.Worksheets(SH_MAIN).Shapes.Count & " shapes (btn3 onAction=" & v7s.OnAction & ")" & vbLf
    gStep = "본예산"
    Set mp = ImportReqCore(baseFile)
    res = res & "year=" & mp("yCur") & "/" & mp("yPrv") & "/" & mp("ySrc") & vbLf
    res = res & "baseNote=" & CfgGet("요구서내용메모") & vbLf
    gStep = "장표"
    dept = CfgGet("부서명")
    res = res & "dept=" & IIf(priv, "(" & Len(dept) & "자)", dept) & vbLf
    If jangFile <> "" Then
        res = res & "jang=" & ImportJangCore(jangFile, dept) & vbLf
        CfgSet "장표가져온시각", NowText()
    End If
    gStep = "추경"
    EnsureBaseRecorded
    For Each f In supps
        If Trim$(CStr(f)) <> "" Then AddOneSupp CStr(f), summary
    Next f
    res = res & "suppCount=" & SuppCount() & vbLf
    '중복 경고: 자체 점검 중이라 표시하지 않음 → 표시했을 문구를 그대로 기록 (비공개 모드는 파일명을 [파일명]으로)
    Dim dw As Variant, dwi As Long, dwT As String
    If gBX_DupWarn <> "" Then
        dw = Split(gBX_DupWarn, vbLf & vbLf)
        res = res & "dupWarnCount=" & (UBound(dw) + 1) & " (창 제목 '" & DUP_TITLE & "', vbExclamation, 자체 점검이라 표시 안 함)" & vbLf
        For dwi = 0 To UBound(dw)
            dwT = Mid$(dw(dwi), InStr(dw(dwi), vbTab) + 1)
            If priv Then dwT = Replace(dwT, Left$(dw(dwi), InStr(dw(dwi), vbTab) - 1), "[파일명]")
            res = res & "dupWarnWouldShow=" & Replace(dwT, vbCrLf, " [줄바꿈] ") & vbLf
        Next dwi
    Else
        res = res & "dupWarnCount=0" & vbLf
    End If
    res = res & "supp=" & Replace(Mid$(summary, 3), vbCrLf, " ; ") & vbLf
    For s = 1 To SuppCount()
        res = res & "  entry" & s & "=" & SuppLabel(s) & " | 추가순번 " & CfgGet("추경" & s & "추가순번") & " | " & FileNameOf(SuppFile(s)) & vbLf
        res = res & "    차수판단=" & CfgGet("추경" & s & "차수근거") & vbLf
        res = res & "    순서=" & CfgGet("추경" & s & "순서점검") & vbLf
        res = res & "    연도=" & CfgGet("추경" & s & "연도점검") & vbLf
    Next s
    gStep = "새로고침"
    새로고침 True
    SpeedOn "자체 점검"
    gStep = "세부산출내역": BuildCalc
    gStep = "추경점검": BuildCheckSheet
    SpeedOff
    gStep = "구조 점검"
    For s = 0 To SuppCount()
        res = res & "parse" & s & "=" & ParseInfo(s) & vbLf
    Next s
    gStep = "결과 읽기"
    Set ws = wb.Worksheets(SH_CALC)
    res = res & "calcTop=" & ws.Range("D2").Value2 & " | " & Replace(CStr(ws.Range("G2").Value2), vbLf, " / ") & vbLf
    res = res & "calcHdr=" & ws.Cells(9, 5).Value2 & " | " & ws.Cells(9, 6).Value2 & " | " & ws.Cells(9, 8).Value2 & " | " & ws.Cells(9, 9).Value2 & " | " & ws.Cells(9, 15).Value2 & vbLf
    lr = ws.Cells(ws.rows.Count, 1).End(xlUp).row
    nBad = 0: T = ""
    Dim nTag As Long
    For r = 10 To lr
        If CStr(ws.Cells(r, 15).Value2) <> "" And CStr(ws.Cells(r, 17).Value2) = "1" Then
            nBad = nBad + 1
            If Not priv Then T = T & ws.Cells(r, 3).Value2 & "=" & ws.Cells(r, 15).Value2 & ", "
        End If
        If CStr(ws.Cells(r, 14).Value2) Like "*#추*" Then nTag = nTag + 1
        If Not priv And CStr(ws.Cells(r, 17).Value2) = "1" Then res = res & "calcItem=" & ws.Cells(r, 3).Value2 & " | " & ws.Cells(r, 13).Value2 & " | " & ws.Cells(r, 8).Value2 & " | " & ws.Cells(r, 9).Value2 & " | " & ws.Cells(r, 10).Value2 & " | 메모 " & ws.Cells(r, 14).Value2 & vbLf
    Next r
    res = res & "calcNoteRows=" & nBad & IIf(priv, "", " (" & T & ")") & vbLf
    res = res & "calcMemoTagRows=" & nTag & vbLf
    Set ws = wb.Worksheets(SH_CHECK)
    res = res & "checkTop=" & ws.Range("E2").Value2 & " | " & Replace(CStr(ws.Range("G2").Value2), vbLf, " / ") & vbLf
    res = res & "checkA7=" & ws.Range("A7").Value2 & vbLf
    lr = ws.Cells(ws.rows.Count, 2).End(xlUp).row
    If ws.Cells(ws.rows.Count, 1).End(xlUp).row > lr Then lr = ws.Cells(ws.rows.Count, 1).End(xlUp).row
    For r = 8 To lr
        T = CStr(ws.Cells(r, 1).Value2)
        If Left$(T, 2) = "■ " Then res = res & "section=" & T & vbLf
        Select Case CStr(ws.Cells(r, 2).Value2)
            Case "중복 파일", "연도 일치", "차수 판단", "차수 순서", "합계 대조", "산출내역 계산", "변경 요약"
                res = res & "  check=" & ws.Cells(r, 2).Value2 & " : " & ws.Cells(r, 3).Value2 & " : " & Left$(CStr(ws.Cells(r, 4).Value2), IIf(priv, 60, 160)) & vbLf
        End Select
        If CStr(ws.Cells(r, 7).Value2) Like "중복*" Or CStr(ws.Cells(r, 7).Value2) Like "같은 차수*" Then res = res & "  log=" & ws.Cells(r, 3).Value2 & " : " & ws.Cells(r, 7).Value2 & vbLf
        If Not priv And CStr(ws.Cells(r, 10).Value2) <> "" And IsNumeric(ws.Cells(r, 1).Value2) Then res = res & "    note=" & ws.Cells(r, 3).Value2 & " : " & ws.Cells(r, 10).Value2 & vbLf
    Next r
    T = CStr(wb.Worksheets(SH_MAIN).Range("D5").Value2)
    If priv Then
        If InStr(T, "확인 필요") > 0 Then T = Mid$(T, InStr(T, "확인 필요")) Else T = "(확인 필요 표시 없음)"
    End If
    res = res & "mainD5=" & T & vbLf
    gStep = "시트 버튼"
    T = "": i = 0
    Dim shp As Shape, nOld As Long, nBadOA As Long
    For Each shp In wb.Worksheets(SH_MAIN).Shapes
        If shp.name Like "bx8_*" Then
            i = i + 1
            T = T & shp.TextFrame2.TextRange.text & "(" & shp.TextFrame2.TextRange.Font.Size & "pt)" & ", "
            If OnActionBook(shp.OnAction) <> LCase$(ThisWorkbook.name) Then nBadOA = nBadOA + 1
        ElseIf shp.name Like "btn*" Then
            nOld = nOld + 1
        End If
    Next shp
    res = res & "sheetButtons=" & i & " old=" & nOld & " badOnAction=" & nBadOA & " addin=" & ThisWorkbook.name & " onAction(예)=" & wb.Worksheets(SH_MAIN).Shapes("bx8_supp").OnAction & " : " & T & vbLf
    res = res & "buttonsCurrent=" & BoardButtonsCurrent(wb) & vbLf
    gStep = "저장"
    Application.DisplayAlerts = False
    wb.SaveAs fileName:=outFile, FileFormat:=51
    Application.DisplayAlerts = True
    res = res & "saved=" & outFile & vbLf
    gBX_NoUI = False
    RunSelfTest = res & "RESULT=OK"
    Exit Function
EH:
    res = res & "ERROR=" & Err.Description & " (" & Err.Number & ") @" & gStep & vbLf
    gBX_NoUI = False
    On Error Resume Next
    SpeedOff
    Application.StatusBar = False
    RunSelfTest = res & "RESULT=FAIL"
End Function

'구조 점검 한 줄: 머리글 행 · 고른 열 · 단위 · 사업 수 · 세출 금액과 세목 행 합계가 다른 사업 수 · 산출식을 못 읽은 사업 수
Private Function ParseInfo(ByVal s As Long) As String
    Dim ws As Worksheet, mp As Object, d As Object, k As Variant, p As clsProject, nNe As Long, nUnk As Long, nOldUnk As Long
    Dim g As Double, fsum As Double
    On Error GoTo EH
    If s = 0 Then Set ws = BW().Worksheets(SH_REQ) Else Set ws = BW().Worksheets(SuppSheetName(s))
    Set mp = YH_Map(ws, "", "", False)
    ParseInfo = SuppLabel(s) & " | hr=" & mp("hr") & " calc=" & ColLetter(mp("calc")) & " calcP=" & ColLetter(mp("calcP")) & " calcG=" & ColLetter(mp("calcG")) & _
        " amt=" & ColLetter(mp("amt")) & IIf(mp("gd"), "(기정 " & ColLetter(mp("gij")) & " + 추경요구액 " & ColLetter(mp("delta")) & ")", "") & _
        " amtP=" & ColLetter(mp("amtP")) & " code=" & ColLetter(mp("code")) & " unit=" & mp("unit") & " year=" & mp("yCur") & "/" & mp("yPrv") & "(" & mp("ySrc") & ")"
    If Not gChasu Is Nothing Then
        If gChasu.Exists(s) Then
            Set d = gChasu(s)
            For Each k In d.keys
                Set p = d(k)
                If p.산출합 >= 0 And IsMismatch(p.요구액, p.산출합) Then nNe = nNe + 1
                If p.산출식상태 = 1 Then nUnk = nUnk + 1
                If p.산출식전상태 = 1 Then nOldUnk = nOldUnk + 1
            Next k
            ParseInfo = ParseInfo & " | items=" & d.Count & " 세출금액<>세목합=" & nNe & " 산출식못읽음=" & nUnk & " 비교산출식못읽음=" & nOldUnk
        End If
    End If
    FileTotals ws, mp, g, fsum
    ParseInfo = ParseInfo & IIf(g >= 0, " | 기정합=최종합(앞)? " & ChainText(s, g), "")
    Exit Function
EH:
    ParseInfo = ParseInfo & " | ERROR " & Err.Description
End Function

Private Function ChainText(ByVal s As Long, ByVal g As Double) As String
    Dim j As Long
    If Abs(g - val(CfgGet("요구서최종합"))) < 1 Then ChainText = "본예산": Exit Function
    For j = 1 To SuppCount()
        If j <> s And CfgGet("추경" & j & "최종합") <> "" Then
            If Abs(g - val(CfgGet("추경" & j & "최종합"))) < 1 Then ChainText = SuppLabel(j): Exit Function
        End If
    Next j
    ChainText = "없음"
End Function

Private Function ColLetter(ByVal c As Long) As String
    If c <= 0 Then ColLetter = "-": Exit Function
    Dim n As Long
    n = c
    Do While n > 0
        ColLetter = Chr$(65 + (n - 1) Mod 26) & ColLetter
        n = (n - 1) \ 26
    Loop
End Function


