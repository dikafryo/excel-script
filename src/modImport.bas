Attribute VB_Name = "modImport"
Option Explicit
'=====================================================================
' modImport : 세출예산요구서 / 장표(102-A, 102-B) 파일 가져오기
'=====================================================================

'① 세출예산요구서 선택
Public Sub 요구서선택()
    Dim fp As Variant, dept As String, mp As Object, yTxt As String
    'V8: 이미 추경이 추가돼 있으면 본예산을 바꾸기 전에 확인 (본예산을 바꾸면 추경 차수는 모두 지워짐)
    If SuppCount() > 0 Then
        If MsgBox("추가된 추경 " & SuppCount() & "건이 있습니다." & vbCrLf & _
                  "세출예산요구서(본예산)를 새로 불러오면 추경 차수가 모두 지워집니다. 계속할까요?", _
                  vbYesNo + vbExclamation, "본예산 다시 불러오기") <> vbYes Then Exit Sub
    End If
    fp = Application.GetOpenFilename("Excel 파일 (*.xls*), *.xls*", , "세출예산요구서(본예산) 파일을 선택하세요", , False)
    If VarType(fp) = vbBoolean Then Exit Sub

    On Error GoTo EH
    SpeedOn "세출예산요구서 가져오는 중..."
    Set mp = ImportReqCore(CStr(fp))
    dept = CfgGet("부서명")
    yTxt = "올해 " & mp("yCur") & "년" & IIf(mp("yPrv") > 0, " · 작년 " & mp("yPrv") & "년", " (작년 열 없음)") & "  [" & mp("ySrc") & "]"
    SpeedOff
    AutoRefresh
    MsgBox "세출예산요구서(본예산)를 가져왔습니다." & vbCrLf & "부서: " & dept & vbCrLf & _
           "연도: " & yTxt & vbCrLf & IIf(CfgGet("요구서내용메모") <> "", CfgGet("요구서내용메모") & vbCrLf, "") & _
           IIf(SheetExists(SH_B), "보드가 갱신되었습니다.", "이제 ② 장표(102-B) 선택으로 집행 장표를 불러오세요.") & vbCrLf & _
           "추경이 있으면 [추경예산 추가]로 추가하세요. (차수는 파일 내용으로 판단해 차수 순서로 정렬)", vbInformation
    Exit Sub
EH:
    SpeedOff
    MsgBox "요구서 가져오기 오류: " & Err.Description, vbCritical
End Sub

'V8: 세출예산요구서(본예산) 가져오기 본체 (대화상자 없음 - 버튼과 자체 점검이 함께 사용). 반환: 연도 지도(YH_Map)
Public Function ImportReqCore(ByVal fp As String) As Object
    Dim wbS As Workbook, wsS As Worksheet, wsD As Worksheet, dept As String, mp As Object, eN As Long, eD As String
    Dim gS As Double, fs As Double, nb As Long, isSupp As Boolean, why As String
    On Error GoTo EH
    Set wbS = Workbooks.Open(fileName:=fp, ReadOnly:=True, UpdateLinks:=0)
    Set wsS = wbS.Worksheets(1)
    Set wsD = GetSheet(SH_REQ, True)
    CopyValues wsS, wsD, "", ""
    YH_FillMerged wsS, wsD                       '병합된 연도 묶음 머리글(예: 2027 / 2026)을 값으로 채움
    dept = ExtractDept(wsD)
    wbS.Close SaveChanges:=False
    Set wbS = Nothing
    If dept <> "" Then CfgSet "부서명", dept
    CfgSet "요구서파일", fp
    CfgSet "요구서가져온시각", NowText()
    '머리글에서 올해·작년 연도 자동 인식 (못 찾으면 파일 제목·이름 → 직접 입력, 자체 점검 중에는 묻지 않고 오류)
    CfgSet "올해연도", ""
    Set mp = YH_Map(wsD, fp, "올해연도", Not gBX_NoUI)
    CfgSet "올해연도", CStr(mp("yCur")): CfgSet "작년연도", CStr(mp("yPrv")): CfgSet "연도출처", CStr(mp("ySrc"))
    CfgSet "요구서지문", DataFingerprint(wsD)
    'V8: 본예산 합계(추경 기정액 연결 확인용) + 내용에 추경 표시가 있으면 참고 메모만 (파일 이름은 보지 않음)
    FileTotals wsD, mp, gS, fs
    CfgSet "요구서최종합", Format$(fs, "0")
    nb = ChasuFromContent(wsD, mp, isSupp, why)
    If isSupp Then
        CfgSet "요구서내용메모", "참고: ① 파일 내용에 추경 표시가 있습니다" & IIf(why <> "", " (" & why & ")", " (기정예산액 열)") & ". 본예산(기준) 파일로 그대로 사용합니다."
    Else
        CfgSet "요구서내용메모", ""
    End If
    ClearAllSupp
    CfgSet "추경기록수", "0"
    wsD.Visible = xlSheetHidden
    Set ImportReqCore = mp
    Exit Function
EH:
    eN = Err.Number: eD = Err.Description
    On Error Resume Next
    If Not wbS Is Nothing Then wbS.Close SaveChanges:=False
    On Error GoTo 0
    Err.Raise eN, "ImportReqCore", eD
End Function

'② 장표(102-A, 102-B) 선택 (복수 선택 가능)
Public Sub 장표선택()
    Dim fps As Variant, i As Long, dept As String, Done As String, res As String
    fps = Application.GetOpenFilename("Excel 파일 (*.xls*), *.xls*", , "장표 파일을 선택하세요 (102-B 필수, 102-A 선택 / Ctrl로 함께 선택 가능)", , True)
    If VarType(fps) = vbBoolean Then Exit Sub

    On Error GoTo EH
    SpeedOn "장표 가져오는 중..."
    dept = CfgGet("부서명")
    For i = LBound(fps) To UBound(fps)
        res = ImportJangCore(CStr(fps(i)), dept)
        If res = "" Then
            MsgBox "장표 형식을 인식하지 못해 건너뜁니다:" & vbCrLf & fps(i), vbExclamation
        Else
            Done = Done & vbCrLf & " - " & res
        End If
    Next i
    If dept <> "" Then CfgSet "부서명", dept
    If dept = "" Then MsgBox "부서를 지정하지 않아 장표 전체 행을 가져왔습니다. [① 세출예산요구서 선택] 또는 [부서 선택] 후 다시 ② 를 실행하세요.", vbExclamation
    CfgSet "장표가져온시각", NowText()
    SpeedOff
    AutoRefresh
    MsgBox "장표를 가져왔습니다. (부서: " & dept & ", 기준일: " & CfgGet("기준일") & ")" & Done, vbInformation
    Exit Sub
EH:
    SpeedOff
    MsgBox "장표 가져오기 오류: " & Err.Description, vbCritical
End Sub

'V8: 장표 파일 하나 가져오기 본체 (대화상자 없음). 반환: "102-B : n행", 형식을 모르면 ""
'  dept 가 비어 있으면 (버튼 사용 시) 부서 목록에서 고르게 함 - 자체 점검 중에는 묻지 않고 전체 행
Public Function ImportJangCore(ByVal fp As String, ByRef dept As String) As String
    Dim wbS As Workbook, wsS As Worksheet, wsD As Worksheet, kind As String, hr As Long, hm As Object, cnt As Long
    Dim eN As Long, eD As String
    On Error GoTo EH
    Set wbS = Workbooks.Open(fileName:=fp, ReadOnly:=True, UpdateLinks:=0)
    Set wsS = wbS.Worksheets(1)
    hr = FindHeaderRow(wsS, "회계연도", 10)
    If hr > 0 Then
        Set hm = HeaderMap(wsS, hr)
        If hm.Exists("사업항목명1") Then
            kind = "B"
        ElseIf hm.Exists("사업항목1레벨") Then
            kind = "A"
        End If
    End If
    If kind = "" Then
        If InStr(1, fp, "102-B", vbTextCompare) > 0 Then kind = "B"
        If InStr(1, fp, "102-A", vbTextCompare) > 0 Then kind = "A"
    End If
    If kind <> "" Then
        If dept = "" And hr > 0 And Not gBX_NoUI Then dept = PickDept(wsS, hr, hm)
        Set wsD = GetSheet(IIf(kind = "A", SH_A, SH_B), True)
        cnt = CopyValues(wsS, wsD, "예산소관부서명", dept)
        If cnt = 0 Then cnt = CopyValues(wsS, wsD, "", "")      '부서 필터 결과가 없으면 전체 복사
        CfgSet "기준일", ExtractBaseDate(wsS)
        CfgSet "장표" & kind & "파일", fp
        CfgSet "장표" & kind & "행수", CStr(cnt)
        wsD.Visible = xlSheetHidden
        ImportJangCore = "102-" & kind & " : " & cnt & "행"
    End If
    wbS.Close SaveChanges:=False
    Exit Function
EH:
    eN = Err.Number: eD = Err.Description
    On Error Resume Next
    If Not wbS Is Nothing Then wbS.Close SaveChanges:=False
    On Error GoTo 0
    Err.Raise eN, "ImportJangCore", eD
End Function

'버튼: 부서명 직접 지정 (요구서 없이 사용할 때)
Public Sub 부서선택()
    Dim ans As String
    ans = Trim(InputBox("장표에서 추출할 우리 부서명을 입력하세요 (예: ○○교육과)." & vbCrLf & _
                        "장표 '예산소관부서명' 열에 이 글자가 포함된 행만 가져옵니다.", "부서 선택", CfgGet("부서명")))
    If ans = "" Then Exit Sub
    CfgSet "부서명", ans
    MsgBox "부서명이 '" & ans & "' 로 저장되었습니다. 이제 ② 장표(102-B) 선택을 실행하세요.", vbInformation
End Sub

'값 복사 (필터 열/텍스트 지정 시 해당 텍스트가 포함된 행 + 헤더 위 행만 복사). 반환: 데이터 행수
Public Function CopyValues(wsS As Worksheet, wsD As Worksheet, ByVal filterHeader As String, ByVal filterText As String) As Long
    Dim arr As Variant, ur As Range, nR As Long, nC As Long, r As Long, c As Long
    Dim hr As Long, fC As Long, out As Variant, k As Long, keep As Boolean
    Set ur = wsS.UsedRange
    arr = ur.Value2
    If Not IsArray(arr) Then CopyValues = 0: Exit Function
    nR = UBound(arr, 1): nC = UBound(arr, 2)
    '원본 좌표 유지 (UsedRange가 A1에서 시작하지 않는 경우 대비)
    Dim r0 As Long, c0 As Long
    r0 = ur.row: c0 = ur.Column

    fC = 0: hr = 0
    If filterHeader <> "" And filterText <> "" Then
        hr = FindHeaderRow(wsS, "회계연도", 10)
        If hr > 0 Then
            For c = 1 To nC
                If CleanText(arr(hr - r0 + 1, c)) = filterHeader Then fC = c: Exit For
            Next c
        End If
    End If

    If fC = 0 Then
        wsD.Cells(r0, c0).Resize(nR, nC).Value2 = arr
        CopyValues = nR - IIf(hr > 0, hr, 0)
    Else
        ReDim out(1 To nR, 1 To nC)
        k = 0
        For r = 1 To nR
            keep = (r <= hr - r0 + 1)
            If Not keep Then keep = (InStr(1, CleanText(arr(r, fC)), filterText, vbTextCompare) > 0)
            If keep Then
                k = k + 1
                For c = 1 To nC: out(k, c) = arr(r, c): Next c
            End If
        Next r
        If k - (hr - r0 + 1) <= 0 Then CopyValues = 0: Exit Function
        wsD.Cells(r0, c0).Resize(k, nC).Value2 = out
        CopyValues = k - (hr - r0 + 1)
    End If
End Function

'요구서에서 부서명 추출 ("□ 기관(부서)명 : ○○교육과")
Private Function ExtractDept(ws As Worksheet) As String
    Dim r As Long, c As Long, T As String, p As Long
    For r = 1 To 10
        For c = 1 To 15
            T = CleanText(ws.Cells(r, c).Value2)
            If InStr(T, "부서)명") > 0 Or InStr(T, "부서명") > 0 Then
                p = InStr(T, ":")
                If p > 0 Then ExtractDept = Trim(Mid(T, p + 1)): Exit Function
            End If
        Next c
    Next r
    ExtractDept = ""
End Function

'장표에 들어있는 예산소관부서 목록을 보여주고 우리 부서를 고르게 함 (요구서를 안 불러온 경우)
Public Function PickDept(ws As Worksheet, ByVal hr As Long, hm As Object) As String
    Dim c As Long, lr As Long, arr As Variant, d As Object, r As Long, T As String, k As Variant
    Dim names() As String, n As Long, i As Long, j As Long, tmp As String, msg As String, ans As String
    c = HCol(hm, "예산소관부서명", False)
    If c = 0 Then Exit Function
    lr = lastRow(ws, c)
    If lr <= hr Then Exit Function
    arr = ws.Range(ws.Cells(hr + 1, c), ws.Cells(lr, c)).Value2
    Set d = CreateObject("Scripting.Dictionary")
    For r = 1 To UBound(arr, 1)
        T = CleanText(arr(r, 1))
        If T <> "" Then d(T) = 1
    Next r
    n = d.Count
    If n = 0 Then Exit Function
    ReDim names(1 To n): i = 0
    For Each k In d.keys: i = i + 1: names(i) = k: Next k
    For i = 2 To n
        tmp = names(i): j = i - 1
        Do While j >= 1
            If StrComp(names(j), tmp, vbTextCompare) <= 0 Then Exit Do
            names(j + 1) = names(j): j = j - 1
        Loop
        names(j + 1) = tmp
    Next i
    If n = 1 Then
        arr = Split(names(1), " "): PickDept = arr(UBound(arr)): Exit Function
    End If
    msg = "장표에 " & n & "개 부서가 있습니다. 우리 부서의 번호를 입력하세요 (또는 부서명 일부 직접 입력)" & vbCrLf & vbCrLf
    For i = 1 To n
        msg = msg & i & ". " & names(i) & vbCrLf
        If Len(msg) > 900 And i < n Then msg = msg & "... (이하 생략 - 부서명을 직접 입력하세요)" & vbCrLf: Exit For
    Next i
    ans = Trim(InputBox(msg, "우리 부서 선택", ""))
    If ans = "" Then Exit Function
    If IsNumeric(ans) Then
        If CLng(ans) >= 1 And CLng(ans) <= n Then
            arr = Split(names(CLng(ans)), " "): PickDept = arr(UBound(arr)): Exit Function
        End If
    End If
    PickDept = ans
End Function

'장표에서 부서명(짧은 이름) 추출: "교육국 ○○교육과" -> "○○교육과"
Private Function DeptFromSheet(ws As Worksheet, ByVal hr As Long, hm As Object) As String
    Dim c As Long, T As String, arr As Variant
    c = HCol(hm, "예산소관부서명", False)
    If c = 0 Then Exit Function
    T = CleanText(ws.Cells(hr + 1, c).Value2)
    arr = Split(T, " ")
    DeptFromSheet = arr(UBound(arr))
End Function

'장표 상단 "기준일 : 2026-09-24"
Private Function ExtractBaseDate(ws As Worksheet) As String
    Dim r As Long, c As Long, T As String, p As Long
    For r = 1 To 5
        For c = 1 To 5
            T = CleanText(ws.Cells(r, c).text)      '표시 형식 그대로 (날짜 셀도 텍스트로)
            If InStr(T, "기준일") > 0 Then
                p = InStr(T, ":")
                If p > 0 Then
                    ExtractBaseDate = FmtDate(Trim(Mid(T, p + 1)))
                Else
                    '"기준일" 라벨 옆 셀에 날짜가 있는 경우
                    ExtractBaseDate = FmtDate(ws.Cells(r, c + 1).Value2)
                End If
                Exit Function
            End If
        Next c
    Next r
    ExtractBaseDate = FmtDate(Date)
End Function


