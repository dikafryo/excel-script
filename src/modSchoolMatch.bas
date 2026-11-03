Attribute VB_Name = "modSchoolMatch"
Option Explicit
' =====================================================================================
'  modSchoolMatch : 줄여 쓴 학교명 → 정식 학교명 판정 (규칙 + 마스터 데이터 대조)
'   1) 공백 제거 후 마스터에 그대로 있으면 → 정확
'   2) 접미사 규칙(여고·여중·여상·공고·상고·외고·과고·예고·체고·초·중·고 …)으로 후보를 만들어 마스터 대조
'      (지역명 접두어 생략형도 함께 시도: 개성중 → (지역명)개성중학교)
'   3) 그래도 없으면 같은 급의 마스터에서 앞글자·유형(여자/공업 …)이 맞는 학교를 찾음
'      → 1개면 변환, 2개 이상이면 ‘확인 필요’, 0개면 ‘목록에 없음’
'   - 학교가 아닌 흔한 낱말(비고·참고·최고 …)과 문장(공백 포함)은 건드리지 않는다.
' =====================================================================================
Public Const SM_SKIP As Long = 0        ' 학교명이 아님(그대로 둠)
Public Const SM_EXACT As Long = 1       ' 이미 정식명
Public Const SM_CONVERTED As Long = 2   ' 변환됨
Public Const SM_AMBIGUOUS As Long = 3   ' 후보 여러 개(확인 필요)
Public Const SM_NOTFOUND As Long = 4    ' 학교명 같지만 목록에 없음

' 접미사 규칙: 긴 것부터 검사. "약칭|정식" 쌍
Private Function SM_Rules() As Variant
    SM_Rules = Array( _
        "여자상업고등학교|여자상업고등학교", "여자상고|여자상업고등학교", "여상|여자상업고등학교", _
        "여자고등학교|여자고등학교", "여자고|여자고등학교", "여고|여자고등학교", _
        "여자중학교|여자중학교", "여자중|여자중학교", "여중|여자중학교", _
        "공업고등학교|공업고등학교", "공업고|공업고등학교", "공고|공업고등학교", _
        "상업고등학교|상업고등학교", "상업고|상업고등학교", "상고|상업고등학교", _
        "외국어고|외국어고등학교", "외고|외국어고등학교", _
        "과학고|과학고등학교", "과고|과학고등학교", _
        "예술고|예술고등학교", "예고|예술고등학교", _
        "체육고|체육고등학교", "체고|체육고등학교", _
        "초등학교|초등학교", "초등교|초등학교", "초교|초등학교", "초등|초등학교", "초|초등학교", _
        "중학교|중학교", "중교|중학교", "중|중학교", _
        "고등학교|고등학교", "고교|고등학교", "고등|고등학교", "고|고등학교")
End Function

' 학교가 아닌 흔한 낱말
Private Function SM_IsCommonWord(ByVal s As String) As Boolean
    Dim W As Variant
    For Each W In Array("비고", "참고", "최고", "재고", "광고", "보고", "경고", "사고", "창고", "신고", "권고", "예고", "공고", "회고", "원고", "피고", "제고", "선고", "충고", "숙고", "심고", "학교", _
                        "집중", "도중", "나중", "적중", "명중", "공중", "군중", "대중", "수중", "이중", "다중", "만중", "체중", "상중", "하중", "가중", "주중", "월중", "연중", "무중", "고등", _
                        "기초", "당초", "최초", "애초", "월초", "연초", "주초", "정초", "초등", "초중고", "초중", "중고", "고교", "학교명", "학교급", "교명")
        If s = W Then SM_IsCommonWord = True: Exit Function
    Next W
End Function

' 공백·전각공백 제거, 앞뒤 정리
Public Function SM_Clean(ByVal s As String) As String
    s = Replace(s, ChrW(12288), "")
    s = Replace(s, " ", "")
    s = Replace(s, vbTab, "")
    s = Replace(s, vbLf, ""): s = Replace(s, vbCr, "")
    SM_Clean = Trim$(s)
End Function

' 급 추정: 1 초 / 2 중 / 3 고 / 0 모름
Public Function SM_GuessLevel(ByVal s As String) As Long
    If Right$(s, 4) = "초등학교" Or Right$(s, 1) = "초" Or Right$(s, 2) = "초교" Or Right$(s, 2) = "초등" Then SM_GuessLevel = 1: Exit Function
    If Right$(s, 3) = "중학교" Or Right$(s, 1) = "중" Or Right$(s, 2) = "중교" Then SM_GuessLevel = 2: Exit Function
    If Right$(s, 4) = "고등학교" Or Right$(s, 1) = "고" Or Right$(s, 2) = "고교" Or Right$(s, 2) = "고등" Then SM_GuessLevel = 3: Exit Function
    If InStr(s, "초등") > 0 Then SM_GuessLevel = 1: Exit Function
    If InStr(s, "중학") > 0 Then SM_GuessLevel = 2: Exit Function
    If InStr(s, "고등") > 0 Or Right$(s, 4) = "영재학교" Or InStr(s, "마이스터") > 0 Then SM_GuessLevel = 3: Exit Function
End Function

' 핵심 판정. 반환: SM_* 코드. fullName/office/found/note 는 결과 채움
Public Function SM_Resolve(ByVal raw As String, ByRef fullName As String, ByRef office As String, ByRef found As String, ByRef note As String) As Long
    Dim s As String, lv As Long, cands As Variant, c As Variant, i As Long
    fullName = "": office = "": found = "": note = ""
    s = SM_Clean(raw)
    If Len(s) < 2 Then SM_Resolve = SM_SKIP: Exit Function
    If IsNumeric(s) Then SM_Resolve = SM_SKIP: Exit Function
    If InStr(raw, " ") > 0 And Len(raw) - Len(s) >= 2 Then SM_Resolve = SM_SKIP: Exit Function   ' 문장(공백 2개 이상)은 제외
    If SM_IsCommonWord(s) Then SM_Resolve = SM_SKIP: Exit Function

    ' 1) 정확히 일치
    If SD_Lookup(s, lv, office, found) Then
        fullName = s
        SM_Resolve = SM_EXACT
        Exit Function
    End If

    ' 2) 접미사 규칙 후보
    cands = SM_Candidates(s)
    For i = LBound(cands) To UBound(cands)
        c = cands(i)
        If Len(c) > 0 Then
            If SD_Lookup(CStr(c), lv, office, found) Then
                fullName = CStr(c)
                SM_Resolve = SM_CONVERTED
                Exit Function
            End If
        End If
    Next i

    ' 3) 마스터에서 유사 검색
    lv = SM_GuessLevel(s)
    If lv = 0 Then SM_Resolve = SM_SKIP: Exit Function
    SM_Resolve = SM_Fuzzy(s, lv, fullName, office, found, note)
End Function

' 접미사 규칙으로 후보 목록 생성(지역명 접두어 생략형 포함)
Private Function SM_Candidates(ByVal s As String) As Variant
    Dim rules As Variant, r As Variant, p() As String, stem As String, out() As String, n As Long, full As String
    ReDim out(0 To 40)
    rules = SM_Rules()
    For Each r In rules
        p = Split(r, "|")
        If Len(s) > Len(p(0)) And Right$(s, Len(p(0))) = p(0) Then
            stem = Left$(s, Len(s) - Len(p(0)))
            full = stem & p(1)
            out(n) = full: n = n + 1
            If Left$(stem, 2) <> "부산" Then out(n) = "부산" & full: n = n + 1
            If Right$(stem, 1) = "여" And p(1) = "고등학교" Then out(n) = Left$(stem, Len(stem) - 1) & "여자고등학교": n = n + 1
            If Right$(stem, 1) = "여" And p(1) = "중학교" Then out(n) = Left$(stem, Len(stem) - 1) & "여자중학교": n = n + 1
        End If
    Next r
    ' "학교"만 빠진 경우
    If Right$(s, 2) <> "학교" Then out(n) = s & "학교": n = n + 1
    If Left$(s, 2) <> "부산" Then out(n) = "부산" & s: n = n + 1
    ReDim Preserve out(0 To IIf(n = 0, 0, n - 1))
    SM_Candidates = out
End Function

' 같은 급 마스터에서 앞글자 + 유형이 맞는 학교 찾기
Private Function SM_Fuzzy(ByVal s As String, ByVal lv As Long, ByRef fullName As String, ByRef office As String, ByRef found As String, ByRef note As String) As Long
    Dim names As Object, k As Variant, stem As String, typ As String, hits As String, nHit As Long, first As String, dummy As Long
    Set names = SD_Names(lv)
    SM_SplitStem s, stem, typ
    If Len(stem) < 1 Then SM_Fuzzy = SM_NOTFOUND: Exit Function
    ' (1) 앞글자 일치
    For Each k In names.keys
        If SM_StemMatch(CStr(k), stem, True) And SM_TypeMatch(CStr(k), typ) Then
            nHit = nHit + 1: If nHit = 1 Then first = CStr(k)
            hits = hits & IIf(hits = "", "", ", ") & k
        End If
    Next k
    ' (2) 없으면 포함 검색
    If nHit = 0 Then
        For Each k In names.keys
            If SM_StemMatch(CStr(k), stem, False) And SM_TypeMatch(CStr(k), typ) Then
                nHit = nHit + 1: If nHit = 1 Then first = CStr(k)
                hits = hits & IIf(hits = "", "", ", ") & k
            End If
        Next k
    End If
    Select Case nHit
        Case 0
            note = "'" & s & "' → " & SD_LevelName(lv) & " 목록에 없음"
            SM_Fuzzy = SM_NOTFOUND
        Case 1
            fullName = first
            SD_Lookup first, dummy, office, found
            note = "유사 검색으로 변환: " & s & " → " & first
            SM_Fuzzy = SM_CONVERTED
        Case Else
            note = "후보 " & nHit & "개: " & hits
            SM_Fuzzy = SM_AMBIGUOUS
    End Select
End Function

' 약칭에서 앞글자(stem)와 유형 키(여/공/상/과/외/예/체/기타) 분리
Private Sub SM_SplitStem(ByVal s As String, ByRef stem As String, ByRef typ As String)
    Dim rules As Variant, r As Variant, p() As String, suf As String
    stem = s: typ = ""
    rules = SM_Rules()
    For Each r In rules
        p = Split(r, "|")
        If Len(s) > Len(p(0)) And Right$(s, Len(p(0))) = p(0) Then
            stem = Left$(s, Len(s) - Len(p(0)))
            suf = p(1)
            Exit For
        End If
    Next r
    If InStr(suf, "여자상업") > 0 Then
        typ = "여상"
    ElseIf InStr(suf, "여자") > 0 Then
        typ = "여"
    ElseIf InStr(suf, "공업") > 0 Then
        typ = "공"
    ElseIf InStr(suf, "상업") > 0 Then
        typ = "상"
    ElseIf InStr(suf, "과학") > 0 Then
        typ = "과"
    ElseIf InStr(suf, "외국어") > 0 Then
        typ = "외"
    ElseIf InStr(suf, "예술") > 0 Then
        typ = "예"
    ElseIf InStr(suf, "체육") > 0 Then
        typ = "체"
    End If
    If Right$(stem, 1) = "여" And typ = "" Then stem = Left$(stem, Len(stem) - 1): typ = "여"
    If Left$(stem, 2) = "부산" And Len(stem) > 2 Then stem = Mid$(stem, 3)
End Sub

Private Function SM_StemMatch(ByVal nm As String, ByVal stem As String, ByVal prefixOnly As Boolean) As Boolean
    Dim body As String
    body = nm
    If Left$(body, 2) = "부산" And Len(body) > 2 Then body = Mid$(body, 3)
    If prefixOnly Then
        SM_StemMatch = (Left$(body, Len(stem)) = stem)
    Else
        SM_StemMatch = (InStr(body, stem) > 0)
    End If
End Function

Private Function SM_TypeMatch(ByVal nm As String, ByVal typ As String) As Boolean
    Select Case typ
        Case "여상": SM_TypeMatch = (InStr(nm, "여자상업") > 0)
        Case "여": SM_TypeMatch = (InStr(nm, "여자") > 0)
        Case "공": SM_TypeMatch = (InStr(nm, "공업") > 0)
        Case "상": SM_TypeMatch = (InStr(nm, "상업") > 0)
        Case "과": SM_TypeMatch = (InStr(nm, "과학") > 0)
        Case "외": SM_TypeMatch = (InStr(nm, "외국어") > 0)
        Case "예": SM_TypeMatch = (InStr(nm, "예술") > 0)
        Case "체": SM_TypeMatch = (InStr(nm, "체육") > 0)
        Case Else: SM_TypeMatch = True
    End Select
End Function

' 정식명 → 약칭 (현황시트 '약칭' 열용)
Public Function SM_Abbrev(ByVal fullName As String) As String
    Dim s As String
    s = fullName
    If Right$(s, 8) = "여자상업고등학교" Then SM_Abbrev = Left$(s, Len(s) - 8) & "여상": Exit Function
    If Right$(s, 6) = "여자고등학교" Then SM_Abbrev = Left$(s, Len(s) - 6) & "여고": Exit Function
    If Right$(s, 5) = "여자중학교" Then SM_Abbrev = Left$(s, Len(s) - 5) & "여중": Exit Function
    If Right$(s, 6) = "공업고등학교" Then SM_Abbrev = Left$(s, Len(s) - 6) & "공고": Exit Function
    If Right$(s, 6) = "상업고등학교" Then SM_Abbrev = Left$(s, Len(s) - 6) & "상고": Exit Function
    If Right$(s, 6) = "과학고등학교" Then SM_Abbrev = Left$(s, Len(s) - 6) & "과고": Exit Function
    If Right$(s, 7) = "외국어고등학교" Then SM_Abbrev = Left$(s, Len(s) - 7) & "외고": Exit Function
    If Right$(s, 6) = "예술고등학교" Then SM_Abbrev = Left$(s, Len(s) - 6) & "예고": Exit Function
    If Right$(s, 6) = "체육고등학교" Then SM_Abbrev = Left$(s, Len(s) - 6) & "체고": Exit Function
    If Right$(s, 4) = "초등학교" Then SM_Abbrev = Left$(s, Len(s) - 4) & "초": Exit Function
    If Right$(s, 3) = "중학교" Then SM_Abbrev = Left$(s, Len(s) - 3) & "중": Exit Function
    If Right$(s, 4) = "고등학교" Then SM_Abbrev = Left$(s, Len(s) - 4) & "고": Exit Function
    SM_Abbrev = s
End Function
