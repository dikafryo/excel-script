Attribute VB_Name = "modSchoolCheck"
Option Explicit
' =============================================================================================
'  수신처 자동체크 (업무관리시스템 문서관리카드 > [수신자지정] 창)
'  - 엑셀 학교명(여러 칸을 선택했으면 그 칸들, 아니면 활성 시트 A2부터 아래로)을 읽어
'    공용그룹 탭의 '전체-고등학교 / 전체-중학교 / 전체-초등학교 / 전체-유치원 / 전체-기타학교' 중
'    필요한 그룹만 + 를 눌러 펼친 뒤, 같은 이름의 학교만 체크하고 [>>]로 수신자에 추가한다.
'  - 마우스·키보드 입력 없이 창 안의 웹 문서(DOM)를 직접 다룬다. 화면 위치와 상관없이 동작.
'  - 이미 수신자에 있는 학교는 건너뛴다(중복 경고창 방지). [확인]은 사용자가 직접 누른다.
'  - A열 모드에서는 B열에 학교별 결과(추가됨/이미 수신자/못 찾음)를 적는다.
' =============================================================================================
Private Type SC_GUID
    D1 As Long
    D2 As Integer
    D3 As Integer
    D4(0 To 7) As Byte
End Type
#If VBA7 Then
Private Declare PtrSafe Function SC_FindWindowEx Lib "user32" Alias "FindWindowExW" (ByVal hParent As LongPtr, ByVal hAfter As LongPtr, ByVal lpClass As LongPtr, ByVal lpTitle As LongPtr) As LongPtr
Private Declare PtrSafe Function SC_GetWindowText Lib "user32" Alias "GetWindowTextW" (ByVal h As LongPtr, ByVal buf As LongPtr, ByVal n As Long) As Long
Private Declare PtrSafe Function SC_GetClassName Lib "user32" Alias "GetClassNameW" (ByVal h As LongPtr, ByVal buf As LongPtr, ByVal n As Long) As Long
Private Declare PtrSafe Function SC_EnumChildWindows Lib "user32" Alias "EnumChildWindows" (ByVal hParent As LongPtr, ByVal cb As LongPtr, ByVal lp As LongPtr) As Long
Private Declare PtrSafe Function SC_RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageW" (ByVal s As LongPtr) As Long
Private Declare PtrSafe Function SC_SendMessageTimeout Lib "user32" Alias "SendMessageTimeoutW" (ByVal h As LongPtr, ByVal m As Long, ByVal W As LongPtr, ByVal L As LongPtr, ByVal f As Long, ByVal T As Long, ByRef r As LongPtr) As LongPtr
Private Declare PtrSafe Function SC_ObjectFromLresult Lib "oleacc" Alias "ObjectFromLresult" (ByVal L As LongPtr, ByRef riid As SC_GUID, ByVal W As LongPtr, ByRef obj As Object) As Long
Private Declare PtrSafe Sub SC_Sleep Lib "kernel32" Alias "Sleep" (ByVal ms As Long)
Private mIE As LongPtr
#Else
Private Declare Function SC_FindWindowEx Lib "user32" Alias "FindWindowExW" (ByVal hParent As Long, ByVal hAfter As Long, ByVal lpClass As Long, ByVal lpTitle As Long) As Long
Private Declare Function SC_GetWindowText Lib "user32" Alias "GetWindowTextW" (ByVal h As Long, ByVal buf As Long, ByVal n As Long) As Long
Private Declare Function SC_GetClassName Lib "user32" Alias "GetClassNameW" (ByVal h As Long, ByVal buf As Long, ByVal n As Long) As Long
Private Declare Function SC_EnumChildWindows Lib "user32" Alias "EnumChildWindows" (ByVal hParent As Long, ByVal cb As Long, ByVal lp As Long) As Long
Private Declare Function SC_RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageW" (ByVal s As Long) As Long
Private Declare Function SC_SendMessageTimeout Lib "user32" Alias "SendMessageTimeoutW" (ByVal h As Long, ByVal m As Long, ByVal w As Long, ByVal l As Long, ByVal f As Long, ByVal t As Long, ByRef r As Long) As Long
Private Declare Function SC_ObjectFromLresult Lib "oleacc" Alias "ObjectFromLresult" (ByVal l As Long, ByRef riid As SC_GUID, ByVal w As Long, ByRef obj As Object) As Long
Private Declare Sub SC_Sleep Lib "kernel32" Alias "Sleep" (ByVal ms As Long)
Private mIE As Long
#End If

Private Const SC_TITLE As String = "수신처 자동체크"
Private Const SC_PREFIX As String = "frmRecpMainOrgGroup"

' ------------------------------------------------------------------ 리본
Public Sub SC_Ribbon(control As IRibbonControl)
    Dim dry As Boolean
    dry = (control.tag = "preview")
    MsgBox SC_RunText(dry), vbInformation, SC_TITLE & IIf(dry, " - 미리보기", "")
End Sub

' ------------------------------------------------------------------ 창 찾기
#If VBA7 Then
Private Function SC_WinText(ByVal h As LongPtr) As String
#Else
Private Function SC_WinText(ByVal h As Long) As String
#End If
    Dim b As String, n As Long
    b = String$(512, vbNullChar)
    n = SC_GetWindowText(h, StrPtr(b), 512)
    SC_WinText = Left$(b, n)
End Function

#If VBA7 Then
Private Function SC_ClsName(ByVal h As LongPtr) As String
#Else
Private Function SC_ClsName(ByVal h As Long) As String
#End If
    Dim b As String, n As Long
    b = String$(256, vbNullChar)
    n = SC_GetClassName(h, StrPtr(b), 256)
    SC_ClsName = Left$(b, n)
End Function

#If VBA7 Then
Public Function SC_EnumCb(ByVal h As LongPtr, ByVal lp As LongPtr) As Long
#Else
Public Function SC_EnumCb(ByVal h As Long, ByVal lp As Long) As Long
#End If
    If SC_ClsName(h) = "Internet Explorer_Server" Then
        mIE = h
        SC_EnumCb = 0
    Else
        SC_EnumCb = 1
    End If
End Function

' [수신자지정] 웹 페이지 대화 상자의 문서(DOM). 없으면 Nothing
Private Function SC_GetDoc() As Object
#If VBA7 Then
    Dim h As LongPtr, topW As LongPtr, res As LongPtr
#Else
    Dim h As Long, topW As Long, res As Long
#End If
    Dim g As SC_GUID, o As Object
    h = SC_FindWindowEx(0, 0, 0, 0)
    Do While h <> 0
        If InStr(SC_WinText(h), "수신자지정") > 0 Then topW = h: Exit Do
        h = SC_FindWindowEx(0, h, 0, 0)
    Loop
    If topW = 0 Then Exit Function
    mIE = 0
    SC_EnumChildWindows topW, AddressOf SC_EnumCb, 0
    If mIE = 0 Then Exit Function
    SC_SendMessageTimeout mIE, SC_RegisterWindowMessage(StrPtr("WM_HTML_GETOBJECT")), 0, 0, 2, 5000, res
    If res = 0 Then Exit Function
    g.D1 = &H626FC520: g.D2 = &HA41E: g.D3 = &H11CF          ' IHTMLDocument2
    g.D4(0) = &HA7: g.D4(1) = &H31: g.D4(2) = &H0: g.D4(3) = &HA0
    g.D4(4) = &HC9: g.D4(5) = &H8: g.D4(6) = &H26: g.D4(7) = &H37
    On Error Resume Next
    SC_ObjectFromLresult res, g, 0, o
    On Error GoTo 0
    Set SC_GetDoc = o
End Function

Private Sub SC_Wait(ByVal ms As Long)
    Dim T As Single
    T = Timer
    Do
        DoEvents
        SC_Sleep 50
    Loop While (Timer - T) * 1000 < ms And Timer >= T
End Sub

' ------------------------------------------------------------------ 이름 정리
Private Function SC_Norm(ByVal s As String) As String
    s = Replace(s, ChrW(160), " ")
    s = Replace(s, vbTab, "")
    s = Replace(s, " ", "")
    SC_Norm = Trim$(s)
End Function

' 학교급 → 공용그룹 이름 앞부분
Private Function SC_GroupOf(ByVal nm As String) As String
    If Right$(nm, 4) = "고등학교" Then
        SC_GroupOf = "전체-고등학교"
    ElseIf Right$(nm, 3) = "중학교" Then
        SC_GroupOf = "전체-중학교"
    ElseIf Right$(nm, 4) = "초등학교" Then
        SC_GroupOf = "전체-초등학교"
    ElseIf Right$(nm, 3) = "유치원" Then
        SC_GroupOf = "전체-유치원"
    Else
        SC_GroupOf = "전체-기타학교"
    End If
End Function

' 표시 이름 '○○교육지원청 △△중학교' → '△△중학교'
Private Function SC_ShortName(ByVal label As String) As String
    Dim p As Long
    label = Trim$(Replace(label, ChrW(160), " "))
    p = InStrRev(label, " ")
    If p > 0 Then label = Mid$(label, p + 1)
    SC_ShortName = SC_Norm(label)
End Function

' ------------------------------------------------------------------ 그룹 펼치기
' 반환: 그룹 코드(없으면 ""). 펼친 뒤 학교 칸이 다 나타날 때까지 최대 15초 기다린다.
Private Function SC_ExpandGroup(ByVal doc As Object, ByVal head As String) As String
    Dim spans As Object, sp As Object, i As Long, sid As String, code As String
    Dim img As Object, box As Object, n As Long, lastN As Long, stable As Long, T As Single
    Set spans = doc.getElementsByTagName("span")
    For i = 0 To spans.Length - 1
        Set sp = spans.item(i)
        sid = ""
        On Error Resume Next
        sid = CStr(sp.id)
        On Error GoTo 0
        If Left$(sid, Len(SC_PREFIX)) = SC_PREFIX And Right$(sid, 4) = "Font" Then
            If Left$(SC_Norm(sp.innerText), Len(head)) = head Then
                code = Mid$(sid, Len(SC_PREFIX) + 1, Len(sid) - Len(SC_PREFIX) - 4)
                Exit For
            End If
        End If
    Next i
    If code = "" Then Exit Function
    Set img = doc.getElementById(SC_PREFIX & code & "Img")
    If Not img Is Nothing Then
        If InStr(1, CStr(img.src), "plus", vbTextCompare) > 0 Then img.parentElement.Click
    End If
    T = Timer
    Do
        SC_Wait 300
        Set box = doc.getElementById(SC_PREFIX & code & "Div")
        n = 0
        If Not box Is Nothing Then n = box.getElementsByTagName("input").Length
        If n > 1 And n = lastN Then stable = stable + 1 Else stable = 0
        lastN = n
    Loop While stable < 2 And (Timer - T) < 15 And Timer >= T
    SC_ExpandGroup = code
End Function

' ------------------------------------------------------------------ 본체
Public Function SC_RunText(ByVal dry As Boolean) As String
    Dim ws As Object, names As Collection, cellList As Collection, c As Object, rg As Object, v As String, r As Long, lastR As Long
    Dim colMode As Boolean, doc As Object, out As String
    Set names = New Collection: Set cellList = New Collection

    On Error Resume Next
    Set ws = ActiveSheet
    On Error GoTo 0
    If ws Is Nothing Then SC_RunText = "학교 목록이 있는 시트를 먼저 선택하세요.": Exit Function

    ' 1) 엑셀 학교명: 여러 칸을 선택했으면 그 칸들, 아니면 A2부터 아래로
    If TypeName(Selection) = "Range" Then
        If Selection.Cells.Count > 1 Then
            Set rg = Intersect(Selection, ws.UsedRange)
            If Not rg Is Nothing Then
                For Each c In rg.Cells
                    v = SC_Norm(CStr(c.value))
                    If v <> "" Then names.Add v: cellList.Add c
                Next c
            End If
        End If
    End If
    If names.Count = 0 Then
        colMode = True
        lastR = ws.Cells(ws.rows.Count, 1).End(-4162).row        ' xlUp
        For r = 2 To lastR
            v = SC_Norm(CStr(ws.Cells(r, 1).value))
            If v <> "" Then names.Add v: cellList.Add ws.Cells(r, 1)
        Next r
    End If
    If names.Count = 0 Then
        SC_RunText = "학교명이 없습니다." & vbCrLf & "A2부터 아래로 학교명을 적거나, 학교명 칸들을 선택한 뒤 실행하세요."
        Exit Function
    End If

    ' 2) 수신자지정 창
    Set doc = SC_GetDoc()
    If doc Is Nothing Then
        SC_RunText = "엑셀 학교 " & names.Count & "개를 읽었습니다." & vbCrLf & vbCrLf & _
                     "'수신자지정' 창을 찾지 못했습니다." & vbCrLf & _
                     "업무관리 기안 화면에서 [수신자지정]을 눌러 창을 열어 둔 뒤 다시 실행하세요."
        Exit Function
    End If

    ' 3) 공용그룹 탭으로
    Dim tabLi As Object, tabA As Object
    On Error Resume Next
    Set tabLi = doc.getElementById("RecpMainOrgGroup-cur")
    If Not tabLi Is Nothing Then
        If InStr(CStr(tabLi.className), "Lcurrent") = 0 Then
            Set tabA = doc.getElementById("RecpMainOrgGroup-tab")
            If Not tabA Is Nothing Then tabA.Click: SC_Wait 800
        End If
    End If
    On Error GoTo 0

    ' 4) 필요한 그룹만 펼치고 학교 칸 모으기 (이름 → 체크박스)
    Dim groups As Object, gname As Variant, code As String, box As Object, inps As Object, inp As Object
    Dim i As Long, f() As String, lab As String, key As String, map As Object, notFoundGroup As String
    Set groups = CreateObject("Scripting.Dictionary")
    Set map = CreateObject("Scripting.Dictionary")
    For Each gname In names
        If Not groups.Exists(SC_GroupOf(CStr(gname))) Then groups.Add SC_GroupOf(CStr(gname)), ""
    Next gname
    For Each gname In groups.keys
        Application.StatusBar = SC_TITLE & ": '" & gname & "' 펼치는 중..."
        code = SC_ExpandGroup(doc, CStr(gname))
        If code = "" Then
            notFoundGroup = notFoundGroup & "  " & gname & vbCrLf
        Else
            Set box = doc.getElementById(SC_PREFIX & code & "Div")
            Set inps = box.getElementsByTagName("input")
            For i = 0 To inps.Length - 1
                Set inp = inps.item(i)
                f = Split(CStr(inp.value), "|")
                If UBound(f) >= 3 Then
                    If f(3) = "2" Then                                    ' 1=그룹 칸, 2=학교 칸
                        lab = ""
                        On Error Resume Next
                        lab = CStr(inp.NextSibling.innerText)
                        On Error GoTo 0
                        If lab = "" And UBound(f) >= 11 Then lab = f(11)
                        key = SC_ShortName(lab)
                        If key <> "" And Not map.Exists(key) Then map.Add key, inp
                    End If
                End If
            Next i
        End If
    Next gname
    Application.StatusBar = False

    ' 5) 이미 수신자 목록에 있는 학교
    Dim already As Object, sel As Object
    Set already = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    Set sel = doc.getElementById("IDC_HDUAGENT_RECVDEPT")
    If Not sel Is Nothing Then
        For i = 0 To sel.options.Length - 1
            key = SC_ShortName(CStr(sel.options.item(i).text))
            ' 수신자 목록은 기관장으로 표시됨: '△△중학교장' → '△△중학교'
            If Right$(key, 3) = "학교장" Or Right$(key, 4) = "유치원장" Then key = Left$(key, Len(key) - 1)
            If key <> "" And Not already.Exists(key) Then already.Add key, True
        Next i
    End If
    On Error GoTo 0
    Dim before As Long: before = already.Count

    ' 6) 체크
    Dim k As Long, nOk As Long, nAl As Long, nMiss As Long, miss As String, st As String, hint As String, mk As Variant
    For k = 1 To names.Count
        v = names(k)
        If already.Exists(v) Then
            st = "이미 수신자": nAl = nAl + 1
        ElseIf map.Exists(v) Then
            If Not dry Then
                If Not map(v).Checked Then map(v).Click
            End If
            st = IIf(dry, "추가 예정", "추가됨"): nOk = nOk + 1
        Else
            hint = ""
            For Each mk In map.keys
                If InStr(CStr(mk), Left$(v, 2)) = 1 And SC_GroupOf(CStr(mk)) = SC_GroupOf(v) And Len(hint) < 30 Then hint = hint & IIf(hint = "", "", ", ") & mk
            Next mk
            st = "못 찾음" & IIf(hint <> "", " (비슷한 이름: " & hint & ")", "")
            nMiss = nMiss + 1
            If nMiss <= 15 Then miss = miss & "  - " & v & IIf(hint <> "", "  → " & hint, "") & vbCrLf
        End If
        If colMode Then SC_Mark ws, cellList(k), st
    Next k

    ' 7) [>>] 눌러 수신자에 추가
    Dim added As Long, aList As Object, a As Object
    If Not dry And nOk > 0 Then
        Set aList = doc.getElementsByTagName("a")
        For i = 0 To aList.Length - 1
            Set a = aList.item(i)
            If InStr(CStr(a.outerHTML), "fncAdd()") > 0 Then a.Click: Exit For
        Next i
        SC_Wait 1500
        On Error Resume Next
        Set sel = doc.getElementById("IDC_HDUAGENT_RECVDEPT")
        added = sel.options.Length - before
        On Error GoTo 0
    End If

    out = "엑셀 학교: " & names.Count & "개" & IIf(colMode, " (A열)", " (선택 범위)") & vbCrLf & _
          "펼친 그룹: " & Join(groups.keys, ", ") & vbCrLf & vbCrLf & _
          IIf(dry, "추가 예정: ", "체크·추가: ") & nOk & "개" & vbCrLf & _
          "이미 수신자: " & nAl & "개" & vbCrLf & _
          "못 찾음: " & nMiss & "개" & vbCrLf
    If Not dry And nOk > 0 Then out = out & "수신자 목록에 늘어난 수: " & added & "개" & vbCrLf
    If notFoundGroup <> "" Then out = out & vbCrLf & "공용그룹에 없는 그룹:" & vbCrLf & notFoundGroup
    If miss <> "" Then out = out & vbCrLf & "못 찾은 학교" & IIf(nMiss > 15, " (앞 15개)", "") & ":" & vbCrLf & miss
    If colMode Then out = out & vbCrLf & "학교별 결과는 B열에 적었습니다."
    If Not dry Then out = out & vbCrLf & "※ 수신자 목록을 확인한 뒤 [확인]은 직접 누르세요."
    SC_RunText = out
End Function

' A열 모드: B열에 결과 (B1 머리글이 비어 있거나 '수신처 결과'일 때만)
Private Sub SC_Mark(ByVal ws As Object, ByVal c As Object, ByVal st As String)
    Dim h As String
    h = Trim$(CStr(ws.Cells(1, 2).value))
    If h <> "" And h <> "수신처 결과" Then Exit Sub
    If h = "" Then ws.Cells(1, 2).value = "수신처 결과"
    c.Offset(0, 1).value = st
    If Left$(st, 3) = "못 찾" Then
        c.Offset(0, 1).Font.color = RGB(192, 0, 0)
    Else
        c.Offset(0, 1).Font.color = RGB(0, 0, 0)
    End If
End Sub


