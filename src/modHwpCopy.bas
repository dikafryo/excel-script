Attribute VB_Name = "modHwpCopy"
Option Explicit
' 한글로 복사하기
'  - 글자만 : 선택 범위의 표시값을 탭 구분 텍스트로 클립보드에 올린다 (색·음영·서식 없음, 빈 행/열 제거)
'  - 양식+  : 임시 시트에 복사해 정리(빈 행/열 제거, 채우기·글자색·조건부서식 제거, 테두리 얇은 실선, 맑은 고딕)한 뒤
'             복사한다 → 한글에서 Ctrl+V 하면 깔끔한 표로 붙는다. 원본 시트는 바꾸지 않는다.

' 양식+ 는 엑셀의 복사 기능을 쓰지 않는다: 엑셀 복사는 원본(임시 시트/숨김 창)이 사라지거나 보이지 않으면
' 클립보드를 비우기 때문. 대신 정리된 HTML 표를 직접 만들어 클립보드(HTML Format + 유니코드 텍스트)에 올린다.
#If VBA7 Then
    Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
    Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
    Private Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Function RegisterClipboardFormat Lib "user32" Alias "RegisterClipboardFormatA" (ByVal lpString As String) As Long
    Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal dest As LongPtr, ByVal src As LongPtr, ByVal n As LongPtr)
#Else
    Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function EmptyClipboard Lib "user32" () As Long
    Private Declare Function CloseClipboard Lib "user32" () As Long
    Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
    Private Declare Function RegisterClipboardFormat Lib "user32" Alias "RegisterClipboardFormatA" (ByVal lpString As String) As Long
    Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
    Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
    Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal dest As Long, ByVal src As Long, ByVal n As Long)
#End If
Private Const CF_UNICODETEXT As Long = 13
Private Const GMEM_MOVEABLE_ZERO As Long = &H42

Public Sub HwpCopyText(control As IRibbonControl)
    Dim rng As Range, r As Long, c As Long, line As String, s As String, v As String, dobj As Object
    Set rng = HC_TargetRange()
    If rng Is Nothing Then Exit Sub
    For r = 1 To rng.rows.Count
        line = ""
        For c = 1 To rng.Columns.Count
            v = HC_CellText(rng.Cells(r, c))
            If c > 1 Then line = line & vbTab
            line = line & v
        Next c
        If Len(Replace(line, vbTab, "")) > 0 Then   ' 완전히 빈 행은 건너뜀
            s = s & IIf(Len(s) > 0, vbCrLf, "") & line
        End If
    Next r
    On Error GoTo Fail
    Set dobj = CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    dobj.SetText s
    dobj.PutInClipboard
    MsgBox "글자만 복사했습니다: " & rng.rows.Count & "행 × " & rng.Columns.Count & "열 (" & rng.Address(False, False) & ")" & vbCrLf & vbCrLf & _
           "한글에서 Ctrl+V 로 붙여넣으세요. 색·음영·서식 없이 글자만 들어갑니다." & vbCrLf & _
           "(표로 만들려면 한글에서 붙여넣은 글자를 블록 지정 후 [입력 > 표 > 문자열을 표로])", vbInformation, "한글로 복사(글자만)"
    Exit Sub
Fail:
    MsgBox "클립보드에 올리지 못했습니다: " & Err.Description, vbExclamation, "한글로 복사(글자만)"
End Sub

Public Sub HwpCopyTable(control As IRibbonControl)
    Dim rng As Range, html As String, plain As String, nRows As Long, nCols As Long
    Set rng = HC_TargetRange()
    If rng Is Nothing Then Exit Sub
    On Error GoTo Fail
    html = HC_BuildHtml(rng, plain, nRows, nCols)
    If Not HC_SetClipboardHtml(html, plain) Then
        MsgBox "클립보드에 올리지 못했습니다. 다른 프로그램이 클립보드를 잡고 있으면 잠시 후 다시 시도하세요.", vbExclamation, "한글로 복사(양식+)"
        Exit Sub
    End If
    MsgBox "양식을 정리해 복사했습니다: " & nRows & "행 × " & nCols & "열 (" & rng.Address(False, False) & ")" & vbCrLf & vbCrLf & _
           "한글에서 Ctrl+V 로 붙여넣으면 표로 들어갑니다." & vbCrLf & _
           "채우기색·글자색·조건부서식·빈 행/열은 빼고, 테두리(얇은 실선)·정렬·굵게·병합·표시형식은 유지됩니다." & vbCrLf & _
           "열 너비는 글자 길이에 맞춰 자동 배분됩니다(A4 폭 기준).", vbInformation, "한글로 복사(양식+)"
    Exit Sub
Fail:
    MsgBox "복사 준비 중 오류: " & Err.Description, vbExclamation, "한글로 복사(양식+)"
End Sub


' [v2.6] 리본 [HWPX활용기능 ▶ 한글로 복사] 통합 메뉴 (tag "HwpCopyMenu")
Public Sub HwpCopyMenu(control As IRibbonControl)
    Dim choice As Long
    If TypeName(Selection) <> "Range" Then
        MsgBox "복사할 범위를 먼저 선택하세요.", vbExclamation, "한글로 복사"
        Exit Sub
    End If
    choice = UI_Choose("한글로 복사", _
        "선택 범위(" & Selection.Address(False, False) & IIf(Selection.Cells.CountLarge = 1, " → 연결된 표 전체", "") & ")를 한글에 붙여 넣을 형태를 고르세요.", _
        Array("글자만 복사   (색 · 서식 없이 탭 구분 글자, 빈 행 제외)", _
              "양식 포함 복사   (빈 행/열 · 채우기색 · 조건부서식 제거, 얇은 실선 표로)"))
    Select Case choice
        Case 1: HwpCopyText Nothing
        Case 2: HwpCopyTable Nothing
    End Select
End Sub

' [v2.6] 범위를 정리된 표(HTML Format + 텍스트)로 클립보드에 올린다 - 한글계산·표계산 결과 복사용. 반환: 성공 여부
Public Function HC_CopyRangeAsTable(ByVal rng As Range, ByRef nRows As Long, ByRef nCols As Long, Optional ByVal keepColors As Boolean = False, Optional ByVal forceFontPt As Long = 0) As Boolean
    Dim html As String, plain As String
    On Error GoTo Fail
    Set rng = HC_TrimEmpty(rng)
    If rng Is Nothing Then Exit Function
    html = HC_BuildHtml(rng, plain, nRows, nCols, keepColors, forceFontPt)
    HC_CopyRangeAsTable = HC_SetClipboardHtml(html, plain)
    Exit Function
Fail:
    HC_CopyRangeAsTable = False
End Function

' 범위를 정리된 HTML 표로 (빈 행/열 제외, 병합 반영, 정렬/굵게/글자크기 유지)
' [v2.7] 열 너비: 열마다 가장 긴 표시 글자 폭(한글 2·그 외 1 단위)으로 mm 를 정한다. 합이 한글 기본 본문 폭(150mm)을 넘으면
'        숫자 열은 그대로 두고 글 열만 줄인다(숫자가 두 줄로 갈라지지 않게). 한글은 <td width>/style width 만 따르고 <colgroup>은 무시함
'        → 한글에 붙였을 때 열이 좁아 줄바꿈되던 문제 해결. keepColors 면 채우기색·글자색을 유지(계산 결과의 불일치·부분합 표시용)
Private Function HC_BuildHtml(ByVal rng As Range, ByRef plain As String, ByRef nRows As Long, ByRef nCols As Long, Optional ByVal keepColors As Boolean = False, Optional ByVal forceFontPt As Long = 0) As String
    Dim r As Long, c As Long, k As Long, cell As Range, h As String, rowHtml As String, line As String, txt As String
    Dim emptyCol() As Boolean, nR As Long, nC As Long, ma As Range, span As String, al As String, st As String, W As Long
    Dim colU() As Long, colMm() As Single, totMm As Single, u As Long, sc As Single, clr As String, fs As Long
    Dim numCnt() As Long, nonCnt() As Long, isNum() As Boolean, fixMm As Single, txtMm As Single, colFs() As Single, ff As Single
    nR = rng.rows.Count: nC = rng.Columns.Count
    ReDim emptyCol(1 To nC): ReDim colU(1 To nC): ReDim colMm(1 To nC): ReDim numCnt(1 To nC): ReDim nonCnt(1 To nC): ReDim isNum(1 To nC): ReDim colFs(1 To nC)
    For c = 1 To nC
        emptyCol(c) = (Application.WorksheetFunction.CountA(rng.Columns(c)) = 0)
        If Not emptyCol(c) Then nCols = nCols + 1
    Next c
    ' 열 너비 계산 (병합 셀은 첫 열에만 반영하지 않고 건너뜀)
    For r = 1 To nR
        For c = 1 To nC
            If Not emptyCol(c) Then
                Set cell = rng.Cells(r, c)
                If Not cell.MergeCells Then
                    u = HC_TextUnits(HC_CellText(cell))
                    If u > colU(c) Then colU(c) = u
                    If Not IsEmpty(cell.value) And Not IsError(cell.value) Then
                        nonCnt(c) = nonCnt(c) + 1
                        If IsNumeric(cell.value) Then numCnt(c) = numCnt(c) + 1
                        If cell.Font.Size > colFs(c) Then colFs(c) = cell.Font.Size
                    End If
                End If
            End If
        Next c
    Next r
    For c = 1 To nC
        If Not emptyCol(c) Then
            If colU(c) < 3 Then colU(c) = 3
            If colU(c) > 44 Then colU(c) = 44
            ' 글꼴 크기에 비례 (10pt 기준 한글 1자 4.2mm·숫자 2.1mm)
            If forceFontPt > 0 Then ff = forceFontPt / 10 Else ff = IIf(colFs(c) > 0, colFs(c), 10) / 10
            colMm(c) = colU(c) * 2.1 * ff + 6
            totMm = totMm + colMm(c)
        End If
    Next c
    ' 한글 기본 본문 폭(A4 세로, 좌우 여백 30mm) 150mm 를 넘으면: 숫자 열(숫자 60% 이상)은 고정, 글 열만 비례 축소(최소 14mm). 그래도 넘치면 전체 축소
    For c = 1 To nC
        If Not emptyCol(c) Then
            isNum(c) = (nonCnt(c) > 0 And numCnt(c) >= nonCnt(c) * 0.6)
            If isNum(c) Then fixMm = fixMm + colMm(c) Else txtMm = txtMm + colMm(c)
        End If
    Next c
    If totMm > 150 Then
        If txtMm > 0 And fixMm < 150 - 14 Then
            sc = (150 - fixMm) / txtMm
            For c = 1 To nC
                If Not emptyCol(c) And Not isNum(c) Then
                    colMm(c) = colMm(c) * sc
                    If colMm(c) < 14 Then colMm(c) = 14
                End If
            Next c
        End If
        totMm = 0
        For c = 1 To nC
            If Not emptyCol(c) Then totMm = totMm + colMm(c)
        Next c
        If totMm > 150 Then
            sc = 150 / totMm
            For c = 1 To nC
                colMm(c) = colMm(c) * sc
            Next c
            totMm = 150
        End If
    End If
    h = "<table border=""1"" cellspacing=""0"" cellpadding=""3"" width=""" & Int(totMm * 3.78) & """ style=""border-collapse:collapse;border:1px solid #000000;font-family:'맑은 고딕';font-size:10pt"">"
    For r = 1 To nR
        If Application.WorksheetFunction.CountA(rng.rows(r)) > 0 Then
            nRows = nRows + 1
            rowHtml = "<tr>": line = "": k = 0
            For c = 1 To nC
                If Not emptyCol(c) Then
                    Set cell = rng.Cells(r, c)
                    txt = HC_CellText(cell)
                    k = k + 1
                    If k > 1 Then line = line & vbTab
                    line = line & txt
                    If cell.MergeCells And cell.MergeArea.Cells(1, 1).Address <> cell.Address Then
                        ' 병합 영역의 첫 셀이 아니면 건너뜀 (colspan/rowspan 으로 표현)
                    Else
                        span = ""
                        If cell.MergeCells Then
                            Set ma = Intersect(cell.MergeArea, rng)
                            If ma.Columns.Count > 1 Then span = span & " colspan=""" & ma.Columns.Count & """"
                            If ma.rows.Count > 1 Then span = span & " rowspan=""" & ma.rows.Count & """"
                        End If
                        Select Case cell.HorizontalAlignment
                            Case xlCenter, xlCenterAcrossSelection: al = "center"
                            Case xlRight: al = "right"
                            Case xlLeft: al = "left"
                            Case Else
                                If Not IsEmpty(cell.value) And IsNumeric(cell.value) And Not IsError(cell.value) Then al = "right" Else al = "left"
                        End Select
                        W = Int(colMm(c) * 3.78)
                        If forceFontPt > 0 Then fs = forceFontPt Else fs = cell.Font.Size
                        st = "border:1px solid #000000;width:" & Format$(colMm(c), "0.0") & "mm;text-align:" & al & ";vertical-align:middle;font-family:'맑은 고딕';font-size:" & fs & "pt" & IIf(cell.Font.Bold, ";font-weight:bold", "")
                        If Not IsEmpty(cell.value) And IsNumeric(cell.value) And Not IsError(cell.value) Then st = st & ";white-space:nowrap"
                        If keepColors Then
                            clr = HC_FillHex(cell)
                            If Len(clr) > 0 Then st = st & ";background-color:#" & clr
                            clr = HC_FontHex(cell)
                            If Len(clr) > 0 Then st = st & ";color:#" & clr
                        End If
                        rowHtml = rowHtml & "<td" & span & " width=""" & W & """ style=""" & st & """>" & IIf(Len(txt) = 0, "&nbsp;", HC_HtmlEscape(txt)) & "</td>"
                    End If
                End If
            Next c
            h = h & rowHtml & "</tr>"
            plain = plain & IIf(Len(plain) > 0, vbCrLf, "") & Trim$(line)
        End If
    Next r
    HC_BuildHtml = h & "</table>"
End Function

' 글자 폭 단위: 한글 등 전각 2, 그 외 1 (여러 줄이면 가장 긴 줄)
Private Function HC_TextUnits(ByVal s As String) As Long
    Dim ln As Variant, i As Long, n As Long
    For Each ln In Split(Replace(s, vbCrLf, vbLf), vbLf)
        n = 0
        For i = 1 To Len(ln)
            If AscW(Mid$(ln, i, 1)) > 255 Then n = n + 2 Else n = n + 1
        Next i
        If n > HC_TextUnits Then HC_TextUnits = n
    Next ln
End Function

' 채우기색 → "RRGGBB" (채우기 없음·흰색이면 "")
Private Function HC_FillHex(ByVal cell As Range) As String
    Dim v As Variant
    On Error Resume Next
    If cell.Interior.ColorIndex = xlColorIndexNone Then Exit Function
    v = cell.Interior.color
    If IsNull(v) Then Exit Function
    If CLng(v) = 16777215 Then Exit Function
    HC_FillHex = HC_HexBgr(CLng(v))
End Function

' 글자색 → "RRGGBB" (자동·검정이면 "")
Private Function HC_FontHex(ByVal cell As Range) As String
    Dim v As Variant
    On Error Resume Next
    v = cell.Font.color
    If IsNull(v) Then Exit Function
    If CLng(v) = 0 Then Exit Function
    HC_FontHex = HC_HexBgr(CLng(v))
End Function

Private Function HC_HexBgr(ByVal c As Long) As String
    HC_HexBgr = Right$("0" & Hex$(c And &HFF), 2) & Right$("0" & Hex$((c \ 256) And &HFF), 2) & Right$("0" & Hex$((c \ 65536) And &HFF), 2)
End Function

Private Function HC_HtmlEscape(ByVal s As String) As String
    s = Replace(s, "&", "&amp;")
    s = Replace(s, "<", "&lt;")
    s = Replace(s, ">", "&gt;")
    s = Replace(s, vbCrLf, "<br>")
    s = Replace(s, vbLf, "<br>")
    HC_HtmlEscape = s
End Function

' 클립보드에 HTML Format(CF_HTML, UTF-8) + 유니코드 텍스트를 올린다
Private Function HC_SetClipboardHtml(ByVal html As String, ByVal plain As String) As Boolean
    Dim pre As String, post As String, hdr As String, full As String
    Dim hdrLen As Long, startFrag As Long, endFrag As Long, endHtml As Long
    Dim b() As Byte, n As Long, cfHtml As Long, i As Long, opened As Long
#If VBA7 Then
    Dim hMem As LongPtr, p As LongPtr
#Else
    Dim hMem As Long, p As Long
#End If
    pre = "<html><body><!--StartFragment-->"
    post = "<!--EndFragment--></body></html>"
    hdr = "Version:0.9" & vbCrLf & "StartHTML:0000000000" & vbCrLf & "EndHTML:0000000000" & vbCrLf & _
          "StartFragment:0000000000" & vbCrLf & "EndFragment:0000000000" & vbCrLf
    hdrLen = Len(hdr)   ' 헤더는 ASCII
    startFrag = hdrLen + HC_Utf8Len(pre)
    endFrag = startFrag + HC_Utf8Len(html)
    endHtml = endFrag + HC_Utf8Len(post)
    hdr = "Version:0.9" & vbCrLf & "StartHTML:" & Format$(hdrLen, "0000000000") & vbCrLf & "EndHTML:" & Format$(endHtml, "0000000000") & vbCrLf & _
          "StartFragment:" & Format$(startFrag, "0000000000") & vbCrLf & "EndFragment:" & Format$(endFrag, "0000000000") & vbCrLf
    full = hdr & pre & html & post
    b = HC_Utf8Bytes(full)          ' 끝에 0 바이트 포함
    For i = 1 To 10
        opened = OpenClipboard(0)
        If opened <> 0 Then Exit For
        HC_SleepMs 50
    Next i
    If opened = 0 Then Exit Function
    EmptyClipboard
    ' HTML Format
    cfHtml = RegisterClipboardFormat("HTML Format")
    n = UBound(b) - LBound(b) + 1
    hMem = GlobalAlloc(GMEM_MOVEABLE_ZERO, n)
    p = GlobalLock(hMem)
    CopyMemory p, VarPtr(b(LBound(b))), n
    GlobalUnlock hMem
    SetClipboardData cfHtml, hMem
    ' 유니코드 텍스트 (탭 구분)
    n = LenB(plain) + 2
    hMem = GlobalAlloc(GMEM_MOVEABLE_ZERO, n)
    p = GlobalLock(hMem)
    CopyMemory p, StrPtr(plain), LenB(plain)
    GlobalUnlock hMem
    SetClipboardData CF_UNICODETEXT, hMem
    CloseClipboard
    HC_SetClipboardHtml = True
End Function

Private Sub HC_SleepMs(ByVal ms As Long)
    Dim T As Single: T = Timer
    Do While Timer - T < ms / 1000!
        DoEvents
    Loop
End Sub

' UTF-8 바이트 (BOM 없음, 끝에 0 바이트 추가)
Private Function HC_Utf8Bytes(ByVal s As String) As Byte()
    Dim st As Object, b() As Byte, n As Long
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2: st.Charset = "utf-8": st.Open
    st.WriteText s
    st.Position = 0: st.Type = 1: st.Position = 3   ' BOM 3바이트 건너뜀
    b = st.Read
    st.Close
    n = UBound(b) - LBound(b) + 1
    ReDim Preserve b(LBound(b) To LBound(b) + n)   ' 0 종료 바이트
    HC_Utf8Bytes = b
End Function

Private Function HC_Utf8Len(ByVal s As String) As Long
    Dim b() As Byte
    If Len(s) = 0 Then Exit Function
    b = HC_Utf8Bytes(s)
    HC_Utf8Len = UBound(b) - LBound(b)   ' 종료 0 제외
End Function

' 대상 범위: 여러 셀 선택이면 그 범위, 1셀이면 연결 영역. 빈 행/열을 잘라낸 사각 범위
Private Function HC_TargetRange() As Range
    Dim rng As Range
    If TypeName(Selection) <> "Range" Then
        MsgBox "복사할 범위를 먼저 선택하세요.", vbExclamation, "한글로 복사"
        Exit Function
    End If
    Set rng = Selection
    If rng.Cells.CountLarge = 1 Then Set rng = rng.CurrentRegion
    If rng.Areas.Count > 1 Then Set rng = rng.Areas(1)
    If rng.rows.Count = rng.Worksheet.rows.Count Or rng.Columns.Count = rng.Worksheet.Columns.Count Then
        If Intersect(rng, rng.Worksheet.UsedRange) Is Nothing Then Exit Function
        Set rng = Intersect(rng, rng.Worksheet.UsedRange)
    End If
    Set rng = HC_TrimEmpty(rng)
    If rng Is Nothing Then
        MsgBox "선택한 범위가 비어 있습니다.", vbExclamation, "한글로 복사"
        Exit Function
    End If
    Set HC_TargetRange = rng
End Function

' 가장자리의 빈 행/열 제거
Private Function HC_TrimEmpty(ByVal rng As Range) As Range
    Dim r1 As Long, r2 As Long, c1 As Long, c2 As Long
    r1 = 1: r2 = rng.rows.Count: c1 = 1: c2 = rng.Columns.Count
    Do While r1 <= r2
        If Application.WorksheetFunction.CountA(rng.rows(r1)) > 0 Then Exit Do
        r1 = r1 + 1
    Loop
    If r1 > r2 Then Exit Function
    Do While r2 > r1
        If Application.WorksheetFunction.CountA(rng.rows(r2)) > 0 Then Exit Do
        r2 = r2 - 1
    Loop
    Do While c1 <= c2
        If Application.WorksheetFunction.CountA(rng.Columns(c1)) > 0 Then Exit Do
        c1 = c1 + 1
    Loop
    Do While c2 > c1
        If Application.WorksheetFunction.CountA(rng.Columns(c2)) > 0 Then Exit Do
        c2 = c2 - 1
    Loop
    Set HC_TrimEmpty = rng.Worksheet.Range(rng.Cells(r1, c1), rng.Cells(r2, c2))
End Function

Private Function HC_CellText(ByVal c As Range) As String
    Dim T As String
    If IsError(c.value) Then
        T = ""
    Else
        T = c.text
        If Len(T) = 0 And Not IsEmpty(c.value) Then T = CStr(c.value)
    End If
    T = Replace(Replace(T, vbCrLf, " "), vbLf, " ")
    HC_CellText = Trim$(T)
End Function

