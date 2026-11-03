Attribute VB_Name = "modStaffImport"
Option Explicit
'=====================================================================
' modStaffImport : [부서명단 파일] - 성명·팀/부서 명단 파일 불러오기
'   지원: XLSX/XLS/XLSM (엑셀), DOCX/DOC (Word 자동화, 표), HWPX (압축 해제 후 XML 표 파싱),
'         HWP (한컴 오피스 자동화로 HWPX 변환 후 파싱 - 한글이 설치되어 있어야 함)
'   규칙: 표에서 '성명/이름' 열과 '팀/부서/소속' 열을 머리글로 찾음 → 없으면 1열=성명, 2열=팀
'         '직위/직급' 열이 있으면 직위도 입력. 부서담당자 시트 5행부터 기록(기존 내용 교체)
'=====================================================================

Public Sub 부서명단파일선택()
    Dim f As Variant, rows As Collection, ext As String
    f = Application.GetOpenFilename("명단 파일 (*.xlsx;*.xls;*.xlsm;*.docx;*.doc;*.hwpx;*.hwp),*.xlsx;*.xls;*.xlsm;*.docx;*.doc;*.hwpx;*.hwp", , "부서명단 파일 선택")
    If VarType(f) = vbBoolean Then Exit Sub
    ext = LCase(Mid(f, InStrRev(f, ".") + 1))
    On Error GoTo EH
    SpeedOn "명단 파일 읽는 중..."
    Select Case ext
        Case "xlsx", "xls", "xlsm": Set rows = ReadExcelRows(CStr(f))
        Case "docx", "doc": Set rows = ReadWordRows(CStr(f))
        Case "hwpx": Set rows = ReadHwpxRows(CStr(f))
        Case "hwp": Set rows = ReadHwpRows(CStr(f))
        Case Else: Err.Raise 1000, , "지원하지 않는 형식: " & ext
    End Select
    SpeedOff
    If rows Is Nothing Then Exit Sub
    If rows.Count = 0 Then
        MsgBox "파일에서 표(명단)를 찾지 못했습니다." & vbCrLf & "성명·팀/부서 열이 있는 표 형태의 파일을 선택하세요.", vbExclamation
        Exit Sub
    End If
    WriteStaffRows rows
    Exit Sub
EH:
    SpeedOff
    MsgBox "부서명단 파일 오류: " & Err.Description, vbCritical
End Sub

'---------------------------------------------------------------- 시트 기록
Private Sub WriteStaffRows(rows As Collection)
    Dim cName As Long, cTeam As Long, cPos As Long, hdrIdx As Long, i As Long, row As Variant, j As Long, T As String
    Dim ws As Worksheet, r As Long, n As Long, nm As String, seen As Object
    '머리글 행 탐색 (처음 20행)
    cName = 0: cTeam = 0: cPos = 0: hdrIdx = 0
    For i = 1 To IIf(rows.Count < 20, rows.Count, 20)
        row = rows(i)
        For j = LBound(row) To UBound(row)
            T = Replace(CleanText(row(j)), " ", "")
            If cName = 0 Then If T = "성명" Or T = "이름" Or T = "담당자" Or T = "성명(이름)" Then cName = j: hdrIdx = i
            If cTeam = 0 Then If InStr(T, "팀") > 0 Or InStr(T, "부서") > 0 Or InStr(T, "소속") > 0 Then If Len(T) <= 8 Then cTeam = j: hdrIdx = i
            If cPos = 0 Then If T = "직위" Or T = "직급" Or T = "직책" Then cPos = j
        Next j
        If cName > 0 And cTeam > 0 Then Exit For
    Next i
    If cName = 0 Then
        '머리글 없음: 1열 성명, 2열 팀
        row = rows(1)
        cName = LBound(row): cTeam = LBound(row) + 1: hdrIdx = 0
        If cTeam > UBound(row) Then cTeam = 0
    ElseIf cTeam = 0 Then
        row = rows(hdrIdx)
        cTeam = IIf(cName + 1 <= UBound(row), cName + 1, 0)
    End If

    If Not SheetExists(SH_STAFF) Then RefreshStaffSheet False
    Set ws = BW().Worksheets(SH_STAFF)
    '기존 명단(5행~, A:E) 지우기
    r = lastRow(ws, 1)
    If r >= 5 Then ws.Range(ws.Cells(5, 1), ws.Cells(r, 5)).ClearContents
    Set seen = CreateObject("Scripting.Dictionary")
    r = 4
    For i = hdrIdx + 1 To rows.Count
        row = rows(i)
        If cName >= LBound(row) And cName <= UBound(row) Then
            nm = CleanText(row(cName))
            nm = Replace(nm, " ", "")
            If nm <> "" And Len(nm) <= 10 And Not seen.Exists(nm) And nm <> "성명" And nm <> "이름" Then
                seen.Add nm, 1
                r = r + 1: n = n + 1
                ws.Cells(r, 1).Value2 = nm
                If cTeam > 0 Then If cTeam <= UBound(row) Then ws.Cells(r, 2).Value2 = CleanText(row(cTeam))
                If cPos > 0 Then If cPos <= UBound(row) Then ws.Cells(r, 3).Value2 = CleanText(row(cPos))
            End If
        End If
    Next i
    If n = 0 Then
        MsgBox "성명을 찾지 못했습니다. (성명/이름 열 확인)", vbExclamation
        Exit Sub
    End If
    StyleBody ws.Range(ws.Cells(5, 1), ws.Cells(r, 5))
    ws.Range(ws.Cells(5, 2), ws.Cells(r, 4)).Interior.color = RGB(255, 255, 224)
    ws.Activate
    If MsgBox(n & "명을 불러왔습니다 (성명·팀/부서" & IIf(cPos > 0, "·직위", "") & ")." & vbCrLf & vbCrLf & _
              "지금 [적용] 하여 메인 보드에 반영할까요?", vbYesNo + vbQuestion, "부서명단 파일") = vbYes Then 적용
End Sub

'---------------------------------------------------------------- 엑셀
Private Function ReadExcelRows(ByVal path As String) As Collection
    Dim wb As Workbook, ws As Worksheet, col As New Collection, arr As Variant, r As Long, c As Long, v() As Variant, best As Worksheet, bestN As Long
    Set wb = Workbooks.Open(path, ReadOnly:=True, UpdateLinks:=0)
    '데이터가 가장 많은 시트 사용
    For Each ws In wb.Worksheets
        If Application.WorksheetFunction.CountA(ws.Cells) > bestN Then bestN = Application.WorksheetFunction.CountA(ws.Cells): Set best = ws
    Next ws
    If Not best Is Nothing Then
        arr = best.UsedRange.Value2
        If IsArray(arr) Then
            For r = LBound(arr, 1) To UBound(arr, 1)
                ReDim v(1 To UBound(arr, 2) - LBound(arr, 2) + 1)
                For c = LBound(arr, 2) To UBound(arr, 2)
                    v(c - LBound(arr, 2) + 1) = arr(r, c)
                Next c
                col.Add v
            Next r
        End If
    End If
    wb.Close False
    Set ReadExcelRows = col
End Function

'---------------------------------------------------------------- Word (표)
Private Function ReadWordRows(ByVal path As String) As Collection
    Dim wa As Object, doc As Object, tB As Object, col As New Collection, r As Long, c As Long, v() As Variant, s As String
    On Error Resume Next
    Set wa = CreateObject("Word.Application")
    On Error GoTo 0
    If wa Is Nothing Then Err.Raise 1001, , "Word 를 열 수 없습니다. DOCX 는 Word 가 설치되어 있어야 합니다."
    wa.Visible = False
    Set doc = wa.Documents.Open(path, ReadOnly:=True, AddToRecentFiles:=False)
    For Each tB In doc.Tables
        For r = 1 To tB.rows.Count
            ReDim v(1 To tB.Columns.Count)
            For c = 1 To tB.Columns.Count
                On Error Resume Next
                s = tB.cell(r, c).Range.text
                On Error GoTo 0
                s = Replace(Replace(s, Chr(13), ""), Chr(7), "")
                v(c) = s
            Next c
            col.Add v
        Next r
    Next tB
    doc.Close False
    wa.Quit
    Set ReadWordRows = col
End Function

'---------------------------------------------------------------- HWPX (zip 안의 Contents/section*.xml 표 파싱)
Private Function ReadHwpxRows(ByVal path As String) As Collection
    Dim fso As Object, tmp As String, zipPath As String, Sh As Object, col As New Collection, f As Object, xml As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    tmp = fso.GetSpecialFolder(2) & "\hwpx_" & Format(Now, "hhmmss")
    fso.CreateFolder tmp
    zipPath = tmp & "\src.zip"
    fso.CopyFile path, zipPath
    Set Sh = CreateObject("Shell.Application")
    Sh.Namespace(CVar(tmp)).CopyHere Sh.Namespace(CVar(zipPath)).items, 4 + 16 + 1024
    Dim t0 As Single: t0 = Timer
    Do While Not fso.FolderExists(tmp & "\Contents") And Timer - t0 < 10: DoEvents: Loop
    If Not fso.FolderExists(tmp & "\Contents") Then Err.Raise 1002, , "HWPX 압축 해제 실패"
    For Each f In fso.GetFolder(tmp & "\Contents").files
        If LCase(Left(f.name, 7)) = "section" And LCase(Right(f.name, 4)) = ".xml" Then
            xml = ReadUtf8(f.path)
            ParseHwpxTables xml, col
        End If
    Next f
    On Error Resume Next
    fso.DeleteFolder tmp, True
    On Error GoTo 0
    Set ReadHwpxRows = col
End Function

Private Sub ParseHwpxTables(ByVal xml As String, col As Collection)
    Dim re As Object, mTr As Object, mTc As Object, mT As Object, tr As Variant, tc As Variant, v() As Variant, n As Long, s As String, m As Variant
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True: re.IgnoreCase = True
    re.Pattern = "<hp:tr\b[\s\S]*?</hp:tr>"
    Set mTr = re.Execute(xml)
    Dim reTc As Object, reT As Object, reTag As Object
    Set reTc = CreateObject("VBScript.RegExp"): reTc.Global = True: reTc.Pattern = "<hp:tc\b[\s\S]*?</hp:tc>"
    Set reT = CreateObject("VBScript.RegExp"): reT.Global = True: reT.Pattern = "<hp:t\b[^>]*>([\s\S]*?)</hp:t>"
    Set reTag = CreateObject("VBScript.RegExp"): reTag.Global = True: reTag.Pattern = "<[^>]+>"
    For Each tr In mTr
        Set mTc = reTc.Execute(tr.value)
        If mTc.Count > 0 Then
            ReDim v(1 To mTc.Count)
            n = 0
            For Each tc In mTc
                n = n + 1: s = ""
                Set mT = reT.Execute(tc.value)
                For Each m In mT
                    s = s & reTag.Replace(m.SubMatches(0), "")
                Next m
                s = Replace(Replace(Replace(s, "&amp;", "&"), "&lt;", "<"), "&gt;", ">")
                v(n) = s
            Next tc
            col.Add v
        End If
    Next tr
End Sub

Private Function ReadUtf8(ByVal path As String) As String
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2: st.Charset = "utf-8": st.Open: st.LoadFromFile path
    ReadUtf8 = st.ReadText
    st.Close
End Function

'---------------------------------------------------------------- HWP (한컴 자동화 → HWPX 변환)
Private Function ReadHwpRows(ByVal path As String) As Collection
    Dim hwp As Object, tmp As String, fso As Object
    On Error Resume Next
    Set hwp = CreateObject("HWPFrame.HwpObject")
    On Error GoTo 0
    If hwp Is Nothing Then Err.Raise 1003, , "한글(HWP) 자동화를 사용할 수 없습니다. 한글에서 HWPX 또는 XLSX 로 저장한 뒤 다시 선택하세요."
    On Error Resume Next
    hwp.RegisterModule "FilePathCheckDLL", "FilePathCheckerModule"
    On Error GoTo 0
    Set fso = CreateObject("Scripting.FileSystemObject")
    tmp = fso.GetSpecialFolder(2) & "\hwp_" & Format(Now, "hhmmss") & ".hwpx"
    If Not hwp.Open(path, "HWP", "forceopen:true") Then Err.Raise 1004, , "HWP 파일을 열 수 없습니다."
    hwp.SaveAs tmp, "HWPX"
    hwp.Quit
    Set ReadHwpRows = ReadHwpxRows(tmp)
    On Error Resume Next
    fso.DeleteFile tmp
    On Error GoTo 0
End Function
