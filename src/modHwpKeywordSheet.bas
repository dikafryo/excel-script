Attribute VB_Name = "modHwpKeywordSheet"
Option Explicit
' =====================================================================================
'  modHwpKeywordSheet : 한글(HWP/HWPX)·PDF 키워드값 가져오기 (시트 방식)
'
'  화면 구성 (리본 [키워드값 가져오기] 클릭 시 생성)
'    1행 : [1 폴더선택] [2 키워드값 가져오기] 도형 버튼
'    A2  : "대상폴더"      B2 : 폴더 링크(클릭 시 탐색기)
'    A3  : "출처파일명(페이지)"  B3~J3 : 키워드1~9 (사용자가 덮어써서 입력)
'    4행~: 파일별 1행. A=파일명 링크(값을 찾은 페이지), B~J=키워드별 값
'
'  값 판별 규칙 (문서 순서대로 첫 번째 일치)
'    표 안에서 키워드 셀을 찾으면
'      ① 같은 셀에 "키워드: 값" 형태면 그 값
'      ② 오른쪽 셀이 값이면 오른쪽   (병합 셀은 건너뛰어 다음 실제 셀)
'      ③ 오른쪽이 비었거나 다른 키워드면 아래 셀
'      ④ 둘 다 아니면 다음 일치 위치를 계속 탐색
'    문단에서는 "키워드 : 값" 의 값, 없으면 바로 다음 줄
'    "값"의 기준 : 비어 있지 않고, 사용자가 입력한 다른 키워드가 아닌 것
'
'  구조 읽기 : hwpx → XML 표 구조 / hwp → 한글로 hwpx 임시 변환 / pdf → Word 표 객체
' =====================================================================================
Private Const KW_FIRST_COL As Long = 2      ' B
Private Const KW_LAST_COL As Long = 10      ' J
Private Const DATA_ROW As Long = 4
Private Const MAX_VALUE_LEN As Long = 200
Private hkNote As String   ' 파일별 처리 상태 메모
Private hkFuzzy As Boolean ' [v2.7] OCR 문서: 키워드 1글자 오차 허용

' ---- 문서 구조 (파일 1개 분량, 모듈 변수) ----
Private Type HKTable
    nRows As Long
    nCols As Long
    txt() As String
    cid() As Long
    page As Long
End Type
Private tbls() As HKTable
Private tblCount As Long
Private lines() As String
Private linePage() As Long
Private lineCount As Long
Private blocks() As Long        ' 문서 순서: 양수=문단 index, 음수=-표 index
Private blockCount As Long
Private cellSeq As Long

' =====================================================================================
'  [리본] 화면 만들기
' =====================================================================================
Public Sub HK_Start(control As IRibbonControl)
    Dim ws As Worksheet, i As Long
    Set ws = HK_NewSheet(ActiveWorkbook)
    HK_CreateButtons ws
    ws.Range("A2").value = "대상폴더: (1 폴더선택 버튼을 누르세요)"
    ws.Range("A3").value = "출처파일명(페이지)"
    For i = KW_FIRST_COL To KW_LAST_COL
        ws.Cells(3, i).value = "키워드" & (i - KW_FIRST_COL + 1)
        ws.Cells(3, i).Font.color = RGB(150, 150, 150)
    Next i
    With ws.Range(ws.Cells(2, 1), ws.Cells(3, KW_LAST_COL))
        .Font.name = "맑은 고딕"
        .Font.Size = 11
    End With
    ws.Range("A2:A3").Font.Bold = True
    ws.Range("A2:A3").Interior.color = RGB(221, 235, 247)
    With ws.Range(ws.Cells(3, KW_FIRST_COL), ws.Cells(3, KW_LAST_COL))
        .Interior.color = RGB(255, 242, 204)
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
    End With
    ws.Range("A3").Borders.LineStyle = xlContinuous
    ws.Columns("A").ColumnWidth = 30
    ws.Range(ws.Columns(KW_FIRST_COL), ws.Columns(KW_LAST_COL)).ColumnWidth = 16
    ws.Range("A5").value = "사용 순서:  ① [1 폴더선택]  ② 3행 B~J에 찾을 키워드 입력(예: 성명, 계좌번호)  ③ [2 키워드값 가져오기]"
    ws.Range("A6").value = "값 위치는 자동 판별합니다: 같은 셀(키워드: 값) → 오른쪽 셀 → 아래 셀 → 문단의 키워드 뒤 / 다음 줄"
    ws.Range("A5:A6").Font.color = RGB(100, 100, 100)
    ws.Range("A5:A6").Font.Size = 10
    ws.Activate
    ws.Range("B3").Select
End Sub

Private Function HK_NewSheet(ByVal wb As Workbook) As Worksheet
    Dim nm As String, n As Long, exist As Worksheet
    nm = "키워드값"
    Do
        Set exist = Nothing
        On Error Resume Next
        Set exist = wb.Worksheets(nm)
        On Error GoTo 0
        If exist Is Nothing Then Exit Do
        n = n + 1
        nm = "키워드값_" & n
    Loop
    Set HK_NewSheet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    HK_NewSheet.name = nm
End Function

Private Sub HK_CreateButtons(ByVal ws As Worksheet)
    Dim s As Shape, i As Long, names As Variant, captions As Variant, actions As Variant
    names = Array("btnHkFolder", "btnHkExtract")
    captions = Array("1 폴더선택", "2 키워드값 가져오기")
    actions = Array("HK_PickFolder", "HK_Extract")
    ws.rows(1).RowHeight = 42
    For i = 0 To 1
        On Error Resume Next
        ws.Shapes(CStr(names(i))).Delete
        On Error GoTo 0
        Set s = ws.Shapes.AddShape(msoShapeRoundedRectangle, 5 + i * 150, 6, 140, 30)
        s.name = CStr(names(i))
        UI_StyleButton s, CStr(captions(i)), CStr(actions(i)), i = 1
    Next i
End Sub

' =====================================================================================
'  [버튼 1] 폴더 선택 → B2 링크
' =====================================================================================
Public Sub HK_PickFolder()
    Dim ws As Worksheet, fldr As String, fso As Object, f As Object, n As Long
    Set ws = ActiveSheet
    With Application.FileDialog(msoFileDialogFolderPicker)
        .title = "HWP / HWPX / PDF 파일이 있는 폴더를 선택하세요"
        If .Show <> -1 Then Exit Sub
        fldr = .SelectedItems(1)
    End With
    Set fso = CreateObject("Scripting.FileSystemObject")
    For Each f In fso.GetFolder(fldr).files
        If HK_IsSupported(f.name) Then n = n + 1
    Next f
    ws.Range("A2").Hyperlinks.Delete
    ws.Range("A2").value = "대상폴더: " & fldr                 ' 09-29: 링크는 A2 (키워드1 열이 넓어지지 않게)
    ws.Hyperlinks.Add anchor:=ws.Range("A2"), Address:=fldr, TextToDisplay:="대상폴더: " & fldr, ScreenTip:="클릭하면 탐색기에서 폴더를 엽니다"
    ws.Range("A2").Font.Bold = True
    ws.Range("B2").value = "대상 파일 " & n & "개"
    ws.Range("B2").Font.color = RGB(100, 100, 100)
    If n = 0 Then MsgBox "선택한 폴더에 hwp/hwpx/pdf 파일이 없습니다.", vbExclamation, "키워드값 가져오기"
End Sub

' =====================================================================================
'  [버튼 2] 키워드값 가져오기
' =====================================================================================
Public Sub HK_Extract()
    Dim ws As Worksheet, fldr As String, kws() As String, kwCols() As Long, kwCount As Long
    Dim fso As Object, f As Object, files As Collection, r As Long, i As Long, k As Long
    Dim hwp As Object, wdApp As Object, ext As String, ok As Boolean, v As String
    Dim pages As Object, pageList As String, hwpPage As Long, t0 As Single
    Dim errMsg As String, errFile As String
    Set ws = ActiveSheet
    ws.Cells(3, 12).value = "처리 상태"
    fldr = HK_FolderOf(ws)
    If Len(fldr) = 0 Then
        MsgBox "먼저 [1 폴더선택] 버튼으로 대상 폴더를 지정하세요.", vbExclamation, "키워드값 가져오기"
        Exit Sub
    End If
    kwCount = HK_ReadKeywords(ws, kws, kwCols)
    If kwCount = 0 Then
        MsgBox "3행 B~J 칸에 찾을 키워드를 입력하세요. (예: 성명, 계좌번호, 연락처)", vbExclamation, "키워드값 가져오기"
        ws.Range("B3").Select
        Exit Sub
    End If
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set files = New Collection
    For Each f In fso.GetFolder(fldr).files
        If HK_IsSupported(f.name) Then files.Add f.path
    Next f
    If files.Count = 0 Then
        MsgBox "대상 폴더에 hwp/hwpx/pdf/이미지(png·jpg) 파일이 없습니다.", vbExclamation, "키워드값 가져오기"
        Exit Sub
    End If
    On Error GoTo HK_Fail   ' [수정] 파일 하나의 예기치 않은 오류로 한글/워드 프로세스가 남지 않도록
    ' 이전 결과 지우기
    With ws.Range(ws.Cells(DATA_ROW, 1), ws.Cells(ws.rows.Count, KW_LAST_COL))
        .Hyperlinks.Delete
        .Clear
    End With
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    t0 = Timer
    r = DATA_ROW
    For i = 1 To files.Count
        Application.StatusBar = "키워드값 가져오기 " & i & "/" & files.Count & " : " & fso.GetFileName(files(i))
        DoEvents
        ext = LCase$(fso.GetExtensionName(files(i)))
        ok = False
        Set pages = CreateObject("Scripting.Dictionary")
        If ext = "hwpx" Then
            ok = HK_LoadHwpx(files(i))
        ElseIf ext = "hwp" Then
            If hwp Is Nothing Then Set hwp = HK_CreateHwp()
            If hwp Is Nothing Then
                ws.Cells(r, KW_FIRST_COL).value = "한글(HWP) 프로그램을 찾을 수 없습니다"
            Else
                ok = HK_LoadHwpViaHwpx(files(i), hwp, fso)
            End If
        ElseIf ext = "pdf" Then
            If wdApp Is Nothing Then Set wdApp = HK_CreateWord()
            If wdApp Is Nothing Then
                ws.Cells(r, KW_FIRST_COL).value = "Word 프로그램을 찾을 수 없습니다(PDF 변환용)"
            Else
                ok = HK_LoadPdfViaWord(files(i), wdApp)
            End If
        ElseIf HK_IsImage(ext) Then
            ok = HK_LoadImageViaOcr(files(i), fso)
            If Not ok Then ws.Cells(r, KW_FIRST_COL).value = "OCR 실패: Windows 10 이상 + 한국어 언어팩(OCR) 필요"
        End If
        If ok Then
            For k = 1 To kwCount
                v = HK_FindValue(kws(k), kws, kwCount, hwpPage)
                ws.Cells(r, kwCols(k)).NumberFormat = "@"
                ws.Cells(r, kwCols(k)).value = v
                If Len(v) > 0 Then
                    If ext = "hwp" And Not hwp Is Nothing Then hwpPage = HK_HwpPageOf(hwp, kws(k))
                    If hwpPage > 0 Then pages(hwpPage) = 1
                End If
            Next k
            If ext = "hwp" And Not hwp Is Nothing Then HK_HwpClose hwp
        End If
        pageList = HK_PageText(pages)
        ws.Cells(r, 12).value = ext & IIf(ok, " 읽기 성공(" & tblCount & "표/" & lineCount & "줄)", " 읽기 실패") & IIf(Len(hkNote) > 0, " " & hkNote, "")
        hkNote = ""
        ws.Cells(r, 1).value = fso.GetFileName(files(i)) & IIf(Len(pageList) > 0, " (" & pageList & ")", "")
        On Error Resume Next
        ws.Hyperlinks.Add anchor:=ws.Cells(r, 1), Address:=files(i), TextToDisplay:=CStr(ws.Cells(r, 1).value)
        On Error GoTo 0
        r = r + 1
    Next i
    HK_FormatResult ws, r - 1, kwCols, kwCount
    On Error Resume Next
    If Not hwp Is Nothing Then hwp.Quit
    If Not wdApp Is Nothing Then wdApp.Quit
    On Error GoTo 0
    Application.StatusBar = False
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "완료: 파일 " & files.Count & "개, 키워드 " & kwCount & "개 (" & Format$(Timer - t0, "0.0") & "초)" & vbCrLf & vbCrLf & _
           "빈 칸은 문서에서 값을 찾지 못한 경우입니다. 키워드 표기(띄어쓰기·괄호)를 문서와 맞춰 다시 실행해 보세요.", vbInformation, "키워드값 가져오기"
    Exit Sub
HK_Fail:
    errMsg = Err.Description
    On Error Resume Next
    errFile = fso.GetFileName(files(i))
    If Not hwp Is Nothing Then HK_HwpClose hwp: hwp.Quit
    If Not wdApp Is Nothing Then wdApp.Quit
    Application.StatusBar = False
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "파일 처리 중 오류가 나서 중단했습니다." & vbCrLf & vbCrLf & _
           "파일: " & errFile & vbCrLf & "오류: " & errMsg & vbCrLf & vbCrLf & _
           "이 파일 앞까지의 결과는 시트에 남아 있습니다. 문제 파일을 폴더에서 빼고 다시 실행하세요.", vbExclamation, "키워드값 가져오기"
End Sub

Private Function HK_FolderOf(ByVal ws As Worksheet) As String
    ' 09-29: 폴더 링크는 A2("대상폴더: 경로"). 예전 시트(B2 링크)도 읽는다.
    Dim p As String, cands As Variant, i As Long, a As String
    cands = Array("A2", "B2")
    For i = 0 To 1
        a = CStr(cands(i)): p = ""
        On Error Resume Next
        If ws.Range(a).Hyperlinks.Count > 0 Then p = ws.Range(a).Hyperlinks(1).Address
        On Error GoTo 0
        If Len(p) = 0 Then p = Trim$(CStr(ws.Range(a).value))
        If Left$(p, 5) = "대상폴더:" Then p = Trim$(Mid$(p, 6))
        If Len(p) > 0 Then
            If dir(p, vbDirectory) <> "" And InStr(p, "\") > 0 Then HK_FolderOf = p: Exit Function
        End If
    Next i
    HK_FolderOf = ""
End Function

Private Function HK_ReadKeywords(ByVal ws As Worksheet, ByRef kws() As String, ByRef kwCols() As Long) As Long
    Dim c As Long, n As Long, s As String
    ReDim kws(1 To KW_LAST_COL): ReDim kwCols(1 To KW_LAST_COL)
    For c = KW_FIRST_COL To KW_LAST_COL
        s = Trim$(CStr(ws.Cells(3, c).value))
        If Len(s) > 0 And Not (Left$(s, 3) = "키워드" And IsNumeric(Mid$(s, 4)) And ws.Cells(3, c).Font.color = RGB(150, 150, 150)) Then   ' 회색 자리표시자만 건너뜀
            n = n + 1
            kws(n) = s: kwCols(n) = c
            ws.Cells(3, c).Font.color = RGB(0, 0, 0)
        End If
    Next c
    HK_ReadKeywords = n
End Function

Private Function HK_IsSupported(ByVal name As String) As Boolean
    Dim e As String
    If Left$(name, 2) = "~$" Then Exit Function
    e = LCase$(Mid$(name, InStrRev(name, ".") + 1))
    HK_IsSupported = (e = "hwp" Or e = "hwpx" Or e = "pdf" Or HK_IsImage(e))
End Function

Private Function HK_IsImage(ByVal e As String) As Boolean
    HK_IsImage = (e = "png" Or e = "jpg" Or e = "jpeg" Or e = "bmp" Or e = "tif" Or e = "tiff" Or e = "gif")
End Function

Private Function HK_PageText(ByVal pages As Object) As String
    Dim keys As Variant, i As Long, j As Long, T As Variant, s As String
    If pages.Count = 0 Then Exit Function
    keys = pages.keys
    For i = 0 To UBound(keys) - 1
        For j = i + 1 To UBound(keys)
            If keys(j) < keys(i) Then T = keys(i): keys(i) = keys(j): keys(j) = T
        Next j
    Next i
    For i = 0 To UBound(keys)
        s = s & IIf(i > 0, ",", "") & keys(i)
    Next i
    HK_PageText = "p." & s
End Function

Private Sub HK_FormatResult(ByVal ws As Worksheet, ByVal lastRow As Long, ByRef kwCols() As Long, ByVal kwCount As Long)
    Dim rng As Range, r As Long
    If lastRow < DATA_ROW Then Exit Sub
    Set rng = ws.Range(ws.Cells(DATA_ROW, 1), ws.Cells(lastRow, KW_LAST_COL))
    rng.Font.name = "맑은 고딕"
    rng.Font.Size = 11
    rng.Borders.LineStyle = xlContinuous
    rng.Borders.color = RGB(191, 191, 191)
    rng.VerticalAlignment = xlCenter
    For r = DATA_ROW To lastRow
        If r Mod 2 = 0 Then ws.Range(ws.Cells(r, 1), ws.Cells(r, KW_LAST_COL)).Interior.color = RGB(242, 242, 242)
    Next r
    ws.Range(ws.Columns(KW_FIRST_COL), ws.Columns(KW_LAST_COL)).AutoFit
    HK_CapWidths ws
    ' (A5:A6 안내문은 결과 영역 Clear 때 이미 지워짐 - 여기서 지우면 5·6행 파일명이 사라지는 버그였음)
End Sub

' =====================================================================================
'  값 찾기 (문서 구조 사용)
' =====================================================================================
Private Function HK_FindValue(ByVal kw As String, ByRef kws() As String, ByVal kwCount As Long, ByRef pageOut As Long) As String
    Dim b As Long, idx As Long, v As String
    pageOut = 0
    For b = 1 To blockCount
        idx = blocks(b)
        If idx < 0 Then
            v = HK_SearchTable(-idx, kw, kws, kwCount)
            If Len(v) > 0 Then pageOut = tbls(-idx).page
        Else
            v = HK_SearchLine(idx, kw, kws, kwCount)
            If Len(v) > 0 Then pageOut = linePage(idx)
        End If
        If Len(v) > 0 Then Exit For
    Next b
    HK_FindValue = HK_Clean(v, kw, kws, kwCount)
End Function

Private Function HK_SearchTable(ByVal ti As Long, ByVal kw As String, ByRef kws() As String, ByVal kwCount As Long) As String
    Dim r As Long, c As Long, cC As Long, rr As Long, CellText As String, v As String, kn As String
    Dim vR As String, vB As String, rowKw As Long, colKw As Long, belowFirst As Boolean
    kn = HK_Norm(kw)
    If Len(kn) = 0 Then Exit Function
    With tbls(ti)
        For r = 1 To .nRows
            For c = 1 To .nCols
                CellText = .txt(r, c)
                If hkFuzzy Then CellText = HK_FuzzyFix(CellText, kw)
                If HK_IsKeywordCell(CellText, kw, kn) Then
                    ' ① 같은 셀 "키워드: 값"
                    v = HK_AfterKeyword(CellText, kw)
                    If HK_IsValue(v, kws, kwCount) Then HK_SearchTable = v: Exit Function
                    ' 오른쪽 셀 / 아래 셀 (병합된 같은 셀은 건너뜀)
                    vR = "": vB = ""
                    cC = c + 1
                    Do While cC <= .nCols
                        If .cid(r, cC) <> .cid(r, c) Then Exit Do
                        cC = cC + 1
                    Loop
                    If cC <= .nCols Then vR = .txt(r, cC)
                    rr = r + 1
                    Do While rr <= .nRows
                        If .cid(rr, c) <> .cid(r, c) Then Exit Do
                        rr = rr + 1
                    Loop
                    If rr <= .nRows Then vB = .txt(rr, c)
                    ' [v2.7] 머리글 표(항목명이 가로로 나열되고 값이 아래 행) 판별:
                    '   같은 행의 다른 키워드 수 > 같은 열의 다른 키워드 수 → 아래 셀 우선, 아니면 오른쪽 셀 우선
                    rowKw = HK_CountKw(ti, r, 0, kws, kwCount, c)
                    colKw = HK_CountKw(ti, 0, c, kws, kwCount, r)
                    belowFirst = (rowKw > colKw)
                    ' 09-29: 서식형 표(라벨|값|라벨|값)는 오른쪽 칸에 값이 있으므로, 오른쪽이 비었거나 라벨/키워드이거나 머리글 행일 때만 아래 칸 우선
                    If belowFirst And Len(HK_Norm(vR)) > 0 Then
                        If Not (HK_LooksLabel(vR) Or Not HK_IsValue(vR, kws, kwCount) Or HK_IsAnyKw(vR, kws, kwCount) Or HK_IsHeaderRow(ti, r)) Then belowFirst = False
                    End If
                    If belowFirst Then
                        If HK_IsValue(vB, kws, kwCount) And Not HK_LooksLabel(vB) Then HK_SearchTable = vB: Exit Function
                        If HK_IsValue(vR, kws, kwCount) And Not HK_LooksLabel(vR) Then HK_SearchTable = vR: Exit Function
                        If HK_IsValue(vB, kws, kwCount) Then HK_SearchTable = vB: Exit Function
                    Else
                        If HK_IsValue(vR, kws, kwCount) And Not HK_LooksLabel(vR) Then HK_SearchTable = vR: Exit Function
                        If HK_IsValue(vB, kws, kwCount) And Not HK_LooksLabel(vB) Then HK_SearchTable = vB: Exit Function
                    End If
                    ' 둘 다 라벨처럼 보이면 오른쪽 셀이라도 채택 (값이 라벨처럼 생긴 드문 경우)
                    If HK_IsValue(vR, kws, kwCount) Then HK_SearchTable = vR: Exit Function
                    ' 계속 탐색
                End If
            Next c
        Next r
    End With
End Function

' 같은 행(r>0) 또는 같은 열(c>0)에서 사용자 키워드에 해당하는 칸 수 (skip 위치는 제외)
Private Function HK_CountKw(ByVal ti As Long, ByVal r As Long, ByVal c As Long, ByRef kws() As String, ByVal kwCount As Long, ByVal skip As Long) As Long
    ' 09-29: 병합 칸은 한 번만 센다(같은 cid), skip 칸(키워드 자신)과 같은 병합 칸도 제외
    Dim i As Long, k As Long, s As String, n As Long, lim As Long, id As Long, lastId As Long, skipId As Long
    With tbls(ti)
        If r > 0 Then lim = .nCols: skipId = .cid(r, skip) Else lim = .nRows: skipId = .cid(skip, c)
        lastId = -1
        For i = 1 To lim
            If r > 0 Then id = .cid(r, i): s = .txt(r, i) Else id = .cid(i, c): s = .txt(i, c)
            If id <> skipId And id <> lastId Then
                lastId = id
                For k = 1 To kwCount
                    If hkFuzzy Then s = HK_FuzzyFix(s, kws(k))
                    If HK_IsKeywordCell(s, kws(k), HK_Norm(kws(k))) Then n = n + 1: Exit For
                Next k
            End If
        Next i
    End With
    HK_CountKw = n
End Function

' 09-29: 글이 사용자 키워드 중 하나의 라벨 칸인지
Private Function HK_IsAnyKw(ByVal s As String, ByRef kws() As String, ByVal kwCount As Long) As Boolean
    Dim k As Long
    For k = 1 To kwCount
        If HK_IsKeywordCell(s, kws(k), HK_Norm(kws(k))) Then HK_IsAnyKw = True: Exit Function
    Next k
End Function

Private Function HK_SearchLine(ByVal li As Long, ByVal kw As String, ByRef kws() As String, ByVal kwCount As Long) As String
    Dim s As String, v As String, p As Long, nb As Long, nxt As Long
    s = lines(li)
    If hkFuzzy Then s = HK_FuzzyFix(s, kw)
    p = InStr(1, s, kw, vbTextCompare)
    If p = 0 Then Exit Function
    v = HK_AfterKeyword(s, kw)
    If HK_IsValue(v, kws, kwCount) Then HK_SearchLine = v: Exit Function
    If Len(HK_Norm(s)) = Len(HK_Norm(kw)) Then
        ' 줄 전체가 키워드뿐이면 다음 문단(비어 있지 않은 것)
        For nb = 1 To blockCount
            If blocks(nb) = li Then
                For nxt = nb + 1 To blockCount
                    If blocks(nxt) > 0 Then
                        v = lines(blocks(nxt))
                        If Len(Trim$(v)) > 0 Then
                            If HK_IsValue(v, kws, kwCount) Then HK_SearchLine = v
                            Exit Function
                        End If
                    Else
                        Exit Function
                    End If
                Next nxt
                Exit Function
            End If
        Next nb
    End If
End Function

' 헤더행 판정: 그 행의 채워진 셀이 2개 이상이고 모두 숫자 없는 짧은 라벨이며, 바로 아래 행에 값이 있음
Private Function HK_IsHeaderRow(ByVal ti As Long, ByVal r As Long) As Boolean
    Dim c As Long, s As String, filled As Long, below As Long, i As Long
    With tbls(ti)
        If r >= .nRows Then Exit Function
        For c = 1 To .nCols
            s = Trim$(.txt(r, c))
            If Len(s) > 0 Then
                filled = filled + 1
                If Len(s) > 12 Then Exit Function
                For i = 1 To Len(s)
                    If Mid$(s, i, 1) Like "#" Then Exit Function
                Next i
            End If
            If Len(Trim$(.txt(r + 1, c))) > 0 Then below = below + 1
        Next c
    End With
    HK_IsHeaderRow = (filled >= 2 And below >= 1)
End Function

Private Function HK_IsKeywordCell(ByVal CellText As String, ByVal kw As String, ByVal kn As String) As Boolean
    Dim cn As String
    cn = HK_Norm(CellText)
    If Len(cn) = 0 Then Exit Function
    If cn = kn Then HK_IsKeywordCell = True: Exit Function
    ' "성명(한글)" 처럼 키워드로 시작하고 짧게 덧붙은 라벨, 또는 "성명: 홍길동" 형태
    If Left$(cn, Len(kn)) = kn Then
        If InStr(1, CellText, kw, vbTextCompare) > 0 Then HK_IsKeywordCell = True
    End If
End Function

Private Function HK_AfterKeyword(ByVal s As String, ByVal kw As String) As String
    Dim p As Long, v As String, i As Long, ch As String
    p = InStr(1, s, kw, vbTextCompare)
    If p = 0 Then Exit Function
    v = Mid$(s, p + Len(kw))
    ' [수정] 키워드 바로 뒤에 글자/숫자가 이어지면("학교명", "성명란") 더 긴 라벨이므로 같은 셀 값으로 보지 않는다
    If Len(v) > 0 Then
        If Left$(v, 1) Like "[가-힣A-Za-z0-9]" Then Exit Function
    End If
    ' 괄호 보충어 제거: "(한글)" 등
    If Left$(LTrim$(v), 1) = "(" And InStr(v, ")") > 0 Then v = Mid$(v, InStr(v, ")") + 1)
    ' 앞쪽 구분 기호 제거
    For i = 1 To Len(v)
        ch = Mid$(v, i, 1)
        If InStr(" :：=-·_|/,、，" & ChrW(8226) & ChrW(8211) & ChrW(8212) & vbTab & Chr(7), ch) = 0 Then Exit For
    Next i
    v = Mid$(v, i)
    HK_AfterKeyword = Trim$(v)
End Function

' v 의 p 위치에 있는 키워드가 라벨로 쓰인 것인지: 앞 글자가 공백/구분자(또는 시작)이고 뒤 글자가 구분자(:, 공백 등)이거나 끝
Private Function HK_IsLabelAt(ByVal v As String, ByVal p As Long, ByVal kwLen As Long) As Boolean
    Dim prevCh As String, nextCh As String
    If p > 1 Then prevCh = Mid$(v, p - 1, 1)
    If p + kwLen <= Len(v) Then nextCh = Mid$(v, p + kwLen, 1)
    If Len(prevCh) > 0 Then
        If InStr(" ,;|/()[]" & vbTab, prevCh) = 0 Then Exit Function
    End If
    If Len(nextCh) = 0 Then HK_IsLabelAt = True: Exit Function
    HK_IsLabelAt = (InStr(" :：=-_|/(" & vbTab, nextCh) > 0)
End Function

' 항목 라벨처럼 보이는 짧은 문구 (값이 아니라 다음 항목명): 숫자 없이 6자 이하이고 명/번호/일자/주소 등으로 끝남
Private Function HK_LooksLabel(ByVal v As String) As Boolean
    Dim n As String, i As Long, tails As Variant, T As Variant
    ' 09-29: 가림 표시(* ○ ●)가 든 글은 가려진 값이지 라벨이 아님
    If InStr(v, "*") > 0 Or InStr(v, ChrW(9675)) > 0 Or InStr(v, ChrW(9679)) > 0 Then Exit Function
    n = HK_Norm(v)
    If Len(n) = 0 Or Len(n) > 6 Then Exit Function
    For i = 1 To Len(n)
        If Mid$(n, i, 1) Like "#" Then Exit Function
    Next i
    tails = Array("명", "번호", "일자", "일시", "주소", "성명", "연락처", "이메일", "비고", "구분", "성별", "직위", "소속", "학년", "학교", "생년월일", "은행", "예금주", "금액", "계좌", "전화", "휴대폰", "메일", "email", "부서", "직급", "과목", "기관", "업무", "학력", "경력", "점수", "학위")
    For Each T In tails
        If Right$(n, Len(T)) = T Then HK_LooksLabel = True: Exit Function
    Next T
End Function

Private Function HK_IsValue(ByVal v As String, ByRef kws() As String, ByVal kwCount As Long) As Boolean
    Dim i As Long, vn As String
    vn = HK_Norm(v)
    If Len(vn) = 0 Then Exit Function
    For i = 1 To kwCount
        If vn = HK_Norm(kws(i)) Then Exit Function
        If hkFuzzy Then
            If HK_Norm(HK_FuzzyFix(v, kws(i))) = HK_Norm(kws(i)) Then Exit Function
        End If
    Next i
    HK_IsValue = True
End Function

Private Function HK_Norm(ByVal s As String) As String
    Dim i As Long, ch As String, o As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If InStr(" " & vbTab & vbCr & vbLf & Chr(7) & ":：*※▶·-_|/()（）[]〔〕" & ChrW(8211) & ChrW(8212) & ChrW(9702), ch) = 0 Then o = o & ch
    Next i
    HK_Norm = LCase$(o)
End Function

Private Function HK_Clean(ByVal v As String, ByVal kw As String, ByRef kws() As String, ByVal kwCount As Long) As String
    Dim i As Long, p As Long, cut As Long
    v = Replace(Replace(Replace(v, vbCr, " "), vbLf, " "), vbTab, " ")
    v = Replace(v, Chr(7), "")
    Do While InStr(v, "  ") > 0
        v = Replace(v, "  ", " ")
    Loop
    v = Trim$(v)
    ' 값 뒤에 다른 키워드가 이어지면 그 앞에서 자르기 ("(주)가나다 대표: 홍길동")
    cut = 0
    For i = 1 To kwCount
        If StrComp(kws(i), kw, vbTextCompare) <> 0 Then
            p = InStr(2, v, kws(i), vbTextCompare)
            Do While p > 0
                If HK_IsLabelAt(v, p, Len(kws(i))) Then
                    If cut = 0 Or p < cut Then cut = p
                    Exit Do
                End If
                p = InStr(p + 1, v, kws(i), vbTextCompare)
            Loop
        End If
    Next i
    If cut > 0 Then v = Trim$(Left$(v, cut - 1))
    Do While Len(v) > 0
        If InStr(",;|:-·", Right$(v, 1)) > 0 Then v = Left$(v, Len(v) - 1) Else Exit Do
    Loop
    If Len(v) > MAX_VALUE_LEN Then v = Left$(v, MAX_VALUE_LEN) & "…"
    HK_Clean = Trim$(v)
End Function

' =====================================================================================
'  문서 읽기 : HWPX
' =====================================================================================
Private Sub HK_Reset()
    tblCount = 0: lineCount = 0: blockCount = 0: cellSeq = 0: hkFuzzy = False
    ReDim tbls(1 To 64): ReDim lines(1 To 512): ReDim linePage(1 To 512): ReDim blocks(1 To 576)
End Sub

Private Sub HK_AddLine(ByVal s As String, ByVal page As Long)
    s = Trim$(s)
    If Len(s) = 0 Then Exit Sub
    lineCount = lineCount + 1
    If lineCount > UBound(lines) Then
        ReDim Preserve lines(1 To UBound(lines) * 2)
        ReDim Preserve linePage(1 To UBound(lines))
    End If
    lines(lineCount) = s: linePage(lineCount) = page
    HK_AddBlock lineCount
End Sub

Private Sub HK_AddBlock(ByVal v As Long)
    blockCount = blockCount + 1
    If blockCount > UBound(blocks) Then ReDim Preserve blocks(1 To UBound(blocks) * 2)
    blocks(blockCount) = v
End Sub

Private Function HK_NewTable(ByVal nRows As Long, ByVal nCols As Long, ByVal page As Long) As Long
    tblCount = tblCount + 1
    If tblCount > UBound(tbls) Then ReDim Preserve tbls(1 To UBound(tbls) * 2)
    With tbls(tblCount)
        .nRows = nRows: .nCols = nCols: .page = page
        ReDim .txt(1 To nRows, 1 To nCols)
        ReDim .cid(1 To nRows, 1 To nCols)
    End With
    HK_AddBlock -tblCount
    HK_NewTable = tblCount
End Function

Private Function HK_LoadHwpx(ByVal path As String) As Boolean
    Dim fso As Object, tmp As String, secFiles As Collection, i As Long, xml As Object
    On Error GoTo Fail
    Set fso = CreateObject("Scripting.FileSystemObject")
    HK_Reset
    tmp = HK_Unzip(path, fso)
    If Len(tmp) = 0 Then Exit Function
    Set secFiles = New Collection
    For i = 0 To 30
        If fso.FileExists(fso.BuildPath(tmp, "Contents\section" & i & ".xml")) Then
            secFiles.Add fso.BuildPath(tmp, "Contents\section" & i & ".xml")
        Else
            If i > 0 Then Exit For
        End If
    Next i
    For i = 1 To secFiles.Count
        Set xml = CreateObject("MSXML2.DOMDocument.6.0")
        xml.async = False: xml.validateOnParse = False
        If xml.Load(secFiles(i)) Then HK_WalkParent xml.documentElement, 0
    Next i
    HK_LoadHwpx = (blockCount > 0)
Cleanup:
    On Error Resume Next
    If Len(tmp) > 0 Then fso.DeleteFolder tmp, True
    Exit Function
Fail:
    Resume Cleanup
End Function

' 부모 노드 아래의 문단(hp:p)들을 순서대로 처리
Private Sub HK_WalkParent(ByVal parent As Object, ByVal page As Long)
    Dim node As Object
    For Each node In parent.childNodes
        If node.nodeName = "hp:p" Then
            HK_WalkPara node, page
        ElseIf node.nodeType = 1 Then
            HK_WalkParent node, page
        End If
    Next node
End Sub

' 문단: 텍스트는 한 줄로, 표는 표 블록으로 (문서 순서 유지)
Private Sub HK_WalkPara(ByVal pNode As Object, ByVal page As Long)
    Dim run As Object, child As Object, s As String
    For Each run In pNode.childNodes
        If run.nodeName = "hp:run" Then
            For Each child In run.childNodes
                If child.nodeName = "hp:t" Then
                    s = s & child.text
                ElseIf child.nodeName = "hp:tbl" Then
                    If Len(Trim$(s)) > 0 Then HK_AddLine s, page: s = ""
                    HK_ParseTable child, page
                ElseIf child.nodeType = 1 Then
                    ' 글상자 등 다른 개체 안의 문단
                    If Len(Trim$(s)) > 0 Then HK_AddLine s, page: s = ""
                    HK_WalkParent child, page
                End If
            Next child
        End If
    Next run
    If Len(Trim$(s)) > 0 Then HK_AddLine s, page
End Sub

Private Sub HK_ParseTable(ByVal tblNode As Object, ByVal page As Long)
    Dim nRows As Long, nCols As Long, ti As Long, tr As Object, tc As Object, node As Object
    Dim r As Long, c As Long, rs As Long, cs As Long, i As Long, j As Long, txt As String, curRow As Long
    nRows = val(tblNode.getAttribute("rowCnt")): nCols = val(tblNode.getAttribute("colCnt"))
    If nRows <= 0 Or nCols <= 0 Then Exit Sub
    ti = HK_NewTable(nRows, nCols, page)
    curRow = 0
    For Each tr In tblNode.childNodes
        If tr.nodeName = "hp:tr" Then
            curRow = curRow + 1
            For Each tc In tr.childNodes
                If tc.nodeName = "hp:tc" Then
                    r = curRow: c = 0: rs = 1: cs = 1: txt = ""
                    For Each node In tc.childNodes
                        Select Case node.nodeName
                            Case "hp:cellAddr"
                                r = val(node.getAttribute("rowAddr")) + 1
                                c = val(node.getAttribute("colAddr")) + 1
                            Case "hp:cellSpan"
                                cs = val(node.getAttribute("colSpan")): rs = val(node.getAttribute("rowSpan"))
                                If cs < 1 Then cs = 1
                                If rs < 1 Then rs = 1
                            Case "hp:subList"
                                txt = HK_CellText(node, page)
                        End Select
                    Next node
                    If c = 0 Then c = 1
                    cellSeq = cellSeq + 1
                    For i = r To r + rs - 1
                        For j = c To c + cs - 1
                            If i >= 1 And i <= nRows And j >= 1 And j <= nCols Then
                                tbls(ti).txt(i, j) = txt
                                tbls(ti).cid(i, j) = cellSeq
                            End If
                        Next j
                    Next i
                End If
            Next tc
        End If
    Next tr
End Sub

' 셀 텍스트: 셀 안 문단들을 공백으로 이어 붙임. 셀 안의 표(중첩)는 별도 표 블록으로 등록
Private Function HK_CellText(ByVal subList As Object, ByVal page As Long) As String
    Dim p As Object, run As Object, child As Object, s As String
    For Each p In subList.childNodes
        If p.nodeName = "hp:p" Then
            For Each run In p.childNodes
                If run.nodeName = "hp:run" Then
                    For Each child In run.childNodes
                        If child.nodeName = "hp:t" Then
                            s = s & child.text
                        ElseIf child.nodeName = "hp:tbl" Then
                            HK_ParseTable child, page
                        End If
                    Next child
                End If
            Next run
            s = s & " "
        End If
    Next p
    HK_CellText = Trim$(s)
End Function

Private Function HK_Unzip(ByVal path As String, ByVal fso As Object) As String
    Dim shellApp As Object, tmpRoot As String, stamp As String, tmpFolder As String, zipPath As String
    Dim nsZip As Object, nsDest As Object, T As Single
    On Error GoTo Fail
    Set shellApp = CreateObject("Shell.Application")
    tmpRoot = Environ$("TEMP")
    stamp = Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(Int(Rnd() * 100000))
    tmpFolder = fso.BuildPath(tmpRoot, "hk_" & stamp)
    zipPath = fso.BuildPath(tmpRoot, "hk_" & stamp & ".zip")
    fso.CopyFile path, zipPath, True
    ' PowerShell Expand-Archive (동기 실행) - Shell.Application CopyHere는 환경에 따라 아무것도 풀지 않는 경우가 있음
    Set shellApp = CreateObject("WScript.Shell")
    shellApp.run "powershell.exe -NoProfile -ExecutionPolicy Bypass -Command ""Expand-Archive -LiteralPath '" & zipPath & "' -DestinationPath '" & tmpFolder & "' -Force""", 0, True
    If Not fso.FileExists(fso.BuildPath(tmpFolder, "Contents\section0.xml")) Then
        ' 예비: Shell.Application 방식
        Set nsZip = CreateObject("Shell.Application").Namespace(CVar(zipPath))
        If Not fso.FolderExists(tmpFolder) Then fso.CreateFolder tmpFolder
        Set nsDest = CreateObject("Shell.Application").Namespace(CVar(tmpFolder))
        If nsZip Is Nothing Or nsDest Is Nothing Then GoTo Fail
        nsDest.CopyHere nsZip.items, 4 + 16 + 1024
        T = Timer
        Do While Timer - T < 20
            DoEvents
            If fso.FileExists(fso.BuildPath(tmpFolder, "Contents\section0.xml")) Then Exit Do
        Loop
    End If
    If Not fso.FileExists(fso.BuildPath(tmpFolder, "Contents\section0.xml")) Then GoTo Fail
    On Error Resume Next
    fso.DeleteFile zipPath, True
    HK_Unzip = tmpFolder
    Exit Function
Fail:
    On Error Resume Next
    If Len(zipPath) > 0 Then fso.DeleteFile zipPath, True
    HK_Unzip = ""
End Function

' =====================================================================================
'  문서 읽기 : HWP (한글로 열어 hwpx 임시 변환)
' =====================================================================================
Private Function HK_CreateHwp() As Object
    Dim h As Object
    On Error Resume Next
    Set h = CreateObject("HWPFrame.HwpObject")
    If h Is Nothing Then Set h = CreateObject("Hancom.HwpObject")
    If Not h Is Nothing Then
        h.RegisterModule "FilePathCheckDLL", "FilePathCheckerModule"
        h.XHwpWindows.item(0).Visible = False
    End If
    Set HK_CreateHwp = h
End Function

Private Function HK_LoadHwpViaHwpx(ByVal path As String, ByVal hwp As Object, ByVal fso As Object) As Boolean
    Dim tmpX As String, ok As Boolean
    On Error Resume Next
    Err.Clear
    ' [핵심 수정] 한글 COM 의 Open/SaveAs 는 Boolean 1 을 돌려주므로 Not ok 가 True 가 되는 VBA 함정 → (<> 0) 으로 정규화
    ok = (hwp.Open(path, "HWP", "forceopen:true") <> 0)
    hkNote = "open=" & ok
    If Err.Number <> 0 Then hkNote = hkNote & " openerr=" & Err.Description: Err.Clear
    If Not ok Then
        HK_LoadHwpViaHwpx = HK_LoadHwpText(hwp)
        hkNote = hkNote & " text=" & HK_LoadHwpViaHwpx
        Exit Function
    End If
    tmpX = fso.BuildPath(Environ$("TEMP"), "hk_" & Format$(Now, "hhnnss") & "_" & Int(Rnd() * 10000) & ".hwpx")
    Err.Clear
    ok = (hwp.SaveAs(tmpX, "HWPX", "") <> 0)
    hkNote = hkNote & " saveas=" & ok & "/" & fso.FileExists(tmpX)
    If Err.Number <> 0 Then hkNote = hkNote & " saveerr=" & Err.Description: Err.Clear
    If ok And fso.FileExists(tmpX) Then
        HK_LoadHwpViaHwpx = HK_LoadHwpx(tmpX)
        hkNote = hkNote & " hwpx=" & HK_LoadHwpViaHwpx
        If Err.Number <> 0 Then hkNote = hkNote & " parseerr=" & Err.Description: Err.Clear
        fso.DeleteFile tmpX, True
    End If
    Err.Clear
    ' 원본 문서를 다시 열어 둔다(페이지 찾기용). SaveAs 뒤에는 현재 문서가 hwpx 사본이므로 재오픈
    hwp.Open path, "HWP", "forceopen:true"
    If Not HK_LoadHwpViaHwpx Then
        HK_LoadHwpViaHwpx = HK_LoadHwpText(hwp)
        hkNote = hkNote & " text=" & HK_LoadHwpViaHwpx
    End If
End Function

' 예비: 한글 텍스트 추출(표 구조 없음). 표 셀이 한 줄씩 나오므로 "키워드 다음 줄" 규칙으로 판별
Private Function HK_LoadHwpText(ByVal hwp As Object) As Boolean
    Dim T As String, ln As Variant
    On Error Resume Next
    HK_Reset
    T = hwp.GetTextFile("UNICODE", "")
    If Len(T) = 0 Then T = hwp.GetTextFile("TEXT", "")
    If Len(T) = 0 Then Exit Function
    T = Replace(Replace(T, vbCr, vbLf), Chr(7), vbLf)
    For Each ln In Split(T, vbLf)
        HK_AddLine CStr(ln), 0
    Next ln
    HK_LoadHwpText = (lineCount > 0)
End Function

Private Sub HK_HwpClose(ByVal hwp As Object)
    On Error Resume Next
    hwp.HAction.run "FileClose"
End Sub

' 현재 열린 한글 문서에서 키워드가 처음 나오는 쪽 번호 (실패 시 0)
Private Function HK_HwpPageOf(ByVal hwp As Object, ByVal kw As String) As Long
    Dim act As Object, frSet As Object, found As Boolean
    Dim seccnt As Variant, secno As Variant, prnpageno As Variant, colno As Variant
    Dim ln As Variant, pos As Variant, over As Variant, ctrlname As Variant
    On Error GoTo Fail
    hwp.HAction.run "MoveDocBegin"
    Set act = hwp.HAction
    Set frSet = hwp.HParameterSet.HFindReplace
    act.GetDefault "RepeatFind", frSet.HSet
    frSet.FindString = kw
    frSet.IgnoreMessage = 1
    frSet.Direction = hwp.FindDir("Forward")
    frSet.MatchCase = 0
    frSet.FindType = 1
    found = act.Execute("RepeatFind", frSet.HSet)
    If Not found Then Exit Function
    hwp.KeyIndicator seccnt, secno, prnpageno, colno, ln, pos, over, ctrlname
    HK_HwpPageOf = CLng(prnpageno)
    Exit Function
Fail:
    HK_HwpPageOf = 0
End Function

' =====================================================================================
'  문서 읽기 : PDF (Word로 열어 표/문단 읽기)
' =====================================================================================
Private Function HK_CreateWord() As Object
    Dim W As Object
    On Error Resume Next
    Set W = CreateObject("Word.Application")
    If Not W Is Nothing Then
        W.Visible = False
        W.DisplayAlerts = 0
    End If
    Set HK_CreateWord = W
End Function

Private Function HK_LoadPdfViaWord(ByVal path As String, ByVal wdApp As Object) As Boolean
    Dim doc As Object, para As Object, tbl As Object, ti As Long, r As Long, c As Long
    Dim Done As Object, key As String, txt As String, page As Long, cel As Object
    On Error GoTo Fail
    HK_Reset
    Set doc = wdApp.Documents.Open(path, False, True, False)
    Set Done = CreateObject("Scripting.Dictionary")
    For Each para In doc.Paragraphs
        page = para.Range.Information(3)          ' wdActiveEndPageNumber
        If para.Range.Information(12) Then          ' wdWithInTable
            Set tbl = para.Range.Tables(1)
            key = CStr(tbl.Range.Start)
            If Not Done.Exists(key) Then
                Done.Add key, 1
                ti = HK_NewTable(tbl.rows.Count, tbl.Columns.Count, page)
                On Error Resume Next
                For Each cel In tbl.Range.Cells
                    r = cel.RowIndex: c = cel.ColumnIndex
                    txt = cel.Range.text
                    If Len(txt) >= 2 Then txt = Left$(txt, Len(txt) - 2)   ' 셀 끝 표식 제거
                    If r >= 1 And r <= tbls(ti).nRows And c >= 1 And c <= tbls(ti).nCols Then
                        cellSeq = cellSeq + 1
                        tbls(ti).txt(r, c) = Trim$(Replace(txt, vbCr, " "))
                        tbls(ti).cid(r, c) = cellSeq
                    End If
                Next cel
                On Error GoTo Fail
            End If
        Else
            HK_AddLine para.Range.text, page
        End If
    Next para
    doc.Close False
    HK_LoadPdfViaWord = (blockCount > 0)
    Exit Function
Fail:
    On Error Resume Next
    If Not doc Is Nothing Then doc.Close False
    HK_LoadPdfViaWord = False
End Function

' =====================================================================================
'  문서 읽기 : 이미지 (Windows 내장 OCR, Windows 10 이상 / Office 버전 무관)
'   [v2.7] OCR 단어의 좌표(BoundingRect)를 받아 표 구조를 흉내 낸다:
'     ① 단어를 세로 위치로 행으로 묶고 ② 행 안에서 글자 폭의 2배가 넘는 빈틈을 칸 경계로 삼아 칸을 나눈 뒤
'     ③ 한 표(행 × 칸)로 등록 → hwpx 표와 같은 규칙(같은 칸 "키워드: 값" → 오른쪽 칸 → 아래 칸)으로 값을 찾는다.
'     좌표가 없으면(구형 결과) 줄 단위 규칙으로 처리
' =====================================================================================
Private Function HK_LoadImageViaOcr(ByVal path As String, ByVal fso As Object) As Boolean
    Dim tmp As String, ps1 As String, outTxt As String, Sh As Object, cmd As Object
    Dim T As Single, txt As String, ln As Variant, parts() As String
    Dim wt() As String, wx() As Single, wy() As Single, ww() As Single, wh() As Single, nw As Long, hasWords As Boolean
    On Error GoTo Fail
    HK_Reset
    hkFuzzy = True
    tmp = Environ$("TEMP")
    ps1 = fso.BuildPath(tmp, "bx_winocr2.ps1")
    outTxt = fso.BuildPath(tmp, "bx_winocr_" & Format$(Now, "hhnnss") & Int(Rnd() * 1000) & ".txt")
    HK_WriteUtf8File ps1, HK_OcrScript()
    Set Sh = CreateObject("WScript.Shell")
    Set cmd = Sh.Exec("powershell.exe -NoProfile -ExecutionPolicy Bypass -File """ & ps1 & """ -path """ & path & """ -out """ & outTxt & """")
    T = Timer
    Do While cmd.status = 0
        DoEvents
        If Timer - T > 60 Then Exit Do
    Loop
    If Not fso.FileExists(outTxt) Then Exit Function
    txt = HK_ReadUtf8File(outTxt)
    fso.DeleteFile outTxt, True
    If Left$(txt, 5) = "ERROR" Then hkNote = Mid$(txt, 1, 80): Exit Function
    ReDim wt(1 To 1024): ReDim wx(1 To 1024): ReDim wy(1 To 1024): ReDim ww(1 To 1024): ReDim wh(1 To 1024)
    For Each ln In Split(Replace(txt, vbCr, ""), vbLf)
        If Left$(CStr(ln), 2) = "S" & vbTab Then
            ' 새 배율 구간: 앞 구간의 단어로 표 등록
            If nw > 0 Then HK_BuildOcrTable wt, wx, wy, ww, wh, nw
            nw = 0: hasWords = True
        ElseIf Left$(CStr(ln), 2) = "W" & vbTab Then
            parts = Split(CStr(ln), vbTab)
            If UBound(parts) >= 6 Then
                nw = nw + 1
                If nw > UBound(wt) Then
                    ReDim Preserve wt(1 To nw * 2): ReDim Preserve wx(1 To nw * 2): ReDim Preserve wy(1 To nw * 2)
                    ReDim Preserve ww(1 To nw * 2): ReDim Preserve wh(1 To nw * 2)
                End If
                wt(nw) = parts(2): wx(nw) = val(parts(3)): wy(nw) = val(parts(4)): ww(nw) = val(parts(5)): wh(nw) = val(parts(6))
                hasWords = True
            End If
        ElseIf Left$(CStr(ln), 2) = "L" & vbTab Then
            ' 줄 글자는 표 등록 때 행 단위로 넣으므로 여기서는 생략
        ElseIf Not hasWords Then
            HK_AddLine CStr(ln), 1
        End If
    Next ln
    If hasWords Then HK_BuildOcrTable wt, wx, wy, ww, wh, nw
    HK_LoadImageViaOcr = (lineCount > 0 Or tblCount > 0)
    Exit Function
Fail:
    hkNote = "OCR 오류: " & Err.Description
    HK_LoadImageViaOcr = False
End Function

' 단어 좌표 → 행/칸 표. 칸 순서(왼쪽부터 1,2,3…)를 열 번호로 써서 "오른쪽 칸"이 항상 같은 행의 다음 칸이 되게 한다
Private Sub HK_BuildOcrTable(ByRef wt() As String, ByRef wx() As Single, ByRef wy() As Single, ByRef ww() As Single, ByRef wh() As Single, ByVal nw As Long)
    Dim idx() As Long, i As Long, j As Long, T As Long, cy() As Single, sumH As Single, avgH As Single
    Dim rowOf() As Long, nRows As Long, rowCy As Single, rowN As Long, units As Long, sumW As Single, avgCharW As Single, gapThr As Single
    Dim rowStart() As Long, rowEnd() As Long, r As Long, cellTxt As String, prevRight As Single, nCells As Long, maxCells As Long
    Dim ti As Long, rowCells() As String, rowCellN() As Long, k As Long, rowText As String
    Dim rowCellX() As Single, rowCellCol() As Long, ax() As Single, ar() As Long, ac() As Long, nAll As Long, aIdx() As Long, nColX As Long, lastX As Single, colTol As Single
    If nw = 0 Then Exit Sub
    ' 평균 글자 폭(한글 1자 기준): 단어 폭 / 글자 단위 수
    For i = 1 To nw
        units = HK_TextUnits(wt(i))
        If units > 0 Then sumW = sumW + ww(i) / (units / 2)
        sumH = sumH + wh(i)
    Next i
    avgH = sumH / nw
    ' 글자 폭은 단어 높이로 추정(한글 한 자 = 대략 정사각형). 단어 폭/글자수 평균은 "값1" 같은 짧은 단어에 왜곡됨
    avgCharW = avgH * 0.9
    If avgCharW < 4 Then avgCharW = 4
    gapThr = avgCharW * 2.2
    ' 세로 중심으로 정렬
    ReDim idx(1 To nw): ReDim cy(1 To nw)
    For i = 1 To nw
        idx(i) = i: cy(i) = wy(i) + wh(i) / 2
    Next i
    For i = 1 To nw - 1
        For j = i + 1 To nw
            If cy(idx(j)) < cy(idx(i)) Then T = idx(i): idx(i) = idx(j): idx(j) = T
        Next j
    Next i
    ' 행으로 묶기 (중심 차이가 글자 높이의 0.6배 이내)
    ReDim rowOf(1 To nw): ReDim rowStart(1 To nw): ReDim rowEnd(1 To nw)
    nRows = 0
    For i = 1 To nw
        If nRows = 0 Then
            nRows = 1: rowCy = cy(idx(i)): rowN = 1: rowStart(1) = i
        ElseIf Abs(cy(idx(i)) - rowCy) <= avgH * 0.6 Then
            rowCy = (rowCy * rowN + cy(idx(i))) / (rowN + 1): rowN = rowN + 1
        Else
            rowEnd(nRows) = i - 1
            nRows = nRows + 1: rowCy = cy(idx(i)): rowN = 1: rowStart(nRows) = i
        End If
        rowOf(idx(i)) = nRows
    Next i
    rowEnd(nRows) = nw
    ' 행 안에서 가로로 정렬한 뒤 빈틈으로 칸 나누기
    ReDim rowCells(1 To nRows, 1 To 64): ReDim rowCellN(1 To nRows): ReDim rowCellX(1 To nRows, 1 To 64): ReDim rowCellCol(1 To nRows, 1 To 64)
    For r = 1 To nRows
        For i = rowStart(r) To rowEnd(r) - 1
            For j = i + 1 To rowEnd(r)
                If wx(idx(j)) < wx(idx(i)) Then T = idx(i): idx(i) = idx(j): idx(j) = T
            Next j
        Next i
        nCells = 0: cellTxt = "": prevRight = -1
        For i = rowStart(r) To rowEnd(r)
            k = idx(i)
            If prevRight >= 0 And wx(k) - prevRight > gapThr And Not (Right$(cellTxt, 1) = "-" Or Left$(wt(k), 1) = "-") Then
                nCells = nCells + 1
                If nCells <= 64 Then rowCells(r, nCells) = Trim$(cellTxt)
                cellTxt = ""
            End If
            If Len(cellTxt) = 0 And nCells + 1 <= 64 Then rowCellX(r, nCells + 1) = wx(k)   ' 칸의 왼쪽 x
            cellTxt = cellTxt & IIf(Len(cellTxt) > 0, " ", "") & wt(k)
            If wx(k) + ww(k) > prevRight Then prevRight = wx(k) + ww(k)
        Next i
        nCells = nCells + 1
        If nCells <= 64 Then rowCells(r, nCells) = Trim$(cellTxt)
        If nCells > 64 Then nCells = 64
        rowCellN(r) = nCells
        If nCells > maxCells Then maxCells = nCells
    Next r
    ' 열 번호: 모든 칸의 왼쪽 x 를 모아 가까운 것끼리 묶는다(같은 열의 칸은 x 가 비슷함) → OCR 이 한 칸을 놓쳐도 아래 칸이 어긋나지 않음
    nAll = 0
    For r = 1 To nRows
        nAll = nAll + rowCellN(r)
    Next r
    ReDim ax(1 To nAll): ReDim ar(1 To nAll): ReDim ac(1 To nAll): ReDim aIdx(1 To nAll)
    i = 0
    For r = 1 To nRows
        For k = 1 To rowCellN(r)
            i = i + 1: ax(i) = rowCellX(r, k): ar(i) = r: ac(i) = k: aIdx(i) = i
        Next k
    Next r
    For i = 1 To nAll - 1
        For j = i + 1 To nAll
            If ax(aIdx(j)) < ax(aIdx(i)) Then T = aIdx(i): aIdx(i) = aIdx(j): aIdx(j) = T
        Next j
    Next i
    colTol = avgCharW * 1.5
    nColX = 0: lastX = -1000000000#
    For i = 1 To nAll
        k = aIdx(i)
        If ax(k) - lastX > colTol Then nColX = nColX + 1
        rowCellCol(ar(k), ac(k)) = nColX
        lastX = ax(k)
    Next i
    If nColX < 1 Then nColX = 1
    ' 표 등록 + 행 글자를 줄로도 등록(예비 규칙용)
    ti = HK_NewTable(nRows, nColX, 1)
    For r = 1 To nRows
        rowText = ""
        For k = 1 To rowCellN(r)
            j = rowCellCol(r, k)
            If j < 1 Then j = 1
            cellSeq = cellSeq + 1
            If Len(tbls(ti).txt(r, j)) > 0 Then
                tbls(ti).txt(r, j) = tbls(ti).txt(r, j) & " " & rowCells(r, k)
            Else
                tbls(ti).txt(r, j) = rowCells(r, k)
                tbls(ti).cid(r, j) = cellSeq
            End If
            rowText = rowText & IIf(Len(rowText) > 0, "  ", "") & rowCells(r, k)
        Next k
        HK_AddLine rowText, 1
    Next r
    hkNote = hkNote & IIf(Len(hkNote) > 0, " ", "OCR ") & nRows & "행/" & nColX & "칸"
End Sub

' 글자 단위 수: 한글·한자 등 전각 2, 그 외 1
Private Function HK_TextUnits(ByVal s As String) As Long
    Dim i As Long, n As Long
    For i = 1 To Len(s)
        If AscW(Mid$(s, i, 1)) > 255 Then n = n + 2 Else n = n + 1
    Next i
    HK_TextUnits = n
End Function

Private Sub HK_WriteUtf8File(ByVal path As String, ByVal content As String)
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2: st.Charset = "utf-8": st.Open
    st.WriteText content
    st.SaveToFile path, 2
    st.Close
End Sub

Private Function HK_ReadUtf8File(ByVal path As String) As String
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2: st.Charset = "utf-8": st.Open
    st.LoadFromFile path
    HK_ReadUtf8File = st.ReadText(-1)
    st.Close
End Function

' OCR 결과: 배율마다 "S<TAB>배율" 뒤에 단어 "W<TAB>줄번호<TAB>글자<TAB>x<TAB>y<TAB>폭<TAB>높이"(원본 좌표), 줄 "L<TAB>줄글자"
' 배율에 따라 읽히는 글자가 달라 1.4·1.2·1.0·1.7·2.0 배로 각각 읽고 모두 표로 등록한다(앞 배율부터 검색)
Private Function HK_OcrScript() As String
    Dim s As String
    s = s & "param($path, $out, $lang=""ko"")" & vbCrLf
    s = s & "try {" & vbCrLf
    s = s & "Add-Type -AssemblyName System.Runtime.WindowsRuntime" & vbCrLf
    s = s & "$null = [Windows.Media.Ocr.OcrEngine, Windows.Foundation, ContentType=WindowsRuntime]" & vbCrLf
    s = s & "$null = [Windows.Storage.StorageFile, Windows.Storage, ContentType=WindowsRuntime]" & vbCrLf
    s = s & "$null = [Windows.Graphics.Imaging.BitmapDecoder, Windows.Graphics.Imaging, ContentType=WindowsRuntime]" & vbCrLf
    s = s & "$null = [Windows.Globalization.Language, Windows.Globalization, ContentType=WindowsRuntime]" & vbCrLf
    s = s & "$g = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' })[0]" & vbCrLf
    s = s & "function Await($op, $type){ $t=$g.MakeGenericMethod($type).Invoke($null,@($op)); $t.Wait(-1) | Out-Null; $t.Result }" & vbCrLf
    s = s & "$eng = [Windows.Media.Ocr.OcrEngine]::TryCreateFromLanguage((New-Object Windows.Globalization.Language($lang)))" & vbCrLf
    s = s & "if (-not $eng) { $eng = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages() }" & vbCrLf
    s = s & "if (-not $eng) { throw ""OCR engine not available"" }" & vbCrLf
    s = s & "Add-Type -AssemblyName System.Drawing" & vbCrLf
    s = s & "$src = [System.Drawing.Image]::FromFile($path)" & vbCrLf
    s = s & "$lines = @()" & vbCrLf
    s = s & "foreach ($scale in @(1.4, 1.2, 1.0, 1.7, 2.0)) {" & vbCrLf
    s = s & "  $nw = [int]($src.Width * $scale); $nh = [int]($src.Height * $scale)" & vbCrLf
    s = s & "  if ($nw -gt 6000 -or $nh -gt 6000) { continue }" & vbCrLf
    s = s & "  $bmp = New-Object System.Drawing.Bitmap $nw, $nh" & vbCrLf
    s = s & "  $gr = [System.Drawing.Graphics]::FromImage($bmp); $gr.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic" & vbCrLf
    s = s & "  $gr.DrawImage($src, 0, 0, $nw, $nh); $gr.Dispose()" & vbCrLf
    s = s & "  $tmpPng = [IO.Path]::Combine([IO.Path]::GetTempPath(), 'bx_ocr_' + [IO.Path]::GetRandomFileName() + '.png')" & vbCrLf
    s = s & "  $bmp.Save($tmpPng, [System.Drawing.Imaging.ImageFormat]::Png); $bmp.Dispose()" & vbCrLf
    s = s & "  $file = Await ([Windows.Storage.StorageFile]::GetFileFromPathAsync($tmpPng)) ([Windows.Storage.StorageFile])" & vbCrLf
    s = s & "  $stream = Await ($file.OpenAsync([Windows.Storage.FileAccessMode]::Read)) ([Windows.Storage.Streams.IRandomAccessStream])" & vbCrLf
    s = s & "  $dec = Await ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)) ([Windows.Graphics.Imaging.BitmapDecoder])" & vbCrLf
    s = s & "  $sb = Await ($dec.GetSoftwareBitmapAsync()) ([Windows.Graphics.Imaging.SoftwareBitmap])" & vbCrLf
    s = s & "  $res = Await ($eng.RecognizeAsync($sb)) ([Windows.Media.Ocr.OcrResult])" & vbCrLf
    s = s & "  $stream.Dispose(); Remove-Item $tmpPng -ErrorAction SilentlyContinue" & vbCrLf
    s = s & "  $lines += ('S' + [char]9 + $scale); $li = 0" & vbCrLf
    s = s & "  foreach ($l in $res.Lines) { $li++; $lines += ('L' + [char]9 + $l.Text); foreach ($w in $l.Words) { $b = $w.BoundingRect; $lines += ('W' + [char]9 + $li + [char]9 + $w.Text + [char]9 + [int]($b.X / $scale) + [char]9 + [int]($b.Y / $scale) + [char]9 + [int]($b.Width / $scale) + [char]9 + [int]($b.Height / $scale)) } }" & vbCrLf
    s = s & "}" & vbCrLf
    s = s & "$src.Dispose()" & vbCrLf
    s = s & "[IO.File]::WriteAllText($out, ($lines -join ""`n""), (New-Object Text.UTF8Encoding($false)))" & vbCrLf
    s = s & "} catch { [IO.File]::WriteAllText($out, ""ERROR "" + $_.Exception.Message, (New-Object Text.UTF8Encoding($false))) }" & vbCrLf
    HK_OcrScript = s
End Function

' [v2.7] OCR 오독 보정: 글 s 안에 키워드(3자 이상)와 길이가 같고 한 글자만 다른 조각이 있으면 그 조각을 키워드로 바꿔 돌려준다
'   숫자 자리는 정확히 같아야 함 → "키워드2"가 "키워드1"로 바뀌거나 "010-1234" 같은 값이 키워드로 오인되지 않게
Private Function HK_FuzzyFix(ByVal s As String, ByVal kw As String) As String
    Dim i As Long, j As Long, n As Long, W As String, diff As Long, ch As String, kc As String, bad As Boolean
    HK_FuzzyFix = s
    n = Len(kw)
    If n < 3 Then Exit Function
    If InStr(1, s, kw, vbTextCompare) > 0 Then Exit Function
    For i = 1 To Len(s) - n + 1
        W = Mid$(s, i, n)
        diff = 0: bad = False
        For j = 1 To n
            ch = Mid$(W, j, 1): kc = Mid$(kw, j, 1)
            If StrComp(ch, kc, vbTextCompare) <> 0 Then
                If (ch Like "#") Or (kc Like "#") Then bad = True: Exit For
                diff = diff + 1
                If diff > 1 Then Exit For
            End If
        Next j
        If Not bad And diff = 1 Then
            HK_FuzzyFix = Left$(s, i - 1) & kw & Mid$(s, i + n)
            Exit Function
        End If
    Next i
End Function

' 09-29: 열 너비 상한 60, 넘치는 글은 줄 바꿈(셀 맞춤)
Private Sub HK_CapWidths(ws As Worksheet)
    Dim c As Long, lastCol As Long
    On Error Resume Next
    lastCol = KW_LAST_COL
    If ws.Columns("A").ColumnWidth > 60 Then ws.Columns("A").ColumnWidth = 60
    For c = KW_FIRST_COL To lastCol
        If ws.Columns(c).ColumnWidth > 60 Then
            ws.Columns(c).ColumnWidth = 60
            ws.Range(ws.Cells(4, c), ws.Cells(ws.rows.Count, c).End(xlUp)).WrapText = True
        End If
    Next c
    ws.Range("A2").WrapText = False
    ws.Range(ws.Cells(4, 1), ws.Cells(ws.rows.Count, 1).End(xlUp)).WrapText = True
    ws.rows("4:" & ws.Cells(ws.rows.Count, 1).End(xlUp).row).VerticalAlignment = xlTop
    On Error GoTo 0
End Sub


