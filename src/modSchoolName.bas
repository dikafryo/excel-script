Attribute VB_Name = "modSchoolName"
Option Explicit
' =====================================================================================
'  modSchoolName : 리본 [교명 확인 ▶ 학교명·지원청] 진입점 (tag "학교명변환")
'   메뉴 1 학교명 정리 + 지원청 표시   2 학교명 정리만
'        3/4/5 초·중·고 현황시트 만들기   6 세 시트 모두
'   - 판정: modSchoolMatch, 데이터: modSchoolData, 시트: modSchoolSheet
'   - 실행 직전 활성 시트는 modRibbonUI 가 자동 백업(백업_…)한다.
'   [v2.6 속도 개선] 셀마다 하던 클립보드 복사(Copy/PasteSpecial)·셀 접근을 없애고
'     ① 값을 배열로 한 번에 읽어 메모리에서 판정 ② 바뀐 셀만 기록 ③ 지원청 서식은 연속 구간 단위로 한 번에 복사
'     ④ 계산·이벤트를 잠시 끔. (1,000행 기준 수십 초 → 1~2초)
' =====================================================================================
Private Const APP_TITLE As String = "학교명·지원청"
Private Const CLR_MARK_AMB As Long = &HB3F0FF     ' 확인 필요 (255,240,179) 연노랑
Private Const CLR_MARK_NF As Long = &HE0E0FF      ' 목록에 없음 (255,224,224) 연빨강
Public SN_Prof As String   ' 구간별 처리 시간(점검용)

Public Sub 학교명변환(control As IRibbonControl)
    Dim choice As Long
    choice = UI_Choose(APP_TITLE, _
        "줄여 쓴 학교명(○○초·○○여중·○○공고 …)을 정식 학교명으로 바꾸고, 오른쪽 셀에 교육지원청을 표시합니다." & vbCrLf & _
        "현황시트(초등학교현황·중학교현황·고등학교현황)가 이 문서에 있으면 그 시트의 학교명·지원청이 우선 적용됩니다.", _
        Array("학교명 정리 + 지원청 표시 (범위 선택)", _
              "학교명 정리만 (범위 선택)", _
              "초등학교 현황시트 만들기", _
              "중학교 현황시트 만들기", _
              "고등학교 현황시트 만들기", _
              "초·중·고 현황시트 모두 만들기"))
    Select Case choice
        Case 1: SN_RunOnRange True
        Case 2: SN_RunOnRange False
        Case 3: SN_MakeSheet 1
        Case 4: SN_MakeSheet 2
        Case 5: SN_MakeSheet 3
        Case 6: SS_CreateAll
    End Select
End Sub

Private Sub SN_MakeSheet(ByVal lv As Long)
    Dim ws As Worksheet
    Set ws = SS_Create(lv, True)
    If ws Is Nothing Then Exit Sub
    SD_Load True
    MsgBox "'" & ws.name & "' 시트 준비 완료 (" & SD_Count(lv) & "교)." & vbCrLf & vbCrLf & _
           "학교명·교육지원청·설립구분 열을 직접 고쳐 두면 이후 학교명 정리 때 우선 적용됩니다.", vbInformation, APP_TITLE
End Sub

' ---------------------------------------------------------------- 범위 실행
Private Sub SN_RunOnRange(ByVal withOffice As Boolean)
    Dim rng As Range
    Dim nExact As Long, nConv As Long, nAmb As Long, nNF As Long, nSkip As Long, nOff As Long
    Dim offMode As Long, srcInfo As String, msg As String, t0 As Single

    On Error Resume Next
    Set rng = Application.InputBox( _
        prompt:="학교명이 있는 범위를 마우스로 선택하세요." & vbCrLf & _
                "(열 머리글 클릭으로 열 전체 선택 가능 · 여러 열을 선택하면 첫 열만 처리)" & vbCrLf & vbCrLf & _
                "변환 예: 부산여상 → 부산여자상업고등학교,  개성중 → 부산개성중학교,  경남공고 → 경남공업고등학교", _
        title:=APP_TITLE & " - 범위 선택", Default:=Selection.Address, Type:=8)
    On Error GoTo 0
    If rng Is Nothing Then Exit Sub
    If rng.Columns.Count > 1 Then Set rng = rng.Columns(1)
    Set rng = Intersect(rng, rng.parent.UsedRange)
    If rng Is Nothing Then MsgBox "선택 범위에 데이터가 없습니다.", vbExclamation, APP_TITLE: Exit Sub
    If rng.Areas.Count > 1 Then Set rng = rng.Areas(1)

    t0 = Timer
    SD_Load True   ' 현황시트 유무를 매번 다시 확인
    srcInfo = "초 " & SD_Source(1) & " " & SD_Count(1) & "교 · 중 " & SD_Source(2) & " " & SD_Count(2) & "교 · 고 " & SD_Source(3) & " " & SD_Count(3) & "교"

    ' 지원청 열 처리 방식
    If withOffice Then
        offMode = SN_OfficeMode(rng)
        If offMode = 0 Then withOffice = False
        If offMode = 3 Then Exit Sub
        If offMode = 2 Then
            rng.Offset(0, 1).EntireColumn.Insert Shift:=xlToRight
            offMode = 1
        End If
    End If

    SN_Process rng, withOffice, nExact, nConv, nAmb, nNF, nSkip, nOff

    msg = "정리 완료  (" & Format$(Timer - t0, "0.0") & "초, " & rng.rows.Count & "행)" & vbCrLf & vbCrLf & _
          "이미 정식명: " & nExact & "개" & vbCrLf & _
          "변환됨: " & nConv & "개" & vbCrLf & _
          "확인 필요(노랑, 후보 여러 개): " & nAmb & "개" & vbCrLf & _
          "목록에 없음(빨강): " & nNF & "개" & vbCrLf & _
          "학교명 아님(그대로): " & nSkip & "개"
    If withOffice Then msg = msg & vbCrLf & "지원청 표시: " & nOff & "개"
    msg = msg & vbCrLf & vbCrLf & "데이터: " & srcInfo & vbCrLf & _
          "※ 변경 내역(이전 학교명 → 변경된 학교명)은 '변경내용' 시트에 있습니다. 되돌리려면 [백업 시트 복원]"
    MsgBox msg, vbInformation, APP_TITLE
End Sub

' 핵심 처리: 배열 판정 → 바뀐 셀만 기록 → 지원청 서식은 연속 구간 단위 복사
Public Sub SN_Process(ByVal rng As Range, ByVal withOffice As Boolean, ByRef nExact As Long, ByRef nConv As Long, _
                      ByRef nAmb As Long, ByRef nNF As Long, ByRef nSkip As Long, ByRef nOff As Long)
    Dim vals As Variant, n As Long, i As Long, raw As String, fullName As String, office As String, found As String, note As String
    Dim resCode() As Long, newVal() As String, offVal() As String, noteTxt() As String
    Dim cell As Range, anyFormula As Variant, anyMerged As Variant, anyColor As Variant, hasComments As Boolean
    Dim chkFormula As Boolean, chkMerged As Boolean, chkColor As Boolean, bulkFmt As Boolean
    Dim calcMode As XlCalculation, evMode As Boolean, scr As Boolean
    Dim runA As Long, Target As Boolean, ws As Worksheet, tA As Single, tB As Single, tc As Single
    Dim chg() As String, nChg As Long   ' 변경내용 시트 기록(2.7i: 메모 대신)

    Set ws = rng.parent
    n = rng.rows.Count
    If n = 1 Then
        ReDim vals(1 To 1, 1 To 1)
        vals(1, 1) = rng.Value2
    Else
        vals = rng.Value2
    End If
    ReDim resCode(1 To n): ReDim newVal(1 To n): ReDim offVal(1 To n): ReDim noteTxt(1 To n)

    tA = Timer
    ' 1) 메모리 판정 (시트 접근 없음)
    For i = 1 To n
        resCode(i) = -1
        If Not IsEmpty(vals(i, 1)) And Not IsError(vals(i, 1)) Then
            raw = CStr(vals(i, 1))
            If Len(Trim$(raw)) > 0 Then
                resCode(i) = SM_Resolve(raw, fullName, office, found, note)
                newVal(i) = fullName: offVal(i) = office
                Select Case resCode(i)
                    Case SM_EXACT: nExact = nExact + 1
                    Case SM_CONVERTED
                        nConv = nConv + 1
                        noteTxt(i) = "변환 전: " & raw & IIf(note <> "", vbLf & note, "")
                    Case SM_AMBIGUOUS
                        nAmb = nAmb + 1
                        noteTxt(i) = "확인 필요 - " & note
                    Case SM_NOTFOUND
                        nNF = nNF + 1
                        noteTxt(i) = note & vbLf & "(현황시트에 학교를 추가하면 다음부터 인식됩니다)"
                    Case Else: nSkip = nSkip + 1
                End Select
            End If
        End If
    Next i

    tB = Timer
    ' 2) 시트 반영
    scr = Application.ScreenUpdating: calcMode = Application.Calculation: evMode = Application.EnableEvents
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    On Error GoTo Done
    anyFormula = rng.HasFormula                 ' False 면 수식 없음 → 셀마다 확인할 필요 없음 (Null=섞임)
    chkFormula = True
    If Not IsNull(anyFormula) Then chkFormula = CBool(anyFormula)
    anyColor = rng.Interior.color               ' Null 이면 색이 섞여 있음
    chkColor = IsNull(anyColor)
    If Not chkColor Then chkColor = (anyColor = CLR_MARK_AMB Or anyColor = CLR_MARK_NF)
    hasComments = (ws.Comments.Count > 0)
    ' 2a) 수식 셀 제외 · 이전 표시(색·메모) 제거
    For i = 1 To n
        If resCode(i) >= 0 Then
            Set cell = rng.Cells(i, 1)
            If chkFormula Then
                If cell.HasFormula Then resCode(i) = -1
            End If
            If resCode(i) >= 0 And (hasComments Or chkColor) Then SN_ClearMark cell, hasComments, chkColor
        End If
    Next i
    ' 2b) 지원청 열 서식: 오른쪽 열에 병합이 없으면 학교명 열 서식을 통째로 1회 복사 (셀마다 복사하던 방식이 느렸던 원인)
    bulkFmt = False
    If withOffice Then
        anyMerged = rng.Offset(0, 1).MergeCells   ' False 면 병합 없음 (Null=섞임)
        chkMerged = True
        If Not IsNull(anyMerged) Then chkMerged = CBool(anyMerged)
        bulkFmt = Not chkMerged
        If bulkFmt Then
            rng.Copy
            rng.Offset(0, 1).PasteSpecial xlPasteFormats
            Application.CutCopyMode = False
        End If
    End If
    ' 2c) 변환값 · 확인 표시 · 메모 기록 (바뀐 셀만)
    For i = 1 To n
        If resCode(i) >= 0 Then
            Set cell = rng.Cells(i, 1)
            Select Case resCode(i)
                Case SM_CONVERTED
                    cell.value = newVal(i)
                    SN_LogChange chg, nChg, cell, CStr(vals(i, 1)), newVal(i), "변환", SN_AfterLf(noteTxt(i))
                Case SM_AMBIGUOUS
                    cell.Interior.color = CLR_MARK_AMB
                    SN_LogChange chg, nChg, cell, CStr(vals(i, 1)), "(확인 필요)", "확인 필요", SN_AfterSep(noteTxt(i))
                    If bulkFmt Then cell.Offset(0, 1).Interior.Pattern = xlNone
                Case SM_NOTFOUND
                    cell.Interior.color = CLR_MARK_NF
                    SN_LogChange chg, nChg, cell, CStr(vals(i, 1)), "(목록에 없음)", "목록에 없음", SN_FirstLine(noteTxt(i))
                    If bulkFmt Then cell.Offset(0, 1).Interior.Pattern = xlNone
            End Select
        End If
    Next i

    tc = Timer
    ' 3) 지원청 값: 대상(정식명·변환) 연속 구간마다 배열로 입력 (병합이 섞여 있으면 구간마다 서식도 복사)
    If withOffice Then
        runA = 0
        For i = 1 To n + 1
            Target = False
            If i <= n Then
                If resCode(i) = SM_EXACT Or resCode(i) = SM_CONVERTED Then
                    Target = True
                    If chkMerged Then
                        If rng.Cells(i, 2).MergeCells Then Target = False
                    End If
                End If
            End If
            If Target Then
                If runA = 0 Then runA = i
            ElseIf runA > 0 Then
                SN_WriteOfficeRun rng, runA, i - 1, offVal, Not bulkFmt
                nOff = nOff + (i - runA)
                runA = 0
            End If
        Next i
    End If
    SN_WriteChangeSheet ws, chg, nChg
Done:
    On Error Resume Next
    SN_Prof = "판정 " & Format$(tB - tA, "0.00") & "s / 기록 " & Format$(tc - tB, "0.00") & "s / 지원청 " & Format$(Timer - tc, "0.00") & "s"
    Application.CutCopyMode = False
    Application.Calculation = calcMode
    Application.EnableEvents = evMode
    Application.ScreenUpdating = scr
    On Error GoTo 0
End Sub

' 오른쪽 열 상태에 따라: 1 바로 쓰기 / 2 열 삽입 후 쓰기 / 3 취소 / 0 지원청 표시 안 함
Private Function SN_OfficeMode(ByVal rng As Range) As Long
    Dim v As Variant, conflict As Long, i As Long, s As String
    If rng.Column = rng.parent.Columns.Count Then SN_OfficeMode = 0: Exit Function
    If rng.rows.Count = 1 Then
        ReDim v(1 To 1, 1 To 1): v(1, 1) = rng.Offset(0, 1).Value2
    Else
        v = rng.Offset(0, 1).Value2
    End If
    For i = 1 To UBound(v, 1)
        If Not IsError(v(i, 1)) Then
            s = Trim$(CStr(v(i, 1)))
            If s <> "" Then
                If Not SD_IsOfficeText(s) Then conflict = conflict + 1
            End If
        End If
    Next i
    If conflict = 0 Then SN_OfficeMode = 1: Exit Function
    Select Case UI_Choose("오른쪽 열에 다른 값이 있습니다", "학교명 오른쪽 열에 지원청이 아닌 값이 " & conflict & "개 있습니다. 어떻게 할까요?", _
                          Array("오른쪽에 새 열을 끼워 넣고 지원청 표시", "그대로 덮어쓰기", "지원청 표시 없이 학교명만 정리", "취소"))
        Case 1: SN_OfficeMode = 2
        Case 2: SN_OfficeMode = 1
        Case 3: SN_OfficeMode = 0
        Case Else: SN_OfficeMode = 3
    End Select
End Function

' 연속 구간(a~b행)의 지원청 셀: (copyFmt 면 학교명 셀 서식을 복사한 뒤) 값 배열 입력
Private Sub SN_WriteOfficeRun(ByVal rng As Range, ByVal a As Long, ByVal b As Long, ByRef offVal() As String, ByVal copyFmt As Boolean)
    Dim src As Range, dst As Range, arr() As String, i As Long
    Set src = rng.Cells(a, 1).Resize(b - a + 1, 1)
    Set dst = src.Offset(0, 1)
    If copyFmt Then
        src.Copy
        dst.PasteSpecial xlPasteFormats
        Application.CutCopyMode = False
    End If
    dst.NumberFormat = "@"
    ReDim arr(1 To b - a + 1, 1 To 1)
    For i = a To b
        arr(i - a + 1, 1) = offVal(i)
    Next i
    dst.value = arr
End Sub

' ---- 2.7i: 변경내용 시트 ------------------------------------------------------
Private Sub SN_LogChange(ByRef chg() As String, ByRef nChg As Long, ByVal cell As Range, ByVal oldV As String, ByVal newV As String, ByVal kind As String, ByVal memo As String)
    nChg = nChg + 1
    ReDim Preserve chg(1 To 5, 1 To nChg)
    chg(1, nChg) = cell.Address(False, False)
    chg(2, nChg) = oldV
    chg(3, nChg) = newV
    chg(4, nChg) = kind
    chg(5, nChg) = memo
End Sub

Private Function SN_AfterLf(ByVal T As String) As String
    Dim k As Long: k = InStr(T, vbLf)
    If k > 0 Then SN_AfterLf = Replace(Mid$(T, k + 1), vbLf, " / ") Else SN_AfterLf = ""
End Function

Private Function SN_AfterSep(ByVal T As String) As String
    Dim k As Long: k = InStr(T, " - ")
    If k > 0 Then SN_AfterSep = Replace(Mid$(T, k + 3), vbLf, " / ") Else SN_AfterSep = Replace(T, vbLf, " / ")
End Function

Private Function SN_FirstLine(ByVal T As String) As String
    Dim k As Long: k = InStr(T, vbLf)
    If k > 0 Then SN_FirstLine = Left$(T, k - 1) Else SN_FirstLine = T
End Function

' '변경내용' 시트: 이전 학교명 → 변경된 학교명 (셀 주소는 원본 셀로 가는 링크). 이미 있으면 내용을 새로 씀
Private Sub SN_WriteChangeSheet(ByVal ws As Worksheet, ByRef chg() As String, ByVal nChg As Long)
    Dim wb As Workbook, cs As Worksheet, i As Long, arr As Variant, actSh As Object
    Set wb = ws.parent
    Set actSh = ActiveSheet
    On Error Resume Next
    Set cs = wb.Worksheets("변경내용")
    On Error GoTo 0
    If cs Is Nothing Then
        Set cs = wb.Worksheets.Add(After:=ws)
        cs.name = "변경내용"
    Else
        cs.Cells.Clear
        cs.Hyperlinks.Delete
    End If
    cs.Range("A1:F1").value = Array("No.", "셀", "이전 학교명", "변경된 학교명", "구분", "비고(후보·안내)")
    With cs.Range("A1:F1")
        .Font.Bold = True
        .Interior.color = RGB(221, 235, 247)
        .HorizontalAlignment = xlCenter
    End With
    If nChg > 0 Then
        ReDim arr(1 To nChg, 1 To 6)
        For i = 1 To nChg
            arr(i, 1) = i
            arr(i, 2) = chg(1, i)
            arr(i, 3) = chg(2, i)
            arr(i, 4) = chg(3, i)
            arr(i, 5) = chg(4, i)
            arr(i, 6) = chg(5, i)
        Next i
        cs.Range("A2").Resize(nChg, 6).value = arr
        For i = 1 To nChg
            cs.Hyperlinks.Add anchor:=cs.Cells(i + 1, 2), Address:="", SubAddress:="'" & ws.name & "'!" & chg(1, i), TextToDisplay:=chg(1, i)
            Select Case chg(4, i)
                Case "확인 필요": cs.Cells(i + 1, 5).Interior.color = CLR_MARK_AMB
                Case "목록에 없음": cs.Cells(i + 1, 5).Interior.color = CLR_MARK_NF
            End Select
        Next i
        cs.Range("A1:F1").Resize(nChg + 1, 6).Borders.LineStyle = xlContinuous
    Else
        cs.Range("A2").value = "(변경된 학교명 없음)"
    End If
    cs.Columns("A:F").AutoFit
    If cs.Columns("F").ColumnWidth > 60 Then cs.Columns("F").ColumnWidth = 60
    On Error Resume Next
    actSh.Activate
    On Error GoTo 0
End Sub

Private Sub SN_Note(ByVal cell As Range, ByVal text As String)
    Dim ln As Variant, maxLen As Long, nLines As Long, W As Single
    On Error Resume Next
    If Not cell.Comment Is Nothing Then cell.Comment.Delete
    cell.AddComment text
    For Each ln In Split(text, vbLf)
        nLines = nLines + 1
        If Len(ln) > maxLen Then maxLen = Len(ln)
    Next ln
    W = maxLen * 8 + 12
    If W < 120 Then W = 120
    If W > 420 Then W = 420: nLines = nLines + (maxLen * 8) \ 400
    With cell.Comment.Shape
        .Width = W
        .Height = nLines * 15 + 8
    End With
    On Error GoTo 0
End Sub

Private Sub SN_ClearMark(ByVal cell As Range, ByVal chkComment As Boolean, ByVal chkColor As Boolean)
    On Error Resume Next
    If chkColor Then
        If cell.Interior.color = CLR_MARK_AMB Or cell.Interior.color = CLR_MARK_NF Then cell.Interior.Pattern = xlNone
    End If
    If chkComment Then
        If Not cell.Comment Is Nothing Then
            If Left$(cell.Comment.text, 4) = "변환 전" Or Left$(cell.Comment.text, 5) = "확인 필요" Or InStr(cell.Comment.text, "목록에 없음") > 0 Then cell.Comment.Delete
        End If
    End If
    On Error GoTo 0
End Sub

' ---------------------------------------------------------------- 자체 점검(빌드 검증용, 리본 미연결)
Public Function SN_SelfTest() As String
    Dim tests As Variant, T As Variant, fullName As String, office As String, found As String, note As String, res As Long, s As String
    SD_Load True
    tests = Array("부산여상", "개성중", "경남공고", "동래여고", "가남초", "해운대관광고", "부산과고", "부산기계공고", "비고", "장전초등학교", "없는학교고", "부산고", "한국영재학교", "정관2중")
    For Each T In tests
        res = SM_Resolve(CStr(T), fullName, office, found, note)
        s = s & T & " → [" & res & "] " & fullName & " / " & office & " " & IIf(note <> "", "(" & note & ")", "") & vbLf
    Next T
    SN_SelfTest = s
End Function

Public Function SN_Probe(ByVal raw As String) As String
    Dim fullName As String, office As String, found As String, note As String, res As Long
    SD_Load True
    res = SM_Resolve(raw, fullName, office, found, note)
    SN_Probe = raw & " → [" & res & "] " & fullName & " / " & office & " / " & found & IIf(note <> "", " (" & note & ")", "")
End Function

' 빌드 검증용: 현황시트 생성 → 결과 문자열(오류 시 오류 내용)
Public Function SN_TestSheet(ByVal lv As Long) As String
    Dim ws As Worksheet
    On Error GoTo Failed
    Set ws = SS_Create(lv, False)
    SN_TestSheet = ws.name & " | " & ws.UsedRange.Address & " | " & ws.Range("A1").value & " | " & ws.Range("B5").value & " / " & ws.Range("C5").value & " / " & ws.Range("D5").value & " / " & ws.Range("E5").value
    Exit Function
Failed:
    SN_TestSheet = "ERR " & Err.Number & ": " & Err.Description & " @" & SS_LastStep
End Function

' 빌드 검증용: 범위에 바로 실행(폼 없이). 반환 "정식/변환/확인/없음/건너뜀/지원청 (초)"
Public Function SN_TestRange(ByVal rng As Range, ByVal withOffice As Boolean) As String
    Dim a As Long, b As Long, c As Long, d As Long, e As Long, f As Long, t0 As Single
    t0 = Timer
    SD_Load True
    SN_Process rng, withOffice, a, b, c, d, e, f
    SN_TestRange = a & "/" & b & "/" & c & "/" & d & "/" & e & "/" & f & " (" & Format$(Timer - t0, "0.00") & "s: " & SN_Prof & ")"
End Function

