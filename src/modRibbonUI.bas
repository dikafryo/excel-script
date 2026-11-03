Attribute VB_Name = "modRibbonUI"
Option Explicit
' =====================================================================================
'  modRibbonUI : 리본·도형 버튼 공통 진입점
'   - 리본의 모든 버튼은 UI_Ribbon(tag) 하나로 들어와 UI_Run이 실행을 담당한다.
'   - 실행 전 계산방식·이벤트·화면갱신·알림·상태표시줄·자동화보안을 보관하고 종료 시 복원한다.
'   - 삭제·치환 계열 기능은 실행 직전에 활성 시트를 숨김 백업 시트로 복사한다(최근 3개 보관).
' =====================================================================================
Public Const APP_VERSION As String = "8.1"
Public Const APP_BUILD As String = "2026-10-17"
Private Const CFG_AppTitle As String = "B-XLAM V8"
Public BX_Silent As Boolean   ' 검증용: True 면 선택 폼을 띄우지 않고 0(취소) 반환
Public BX_Trace As String    ' 진단용: 긴 작업의 현재 단계
Private Const BACKUP_PREFIX As String = "백업_"
Private Const BACKUP_KEEP As Long = 3
Private busy As Boolean
Public UI_LastChoice As Long   ' UI_Choose 마지막 선택(0=취소)

' 도형 OnAction에 넣는 매크로 이름. 추가 기능 파일명이 바뀌어도 동작하도록 파일명을 붙이지 않는다.
Public Function UI_Macro(ByVal proc As String) As String
    UI_Macro = proc
End Function

Public Sub UI_Ribbon(control As IRibbonControl)
    UI_Run control.tag, True
End Sub

'V8: 버전 정보 (리본 버튼은 없음 - 예산보드 시트 [버전 정보] 버튼에서 호출, 이전 리본 콜백 이름도 유지)
Public Sub BX_AboutRibbon(control As IRibbonControl)
    버전정보
End Sub

'V8: 예산보드 시트 [버전 정보] 버튼
Public Sub 버전정보()
    MsgBox "B-XLAM V" & APP_VERSION & "  (빌드 " & APP_BUILD & ")" & vbCrLf & vbCrLf & _
           "V8 새 기능" & vbCrLf & _
           " - 추경예산 추가: 1차, 2차, 3차 ... 차수 제한 없이 추가, 차수별 색 구분" & vbCrLf & _
           " - 차수는 파일 내용(제목 · 기정액)으로 판단해 차수 순서로 정렬 (추가 순서 · 파일 이름 무관)" & vbCrLf & _
           " - 추가할 때마다 중복 파일 · 연도 · 차수 · 장표 금액 점검 (추경 점검표 한 장에 차수별 정리)" & vbCrLf & _
           " - 세부산출내역 머리글 연도 자동 인식 (예: 2027년 산출내역 / 2026년 산출내역), 메모에 '1추' · '2추' 표시" & vbCrLf & _
           " - 장표 대조 · 산출내역 비교: 비고에 사유 + 맨 위 '확인 필요 n건' (종류별 건수)" & vbCrLf & _
           " - 리본 [교명 확인] > [수신처 자동체크]: 엑셀 학교명을 업무포털 [수신자지정] 창에서 체크 · 수신자 추가" & vbCrLf & _
           " - 리본은 V7과 같고([수신처 자동체크] 하나 추가), 추경 기능은 [UBIS예산확인]으로 만든 예산보드 시트의 버튼에 있습니다" & vbCrLf & vbCrLf & _
           "만든 곳: 엑셀과 한글 매크로", vbInformation, CFG_AppTitle
End Sub

Public Sub UI_ShapeClick()
    Dim s As Shape
    On Error GoTo Failed
    Set s = ActiveSheet.Shapes(CStr(Application.Caller))
    UI_Run s.AlternativeText, False
    Exit Sub
Failed:
    MsgBox "버튼 연결을 확인해 주세요: " & Err.Description, vbExclamation, CFG_AppTitle
End Sub

Public Sub UI_Run(ByVal proc As String, Optional ByVal ribbonCall As Boolean = False)
    Dim calc As XlCalculation, events As Boolean, screen As Boolean, alerts As Boolean
    Dim status As Variant, security As Long, code As Long, detail As String
    If busy Then Exit Sub
    BX_EnsureBoardEvents   ' 2.7m 예산보드 빈 행 자동 숨김 이벤트
    If ActiveWorkbook Is Nothing Then Workbooks.Add
    If ActiveWorkbook Is ThisWorkbook Then Workbooks.Add
    If TypeName(ActiveSheet) <> "Worksheet" Then
        MsgBox "워크시트를 선택한 뒤 실행하세요.", vbExclamation, CFG_AppTitle
        Exit Sub
    End If
    calc = Application.Calculation: events = Application.EnableEvents
    screen = Application.ScreenUpdating: alerts = Application.DisplayAlerts
    status = Application.StatusBar: security = Application.AutomationSecurity
    busy = True
    On Error GoTo Failed
    Application.AutomationSecurity = msoAutomationSecurityForceDisable
    If UI_IsDestructive(proc) Then UI_MakeBackup ActiveSheet
    If UI_NeedsNewFile(proc) Then
        ' 폴더/파일 정리: 현재 문서에 '기준폴더' 시트(저장해 둔 기준폴더 파일)가 있으면 그 문서에서 실행
        If Not (proc = "폴더분석정리" And UI_HasSheet(ActiveWorkbook, "기준폴더")) Then UI_OpenNewFile proc
    End If
    Select Case proc
        Case "시트합치기": 시트합치기 Nothing
        Case "MergeSheetsToCollection": MergeSheetsToCollection Nothing
        Case "b빈칸삭제": b빈칸삭제 Nothing
        Case "c중복셀값제거": c중복셀값제거 Nothing
        Case "p001_폴더선택": p001_폴더선택 Nothing
        Case "폴더분석정리": 폴더분석정리 Nothing
        Case "셀자동맞춤2": 셀자동맞춤2 Nothing
        Case "HwpCopyText": HwpCopyText Nothing
        Case "HwpCopyTable": HwpCopyTable Nothing
        Case "Fake병합": Fake병합 Nothing
        Case "병합취소": 병합취소 Nothing
        Case "셀색표시": 셀색표시 Nothing
        Case "짝수행연한파란색": 짝수행연한파란색 Nothing
        Case "짝수행연한노란색": 짝수행연한노란색 Nothing
        Case "테두리선스타일": 테두리선스타일 Nothing
        Case "MergeHwpInFolder": MergeHwpInFolder Nothing
        Case "dHWP_to_PDF_Converter": dHWP_to_PDF_Converter Nothing
        Case "eMergePdfsInFolder": eMergePdfsInFolder Nothing
        Case "OfficeMergeToPdf": OfficeMergeToPdf Nothing
        Case "한글계산1차": 한글계산1차 Nothing
        Case "표계산": 표계산 Nothing
        Case "메모에이미지넣기": 메모에이미지넣기 Nothing
        Case "학교명변환": 학교명변환 Nothing
        Case "OpenHelpWebsite": OpenHelpWebsite Nothing
        Case "UI_RestoreBackup": UI_RestoreBackup
        Case "HK_Start": HK_Start Nothing
        Case "HK_PickFolder": HK_PickFolder
        Case "HK_Extract": HK_Extract
        Case "p001_목록생성": p001_목록생성
        Case "p002_이름가져오기": p002_이름가져오기
        Case "p_단어입력추가": p_단어입력추가
        Case "p003_이름새로설정": p003_이름새로설정
        Case "p004_이름변경": p004_이름변경
        Case "FC_파일이동실행": FC_파일이동실행
        Case "FC_기준폴더만들기": FC_기준폴더만들기
        Case "FC_기준폴더이동": FC_기준폴더이동
        Case "FC_기준폴더설정": FC_기준폴더설정
        Case "FC_분석_폴더": FC_분석_폴더
        Case "FC_분석_파일": FC_분석_파일
        Case "FC_분석_폴더및파일": FC_분석_폴더및파일
        Case "p005_중복단어삭제": p005_중복단어삭제
        Case "FC_날짜키워드변경": FC_날짜키워드변경
        Case "FC_키워드월일변경": FC_키워드월일변경
        Case "FC_년월일키워드변경": FC_년월일키워드변경
        Case "FC_이름변경후이동": FC_이름변경후이동
        Case "FC_미분류오늘날짜추천": FC_미분류오늘날짜추천
        Case "FC_이름변경실행": FC_이름변경실행
        Case "HwpCopyMenu": HwpCopyMenu Nothing
        Case "HwpCalcMenu": 한글계산확인 Nothing
        Case "OfficeMergeMenu": OfficeMergeMenu Nothing
        Case "CellColorMenu": CellColorMenu Nothing
        Case "FakeMergeMenu": FakeMergeMenu Nothing
        Case "p005_중복단어적용": p005_중복단어적용
        Case Else: Err.Raise vbObjectError + 221, , "등록되지 않은 버튼입니다: " & proc
    End Select
CleanExit:
    On Error Resume Next
    Application.Calculation = calc
    Application.EnableEvents = events
    Application.ScreenUpdating = screen
    Application.DisplayAlerts = alerts
    Application.StatusBar = status
    Application.AutomationSecurity = security
    busy = False
    On Error GoTo 0
    If code <> 0 Then MsgBox "작업: " & proc & vbCrLf & "오류 " & code & ": " & detail & IIf(Len(BX_Trace) > 0, vbCrLf & "단계: " & BX_Trace, ""), vbExclamation, CFG_AppTitle
    Exit Sub
Failed:
    code = Err.Number: detail = Err.Description
    Resume CleanExit
End Sub

' ---------------------------------------------------------------- 새 파일에서 실행
' [v2.3] 버튼(작업 시트)을 만드는 기능과 한글 복사값을 계산하는 기능은 새 통합 문서를 열어 그 안에서 실행한다.
'        (양식·서식 변경형 기능은 현재 시트에서 그대로 실행)
Private Function UI_NeedsNewFile(ByVal proc As String) As Boolean
    Select Case proc
        Case "HK_Start", "p001_폴더선택", "폴더분석정리", "한글계산1차", "표계산"
            UI_NeedsNewFile = True
    End Select
End Function

Private Function UI_HasSheet(ByVal wb As Workbook, ByVal nm As String) As Boolean
    Dim ws As Worksheet
    If wb Is Nothing Then Exit Function
    On Error Resume Next
    Set ws = wb.Worksheets(nm)
    On Error GoTo 0
    UI_HasSheet = Not ws Is Nothing
End Function

Private Sub UI_OpenNewFile(ByVal proc As String)
    Dim wb As Workbook
    Set wb = Workbooks.Add
    wb.Activate
    On Error Resume Next
    Application.StatusBar = "새 통합 문서에서 실행: " & proc
    On Error GoTo 0
End Sub

' ---------------------------------------------------------------- 안전 백업
Private Function UI_IsDestructive(ByVal proc As String) As Boolean
    Select Case proc
        Case "b빈칸삭제", "c중복셀값제거", "학교명변환"
            UI_IsDestructive = True
    End Select
End Function

Private Sub UI_MakeBackup(ByVal ws As Worksheet)
    Dim wb As Workbook, copyWs As Worksheet, nm As String
    If Left$(ws.name, Len(BACKUP_PREFIX)) = BACKUP_PREFIX Then Exit Sub
    On Error GoTo skip
    Set wb = ws.parent
    nm = BACKUP_PREFIX & Left$(ws.name, 18) & "_" & Format$(Now, "hhmmss")
    Dim su As Boolean: su = Application.ScreenUpdating
    Application.ScreenUpdating = False
    ws.Copy After:=wb.Worksheets(wb.Worksheets.Count)
    Set copyWs = wb.Worksheets(wb.Worksheets.Count)
    copyWs.name = nm
    copyWs.Visible = xlSheetHidden
    ws.Activate
    UI_TrimBackups wb
    Application.ScreenUpdating = su   '2.7j: 백업 뒤 화면 갱신 복원(꺼진 채 범위 InputBox 가 뜨면 선택 표시가 안 보임)
    Exit Sub
skip:
    On Error Resume Next
    ws.Activate
    Application.ScreenUpdating = su
End Sub

Private Sub UI_TrimBackups(ByVal wb As Workbook)
    Dim i As Long, n As Long, alerts As Boolean
    For i = 1 To wb.Worksheets.Count
        If Left$(wb.Worksheets(i).name, Len(BACKUP_PREFIX)) = BACKUP_PREFIX Then n = n + 1
    Next i
    alerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    i = 1
    Do While n > BACKUP_KEEP And i <= wb.Worksheets.Count
        If Left$(wb.Worksheets(i).name, Len(BACKUP_PREFIX)) = BACKUP_PREFIX Then
            wb.Worksheets(i).Delete
            n = n - 1
        Else
            i = i + 1
        End If
    Loop
    Application.DisplayAlerts = alerts
End Sub

Public Sub UI_RestoreBackup()
    Dim wb As Workbook, ws As Worksheet, latest As Worksheet, orig As Worksheet, n As Long, choice As Long, stem As String
    Set wb = ActiveWorkbook
    For Each ws In wb.Worksheets
        If Left$(ws.name, Len(BACKUP_PREFIX)) = BACKUP_PREFIX Then
            n = n + 1
            Set latest = ws
        End If
    Next ws
    If latest Is Nothing Then
        MsgBox "이 통합 문서에는 자동 백업 시트가 없습니다." & vbCrLf & vbCrLf & _
               "빈칸삭제·중복셀값제거·학교명변환을 실행하면" & vbCrLf & _
               "실행 직전 상태가 숨김 시트(" & BACKUP_PREFIX & "…)로 자동 저장됩니다.", vbInformation, CFG_AppTitle
        Exit Sub
    End If
    ' 백업 이름: 백업_<원본 앞 18자>_hhmmss → 원본 시트 찾기
    stem = Mid$(latest.name, Len(BACKUP_PREFIX) + 1)
    If InStrRev(stem, "_") > 0 Then stem = Left$(stem, InStrRev(stem, "_") - 1)
    For Each ws In wb.Worksheets
        If Left$(ws.name, Len(BACKUP_PREFIX)) <> BACKUP_PREFIX Then
            If Left$(ws.name, 18) = stem Then Set orig = ws: Exit For
        End If
    Next ws
    choice = UI_Choose("백업 시트 복원", "가장 최근 백업: " & latest.name & IIf(orig Is Nothing, "  (원본 시트를 찾지 못함)", "  → 원본: " & orig.name) & vbCrLf & _
                       "(보관 중인 백업 " & n & "개, 최근 " & BACKUP_KEEP & "개까지 보관)", _
                       Array("백업 시트만 표시 (비교용)", "원본 시트에 백업 내용 덮어쓰기 (복원)", "취소"))
    Select Case choice
        Case 1
            latest.Visible = xlSheetVisible
            latest.Activate
            MsgBox "백업 시트를 표시했습니다: " & latest.name & vbCrLf & "필요한 내용을 복사해 되돌린 뒤 백업 시트는 삭제해도 됩니다.", vbInformation, CFG_AppTitle
        Case 2
            If orig Is Nothing Then
                MsgBox "원본 시트를 찾지 못해 자동 복원할 수 없습니다. [백업 시트만 표시]로 확인해 주세요.", vbExclamation, CFG_AppTitle
                Exit Sub
            End If
            Application.ScreenUpdating = False
            orig.Cells.Clear
            latest.Cells.Copy orig.Cells(1, 1)
            Application.CutCopyMode = False
            orig.Activate
            orig.Range("A1").Select
            Application.ScreenUpdating = True
            MsgBox "'" & orig.name & "' 시트를 백업(" & latest.name & ") 내용으로 되돌렸습니다.", vbInformation, CFG_AppTitle
    End Select
End Sub
' ---------------------------------------------------------------- 선택 폼 (숫자 입력 대신 마우스 클릭)
' options: 문자열 배열(최대 6개). 반환: 선택 번호(1~), 취소/닫기=0
' [v2.3] 애플 스타일 카드 목록(왼쪽 정렬), 키보드 1~6·Esc 지원. 레이아웃은 frmChoice.Setup / modUITheme 담당
Public Function UI_Choose(ByVal title As String, ByVal prompt As String, ByVal options As Variant) As Long
    Dim f As frmChoice
    If BX_Silent Then Exit Function
    Set f = New frmChoice
    f.Setup title, prompt, options
    f.Show vbModal
    UI_Choose = f.Result
    UI_LastChoice = f.Result
    Unload f
End Function

' 기존 MsgBox(예/아니오/취소) 코드와 호환: 1→vbYes, 2→vbNo, 3 또는 취소→vbCancel
Public Function UI_ChoiceToMsg(ByVal choice As Long) As VbMsgBoxResult
    Select Case choice
        Case 1: UI_ChoiceToMsg = vbYes
        Case 2: UI_ChoiceToMsg = vbNo
        Case Else: UI_ChoiceToMsg = vbCancel
    End Select
End Function

' ---------------------------------------------------------------- 도형 버튼 스타일
' [v2.3] fillRGB/textRGB 를 주면 그 색으로(버튼 그룹 구분용). 예: 초록 RGB(52,168,83)/흰색, 주황 RGB(255,149,0)/흰색
Public Sub UI_StyleButton(ByVal s As Shape, ByVal caption As String, ByVal proc As String, Optional ByVal primary As Boolean = False, _
                          Optional ByVal fillRGB As Long = -1, Optional ByVal textRGB As Long = -1)
    With s
        .OnAction = UI_Macro("UI_ShapeClick")
        .AlternativeText = proc
        .Placement = xlFreeFloating
        .Height = 30
        .line.Visible = msoFalse
        .Shadow.Visible = msoFalse
        .Fill.Solid
        If fillRGB >= 0 Then
            .Fill.ForeColor.RGB = fillRGB
        ElseIf primary Then
            .Fill.ForeColor.RGB = RGB(26, 64, 103)
        Else
            .Fill.ForeColor.RGB = RGB(232, 239, 247)
        End If
        With .TextFrame2
            .AutoSize = msoAutoSizeNone
            .WordWrap = msoFalse
            .MarginLeft = 10: .MarginRight = 10
            .MarginTop = 3: .MarginBottom = 3
            .VerticalAnchor = msoAnchorMiddle
            With .TextRange
                .text = caption
                .Font.name = "맑은 고딕"
                .Font.NameFarEast = "맑은 고딕"
                .Font.Size = 10
                .Font.Bold = msoTrue
                .ParagraphFormat.Alignment = msoAlignCenter
                If textRGB >= 0 Then
                    .Font.Fill.ForeColor.RGB = textRGB
                ElseIf primary Then
                    .Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
                Else
                    .Font.Fill.ForeColor.RGB = RGB(26, 64, 103)
                End If
            End With
        End With
    End With
End Sub


' ---------------------------------------------------------------- 이름 검사
Public Function UI_ValidFolderName(ByVal value As String) As Boolean
    Dim i As Long, stem As String
    value = Trim$(value)
    If Len(value) = 0 Or value = "." Or value = ".." Then Exit Function
    If Right$(value, 1) = "." Or Right$(value, 1) = " " Then Exit Function
    For i = 1 To Len(value)
        If AscW(Mid$(value, i, 1)) >= 0 And AscW(Mid$(value, i, 1)) < 32 Then Exit Function
        If InStr("\/:*?""<>|", Mid$(value, i, 1)) > 0 Then Exit Function
    Next i
    stem = UCase$(Split(value, ".")(0))
    Select Case stem
        Case "CON", "PRN", "AUX", "NUL", "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9", "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9"
            Exit Function
    End Select
    UI_ValidFolderName = True
End Function

' ---------------------------------------------------------------- 입력 폼 (InputBox 대신, 애플 스타일)
' 반환: 입력 문자열(Trim). 취소/닫기 = "" (빈 문자열)   Enter=확인, Esc=취소
Public Function UI_Input(ByVal title As String, ByVal prompt As String, Optional ByVal defaultText As String = "", Optional ByVal numericOnly As Boolean = False) As String
    Dim f As frmInput
    Set f = New frmInput
    f.Setup title, prompt, defaultText, numericOnly
    f.Show vbModal
    If f.Cancelled Then
        UI_Input = ""
    Else
        UI_Input = Trim$(f.txtValue.text)
    End If
    Unload f
End Function

' 폼 구성 자체 점검(표시 없이 컨트롤 생성만) - 빌드 검증용
Public Function UI_FormSelfTest() As String
    Dim f As frmChoice, g As frmInput
    Set f = New frmChoice
    f.Setup "테스트 제목", "안내문 줄입니다." & vbCrLf & "둘째 줄", Array("첫 번째 항목", "두 번째 항목", "세 번째 항목 - 조금 더 긴 문구가 들어가는 경우 두 줄로 접힘 확인용 텍스트입니다", "넷째")
    UI_FormSelfTest = "frmChoice H=" & f.Height & " W=" & f.Width & " ctrls=" & f.Controls.Count
    Unload f
    Set g = New frmInput
    g.Setup "입력 테스트", "값을 입력하세요.", "기본값", False
    UI_FormSelfTest = UI_FormSelfTest & " | frmInput H=" & g.Height & " W=" & g.Width & " ctrls=" & g.Controls.Count
    Unload g
End Function

' [v2.6] 미리보기·체크 항목이 있는 선택 폼. previews/extras 는 배열 또는 Empty. extraOut ← 체크 항목 번호(1~)
Public Function UI_ChooseEx(ByVal title As String, ByVal prompt As String, ByVal options As Variant, ByVal previews As Variant, _
                            ByVal extras As Variant, ByVal extraDefault As Long, ByRef extraOut As Long) As Long
    Dim f As frmChoice
    If BX_Silent Then Exit Function
    Set f = New frmChoice
    f.SetupEx title, prompt, options, previews, extras, extraDefault
    f.Show vbModal
    UI_ChooseEx = f.Result
    extraOut = f.Extra
    UI_LastChoice = f.Result
    Unload f
End Function

' 폼 구성 자체 점검(표시 없이) - 미리보기·체크 항목 포함
Public Function UI_FormSelfTestEx() As String
    Dim f As frmChoice
    Set f = New frmChoice
    f.SetupEx "테스트", "미리보기 점검", Array("실선", "점선", "짧은 점선", "점-실선", "이중선", "없음"), _
              Array("line:solid", "line:dot", "line:dash", "line:dashdot", "line:double", "line:none"), Array("얇게", "중간", "굵게"), 2
    UI_FormSelfTestEx = "line H=" & f.Height & " ctrls=" & f.Controls.Count & " extra=" & f.Extra
    f.OnExtra 3
    UI_FormSelfTestEx = UI_FormSelfTestEx & " →extra=" & f.Extra & " ctrls=" & f.Controls.Count
    Unload f
    Set f = New frmChoice
    f.SetupEx "테스트2", "", Array("조건", "지우기", "파랑", "노랑", "핑크", "녹색"), Array("", "", "stripe:DDEBF7", "stripe:FFFF99", "stripe:FCE4EC", "stripe:E2EFDA"), Empty, 0
    UI_FormSelfTestEx = UI_FormSelfTestEx & " | stripe H=" & f.Height & " ctrls=" & f.Controls.Count
    Unload f
    Set f = New frmChoice
    f.SetupEx "테스트3", "설명", Array("적용", "취소"), Array("merge:after", "merge:before"), Empty, 0
    UI_FormSelfTestEx = UI_FormSelfTestEx & " | merge H=" & f.Height & " ctrls=" & f.Controls.Count
    Unload f
End Function


