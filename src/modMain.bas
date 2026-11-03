Attribute VB_Name = "modMain"
Option Explicit
'=====================================================================
' modMain : 설치(버튼 생성) / 전체 새로고침 오케스트레이션  - V8 (추경 차수 추가 · 추경 점검표)
'=====================================================================
Public gAutoBusy As Boolean
Public gStep As String

'★ 최초 1회 실행: 메인 시트 + 버튼 생성  (Alt+F8 → 설치)
Public Sub 설치()
    SpeedOn "설치 중..."
    CfgSheet
    SetupMainSheet
    RemoveBlankDefaultSheets
    BW().Worksheets(SH_MAIN).Activate
    SpeedOff
    MsgBox "예산집행보드 V8이 설치되었습니다." & vbCrLf & vbCrLf & _
           "① 세출예산요구서(본예산) 선택 (부서명·담당자·연도 자동 인식) → ② 장표(102-B) 선택 → 메인 보드 자동 생성" & vbCrLf & _
           "추경이 있으면 [추경예산 추가]로 1차, 2차 ... 차수별로 추가합니다. (추가할 때마다 중복·연도·차수·장표 금액 점검)" & vbCrLf & _
           "요구서가 없으면 ② 실행 시 부서 목록에서 우리 부서를 고르면 됩니다.", vbInformation, "예산집행보드 V8"
End Sub

'새 통합문서의 기본 빈 시트(Sheet1/시트1 등) 제거
Public Sub RemoveBlankDefaultSheets()
    Dim i As Long, ws As Worksheet
    For i = BW().Worksheets.Count To 1 Step -1
        Set ws = BW().Worksheets(i)
        If ws.name <> SH_MAIN And BW().Worksheets.Count > 1 Then
            If (LCase(Left(ws.name, 5)) = "sheet" Or Left(ws.name, 2) = "시트") Then
                If Application.WorksheetFunction.CountA(ws.Cells) = 0 And ws.Shapes.Count = 0 Then
                    Application.DisplayAlerts = False
                    ws.Delete
                    Application.DisplayAlerts = True
                End If
            End If
        End If
    Next i
End Sub

'---------------------------------------------------------------- 리본(추가기능 xlam) 진입점: 탭 "UBIS예산확인"
Public Sub UB_Ribbon(control As IRibbonControl)
    UB_Run control.tag
End Sub

Public Sub UB_Run(ByVal tag As String)
    On Error GoTo EH

    If tag = "UB_새보드" Then
        U2_Start
        Exit Sub
    End If

    '활성 통합문서가 예산집행보드가 아니면 → 새 통합문서 + 메인 시트 + 버튼 생성
    Dim isNew As Boolean
    isNew = EnsureBoard()
    If tag = "UB_새보드" Then
        If Not isNew Then
            'V8: 지금 문서가 V7(또는 다른 추가기능)에서 만든 보드면 버튼을 V8 시트 버튼으로 바꿀지 먼저 묻는다
            If Not BoardButtonsCurrent(gBoardWB) Then
                If MsgBox("지금 문서는 이전 버전(V7 등)에서 만든 예산보드입니다." & vbCrLf & vbCrLf & _
                          "이 보드의 버튼을 V8 시트 버튼으로 바꿀까요?" & vbCrLf & _
                          "(추경예산 추가 · 마지막 추경 취소 · 추경 점검표 · ①② · 새로고침 · 세부산출내역 · 버전 정보)" & vbCrLf & vbCrLf & _
                          "예 = 이 보드 버튼 바꾸기 (자료는 그대로)   /   아니요 = 새 보드 만들기", _
                          vbYesNo + vbQuestion + vbDefaultButton1, "UBIS 예산확인 - V8 버튼") = vbYes Then
                    SetupMainSheet
                    gBoardWB.Worksheets(SH_MAIN).Activate
                    Exit Sub
                End If
            End If
            Set gBoardWB = Workbooks.Add
            설치
        End If
        Exit Sub
    End If
    Select Case tag
        Case "요구서선택": 요구서선택
        Case "장표선택": 장표선택
        Case "새로고침버튼": 새로고침버튼
        Case "메인으로": 메인으로
        Case "예산보드": 예산보드
        Case "세목별현황": 세목별현황
        Case "세부산출내역": 세부산출내역
        Case "인쇄": 인쇄
        Case "부서담당자": 부서담당자
        Case "단위사업카드": 단위사업카드
        Case "부서명단파일선택": 부서명단파일선택
        Case "추경추가": 추경예산추가                 'V8
        Case "추경점검": 추경점검표                   'V8
        Case "마지막추경취소": 마지막추경취소         'V8
        Case Else: MsgBox "알 수 없는 기능: " & tag, vbExclamation
    End Select
    Exit Sub
EH:
    MsgBox "UBIS예산확인 오류: " & Err.Description, vbCritical
End Sub

'보드 통합문서 확보. 활성 문서가 보드가 아니면 새 통합문서를 만들고 메인 시트·버튼 생성 → True 반환
Public Function EnsureBoard() As Boolean
    If ThisWorkbook.IsAddin Then
        If Not ActiveWorkbook Is Nothing Then
            If IsBoardWB(ActiveWorkbook) Then Set gBoardWB = ActiveWorkbook: Exit Function
        End If
        Set gBoardWB = Workbooks.Add
        gBoardWB.Activate
        설치
        EnsureBoard = True
    ElseIf Not SheetExists(SH_MAIN) Then
        설치
        EnsureBoard = True
    End If
End Function

'버튼: 데이터 새로고침
Public Sub 새로고침버튼()
    '모든 슬라이서 필터 해제(리셋) 후 새로고침
    DeleteOrClearSlicers "슬라이서_", False
    새로고침 True
End Sub

'전체 새로고침 (silent=True 면 완료 메시지 생략)
Public Sub 새로고침(Optional ByVal silent As Boolean = False)
    Dim cur As String
    If gAutoBusy Then Exit Sub
    If Not SheetExists(SH_REQ) And Not SheetExists(SH_B) And Not SheetExists(SH_A) Then
        If Not silent Then MsgBox "먼저 ① 세출예산요구서 → ② 장표(102-B)를 선택하세요.", vbExclamation
        Exit Sub
    End If
    gAutoBusy = True
    On Error GoTo EH
    SpeedOn "데이터 새로고침 중..."
    gStep = "구버전 시트 정리": CleanupLegacySheets
    gStep = "모델구성": BuildModel
    '입력 시트는 이미 만들어져 있을 때만 갱신 (입력값 보존)
    If SheetExists(SH_STAFF) Then gStep = "부서담당자 시트": RefreshStaffSheet False
    If SheetExists(SH_CARD) Then gStep = "단위사업카드 시트": RefreshCardSheet False
    gStep = "메인 보드": BuildMain
    gStep = "세부현황(세목별)": BuildGrouped
    If SheetExists(SH_CHART) Then gStep = "예산보드": BuildChartBoard
    If SheetExists(SH_CALC) Then gStep = "세부산출내역": BuildCalc
    If SheetExists(SH_CHECK) Then gStep = "추경 점검표": BuildCheckSheet     'V8
    gStep = "마무리"
    BW().Worksheets(SH_MAIN).Activate
    BW().Worksheets(SH_MAIN).Range("A1").Select
    CfgSet "마지막갱신", NowText()
    SpeedOff
    gAutoBusy = False
    If Not silent Then MsgBox "새로고침 완료 (" & gItems.Count & "개 사업항목)", vbInformation
    Exit Sub
EH:
    SpeedOff
    gAutoBusy = False
    MsgBox "새로고침 오류 [" & gStep & " 단계]: " & Err.Description & " (오류번호 " & Err.Number & ")", vbCritical
End Sub

'이전 버전에서 만들어진 시트 제거 (최종 시트: 메인 / 세부현황_세목별 / 부서담당자 / 단위사업카드)
Public Sub CleanupLegacySheets()
    Dim nm As Variant
    DeleteOrClearSlicers "슬라이서_현황_", True
    DeleteOrClearSlicers "슬라이서_월별_", True
    If SheetExists("_보드데이터") Then DeleteOrClearSlicers "슬라이서_차트_", True
    '시트 이름 변경 호환: 세부현황_세목별 → 세부현황
    If SheetExists("세부현황_세목별") And Not SheetExists(SH_GROUP) Then BW().Worksheets("세부현황_세목별").name = SH_GROUP
    If SheetExists("팀별분석") Or SheetExists("차트보드") Then DeleteOrClearSlicers "슬라이서_보드_", True
    If SheetExists("전년비교") Then DeleteOrClearSlicers "슬라이서_전년_", True
    For Each nm In Array("사업비현황", "사업상세", "세목현황", "슬라이서보드", "월별현황", "월별파일목록", "_월별데이터", "_보드데이터", "팀별분석", "차트보드", "전년비교", "원본_전년")
        DeleteSheetIfExists CStr(nm)
    Next nm
    On Error Resume Next
    BW().names("담당자목록").Delete
    On Error GoTo 0
End Sub

'가져오기 직후 자동 새로고침
Public Sub AutoRefresh()
    If SheetExists(SH_REQ) Or SheetExists(SH_B) Or SheetExists(SH_A) Then 새로고침 True
End Sub

'통합문서 열 때 자동 처리
Public Sub OnWorkbookOpen()
    If Not SheetExists(SH_MAIN) Then
        설치
    ElseIf SheetExists(SH_REQ) Or SheetExists(SH_B) Then
        If CfgGet("열때자동갱신") <> "N" Then 새로고침 True
    End If
End Sub

'재계산 이벤트: 차트보드에서 슬라이서(차트종류) 선택 → 인덱스 재계산 → 차트 회전
Public Sub OnSheetCalculate(Sh As Object)
    '현재 사용 안 함 (차트보드 제거)
End Sub

'시트 변경 이벤트 (현재는 사용 안 함 - 대시보드는 슬라이서로 동작)
Public Sub OnSheetChange(Sh As Object, Target As Range)
    If gAutoBusy Then Exit Sub
End Sub


