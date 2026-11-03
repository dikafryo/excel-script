Attribute VB_Name = "modCellColorMenu"
Option Explicit
' =====================================================================================
'  modCellColorMenu : 리본 [셀정리하기 ▶ 셀색표시] 통합 메뉴 (tag "CellColorMenu")  [v2.6]
'   1 조건에 맞는 행만 색 표시 (modCellColorMark.셀색표시)
'   2 색 지우기 (채우기색·조건부서식 제거)
'   3~6 교차형 파랑·노랑·핑크·연한녹색 (modAltRowBlue.AR_Stripe) - 폼에 미리보기 표시
' =====================================================================================
Private Const APP_TITLE As String = "셀색표시"

Public Sub CellColorMenu(control As IRibbonControl)
    Dim choice As Long, dummy As Long
    choice = UI_ChooseEx(APP_TITLE, _
        "표에 색을 입히는 방법을 고르세요. (셀 1개만 선택했으면 연결된 표 전체가 대상)", _
        Array("조건에 맞는 행만 색 표시  (기준열 · 조건 · 색 선택)", _
              "색 지우기  (선택 범위의 채우기색 · 조건부서식 제거)", _
              "교차형 파랑  (짝수 행)", _
              "교차형 노랑  (짝수 행)", _
              "교차형 핑크  (짝수 행)", _
              "교차형 연한녹색  (짝수 행)"), _
        Array("", "", "stripe:DDEBF7", "stripe:FFFF99", "stripe:FCE4EC", "stripe:E2EFDA"), Empty, 0, dummy)
    Select Case choice
        Case 1: 셀색표시 Nothing
        Case 2: CC_ClearColors
        Case 3: AR_Stripe RGB(221, 235, 247), "연한 파랑"
        Case 4: AR_Stripe RGB(255, 255, 153), "연한 노랑"
        Case 5: AR_Stripe RGB(252, 228, 236), "연한 핑크"
        Case 6: AR_Stripe RGB(226, 239, 218), "연한 녹색"
    End Select
End Sub

' 선택 범위(1셀이면 연결 영역, 열/행 전체면 사용 영역)의 채우기색·조건부서식 제거
Private Sub CC_ClearColors()
    Dim rng As Range, ws As Worksheet, nCF As Long
    If TypeName(Selection) <> "Range" Then
        MsgBox "색을 지울 범위를 먼저 선택하세요.", vbExclamation, APP_TITLE
        Exit Sub
    End If
    Set rng = Selection
    Set ws = rng.Worksheet
    If rng.Cells.CountLarge = 1 Then Set rng = rng.CurrentRegion
    If rng.rows.Count = ws.rows.Count Or rng.Columns.Count = ws.Columns.Count Then
        If Intersect(rng, ws.UsedRange) Is Nothing Then
            MsgBox "선택한 열/행에 데이터가 없습니다.", vbExclamation, APP_TITLE
            Exit Sub
        End If
        Set rng = Intersect(rng, ws.UsedRange)
    End If
    If UI_Choose(APP_TITLE & " - 색 지우기", "범위 " & rng.Address(False, False) & " 의 채우기색과 조건부서식을 모두 지웁니다." & vbCrLf & "(글꼴·테두리·값은 그대로)", _
                 Array("지우기", "취소")) <> 1 Then Exit Sub
    Application.ScreenUpdating = False
    nCF = rng.FormatConditions.Count
    rng.FormatConditions.Delete
    rng.Interior.ColorIndex = xlNone
    Application.ScreenUpdating = True
    MsgBox "채우기색을 지우고 조건부서식 " & nCF & "개를 제거했습니다." & vbCrLf & vbCrLf & "대상 범위: " & rng.Address(False, False), vbInformation, APP_TITLE
End Sub
