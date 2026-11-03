Attribute VB_Name = "modCellColorMark"
Option Explicit

Sub 셀색표시(control As IRibbonControl)

    Dim cfChoice As VbMsgBoxResult
    Dim rngApply As Range
    Dim rngBaseCell As Range
    Dim baseCol As Long
    Dim baseAddr As String

    Dim mode As String
    Dim condText As String
    Dim formula1 As String

    Dim pickedColor As Long
    Dim resp As VbMsgBoxResult
    Dim topLeft As Range

    On Error GoTo CleanFail

    ' [요청사항] 매크로 실행 시 조건부서식 처리 방식 선택
    cfChoice = UI_ChoiceToMsg(UI_Choose("조건부 서식 처리", "조건부 서식 처리 방식을 선택하세요.", Array("해당 범위의 조건부서식 모두 삭제 후 새로 시작", "조건부서식 추가 (새 규칙을 1순위로 적용)")))
    If cfChoice = vbCancel Then Exit Sub

    ' 1) 범위 선택 (Ctrl+A = CurrentRegion)
    resp = UI_ChoiceToMsg(UI_Choose("1. 범위 선택", "적용 범위를 어떻게 정할까요?", Array("현재 셀의 연결된 범위(CurrentRegion) 자동 선택", "직접 범위 선택")))
    If resp = vbCancel Then Exit Sub

    If resp = vbYes Then
        Set rngApply = ActiveCell.CurrentRegion
    Else
        Set rngApply = Application.InputBox( _
            prompt:="색을 적용할 범위를 드래그로 선택하세요.", _
            title:="1. 범위 선택", Type:=8)
    End If
    If rngApply Is Nothing Then Exit Sub

    ' 선택범위의 "가장 작은(좌상단) 셀"
    Set topLeft = GetTopLeftCell(rngApply)

    ' 2) 기준열 선택 (셀 1개)
    Set rngBaseCell = Application.InputBox( _
        prompt:="기준열에서 셀 1개를 선택하세요." & vbCrLf & _
                "※ 이 열의 같은 '행' 값으로 조건을 판단합니다.", _
        title:="2. 기준열 선택", Type:=8)
    If rngBaseCell Is Nothing Then Exit Sub

    If rngBaseCell.Worksheet.name <> rngApply.Worksheet.name Then
        MsgBox "기준열은 적용 범위와 같은 시트에서 선택해야 합니다.", vbExclamation
        Exit Sub
    End If

    baseCol = rngBaseCell.Column

    ' 조건부서식 수식에서: 열은 고정, 행은 적용범위 첫 행 기준 상대
    baseAddr = rngApply.Worksheet.Cells(rngApply.row, baseCol).Address(RowAbsolute:=False, ColumnAbsolute:=True)

    ' 3) 방법 선택
    mode = CStr(UI_Choose("3. 방법 선택", "셀색을 표시할 조건을 고르세요.", Array("ISODD(기준셀) 가 TRUE일 때", "ISODD(기준셀) 가 FALSE일 때", "사용자 조건 (<10, >10, >10%, 또는 =수식)", "텍스트 조건 (같음 / 포함)")))
    mode = Trim$(mode)
    If mode = "0" Then Exit Sub

    Select Case mode
        Case "1"
            formula1 = "=ISODD(" & baseAddr & ")"

        Case "2"
            formula1 = "=NOT(ISODD(" & baseAddr & "))"

        Case "3"
            condText = UI_Input("3-3. 조건 입력", _
                "조건을 입력하세요." & vbCrLf & _
                "예) <10   >10   >10%   =AND(" & baseAddr & ">0," & baseAddr & "<100)", "<10")
            If condText = vbNullString Then Exit Sub

            If Left$(condText, 1) = "=" Then
                formula1 = condText
            Else
                If Left$(condText, 1) = "<" Or Left$(condText, 1) = ">" Or Left$(condText, 1) = "=" Then
                    formula1 = "=" & baseAddr & condText
                Else
                    MsgBox "조건 형식을 인식할 수 없습니다. 예: <10, >10%, 또는 =수식 형태로 입력하세요.", vbExclamation
                    Exit Sub
                End If
            End If

        Case "4"
            formula1 = BuildTextConditionFormula(rngApply, baseAddr)
            If formula1 = vbNullString Then Exit Sub

        Case Else
            MsgBox "방법은 1, 2, 3, 4 중 하나로 입력하세요.", vbExclamation
            Exit Sub
    End Select

    ' 4) 색상표에서 선택 (XFD로 화면 점프 방지 + 끝나면 topLeft로 이동)
    Select Case UI_Choose("4. 색 선택", "조건에 맞는 행에 칠할 색을 고르세요.", Array("연한 노랑", "연한 초록", "연한 파랑", "연한 빨강", "셀 서식 대화상자에서 직접 선택"))
        Case 1: pickedColor = RGB(255, 242, 204)
        Case 2: pickedColor = RGB(226, 239, 218)
        Case 3: pickedColor = RGB(221, 235, 247)
        Case 4: pickedColor = RGB(252, 228, 214)
        Case 5
            pickedColor = PickColorFromFormatCellsDialog_NoJump(topLeft)
            If pickedColor = -1 Then Exit Sub
        Case Else: Exit Sub
    End Select

    Application.ScreenUpdating = False

    ' [요청사항] Yes면 해당범위 조건부서식 전부 삭제
    If cfChoice = vbYes Then
        rngApply.FormatConditions.Delete
    End If

    ' 5) 적용(추가) + No면 1순위로 올리기
    With rngApply.FormatConditions.Add(Type:=xlExpression, formula1:=formula1)
        .Interior.color = pickedColor
        .StopIfTrue = False

        If cfChoice = vbNo Then
            ' 기존 규칙 유지 + 새 규칙을 최우선(1순위)로
            .SetFirstPriority
        End If
    End With

    ' 화면 복귀
    topLeft.Select
    ActiveWindow.ScrollRow = topLeft.row
    ActiveWindow.ScrollColumn = topLeft.Column

CleanExit:
    Application.ScreenUpdating = True
    Exit Sub

CleanFail:
    Application.ScreenUpdating = True
    MsgBox "취소되었거나 입력/선택이 올바르지 않습니다.", vbExclamation

End Sub


' === 4번(텍스트 조건) 수식 만들기 ===
Private Function BuildTextConditionFormula(ByVal rngApply As Range, ByVal baseAddr As String) As String
    Dim subMode As String
    Dim critCell As Range
    Dim critAddrAbs As String
    Dim txt As String

    subMode = CStr(UI_Choose("4. 텍스트 조건", "텍스트 조건을 고르세요.", Array("기준열 값이 '선택한 셀 값'과 같음", "기준열 값에 '입력 텍스트'가 포함됨")))
    If subMode = "0" Then
        BuildTextConditionFormula = vbNullString
        Exit Function
    End If

    Select Case subMode
        Case "1"
            Set critCell = Application.InputBox( _
                prompt:="비교 기준이 될 셀 1개를 선택하세요(그 셀의 값과 같으면 적용).", _
                title:="4-1. 같은 값(기준 셀 선택)", Type:=8)
            If critCell Is Nothing Then
                BuildTextConditionFormula = vbNullString
                Exit Function
            End If

            If critCell.Worksheet.name <> rngApply.Worksheet.name Then
                MsgBox "비교 기준 셀은 적용 범위와 같은 시트에서 선택해야 합니다.", vbExclamation
                BuildTextConditionFormula = vbNullString
                Exit Function
            End If

            critAddrAbs = critCell.Address(RowAbsolute:=True, ColumnAbsolute:=True)
            BuildTextConditionFormula = "=(" & baseAddr & "=" & critAddrAbs & ")"

        Case "2"
            txt = UI_Input("4-2. 텍스트 포함", _
                "포함 여부를 검사할 텍스트를 입력하세요." & vbCrLf & _
                "예) 서울", "")
            If Len(txt) = 0 Then
                BuildTextConditionFormula = vbNullString
                Exit Function
            End If

            ' 대/소문자 구분 없음(SEARCH). 오류는 FALSE 처리.
            BuildTextConditionFormula = "=IFERROR(ISNUMBER(SEARCH(""" & EscapeCFString(txt) & """," & baseAddr & ")),FALSE)"

        Case Else
            MsgBox "4번 텍스트 조건은 1 또는 2로 입력하세요.", vbExclamation
            BuildTextConditionFormula = vbNullString
    End Select
End Function


Private Function EscapeCFString(ByVal s As String) As String
    EscapeCFString = Replace(s, """", """""")
End Function


' 선택한 범위(여러 Area 포함)에서 "가장 작은(좌상단) 셀" 반환
Private Function GetTopLeftCell(ByVal rng As Range) As Range
    Dim a As Range
    Dim candidate As Range

    Set candidate = rng.Areas(1).Cells(1, 1)

    For Each a In rng.Areas
        If a.Cells(1, 1).row < candidate.row Then
            Set candidate = a.Cells(1, 1)
        ElseIf a.Cells(1, 1).row = candidate.row Then
            If a.Cells(1, 1).Column < candidate.Column Then
                Set candidate = a.Cells(1, 1)
            End If
        End If
    Next a

    Set GetTopLeftCell = candidate
End Function


' "셀 서식 > 패턴" 대화상자로 색을 고르게 하고,
' (1) XFD 선택으로 화면 점프 방지
' (2) 끝나면 returnTo 셀로 이동
Private Function PickColorFromFormatCellsDialog_NoJump(ByVal returnTo As Range) As Long

    Dim ws As Worksheet
    Dim tmp As Range
    Dim oldValue As Variant
    Dim oldColor As Variant
    Dim oldPattern As Variant
    Dim savedScrollRow As Long, savedScrollCol As Long
    Dim ok As Boolean

    On Error GoTo Fail

    Set ws = returnTo.Worksheet
    Set tmp = ws.Cells(ws.rows.Count, ws.Columns.Count) ' XFD1048576

    savedScrollRow = ActiveWindow.ScrollRow
    savedScrollCol = ActiveWindow.ScrollColumn

    oldValue = tmp.value
    oldColor = tmp.Interior.color
    oldPattern = tmp.Interior.Pattern
    Dim oldColorIndex As Variant: oldColorIndex = tmp.Interior.ColorIndex

    tmp.value = ""
    tmp.Interior.Pattern = xlSolid
    tmp.Interior.ColorIndex = xlColorIndexNone

    Application.ScreenUpdating = False
    tmp.Select
    ActiveWindow.ScrollRow = savedScrollRow
    ActiveWindow.ScrollColumn = savedScrollCol
    Application.ScreenUpdating = True

    ok = Application.Dialogs(xlDialogPatterns).Show
    If Not ok Then
        PickColorFromFormatCellsDialog_NoJump = -1
        GoTo RestoreAndMove
    End If

    PickColorFromFormatCellsDialog_NoJump = tmp.Interior.color

RestoreAndMove:
    Application.ScreenUpdating = False

    tmp.value = oldValue
    If oldColorIndex = xlColorIndexNone Then
        tmp.Interior.ColorIndex = xlColorIndexNone   ' 원래 채우기 없음이면 그대로 없음으로
    Else
        tmp.Interior.Pattern = oldPattern
        tmp.Interior.color = oldColor
    End If

    returnTo.Select
    ActiveWindow.ScrollRow = returnTo.row
    ActiveWindow.ScrollColumn = returnTo.Column

    Application.ScreenUpdating = True
    Exit Function

Fail:
    PickColorFromFormatCellsDialog_NoJump = -1
End Function


