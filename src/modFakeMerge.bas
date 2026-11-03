Attribute VB_Name = "modFakeMerge"
Option Explicit
' 병합효과: 같은 값이 이어지는 셀을 실제로 병합하지 않고, 가운데 한 칸만 보이게(;;;) 하고 사이 테두리를 조건부서식으로 숨긴다.
' [개선] 날짜·숫자 표시형식을 보존(General 로 덮어쓰지 않음), 오류값 셀 안전, 팔레트 복원, 결과 안내.

Private Const HIDE_FMT As String = ";;;"

Sub Fake병합(control As IRibbonControl)
    Dim rng As Range, startCell As Range, midCell As Range
    Dim i As Long, j As Long, countSame As Long, midOffset As Long
    Dim bgColor As Long, groups As Long, ans As Long
    Dim cur As String, nxt As String

    If TypeName(Selection) <> "Range" Then
        MsgBox "병합효과를 줄 범위를 먼저 선택하세요.", vbExclamation, "병합효과"
        Exit Sub
    End If
    Set rng = Selection
    If rng.Cells.CountLarge < 2 Then
        MsgBox "같은 값이 이어지는 셀 2개 이상을 포함해 범위를 선택하세요.", vbExclamation, "병합효과"
        Exit Sub
    End If

    ans = UI_Choose("병합효과 - 배경색 설정", "같은 값이 이어지는 구간의 배경색을 어떻게 할까요?", Array("색상 직접 선택", "기본 흰색 적용"))
    If ans = 0 Then Exit Sub
    If ans = 1 Then
        bgColor = PickBgColor()
        If bgColor = -1 Then Exit Sub
    Else
        bgColor = RGB(255, 255, 255)
    End If

    Application.ScreenUpdating = False
    ' 1. 초기화: 이전 병합효과만 되돌리고(표시형식 보존) 테두리·정렬 정리
    FM_RestoreHidden rng
    With rng
        .FormatConditions.Delete
        .Borders.LineStyle = xlContinuous
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlCenter
    End With
    ' 2. 같은 값이 위/아래로 이어지면 사이 선을 숨기는 조건부서식 (빈 셀 제외)
    With rng.FormatConditions.Add(Type:=xlExpression, formula1:="=AND(RC<>"""",RC=R[-1]C)")
        .Borders(xlTop).LineStyle = xlNone
    End With
    With rng.FormatConditions.Add(Type:=xlExpression, formula1:="=AND(RC<>"""",RC=R[1]C)")
        .Borders(xlBottom).LineStyle = xlNone
    End With
    ' 3. 열별로 같은 값 구간을 찾아 가운데 셀만 남기고 숨김
    For j = 1 To rng.Columns.Count
        Set startCell = rng.Cells(1, j)
        countSame = 1
        For i = 1 To rng.rows.Count
            cur = FM_Val(rng.Cells(i, j))
            nxt = ""
            If i < rng.rows.Count Then nxt = FM_Val(rng.Cells(i + 1, j))
            If cur <> "" And cur = nxt Then
                countSame = countSame + 1
            Else
                If countSame > 1 Then
                    midOffset = (countSame - 1) \ 2
                    Set midCell = startCell.Offset(midOffset, 0)
                    With Range(startCell, rng.Cells(i, j))
                        .Interior.color = bgColor
                    End With
                    FM_HideExcept Range(startCell, rng.Cells(i, j)), midCell
                    groups = groups + 1
                End If
                If i < rng.rows.Count Then Set startCell = rng.Cells(i + 1, j)
                countSame = 1
            End If
        Next i
    Next j
    Application.ScreenUpdating = True
    MsgBox "병합효과를 적용했습니다." & vbCrLf & vbCrLf & _
           "대상 범위: " & rng.Address(False, False) & vbCrLf & _
           "같은 값 구간: " & groups & "곳 (가운데 셀만 표시)" & vbCrLf & vbCrLf & _
           "되돌리려면 같은 범위를 선택하고 [병합효과취소]를 누르세요.", vbInformation, "병합효과"
End Sub

' 비교용 문자열: 오류값은 주소를 붙여 서로 다르게, 빈 셀은 ""
Private Function FM_Val(ByVal c As Range) As String
    If IsError(c.value) Then
        FM_Val = "#ERR" & c.Address
    Else
        FM_Val = CStr(c.value)
    End If
End Function

' 구간의 셀을 숨기되(;;;) 가운데 셀은 원래 표시형식을 유지
Private Sub FM_HideExcept(ByVal grp As Range, ByVal keep As Range)
    Dim c As Range
    For Each c In grp.Cells
        If c.Address <> keep.Address Then
            If c.NumberFormat <> HIDE_FMT Then c.id = c.NumberFormat   ' 원래 형식 보관 (병합취소 때 복원)
            c.NumberFormat = HIDE_FMT
        End If
    Next c
End Sub

' 숨긴 셀(;;;)의 표시형식을 보관해 둔 원래 형식으로 복원. 보관값이 없으면 같은 열의 표시된 이웃 셀 형식, 그것도 없으면 General
Public Sub FM_RestoreHidden(ByVal rng As Range)
    Dim c As Range, src As Range, k As Long, fmt As String
    For Each c In rng.Cells
        If c.NumberFormat = HIDE_FMT Then
            fmt = ""
            On Error Resume Next
            fmt = c.id
            On Error GoTo 0
            If fmt = "" Or fmt = HIDE_FMT Then
                For k = 1 To 50
                    If c.row - k >= 1 Then
                        Set src = c.Offset(-k, 0)
                        If src.NumberFormat <> HIDE_FMT And FM_Val(src) = FM_Val(c) Then fmt = src.NumberFormat: Exit For
                    End If
                    If c.row + k <= c.parent.rows.Count Then
                        Set src = c.Offset(k, 0)
                        If src.NumberFormat <> HIDE_FMT And FM_Val(src) = FM_Val(c) Then fmt = src.NumberFormat: Exit For
                    End If
                Next k
            End If
            If fmt = "" Or fmt = HIDE_FMT Then fmt = "General"
            c.NumberFormat = fmt
            On Error Resume Next
            c.id = ""
            On Error GoTo 0
        End If
    Next c
End Sub

' 색상 대화상자: 팔레트 10번을 잠시 빌려 쓰고 원래 색으로 되돌린다
Private Function PickBgColor() As Long
    Dim orig As Long, ok As Boolean
    On Error GoTo ColorError
    orig = ActiveWorkbook.Colors(10)
    ok = Application.Dialogs(xlDialogEditColor).Show(10, 255, 255, 255)
    If ok Then
        PickBgColor = ActiveWorkbook.Colors(10)
    Else
        PickBgColor = -1
    End If
    ActiveWorkbook.Colors(10) = orig
    Exit Function
ColorError:
    PickBgColor = -1
End Function
