Attribute VB_Name = "modFakeMergeUndo"
Option Explicit
' 병합효과취소: 선택 범위에서 병합효과가 남긴 숨김(;;;)·조건부서식·배경을 되돌린다.
' [개선] 표시형식을 General 로 덮어쓰지 않고 숨긴 셀만 원래 형식으로 복원(modFakeMerge.FM_RestoreHidden)

Sub 병합취소(control As IRibbonControl)
    Dim rng As Range, c As Range, hidden As Long
    If TypeName(Selection) <> "Range" Then
        MsgBox "병합효과를 취소할 범위를 먼저 선택하세요.", vbExclamation, "병합효과취소"
        Exit Sub
    End If
    Set rng = Selection
    If rng.Cells.CountLarge = 1 Then Set rng = rng.CurrentRegion

    For Each c In rng.Cells
        If c.NumberFormat = ";;;" Then hidden = hidden + 1
    Next c

    Application.ScreenUpdating = False
    FM_RestoreHidden rng              ' 숨긴 셀 → 원래 표시형식
    rng.FormatConditions.Delete       ' 테두리 숨김 규칙 제거
    rng.Borders.LineStyle = xlContinuous
    rng.Interior.ColorIndex = xlNone  ' 병합효과 배경색 제거
    rng.VerticalAlignment = xlCenter
    rng.HorizontalAlignment = xlCenter
    Application.ScreenUpdating = True

    MsgBox "병합효과를 취소했습니다." & vbCrLf & vbCrLf & _
           "대상 범위: " & rng.Address(False, False) & vbCrLf & _
           "다시 표시한 셀: " & hidden & "개", vbInformation, "병합효과취소"
End Sub
