Attribute VB_Name = "modBorderStyle"
Option Explicit
' 테두리 스타일: 선택 범위의 바깥·안쪽 테두리에 선 종류와 굵기를 적용한다.
' [v2.6] 폼 한 장으로 통합: 선 종류마다 선 모양 미리보기 + 굵기(얇게·중간·굵게)는 체크 항목으로 선택
'        1셀 선택 시 연결 영역, 1열/1행 범위의 안쪽 선 오류 방지, 엑셀이 지원하지 않는 조합은 자동 보정

Sub 테두리선스타일(control As IRibbonControl)
    Dim rng As Range, kind As Long, wt As Long, ls As Long, W As Long, b As Variant, fixMsg As String, wtName As String, kindName As String
    If TypeName(Selection) <> "Range" Then
        MsgBox "테두리를 적용할 범위를 먼저 선택하세요.", vbExclamation, "테두리 스타일"
        Exit Sub
    End If
    Set rng = Selection
    If rng.Cells.CountLarge = 1 Then Set rng = rng.CurrentRegion

    kind = UI_ChooseEx("테두리 스타일", _
                       "범위 " & rng.Address(False, False) & " 의 모든 테두리(바깥+안쪽)에 적용합니다." & vbCrLf & _
                       "아래에서 굵기를 체크한 뒤 선 종류를 클릭하면 바로 적용됩니다.", _
                       Array("실선", "점선", "짧은 점선", "점-실선", "이중선", "테두리 없음 (지우기)"), _
                       Array("line:solid", "line:dot", "line:dash", "line:dashdot", "line:double", "line:none"), _
                       Array("얇게", "중간", "굵게"), 1, wt)
    If kind = 0 Then Exit Sub
    Select Case kind
        Case 1: ls = xlContinuous: kindName = "실선"
        Case 2: ls = xlDot: kindName = "점선"
        Case 3: ls = xlDash: kindName = "짧은 점선"
        Case 4: ls = xlDashDot: kindName = "점-실선"
        Case 5: ls = xlDouble: kindName = "이중선"
        Case 6: ls = xlLineStyleNone
    End Select

    If ls = xlLineStyleNone Then
        Application.ScreenUpdating = False
        rng.Borders.LineStyle = xlLineStyleNone
        Application.ScreenUpdating = True
        MsgBox "테두리를 지웠습니다." & vbCrLf & "대상 범위: " & rng.Address(False, False), vbInformation, "테두리 스타일"
        Exit Sub
    End If

    Select Case wt
        Case 2: W = xlMedium: wtName = "중간"
        Case 3: W = xlThick: wtName = "굵게"
        Case Else: W = xlThin: wtName = "얇게"
    End Select
    ' 엑셀이 지원하는 선종류×굵기 조합으로 보정 (점선=얇게만, 짧은점선/점-실선=얇게·중간, 이중선=굵게 고정)
    Select Case ls
        Case xlDouble
            If W <> xlThick Then fixMsg = "이중선은 굵기가 고정입니다."
            W = xlThick
        Case xlDot
            If W <> xlThin Then fixMsg = "점선은 얇게만 지원되어 얇게로 적용했습니다."
            W = xlThin
        Case xlDash, xlDashDot
            If W = xlThick Then fixMsg = "이 선 종류는 굵게가 없어 중간으로 적용했습니다.": W = xlMedium
    End Select

    Application.ScreenUpdating = False
    For Each b In Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight)
        With rng.Borders(b)
            .LineStyle = ls
            .weight = W
            .ColorIndex = xlAutomatic
        End With
    Next b
    If rng.Columns.Count > 1 Then
        With rng.Borders(xlInsideVertical)
            .LineStyle = ls: .weight = W: .ColorIndex = xlAutomatic
        End With
    End If
    If rng.rows.Count > 1 Then
        With rng.Borders(xlInsideHorizontal)
            .LineStyle = ls: .weight = W: .ColorIndex = xlAutomatic
        End With
    End If
    Application.ScreenUpdating = True
    MsgBox "테두리를 적용했습니다. (" & kindName & " · " & wtName & ")" & vbCrLf & vbCrLf & _
           "대상 범위: " & rng.Address(False, False) & IIf(Len(fixMsg) > 0, vbCrLf & "※ " & fixMsg, ""), vbInformation, "테두리 스타일"
End Sub

