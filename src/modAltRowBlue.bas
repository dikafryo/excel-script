Attribute VB_Name = "modAltRowBlue"
Option Explicit
' 교차행 색: 선택 범위의 짝수 행에 연한 색을 칠한다. [v2.6] 리본에서는 셀색표시 메뉴(modCellColorMenu)로 진입
' [개선] 1셀 선택 → 연결 영역, 열/행 전체 선택 → 사용 영역으로 제한(수십만 행 채우기 방지), 범위 미선택 안내, 결과 메시지

Sub 짝수행연한파란색(control As IRibbonControl)
    AR_Stripe RGB(221, 235, 247), "연한 파랑"
End Sub

Public Sub AR_Stripe(ByVal color As Long, ByVal colorName As String)
    Dim rng As Range, ws As Worksheet, i As Long, n As Long
    If TypeName(Selection) <> "Range" Then
        MsgBox "색을 칠할 범위를 먼저 선택하세요.", vbExclamation, "교차행 " & colorName
        Exit Sub
    End If
    Set rng = Selection
    Set ws = rng.Worksheet
    If rng.Cells.CountLarge = 1 Then Set rng = rng.CurrentRegion
    If rng.rows.Count = ws.rows.Count Or rng.Columns.Count = ws.Columns.Count Then
        If Intersect(rng, ws.UsedRange) Is Nothing Then
            MsgBox "선택한 열/행에 데이터가 없습니다.", vbExclamation, "교차행 " & colorName
            Exit Sub
        End If
        Set rng = Intersect(rng, ws.UsedRange)
    End If
    Application.ScreenUpdating = False
    For i = 2 To rng.rows.Count Step 2
        rng.rows(i).Interior.color = color
        n = n + 1
    Next i
    Application.ScreenUpdating = True
    MsgBox "짝수 행 " & n & "개에 " & colorName & " 색을 칠했습니다." & vbCrLf & vbCrLf & _
           "대상 범위: " & rng.Address(False, False), vbInformation, "교차행 " & colorName
End Sub

