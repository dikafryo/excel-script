Attribute VB_Name = "modTestHooks"
Option Explicit
' 빌드 검증용 훅 (리본 미연결). 헤드리스 검증 스크립트가 Application.Run 으로 호출한다.
Public Sub BX_BuildCheck()
    ' 빌드 스크립트가 Application.Run 으로 호출해 전체 프로젝트 컴파일을 확인한다.
End Sub

Public Sub BX_SetSilent(ByVal v As Boolean)
    BX_Silent = v
End Sub

Public Sub BX_SetTestFolder(ByVal p As String)
    MH_TestFolder = p
    HP_TestFolder = p
End Sub

' 범위를 표 HTML 로 클립보드에 복사 (keepColors) → "행x열" 반환
Public Function BX_TestCopyTable(ByVal rng As Range, ByVal keepColors As Boolean, Optional ByVal forceFontPt As Long = 0) As String
    Dim nR As Long, nC As Long
    If HC_CopyRangeAsTable(rng, nR, nC, keepColors, forceFontPt) Then BX_TestCopyTable = nR & "x" & nC Else BX_TestCopyTable = "FAIL"
End Function
