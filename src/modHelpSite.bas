Attribute VB_Name = "modHelpSite"
Option Explicit
' 리본 [사용방법] 도움말 아이콘: 사용설명서 사이트 열기
Sub OpenHelpWebsite(control As IRibbonControl)
    Dim url As String
    url = "https://hwpxlam.com/xlam/"
    On Error GoTo Fail
    ThisWorkbook.FollowHyperlink Address:=url
    Exit Sub
Fail:
    MsgBox "도움말 사이트를 열 수 없습니다." & vbCrLf & url & vbCrLf & vbCrLf & _
           "브라우저에서 위 주소를 직접 열어 주세요. (" & Err.Description & ")", vbExclamation, "B-XLAM V8"
End Sub


