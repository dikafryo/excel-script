VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBoardEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' 예산보드 빈 행 자동 숨김 (2.7m): 슬라이서로 팀을 고르면 재계산이 일어나므로 그때 값이 빈 행을 숨긴다
Public WithEvents App As Application
Attribute App.VB_VarHelpID = -1
Private busyHide As Boolean

Private Sub App_SheetCalculate(ByVal Sh As Object)
    Dim wb As Workbook, nm As name, rg As Range, c As Range, r As Long, d As Object, k As Variant, allEmpty As Boolean
    If busyHide Then Exit Sub
    On Error GoTo Done
    If TypeName(Sh) <> "Worksheet" Then Exit Sub
    Set wb = Sh.parent
    Set d = CreateObject("Scripting.Dictionary")
    For Each nm In wb.names
        If Left$(nm.name, 8) = "BX_HIDE_" Then
            Set rg = Nothing
            On Error Resume Next
            Set rg = nm.RefersToRange
            On Error GoTo Done
            If Not rg Is Nothing Then
                If rg.parent.name = Sh.name Then
                    For Each c In rg.Cells
                        allEmpty = (Len(CStr(c.Value2)) = 0)
                        If d.Exists(c.row) Then d(c.row) = d(c.row) And allEmpty Else d.Add c.row, allEmpty
                    Next c
                End If
            End If
        End If
    Next nm
    If d.Count = 0 Then Exit Sub
    busyHide = True
    Application.ScreenUpdating = False
    For Each k In d.keys
        r = CLng(k)
        If Sh.rows(r).hidden <> d(k) Then Sh.rows(r).hidden = d(k)
    Next k
Done:
    Application.ScreenUpdating = True
    busyHide = False
End Sub
