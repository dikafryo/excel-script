VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsUIButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' clsUIButton : 동적으로 만든 Label 을 버튼처럼 쓰기 위한 이벤트 래퍼
'   Owner 폼의 Public Sub OnPick(idx) 를 호출한다. Kind: 0=옵션 카드, 1=주요 버튼, 2=보조 버튼, 3=글자 라벨(카드 위 투명), 4=체크 항목(OnExtra)

Public WithEvents lbl As MSForms.label
Attribute lbl.VB_VarHelpID = -1
Public Owner As Object
Public Index As Long
Public kind As Long

Private Sub lbl_Click()
    If Owner Is Nothing Then Exit Sub
    If kind = 4 Then
        Owner.OnExtra Index
    Else
        Owner.OnPick Index
    End If
End Sub

Private Sub lbl_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If Owner Is Nothing Then Exit Sub
    If kind = 4 Then Owner.OnHover -1 Else Owner.OnHover Index
End Sub

Public Sub SetHover(ByVal isOn As Boolean)
    Select Case kind
        Case 0: UI_Hover lbl, isOn
        Case 1, 2: UI_HoverButton lbl, (kind = 1), isOn
        Case Else   ' 3 = 글자 라벨(투명), 4 = 체크 항목 : 색 변화 없음
    End Select
End Sub
