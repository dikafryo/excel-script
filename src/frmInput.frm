VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmInput 
   Caption         =   "입력"
   ClientHeight    =   4140
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8955.001
   OleObjectBlob   =   "frmInput.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "frmInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' frmInput : 입력 폼 (애플 스타일). 기존 btnOK·btnCancel 은 숨기고 Label 버튼을 동적으로 만든다.
'   호출: modRibbonUI.UI_Input → Me.Setup(title, prompt, defaultText, numericOnly) → Show
Public Cancelled As Boolean
Public numericOnly As Boolean
Private items As Collection

Private Sub UserForm_Initialize()
    Cancelled = True
    Set items = New Collection
End Sub

Public Sub Setup(ByVal title As String, ByVal prompt As String, ByVal defaultText As String, ByVal numOnly As Boolean)
    Dim y As Single, ctl As Object, txt As Object, lblTitle As Object, lblSub As Object, BW As Single
    numericOnly = numOnly
    UI_StyleForm Me
    Me.caption = title
    btnOK.Visible = False
    btnCancel.Visible = False
    lblPrompt.Visible = False

    Set lblTitle = Me.Controls.Add("Forms.Label.1", "uiTitle", True)
    lblTitle.top = 18
    UI_StyleTitle lblTitle, title
    y = lblTitle.top + lblTitle.Height + 6

    Set lblSub = Me.Controls.Add("Forms.Label.1", "uiPrompt", True)
    lblSub.top = y
    UI_StylePrompt lblSub, prompt
    y = lblSub.top + lblSub.Height + IIf(lblSub.Height > 0, 14, 8)

    UI_StyleTextBox txtValue, y
    txtValue.text = defaultText
    y = y + txtValue.Height + 22

    Set ctl = Me.Controls.Add("Forms.Label.1", "uiDiv", True)
    UI_StyleDivider ctl, y
    y = y + 12

    BW = 120
    Set ctl = Me.Controls.Add("Forms.Label.1", "uiOK", True)
    UI_StyleButton2 ctl, "", True, UI_FormWidth - UI_Margin - 8 - BW, y, BW
    AddBtn ctl, 1, 1
    Set txt = Me.Controls.Add("Forms.Label.1", "uiOKT", True)
    UI_StyleButtonText txt, ctl, "확인  (Enter)", True
    AddBtn txt, 1, 3

    Set ctl = Me.Controls.Add("Forms.Label.1", "uiCancel", True)
    UI_StyleButton2 ctl, "", False, UI_FormWidth - UI_Margin - 8 - BW * 2 - 10, y, BW
    AddBtn ctl, 0, 2
    Set txt = Me.Controls.Add("Forms.Label.1", "uiCancelT", True)
    UI_StyleButtonText txt, ctl, "취소  (Esc)", False
    AddBtn txt, 0, 3

    Me.Height = y + UI_ButtonH + 40
End Sub

Private Sub UserForm_Activate()
    On Error Resume Next
    txtValue.SetFocus
    txtValue.SelStart = 0
    txtValue.SelLength = Len(txtValue.text)
End Sub

Private Sub AddBtn(ByVal ctl As Object, ByVal idx As Long, ByVal kind As Long)
    Dim b As clsUIButton
    Set b = New clsUIButton
    Set b.lbl = ctl: Set b.Owner = Me: b.Index = idx: b.kind = kind
    items.Add b
End Sub

Public Sub OnPick(ByVal idx As Long)
    If idx = 1 Then DoOK Else DoCancel
End Sub

Public Sub OnHover(ByVal idx As Long)
    Dim b As clsUIButton
    For Each b In items
        b.SetHover (b.Index = idx)
    Next b
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    OnHover -1
End Sub

Private Sub DoOK()
    If numericOnly Then
        If Not IsNumeric(Trim$(txtValue.text)) Then
            MsgBox "숫자를 입력하세요.", vbExclamation, Me.caption
            txtValue.SetFocus
            Exit Sub
        End If
    End If
    Cancelled = False
    Me.Hide
End Sub

Private Sub DoCancel()
    Cancelled = True
    Me.Hide
End Sub

' 입력 상자에서 Enter=확인, Esc=취소
Private Sub txtValue_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0: DoOK
    ElseIf KeyCode = vbKeyEscape Then
        KeyCode = 0: DoCancel
    End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        Cancel = 1
        Cancelled = True
        Me.Hide
    End If
End Sub
