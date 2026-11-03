VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmChoice 
   Caption         =   "선택"
   ClientHeight    =   6615
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8565.001
   OleObjectBlob   =   "frmChoice.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "frmChoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' frmChoice : 선택 폼 (애플 스타일). 기존 btn1~6·btnCancel 은 숨기고 Label 카드를 동적으로 만든다.
'   호출: modRibbonUI.UI_Choose → Me.Setup(title, prompt, options) → Show → Result(1~, 취소=0)
'   [v2.6] SetupEx : 카드 오른쪽 미리보기(previews, modUITheme.UI_DrawPreview) + 체크 항목(extras, 단일 선택 → Extra)
Public Result As Long
Public Extra As Long              ' 체크 항목 선택 번호(1~), 없으면 0
Private items As Collection      ' clsUIButton 보관 (이벤트 유지)
Private optCount As Long
Private cancelIdx As Long
Private mPrev As Variant         ' 미리보기 spec 배열
Private mExtras As Variant       ' 체크 항목 캡션 배열
Private mCards As Collection     ' 옵션 카드 라벨(미리보기 다시 그릴 때)
Private mPrevCtls As Collection  ' 그려 둔 미리보기 컨트롤

Private Sub UserForm_Initialize()
    Result = 0
    Extra = 0
    Set items = New Collection
    Set mCards = New Collection
End Sub

Public Sub Setup(ByVal title As String, ByVal prompt As String, ByVal options As Variant)
    SetupEx title, prompt, options, Empty, Empty, 0
End Sub

Public Sub SetupEx(ByVal title As String, ByVal prompt As String, ByVal options As Variant, ByVal previews As Variant, ByVal extras As Variant, ByVal extraDefault As Long)
    Dim i As Long, n As Long, y As Single, ctl As Object, txt As Object, num As Object, lblTitle As Object, lblSub As Object
    Dim spec As String, hasPrev As Boolean, k As Long, nx As Long, ew As Single
    UI_StyleForm Me
    Me.caption = title
    For i = 1 To 6
        Me.Controls("btn" & i).Visible = False
    Next i
    btnCancel.Visible = False
    lblPrompt.Visible = False
    mPrev = previews
    mExtras = extras
    hasPrev = IsArray(previews)

    Set lblTitle = Me.Controls.Add("Forms.Label.1", "uiTitle", True)
    lblTitle.top = 18
    UI_StyleTitle lblTitle, title
    y = lblTitle.top + lblTitle.Height + 6

    Set lblSub = Me.Controls.Add("Forms.Label.1", "uiPrompt", True)
    lblSub.top = y
    UI_StylePrompt lblSub, prompt
    y = lblSub.top + lblSub.Height + IIf(lblSub.Height > 0, 14, 8)

    n = UBound(options) - LBound(options) + 1
    If n > 6 Then n = 6
    optCount = n
    For i = 1 To n
        spec = ""
        If hasPrev Then
            If LBound(previews) + i - 1 <= UBound(previews) Then spec = CStr(previews(LBound(previews) + i - 1))
        End If
        Set ctl = Me.Controls.Add("Forms.Label.1", "uiOpt" & i, True)
        UI_StyleOption ctl, i, "", y
        AddBtn ctl, i, 0
        mCards.Add ctl
        Set txt = Me.Controls.Add("Forms.Label.1", "uiOptT" & i, True)
        UI_StyleOptionText txt, ctl, i, CStr(options(LBound(options) + i - 1)), IIf(Len(spec) > 0, UI_PrevW + 10, 0)
        AddBtn txt, i, 3
        Set num = Me.Controls.Add("Forms.Label.1", "uiOptN" & i, True)
        UI_StyleOptionNum num, ctl, i, txt
        AddBtn num, i, 3
        y = y + ctl.Height + UI_OptionGap
    Next i

    ' 체크 항목(단일 선택) 줄
    If IsArray(extras) Then
        nx = UBound(extras) - LBound(extras) + 1
        If nx > 0 Then
            y = y + 6
            ew = (UI_FormWidth - UI_Margin * 2 - 8) / nx
            For k = 1 To nx
                Set ctl = Me.Controls.Add("Forms.Label.1", "uiExtra" & k, True)
                UI_StyleExtra ctl, CStr(extras(LBound(extras) + k - 1)), UI_Margin + (k - 1) * ew, y, ew - 4
                AddBtn ctl, k, 4
            Next k
            y = y + 26
            Extra = extraDefault
            If Extra < 1 Or Extra > nx Then Extra = 1
            RefreshExtras
        End If
    End If

    y = y + 10
    Set ctl = Me.Controls.Add("Forms.Label.1", "uiDiv", True)
    UI_StyleDivider ctl, y
    y = y + 12

    cancelIdx = 0
    Set ctl = Me.Controls.Add("Forms.Label.1", "uiCancel", True)
    UI_StyleButton2 ctl, "", False, UI_FormWidth - UI_Margin - 8 - 120, y, 120
    AddBtn ctl, 0, 2
    Set txt = Me.Controls.Add("Forms.Label.1", "uiCancelT", True)
    UI_StyleButtonText txt, ctl, "취소  (Esc)", False
    AddBtn txt, 0, 3

    Me.Height = y + UI_ButtonH + 40
    DrawPreviews
End Sub

' 미리보기 그리기(다시 그리기 포함). 선 미리보기는 체크된 굵기(Extra)를 반영
Private Sub DrawPreviews()
    Dim i As Long, c As Object, col As Collection, spec As String
    If Not mPrevCtls Is Nothing Then
        For Each c In mPrevCtls
            On Error Resume Next
            Me.Controls.Remove c.name
            On Error GoTo 0
        Next c
    End If
    Set mPrevCtls = New Collection
    If Not IsArray(mPrev) Then Exit Sub
    For i = 1 To optCount
        spec = ""
        If LBound(mPrev) + i - 1 <= UBound(mPrev) Then spec = CStr(mPrev(LBound(mPrev) + i - 1))
        If Len(spec) > 0 Then
            Set col = UI_DrawPreview(Me, mCards(i), spec, "uiPv" & i & "_", IIf(Extra < 1, 1, Extra))
            For Each c In col
                AddBtn c, i, 3
                mPrevCtls.Add c
            Next c
        End If
    Next i
End Sub

Private Sub RefreshExtras()
    Dim k As Long, nx As Long
    If Not IsArray(mExtras) Then Exit Sub
    nx = UBound(mExtras) - LBound(mExtras) + 1
    For k = 1 To nx
        Me.Controls("uiExtra" & k).caption = IIf(k = Extra, "■ ", "□ ") & CStr(mExtras(LBound(mExtras) + k - 1))
        Me.Controls("uiExtra" & k).Font.Bold = (k = Extra)
    Next k
End Sub

' 라벨을 클릭 가능한 버튼으로 등록
Private Sub AddBtn(ByVal ctl As Object, ByVal idx As Long, ByVal kind As Long)
    Dim b As clsUIButton
    Set b = New clsUIButton
    Set b.lbl = ctl: Set b.Owner = Me: b.Index = idx: b.kind = kind
    items.Add b
End Sub

' 콜백: 클릭
Public Sub OnPick(ByVal idx As Long)
    Result = idx
    Me.Hide
End Sub

' 콜백: 체크 항목 클릭 → 선택 바꾸고 미리보기 다시 그림
Public Sub OnExtra(ByVal idx As Long)
    Extra = idx
    RefreshExtras
    DrawPreviews
End Sub

' 콜백: 마우스 올림 → 해당 항목만 강조
Public Sub OnHover(ByVal idx As Long)
    Dim b As clsUIButton
    For Each b In items
        b.SetHover (b.Index = idx)
    Next b
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    OnHover -1
End Sub

' 키보드: 숫자 1~6 선택, Esc 취소
Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim k As Long
    k = KeyCode
    If k = vbKeyEscape Then
        Result = 0: Me.Hide
    ElseIf k >= vbKey1 And k <= vbKey6 Then
        If k - vbKey0 <= optCount Then Result = k - vbKey0: Me.Hide
    ElseIf k >= vbKeyNumpad1 And k <= vbKeyNumpad6 Then
        If k - vbKeyNumpad0 <= optCount Then Result = k - vbKeyNumpad0: Me.Hide
    End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        Cancel = 1
        Result = 0
        Me.Hide
    End If
End Sub
