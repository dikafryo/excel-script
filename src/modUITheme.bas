Attribute VB_Name = "modUITheme"
Option Explicit
' [v2.6] 옵션 카드 미리보기 크기
Public Const UI_PrevW As Single = 72
Public Const UI_PrevH As Single = 30
Public Const UI_NumW As Single = 28
' =====================================================================================
'  modUITheme : 폼 공통 디자인(애플 스타일) - 색상·글꼴·컨트롤 스타일을 한 곳에서 관리
'   - 폼 디자이너(.frx)를 건드리지 않고 실행 시점에 스타일을 입힌다.
'   - 옵션 버튼은 Label 컨트롤을 동적으로 만들어 왼쪽 정렬·호버 효과를 준다.
' =====================================================================================

' ---- 색상 (iOS/macOS 시스템 컬러 계열)
Public Const UI_ClrWindow As Long = &HFFFFFF        ' 흰 배경
Public Const UI_ClrPanel As Long = &HF7F2F2         ' systemGray6 (242,242,247) → BGR
Public Const UI_ClrHover As Long = &HF0E9E9         ' 호버 (233,233,240)
Public Const UI_ClrLine As Long = &HDCD6D6          ' 구분선 (214,214,220)
Public Const UI_ClrText As Long = &H1C1C1E          ' 본문 글자 (28,28,30)
Public Const UI_ClrSub As Long = &H8E8E8E           ' 보조 글자 (142,142,142)
Public Const UI_ClrAccent As Long = &HFF7A00        ' 애플 블루 (0,122,255) → BGR
Public Const UI_ClrAccentDn As Long = &HE06400      ' 눌림 (0,100,224)
Public Const UI_ClrBadge As Long = &HFCE9E6         ' 번호 배지 배경 (230,233,252)

' ---- 글꼴
Public Const UI_FontName As String = "맑은 고딕"
Public Const UI_FontTitle As Single = 14
Public Const UI_FontBody As Single = 11.5
Public Const UI_FontOption As Single = 12

' ---- 치수
Public Const UI_FormWidth As Single = 440
Public Const UI_Margin As Single = 22
Public Const UI_OptionH As Single = 40
Public Const UI_OptionGap As Single = 6
Public Const UI_ButtonH As Single = 34

' 폼 바탕 스타일
Public Sub UI_StyleForm(ByVal f As Object)
    On Error Resume Next
    f.BackColor = UI_ClrWindow
    f.Width = UI_FormWidth
    f.Font.name = UI_FontName
    f.Font.Size = UI_FontBody
    f.StartUpPosition = 1   ' CenterOwner
End Sub

' 제목 라벨(큰 글씨, 굵게)
Public Sub UI_StyleTitle(ByVal lbl As Object, ByVal text As String)
    With lbl
        .caption = text
        .Font.name = UI_FontName
        .Font.Size = UI_FontTitle
        .Font.Bold = True
        .ForeColor = UI_ClrText
        .BackColor = UI_ClrWindow
        .BackStyle = 1
        .TextAlign = 1        ' fmTextAlignLeft
        .WordWrap = True
        .Left = UI_Margin
        .Width = UI_FormWidth - UI_Margin * 2 - 8
        .AutoSize = True
        .AutoSize = False
        .Width = UI_FormWidth - UI_Margin * 2 - 8
        If .Height < 24 Then .Height = 24
    End With
End Sub

' 안내문 라벨(회색, 보통 글씨) - 길이에 맞춰 높이 자동
Public Sub UI_StylePrompt(ByVal lbl As Object, ByVal text As String)
    With lbl
        .caption = text
        .Font.name = UI_FontName
        .Font.Size = UI_FontBody
        .Font.Bold = False
        .ForeColor = UI_ClrSub
        .BackColor = UI_ClrWindow
        .BackStyle = 1
        .TextAlign = 1
        .WordWrap = True
        .Left = UI_Margin
        .Width = UI_FormWidth - UI_Margin * 2 - 8
        .AutoSize = True
        .AutoSize = False
        .Width = UI_FormWidth - UI_Margin * 2 - 8
        If Len(text) = 0 Then
            .Height = 0
        Else
            If .Height < 20 Then .Height = 20
            If .Height > 320 Then .Height = 320
        End If
    End With
End Sub

' 옵션(선택 항목) 라벨 - 왼쪽 정렬, 번호 + 문구, 카드형 테두리
Public Sub UI_StyleOption(ByVal lbl As Object, ByVal idx As Long, ByVal text As String, ByVal top As Single)
    With lbl
        .caption = ""
        .Font.name = UI_FontName
        .Font.Size = UI_FontOption
        .Font.Bold = False
        .ForeColor = UI_ClrText
        .BackColor = UI_ClrPanel
        .BackStyle = 1
        .BorderStyle = 1
        .BorderColor = UI_ClrLine
        .TextAlign = 1          ' 왼쪽 정렬
        .WordWrap = False
        .Left = UI_Margin
        .top = top
        .Width = UI_FormWidth - UI_Margin * 2 - 8
        .Height = UI_OptionH
    End With
End Sub

' 옵션 카드 위에 얹는 글자 라벨(투명 배경, 세로 가운데). 카드 높이를 글자에 맞춰 늘려 준다.
Public Sub UI_StyleOptionText(ByVal lbl As Object, ByVal card As Object, ByVal idx As Long, ByVal text As String, Optional ByVal reserveRight As Single = 0)
    With lbl
        .caption = text
        .Font.name = UI_FontName
        .Font.Size = UI_FontOption
        .Font.Bold = False
        .ForeColor = UI_ClrText
        .BackStyle = 0          ' 투명
        .BorderStyle = 0
        .TextAlign = 1          ' 왼쪽 정렬
        .WordWrap = True
        .Left = card.Left + 14 + UI_NumW
        .Width = card.Width - 28 - UI_NumW - reserveRight
        .AutoSize = True
        .AutoSize = False
        .Width = card.Width - 28 - UI_NumW - reserveRight
    End With
    If lbl.Height + 14 > card.Height Then card.Height = lbl.Height + 14
    lbl.top = card.top + (card.Height - lbl.Height) / 2
End Sub

' 버튼 위에 얹는 글자 라벨(투명, 가운데 정렬, 세로 가운데)
Public Sub UI_StyleButtonText(ByVal lbl As Object, ByVal card As Object, ByVal text As String, ByVal primary As Boolean)
    With lbl
        .caption = text
        .Font.name = UI_FontName
        .Font.Size = UI_FontBody
        .Font.Bold = primary
        .ForeColor = IIf(primary, UI_ClrWindow, UI_ClrAccent)
        .BackStyle = 0
        .BorderStyle = 0
        .TextAlign = 2
        .WordWrap = False
        .Left = card.Left
        .Width = card.Width
        .AutoSize = True
        .AutoSize = False
        .Width = card.Width
    End With
    lbl.top = card.top + (card.Height - lbl.Height) / 2
End Sub

' 주요 버튼(파란색, 흰 글씨) / 보조 버튼(회색 바탕, 파란 글씨)
Public Sub UI_StyleButton2(ByVal lbl As Object, ByVal text As String, ByVal primary As Boolean, _
                           ByVal px As Single, ByVal py As Single, ByVal pw As Single)
    With lbl
        .caption = ""
        .Font.name = UI_FontName
        .Font.Size = UI_FontBody
        .Font.Bold = primary
        .TextAlign = 2          ' 가운데
        .WordWrap = False
        .BackStyle = 1
        .BorderStyle = 1
        .Left = px
        .top = py
        .Width = pw
        .Height = UI_ButtonH
        If primary Then
            .BackColor = UI_ClrAccent
            .ForeColor = UI_ClrWindow
            .BorderColor = UI_ClrAccent
        Else
            .BackColor = UI_ClrPanel
            .ForeColor = UI_ClrAccent
            .BorderColor = UI_ClrLine
        End If
    End With
End Sub

' 입력 상자 스타일
Public Sub UI_StyleTextBox(ByVal tB As Object, ByVal top As Single)
    With tB
        .Font.name = UI_FontName
        .Font.Size = UI_FontOption
        .ForeColor = UI_ClrText
        .BackColor = UI_ClrWindow
        .BorderStyle = 1
        .BorderColor = UI_ClrLine
        .SpecialEffect = 0
        .Left = UI_Margin
        .top = top
        .Width = UI_FormWidth - UI_Margin * 2 - 8
        .Height = 32
    End With
End Sub

' 얇은 구분선
Public Sub UI_StyleDivider(ByVal lbl As Object, ByVal top As Single)
    With lbl
        .caption = ""
        .BackStyle = 1
        .BackColor = UI_ClrLine
        .BorderStyle = 0
        .Left = UI_Margin
        .top = top
        .Width = UI_FormWidth - UI_Margin * 2 - 8
        .Height = 1
    End With
End Sub

' 호버/기본 색 전환 (옵션 카드)
Public Sub UI_Hover(ByVal lbl As Object, ByVal isOn As Boolean)
    On Error Resume Next
    If isOn Then
        lbl.BackColor = UI_ClrHover
        lbl.BorderColor = UI_ClrAccent
    Else
        lbl.BackColor = UI_ClrPanel
        lbl.BorderColor = UI_ClrLine
    End If
End Sub

Public Sub UI_HoverButton(ByVal lbl As Object, ByVal primary As Boolean, ByVal isOn As Boolean)
    On Error Resume Next
    If primary Then
        lbl.BackColor = IIf(isOn, UI_ClrAccentDn, UI_ClrAccent)
    Else
        lbl.BackColor = IIf(isOn, UI_ClrHover, UI_ClrPanel)
        lbl.BorderColor = IIf(isOn, UI_ClrAccent, UI_ClrLine)
    End If
End Sub

' =====================================================================================
' [v2.6] 옵션 카드 미리보기 + 체크 항목(굵기 선택 등)
'   미리보기 spec : "stripe:RRGGBB" 교차행 표 / "line:solid|dot|dash|dashdot|double|none" 선 모양 /
'                  "merge:before|after" 병합효과 전·후 / "swatch:RRGGBB" 단색
' =====================================================================================

' 16진 "RRGGBB" → Long(RGB)
Public Function UI_HexColor(ByVal hex6 As String) As Long
    hex6 = Replace(Trim$(hex6), "#", "")
    If Len(hex6) <> 6 Then UI_HexColor = vbWhite: Exit Function
    UI_HexColor = RGB(CLng("&H" & Mid$(hex6, 1, 2)), CLng("&H" & Mid$(hex6, 3, 2)), CLng("&H" & Mid$(hex6, 5, 2)))
End Function

' 작은 상자(라벨) 하나 생성
Private Function UI_Box(ByVal f As Object, ByVal nm As String, ByVal L As Single, ByVal T As Single, ByVal W As Single, ByVal h As Single, _
                        ByVal back As Long, ByVal borderClr As Long, Optional ByVal cap As String = "", Optional ByVal fontSize As Single = 7) As Object
    Dim c As Object
    Set c = f.Controls.Add("Forms.Label.1", nm, True)
    With c
        .Left = L: .top = T: .Width = W: .Height = h
        .BackStyle = 1
        .BackColor = back
        If borderClr < 0 Then
            .BorderStyle = 0
        Else
            .BorderStyle = 1
            .BorderColor = borderClr
        End If
        .caption = cap
        .Font.name = UI_FontName
        .Font.Size = fontSize
        .TextAlign = 2
        .ForeColor = UI_ClrText
        .WordWrap = False
    End With
    Set UI_Box = c
End Function

' 카드 오른쪽에 미리보기를 그린다. 반환: 만든 컨트롤 모음(클릭 등록용). weight: 선 굵기 1~3
Public Function UI_DrawPreview(ByVal f As Object, ByVal card As Object, ByVal spec As String, ByVal prefix As String, ByVal weight As Long) As Collection
    Dim col As New Collection, kind As String, arg As String, p As Long
    Dim x0 As Single, y0 As Single, r As Long, c As Long, clr As Long, k As Long, x As Single, th As Single, yc As Single
    p = InStr(spec, ":")
    If p = 0 Then kind = LCase$(spec): arg = "" Else kind = LCase$(Left$(spec, p - 1)): arg = Mid$(spec, p + 1)
    x0 = card.Left + card.Width - 14 - UI_PrevW
    y0 = card.top + (card.Height - UI_PrevH) / 2
    If weight < 1 Then weight = 1
    If weight > 3 Then weight = 3
    Select Case kind
        Case "stripe"
            clr = UI_HexColor(arg)
            For r = 1 To 4
                For c = 1 To 3
                    col.Add UI_Box(f, prefix & r & "_" & c, x0 + (c - 1) * 24, y0 + (r - 1) * 7.5, 24, 7.5, IIf(r Mod 2 = 0, clr, vbWhite), UI_ClrLine)
                Next c
            Next r
        Case "swatch"
            col.Add UI_Box(f, prefix & "1", x0, y0 + 4, UI_PrevW, UI_PrevH - 8, UI_HexColor(arg), UI_ClrLine)
        Case "line"
            th = weight
            yc = y0 + UI_PrevH / 2 - th / 2
            Select Case LCase$(arg)
                Case "solid"
                    col.Add UI_Box(f, prefix & "1", x0, yc, UI_PrevW, th, 0, -1)
                Case "dot"
                    For k = 0 To 11
                        col.Add UI_Box(f, prefix & k, x0 + k * 6, yc, 3, th, 0, -1)
                    Next k
                Case "dash"
                    For k = 0 To 5
                        col.Add UI_Box(f, prefix & k, x0 + k * 12, yc, 8, th, 0, -1)
                    Next k
                Case "dashdot"
                    x = x0
                    For k = 0 To 3
                        col.Add UI_Box(f, prefix & "a" & k, x, yc, 10, th, 0, -1)
                        col.Add UI_Box(f, prefix & "b" & k, x + 13, yc, 3, th, 0, -1)
                        x = x + 18
                    Next k
                Case "double"
                    col.Add UI_Box(f, prefix & "1", x0, y0 + UI_PrevH / 2 - 2.5, UI_PrevW, 1.2, 0, -1)
                    col.Add UI_Box(f, prefix & "2", x0, y0 + UI_PrevH / 2 + 1.5, UI_PrevW, 1.2, 0, -1)
                Case Else   ' none
                    col.Add UI_Box(f, prefix & "1", x0, y0 + 6, UI_PrevW, UI_PrevH - 12, UI_ClrWindow, UI_ClrLine, "없음", 8)
                    col(col.Count).ForeColor = UI_ClrSub
            End Select
        Case "merge"
            ' 라벨은 글자를 세로 가운데에 놓지 못하므로 상자(테두리)와 글자 라벨을 따로 만든다
            If LCase$(arg) = "before" Then
                For r = 1 To 3
                    col.Add UI_Box(f, prefix & r, x0, y0 + (r - 1) * 10, UI_PrevW, 10, vbWhite, UI_ClrLine)
                    col.Add UI_Box(f, prefix & "t" & r, x0 + 1, y0 + (r - 1) * 10 + 1, UI_PrevW - 2, 8, vbWhite, -1, "가나다", 6)
                Next r
            Else
                col.Add UI_Box(f, prefix & "1", x0, y0, UI_PrevW, UI_PrevH, vbWhite, UI_ClrLine)
                col.Add UI_Box(f, prefix & "t1", x0 + 1, y0 + (UI_PrevH - 12) / 2, UI_PrevW - 2, 12, vbWhite, -1, "가나다", 8)
            End If
    End Select
    Set UI_DrawPreview = col
End Function

' 체크 항목(굵기 등) 라벨 스타일. 캡션은 폼이 "■ / □" 로 갱신
Public Sub UI_StyleExtra(ByVal lbl As Object, ByVal text As String, ByVal px As Single, ByVal py As Single, ByVal pw As Single)
    With lbl
        .caption = "□ " & text
        .Font.name = UI_FontName
        .Font.Size = UI_FontBody
        .Font.Bold = False
        .ForeColor = UI_ClrText
        .BackStyle = 1
        .BackColor = UI_ClrWindow
        .BorderStyle = 0
        .TextAlign = 1
        .WordWrap = False
        .Left = px
        .top = py
        .Width = pw
        .Height = 22
    End With
End Sub

' 옵션 카드 왼쪽 번호 라벨("1.") - 본문 라벨(UI_StyleOptionText)과 첫 줄 높이를 맞춘다
Public Sub UI_StyleOptionNum(ByVal lbl As Object, ByVal card As Object, ByVal idx As Long, ByVal textLbl As Object)
    With lbl
        .caption = idx & "."
        .Font.name = UI_FontName
        .Font.Size = UI_FontOption
        .Font.Bold = False
        .ForeColor = UI_ClrText
        .BackStyle = 0
        .BorderStyle = 0
        .TextAlign = 1
        .WordWrap = False
        .AutoSize = False
        .Left = card.Left + 14
        .Width = UI_NumW
        .Height = 18
        .top = textLbl.top
    End With
End Sub


