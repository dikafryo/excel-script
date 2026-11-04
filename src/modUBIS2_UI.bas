Attribute VB_Name = "modUBIS2_UI"
Option Explicit

' =====================================================================================
' modUBIS2_UI : 리본 진입점, 시작 안내, 파일 선택, 버튼 처리
' =====================================================================================

Private Const C_NAVY As Long = &H67401A
Private Const C_TEXT As Long = &H1C1C1E

Public Sub U2_Start(Optional ByVal pick As Boolean = True)
    Dim wb As Workbook

    On Error GoTo EH
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Set wb = Workbooks.Add
    wb.Activate
    Set gU2Book = wb
    U2_ResetData

    U2_BuildGuide wb
    U2_BuildOptions wb, True
    U2_OnlyFourSheets wb

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    wb.Worksheets(U2_SH_GUIDE).Activate
    If pick Then U2_PickAndImport
    Exit Sub

EH:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "UBIS 예산 가져오기를 시작하지 못했습니다." & vbCrLf & Err.Description, vbCritical, "UBIS 예산"
End Sub

Public Sub U2_PickAndImport()
    Dim files As Variant
    Dim msg As String

    If gU2Book Is Nothing Then Set gU2Book = U2_Book()
    If gU2Book Is Nothing Then
        MsgBox "먼저 리본의 [UBIS 예산 가져오기]를 실행하세요.", vbExclamation, "UBIS 예산"
        Exit Sub
    End If

    files = Application.GetOpenFilename( _
        "Excel 파일 (*.xlsx;*.xls;*.xlsm), *.xlsx;*.xls;*.xlsm", , _
        "UBIS에서 내려받은 세입·세출 파일을 모두 선택하세요 (Ctrl 또는 Shift로 여러 개 선택)", , True)
    If VarType(files) = vbBoolean Then Exit Sub

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    msg = U2_ImportFiles(files)

    If gU2IncN = 0 And gU2ExpN = 0 Then
        Application.ScreenUpdating = True
        Application.DisplayAlerts = True
        MsgBox msg, vbExclamation, "UBIS 예산"
        U2_ShowSheet U2_CurrentBook(), U2_SH_GUIDE
        Exit Sub
    End If

    U2_BuildAll gU2Book, False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    MsgBox msg & vbCrLf & vbCrLf & "시작 화면에서 결과를 확인하세요.", vbInformation, "UBIS 예산"
End Sub

' ------------------------------------------------------------------ 버튼

Public Sub U2_AddButton(ByVal ws As Worksheet, ByVal nm As String, ByVal caption As String, _
                        ByVal proc As String, ByVal leftPos As Single, ByVal topPos As Single, _
                        ByVal height As Single, ByVal primary As Boolean)
    Dim s As Shape
    Dim w As Single

    On Error Resume Next
    ws.Shapes(nm).Delete
    On Error GoTo 0

    w = Application.Max(90, Len(caption) * 13)
    Set s = ws.Shapes.AddShape(msoShapeRoundedRectangle, leftPos, topPos, w, height)
    s.Name = nm
    s.OnAction = "U2_ShapeClick"
    s.AlternativeText = proc
    s.Placement = xlFreeFloating
    s.Fill.Solid
    If primary Then
        s.Fill.ForeColor.RGB = C_NAVY
    Else
        s.Fill.ForeColor.RGB = RGB(235, 241, 249)
    End If
    s.Line.Visible = msoFalse
    s.Shadow.Visible = msoFalse

    With s.TextFrame2
        .MarginLeft = 8
        .MarginRight = 8
        .MarginTop = 2
        .MarginBottom = 2
        .VerticalAnchor = msoAnchorMiddle
        .WordWrap = msoTrue
        .AutoSize = msoAutoSizeNone
        With .TextRange
            .Text = caption
            .Font.Name = "맑은 고딕"
            .Font.NameFarEast = "맑은 고딕"
            .Font.Size = 10
            .Font.Bold = msoTrue
            If primary Then
                .Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            Else
                .Font.Fill.ForeColor.RGB = C_NAVY
            End If
            .ParagraphFormat.Alignment = msoAlignCenter
        End With
    End With
End Sub

Public Sub U2_ShapeClick()
    Dim s As Shape
    Dim proc As String
    On Error GoTo EH
    Set s = ActiveSheet.Shapes(CStr(Application.Caller))
    proc = s.AlternativeText
    U2_Dispatch proc
    Exit Sub
EH:
    MsgBox "버튼 연결을 확인해 주세요: " & Err.Description, vbExclamation, "UBIS 예산"
End Sub

Public Sub U2_Dispatch(ByVal proc As String)
    Select Case proc
        Case "pick"
            U2_PickAndImport
        Case "options"
            U2_ShowSheet U2_CurrentBook(), U2_SH_OPTIONS
        Case "summary", "start"
            U2_ShowSheet U2_CurrentBook(), U2_SH_GUIDE
        Case "income"
            U2_ShowSheet U2_CurrentBook(), U2_SH_INCOME
        Case "expense"
            U2_ShowSheet U2_CurrentBook(), U2_SH_EXPENSE
        Case "help"
            U2_ShowHelp
        Case "apply"
            U2_ApplyFilters
        Case "reset"
            U2_ResetFilters
        Case "print"
            U2_Print
        Case "save"
            U2_Save
        Case Else
            MsgBox "알 수 없는 버튼입니다: " & proc, vbExclamation, "UBIS 예산"
    End Select
End Sub

Public Sub U2_ShowHelp()
    MsgBox "UBIS 예산 파일 내려받기" & vbCrLf & vbCrLf & _
           "[세입예산]" & vbCrLf & _
           "1. UBIS > 세입 > 세입요구 > 본예산 > 조회 > 세입예산요구서 클릭" & vbCrLf & _
           "2. 추경이 있으면 세입 > 세입요구 > 추경예산(1차, 2차...) > 조회 > 세입예산요구서 클릭" & vbCrLf & vbCrLf & _
           "[세출예산]" & vbCrLf & _
           "1. UBIS > 세출 > 세출요구 > 본예산 > 조회 > 세출예산요구서 클릭" & vbCrLf & _
           "2. 추경이 있으면 세출 > 세출요구 > 추경예산(1차, 2차...) > 조회 > 세출예산요구서 클릭" & vbCrLf & vbCrLf & _
           "내려받은 파일을 모두 선택한 뒤 [파일 선택하고 분석하기]를 누르세요.", _
           vbInformation, "UBIS 예산 사용 방법"
End Sub

Public Sub U2_Print()
    On Error Resume Next
    ActiveSheet.PrintPreview
    On Error GoTo 0
End Sub

Public Sub U2_Save()
    Dim fp As Variant
    Dim wb As Workbook
    Set wb = gU2Book
    If wb Is Nothing Then Set wb = ActiveWorkbook
    If wb Is Nothing Then Exit Sub

    fp = Application.GetSaveAsFilename( _
        InitialFileName:="UBIS 예산 관리_" & Format$(Date, "yyyymmdd") & ".xlsx", _
        FileFilter:="Excel 통합 문서 (*.xlsx), *.xlsx")
    If VarType(fp) = vbBoolean Then Exit Sub

    On Error GoTo EH
    Application.DisplayAlerts = False
    wb.SaveAs Filename:=CStr(fp), FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = True
    MsgBox "저장했습니다." & vbCrLf & CStr(fp), vbInformation, "UBIS 예산"
    Exit Sub

EH:
    Application.DisplayAlerts = True
    MsgBox "저장하지 못했습니다." & vbCrLf & Err.Description, vbExclamation, "UBIS 예산"
End Sub

' ------------------------------------------------------------------ 자체 점검

Public Function U2_SelfTest(ByVal folder As String, ByVal outFile As String) As String
    Dim files() As String
    Dim n As Long
    Dim f As String
    Dim wb As Workbook
    Dim msg As String

    U2_ResetData
    f = Dir(folder & "\*.xlsx")
    Do While f <> ""
        If InStr(1, f, "~$", vbTextCompare) = 0 Then
            n = n + 1
            ReDim Preserve files(1 To n)
            files(n) = folder & "\" & f
        End If
        f = Dir
    Loop

    If n = 0 Then
        U2_SelfTest = "NO FILES"
        Exit Function
    End If

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    gU2Silent = True

    Set wb = Workbooks.Add
    Set gU2Book = wb
    msg = U2_ImportFiles(files)
    U2_BuildAll wb, True
    If gU2BuildError <> "" Then
        msg = msg & " | BUILD ERROR: " & gU2BuildError
    End If
    gU2Silent = False

    On Error GoTo EH
    wb.SaveAs Filename:=outFile, FileFormat:=xlOpenXMLWorkbook
    wb.Close SaveChanges:=False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    U2_SelfTest = "OK inc=" & gU2IncN & " exp=" & gU2ExpN & " files=" & n & " | " & msg
    Exit Function

EH:
    gU2Silent = False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    U2_SelfTest = "ERROR " & Err.Number & ": " & Err.Description & " | " & msg & _
                  IIf(gU2BuildError <> "", " | BUILD: " & gU2BuildError, "")
End Function

Public Function U2_Version() As String
    U2_Version = "UBIS 예산 가져오기 v1.0"
End Function
