Attribute VB_Name = "modUBIS2_Build"
Option Explicit

' =====================================================================================
' modUBIS2_Build : UBIS 예산 분석 화면(요약·세입·세출·옵션) 만들기
'   - 원본 데이터는 modUBIS2_Data 가 숨김 시트에 보관한다.
'   - 이 모듈은 옵션 값에 따라 화면을 다시 그린다.
' =====================================================================================

Private Const C_NAVY As Long = &H67401A          ' RGB(26,64,103)
Private Const C_BLUE As Long = &HFF7A00          ' RGB(0,122,255)
Private Const C_LIGHT As Long = &HF7F2F2         ' RGB(242,242,247)
Private Const C_LINE As Long = &HDCD6D6          ' RGB(214,214,220)
Private Const C_TEXT As Long = &H1C1C1E          ' RGB(28,28,30)
Private Const C_SUB As Long = &H8E8E8E           ' RGB(142,142,142)
Private Const C_GREEN As Long = &H53A834         ' RGB(52,168,83)
Private Const C_ORANGE As Long = &H0099FF        ' RGB(255,153,0)
Private Const C_PURPLE As Long = &HA0508B        ' RGB(139,80,160)

Public Sub U2_BuildAll(ByVal wb As Workbook, Optional ByVal resetOptions As Boolean = False)
    Dim oldCalc As XlCalculation
    Dim oldScreen As Boolean
    Dim oldAlerts As Boolean

    Set gU2Book = wb
    oldCalc = Application.Calculation
    oldScreen = Application.ScreenUpdating
    oldAlerts = Application.DisplayAlerts

    gU2BuildError = ""
    On Error GoTo EH
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    gU2Trace = "안내 화면"
    U2_BuildGuide wb
    gU2Trace = "옵션 화면"
    U2_BuildOptions wb, resetOptions
    gU2Trace = "세입 화면"
    U2_BuildIncome wb
    gU2Trace = "세출 화면"
    U2_BuildExpense wb
    gU2Trace = "세출 요약"
    U2_BuildExpenseSummary wb
    gU2Trace = "요약 화면"
    U2_BuildSummary wb

    Application.Calculation = oldCalc
    Application.ScreenUpdating = oldScreen
    Application.DisplayAlerts = oldAlerts
    U2_ShowSheet wb, U2_SH_SUMMARY
    Exit Sub

EH:
    Application.Calculation = oldCalc
    Application.ScreenUpdating = oldScreen
    Application.DisplayAlerts = oldAlerts
    gU2BuildError = gU2Trace & " / " & Err.Number & ": " & Err.Description
    If Not gU2Silent Then
        MsgBox "화면을 만드는 중 오류가 발생했습니다." & vbCrLf & gU2BuildError, vbCritical, "UBIS 예산"
    End If
End Sub

' ------------------------------------------------------------------ 시작 안내 시트

Public Sub U2_BuildGuide(ByVal wb As Workbook)
    Dim ws As Worksheet
    Dim lastRow As Long

    Set ws = U2_GetSheet(wb, U2_SH_GUIDE, False)
    If ws Is Nothing Then
        Set ws = wb.Worksheets(1)
        ws.Name = U2_SH_GUIDE
    End If

    ws.Cells.Clear
    ws.Cells.Interior.Color = RGB(255, 255, 255)
    U2_DeleteShapes ws

    With ws.Range("A1:J2")
        .Merge
        .Value2 = "UBIS 예산 가져오기"
        .Font.Name = "맑은 고딕"
        .Font.Size = 22
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = C_NAVY
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .IndentLevel = 2
    End With
    ws.Rows(1).RowHeight = 34
    ws.Rows(2).RowHeight = 10

    With ws.Range("A3:J3")
        .Merge
        .Value2 = "UBIS에서 내려받은 세입·세출 예산요구서를 한 번에 분석해 한눈에 보여 줍니다."
        .Font.Name = "맑은 고딕"
        .Font.Size = 12
        .Font.Color = C_TEXT
        .VerticalAlignment = xlCenter
        .IndentLevel = 2
    End With
    ws.Rows(3).RowHeight = 26

    With ws.Range("A5:E12")
        .Merge
        .Value2 = "[세입예산]" & vbLf & vbLf & _
                  "1. UBIS > 세입 > 세입요구 > 본예산 > 조회 > 세입예산요구서 클릭" & vbLf & vbLf & _
                  "2. 추경이 있으면" & vbLf & _
                  "   세입 > 세입요구 > 추경예산(1차, 2차...)" & vbLf & _
                  "   > 조회 > 세입예산요구서 클릭"
        .Font.Name = "맑은 고딕"
        .Font.Size = 11
        .Font.Color = C_TEXT
        .Interior.Color = RGB(240, 247, 255)
        .WrapText = True
        .VerticalAlignment = xlTop
        .HorizontalAlignment = xlLeft
        .IndentLevel = 1
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(200, 220, 245)
    End With

    With ws.Range("F5:J12")
        .Merge
        .Value2 = "[세출예산]" & vbLf & vbLf & _
                  "1. UBIS > 세출 > 세출요구 > 본예산 > 조회 > 세출예산요구서 클릭" & vbLf & vbLf & _
                  "2. 추경이 있으면" & vbLf & _
                  "   세출 > 세출요구 > 추경예산(1차, 2차...)" & vbLf & _
                  "   > 조회 > 세출예산요구서 클릭"
        .Font.Name = "맑은 고딕"
        .Font.Size = 11
        .Font.Color = C_TEXT
        .Interior.Color = RGB(240, 250, 240)
        .WrapText = True
        .VerticalAlignment = xlTop
        .HorizontalAlignment = xlLeft
        .IndentLevel = 1
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(200, 235, 200)
    End With

    With ws.Range("A13:J13")
        .Merge
        .Value2 = "내려받은 파일을 모두 선택하세요. (본예산 + 추경 1차·2차·3차 …) 세입·세출 파일을 한 번에 선택해도 됩니다."
        .Font.Name = "맑은 고딕"
        .Font.Size = 11
        .Font.Bold = True
        .Font.Color = C_NAVY
        .Interior.Color = RGB(255, 249, 230)
        .WrapText = True
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlLeft
        .IndentLevel = 1
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(245, 220, 150)
    End With
    ws.Rows(13).RowHeight = 26

    ws.Rows(5).RowHeight = 24
    ws.Rows(6).RowHeight = 20
    ws.Rows(7).RowHeight = 20
    ws.Rows(8).RowHeight = 20
    ws.Rows(9).RowHeight = 20
    ws.Rows(10).RowHeight = 20
    ws.Rows(11).RowHeight = 20
    ws.Rows(12).RowHeight = 20

    U2_AddButton ws, "u2_pick", "1. 파일 선택하고 분석하기", "pick", 12, 290, 40, True
    U2_AddButton ws, "u2_prev", "2. 결과 보기", "summary", 410, 290, 40, False
    U2_AddButton ws, "u2_optguide", "옵션 설정", "options", 550, 290, 40, False

    With ws.Range("A18:J18")
        .Merge
        .Value2 = "마지막 가져오기: " & IIf(gU2LastImport = "", "아직 없음", gU2LastImport) & _
                  IIf(gU2LastFiles > 0, "   |   읽은 파일 " & gU2LastFiles & "개", "") & _
                  IIf(gU2LastErrors <> "", "   |   확인할 파일 " & U2_LineCount(gU2LastErrors) & "개", "")
        .Font.Name = "맑은 고딕"
        .Font.Size = 10
        .Font.Color = C_SUB
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlLeft
        .IndentLevel = 2
    End With

    ws.Columns("A:J").ColumnWidth = 16
    ws.Rows(4).RowHeight = 12
    ws.Rows(16).RowHeight = 12
    ws.Rows(17).RowHeight = 12
    ws.Rows(19).RowHeight = 12
    ws.Activate
    ws.Range("A1").Select
End Sub

' ------------------------------------------------------------------ 옵션 시트

Public Sub U2_BuildOptions(ByVal wb As Workbook, ByVal reset As Boolean)
    Dim ws As Worksheet
    Dim old As Object
    Dim i As Long
    Dim arrRound As Variant
    Dim arrDept As Variant
    Dim arrSub As Variant
    Dim arrCode As Variant
    Dim arrItem As Variant

    Set ws = U2_GetSheet(wb, U2_SH_OPTIONS, False)
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = U2_SH_OPTIONS
        reset = True
    End If

    If reset Then
        ws.Cells.Clear
        U2_DeleteShapes ws
    Else
        Set old = U2_ReadOptions()
    End If

    With ws.Range("A1:H2")
        .Merge
        .Value2 = "옵션 설정"
        .Font.Name = "맑은 고딕"
        .Font.Size = 18
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = C_NAVY
        .VerticalAlignment = xlCenter
        .IndentLevel = 1
    End With
    ws.Rows(1).RowHeight = 30
    ws.Rows(2).RowHeight = 6

    U2_OptionLabel ws, "A4", "공통"
    U2_OptionLabel ws, "A5", "차수"
    U2_OptionLabel ws, "A6", "금액 단위"
    U2_OptionLabel ws, "A8", "세입"
    U2_OptionLabel ws, "A9", "사업항목"
    U2_OptionLabel ws, "A11", "세출"
    U2_OptionLabel ws, "A12", "부서"
    U2_OptionLabel ws, "A13", "세부사업"
    U2_OptionLabel ws, "A14", "세목"
    U2_OptionLabel ws, "A15", "신규여부"
    U2_OptionLabel ws, "A16", "공약사업"
    U2_OptionLabel ws, "A17", "홍보예산"
    U2_OptionLabel ws, "A18", "명시이월"
    U2_OptionLabel ws, "A19", "본청주관"
    U2_OptionLabel ws, "A20", "사전절차"

    ws.Range("M2:X400").ClearContents
    U2_WriteList ws, 1, U2_RoundList()
    U2_WriteList ws, 2, Array("원", "천원", "백만원")
    U2_WriteList ws, 3, U2_UniqueIncomeItems()
    U2_WriteList ws, 4, U2_UniqueExpenseDepts()
    U2_WriteList ws, 5, U2_UniqueExpenseSubBiz()
    U2_WriteList ws, 6, U2_UniqueExpenseCodes()
    U2_WriteList ws, 7, Array("전체", "Y", "N")
    U2_WriteList ws, 8, Array("전체", "Y", "N")
    U2_WriteList ws, 9, Array("전체", "Y", "N")
    U2_WriteList ws, 10, Array("전체", "Y", "N")
    U2_WriteList ws, 11, Array("전체", "Y", "N")
    U2_WriteList ws, 12, Array("전체", "중기재정계획", "재정투자심사", "지방보조금", "정책연구용역심의", _
                              "공유재산관리계획", "정수물품취득", "국외연수심의위원회", "사무위탁·대행")

    U2_SetValidation ws, "B5", "=$M$2:$M$40"
    U2_SetValidation ws, "B6", "=$N$2:$N$10"
    U2_SetValidation ws, "B9", "=$O$2:$O$300"
    U2_SetValidation ws, "B12", "=$P$2:$P$300"
    U2_SetValidation ws, "B13", "=$Q$2:$Q$300"
    U2_SetValidation ws, "B14", "=$R$2:$R$300"
    U2_SetValidation ws, "B15", "=$S$2:$S$10"
    U2_SetValidation ws, "B16", "=$T$2:$T$10"
    U2_SetValidation ws, "B17", "=$U$2:$U$10"
    U2_SetValidation ws, "B18", "=$V$2:$V$10"
    U2_SetValidation ws, "B19", "=$W$2:$W$10"
    U2_SetValidation ws, "B20", "=$X$2:$X$20"

    If reset Or old Is Nothing Then
        ws.Range("B5").Value2 = "전체"
        ws.Range("B6").Value2 = "천원"
        ws.Range("B9").Value2 = "전체"
        ws.Range("B12").Value2 = "전체"
        ws.Range("B13").Value2 = "전체"
        ws.Range("B14").Value2 = "전체"
        ws.Range("B15").Value2 = "전체"
        ws.Range("B16").Value2 = "전체"
        ws.Range("B17").Value2 = "전체"
        ws.Range("B18").Value2 = "전체"
        ws.Range("B19").Value2 = "전체"
        ws.Range("B20").Value2 = "전체"
    Else
        U2_RestoreOption ws, "B5", old("Round")
        U2_RestoreOption ws, "B6", old("Unit")
        U2_RestoreOption ws, "B9", old("IncItem")
        U2_RestoreOption ws, "B12", old("Dept")
        U2_RestoreOption ws, "B13", old("SubBiz")
        U2_RestoreOption ws, "B14", old("ItemCode")
        U2_RestoreOption ws, "B15", old("New")
        U2_RestoreOption ws, "B16", old("Promise")
        U2_RestoreOption ws, "B17", old("Promo")
        U2_RestoreOption ws, "B18", old("Carry")
        U2_RestoreOption ws, "B19", old("Central")
        U2_RestoreOption ws, "B20", old("Pre")
    End If

    With ws.Range("B5:B20")
        .Font.Name = "맑은 고딕"
        .Font.Size = 11
        .Interior.Color = RGB(250, 250, 252)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = C_LINE
        .HorizontalAlignment = xlLeft
        .IndentLevel = 1
    End With

    U2_AddButton ws, "u2_apply", "필터 적용", "apply", 30, 456, 34, True
    U2_AddButton ws, "u2_reset", "초기화", "reset", 150, 456, 34, False
    U2_AddButton ws, "u2_sum", "요약 보기", "summary", 270, 456, 34, False
    U2_AddButton ws, "u2_inc", "세입 보기", "income", 390, 456, 34, False
    U2_AddButton ws, "u2_exp", "세출 보기", "expense", 510, 456, 34, False

    ws.Columns("A").ColumnWidth = 16
    ws.Columns("B").ColumnWidth = 34
    ws.Columns("C:L").ColumnWidth = 3
    ws.Columns("M:X").Hidden = True

    ws.Range("A22:H22").Merge
    ws.Range("A22").Value2 = "필터를 바꾼 뒤 [필터 적용]을 누르면 세입·세출·요약 화면이 다시 만들어집니다."
    ws.Range("A22").Font.Name = "맑은 고딕"
    ws.Range("A22").Font.Size = 10
    ws.Range("A22").Font.Color = C_SUB
    ws.Range("A22").VerticalAlignment = xlCenter
    ws.Range("A22").IndentLevel = 1
End Sub

Private Sub U2_OptionLabel(ByVal ws As Worksheet, ByVal addr As String, ByVal text As String)
    With ws.Range(addr)
        .Value2 = text
        .Font.Name = "맑은 고딕"
        .Font.Size = 11
        .Font.Bold = (Len(text) <= 4)
        .Font.Color = C_TEXT
        .VerticalAlignment = xlCenter
        .IndentLevel = 1
    End With
End Sub

Private Sub U2_SetValidation(ByVal ws As Worksheet, ByVal addr As String, ByVal formula As String)
    With ws.Range(addr).Validation
        .Delete
        .Add xlValidateList, xlValidAlertStop, xlBetween, formula
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
End Sub

Private Sub U2_RestoreOption(ByVal ws As Worksheet, ByVal addr As String, ByVal value As String)
    Dim rng As Range
    Set rng = ws.Range(addr)
    On Error Resume Next
    rng.Value2 = value
    If Err.Number <> 0 Then rng.Value2 = "전체"
    On Error GoTo 0
End Sub

Private Sub U2_WriteList(ByVal ws As Worksheet, ByVal col As Long, ByVal values As Variant)
    Dim i As Long
    Dim c As Long
    c = 12 + col
    For i = LBound(values) To UBound(values)
        ws.Cells(i - LBound(values) + 2, c).Value2 = values(i)
    Next i
End Sub

' ------------------------------------------------------------------ 세입 시트

Private Sub U2_BuildIncome(ByVal wb As Workbook)
    Dim ws As Worksheet
    Dim opt As Object
    Dim unit As Double
    Dim maxR As Long
    Dim items As Object
    Dim i As Long
    Dim r As Long
    Dim c As Long
    Dim key As Variant
    Dim keys() As String
    Dim vals() As Double
    Dim dept As String
    Dim n As Long
    Dim sumR As Long
    Dim rec As U2IncomeRec
    Dim detailRow As Long
    Dim lastCol As Long
    Dim outArr As Variant
    Dim j As Long
    Dim kk As Long
    Dim tmpRow() As Variant

    Set ws = U2_ResetSheet(wb, U2_SH_INCOME)
    Set opt = U2_ReadOptions()
    unit = U2_UnitValue(opt("Unit"))
    maxR = U2_MaxRound()

    U2_Title ws, "세입 예산", "사업항목별 세입을 본예산·추경 차수별로 비교합니다. (금액 단위: " & CStr(opt("Unit")) & ")", 16

    Set items = CreateObject("Scripting.Dictionary")
    For i = 1 To gU2IncN
        If U2_IncPass(gU2Inc(i), opt, True) Then
            key = gU2Inc(i).Item
            If key <> "" Then items(key) = 1
        End If
    Next i

    If items.Count = 0 Then
        ws.Range("A4").Value2 = "선택한 조건에 맞는 세입 자료가 없습니다."
        Exit Sub
    End If

    ReDim keys(1 To items.Count)
    n = 0
    For Each key In items.Keys
        n = n + 1
        keys(n) = CStr(key)
    Next key

    U2_Header ws, 4, 1, U2_IncomeSummaryHeaders(maxR)

    lastCol = 3 + maxR + 2
    ReDim outArr(1 To n, 1 To lastCol)
    For i = 1 To n
        key = keys(i)
        ReDim vals(0 To maxR)
        dept = ""
        For sumR = 0 To maxR
            vals(sumR) = U2_IncomeRoundSum(key, sumR, opt)
            If dept = "" Then dept = U2_IncomeDept(key, sumR)
        Next sumR
        ' 없는 차수는 앞 차수 값을 이어받는다.
        For sumR = 1 To maxR
            If Not U2_IncomeRoundExists(key, sumR) Then vals(sumR) = vals(sumR - 1)
        Next sumR
        outArr(i, 1) = key
        outArr(i, 2) = dept
        For sumR = 0 To maxR
            outArr(i, 3 + sumR) = vals(sumR) / unit
        Next sumR
        outArr(i, lastCol - 1) = vals(maxR) / unit
        outArr(i, lastCol) = (vals(maxR) - vals(0)) / unit
    Next i

    ws.Range("A5").Resize(n, lastCol).Value2 = outArr
    ws.Range(ws.Cells(5, 1), ws.Cells(4 + n, lastCol)).Sort _
        Key1:=ws.Cells(5, lastCol - 1), Order1:=xlDescending, Header:=xlNo

    r = 5 + n
    With ws.Range(ws.Cells(5, 1), ws.Cells(r - 1, 3 + maxR + 2))
        .Borders.LineStyle = xlContinuous
        .Borders.Color = C_LINE
        .Font.Name = "맑은 고딕"
        .Font.Size = 10
    End With
    ws.Range(ws.Cells(5, 3), ws.Cells(r - 1, 3 + maxR + 2)).NumberFormat = "#,##0"
    ws.Range(ws.Cells(5, 3 + maxR + 2), ws.Cells(r - 1, 3 + maxR + 2)).Font.Color = C_BLUE

    detailRow = r + 2
    ws.Cells(detailRow - 1, 1).Value2 = "세입 상세"
    ws.Cells(detailRow - 1, 1).Font.Bold = True
    ws.Cells(detailRow - 1, 1).Font.Size = 12
    ws.Cells(detailRow - 1, 1).Font.Color = C_NAVY

    U2_Header ws, detailRow, 1, Array("차수", "장", "관", "항", "목", "사업항목", "소관부서", _
                                      "기정예산", "요구액", "최종예산", "증감", "비고")
    r = detailRow + 1
    For i = 1 To gU2IncN
        If U2_IncPass(gU2Inc(i), opt, False) Then
            ws.Cells(r, 1).Value2 = gU2Inc(i).RoundName
            ws.Cells(r, 2).Value2 = gU2Inc(i).Jang
            ws.Cells(r, 3).Value2 = gU2Inc(i).Gwan
            ws.Cells(r, 4).Value2 = gU2Inc(i).Hang
            ws.Cells(r, 5).Value2 = gU2Inc(i).Mok
            ws.Cells(r, 6).Value2 = gU2Inc(i).Item
            ws.Cells(r, 7).Value2 = gU2Inc(i).Dept
            ws.Cells(r, 8).Value2 = gU2Inc(i).PrevAmt / unit
            ws.Cells(r, 9).Value2 = gU2Inc(i).ReqAmt / unit
            ws.Cells(r, 10).Value2 = gU2Inc(i).FinalAmt / unit
            ws.Cells(r, 11).Value2 = gU2Inc(i).ChangeAmt / unit
            ws.Cells(r, 12).Value2 = gU2Inc(i).Note
            r = r + 1
        End If
    Next i

    If r > detailRow + 1 Then
        With ws.Range(ws.Cells(detailRow + 1, 1), ws.Cells(r - 1, 12))
            .Borders.LineStyle = xlContinuous
            .Borders.Color = C_LINE
            .Font.Name = "맑은 고딕"
            .Font.Size = 10
        End With
        ws.Range(ws.Cells(detailRow + 1, 8), ws.Cells(r - 1, 11)).NumberFormat = "#,##0"
        ws.Range(ws.Cells(detailRow + 1, 1), ws.Cells(r - 1, 12)).AutoFilter
    End If

    ws.Columns("A").ColumnWidth = 30
    ws.Columns("B").ColumnWidth = 18
    ws.Columns("C:Z").ColumnWidth = 13
    ws.Rows(4).RowHeight = 30
    ws.Rows(detailRow).RowHeight = 28
End Sub

Private Function U2_IncomeSummaryHeaders(ByVal maxR As Long) As Variant
    Dim a() As String
    Dim i As Long
    ReDim a(0 To 3 + maxR + 2)
    a(0) = "사업항목"
    a(1) = "소관부서"
    a(2) = "본예산"
    For i = 1 To maxR
        a(2 + i) = CStr(i) & "차 추경"
    Next i
    a(3 + maxR) = "최종예산"
    a(4 + maxR) = "증감"
    U2_IncomeSummaryHeaders = a
End Function

Private Sub U2_SortArray(ByRef arr As Variant, ByVal n As Long, ByVal nCol As Long, ByVal sortCol As Long)
    Dim i As Long, j As Long, c As Long
    Dim v1 As Double, v2 As Double
    Dim tmp As Variant
    For i = 1 To n - 1
        For j = i + 1 To n
            v1 = U2_Num(arr(i, sortCol))
            v2 = U2_Num(arr(j, sortCol))
            If v2 > v1 Then
                For c = 1 To nCol
                    tmp = arr(i, c)
                    arr(i, c) = arr(j, c)
                    arr(j, c) = tmp
                Next c
            End If
        Next j
    Next i
End Sub

Private Function U2_IncomeRoundSum(ByVal key As String, ByVal roundNo As Long, ByRef opt As Object) As Double
    Dim i As Long
    Dim total As Double
    For i = 1 To gU2IncN
        If gU2Inc(i).Item = key And gU2Inc(i).RoundNo = roundNo Then
            If U2_IncPass(gU2Inc(i), opt, True) Then total = total + gU2Inc(i).FinalAmt
        End If
    Next i
    U2_IncomeRoundSum = total
End Function

Private Function U2_IncomeRoundExists(ByVal key As String, ByVal roundNo As Long) As Boolean
    Dim i As Long
    For i = 1 To gU2IncN
        If gU2Inc(i).Item = key And gU2Inc(i).RoundNo = roundNo Then
            U2_IncomeRoundExists = True
            Exit Function
        End If
    Next i
End Function

Private Function U2_IncomeDept(ByVal key As String, ByVal roundNo As Long) As String
    Dim i As Long
    For i = 1 To gU2IncN
        If gU2Inc(i).Item = key And gU2Inc(i).RoundNo = roundNo And gU2Inc(i).Dept <> "" Then
            U2_IncomeDept = gU2Inc(i).Dept
            Exit Function
        End If
    Next i
End Function

' ------------------------------------------------------------------ 세출 시트

Private Sub U2_BuildExpense(ByVal wb As Workbook)
    Dim ws As Worksheet
    Dim opt As Object
    Dim unit As Double
    Dim i As Long
    Dim r As Long
    Dim n As Long
    Dim total As Double
    Dim arr As Variant
    Dim headers As Variant

    Set ws = U2_ResetSheet(wb, U2_SH_EXPENSE)
    Set opt = U2_ReadOptions()
    unit = U2_UnitValue(opt("Unit"))

    U2_Title ws, "세출 예산", "부서·세부사업·세목·산출기초·재원·담당자·사업유형을 한 표에서 봅니다. (금액 단위: " & CStr(opt("Unit")) & ")", 32

    n = 0
    total = 0
    For i = 1 To gU2ExpN
        If U2_ExpPass(gU2Exp(i), opt) Then
            n = n + 1
            total = total + gU2Exp(i).ReqAmt
        End If
    Next i

    ws.Range("A4").Value2 = IIf(CStr(opt("Round")) = "전체", "표시된 전체 차수 합계 ", "선택 차수 합계 ") & _
                            Format$(total / unit, "#,##0") & " " & CStr(opt("Unit")) & _
                            "   |   " & Format$(n, "#,##0") & "건"
    ws.Range("A4").Font.Name = "맑은 고딕"
    ws.Range("A4").Font.Size = 11
    ws.Range("A4").Font.Bold = True
    ws.Range("A4").Font.Color = C_NAVY

    headers = Array("차수", "부서명", "정책사업", "단위사업", "세부사업", "세세부사업", _
                    "산출1", "산출2", "산출3", "산출4", "세목코드", "원가통계목", _
                    "예산요구액", "재원_특교", "재원_국고", "재원_비법정", "재원_기타", "재원_자체", "재원_합계", _
                    "업무담당자", "홍보방법", "홍보예산", "신규여부", "공약사업", "명시이월", _
                    "사업기간", "명시이월액", "명시이월재원", "본청주관", "사전절차", "설명사항", "수정여부")
    U2_Header ws, 6, 1, headers

    If n > 0 Then
        ReDim arr(1 To n, 1 To 32)
        r = 0
        For i = 1 To gU2ExpN
            If U2_ExpPass(gU2Exp(i), opt) Then
                r = r + 1
                arr(r, 1) = gU2Exp(i).RoundName
                arr(r, 2) = gU2Exp(i).Dept
                arr(r, 3) = gU2Exp(i).Policy
                arr(r, 4) = gU2Exp(i).UnitBiz
                arr(r, 5) = gU2Exp(i).SubBiz
                arr(r, 6) = gU2Exp(i).SubSubBiz
                arr(r, 7) = U2_LabelText(gU2Exp(i).Calc1)
                arr(r, 8) = U2_LabelText(gU2Exp(i).Calc2)
                arr(r, 9) = U2_LabelText(gU2Exp(i).Calc3)
                arr(r, 10) = U2_LabelText(gU2Exp(i).Calc4)
                arr(r, 11) = gU2Exp(i).ItemCode
                arr(r, 12) = gU2Exp(i).CostStat
                arr(r, 13) = gU2Exp(i).ReqAmt / unit
                arr(r, 14) = gU2Exp(i).FundSpecial / unit
                arr(r, 15) = gU2Exp(i).FundNational / unit
                arr(r, 16) = gU2Exp(i).FundNonStatutory / unit
                arr(r, 17) = gU2Exp(i).FundOther / unit
                arr(r, 18) = gU2Exp(i).FundSelf / unit
                arr(r, 19) = gU2Exp(i).FundTotal / unit
                arr(r, 20) = gU2Exp(i).Manager
                arr(r, 21) = gU2Exp(i).PromoMethod
                arr(r, 22) = U2_Mark(gU2Exp(i).PromoBudget)
                arr(r, 23) = U2_Mark(gU2Exp(i).IsNew)
                arr(r, 24) = U2_Mark(gU2Exp(i).IsPromise)
                arr(r, 25) = U2_Mark(gU2Exp(i).IsCarryover)
                arr(r, 26) = gU2Exp(i).CarryPeriod
                arr(r, 27) = gU2Exp(i).CarryAmt / unit
                arr(r, 28) = gU2Exp(i).CarryFund
                arr(r, 29) = U2_Mark(gU2Exp(i).IsCentral)
                arr(r, 30) = gU2Exp(i).PreSummary
                arr(r, 31) = gU2Exp(i).Desc
                arr(r, 32) = gU2Exp(i).Revised
            End If
        Next i
        ws.Range("A7").Resize(n, 32).Value2 = arr
    End If

    With ws.Range("A6:AF" & (6 + n))
        .Borders.LineStyle = xlContinuous
        .Borders.Color = C_LINE
        .Font.Name = "맑은 고딕"
        .Font.Size = 9
        .VerticalAlignment = xlCenter
    End With
    ws.Range("A6:AF6").Interior.Color = C_NAVY
    ws.Range("A6:AF6").Font.Color = RGB(255, 255, 255)
    ws.Range("A6:AF6").Font.Bold = True
    ws.Range("A6:AF6").WrapText = True
    ws.Rows(6).RowHeight = 32

    If n > 0 Then
        ws.Range("M7:S" & (6 + n)).NumberFormat = "#,##0"
        ws.Range("AA7:AA" & (6 + n)).NumberFormat = "#,##0"
        U2_ColorFlagColumns ws, 7, n
        ws.Range("A6:AF" & (6 + n)).AutoFilter
    End If

    ws.Columns("A").ColumnWidth = 10
    ws.Columns("B").ColumnWidth = 26
    ws.Columns("C:E").ColumnWidth = 18
    ws.Columns("F").ColumnWidth = 24
    ws.Columns("G:J").ColumnWidth = 26
    ws.Columns("K").ColumnWidth = 11
    ws.Columns("L").ColumnWidth = 14
    ws.Columns("M:S").ColumnWidth = 13
    ws.Columns("T").ColumnWidth = 12
    ws.Columns("U").ColumnWidth = 16
    ws.Columns("V:Z").ColumnWidth = 10
    ws.Columns("AA").ColumnWidth = 14
    ws.Columns("AB:AF").ColumnWidth = 18
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Range("G7").Select
    ActiveWindow.FreezePanes = True
End Sub

Private Sub U2_ColorFlagColumns(ByVal ws As Worksheet, ByVal firstRow As Long, ByVal n As Long)
    If n <= 0 Then Exit Sub
    With ws.Range("V" & firstRow & ":V" & (firstRow + n - 1))
        .Interior.Color = RGB(255, 242, 230)
        .HorizontalAlignment = xlCenter
    End With
    With ws.Range("W" & firstRow & ":W" & (firstRow + n - 1))
        .Interior.Color = RGB(232, 245, 233)
        .HorizontalAlignment = xlCenter
    End With
    With ws.Range("X" & firstRow & ":X" & (firstRow + n - 1))
        .Interior.Color = RGB(255, 248, 225)
        .HorizontalAlignment = xlCenter
    End With
    With ws.Range("Y" & firstRow & ":Y" & (firstRow + n - 1))
        .Interior.Color = RGB(227, 242, 253)
        .HorizontalAlignment = xlCenter
    End With
    With ws.Range("AC" & firstRow & ":AC" & (firstRow + n - 1))
        .Interior.Color = RGB(243, 229, 245)
        .HorizontalAlignment = xlCenter
    End With
End Sub

' ------------------------------------------------------------------ 요약 시트

' ------------------------------------------------------------------ 세출 요약 시트

Private Sub U2_BuildExpenseSummary(ByVal wb As Workbook)
    Dim ws As Worksheet
    Dim opt As Object
    Dim unit As Double
    Dim dDept As Object, dSub As Object, dItem As Object
    Dim i As Long, r As Long, n As Long
    Dim k As Variant
    Dim a As Variant
    Dim total As Double
    Dim r1 As Long, r2 As Long, r3 As Long
    Dim code As String, cost As String
    Dim p As Long

    Set ws = U2_ResetSheet(wb, U2_SH_EXP_SUMMARY)
    Set opt = U2_ReadOptions()
    unit = U2_UnitValue(opt("Unit"))

    U2_Title ws, "세출 요약", "부서·세부사업·세목별로 예산을 한눈에 봅니다. (금액 단위: " & CStr(opt("Unit")) & ")", 12

    Set dDept = CreateObject("Scripting.Dictionary")
    Set dSub = CreateObject("Scripting.Dictionary")
    Set dItem = CreateObject("Scripting.Dictionary")

    For i = 1 To gU2ExpN
        If U2_ExpPassSummary(gU2Exp(i), opt) Then
            total = total + gU2Exp(i).ReqAmt
            U2_AddStat dDept, gU2Exp(i).Dept, gU2Exp(i).ReqAmt, gU2Exp(i)
            U2_AddStat dSub, gU2Exp(i).SubBiz, gU2Exp(i).ReqAmt, gU2Exp(i)
            U2_AddStat dItem, gU2Exp(i).ItemCode & " | " & gU2Exp(i).CostStat, gU2Exp(i).ReqAmt, gU2Exp(i)
        End If
    Next i

    ' 1) 부서별
    ws.Cells(4, 2).Value2 = "부서별"
    U2_SectionTitle ws, "B4:H4", "부서별"
    U2_Header ws, 5, 2, Array("부서명", "예산요구액", "건수", "신규", "공약", "홍보", "명시이월", "본청주관")
    r = 6
    For Each k In dDept.Keys
        a = dDept(k)
        ws.Cells(r, 2).Value2 = CStr(k)
        ws.Cells(r, 3).Value2 = CDbl(a(0)) / unit
        ws.Cells(r, 4).Value2 = a(1)
        ws.Cells(r, 5).Value2 = a(2)
        ws.Cells(r, 6).Value2 = a(3)
        ws.Cells(r, 7).Value2 = a(4)
        ws.Cells(r, 8).Value2 = a(5)
        ws.Cells(r, 9).Value2 = a(6)
        r = r + 1
    Next k
    r1 = r
    ws.Cells(r1, 2).Value2 = "합계"
    ws.Cells(r1, 3).Value2 = total / unit
    ws.Cells(r1, 4).Value2 = Application.WorksheetFunction.Sum(ws.Range("D6:D" & (r1 - 1)))
    ws.Cells(r1, 5).Value2 = Application.WorksheetFunction.Sum(ws.Range("E6:E" & (r1 - 1)))
    ws.Cells(r1, 6).Value2 = Application.WorksheetFunction.Sum(ws.Range("F6:F" & (r1 - 1)))
    ws.Cells(r1, 7).Value2 = Application.WorksheetFunction.Sum(ws.Range("G6:G" & (r1 - 1)))
    ws.Cells(r1, 8).Value2 = Application.WorksheetFunction.Sum(ws.Range("H6:H" & (r1 - 1)))
    ws.Cells(r1, 9).Value2 = Application.WorksheetFunction.Sum(ws.Range("I6:I" & (r1 - 1)))
    If r1 > 6 Then
        ws.Range(ws.Cells(6, 2), ws.Cells(r1 - 1, 9)).Sort Key1:=ws.Cells(6, 3), Order1:=xlDescending, Header:=xlNo
    End If
    U2_FormatBlock ws, 5, 2, r1, 9
    ws.Range(ws.Cells(6, 3), ws.Cells(r1, 3)).NumberFormat = "#,##0"
    ws.Range(ws.Cells(r1, 2), ws.Cells(r1, 9)).Font.Bold = True

    ' 2) 세부사업별
    r2 = r1 + 3
    ws.Cells(r2, 2).Value2 = "세부사업별"
    U2_SectionTitle ws, "B" & r2 & ":H" & r2, "세부사업별"
    U2_Header ws, r2 + 1, 2, Array("세부사업", "예산요구액", "건수")
    r = r2 + 2
    For Each k In dSub.Keys
        a = dSub(k)
        ws.Cells(r, 2).Value2 = CStr(k)
        ws.Cells(r, 3).Value2 = CDbl(a(0)) / unit
        ws.Cells(r, 4).Value2 = a(1)
        r = r + 1
    Next k
    ws.Cells(r, 2).Value2 = "합계"
    ws.Cells(r, 3).Value2 = total / unit
    ws.Cells(r, 4).Value2 = Application.WorksheetFunction.Sum(ws.Range("D" & (r2 + 2) & ":D" & (r - 1)))
    If r > r2 + 2 Then
        ws.Range(ws.Cells(r2 + 2, 2), ws.Cells(r - 1, 4)).Sort Key1:=ws.Cells(r2 + 2, 3), Order1:=xlDescending, Header:=xlNo
    End If
    U2_FormatBlock ws, r2 + 1, 2, r, 4
    ws.Range(ws.Cells(r2 + 2, 3), ws.Cells(r, 3)).NumberFormat = "#,##0"
    ws.Range(ws.Cells(r, 2), ws.Cells(r, 4)).Font.Bold = True

    ' 3) 세목·원가통계목별
    r3 = r + 3
    ws.Cells(r3, 2).Value2 = "세목·원가통계목별"
    U2_SectionTitle ws, "B" & r3 & ":H" & r3, "세목·원가통계목별"
    U2_Header ws, r3 + 1, 2, Array("세목코드", "원가통계목", "예산요구액", "건수")
    r = r3 + 2
    For Each k In dItem.Keys
        a = dItem(k)
        p = InStr(CStr(k), " | ")
        If p > 0 Then
            code = Left$(CStr(k), p - 1)
            cost = Mid$(CStr(k), p + 3)
        Else
            code = CStr(k)
            cost = ""
        End If
        ws.Cells(r, 2).Value2 = code
        ws.Cells(r, 3).Value2 = cost
        ws.Cells(r, 4).Value2 = CDbl(a(0)) / unit
        ws.Cells(r, 5).Value2 = a(1)
        r = r + 1
    Next k
    ws.Cells(r, 2).Value2 = "합계"
    ws.Cells(r, 4).Value2 = total / unit
    ws.Cells(r, 5).Value2 = Application.WorksheetFunction.Sum(ws.Range("E" & (r3 + 2) & ":E" & (r - 1)))
    If r > r3 + 2 Then
        ws.Range(ws.Cells(r3 + 2, 2), ws.Cells(r - 1, 5)).Sort Key1:=ws.Cells(r3 + 2, 4), Order1:=xlDescending, Header:=xlNo
    End If
    U2_FormatBlock ws, r3 + 1, 2, r, 5
    ws.Range(ws.Cells(r3 + 2, 4), ws.Cells(r, 4)).NumberFormat = "#,##0"
    ws.Range(ws.Cells(r, 2), ws.Cells(r, 5)).Font.Bold = True

    U2_AddButton ws, "u2es_sum", "요약 보기", "summary", 30, r * 15 + 40, 34, False
    U2_AddButton ws, "u2es_exp", "세출 상세", "expense", 150, r * 15 + 40, 34, False
    U2_AddButton ws, "u2es_opt", "옵션 설정", "options", 270, r * 15 + 40, 34, False
    U2_AddButton ws, "u2es_print", "인쇄", "print", 390, r * 15 + 40, 34, False

    ws.Columns("A").ColumnWidth = 3
    ws.Columns("B").ColumnWidth = 32
    ws.Columns("C").ColumnWidth = 16
    ws.Columns("D").ColumnWidth = 10
    ws.Columns("E:I").ColumnWidth = 10
End Sub

Private Sub U2_AddStat(ByRef d As Object, ByVal key As String, ByVal amt As Double, ByRef rec As U2ExpenseRec)
    Dim a As Variant
    If key = "" Then key = "(미지정)"
    If Not d.Exists(key) Then d.Add key, Array(0#, 0, 0, 0, 0, 0, 0)
    a = d(key)
    a(0) = CDbl(a(0)) + amt
    a(1) = CLng(a(1)) + 1
    If rec.IsNew = "Y" Then a(2) = CLng(a(2)) + 1
    If rec.IsPromise = "Y" Then a(3) = CLng(a(3)) + 1
    If rec.PromoBudget = "Y" Then a(4) = CLng(a(4)) + 1
    If rec.IsCarryover = "Y" Then a(5) = CLng(a(5)) + 1
    If rec.IsCentral = "Y" Then a(6) = CLng(a(6)) + 1
    d(key) = a
End Sub
Private Sub U2_BuildSummary(ByVal wb As Workbook)
    Dim ws As Worksheet
    Dim opt As Object
    Dim unit As Double
    Dim incTotal As Double
    Dim expTotal As Double
    Dim i As Long
    Dim r As Long
    Dim n As Long
    Dim newAmt As Double, newCnt As Long
    Dim proAmt As Double, proCnt As Long
    Dim promoAmt As Double, promoCnt As Long
    Dim carryAmt As Double, carryCnt As Long
    Dim centralAmt As Double, centralCnt As Long
    Dim preAmt As Double, preCnt As Long
    Dim maxR As Long
    Dim rs As Long
    Dim incR As Double, expR As Double
    Dim dSub As Object
    Dim dFund As Object
    Dim k As Variant
    Dim keys() As String
    Dim vals() As Double
    Dim subCnt As Long
    Dim tmp As Double
    Dim tmpS As String

    On Error GoTo EH
    gU2Trace = "reset"
    Set ws = U2_ResetSheet(wb, U2_SH_SUMMARY)
    Set opt = U2_ReadOptions()
    unit = U2_UnitValue(opt("Unit"))
    maxR = U2_MaxRound()

    U2_Title ws, "UBIS 예산 요약", U2_SummarySubtitle(), 17
    ws.Rows(4).RowHeight = 20
    ws.Rows(5).RowHeight = 30

    For i = 1 To gU2IncN
        If U2_IncPassSummary(gU2Inc(i), opt) Then incTotal = incTotal + gU2Inc(i).FinalAmt
    Next i

    For i = 1 To gU2ExpN
        If U2_ExpPassSummary(gU2Exp(i), opt) Then
            expTotal = expTotal + gU2Exp(i).ReqAmt
            If gU2Exp(i).IsNew = "Y" Then newAmt = newAmt + gU2Exp(i).ReqAmt: newCnt = newCnt + 1
            If gU2Exp(i).IsPromise = "Y" Then proAmt = proAmt + gU2Exp(i).ReqAmt: proCnt = proCnt + 1
            If gU2Exp(i).PromoBudget = "Y" Then promoAmt = promoAmt + gU2Exp(i).ReqAmt: promoCnt = promoCnt + 1
            If gU2Exp(i).IsCarryover = "Y" Then carryAmt = carryAmt + gU2Exp(i).ReqAmt: carryCnt = carryCnt + 1
            If gU2Exp(i).IsCentral = "Y" Then centralAmt = centralAmt + gU2Exp(i).ReqAmt: centralCnt = centralCnt + 1
            If PreAny(gU2Exp(i)) Then preAmt = preAmt + gU2Exp(i).ReqAmt: preCnt = preCnt + 1
        End If
    Next i

    U2_Kpi ws, 4, 2, "세입 최종예산", incTotal / unit, C_BLUE
    U2_Kpi ws, 4, 4, "세출 예산요구액", expTotal / unit, C_NAVY
    U2_Kpi ws, 4, 6, "세입 - 세출", (incTotal - expTotal) / unit, IIf(incTotal >= expTotal, C_GREEN, RGB(200, 60, 60))
    U2_Kpi ws, 4, 8, "신규사업", newAmt / unit, C_GREEN
    U2_Kpi ws, 4, 10, "공약사업", proAmt / unit, C_ORANGE
    U2_Kpi ws, 4, 12, "홍보예산", promoAmt / unit, RGB(180, 120, 0)
    U2_Kpi ws, 4, 14, "명시이월", carryAmt / unit, C_BLUE
    U2_Kpi ws, 4, 16, "본청주관", centralAmt / unit, C_PURPLE

    ' 차수별 요약
    ws.Cells(9, 2).Value2 = "차수별 요약"
    U2_SectionTitle ws, "B9:H9", "차수별 요약"
    U2_Header ws, 10, 2, Array("차수", "세입", "세출", "차이", "세출건수")
    r = 11
    For rs = 0 To maxR
        incR = 0: expR = 0
        n = 0
        For i = 1 To gU2IncN
            If gU2Inc(i).RoundNo = rs And U2_IncPass(gU2Inc(i), opt, True) Then incR = incR + gU2Inc(i).FinalAmt
        Next i
        For i = 1 To gU2ExpN
            If gU2Exp(i).RoundNo = rs And U2_ExpPassNoRound(gU2Exp(i), opt) Then
                expR = expR + gU2Exp(i).ReqAmt
                n = n + 1
            End If
        Next i
        ws.Cells(r, 2).Value2 = IIf(rs = 0, "본예산", CStr(rs) & "차 추경")
        ws.Cells(r, 3).Value2 = incR / unit
        ws.Cells(r, 4).Value2 = expR / unit
        ws.Cells(r, 5).Value2 = (incR - expR) / unit
        ws.Cells(r, 6).Value2 = n
        r = r + 1
    Next rs
    U2_FormatBlock ws, 10, 2, r - 1, 6
    ws.Range(ws.Cells(11, 3), ws.Cells(r - 1, 5)).NumberFormat = "#,##0"

    ' 세부사업 TOP 10
    Set dSub = CreateObject("Scripting.Dictionary")
    Dim subKey As String
    For i = 1 To gU2ExpN
        If U2_ExpPassSummary(gU2Exp(i), opt) And gU2Exp(i).SubBiz <> "" Then
            subKey = gU2Exp(i).SubBiz
            If Not dSub.Exists(subKey) Then dSub.Add subKey, 0
            dSub(subKey) = CDbl(dSub(subKey)) + gU2Exp(i).ReqAmt
        End If
    Next i
    gU2Trace = "top sub"
    U2_SectionTitle ws, "J9:P9", "세부사업 TOP 10"
    U2_Header ws, 10, 10, Array("순위", "세부사업", "예산요구액", "비중")
    r = 11
    For Each k In dSub.Keys
        ws.Cells(r, 10).Value2 = 0
        ws.Cells(r, 11).Value2 = CStr(k)
        ws.Cells(r, 12).Value2 = U2_Num(dSub(k)) / unit
        If expTotal > 0 Then ws.Cells(r, 13).Value2 = U2_Num(dSub(k)) / expTotal
        r = r + 1
    Next k
    If r > 11 Then
        ws.Range(ws.Cells(11, 10), ws.Cells(r - 1, 13)).Sort _
            Key1:=ws.Cells(11, 12), Order1:=xlDescending, Header:=xlNo
        For i = 11 To r - 1
            ws.Cells(i, 10).Value2 = i - 10
        Next i
    End If
    U2_FormatBlock ws, 10, 10, r - 1, 13
    ws.Range(ws.Cells(11, 12), ws.Cells(r - 1, 12)).NumberFormat = "#,##0"
    ws.Range(ws.Cells(11, 13), ws.Cells(r - 1, 13)).NumberFormat = "0.0%"
    gU2Trace = "fund"
    ' 재원별
    Set dFund = CreateObject("Scripting.Dictionary")
    dFund("특교") = 0
    dFund("국고") = 0
    dFund("비법정") = 0
    dFund("기타지원금") = 0
    dFund("자체") = 0
    For i = 1 To gU2ExpN
        If U2_ExpPassSummary(gU2Exp(i), opt) Then
            dFund("특교") = dFund("특교") + gU2Exp(i).FundSpecial
            dFund("국고") = dFund("국고") + gU2Exp(i).FundNational
            dFund("비법정") = dFund("비법정") + gU2Exp(i).FundNonStatutory
            dFund("기타지원금") = dFund("기타지원금") + gU2Exp(i).FundOther
            dFund("자체") = dFund("자체") + gU2Exp(i).FundSelf
        End If
    Next i
    U2_SectionTitle ws, "B24:H24", "재원별"
    U2_Header ws, 25, 2, Array("재원", "금액", "비중")
    r = 26
    For Each k In dFund.Keys
        ws.Cells(r, 2).Value2 = k
        ws.Cells(r, 3).Value2 = U2_Num(dFund(k)) / unit
        If expTotal > 0 Then ws.Cells(r, 4).Value2 = U2_Num(dFund(k)) / expTotal
        r = r + 1
    Next k
    U2_FormatBlock ws, 25, 2, r - 1, 4
    ws.Range(ws.Cells(26, 3), ws.Cells(r - 1, 3)).NumberFormat = "#,##0"
    ws.Range(ws.Cells(26, 4), ws.Cells(r - 1, 4)).NumberFormat = "0.0%"

    gU2Trace = "flags"
    ' 사업 유형별
    U2_SectionTitle ws, "J24:P24", "사업 유형별"
    U2_Header ws, 25, 10, Array("구분", "건수", "금액")
    ws.Cells(26, 10).Value2 = "신규사업": ws.Cells(26, 11).Value2 = newCnt: ws.Cells(26, 12).Value2 = newAmt / unit
    ws.Cells(27, 10).Value2 = "공약사업": ws.Cells(27, 11).Value2 = proCnt: ws.Cells(27, 12).Value2 = proAmt / unit
    ws.Cells(28, 10).Value2 = "홍보예산": ws.Cells(28, 11).Value2 = promoCnt: ws.Cells(28, 12).Value2 = promoAmt / unit
    ws.Cells(29, 10).Value2 = "명시이월": ws.Cells(29, 11).Value2 = carryCnt: ws.Cells(29, 12).Value2 = carryAmt / unit
    ws.Cells(30, 10).Value2 = "본청주관": ws.Cells(30, 11).Value2 = centralCnt: ws.Cells(30, 12).Value2 = centralAmt / unit
    ws.Cells(31, 10).Value2 = "사전절차": ws.Cells(31, 11).Value2 = preCnt: ws.Cells(31, 12).Value2 = preAmt / unit
    U2_FormatBlock ws, 25, 10, 31, 12
    ws.Range("L26:L31").NumberFormat = "#,##0"

    gU2Trace = "buttons"
    U2_AddButton ws, "u2s_inc", "세입 보기", "income", 30, 500, 34, False
    U2_AddButton ws, "u2s_exp", "세출 보기", "expense", 150, 500, 34, False
    U2_AddButton ws, "u2s_opt", "옵션 설정", "options", 270, 500, 34, False
    U2_AddButton ws, "u2s_pick", "파일 다시 선택", "pick", 390, 500, 34, False
    U2_AddButton ws, "u2s_expsum", "세출 요약", "expense_summary", 660, 500, 34, False
    U2_AddButton ws, "u2s_print", "인쇄", "print", 790, 500, 34, False

    ws.Columns("A").ColumnWidth = 3
    ws.Columns("B:H").ColumnWidth = 15
    ws.Columns("I").ColumnWidth = 3
    ws.Columns("J:P").ColumnWidth = 15
    ws.Rows(3).RowHeight = 8
    ws.Rows(8).RowHeight = 8
    ws.Rows(23).RowHeight = 8
    Exit Sub

EH:
    gU2BuildError = "Summary " & gU2Trace & " / " & Err.Number & ": " & Err.Description
End Sub

Private Function U2_SummarySubtitle() As String
    Dim s As String
    Dim opt As Object
    Dim rn As Long
    s = "기준 " & Format$(Now, "yyyy-mm-dd hh:nn") & "   |   읽은 파일 " & gU2LastFiles & "개"
    If gU2IncN > 0 Then s = s & "   |   세입 " & Format$(gU2IncN, "#,##0") & "건"
    If gU2ExpN > 0 Then s = s & "   |   세출 " & Format$(gU2ExpN, "#,##0") & "건"
    Set opt = U2_ReadOptions()
    rn = U2_EffectiveRound(opt)
    s = s & "   |   금액 단위: " & CStr(opt("Unit"))
    s = s & "   |   요약 KPI는 " & IIf(rn = 0, "본예산", CStr(rn) & "차 추경") & " 기준"
    U2_SummarySubtitle = s
End Function

Private Sub U2_Kpi(ByVal ws As Worksheet, ByVal r As Long, ByVal c As Long, ByVal label As String, ByVal value As Double, ByVal color As Long)
    With ws.Range(ws.Cells(r, c), ws.Cells(r, c + 1))
        .Merge
        .Value2 = label
        .Font.Name = "맑은 고딕"
        .Font.Size = 9
        .Font.Color = C_SUB
        .Interior.Color = RGB(250, 250, 252)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = C_LINE
        .VerticalAlignment = xlBottom
        .HorizontalAlignment = xlLeft
        .IndentLevel = 1
    End With
    With ws.Range(ws.Cells(r + 1, c), ws.Cells(r + 1, c + 1))
        .Merge
        .Value2 = value
        .NumberFormat = "#,##0"
        .Font.Name = "맑은 고딕"
        .Font.Size = 16
        .Font.Bold = True
        .Font.Color = color
        .Interior.Color = RGB(250, 250, 252)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = C_LINE
        .VerticalAlignment = xlTop
        .HorizontalAlignment = xlLeft
        .IndentLevel = 1
    End With
End Sub
Private Sub U2_SectionTitle(ByVal ws As Worksheet, ByVal addr As String, ByVal text As String)
    With ws.Range(addr)
        .Merge
        .Value2 = text
        .Font.Name = "맑은 고딕"
        .Font.Size = 12
        .Font.Bold = True
        .Font.Color = C_NAVY
        .VerticalAlignment = xlCenter
    End With
End Sub

Private Sub U2_FormatBlock(ByVal ws As Worksheet, ByVal r1 As Long, ByVal c1 As Long, ByVal r2 As Long, ByVal c2 As Long)
    With ws.Range(ws.Cells(r1, c1), ws.Cells(r2, c2))
        .Borders.LineStyle = xlContinuous
        .Borders.Color = C_LINE
        .Font.Name = "맑은 고딕"
        .Font.Size = 10
        .VerticalAlignment = xlCenter
    End With
End Sub

Private Sub U2_SortPairs(ByRef keys() As String, ByRef vals() As Double, ByVal n As Long)
    Dim i As Long, j As Long
    Dim tk As String
    Dim tv As Double
    For i = 1 To n - 1
        For j = i + 1 To n
            If vals(j) > vals(i) Then
                tv = vals(i): vals(i) = vals(j): vals(j) = tv
                tk = keys(i): keys(i) = keys(j): keys(j) = tk
            End If
        Next j
    Next i
End Sub

' ------------------------------------------------------------------ 필터

Public Sub U2_ApplyFilters()
    Dim wb As Workbook
    Set wb = gU2Book
    If wb Is Nothing Then Set wb = U2_Book()
    If wb Is Nothing Then Exit Sub
    U2_BuildIncome wb
    U2_BuildExpense wb
    U2_BuildSummary wb
    U2_ShowSheet wb, U2_SH_SUMMARY
End Sub

Public Sub U2_ResetFilters()
    Dim ws As Worksheet
    Set ws = U2_GetSheet(gU2Book, U2_SH_OPTIONS)
    If ws Is Nothing Then Exit Sub
    ws.Range("B5").Value2 = "전체"
    ws.Range("B6").Value2 = "천원"
    ws.Range("B9").Value2 = "전체"
    ws.Range("B12").Value2 = "전체"
    ws.Range("B13").Value2 = "전체"
    ws.Range("B14").Value2 = "전체"
    ws.Range("B15").Value2 = "전체"
    ws.Range("B16").Value2 = "전체"
    ws.Range("B17").Value2 = "전체"
    ws.Range("B18").Value2 = "전체"
    ws.Range("B19").Value2 = "전체"
    ws.Range("B20").Value2 = "전체"
    U2_ApplyFilters
End Sub

Public Function U2_ReadOptions() As Object
    Dim d As Object
    Dim ws As Worksheet
    Set d = CreateObject("Scripting.Dictionary")
    Set ws = U2_GetSheet(gU2Book, U2_SH_OPTIONS)
    If ws Is Nothing Then
        d("Round") = "전체": d("Unit") = "천원": d("IncItem") = "전체"
        d("Dept") = "전체": d("SubBiz") = "전체": d("ItemCode") = "전체"
        d("New") = "전체": d("Promise") = "전체": d("Promo") = "전체"
        d("Carry") = "전체": d("Central") = "전체": d("Pre") = "전체"
        Set U2_ReadOptions = d
        Exit Function
    End If
    d("Round") = U2_OptText(ws.Range("B5").Value2, "전체")
    d("Unit") = U2_OptText(ws.Range("B6").Value2, "천원")
    d("IncItem") = U2_OptText(ws.Range("B9").Value2, "전체")
    d("Dept") = U2_OptText(ws.Range("B12").Value2, "전체")
    d("SubBiz") = U2_OptText(ws.Range("B13").Value2, "전체")
    d("ItemCode") = U2_OptText(ws.Range("B14").Value2, "전체")
    d("New") = U2_OptText(ws.Range("B15").Value2, "전체")
    d("Promise") = U2_OptText(ws.Range("B16").Value2, "전체")
    d("Promo") = U2_OptText(ws.Range("B17").Value2, "전체")
    d("Carry") = U2_OptText(ws.Range("B18").Value2, "전체")
    d("Central") = U2_OptText(ws.Range("B19").Value2, "전체")
    d("Pre") = U2_OptText(ws.Range("B20").Value2, "전체")
    Set U2_ReadOptions = d
End Function

Private Function U2_OptText(ByVal v As Variant, ByVal fallback As String) As String
    Dim s As String
    s = Trim$(CStr(v))
    If s = "" Then s = fallback
    U2_OptText = s
End Function

Private Function U2_IncPass(ByRef rec As U2IncomeRec, ByRef opt As Object, ByVal ignoreRound As Boolean) As Boolean
    Dim rn As Long
    If Not ignoreRound Then
        rn = U2_RoundNo(CStr(opt("Round")))
        If rn >= 0 And rec.RoundNo <> rn Then Exit Function
    End If
    If CStr(opt("Dept")) <> "전체" And rec.Dept <> CStr(opt("Dept")) Then Exit Function
    If CStr(opt("IncItem")) <> "전체" And rec.Item <> CStr(opt("IncItem")) Then Exit Function
    U2_IncPass = True
End Function

Private Function U2_ExpPass(ByRef rec As U2ExpenseRec, ByRef opt As Object) As Boolean
    Dim rn As Long
    rn = U2_RoundNo(CStr(opt("Round")))
    If rn >= 0 And rec.RoundNo <> rn Then Exit Function
    If Not U2_ExpPassNoRound(rec, opt) Then Exit Function
    U2_ExpPass = True
End Function

Private Function U2_ExpPassNoRound(ByRef rec As U2ExpenseRec, ByRef opt As Object) As Boolean
    If CStr(opt("Dept")) <> "전체" And rec.Dept <> CStr(opt("Dept")) Then Exit Function
    If CStr(opt("SubBiz")) <> "전체" And rec.SubBiz <> CStr(opt("SubBiz")) Then Exit Function
    If CStr(opt("ItemCode")) <> "전체" Then
        If rec.ItemCode & " | " & rec.CostStat <> CStr(opt("ItemCode")) Then Exit Function
    End If
    If CStr(opt("New")) <> "전체" And rec.IsNew <> CStr(opt("New")) Then Exit Function
    If CStr(opt("Promise")) <> "전체" And rec.IsPromise <> CStr(opt("Promise")) Then Exit Function
    If CStr(opt("Promo")) <> "전체" And rec.PromoBudget <> CStr(opt("Promo")) Then Exit Function
    If CStr(opt("Carry")) <> "전체" And rec.IsCarryover <> CStr(opt("Carry")) Then Exit Function
    If CStr(opt("Central")) <> "전체" And rec.IsCentral <> CStr(opt("Central")) Then Exit Function
    If CStr(opt("Pre")) <> "전체" Then
        If Not U2_PreSelected(rec, CStr(opt("Pre"))) Then Exit Function
    End If
    U2_ExpPassNoRound = True
End Function

Private Function U2_EffectiveRound(ByRef opt As Object) As Long
    Dim rn As Long
    rn = U2_RoundNo(CStr(opt("Round")))
    If rn < 0 Then rn = U2_MaxRound()
    U2_EffectiveRound = rn
End Function

Private Function U2_ExpPassSummary(ByRef rec As U2ExpenseRec, ByRef opt As Object) As Boolean
    If rec.RoundNo <> U2_EffectiveRound(opt) Then Exit Function
    If Not U2_ExpPassNoRound(rec, opt) Then Exit Function
    U2_ExpPassSummary = True
End Function

Private Function U2_IncPassSummary(ByRef rec As U2IncomeRec, ByRef opt As Object) As Boolean
    If rec.RoundNo <> U2_EffectiveRound(opt) Then Exit Function
    If CStr(opt("Dept")) <> "전체" And rec.Dept <> CStr(opt("Dept")) Then Exit Function
    If CStr(opt("IncItem")) <> "전체" And rec.Item <> CStr(opt("IncItem")) Then Exit Function
    U2_IncPassSummary = True
End Function

Private Function U2_PreSelected(ByRef rec As U2ExpenseRec, ByVal name As String) As Boolean
    Select Case name
        Case "중기재정계획": U2_PreSelected = (rec.Pre1 = "Y")
        Case "재정투자심사": U2_PreSelected = (rec.Pre2 = "Y")
        Case "지방보조금": U2_PreSelected = (rec.Pre3 = "Y")
        Case "정책연구용역심의": U2_PreSelected = (rec.Pre4 = "Y")
        Case "공유재산관리계획": U2_PreSelected = (rec.Pre5 = "Y")
        Case "정수물품취득": U2_PreSelected = (rec.Pre6 = "Y")
        Case "국외연수심의위원회": U2_PreSelected = (rec.Pre7 = "Y")
        Case "사무위탁·대행": U2_PreSelected = (rec.Pre8 = "Y")
    End Select
End Function

Private Function PreAny(ByRef rec As U2ExpenseRec) As Boolean
    PreAny = (rec.Pre1 = "Y" Or rec.Pre2 = "Y" Or rec.Pre3 = "Y" Or rec.Pre4 = "Y" Or _
              rec.Pre5 = "Y" Or rec.Pre6 = "Y" Or rec.Pre7 = "Y" Or rec.Pre8 = "Y")
End Function

Private Function U2_RoundNo(ByVal s As String) As Long
    Dim i As Long
    Dim ch As String
    Dim num As String
    If s = "전체" Or s = "" Then
        U2_RoundNo = -1
    ElseIf s = "본예산" Then
        U2_RoundNo = 0
    Else
        For i = 1 To Len(s)
            ch = Mid$(s, i, 1)
            If ch >= "0" And ch <= "9" Then num = num & ch
        Next i
        If num = "" Then U2_RoundNo = 0 Else U2_RoundNo = CLng(num)
    End If
End Function

Private Function U2_UnitValue(ByVal s As String) As Double
    Select Case s
        Case "원": U2_UnitValue = 1
        Case "백만원": U2_UnitValue = 1000000
        Case Else: U2_UnitValue = 1000
    End Select
End Function

Private Function U2_Mark(ByVal v As String) As String
    If v = "Y" Then U2_Mark = "○" Else U2_Mark = ""
End Function

Private Function U2_LineCount(ByVal s As String) As Long
    U2_LineCount = UBound(Split(s, vbCrLf)) + 1
End Function

' ------------------------------------------------------------------ 목록 만들기

Private Function U2_RoundList() As Variant
    Dim maxR As Long
    Dim a() As String
    Dim i As Long
    maxR = U2_MaxRound()
    ReDim a(0 To maxR + 1)
    a(0) = "전체"
    a(1) = "본예산"
    For i = 1 To maxR
        a(i + 1) = CStr(i) & "차 추경"
    Next i
    U2_RoundList = a
End Function

Private Function U2_UniqueIncomeItems() As Variant
    Dim d As Object
    Dim i As Long
    Dim a() As String
    Set d = CreateObject("Scripting.Dictionary")
    d("전체") = 1
    For i = 1 To gU2IncN
        If gU2Inc(i).Item <> "" Then d(gU2Inc(i).Item) = 1
    Next i
    a = U2_DictKeys(d)
    U2_UniqueIncomeItems = a
End Function

Private Function U2_UniqueExpenseDepts() As Variant
    Dim d As Object
    Dim i As Long
    Set d = CreateObject("Scripting.Dictionary")
    d("전체") = 1
    For i = 1 To gU2ExpN
        If gU2Exp(i).Dept <> "" Then d(gU2Exp(i).Dept) = 1
    Next i
    U2_UniqueExpenseDepts = U2_DictKeys(d)
End Function

Private Function U2_UniqueExpenseSubBiz() As Variant
    Dim d As Object
    Dim i As Long
    Set d = CreateObject("Scripting.Dictionary")
    d("전체") = 1
    For i = 1 To gU2ExpN
        If gU2Exp(i).SubBiz <> "" Then d(gU2Exp(i).SubBiz) = 1
    Next i
    U2_UniqueExpenseSubBiz = U2_DictKeys(d)
End Function

Private Function U2_UniqueExpenseCodes() As Variant
    Dim d As Object
    Dim i As Long
    Dim k As String
    Set d = CreateObject("Scripting.Dictionary")
    d("전체") = 1
    For i = 1 To gU2ExpN
        k = gU2Exp(i).ItemCode & " | " & gU2Exp(i).CostStat
        If Len(Trim$(k)) > 3 Then d(k) = 1
    Next i
    U2_UniqueExpenseCodes = U2_DictKeys(d)
End Function

Private Function U2_DictKeys(ByRef d As Object) As Variant
    Dim a() As String
    Dim i As Long
    Dim k As Variant
    ReDim a(0 To d.Count - 1)
    i = 0
    For Each k In d.Keys
        a(i) = CStr(k)
        i = i + 1
    Next k
    U2_DictKeys = a
End Function

Private Function U2_MaxRound() As Long
    Dim i As Long
    Dim m As Long
    For i = 1 To gU2IncN
        If gU2Inc(i).RoundNo > m Then m = gU2Inc(i).RoundNo
    Next i
    For i = 1 To gU2ExpN
        If gU2Exp(i).RoundNo > m Then m = gU2Exp(i).RoundNo
    Next i
    If m > U2_MAX_ROUND Then m = U2_MAX_ROUND
    U2_MaxRound = m
End Function

' ------------------------------------------------------------------ 공통 서식

Public Sub U2_DeleteShapes(ByVal ws As Worksheet)
    Dim i As Long
    For i = ws.Shapes.Count To 1 Step -1
        ws.Shapes(i).Delete
    Next i
End Sub

Public Function U2_ResetSheet(ByVal wb As Workbook, ByVal nm As String) As Worksheet
    Dim ws As Worksheet
    Set ws = U2_GetSheet(wb, nm, False)
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = nm
    End If
    ws.Visible = xlSheetVisible
    ws.Cells.Clear
    U2_DeleteShapes ws
    ws.Cells.Interior.Color = RGB(255, 255, 255)
    Set U2_ResetSheet = ws
End Function

Public Sub U2_Title(ByVal ws As Worksheet, ByVal title As String, ByVal subtitle As String, ByVal lastCol As Long)
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol))
        .Merge
        .Value2 = title
        .Font.Name = "맑은 고딕"
        .Font.Size = 20
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = C_NAVY
        .VerticalAlignment = xlCenter
        .IndentLevel = 1
    End With
    With ws.Range(ws.Cells(2, 1), ws.Cells(2, lastCol))
        .Merge
        .Value2 = subtitle
        .Font.Name = "맑은 고딕"
        .Font.Size = 10
        .Font.Color = C_SUB
        .VerticalAlignment = xlCenter
        .IndentLevel = 1
    End With
    ws.Rows(1).RowHeight = 32
    ws.Rows(2).RowHeight = 22
    ws.Rows(3).RowHeight = 6
End Sub

Public Sub U2_Header(ByVal ws As Worksheet, ByVal r As Long, ByVal c As Long, ByVal headers As Variant)
    Dim i As Long
    For i = LBound(headers) To UBound(headers)
        ws.Cells(r, c + i - LBound(headers)).Value2 = headers(i)
    Next i
    With ws.Range(ws.Cells(r, c), ws.Cells(r, c + UBound(headers) - LBound(headers)))
        .Font.Name = "맑은 고딕"
        .Font.Size = 9
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = C_NAVY
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(255, 255, 255)
    End With
End Sub

Public Sub U2_ShowSheet(ByVal wb As Workbook, ByVal nm As String)
    Dim ws As Worksheet
    Set ws = U2_GetSheet(wb, nm, False)
    If Not ws Is Nothing Then
        ws.Activate
        ws.Range("A1").Select
    End If
End Sub
