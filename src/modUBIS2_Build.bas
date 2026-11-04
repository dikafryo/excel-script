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

    gU2Trace = "옵션 화면"
    U2_BuildOptions wb, resetOptions
    gU2Trace = "세입 화면"
    U2_BuildIncome wb
    gU2Trace = "세출 화면"
    U2_BuildExpense wb
    gU2Trace = "요약 화면"
    U2_BuildSummary wb
    U2_OnlyFourSheets wb

    Application.Calculation = oldCalc
    Application.ScreenUpdating = oldScreen
    Application.DisplayAlerts = oldAlerts
    U2_ShowSheet wb, U2_SH_GUIDE
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

    Set ws = U2_GetSheet(wb, U2_SH_GUIDE, False)
    If ws Is Nothing Then
        Set ws = wb.Worksheets(1)
        ws.Name = U2_SH_GUIDE
    End If

    ws.Cells.UnMerge
    ws.Cells.Clear
    ws.Cells.Interior.Color = RGB(255, 255, 255)
    U2_DeleteShapes ws
    U2_BuildNav ws
    U2_Title ws, "UBIS 예산 가져오기", "UBIS에서 내려받은 세입·세출 예산요구서를 한 번에 분석합니다.", 10

    With ws.Range("A6:E13")
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

    With ws.Range("F6:J13")
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

    With ws.Range("A14:J14")
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
    ws.Rows(14).RowHeight = 26

    ws.Rows(6).RowHeight = 24
    ws.Rows(7).RowHeight = 20
    ws.Rows(8).RowHeight = 20
    ws.Rows(9).RowHeight = 20
    ws.Rows(10).RowHeight = 20
    ws.Rows(11).RowHeight = 20
    ws.Rows(12).RowHeight = 20
    ws.Rows(13).RowHeight = 20

    U2_AddButton ws, "u2_pick", "파일 선택하고 분석하기", "pick", 20, 300, 42, True
    U2_AddButton ws, "u2_optguide", "옵션 설정", "options", 260, 300, 42, False
    U2_AddButton ws, "u2_helpguide", "사용 방법", "help", 380, 300, 42, False

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
    ws.Rows(5).RowHeight = 8
    ws.Rows(15).RowHeight = 8
    ws.Rows(16).RowHeight = 8
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

    U2_BuildNav ws
    U2_Title ws, "옵션 설정", "필터를 선택한 뒤 [필터 적용]을 누르세요.", 8

    U2_OptionLabel ws, "A6", "공통"
    U2_OptionLabel ws, "A7", "차수"
    U2_OptionLabel ws, "A8", "금액 단위"
    U2_OptionLabel ws, "A10", "세입"
    U2_OptionLabel ws, "A11", "사업항목"
    U2_OptionLabel ws, "A13", "세출"
    U2_OptionLabel ws, "A14", "부서"
    U2_OptionLabel ws, "A15", "세부사업"
    U2_OptionLabel ws, "A16", "세목"
    U2_OptionLabel ws, "A17", "신규여부"
    U2_OptionLabel ws, "A18", "공약사업"
    U2_OptionLabel ws, "A19", "홍보예산"
    U2_OptionLabel ws, "A20", "명시이월"
    U2_OptionLabel ws, "A21", "본청주관"
    U2_OptionLabel ws, "A22", "사전절차"

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

    U2_SetValidation ws, "B7", "=$M$2:$M$40"
    U2_SetValidation ws, "B8", "=$N$2:$N$10"
    U2_SetValidation ws, "B11", "=$O$2:$O$300"
    U2_SetValidation ws, "B14", "=$P$2:$P$300"
    U2_SetValidation ws, "B15", "=$Q$2:$Q$300"
    U2_SetValidation ws, "B16", "=$R$2:$R$300"
    U2_SetValidation ws, "B17", "=$S$2:$S$10"
    U2_SetValidation ws, "B18", "=$T$2:$T$10"
    U2_SetValidation ws, "B19", "=$U$2:$U$10"
    U2_SetValidation ws, "B20", "=$V$2:$V$10"
    U2_SetValidation ws, "B21", "=$W$2:$W$10"
    U2_SetValidation ws, "B22", "=$X$2:$X$20"

    If reset Or old Is Nothing Then
        ws.Range("B7").Value2 = "전체"
        ws.Range("B8").Value2 = "천원"
        ws.Range("B11").Value2 = "전체"
        ws.Range("B14").Value2 = "전체"
        ws.Range("B15").Value2 = "전체"
        ws.Range("B16").Value2 = "전체"
        ws.Range("B17").Value2 = "전체"
        ws.Range("B18").Value2 = "전체"
        ws.Range("B19").Value2 = "전체"
        ws.Range("B20").Value2 = "전체"
        ws.Range("B21").Value2 = "전체"
        ws.Range("B22").Value2 = "전체"
    Else
        U2_RestoreOption ws, "B7", old("Round")
        U2_RestoreOption ws, "B8", old("Unit")
        U2_RestoreOption ws, "B11", old("IncItem")
        U2_RestoreOption ws, "B14", old("Dept")
        U2_RestoreOption ws, "B15", old("SubBiz")
        U2_RestoreOption ws, "B16", old("ItemCode")
        U2_RestoreOption ws, "B17", old("New")
        U2_RestoreOption ws, "B18", old("Promise")
        U2_RestoreOption ws, "B19", old("Promo")
        U2_RestoreOption ws, "B20", old("Carry")
        U2_RestoreOption ws, "B21", old("Central")
        U2_RestoreOption ws, "B22", old("Pre")
    End If

    With ws.Range("B7:B22")
        .Font.Name = "맑은 고딕"
        .Font.Size = 11
        .Interior.Color = RGB(250, 250, 252)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = C_LINE
        .HorizontalAlignment = xlLeft
        .IndentLevel = 1
    End With

    U2_AddButton ws, "u2_apply", "필터 적용", "apply", 30, 520, 34, True
    U2_AddButton ws, "u2_reset", "초기화", "reset", 150, 520, 34, False
    U2_AddButton ws, "u2_sum", "요약 보기", "summary", 270, 520, 34, False
    U2_AddButton ws, "u2_inc", "세입 보기", "income", 390, 520, 34, False
    U2_AddButton ws, "u2_exp", "세출 보기", "expense", 510, 520, 34, False

    ws.Columns("A").ColumnWidth = 16
    ws.Columns("B").ColumnWidth = 34
    ws.Columns("C:L").ColumnWidth = 3
    ws.Columns("M:X").Hidden = True

    ws.Range("A24:H24").Merge
    ws.Range("A24").Value2 = "필터를 바꾼 뒤 [필터 적용]을 누르면 세입·세출·요약 화면이 다시 만들어집니다."
    ws.Range("A24").Font.Name = "맑은 고딕"
    ws.Range("A24").Font.Size = 10
    ws.Range("A24").Font.Color = C_SUB
    ws.Range("A24").VerticalAlignment = xlCenter
    ws.Range("A24").IndentLevel = 1
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
    U2_BuildNav ws
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
        ws.Range("A6").Value2 = "선택한 조건에 맞는 세입 자료가 없습니다."
        Exit Sub
    End If

    ReDim keys(1 To items.Count)
    n = 0
    For Each key In items.Keys
        n = n + 1
        keys(n) = CStr(key)
    Next key

    U2_Header ws, 6, 1, U2_IncomeSummaryHeaders(maxR)

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

    ws.Range("A7").Resize(n, lastCol).Value2 = outArr
    ws.Range(ws.Cells(7, 1), ws.Cells(6 + n, lastCol)).Sort _
        Key1:=ws.Cells(7, lastCol - 1), Order1:=xlDescending, Header:=xlNo

    r = 7 + n
    With ws.Range(ws.Cells(7, 1), ws.Cells(r - 1, 3 + maxR + 2))
        .Borders.LineStyle = xlContinuous
        .Borders.Color = C_LINE
        .Font.Name = "맑은 고딕"
        .Font.Size = 10
    End With
    ws.Range(ws.Cells(7, 3), ws.Cells(r - 1, 3 + maxR + 2)).NumberFormat = "#,##0"
    ws.Range(ws.Cells(7, 3 + maxR + 2), ws.Cells(r - 1, 3 + maxR + 2)).Font.Color = C_BLUE

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
    ws.Rows(6).RowHeight = 30
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
    Dim mgrLast As Long

    Set ws = U2_ResetSheet(wb, U2_SH_EXPENSE)
    Set opt = U2_ReadOptions()
    unit = U2_UnitValue(opt("Unit"))

    U2_BuildNav ws
    U2_Title ws, "세출예산", "담당자와 사업을 고르면 아래에 세부 내역이 바로 표시됩니다. (금액 단위: " & CStr(opt("Unit")) & ")", 23

    With ws.Range("A6")
        .Value2 = "담당자"
        .Font.Name = "맑은 고딕"
        .Font.Size = 10
        .Font.Bold = True
        .Font.Color = C_TEXT
        .VerticalAlignment = xlCenter
    End With
    With ws.Range("E6")
        .Value2 = "사업 목록 (세부사업 / 세세부사업)"
        .Font.Name = "맑은 고딕"
        .Font.Size = 10
        .Font.Bold = True
        .Font.Color = C_TEXT
        .VerticalAlignment = xlCenter
    End With

    ws.Range("AY1:BG500").Clear
    mgrLast = U2_FillManagerList(ws)
    U2_CreateExpenseControls ws, mgrLast
    U2_RebuildExpenseList ws
    U2_BuildExpenseDetail ws

    ws.Columns("A").ColumnWidth = 18
    ws.Columns("B").ColumnWidth = 18
    ws.Columns("C:F").ColumnWidth = 14
    ws.Columns("G").ColumnWidth = 9
    ws.Columns("H").ColumnWidth = 12
    ws.Columns("I:O").ColumnWidth = 10
    ws.Columns("P").ColumnWidth = 9
    ws.Columns("Q:U").ColumnWidth = 6
    ws.Columns("V").ColumnWidth = 14
    ws.Columns("W").ColumnWidth = 20
    ws.Columns("AY:BG").Hidden = True
    ws.Rows(6).RowHeight = 18
    ws.Rows(7).RowHeight = 18
    ws.Rows(8).RowHeight = 18
    ws.Rows(9).RowHeight = 18
    ws.Rows(10).RowHeight = 18
    ws.Rows(11).RowHeight = 18
    ws.Rows(12).RowHeight = 18
    ws.Rows(13).RowHeight = 18

    On Error Resume Next
    ActiveWindow.Zoom = 75
    On Error GoTo 0
    With ws.PageSetup
        .Orientation = xlLandscape
        .PaperSize = xlPaperA4
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .PrintArea = "$A$1:$W$" & (14 + 20)
    End With
End Sub

Private Function U2_FillManagerList(ByVal ws As Worksheet) As Long
    Dim d As Object
    Dim i As Long
    Dim r As Long
    Dim k As Variant
    Dim opt As Object
    Dim managers() As String
    Dim n As Long
    Dim n1 As Long
    Dim n2 As Long

    Set opt = U2_ReadOptions()
    Set d = CreateObject("Scripting.Dictionary")
    d("전체") = 1
    For i = 1 To gU2ExpN
        If U2_ExpPassSummary(gU2Exp(i), opt) Then
            If gU2Exp(i).Manager <> "" Then d(gU2Exp(i).Manager) = 1
        End If
    Next i

    n = d.Count
    ReDim managers(1 To n)
    i = 0
    For Each k In d.Keys
        i = i + 1
        managers(i) = CStr(k)
    Next k

    ws.Range("AY2:AZ500").ClearContents
    n1 = (n + 1) \ 2
    n2 = n - n1
    For i = 1 To n1
        ws.Cells(1 + i, "AY").Value2 = managers(i)
    Next i
    For i = 1 To n2
        ws.Cells(1 + i, "AZ").Value2 = managers(n1 + i)
    Next i
    ws.Range("BF1").Value2 = 1 + n1
    ws.Range("BG1").Value2 = 1 + n2
    If n2 = 0 Then ws.Range("BG1").Value2 = 1
    U2_FillManagerList = 1 + Application.Max(n1, n2)
End Function

Private Sub U2_CreateExpenseControls(ByVal ws As Worksheet, ByVal mgrLast As Long)
    Dim shp As Shape
    Dim last1 As Long
    Dim last2 As Long

    On Error Resume Next
    ws.Shapes("u2lstManager1").Delete
    ws.Shapes("u2lstManager2").Delete
    ws.Shapes("u2lstSub").Delete
    On Error GoTo 0

    last1 = CLng(U2_Num(ws.Range("BF1").Value2))
    last2 = CLng(U2_Num(ws.Range("BG1").Value2))
    If last1 < 2 Then last1 = 2
    If last2 < 2 Then last2 = 2

    Set shp = ws.Shapes.AddFormControl(xlListBox, 8, 125, 128, 90)
    shp.Name = "u2lstManager1"
    shp.Placement = xlMove
    shp.ControlFormat.ListFillRange = "$AY$2:$AY$" & last1
    shp.ControlFormat.LinkedCell = "$BA$1"
    shp.ControlFormat.ListIndex = 1
    shp.OnAction = "U2_Manager1Change"
    shp.ZOrder msoBringToFront

    Set shp = ws.Shapes.AddFormControl(xlListBox, 140, 125, 128, 90)
    shp.Name = "u2lstManager2"
    shp.Placement = xlMove
    shp.ControlFormat.ListFillRange = "$AZ$2:$AZ$" & last2
    shp.ControlFormat.LinkedCell = "$BB$1"
    On Error Resume Next
    shp.ControlFormat.ListIndex = -1
    On Error GoTo 0
    shp.OnAction = "U2_Manager2Change"
    shp.ZOrder msoBringToFront

    Set shp = ws.Shapes.AddFormControl(xlListBox, 290, 125, 320, 90)
    shp.Name = "u2lstSub"
    shp.Placement = xlMove
    shp.ControlFormat.LinkedCell = "$BE$1"
    shp.OnAction = "U2_SubListClick"
    shp.ZOrder msoBringToFront
End Sub

Public Sub U2_RebuildExpenseList(Optional ByVal ws As Worksheet = Nothing)
    Dim d As Object
    Dim i As Long
    Dim r As Long
    Dim k As Variant
    Dim mgr As String
    Dim opt As Object

    If ws Is Nothing Then Set ws = U2_GetSheet(U2_CurrentBook(), U2_SH_EXPENSE)
    If ws Is Nothing Then Exit Sub
    U2_EnsureDataLoaded
    Set opt = U2_ReadOptions()
    mgr = U2_SelectedManager(ws)

    Set d = CreateObject("Scripting.Dictionary")
    d("ALL") = "전체"
    For i = 1 To gU2ExpN
        If U2_ExpPassSummary(gU2Exp(i), opt) Then
            If mgr = "" Or mgr = "전체" Or gU2Exp(i).Manager = mgr Then
                If gU2Exp(i).SubBiz <> "" Then d("S|" & gU2Exp(i).SubBiz) = "▸ " & gU2Exp(i).SubBiz
                If gU2Exp(i).SubSubBiz <> "" Then d("SS|" & gU2Exp(i).SubSubBiz) = "    - " & gU2Exp(i).SubSubBiz
            End If
        End If
    Next i

    ws.Range("BC2:BD500").ClearContents
    r = 2
    For Each k In d.Keys
        ws.Cells(r, "BC").Value2 = CStr(d(k))
        ws.Cells(r, "BD").Value2 = CStr(k)
        r = r + 1
    Next k
    If r = 2 Then
        ws.Cells(2, "BC").Value2 = "전체"
        ws.Cells(2, "BD").Value2 = "ALL"
        r = 3
    End If

    On Error Resume Next
    ws.Shapes("u2lstSub").OnAction = ""
    ws.Shapes("u2lstSub").ControlFormat.ListFillRange = "$BC$2:$BC$" & (r - 1)
    ws.Shapes("u2lstSub").ControlFormat.ListIndex = 1
    ws.Shapes("u2lstSub").OnAction = "U2_SubListClick"
    On Error GoTo 0
End Sub

Public Sub U2_Manager1Change()
    Dim ws As Worksheet
    U2_EnsureDataLoaded
    Set ws = U2_GetSheet(U2_CurrentBook(), U2_SH_EXPENSE)
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    ws.Shapes("u2lstManager2").ControlFormat.ListIndex = -1
    ws.Range("BB1").Value2 = 0
    On Error GoTo 0
    U2_RebuildExpenseList ws
    U2_BuildExpenseDetail ws
End Sub

Public Sub U2_Manager2Change()
    Dim ws As Worksheet
    U2_EnsureDataLoaded
    Set ws = U2_GetSheet(U2_CurrentBook(), U2_SH_EXPENSE)
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    ws.Shapes("u2lstManager1").ControlFormat.ListIndex = -1
    ws.Range("BA1").Value2 = 0
    On Error GoTo 0
    U2_RebuildExpenseList ws
    U2_BuildExpenseDetail ws
End Sub

Public Sub U2_SubListClick()
    U2_EnsureDataLoaded
    U2_BuildExpenseDetail
End Sub

Public Sub U2_BuildExpenseDetail(Optional ByVal ws As Worksheet = Nothing)
    Dim opt As Object
    Dim unit As Double
    Dim mgr As String
    Dim sel As String
    Dim key As String
    Dim i As Long
    Dim n As Long
    Dim r As Long
    Dim total As Double
    Dim arr As Variant
    Dim headers As Variant
    Dim unitLabel As String
    Dim rn As Long
    Dim prevLast As Long

    If ws Is Nothing Then Set ws = U2_GetSheet(U2_CurrentBook(), U2_SH_EXPENSE)
    If ws Is Nothing Then Exit Sub
    U2_EnsureDataLoaded
    Set opt = U2_ReadOptions()
    unit = U2_UnitValue(opt("Unit"))
    unitLabel = CStr(opt("Unit"))
    mgr = U2_SelectedManager(ws)
    sel = U2_SelectedItem(ws)
    If sel = "" Then sel = "전체"
    key = U2_FindListKey(ws, sel)
    rn = U2_EffectiveRound(opt)

    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    prevLast = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If prevLast < 15 Then prevLast = 15
    ws.Range("A14:W" & prevLast).Clear

    For i = 1 To gU2ExpN
        If U2_ExpPassSummary(gU2Exp(i), opt) Then
            If mgr = "" Or mgr = "전체" Or gU2Exp(i).Manager = mgr Then
                If U2_ItemMatch(gU2Exp(i), key) Then
                    n = n + 1
                    total = total + gU2Exp(i).ReqAmt
                End If
            End If
        End If
    Next i

    With ws.Range("A5:W5")
        .Merge
        .Value2 = "담당자: " & IIf(mgr = "", "전체", mgr) & _
                  "   |   사업: " & sel & _
                  "   |   차수: " & IIf(rn = 0, "본예산", CStr(rn) & "차 추경") & _
                  "   |   합계 " & Format$(total / unit, "#,##0") & " " & unitLabel & _
                  "   |   " & Format$(n, "#,##0") & "건"
        .Font.Name = "맑은 고딕"
        .Font.Size = 10
        .Font.Bold = True
        .Font.Color = C_NAVY
        .Interior.Color = RGB(240, 247, 255)
        .VerticalAlignment = xlCenter
        .IndentLevel = 1
    End With

    headers = Array("세부사업", "세세부사업", "산출1", "산출2", "산출3", "산출4", "세목코드", "원가통계목", _
                    "예산요구액", "재원_특교", "재원_국고", "재원_비법정", "재원_기타", "재원_자체", "재원_합계", _
                    "업무담당자", "신규", "공약", "홍보", "명시이월", "본청주관", "사전절차", "설명사항")
    U2_Header ws, 14, 1, headers

    If n > 0 Then
        ReDim arr(1 To n, 1 To 23)
        r = 0
        For i = 1 To gU2ExpN
            If U2_ExpPassSummary(gU2Exp(i), opt) Then
                If mgr = "" Or mgr = "전체" Or gU2Exp(i).Manager = mgr Then
                    If U2_ItemMatch(gU2Exp(i), key) Then
                        r = r + 1
                        arr(r, 1) = gU2Exp(i).SubBiz
                        arr(r, 2) = gU2Exp(i).SubSubBiz
                        arr(r, 3) = U2_LabelText(gU2Exp(i).Calc1)
                        arr(r, 4) = U2_LabelText(gU2Exp(i).Calc2)
                        arr(r, 5) = U2_LabelText(gU2Exp(i).Calc3)
                        arr(r, 6) = U2_LabelText(gU2Exp(i).Calc4)
                        arr(r, 7) = gU2Exp(i).ItemCode
                        arr(r, 8) = gU2Exp(i).CostStat
                        arr(r, 9) = gU2Exp(i).ReqAmt / unit
                        arr(r, 10) = gU2Exp(i).FundSpecial / unit
                        arr(r, 11) = gU2Exp(i).FundNational / unit
                        arr(r, 12) = gU2Exp(i).FundNonStatutory / unit
                        arr(r, 13) = gU2Exp(i).FundOther / unit
                        arr(r, 14) = gU2Exp(i).FundSelf / unit
                        arr(r, 15) = gU2Exp(i).FundTotal / unit
                        arr(r, 16) = gU2Exp(i).Manager
                        arr(r, 17) = U2_Mark(gU2Exp(i).IsNew)
                        arr(r, 18) = U2_Mark(gU2Exp(i).IsPromise)
                        arr(r, 19) = U2_Mark(gU2Exp(i).PromoBudget)
                        arr(r, 20) = U2_Mark(gU2Exp(i).IsCarryover)
                        arr(r, 21) = U2_Mark(gU2Exp(i).IsCentral)
                        arr(r, 22) = gU2Exp(i).PreSummary
                        arr(r, 23) = gU2Exp(i).Desc
                    End If
                End If
            End If
        Next i
        ws.Range("A15").Resize(n, 23).Value2 = arr
    Else
        ws.Range("A15").Value2 = "선택한 조건에 맞는 세출 자료가 없습니다."
    End If

    With ws.Range("A14:W" & (14 + Application.Max(1, n)))
        .Borders.LineStyle = xlContinuous
        .Borders.Color = C_LINE
        .Font.Name = "맑은 고딕"
        .Font.Size = 9
        .VerticalAlignment = xlCenter
    End With
    ws.Range("A14:W14").Interior.Color = C_NAVY
    ws.Range("A14:W14").Font.Color = RGB(255, 255, 255)
    ws.Range("A14:W14").Font.Bold = True
    ws.Range("A14:W14").WrapText = True
    ws.Rows(14).RowHeight = 32

    If n > 0 Then
        ws.Range("I15:O" & (14 + n)).NumberFormat = "#,##0"
        ws.Range("Q15:Q" & (14 + n)).Interior.Color = RGB(232, 245, 233)
        ws.Range("R15:R" & (14 + n)).Interior.Color = RGB(255, 248, 225)
        ws.Range("S15:S" & (14 + n)).Interior.Color = RGB(255, 248, 225)
        ws.Range("T15:T" & (14 + n)).Interior.Color = RGB(227, 242, 253)
        ws.Range("U15:U" & (14 + n)).Interior.Color = RGB(243, 229, 245)
        ws.Range("A14:W" & (14 + n)).AutoFilter
    End If

    ws.Activate
    On Error Resume Next
    ActiveWindow.FreezePanes = False
    ws.Range("A15").Select
    ActiveWindow.FreezePanes = True
    ActiveWindow.Zoom = 75
    On Error GoTo 0
End Sub

Private Function U2_FindListKey(ByVal ws As Worksheet, ByVal sel As String) As String
    Dim r As Long
    Dim lastR As Long
    lastR = ws.Cells(ws.Rows.Count, "BC").End(xlUp).Row
    For r = 2 To lastR
        If U2_CellText(ws.Cells(r, "BC").Value2) = sel Then
            U2_FindListKey = U2_CellText(ws.Cells(r, "BD").Value2)
            Exit Function
        End If
    Next r
    U2_FindListKey = "ALL"
End Function

Private Function U2_SelectedManager(ByVal ws As Worksheet) As String
    Dim idx As Long
    idx = CLng(U2_Num(ws.Range("BA1").Value2))
    If idx >= 1 Then
        U2_SelectedManager = U2_CellText(ws.Cells(1 + idx, "AY").Value2)
        If U2_SelectedManager <> "" Then Exit Function
    End If
    idx = CLng(U2_Num(ws.Range("BB1").Value2))
    If idx >= 1 Then
        U2_SelectedManager = U2_CellText(ws.Cells(1 + idx, "AZ").Value2)
        If U2_SelectedManager <> "" Then Exit Function
    End If
    U2_SelectedManager = "전체"
End Function

Private Function U2_SelectedItem(ByVal ws As Worksheet) As String
    Dim idx As Long
    idx = CLng(U2_Num(ws.Range("BE1").Value2))
    If idx >= 1 Then U2_SelectedItem = U2_CellText(ws.Cells(1 + idx, "BC").Value2)
    If U2_SelectedItem = "" Then U2_SelectedItem = "전체"
End Function

Private Function U2_ItemMatch(ByRef rec As U2ExpenseRec, ByVal key As String) As Boolean
    If key = "" Or key = "ALL" Then
        U2_ItemMatch = True
    ElseIf Left$(key, 2) = "S|" Then
        U2_ItemMatch = (rec.SubBiz = Mid$(key, 3))
    ElseIf Left$(key, 3) = "SS|" Then
        U2_ItemMatch = (rec.SubSubBiz = Mid$(key, 4))
    Else
        U2_ItemMatch = True
    End If
End Function

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
    Set ws = U2_ResetSheet(wb, U2_SH_GUIDE)
    U2_BuildNav ws
    Set opt = U2_ReadOptions()
    unit = U2_UnitValue(opt("Unit"))
    maxR = U2_MaxRound()

    U2_Title ws, "UBIS 예산 관리", U2_SummarySubtitle(), 17
    ws.Rows(6).RowHeight = 20
    ws.Rows(7).RowHeight = 30

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

    U2_Kpi ws, 6, 2, "세입 최종예산", incTotal / unit, C_BLUE
    U2_Kpi ws, 6, 4, "세출 예산요구액", expTotal / unit, C_NAVY
    U2_Kpi ws, 6, 6, "세입 - 세출", (incTotal - expTotal) / unit, IIf(incTotal >= expTotal, C_GREEN, RGB(200, 60, 60))
    U2_Kpi ws, 6, 8, "신규사업", newAmt / unit, C_GREEN
    U2_Kpi ws, 6, 10, "공약사업", proAmt / unit, C_ORANGE
    U2_Kpi ws, 6, 12, "홍보예산", promoAmt / unit, RGB(180, 120, 0)
    U2_Kpi ws, 6, 14, "명시이월", carryAmt / unit, C_BLUE
    U2_Kpi ws, 6, 16, "본청주관", centralAmt / unit, C_PURPLE

    ' 차수별 요약
    ws.Cells(11, 2).Value2 = "차수별 요약"
    U2_SectionTitle ws, "B11:H11", "차수별 요약"
    U2_Header ws, 12, 2, Array("차수", "세입", "세출", "차이", "세출건수")
    r = 13
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
    U2_FormatBlock ws, 12, 2, r - 1, 6
    ws.Range(ws.Cells(13, 3), ws.Cells(r - 1, 5)).NumberFormat = "#,##0"

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
    U2_SectionTitle ws, "J11:P11", "세부사업 TOP 10"
    U2_Header ws, 12, 10, Array("순위", "세부사업", "예산요구액", "비중")
    r = 13
    For Each k In dSub.Keys
        ws.Cells(r, 10).Value2 = 0
        ws.Cells(r, 11).Value2 = CStr(k)
        ws.Cells(r, 12).Value2 = U2_Num(dSub(k)) / unit
        If expTotal > 0 Then ws.Cells(r, 13).Value2 = U2_Num(dSub(k)) / expTotal
        r = r + 1
    Next k
    If r > 13 Then
        ws.Range(ws.Cells(13, 10), ws.Cells(r - 1, 13)).Sort _
            Key1:=ws.Cells(13, 12), Order1:=xlDescending, Header:=xlNo
        For i = 13 To r - 1
            ws.Cells(i, 10).Value2 = i - 12
        Next i
    End If
    U2_FormatBlock ws, 12, 10, r - 1, 13
    ws.Range(ws.Cells(13, 12), ws.Cells(r - 1, 12)).NumberFormat = "#,##0"
    ws.Range(ws.Cells(13, 13), ws.Cells(r - 1, 13)).NumberFormat = "0.0%"
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
    U2_SectionTitle ws, "B26:H26", "재원별"
    U2_Header ws, 27, 2, Array("재원", "금액", "비중")
    r = 28
    For Each k In dFund.Keys
        ws.Cells(r, 2).Value2 = k
        ws.Cells(r, 3).Value2 = U2_Num(dFund(k)) / unit
        If expTotal > 0 Then ws.Cells(r, 4).Value2 = U2_Num(dFund(k)) / expTotal
        r = r + 1
    Next k
    U2_FormatBlock ws, 27, 2, r - 1, 4
    ws.Range(ws.Cells(28, 3), ws.Cells(r - 1, 3)).NumberFormat = "#,##0"
    ws.Range(ws.Cells(28, 4), ws.Cells(r - 1, 4)).NumberFormat = "0.0%"

    gU2Trace = "flags"
    ' 사업 유형별
    U2_SectionTitle ws, "J26:P26", "사업 유형별"
    U2_Header ws, 27, 10, Array("구분", "건수", "금액")
    ws.Cells(32, 10).Value2 = "신규사업": ws.Cells(26, 11).Value2 = newCnt: ws.Cells(26, 12).Value2 = newAmt / unit
    ws.Cells(33, 10).Value2 = "공약사업": ws.Cells(27, 11).Value2 = proCnt: ws.Cells(27, 12).Value2 = proAmt / unit
    ws.Cells(32, 10).Value2 = "홍보예산": ws.Cells(28, 11).Value2 = promoCnt: ws.Cells(28, 12).Value2 = promoAmt / unit
    ws.Cells(33, 10).Value2 = "명시이월": ws.Cells(29, 11).Value2 = carryCnt: ws.Cells(29, 12).Value2 = carryAmt / unit
    ws.Cells(32, 10).Value2 = "본청주관": ws.Cells(30, 11).Value2 = centralCnt: ws.Cells(30, 12).Value2 = centralAmt / unit
    ws.Cells(33, 10).Value2 = "사전절차": ws.Cells(31, 11).Value2 = preCnt: ws.Cells(31, 12).Value2 = preAmt / unit
    U2_FormatBlock ws, 27, 10, 33, 12
    ws.Range("L28:L33").NumberFormat = "#,##0"

    gU2Trace = "buttons"
    ws.Columns("A").ColumnWidth = 3
    ws.Columns("B:H").ColumnWidth = 15
    ws.Columns("I").ColumnWidth = 3
    ws.Columns("J:P").ColumnWidth = 15
    ws.Rows(5).RowHeight = 6
    ws.Rows(10).RowHeight = 8
    ws.Rows(25).RowHeight = 8
    ws.Range("B36:P36").Merge
    ws.Range("B36").Value2 = "담당자별 상세는 상단 [세출예산] 버튼에서 확인하세요."
    ws.Range("B36").Font.Name = "맑은 고딕"
    ws.Range("B36").Font.Size = 10
    ws.Range("B36").Font.Color = C_SUB
    ws.Range("B36").VerticalAlignment = xlCenter
    ws.Range("B36").IndentLevel = 1
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
    U2_EnsureDataLoaded
    Set wb = gU2Book
    If wb Is Nothing Then Set wb = U2_Book()
    If wb Is Nothing Then Exit Sub
    U2_BuildIncome wb
    U2_BuildExpense wb
    U2_BuildSummary wb
    U2_OnlyFourSheets wb
    U2_ShowSheet wb, U2_SH_GUIDE
End Sub

Public Sub U2_ResetFilters()
    Dim ws As Worksheet
    Set ws = U2_GetSheet(gU2Book, U2_SH_OPTIONS)
    If ws Is Nothing Then Exit Sub
    ws.Range("B7").Value2 = "전체"
    ws.Range("B8").Value2 = "천원"
    ws.Range("B11").Value2 = "전체"
    ws.Range("B14").Value2 = "전체"
    ws.Range("B15").Value2 = "전체"
    ws.Range("B16").Value2 = "전체"
    ws.Range("B17").Value2 = "전체"
    ws.Range("B18").Value2 = "전체"
    ws.Range("B19").Value2 = "전체"
    ws.Range("B20").Value2 = "전체"
    ws.Range("B21").Value2 = "전체"
    ws.Range("B22").Value2 = "전체"
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
    d("Round") = U2_OptText(ws.Range("B7").Value2, "전체")
    d("Unit") = U2_OptText(ws.Range("B8").Value2, "천원")
    d("IncItem") = U2_OptText(ws.Range("B11").Value2, "전체")
    d("Dept") = U2_OptText(ws.Range("B14").Value2, "전체")
    d("SubBiz") = U2_OptText(ws.Range("B15").Value2, "전체")
    d("ItemCode") = U2_OptText(ws.Range("B16").Value2, "전체")
    d("New") = U2_OptText(ws.Range("B17").Value2, "전체")
    d("Promise") = U2_OptText(ws.Range("B18").Value2, "전체")
    d("Promo") = U2_OptText(ws.Range("B19").Value2, "전체")
    d("Carry") = U2_OptText(ws.Range("B20").Value2, "전체")
    d("Central") = U2_OptText(ws.Range("B21").Value2, "전체")
    d("Pre") = U2_OptText(ws.Range("B22").Value2, "전체")
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

Public Function U2_CurrentBook() As Workbook
    If gU2Book Is Nothing Then
        On Error Resume Next
        Set gU2Book = ActiveWorkbook
        On Error GoTo 0
    End If
    Set U2_CurrentBook = gU2Book
End Function

Public Sub U2_OnlyFourSheets(ByVal wb As Workbook)
    Dim ws As Worksheet
    If U2_HasSheet(wb, U2_SH_GUIDE) Then wb.Worksheets(U2_SH_GUIDE).Activate
    For Each ws In wb.Worksheets
        Select Case ws.Name
            Case U2_SH_GUIDE, U2_SH_OPTIONS, U2_SH_INCOME, U2_SH_EXPENSE
                ws.Visible = xlSheetVisible
            Case Else
                ws.Visible = xlSheetVeryHidden
        End Select
    Next ws
End Sub

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
    ws.Cells.UnMerge
    ws.Cells.Clear
    U2_DeleteShapes ws
    ws.Cells.Interior.Color = RGB(255, 255, 255)
    Set U2_ResetSheet = ws
End Function

Public Sub U2_Title(ByVal ws As Worksheet, ByVal title As String, ByVal subtitle As String, ByVal lastCol As Long)
    With ws.Range(ws.Cells(3, 1), ws.Cells(3, lastCol))
        .Merge
        .Value2 = title
        .Font.Name = "맑은 고딕"
        .Font.Size = 18
        .Font.Bold = True
        .Font.Color = C_NAVY
        .VerticalAlignment = xlCenter
        .IndentLevel = 1
    End With
    With ws.Range(ws.Cells(4, 1), ws.Cells(4, lastCol))
        .Merge
        .Value2 = subtitle
        .Font.Name = "맑은 고딕"
        .Font.Size = 10
        .Font.Color = C_SUB
        .VerticalAlignment = xlCenter
        .IndentLevel = 1
    End With
    ws.Rows(2).RowHeight = 6
    ws.Rows(3).RowHeight = 30
    ws.Rows(4).RowHeight = 22
End Sub

Public Sub U2_BuildNav(ByVal ws As Worksheet)
    Dim captions As Variant
    Dim procs As Variant
    Dim i As Long
    Dim leftPos As Single

    captions = Array("시작", "옵션", "세입예산", "세출예산", "파일 다시 선택", "사용 방법", "인쇄")
    procs = Array("start", "options", "income", "expense", "pick", "help", "print")

    ws.Range("A1:AG1").Clear
    With ws.Range("A1:AG1")
        .Interior.Color = C_NAVY
        .RowHeight = 34
    End With

    leftPos = 8
    For i = 0 To UBound(captions)
        U2_AddButton ws, "u2nav_" & CStr(procs(i)), CStr(captions(i)), CStr(procs(i)), leftPos, 4, 26, False
        ws.Shapes("u2nav_" & CStr(procs(i))).Placement = xlMove
        ws.Shapes("u2nav_" & CStr(procs(i))).ZOrder msoBringToFront
        leftPos = leftPos + Application.Max(90, Len(CStr(captions(i))) * 13) + 8
    Next i

    ws.Activate
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True
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
        If nm = U2_SH_EXPENSE Then
            On Error Resume Next
            ActiveWindow.Zoom = 75
            On Error GoTo 0
        End If
    End If
End Sub
