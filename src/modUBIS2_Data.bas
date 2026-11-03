Attribute VB_Name = "modUBIS2_Data"
Option Explicit

' =====================================================================================
' modUBIS2_Data : UBIS 세입·세출 예산요구서 읽기와 원본 데이터 보관
'   - 여러 파일을 한 번에 읽어 세입/세출을 구분하고 연도·차수를 자동 인식한다.
'   - 금액은 모두 '원' 단위로 저장한다. (UBIS 원본은 보통 천원 단위)
'   - 원본 데이터는 숨김 시트(_세입원본, _세출원본)에 보관하고 화면은 그 데이터로 다시 그린다.
' =====================================================================================

Public Const U2_SH_GUIDE As String = "시작"
Public Const U2_SH_SUMMARY As String = "요약"
Public Const U2_SH_INCOME As String = "세입"
Public Const U2_SH_EXPENSE As String = "세출"
Public Const U2_SH_EXP_SUMMARY As String = "세출요약"
Public Const U2_SH_OPTIONS As String = "옵션"
Public Const U2_SH_INC_RAW As String = "_세입원본"
Public Const U2_SH_EXP_RAW As String = "_세출원본"
Public Const U2_SH_LISTS As String = "_목록"

Public Const U2_MAX_ROUND As Long = 20

Public Type U2IncomeRec
    Yr As Long
    RoundNo As Long
    RoundName As String
    Jang As String
    Gwan As String
    Hang As String
    Mok As String
    Item As String
    Dept As String
    PrevAmt As Double
    ReqAmt As Double
    FinalAmt As Double
    ChangeAmt As Double
    Note As String
    SrcFile As String
    SrcRow As Long
End Type

Public Type U2ExpenseRec
    Yr As Long
    RoundNo As Long
    RoundName As String
    Dept As String
    Policy As String
    UnitBiz As String
    SubBiz As String
    SubSubBiz As String
    Calc1 As String
    Calc2 As String
    Calc3 As String
    Calc4 As String
    ItemCode As String
    CostStat As String
    ReqAmt As Double
    BaseAmt As Double
    SuppAmt As Double
    FundSpecial As Double
    FundNational As Double
    FundNonStatutory As Double
    FundOther As Double
    FundSelf As Double
    FundTotal As Double
    Manager As String
    PromoMethod As String
    PromoBudget As String
    IsNew As String
    IsPromise As String
    IsCarryover As String
    CarryPeriod As String
    CarryAmt As Double
    CarryFund As String
    IsCentral As String
    Pre1 As String
    Pre2 As String
    Pre3 As String
    Pre4 As String
    Pre5 As String
    Pre6 As String
    Pre7 As String
    Pre8 As String
    PreSummary As String
    Desc As String
    Revised As String
    SrcFile As String
    SrcRow As Long
End Type

Public gU2Inc() As U2IncomeRec
Public gU2Exp() As U2ExpenseRec
Public gU2IncN As Long
Public gU2ExpN As Long
Public gU2LastImport As String
Public gU2LastFiles As Long
Public gU2LastErrors As String
Public gU2Book As Workbook
Public gU2Silent As Boolean
Public gU2Trace As String
Public gU2BuildError As String

' ------------------------------------------------------------------ 초기화 / 추가


Public Sub U2_ResetData()
    Erase gU2Inc
    Erase gU2Exp
    gU2IncN = 0
    gU2ExpN = 0
    gU2LastImport = ""
    gU2LastFiles = 0
    gU2LastErrors = ""
End Sub

Private Sub U2_AppendIncome(ByRef rec As U2IncomeRec)
    If gU2IncN = 0 Then ReDim gU2Inc(1 To 1000)
    If gU2IncN >= UBound(gU2Inc) Then ReDim Preserve gU2Inc(1 To UBound(gU2Inc) + 1000)
    gU2IncN = gU2IncN + 1
    gU2Inc(gU2IncN) = rec
End Sub

Private Sub U2_AppendExpense(ByRef rec As U2ExpenseRec)
    If gU2ExpN = 0 Then ReDim gU2Exp(1 To 1000)
    If gU2ExpN >= UBound(gU2Exp) Then ReDim Preserve gU2Exp(1 To UBound(gU2Exp) + 1000)
    gU2ExpN = gU2ExpN + 1
    gU2Exp(gU2ExpN) = rec
End Sub

' ------------------------------------------------------------------ 파일 읽기

Public Function U2_ImportFiles(ByVal files As Variant) As String
    Dim i As Long
    Dim fp As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim opened As Boolean
    Dim okFiles As Long
    Dim errText As String
    Dim res As String
    Dim used As Boolean

    U2_ResetData

    If IsArray(files) Then
        For i = LBound(files) To UBound(files)
            fp = CStr(files(i))
            If U2_IsExcelFile(fp) Then
                On Error GoTo FileError
                Set wb = Nothing
                opened = False
                If Not U2_IsOpenWorkbook(fp, wb) Then
                    Application.StatusBar = "UBIS 파일 읽는 중... " & U2_FileName(fp)
                    Set wb = Workbooks.Open(Filename:=fp, ReadOnly:=True, UpdateLinks:=0)
                    opened = True
                End If

                Application.Calculate
                used = False
                For Each ws In wb.Worksheets
                    If U2_HasData(ws) Then
                        res = U2_ParseWorkbook(wb, ws, fp)
                        If Left$(res, 2) = "OK" Then
                            okFiles = okFiles + 1
                            used = True
                        Else
                            errText = errText & IIf(errText = "", "", vbCrLf) & res
                        End If
                        Exit For
                    End If
                Next ws

                If Not used Then
                    errText = errText & IIf(errText = "", "", vbCrLf) & _
                              U2_FileName(fp) & ": 읽을 자료가 없습니다."
                End If

                If opened Then wb.Close SaveChanges:=False
                Set wb = Nothing
                On Error GoTo 0
            End If
        Next i
    End If

    Application.StatusBar = False
    gU2LastFiles = okFiles
    gU2LastErrors = errText
    gU2LastImport = Format$(Now, "yyyy-mm-dd hh:nn")

    If gU2IncN = 0 And gU2ExpN = 0 Then
        U2_ImportFiles = "선택한 파일에서 세입·세출 자료를 찾지 못했습니다." & _
                         IIf(errText = "", "", vbCrLf & vbCrLf & errText)
        Exit Function
    End If

    U2_WriteMaster

    U2_ImportFiles = "세입 " & Format$(gU2IncN, "#,##0") & "건, 세출 " & _
                     Format$(gU2ExpN, "#,##0") & "건을 읽었습니다." & _
                     IIf(errText = "", "", vbCrLf & vbCrLf & "확인할 파일:" & vbCrLf & errText)
    Exit Function

FileError:
    errText = errText & IIf(errText = "", "", vbCrLf) & U2_FileName(fp) & ": " & Err.Description
    If opened Then
        On Error Resume Next
        wb.Close SaveChanges:=False
        On Error GoTo 0
    End If
    Set wb = Nothing
    Resume Next
End Function

Private Function U2_ParseWorkbook(ByVal wb As Workbook, ByVal ws As Worksheet, ByVal fp As String) As String
    Dim kind As String
    Dim hr As Long
    Dim info As Object
    Dim yr As Long
    Dim roundNo As Long
    Dim roundName As String
    Dim before As Long

    kind = U2_DetectKind(ws)
    If kind = "" Then
        U2_ParseWorkbook = U2_FileName(fp) & ": 세입·세출 예산요구서 형식을 인식하지 못했습니다."
        Exit Function
    End If

    hr = U2_FindHeaderRow(ws, kind)
    If hr = 0 Then
        U2_ParseWorkbook = U2_FileName(fp) & ": 머리글 행을 찾지 못했습니다."
        Exit Function
    End If

    Set info = U2_FileInfo(ws, fp)
    yr = CLng(info("Year"))
    roundNo = CLng(info("RoundNo"))
    roundName = CStr(info("RoundName"))

    On Error GoTo ParseError
    If kind = "income" Then
        before = gU2IncN
        U2_ParseIncome ws, fp, yr, roundNo, roundName, hr
        If gU2IncN = before Then
            U2_ParseWorkbook = U2_FileName(fp) & ": 세입 자료 행을 찾지 못했습니다."
        Else
            U2_ParseWorkbook = "OK"
        End If
    Else
        before = gU2ExpN
        U2_ParseExpense ws, fp, yr, roundNo, roundName, hr
        If gU2ExpN = before Then
            U2_ParseWorkbook = U2_FileName(fp) & ": 세출 자료 행을 찾지 못했습니다."
        Else
            U2_ParseWorkbook = "OK"
        End If
    End If
    Exit Function

ParseError:
    U2_ParseWorkbook = U2_FileName(fp) & ": " & Err.Description
End Function

Private Function U2_DetectKind(ByVal ws As Worksheet) As String
    Dim s As String
    Dim r As Long
    Dim c As Long

    s = U2_Norm(ws.Name)
    For r = 1 To 5
        For c = 1 To Application.Min(ws.UsedRange.Columns.Count, 20)
            s = s & U2_Norm(ws.Cells(r, c).Value2)
        Next c
    Next r

    If InStr(s, "세입예산") > 0 Or InStr(s, "세입요구서") > 0 Then
        U2_DetectKind = "income"
        Exit Function
    End If
    If InStr(s, "세출예산") > 0 Or InStr(s, "세출요구서") > 0 Or InStr(s, "세출요구") > 0 Then
        U2_DetectKind = "expense"
        Exit Function
    End If

    If U2_FindHeaderRow(ws, "expense") > 0 Then
        U2_DetectKind = "expense"
    ElseIf U2_FindHeaderRow(ws, "income") > 0 Then
        U2_DetectKind = "income"
    End If
End Function

Private Function U2_FindHeaderRow(ByVal ws As Worksheet, ByVal kind As String) As Long
    Dim r As Long
    Dim c As Long
    Dim s As String
    Dim lastC As Long
    Dim hit As Boolean

    lastC = Application.Min(ws.UsedRange.Columns.Count, 80)
    For r = 1 To Application.Min(ws.UsedRange.Rows.Count, 25)
        s = ""
        For c = 1 To lastC
            s = s & U2_Norm(ws.Cells(r, c).Value2) & "|"
        Next c

        If kind = "expense" Then
            hit = (InStr(s, "세부사업") > 0) And (InStr(s, "산출기초") > 0) And _
                  ((InStr(s, "세목코드") > 0) Or (InStr(s, "원가통계목") > 0))
        Else
            hit = (InStr(s, "사업항목") > 0) And (InStr(s, "산출기초") > 0)
        End If

        If hit Then U2_FindHeaderRow = r
    Next r
End Function

Private Function U2_FileInfo(ByVal ws As Worksheet, ByVal fp As String) As Object
    Dim d As Object
    Dim r As Long
    Dim c As Long
    Dim s As String
    Dim y As Long
    Dim rn As Long
    Dim rname As String
    Dim p As Long
    Dim t As String

    Set d = CreateObject("Scripting.Dictionary")
    s = ""
    For r = 1 To Application.Min(ws.UsedRange.Rows.Count, 5)
        For c = 1 To Application.Min(ws.UsedRange.Columns.Count, 30)
            t = U2_Clean(ws.Cells(r, c).Value2)
            If t <> "" Then s = s & " " & t
        Next c
    Next r

    y = U2_YearIn(s)
    If y = 0 Then y = U2_YearIn(U2_FileName(fp))

    rn = 0
    rname = "본예산"
    p = InStr(1, s, "제", vbTextCompare)
    Do While p > 0
        t = Mid$(s, p, 12)
        If t Like "제#회*추경*" Or t Like "제##회*추경*" Then
            rn = CLng(U2_FirstNumber(t))
            Exit Do
        End If
        p = InStr(p + 1, s, "제", vbTextCompare)
    Loop
    If rn = 0 Then
        p = InStr(1, s, "추경", vbTextCompare)
        If p > 0 Then
            t = Mid$(s, Application.Max(1, p - 8), 12)
            rn = CLng(U2_FirstNumber(t))
            If rn = 0 Then rn = 1
        End If
    End If
    If rn > 0 Then rname = CStr(rn) & "차 추경"

    d("Year") = y
    d("RoundNo") = rn
    d("RoundName") = rname
    Set U2_FileInfo = d
End Function

Private Function U2_FirstNumber(ByVal s As String) As String
    Dim i As Long
    Dim ch As String
    Dim out As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch >= "0" And ch <= "9" Then
            out = out & ch
        ElseIf out <> "" Then
            Exit For
        End If
    Next i
    U2_FirstNumber = out
End Function

Private Function U2_YearIn(ByVal s As String) As Long
    Dim i As Long
    Dim y As Long
    Dim best As Long
    For i = 1 To Len(s) - 3
        If Mid$(s, i, 4) Like "20##" Then
            y = CLng(Mid$(s, i, 4))
            If y > best Then best = y
        End If
    Next i
    U2_YearIn = best
End Function

' ------------------------------------------------------------------ 세입 파싱

Private Sub U2_ParseIncome(ByVal ws As Worksheet, ByVal fp As String, ByVal yr As Long, _
                           ByVal roundNo As Long, ByVal roundName As String, ByVal hr As Long)
    Dim map As Object
    Dim unit As Double
    Dim lastRow As Long
    Dim r As Long
    Dim rec As U2IncomeRec
    Dim dept As String
    Dim item As String
    Dim v As Variant

    Set map = U2_MapColumns(ws, hr, "income", yr, roundNo)
    U2_RequireIncomeColumns map, fp
    unit = U2_Unit(ws)
    lastRow = U2_LastRow(ws, map)

    For r = hr + 1 To lastRow
        dept = U2_CellText(ws.Cells(r, map("소관부서")).Value2)
        item = U2_CellText(ws.Cells(r, map("사업항목")).Value2)
        If dept <> "" And item <> "" Then
            rec.Yr = yr
            rec.RoundNo = roundNo
            rec.RoundName = roundName
            rec.Jang = U2_MergedText(ws, r, map("장"))
            rec.Gwan = U2_MergedText(ws, r, map("관"))
            rec.Hang = U2_MergedText(ws, r, map("항"))
            rec.Mok = U2_MergedText(ws, r, map("목"))
            rec.Item = item
            rec.Dept = dept
            rec.PrevAmt = U2_Num(ws.Cells(r, map("기정")).Value2) * unit
            rec.ReqAmt = U2_Num(ws.Cells(r, map("요구")).Value2) * unit
            If map.Exists("최종") Then
                rec.FinalAmt = U2_Num(ws.Cells(r, map("최종")).Value2) * unit
            Else
                rec.FinalAmt = rec.ReqAmt
            End If
            rec.ChangeAmt = rec.FinalAmt - rec.PrevAmt
            If map.Exists("비고") Then rec.Note = U2_CellText(ws.Cells(r, map("비고")).Value2)
            rec.SrcFile = U2_FileName(fp)
            rec.SrcRow = r
            U2_AppendIncome rec
        End If
    Next r
End Sub

' ------------------------------------------------------------------ 세출 파싱

Private Sub U2_ParseExpense(ByVal ws As Worksheet, ByVal fp As String, ByVal yr As Long, _
                            ByVal roundNo As Long, ByVal roundName As String, ByVal hr As Long)
    Dim map As Object
    Dim unit As Double
    Dim lastRow As Long
    Dim r As Long
    Dim rec As U2ExpenseRec
    Dim level As Long
    Dim levelText As String
    Dim calcText As String
    Dim h1 As String
    Dim h2 As String
    Dim h3 As String
    Dim h4 As String
    Dim code As String
    Dim cost As String
    Dim amt As Double
    Dim nextLevel As Long
    Dim isLeaf As Boolean

    Set map = U2_MapColumns(ws, hr, "expense", yr, roundNo)
    U2_RequireExpenseColumns map, fp
    unit = U2_Unit(ws)
    lastRow = U2_LastRow(ws, map)

    For r = hr + 1 To lastRow
        levelText = U2_CellText(ws.Cells(r, map("레벨")).Value2)
        level = CLng(U2_Num(levelText))
        calcText = U2_CellText(ws.Cells(r, map("산출기초")).Value2)
        If level = 0 Then level = U2_LevelFromText(calcText)

        If U2_CellText(ws.Cells(r, map("세세부사업")).Value2) <> "" Then
            h1 = U2_CellText(ws.Cells(r, map("세세부사업")).Value2)
        End If
        If level <= 2 Then
            If calcText <> "" Then h2 = calcText
            h3 = ""
            h4 = ""
        ElseIf level = 3 Then
            If calcText <> "" Then h3 = calcText
            h4 = ""
        ElseIf level >= 4 Then
            If calcText <> "" Then h4 = calcText
        End If

        code = U2_CellText(ws.Cells(r, map("세목코드")).Value2)
        cost = U2_CellText(ws.Cells(r, map("원가통계목")).Value2)
        amt = U2_ExpenseAmount(ws, r, map, unit)
        nextLevel = U2_NextLevel(ws, r, lastRow, map("레벨"))

        isLeaf = (code <> "" Or cost <> "")
        If isLeaf Then
            rec.Yr = yr
            rec.RoundNo = roundNo
            rec.RoundName = roundName
            rec.Dept = U2_CellText(ws.Cells(r, map("부서명")).Value2)
            rec.Policy = U2_CellText(ws.Cells(r, map("정책사업")).Value2)
            rec.UnitBiz = U2_CellText(ws.Cells(r, map("단위사업")).Value2)
            rec.SubBiz = U2_CellText(ws.Cells(r, map("세부사업")).Value2)
            rec.SubSubBiz = U2_CellText(ws.Cells(r, map("세세부사업")).Value2)
            rec.Calc1 = h1
            rec.Calc2 = h2
            rec.Calc3 = h3
            rec.Calc4 = h4
            rec.ItemCode = code
            rec.CostStat = cost
            rec.ReqAmt = amt
            If map.Exists("기정액") Then rec.BaseAmt = U2_Num(ws.Cells(r, map("기정액")).Value2) * unit
            If map.Exists("추경요구액") Then rec.SuppAmt = U2_Num(ws.Cells(r, map("추경요구액")).Value2) * unit
            rec.FundSpecial = U2_Fund(ws, r, map, "재원_특교", unit)
            rec.FundNational = U2_Fund(ws, r, map, "재원_국고", unit)
            rec.FundNonStatutory = U2_Fund(ws, r, map, "재원_비법정", unit)
            rec.FundOther = U2_Fund(ws, r, map, "재원_기타지원금", unit)
            rec.FundSelf = U2_Fund(ws, r, map, "재원_자체", unit)
            rec.FundTotal = U2_Fund(ws, r, map, "재원_합계", unit)
            If rec.FundTotal = 0 Then
                rec.FundTotal = rec.FundSpecial + rec.FundNational + rec.FundNonStatutory + _
                                rec.FundOther + rec.FundSelf
            End If
            If map.Exists("담당자") Then rec.Manager = U2_CellText(ws.Cells(r, map("담당자")).Value2)
            If map.Exists("홍보방법") Then rec.PromoMethod = U2_CellText(ws.Cells(r, map("홍보방법")).Value2)
            If map.Exists("홍보예산") Then rec.PromoBudget = U2_Yn(ws.Cells(r, map("홍보예산")).Value2)
            If map.Exists("신규") Then rec.IsNew = U2_Yn(ws.Cells(r, map("신규")).Value2)
            If map.Exists("공약") Then rec.IsPromise = U2_Yn(ws.Cells(r, map("공약")).Value2)
            If map.Exists("명시이월") Then rec.IsCarryover = U2_Yn(ws.Cells(r, map("명시이월")).Value2)
            If map.Exists("사업기간") Then rec.CarryPeriod = U2_CellText(ws.Cells(r, map("사업기간")).Value2)
            If map.Exists("명시이월액") Then rec.CarryAmt = U2_Num(ws.Cells(r, map("명시이월액")).Value2) * unit
            If map.Exists("명시이월재원") Then rec.CarryFund = U2_CellText(ws.Cells(r, map("명시이월재원")).Value2)
            If map.Exists("본청주관") Then rec.IsCentral = U2_Yn(ws.Cells(r, map("본청주관")).Value2)
            rec.Pre1 = U2_Flag(ws, r, map, "사전1")
            rec.Pre2 = U2_Flag(ws, r, map, "사전2")
            rec.Pre3 = U2_Flag(ws, r, map, "사전3")
            rec.Pre4 = U2_Flag(ws, r, map, "사전4")
            rec.Pre5 = U2_Flag(ws, r, map, "사전5")
            rec.Pre6 = U2_Flag(ws, r, map, "사전6")
            rec.Pre7 = U2_Flag(ws, r, map, "사전7")
            rec.Pre8 = U2_Flag(ws, r, map, "사전8")
            rec.PreSummary = U2_PreSummary(rec)
            If map.Exists("설명") Then rec.Desc = U2_CellText(ws.Cells(r, map("설명")).Value2)
            If map.Exists("수정") Then rec.Revised = U2_CellText(ws.Cells(r, map("수정")).Value2)
            rec.SrcFile = U2_FileName(fp)
            rec.SrcRow = r
            U2_AppendExpense rec
        End If
    Next r
End Sub

Private Function U2_ExpenseAmount(ByVal ws As Worksheet, ByVal r As Long, ByVal map As Object, ByVal unit As Double) As Double
    Dim v As Double
    If map.Exists("최종액") Then
        v = U2_Num(ws.Cells(r, map("최종액")).Value2)
    ElseIf map.Exists("요구액") Then
        v = U2_Num(ws.Cells(r, map("요구액")).Value2)
    End If
    U2_ExpenseAmount = v * unit
End Function

Private Function U2_Fund(ByVal ws As Worksheet, ByVal r As Long, ByVal map As Object, ByVal key As String, ByVal unit As Double) As Double
    If map.Exists(key) Then U2_Fund = U2_Num(ws.Cells(r, map(key)).Value2) * unit
End Function

Private Function U2_Flag(ByVal ws As Worksheet, ByVal r As Long, ByVal map As Object, ByVal key As String) As String
    If map.Exists(key) Then U2_Flag = U2_Yn(ws.Cells(r, map(key)).Value2)
End Function

Private Function U2_PreSummary(ByRef rec As U2ExpenseRec) As String
    Dim s As String
    If rec.Pre1 = "Y" Then s = s & IIf(s = "", "", ", ") & "중기재정계획"
    If rec.Pre2 = "Y" Then s = s & IIf(s = "", "", ", ") & "재정투자심사"
    If rec.Pre3 = "Y" Then s = s & IIf(s = "", "", ", ") & "지방보조금"
    If rec.Pre4 = "Y" Then s = s & IIf(s = "", "", ", ") & "정책연구용역심의"
    If rec.Pre5 = "Y" Then s = s & IIf(s = "", "", ", ") & "공유재산관리계획"
    If rec.Pre6 = "Y" Then s = s & IIf(s = "", "", ", ") & "정수물품취득"
    If rec.Pre7 = "Y" Then s = s & IIf(s = "", "", ", ") & "국외연수심의위원회"
    If rec.Pre8 = "Y" Then s = s & IIf(s = "", "", ", ") & "사무위탁·대행"
    U2_PreSummary = s
End Function

Private Function U2_NextLevel(ByVal ws As Worksheet, ByVal r As Long, ByVal lastRow As Long, ByVal col As Long) As Long
    Dim i As Long
    Dim v As Long
    For i = r + 1 To lastRow
        v = CLng(U2_Num(ws.Cells(i, col).Value2))
        If v > 0 Then
            U2_NextLevel = v
            Exit Function
        End If
    Next i
End Function

Private Function U2_LevelFromText(ByVal s As String) As Long
    Dim ch As String
    Dim code As Long
    s = Trim$(s)
    If s = "" Then Exit Function
    ch = Left$(s, 1)
    code = AscW(ch)
    If ch Like "#" Then
        If Mid$(s, 2, 1) = ")" Then U2_LevelFromText = 3 Else U2_LevelFromText = 1
    ElseIf code >= &HAC00 And code <= &HD7A3 Then
        If Mid$(s, 2, 1) = ")" Then U2_LevelFromText = 4 Else U2_LevelFromText = 2
    End If
End Function

' ------------------------------------------------------------------ 머리글 지도

Private Function U2_MapColumns(ByVal ws As Worksheet, ByVal hr As Long, ByVal kind As String, _
                               ByVal yr As Long, ByVal roundNo As Long) As Object
    Dim map As Object
    Dim calcCols As Object
    Dim c As Long
    Dim r As Long
    Dim lastC As Long
    Dim low As String
    Dim grp As String
    Dim full As String
    Dim t As String
    Dim best As Long
    Dim bestScore As Long
    Dim score As Long
    Dim k As Variant

    Set map = CreateObject("Scripting.Dictionary")
    Set calcCols = CreateObject("Scripting.Dictionary")
    lastC = Application.Min(ws.UsedRange.Columns.Count, 80)

    For c = 1 To lastC
        low = U2_HeaderText(ws, hr, c)
        grp = ""
        For r = hr - 1 To Application.Max(1, hr - 5) Step -1
            t = U2_HeaderText(ws, r, c)
            If t <> "" Then grp = t & "|" & grp
        Next r
        full = grp & "|" & low

        If kind = "income" Then
            If InStr(full, "사업항목") > 0 Then U2_Put map, "사업항목", c
            If InStr(full, "소관부서") > 0 Then U2_Put map, "소관부서", c
            If low = "장" Then U2_Put map, "장", c
            If low = "관" Then U2_Put map, "관", c
            If low = "항" Then U2_Put map, "항", c
            If low = "목" Then U2_Put map, "목", c
            If (InStr(low, "예산액") > 0 Or InStr(full, "기정예산액") > 0) And _
               InStr(low, "산출기초") = 0 And InStr(low, "금액") = 0 Then
                U2_Put map, "기정", c
            End If
            If InStr(low, "금액") > 0 Then
                If InStr(full, "예산요구액") > 0 Or InStr(full, "세입예산요구액") > 0 Or _
                   InStr(full, "요구액") > 0 Then U2_Put map, "요구", c
            End If
            If InStr(low, "최종예산액") > 0 Or InStr(full, "최종예산액") > 0 Then U2_Put map, "최종", c
            If InStr(full, "비고") > 0 Then U2_Put map, "비고", c
        Else
            If InStr(full, "세목코드") > 0 Then U2_Put map, "세목코드", c
            If InStr(full, "원가통계목") > 0 Then U2_Put map, "원가통계목", c
            If InStr(full, "세세부사업") > 0 Then U2_Put map, "세세부사업", c
            If InStr(full, "세부사업") > 0 And InStr(full, "세세부사업") = 0 Then
                If Not map.Exists("세부사업") Then
                    map("세부사업") = c
                ElseIf Not map.Exists("세부사업2") Then
                    map("세부사업2") = c
                End If
            End If
            If InStr(full, "부서명") > 0 Then U2_Put map, "부서명", c
            If InStr(full, "정책사업") > 0 Then U2_Put map, "정책사업", c
            If InStr(full, "단위사업") > 0 Then U2_Put map, "단위사업", c
            If InStr(full, "목적지정사업") > 0 Then U2_Put map, "목적지정사업", c
            If InStr(full, "레벨") > 0 Then U2_Put map, "레벨", c
            If InStr(full, "산출기초") > 0 Then calcCols.Add c, full
            If InStr(full, "기정예산액") > 0 Then U2_Put map, "기정액", c
            If InStr(full, "예산요구액") > 0 And InStr(full, "조정후") = 0 Then U2_Put map, "요구액", c
            If InStr(full, "최종예산") > 0 Then U2_Put map, "최종액", c
            If InStr(full, "추경요구액") > 0 Then U2_Put map, "추경요구액", c
            If InStr(full, "전년도예산액") > 0 Then U2_Put map, "전년액", c

            If InStr(full, "재원별") > 0 Then
                If InStr(full, "특교") > 0 And InStr(full, "금액") > 0 Then U2_Put map, "재원_특교", c
                If InStr(full, "국고") > 0 And InStr(full, "금액") > 0 Then U2_Put map, "재원_국고", c
                If InStr(full, "비법정") > 0 And InStr(full, "금액") > 0 Then U2_Put map, "재원_비법정", c
                If InStr(full, "기타") > 0 And InStr(full, "금액") > 0 Then U2_Put map, "재원_기타지원금", c
                If InStr(full, "자체") > 0 Then
                    If InStr(full, "금액") > 0 Then
                        map("재원_자체") = c
                    Else
                        U2_Put map, "재원_자체", c
                    End If
                End If
                If InStr(full, "합계") > 0 And InStr(full, "사업명") = 0 Then U2_Put map, "재원_합계", c
            End If

            If (InStr(full, "업무담당자") > 0 Or InStr(full, "담당자") > 0) And InStr(full, "홍보") = 0 Then U2_Put map, "담당자", c
            If InStr(full, "홍보") > 0 And InStr(full, "방법") > 0 Then U2_Put map, "홍보방법", c
            If InStr(full, "홍보") > 0 And (InStr(full, "여부") > 0 Or InStr(full, "예산") > 0) Then U2_Put map, "홍보예산", c
            If InStr(full, "신규") > 0 Then U2_Put map, "신규", c
            If InStr(full, "공약") > 0 Then U2_Put map, "공약", c
            If InStr(full, "명시이월") > 0 And InStr(full, "여부") > 0 Then U2_Put map, "명시이월", c
            If InStr(full, "사업기간") > 0 Then U2_Put map, "사업기간", c
            If InStr(full, "명시이월") > 0 And InStr(full, "액") > 0 Then U2_Put map, "명시이월액", c
            If InStr(full, "명시이월") > 0 And InStr(full, "재원") > 0 Then U2_Put map, "명시이월재원", c
            If InStr(full, "본청주관") > 0 Then U2_Put map, "본청주관", c
            If InStr(full, "중기재정계획") > 0 Then U2_Put map, "사전1", c
            If InStr(full, "재정투자심사") > 0 Then U2_Put map, "사전2", c
            If InStr(full, "지방보조금") > 0 Then U2_Put map, "사전3", c
            If InStr(full, "정책연구용역심의") > 0 Then U2_Put map, "사전4", c
            If InStr(full, "공유재산관리계획") > 0 Then U2_Put map, "사전5", c
            If InStr(full, "정수물품취득") > 0 Then U2_Put map, "사전6", c
            If InStr(full, "국외연수심의위원회") > 0 Then U2_Put map, "사전7", c
            If InStr(full, "사무위탁") > 0 Then U2_Put map, "사전8", c
            If InStr(full, "설명사항") > 0 Then U2_Put map, "설명", c
            If InStr(full, "수정여부") > 0 Then U2_Put map, "수정", c
        End If
    Next c

    If kind = "expense" Then
        best = 0
        bestScore = -1
        For Each k In calcCols.Keys
            full = CStr(calcCols(k))
            score = 0
            If roundNo > 0 Then
                If InStr(full, "최종") > 0 Or InStr(full, "요구") > 0 Then score = 3
                If InStr(full, "기정") > 0 Then score = 2
            Else
                If yr > 0 And InStr(full, CStr(yr)) > 0 Then score = 3
                If InStr(full, "올해") > 0 Then score = 3
                If InStr(full, "전년") > 0 Then score = 1
            End If
            If score > bestScore Then
                bestScore = score
                best = CLng(k)
            End If
        Next k
        If best = 0 Then
            For Each k In calcCols.Keys
                If CLng(k) > best Then best = CLng(k)
            Next k
        End If
        If best > 0 Then map("산출기초") = best
    End If

    Set U2_MapColumns = map
End Function

Private Sub U2_Put(ByRef map As Object, ByVal key As String, ByVal c As Long)
    If Not map.Exists(key) Then map(key) = c
End Sub

Private Sub U2_RequireIncomeColumns(ByRef map As Object, ByVal fp As String)
    Dim missing As String
    If Not map.Exists("사업항목") Then missing = missing & " 사업항목"
    If Not map.Exists("소관부서") Then missing = missing & " 소관부서"
    If Not map.Exists("기정") Then missing = missing & " 기정예산액"
    If Not map.Exists("요구") Then missing = missing & " 예산요구액"
    If missing <> "" Then
        Err.Raise vbObjectError + 8101, "U2_ParseIncome", _
                  U2_FileName(fp) & "에서 다음 열을 찾지 못했습니다:" & missing & vbCrLf & _
                  "UBIS에서 내려받은 세입예산요구서 원본인지 확인하세요."
    End If
End Sub

Private Sub U2_RequireExpenseColumns(ByRef map As Object, ByVal fp As String)
    Dim missing As String
    If Not map.Exists("세목코드") Then missing = missing & " 세목코드"
    If Not map.Exists("원가통계목") Then missing = missing & " 원가통계목"
    If Not map.Exists("세부사업") Then missing = missing & " 세부사업"
    If Not map.Exists("부서명") Then missing = missing & " 부서명"
    If Not map.Exists("산출기초") Then missing = missing & " 산출기초"
    If Not map.Exists("요구액") And Not map.Exists("최종액") Then missing = missing & " 예산요구액/최종예산"
    If missing <> "" Then
        Err.Raise vbObjectError + 8102, "U2_ParseExpense", _
                  U2_FileName(fp) & "에서 다음 열을 찾지 못했습니다:" & missing & vbCrLf & _
                  "UBIS에서 내려받은 세출예산요구서 원본인지 확인하세요."
    End If
End Sub

' ------------------------------------------------------------------ 공통 셀 읽기

Private Function U2_HeaderText(ByVal ws As Worksheet, ByVal r As Long, ByVal c As Long) As String
    Dim v As Variant
    On Error Resume Next
    If ws.Cells(r, c).MergeCells Then
        v = ws.Cells(r, c).MergeArea.Cells(1, 1).Value2
    Else
        v = ws.Cells(r, c).Value2
    End If
    On Error GoTo 0
    U2_HeaderText = U2_Norm(v)
End Function

Private Function U2_MergedText(ByVal ws As Worksheet, ByVal r As Long, ByVal c As Long) As String
    Dim v As Variant
    On Error Resume Next
    If ws.Cells(r, c).MergeCells Then
        v = ws.Cells(r, c).MergeArea.Cells(1, 1).Value2
    Else
        v = ws.Cells(r, c).Value2
    End If
    On Error GoTo 0
    U2_MergedText = U2_CellText(v)
End Function

Private Function U2_CellText(ByVal v As Variant) As String
    If IsError(v) Then
        U2_CellText = ""
    ElseIf IsEmpty(v) Then
        U2_CellText = ""
    Else
        U2_CellText = Trim$(CStr(v))
    End If
End Function

Private Function U2_Clean(ByVal v As Variant) As String
    U2_Clean = U2_CellText(v)
End Function

Private Function U2_Norm(ByVal v As Variant) As String
    Dim s As String
    s = U2_CellText(v)
    s = Replace(s, vbCr, "")
    s = Replace(s, vbLf, "")
    s = Replace(s, vbTab, "")
    s = Replace(s, " ", "")
    s = Replace(s, "(", "")
    s = Replace(s, ")", "")
    s = Replace(s, "（", "")
    s = Replace(s, "）", "")
    s = Replace(s, "①", "")
    s = Replace(s, "②", "")
    s = Replace(s, "③", "")
    U2_Norm = UCase$(s)
End Function

Public Function U2_LabelText(ByVal s As String) As String
    Dim p As Long
    Dim q As Long
    Dim ch As String
    s = Trim$(s)
    If s = "" Then Exit Function
    For p = 2 To Len(s)
        If Mid$(s, p, 1) = "원" Then
            ch = Mid$(s, p - 1, 1)
            If ch >= "0" And ch <= "9" Then
                q = InStrRev(Left$(s, p), " ")
                If q > 0 Then s = Left$(s, q - 1)
                Exit For
            End If
        End If
    Next p
    s = Trim$(s)
    Do
        If Right$(s, 2) = "0=" Then
            s = Trim$(Left$(s, Len(s) - 2))
        ElseIf Right$(s, 1) = "=" Then
            s = Trim$(Left$(s, Len(s) - 1))
        Else
            Exit Do
        End If
    Loop
    U2_LabelText = Trim$(s)
End Function

Public Function U2_Num(ByVal v As Variant) As Double
    Dim s As String
    If IsEmpty(v) Or IsError(v) Then Exit Function
    If IsNumeric(v) Then
        U2_Num = CDbl(v)
    Else
        s = Replace(CStr(v), ",", "")
        s = Replace(s, "원", "")
        s = Replace(s, "천원", "")
        s = Trim$(s)
        If IsNumeric(s) Then U2_Num = CDbl(s)
    End If
End Function

Private Function U2_Yn(ByVal v As Variant) As String
    Dim s As String
    s = UCase$(Replace(U2_CellText(v), " ", ""))
    If s = "" Or s = "N" Or s = "NO" Or s = "X" Or s = "-" Or s = "미해당" Or s = "해당없음" Then
        U2_Yn = "N"
    Else
        U2_Yn = "Y"
    End If
End Function

Private Function U2_LastRow(ByVal ws As Worksheet, ByRef map As Object) As Long
    Dim key As String
    Dim col As Long
    If map.Exists("세목코드") Then
        col = map("세목코드")
    ElseIf map.Exists("사업항목") Then
        col = map("사업항목")
    Else
        col = 1
    End If
    U2_LastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
    Dim lastUsed As Long
    lastUsed = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    If U2_LastRow < lastUsed Then U2_LastRow = lastUsed
End Function

Private Function U2_Unit(ByVal ws As Worksheet) As Double
    Dim r As Long
    Dim c As Long
    Dim s As String
    For r = 1 To Application.Min(ws.UsedRange.Rows.Count, 6)
        For c = 1 To ws.UsedRange.Columns.Count
            s = s & U2_Norm(ws.Cells(r, c).Value2)
        Next c
    Next r
    If InStr(s, "천원") > 0 Then
        U2_Unit = 1000
    ElseIf InStr(s, "단위") > 0 And InStr(s, "원") > 0 Then
        U2_Unit = 1
    Else
        U2_Unit = 1000
    End If
End Function

Private Function U2_IsExcelFile(ByVal fp As String) As Boolean
    Dim e As String
    e = LCase$(Mid$(fp, InStrRev(fp, ".") + 1))
    U2_IsExcelFile = (e = "xls" Or e = "xlsx" Or e = "xlsm" Or e = "xlsb")
End Function

Private Function U2_IsOpenWorkbook(ByVal fp As String, ByRef wb As Workbook) As Boolean
    Dim w As Workbook
    For Each w In Application.Workbooks
        If StrComp(w.FullName, fp, vbTextCompare) = 0 Then
            Set wb = w
            U2_IsOpenWorkbook = True
            Exit Function
        End If
    Next w
End Function

Private Function U2_HasData(ByVal ws As Worksheet) As Boolean
    On Error Resume Next
    U2_HasData = (ws.UsedRange.Cells.CountLarge > 1) And (ws.Visible = xlSheetVisible)
    On Error GoTo 0
End Function

Private Function U2_FileName(ByVal fp As String) As String
    Dim p As Long
    p = InStrRev(fp, "\")
    If p > 0 Then U2_FileName = Mid$(fp, p + 1) Else U2_FileName = fp
End Function

' ------------------------------------------------------------------ 숨김 원본 시트 쓰기

Public Sub U2_WriteMaster()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim arr As Variant
    Dim i As Long
    Dim c As Long

    Set wb = U2_Book()
    If wb Is Nothing Then Exit Sub

    ' 세입
    Set ws = U2_GetSheet(wb, U2_SH_INC_RAW, True)
    ws.Cells.Clear
    ws.Range("A1").Resize(1, 16).Value2 = Array( _
        "연도", "차수", "차수명", "장", "관", "항", "목", "사업항목", "소관부서", _
        "기정예산", "요구액", "최종예산", "증감", "비고", "출처파일", "원본행")
    If gU2IncN > 0 Then
        ReDim arr(1 To gU2IncN, 1 To 16)
        For i = 1 To gU2IncN
            arr(i, 1) = gU2Inc(i).Yr
            arr(i, 2) = gU2Inc(i).RoundNo
            arr(i, 3) = gU2Inc(i).RoundName
            arr(i, 4) = gU2Inc(i).Jang
            arr(i, 5) = gU2Inc(i).Gwan
            arr(i, 6) = gU2Inc(i).Hang
            arr(i, 7) = gU2Inc(i).Mok
            arr(i, 8) = gU2Inc(i).Item
            arr(i, 9) = gU2Inc(i).Dept
            arr(i, 10) = gU2Inc(i).PrevAmt
            arr(i, 11) = gU2Inc(i).ReqAmt
            arr(i, 12) = gU2Inc(i).FinalAmt
            arr(i, 13) = gU2Inc(i).ChangeAmt
            arr(i, 14) = gU2Inc(i).Note
            arr(i, 15) = gU2Inc(i).SrcFile
            arr(i, 16) = gU2Inc(i).SrcRow
        Next i
        ws.Range("A2").Resize(gU2IncN, 16).Value2 = arr
    End If
    U2_FormatRaw ws, 16
    ws.Visible = xlSheetVeryHidden

    ' 세출
    Set ws = U2_GetSheet(wb, U2_SH_EXP_RAW, True)
    ws.Cells.Clear
    ws.Range("A1").Resize(1, 46).Value2 = Array( _
        "연도", "차수", "차수명", "부서명", "정책사업", "단위사업", "세부사업", "세세부사업", _
        "산출1", "산출2", "산출3", "산출4", "세목코드", "원가통계목", "예산요구액", "기정예산", "추경요구액", _
        "재원_특교", "재원_국고", "재원_비법정", "재원_기타지원금", "재원_자체", "재원_합계", _
        "업무담당자", "홍보방법", "홍보예산여부", "신규여부", "공약사업", "명시이월여부", "사업기간", _
        "명시이월액", "명시이월재원", "본청주관", "사전1", "사전2", "사전3", "사전4", "사전5", "사전6", _
        "사전7", "사전8", "사전절차요약", "설명사항", "수정여부", "출처파일", "원본행")
    If gU2ExpN > 0 Then
        ReDim arr(1 To gU2ExpN, 1 To 46)
        For i = 1 To gU2ExpN
            arr(i, 1) = gU2Exp(i).Yr
            arr(i, 2) = gU2Exp(i).RoundNo
            arr(i, 3) = gU2Exp(i).RoundName
            arr(i, 4) = gU2Exp(i).Dept
            arr(i, 5) = gU2Exp(i).Policy
            arr(i, 6) = gU2Exp(i).UnitBiz
            arr(i, 7) = gU2Exp(i).SubBiz
            arr(i, 8) = gU2Exp(i).SubSubBiz
            arr(i, 9) = gU2Exp(i).Calc1
            arr(i, 10) = gU2Exp(i).Calc2
            arr(i, 11) = gU2Exp(i).Calc3
            arr(i, 12) = gU2Exp(i).Calc4
            arr(i, 13) = gU2Exp(i).ItemCode
            arr(i, 14) = gU2Exp(i).CostStat
            arr(i, 15) = gU2Exp(i).ReqAmt
            arr(i, 16) = gU2Exp(i).BaseAmt
            arr(i, 17) = gU2Exp(i).SuppAmt
            arr(i, 18) = gU2Exp(i).FundSpecial
            arr(i, 19) = gU2Exp(i).FundNational
            arr(i, 20) = gU2Exp(i).FundNonStatutory
            arr(i, 21) = gU2Exp(i).FundOther
            arr(i, 22) = gU2Exp(i).FundSelf
            arr(i, 23) = gU2Exp(i).FundTotal
            arr(i, 24) = gU2Exp(i).Manager
            arr(i, 25) = gU2Exp(i).PromoMethod
            arr(i, 26) = gU2Exp(i).PromoBudget
            arr(i, 27) = gU2Exp(i).IsNew
            arr(i, 28) = gU2Exp(i).IsPromise
            arr(i, 29) = gU2Exp(i).IsCarryover
            arr(i, 30) = gU2Exp(i).CarryPeriod
            arr(i, 31) = gU2Exp(i).CarryAmt
            arr(i, 32) = gU2Exp(i).CarryFund
            arr(i, 33) = gU2Exp(i).IsCentral
            arr(i, 34) = gU2Exp(i).Pre1
            arr(i, 35) = gU2Exp(i).Pre2
            arr(i, 36) = gU2Exp(i).Pre3
            arr(i, 37) = gU2Exp(i).Pre4
            arr(i, 38) = gU2Exp(i).Pre5
            arr(i, 39) = gU2Exp(i).Pre6
            arr(i, 40) = gU2Exp(i).Pre7
            arr(i, 41) = gU2Exp(i).Pre8
            arr(i, 42) = gU2Exp(i).PreSummary
            arr(i, 43) = gU2Exp(i).Desc
            arr(i, 44) = gU2Exp(i).Revised
            arr(i, 45) = gU2Exp(i).SrcFile
            arr(i, 46) = gU2Exp(i).SrcRow
        Next i
        ws.Range("A2").Resize(gU2ExpN, 46).Value2 = arr
    End If
    U2_FormatRaw ws, 46
    ws.Visible = xlSheetVeryHidden

    U2_WriteLists wb
End Sub

Private Sub U2_FormatRaw(ByVal ws As Worksheet, ByVal nCol As Long)
    With ws.Range("A1").Resize(1, nCol)
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(26, 64, 103)
    End With
    ws.Rows(1).RowHeight = 22
End Sub

Private Sub U2_WriteLists(ByVal wb As Workbook)
    Dim ws As Worksheet
    Dim dDept As Object, dSub As Object, dCode As Object, dRound As Object
    Dim i As Long
    Dim k As Variant
    Dim r As Long

    Set dDept = CreateObject("Scripting.Dictionary")
    Set dSub = CreateObject("Scripting.Dictionary")
    Set dCode = CreateObject("Scripting.Dictionary")
    Set dRound = CreateObject("Scripting.Dictionary")

    For i = 1 To gU2ExpN
        If gU2Exp(i).Dept <> "" Then dDept(gU2Exp(i).Dept) = 1
        If gU2Exp(i).SubBiz <> "" Then dSub(gU2Exp(i).SubBiz) = 1
        If gU2Exp(i).ItemCode <> "" Then dCode(gU2Exp(i).ItemCode & " | " & gU2Exp(i).CostStat) = 1
        dRound(CStr(gU2Exp(i).RoundNo) & "|" & gU2Exp(i).RoundName) = 1
    Next i
    For i = 1 To gU2IncN
        If gU2Inc(i).Dept <> "" Then dDept(gU2Inc(i).Dept) = 1
        dRound(CStr(gU2Inc(i).RoundNo) & "|" & gU2Inc(i).RoundName) = 1
    Next i

    Set ws = U2_GetSheet(wb, U2_SH_LISTS, True)
    ws.Cells.Clear
    ws.Range("A1").Value2 = "부서"
    ws.Range("B1").Value2 = "세부사업"
    ws.Range("C1").Value2 = "세목"
    ws.Range("D1").Value2 = "차수"

    r = 2
    For Each k In dDept.Keys
        ws.Cells(r, 1).Value2 = k
        r = r + 1
    Next k
    r = 2
    For Each k In dSub.Keys
        ws.Cells(r, 2).Value2 = k
        r = r + 1
    Next k
    r = 2
    For Each k In dCode.Keys
        ws.Cells(r, 3).Value2 = k
        r = r + 1
    Next k
    r = 2
    For Each k In dRound.Keys
        ws.Cells(r, 4).Value2 = k
        r = r + 1
    Next k

    ws.Visible = xlSheetVeryHidden
End Sub

' ------------------------------------------------------------------ 시트 도우미

Public Function U2_Book() As Workbook
    On Error Resume Next
    If Not ActiveWorkbook Is Nothing Then
        If ActiveWorkbook.Name Like "UBIS*" Or U2_HasSheet(ActiveWorkbook, U2_SH_GUIDE) Then
            Set U2_Book = ActiveWorkbook
            Exit Function
        End If
    End If
    If Not gU2Book Is Nothing Then
        Set U2_Book = gU2Book
        Exit Function
    End If
    Set U2_Book = ActiveWorkbook
End Function

Public Function U2_GetSheet(ByVal wb As Workbook, ByVal nm As String, Optional ByVal create As Boolean = False) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(nm)
    On Error GoTo 0
    If ws Is Nothing And create Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = nm
    End If
    Set U2_GetSheet = ws
End Function

Public Function U2_HasSheet(ByVal wb As Workbook, ByVal nm As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(nm)
    On Error GoTo 0
    U2_HasSheet = Not ws Is Nothing
End Function
