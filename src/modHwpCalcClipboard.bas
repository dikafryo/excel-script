Attribute VB_Name = "modHwpCalcClipboard"
Option Explicit

' =================================================================================
' [v2.6] 리본 [HWPX활용기능 ▶ 한글(HWPX) 계산확인] 통합 메뉴 (tag "HwpCalcMenu")
'   1 계산 내용 확인 : 한글에서 산출 근거 글을 복사(Ctrl+C) → 한글계산1차
'   2 표 계산 확인   : 한글 표를 복사 → 표계산 (열 자동 인식: 항목명·차이 열이 함께 복사돼도 됨)
'   결과는 새 통합 문서에 만들고, 한글에 바로 붙일 수 있게 표 형태로 클립보드에 복사한다.
' =================================================================================
Public HCM_Quiet As Boolean   ' 메뉴에서 호출 중이면 개별 완료 메시지 생략

Public Sub 한글계산확인(control As IRibbonControl)
    Dim choice As Long, ws As Worksheet, rng As Range, lastRow As Long, nR As Long, nC As Long, Extra As String
    choice = UI_Choose("한글(HWPX) 계산 확인", _
        "한글에서 내용을 복사(Ctrl+C)한 뒤 실행하세요. 결과는 새 통합 문서에 만들고," & vbCrLf & _
        "한글에 바로 붙여 넣을 수 있게 표 형태로 클립보드에도 복사합니다.", _
        Array("계산 내용 확인   (산출 근거 글을 복사한 뒤 - 줄마다 금액 계산 · 합계 대조)", _
              "표 계산 확인   (표를 블록 복사한 뒤 - 산출내용·금액·전년도·차이 열 자동 인식)"))
    If choice = 0 Then Exit Sub
    Workbooks.Add
    HCM_Quiet = True
    On Error GoTo Done
    If choice = 1 Then
        한글계산1차 Nothing
        Set ws = HCM_Sheet("한글계산")
        If Not ws Is Nothing Then
            lastRow = ws.Cells(ws.rows.Count, "B").End(xlUp).row
            If lastRow >= 3 Then Set rng = ws.Range("A2:D" & lastRow)
        End If
    Else
        표계산 Nothing
        Set ws = HCM_Sheet("표계산")
        If Not ws Is Nothing Then
            lastRow = ws.Cells(ws.rows.Count, "A").End(xlUp).row
            If lastRow >= 2 Then Set rng = ws.Range("A1:G" & lastRow)
        End If
        Extra = TC_LastColInfo
    End If
Done:
    HCM_Quiet = False
    On Error GoTo 0
    If rng Is Nothing Then Exit Sub   ' 자료 없음 등 (안내는 각 기능이 표시)
    If HC_CopyRangeAsTable(rng, nR, nC, True, 10) Then
        MsgBox "계산을 마치고 결과 표(" & nR & "행 × " & nC & "열)를 클립보드에 복사했습니다." & vbCrLf & vbCrLf & _
               "한글에서 Ctrl+V 로 붙여 넣으면 표로 들어갑니다. (불일치·합계 표시 색 유지 · 얇은 실선 · 맑은 고딕 · 열 너비 자동)" & _
               IIf(Len(Extra) > 0, vbCrLf & vbCrLf & Extra, ""), vbInformation, "한글(HWPX) 계산 확인"
    Else
        MsgBox "계산은 마쳤지만 클립보드 복사에 실패했습니다. 결과 시트를 직접 복사해 주세요." & _
               IIf(Len(Extra) > 0, vbCrLf & vbCrLf & Extra, ""), vbExclamation, "한글(HWPX) 계산 확인"
    End If
End Sub

Private Function HCM_Sheet(ByVal nm As String) As Worksheet
    On Error Resume Next
    Set HCM_Sheet = ActiveWorkbook.Worksheets(nm)
    On Error GoTo 0
End Function


' =================================================================================
' 메인 프로시저: 클립보드 텍스트를 분석, 계산하고 서식을 적용하는 전체 프로세스를 실행합니다.
' =================================================================================
Public Sub 한글계산1차(control As IRibbonControl)

    ' --- 1. 변수 선언 ---
    Dim ws As Worksheet
    Dim prevCalc As XlCalculation
    Dim i As Long, r As Long, k As Long
        Dim tempExpr As String
    
    ' 데이터 처리 관련 변수
    Dim dataText As String, normalized As String, lineText As String, parsedText As String
    Dim lines() As String
    Dim lhs As String, rhs As String
    Dim resultVal As Double, eqVal As Double
    Dim remarks As String
    
    ' 출력 및 요약 관련 변수
    Dim outRow As Long, dataEndRow As Long, lastRow As Long
    Dim noCounter As Long
    Dim calcTotal As Double, budgetAmount As Double
    
    ' 상태 플래그 변수
    Dim isBudget As Boolean, hasMoney As Boolean, excludeLow As Boolean, showRow As Boolean
    
    ' 부모-자식 항목 비교를 위한 변수
    Dim parentValue As Double, childSum As Double, childCount As Long

    ' --- 2. 초기 설정 및 성능 최적화 ---
    On Error GoTo SAFE_EXIT
    ' [수정] 클립보드를 먼저 확인한다 (비어 있으면 시트를 건드리지 않음)
    dataText = GetClipboardText()
    If Len(Trim$(dataText)) = 0 Then
        MsgBox "클립보드에 텍스트가 없습니다." & vbCrLf & vbCrLf & _
               "한글에서 계산할 사업설명서 내용을 복사(Ctrl+C)한 뒤 [한글계산]을 누르세요.", vbExclamation, "한글계산"
        Exit Sub
    End If
    ' [수정] 결과는 활성 시트를 지우지 않고 '한글계산' 시트에 기록
    Set ws = HC_ResultSheet(ActiveWorkbook)
    ws.Activate
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        prevCalc = .Calculation
        .Calculation = xlCalculationManual
    End With
    ws.Range("A1:D200000").Clear
    ws.Columns("A:D").NumberFormat = "General"
    
    ws.Range("A2:D2").value = Array("No.", "원문", "계산값", "비고")
    With ws.Range("A2:D2")
        .Font.name = "맑은 고딕"
        .Font.Size = 14
        .Font.Bold = True
        .Interior.color = RGB(199, 230, 255) ' 하늘색
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ' --- 4. 클립보드 데이터 읽기 및 기본 처리 (1차 Pass) ---
    ' (클립보드는 위에서 이미 읽음)
    
    ' 줄바꿈 정규화 후 배열로 분리
    normalized = Replace(Replace(dataText, vbCrLf, vbLf), vbCr, vbLf)
    lines = Split(normalized, vbLf)
    
    ' 변수 초기화
    outRow = 3
    noCounter = 0
    budgetAmount = 0
    
    ' 각 라인 처리
    For i = LBound(lines) To UBound(lines)
        lineText = lines(i) ' Trim 제거, 원본 공백 유지를 위해
        
        ' 빈 줄이거나 제외할 라인이면 건너뛰기
        If Len(Trim(lineText)) = 0 Or IsExcludedLine(lineText) Then GoTo ContinueNext
        
        parsedText = StripPrefixMarkers(lineText)
        If Len(Trim(parsedText)) = 0 Then GoTo ContinueNext
        
        isBudget = IsBudgetOrTotalLine(parsedText)
        
        ' "=" 기준으로 좌/우 분리
        Dim eqPos As Long
        eqPos = InStr(parsedText, "=")
        If eqPos > 0 Then
            lhs = Trim$(Left$(parsedText, eqPos - 1))
            rhs = Trim$(Mid$(parsedText, eqPos + 1))
        Else
            lhs = parsedText
            rhs = ""
        End If
        
        remarks = ""
        resultVal = 0
        showRow = False
        excludeLow = False
        
        ' 단위 없는 숫자 검사
        If HasNumberWithoutUnit(lhs) Then
            remarks = AppendRemark(remarks, "수량단위 확인")
            showRow = True
        End If
        
        hasMoney = HasMoneyUnit(lhs)
        Dim hasMulDivOperator As Boolean
        hasMulDivOperator = InStr(1, lhs, "*") > 0 Or InStr(1, lhs, "×") > 0 Or InStr(1, lhs, "x") > 0 Or InStr(1, lhs, "X") > 0 Or InStr(1, lhs, "÷") > 0 Or InStr(1, lhs, "/") > 0

        If hasMoney Or hasMulDivOperator Then
            If hasMulDivOperator Then
               ' resultVal = ComputeMulDivFromText(lhs)
               ' resultVal = ApplyUnitFromLine(lhs, resultVal)
            resultVal = EvaluateFormulaInText(lhs)
            resultVal = ApplyUnitFromLine(lhs, resultVal)

            
            
            Else
                resultVal = SumMoneyValuesWon(lhs)
            End If
            
            If hasMoney And resultVal > 0 And resultVal < 1000 Then
                excludeLow = True
                remarks = AppendRemark(remarks, "확인 필요(숫자: " & Format(resultVal, "#,##0") & "원)")
                showRow = True
            End If
        Else
            remarks = AppendRemark(remarks, "계산 제외")
            showRow = True
        End If
        
        If isBudget Then
            Dim bVal As Double
            bVal = SumMoneyValuesWon(parsedText)
            If bVal = 0 Then bVal = ParseFirstNumberInText(parsedText)
            If bVal <> 0 Then budgetAmount = bVal
        End If
        
        If eqPos > 0 And Not excludeLow And resultVal > 0 Then
            eqVal = ExtractEqualsValueInWon(rhs)
            If eqVal > 0 And CLng(resultVal) <> CLng(eqVal) Then
                ws.Range("B" & outRow & ":C" & outRow).Interior.color = RGB(255, 230, 230)
                remarks = AppendRemark(remarks, "원문표기=" & Format(eqVal, "#,##0") & "원, 계산=" & Format(resultVal, "#,##0") & "원 불일치")
            End If
        End If
        
        If showRow Or (resultVal > 0 And Not excludeLow) Then
            noCounter = noCounter + 1
            ws.Cells(outRow, "A").value = noCounter
            ws.Cells(outRow, "B").value = lineText
            
            If resultVal > 0 And Not excludeLow Then
                ws.Cells(outRow, "C").value = resultVal
            Else
                ws.Cells(outRow, "C").value = ""
            End If
            
            ws.Cells(outRow, "D").value = remarks
            outRow = outRow + 1
        End If
        
ContinueNext:
    Next i
    
    dataEndRow = outRow - 1
    
    ' --- 5. 후처리, 최종 계산, 서식 지정 (2차 통합 Pass) ---
    If dataEndRow < 3 Then GoTo NO_DATA

    ' C열을 제외한 일괄 서식 적용
    ws.Range("C3:C" & dataEndRow).HorizontalAlignment = xlRight
    ws.Range("B3:D" & dataEndRow).Font.name = "맑은 고딕"
    ws.Range("B3:C" & dataEndRow).Font.Size = 14
    ws.Range("D3:D" & dataEndRow).Font.Size = 12
    ws.Range("A3:A" & dataEndRow & ",D3:D" & dataEndRow).HorizontalAlignment = xlCenter
    ws.Range("A3:A" & dataEndRow & ",D3:D" & dataEndRow).VerticalAlignment = xlCenter
    
    ' 최종 합계 계산 및 조건부 서식 적용
    calcTotal = 0
    For r = 3 To dataEndRow
        Dim cellValueB As String
        Dim currentValue As Double
        Dim isMatchedParent As Boolean
        
        cellValueB = ws.Cells(r, "B").value
        currentValue = val(ws.Cells(r, "C").value)
        ' [수정] 예산액·총액·합계·소계 줄은 문서의 총액이므로 항목 합산에서 제외
        If IsBudgetOrTotalLine(cellValueB) And currentValue <> 0 Then
            ws.Cells(r, "D").value = AppendRemark(CStr(ws.Cells(r, "D").value), "총액 줄(합산 제외)")
            currentValue = 0
        End If
        isMatchedParent = False
        
        Dim parentIndent As Integer
        parentIndent = GetIndentLevel(cellValueB)
        
        ' 다음 행이 있고, 다음 행의 들여쓰기가 현재 행보다 깊은 경우, 현재 행을 '부모'로 간주
        If r + 1 <= dataEndRow Then
            If GetIndentLevel(ws.Cells(r + 1, "B").value) > parentIndent Then
                parentValue = currentValue
                childSum = 0
                childCount = 0
                k = r + 1
                
                ' 자식 항목 탐색: 들여쓰기가 부모보다 깊은 모든 하위 항목을 찾음
                Do While k <= dataEndRow
                    If GetIndentLevel(ws.Cells(k, "B").value) > parentIndent Then
                        childSum = childSum + val(ws.Cells(k, "C").value)
                        childCount = childCount + 1
                        k = k + 1
                    Else
                        ' 들여쓰기가 같거나 얕아지면 자식 블록이 끝난 것으로 간주
                        Exit Do
                    End If
                Loop
                
                ' 자식 항목이 있고, 부모 값과 자식 합계가 일치하는지 확인
                If childCount > 0 And parentValue > 0 And parentValue = childSum Then
                    isMatchedParent = True ' 합계에서 제외할 부모로 플래그 설정
                    ws.Range("A" & r & ":D" & r).Interior.color = RGB(204, 255, 204) ' 연녹색
                    ws.Cells(r, "D").value = "하위 항목과 일치"
                End If
            End If
        End If
        
        ' 합계가 일치하는 부모 항목이 아닐 경우에만 최종 합계(calcTotal)에 더함
        If Not isMatchedParent Then
            calcTotal = calcTotal + currentValue
        End If

        ' 짝수 행 배경색 적용
        If ws.Range("A" & r & ":D" & r).Interior.ColorIndex = xlNone Then
            If IsNumeric(ws.Cells(r, "A").value) And ws.Cells(r, "A").value Mod 2 = 0 Then
                ws.Range("A" & r & ":D" & r).Interior.color = RGB(242, 242, 242) ' 연회색
            End If
        End If
    Next r

    ' --- 6. 요약 정보 출력 ---
    outRow = dataEndRow + 2
    
    ws.Cells(outRow, "B").value = "합계(계산)"
    ws.Cells(outRow, "C").value = calcTotal
    ws.Range("A" & outRow & ":D" & outRow).Interior.color = RGB(255, 255, 204)
    ws.rows(outRow).Font.Bold = True
    outRow = outRow + 1
    
    If budgetAmount <> 0 Then
        ws.Cells(outRow, "B").value = "예산액/총액"
        ws.Cells(outRow, "C").value = budgetAmount
        ws.Range("A" & outRow & ":D" & outRow).Interior.color = RGB(255, 255, 204)
        ws.rows(outRow).Font.Bold = True
        outRow = outRow + 1
        
        ws.Cells(outRow, "B").value = "차이(예산-계산)"
        ws.Cells(outRow, "C").value = budgetAmount - calcTotal
        ws.Range("A" & outRow & ":D" & outRow).Interior.color = RGB(255, 204, 204)
        ws.rows(outRow).Font.Bold = True
    End If
    
    ' =================================================================================
    '                   ↓↓↓ [요청사항 반영] 단위 선택 로직 수정 ↓↓↓
    ' =================================================================================
    Dim unitChoice As VbMsgBoxResult
    Dim displayFormat As String
    lastRow = ws.Cells(ws.rows.Count, "B").End(xlUp).row
    
    ' 메시지 박스 문구 수정
    unitChoice = UI_ChoiceToMsg(UI_Choose("완료 및 표시 단위 선택", "계산이 완료되었습니다. 결과 표시 단위를 선택하세요.", Array("원 단위", "천원 단위")))
    
    If unitChoice <> vbNo Then ' 원 단위 (폼을 취소해도 원 단위 유지)
        displayFormat = "#,##0"
    Else ' [아니요] 클릭 시: 천원 단위로 표시
        displayFormat = "#,##0""천원"""
        
        ' 데이터 및 요약 값들을 1000으로 나눔
        If lastRow >= 3 Then
            For r = 3 To lastRow
                If IsNumeric(ws.Cells(r, "C").value) And ws.Cells(r, "C").value <> "" Then
                    If ws.Cells(r, "C").value <> 0 Then
                        ws.Cells(r, "C").value = ws.Cells(r, "C").value / 1000
                    End If
                End If
            Next r
        End If
    End If
    
    ' C열 전체에 서식 일괄 적용
    If lastRow >= 3 Then
        ws.Range("C3:C" & lastRow).NumberFormat = displayFormat
    End If
    ' =================================================================================
    '                   ↑↑↑ [수정 완료] ↑↑↑
    ' =================================================================================
    
    GoTo FORMAT_END

NO_DATA:
    MsgBox "계산할 항목을 찾지 못했습니다.", vbInformation

FORMAT_END:
    ' 전체 테두리 및 열 너비 자동 맞춤
    lastRow = ws.Cells(ws.rows.Count, "B").End(xlUp).row
    ws.Columns("A:D").AutoFit
    If lastRow >= 2 Then
        With ws.Range("A2:D" & lastRow)
            .Borders.LineStyle = xlContinuous
            .Borders.color = vbBlack
        End With
    End If
    
    ' 별도의 완료 메시지 박스 제거
    ' MsgBox "완료!", vbInformation

SAFE_EXIT:
    ' --- 7. 종료 처리 ---
    If Err.Number <> 0 Then MsgBox "한글계산 중 오류가 나서 중단했습니다." & vbCrLf & Err.Description & vbCrLf & vbCrLf & "복사한 내용의 형식을 확인해 주세요.", vbExclamation, "한글계산"
    On Error Resume Next
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = prevCalc
    End With
End Sub


' =================================================================================
' 도우미 함수 (Helper Functions)
' =================================================================================

' [새 함수] 텍스트 앞쪽의 공백(들여쓰기) 수를 세는 함수
Private Function GetIndentLevel(s As String) As Integer
    GetIndentLevel = Len(s) - Len(LTrim(s))
End Function

' B열의 '천원' 항목들을 합산하는 함수 (곱하기 없을 때)
Private Function SumCheonwonValues(ByVal sourceString As String) As Double
    Dim processedString As String, totalSum As Double
    Dim parts As Variant, numPart As Variant, numStr As String
    
    processedString = sourceString
    
    Do While InStr(1, processedString, "(") > 0
        Dim startPos As Integer, endPos As Integer
        startPos = InStr(1, processedString, "(")
        endPos = InStr(startPos, processedString, ")")
        If endPos > startPos Then
            processedString = Left(processedString, startPos - 1) & Mid(processedString, endPos + 1)
        Else
            Exit Do
        End If
    Loop
    
    parts = Split(processedString, "천원")
    For Each numPart In parts
        numStr = Trim(numPart)
        If numStr <> "" Then
            numStr = Trim(Mid(numStr, InStrRev(numStr, " ") + 1))
            totalSum = totalSum + val(Replace(numStr, ",", ""))
        End If
    Next numPart
    
    SumCheonwonValues = totalSum
End Function

' 클립보드 텍스트를 가져오는 함수
Private Function GetClipboardText() As String
    Dim html As Object, dobj As Object
    On Error Resume Next
    Set html = CreateObject("htmlfile")
    GetClipboardText = html.parentWindow.clipboardData.GetData("Text")
    If Len(GetClipboardText) = 0 Then
        ' 예비: MSForms DataObject (htmlfile 이 막힌 환경)
        Set dobj = CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
        dobj.GetFromClipboard
        GetClipboardText = dobj.GetText(1)
    End If
End Function

' 결과 시트: 있으면 재사용, 없으면 생성
Private Function HC_ResultSheet(ByVal wb As Workbook) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets("한글계산")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.name = "한글계산"
    End If
    Set HC_ResultSheet = ws
End Function

' 특정 키워드가 포함된 라인을 건너뛸지 결정하는 함수
Private Function IsExcludedLine(ByVal s As String) As Boolean
    Dim tempS As String
    tempS = Trim(s)
    If InStr(1, tempS, "사업기간", vbTextCompare) > 0 Then IsExcludedLine = True: Exit Function
    If InStr(1, tempS, "지방교육재정교부금법", vbTextCompare) > 0 Then IsExcludedLine = True: Exit Function
    If InStr(1, tempS, "보통교부금 기준재정수요액", vbTextCompare) > 0 Then IsExcludedLine = True: Exit Function
    If InStr(1, tempS, "공약 ?", vbTextCompare) > 0 Then IsExcludedLine = True: Exit Function
    If InStr(1, tempS, "(교육부 직업교육정책과", vbTextCompare) > 0 Then IsExcludedLine = True: Exit Function
End Function

' 라인 앞의 번호(예: 1), 가))를 제거하는 함수
Private Function StripPrefixMarkers(ByVal s As String) As String
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "^\s*((\d+\s*[)\.])|([가-힣]\s*[)\.]))\s*"
    StripPrefixMarkers = re.Replace(s, "")
End Function

' 예산 또는 총액 관련 키워드가 있는지 확인하는 함수
Private Function IsBudgetOrTotalLine(ByVal s As String) As Boolean
    Dim tempS As String
    tempS = Trim(s)
    IsBudgetOrTotalLine = (InStr(1, tempS, "예산액", vbTextCompare) > 0) Or _
                          (InStr(1, tempS, "예 산 액", vbTextCompare) > 0) Or _
                          (InStr(1, tempS, "총액", vbTextCompare) > 0) Or _
                          (InStr(1, tempS, "총 액", vbTextCompare) > 0) Or _
                          (InStr(1, tempS, "합계", vbTextCompare) > 0) Or _
                          (InStr(1, tempS, "소계", vbTextCompare) > 0)   ' [수정] 합계·소계 줄이 항목으로 합산돼 총액이 두 배가 되던 문제
End Function

' 계산된 값에 '천원', '백만원' 단위를 적용하는 함수
Private Function ApplyUnitFromLine(ByVal s As String, ByVal v As Double) As Double
    If v = 0 Then
        ApplyUnitFromLine = 0
    ElseIf InStr(1, s, "백만원", vbTextCompare) > 0 Then
        ApplyUnitFromLine = v * 1000000#
    ElseIf InStr(1, s, "천원", vbTextCompare) > 0 Then
        ApplyUnitFromLine = v * 1000#
    Else
        ApplyUnitFromLine = v
    End If
End Function

' 라인에서 단위가 붙은 금액만 합산하는 함수
Private Function SumMoneyValuesWon(ByVal s As String) As Double
    Dim re As Object, ms As Object, m As Object
    Dim total As Double, num As Double, unitTxt As String
    
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.Pattern = "([0-9][0-9,\.]*)\s*(백만원|천원|원)"
    
    If re.Test(s) Then
        Set ms = re.Execute(s)
        For Each m In ms
            num = CDbl(Replace(m.SubMatches(0), ",", ""))
            unitTxt = CStr(m.SubMatches(1))
            Select Case unitTxt
                Case "백만원": total = total + (num * 1000000#)
                Case "천원":   total = total + (num * 1000#)
                Case "원":     total = total + num
            End Select
        Next m
    End If
    SumMoneyValuesWon = total
End Function

' 텍스트에서 첫 번째 숫자를 파싱하는 함수
Private Function ParseFirstNumberInText(ByVal s As String) As Double
    Dim re As Object, m As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "([0-9][0-9,\.]*)"
    If re.Test(s) Then
        Set m = re.Execute(s)(0)
        ParseFirstNumberInText = CDbl(Replace(m.value, ",", ""))
    Else
        ParseFirstNumberInText = 0
    End If
End Function

' "=" 우변의 값을 원 단위로 추출하는 함수
Private Function ExtractEqualsValueInWon(ByVal rhs As String) As Double
    Dim re As Object, m As Object
    Dim num As Double, unitTxt As String, v As Double
    
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "([0-9][0-9,\.]*)\s*(백만원|천원|원)?"
    
    If re.Test(rhs) Then
        Set m = re.Execute(rhs)(0)
        num = CDbl(Replace(m.SubMatches(0), ",", ""))
        If m.SubMatches.Count >= 2 Then unitTxt = CStr(m.SubMatches(1)) Else unitTxt = ""
        
        Select Case unitTxt
            Case "백만원": v = num * 1000000#
            Case "천원":   v = num * 1000#
            Case Else:     v = num ' 단위 없으면 원으로 간주
        End Select
        ExtractEqualsValueInWon = v
    Else
        ExtractEqualsValueInWon = 0
    End If
End Function

' 곱하기/나누기 계산을 수행하는 함수
Private Function ComputeMulDivFromText(ByVal lhs As String) As Double
    Dim s As String, valSet As Boolean
    Dim re As Object, ms As Object, m As Object
    Dim op As String, seg As String
    Dim factor As Double, acc As Double
    
    s = Replace(Replace(Replace(Replace(lhs, "×", "*"), "X", "*"), "x", "*"), "÷", "/")
    
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.Pattern = "([*/])?([^*/]+)"
    
    If Not re.Test(s) Then Exit Function
    
    Set ms = re.Execute(s)
    acc = 1#
    valSet = False
    
    For Each m In ms
        op = m.SubMatches(0)
        If op = "" Then op = "*"
        seg = Trim$(m.SubMatches(1))
        
        factor = ExtractRelevantNumber(seg)
        If factor <> 0 Then
            If Not valSet Then
                acc = factor
                valSet = True
            Else
                If op = "*" Then
                    acc = acc * factor
                ElseIf op = "/" Then
                    If factor = 0 Then Exit Function
                    acc = acc / factor
                End If
            End If
        End If
    Next m
    
    If valSet Then ComputeMulDivFromText = acc
End Function

' 계산에 필요한 숫자(단위가 있거나 없는)를 우선순위에 따라 추출하는 함수
Private Function ExtractRelevantNumber(ByVal seg As String) As Double
    Dim re As Object, m As Object
    Set re = CreateObject("VBScript.RegExp")
    
    ' 1순위: 통화 단위
    re.Pattern = "([0-9][0-9,\.]*)\s*(백만원|천원|원)"
    If re.Test(seg) Then
        Set m = re.Execute(seg)(0)
        ExtractRelevantNumber = CDbl(Replace(m.SubMatches(0), ",", ""))
        Exit Function
    End If
    
    ' 2순위: 수량 단위
    re.Pattern = "([0-9][0-9,\.]*)\s*(개|대|면|g|kg|mg|ton|ℓ|ml|병|mm|m|cm|km|cc|권|부|식|상자|m2|명|월|건|m3|실|조|각|장|타|set|봉|매|통|롤|마|저항|벌|일|동|개소|학급|교|회|개반|시간|시수|종목|개원|점수|과정|기관|연구회|점|팀)"
    If re.Test(seg) Then
        Set m = re.Execute(seg)(0)
        ExtractRelevantNumber = CDbl(Replace(m.SubMatches(0), ",", ""))
        Exit Function
    End If
    
    ' 3순위: 단위가 없는 숫자 (가장 마지막 순위로 검사)
    re.Pattern = "([0-9][0-9,\.]*)"
    If re.Test(seg) Then
        Set m = re.Execute(seg)(0)
        ExtractRelevantNumber = CDbl(Replace(m.value, ",", ""))
        Exit Function
    End If
End Function

' 텍스트에 금액 단위가 있는지 확인하는 함수
Private Function HasMoneyUnit(ByVal s As String) As Boolean
    HasMoneyUnit = (InStr(1, s, "백만원", vbTextCompare) > 0) Or _
                   (InStr(1, s, "천원", vbTextCompare) > 0) Or _
                   (InStr(1, s, "원", vbTextCompare) > 0)
End Function

' 비고란에 텍스트를 추가하는 함수
Private Function AppendRemark(ByVal base As String, ByVal addMsg As String) As String
    If Len(base) = 0 Then
        AppendRemark = addMsg
    Else
        AppendRemark = base & "; " & addMsg
    End If
End Function

' 텍스트에 지정된 단위가 없는 숫자가 있는지 확인하는 함수
Private Function HasNumberWithoutUnit(ByVal s As String) As Boolean
    Dim re As Object
    Dim tempStr As String
    
    ' 괄호와 그 안의 내용 제거 (계산에 영향을 주지 않도록)
    tempStr = s
    Do While InStr(1, tempStr, "(") > 0
        Dim startPos As Integer, endPos As Integer
        startPos = InStr(1, tempStr, "(")
        endPos = InStr(startPos, tempStr, ")")
        If endPos > startPos Then
            tempStr = Left(tempStr, startPos - 1) & Mid(tempStr, endPos + 1)
        Else
            Exit Do
        End If
    Loop
    
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    
    ' 1. 먼저 모든 '유효한 단위가 붙은 숫자'를 텍스트에서 제거합니다.
    re.Pattern = "([0-9][0-9,\.]*)\s*(백만원|천원|원|개|대|면|g|kg|mg|ton|ℓ|ml|병|mm|m|cm|km|cc|권|부|식|상자|m2|명|월|건|m3|실|조|각|장|타|set|봉|매|통|롤|마|저항|벌|일|동|개소|학급|교|회|개반|시간|시수|종목|개원|점수|과정|기관|연구회|점|팀)"
    tempStr = re.Replace(tempStr, "")
    
    ' 2. 단위가 제거된 후에도 텍스트에 숫자가 남아있다면, 그 숫자는 '단위 없는 숫자'입니다.
    re.Pattern = "[0-9]"
    If re.Test(tempStr) Then
        HasNumberWithoutUnit = True
    Else
        HasNumberWithoutUnit = False
    End If
End Function


' 문자열 내 수식을 실제 수치 값으로 계산하는 함수
Private Function EvaluateFormulaInText(ByVal expr As String) As Double
    Dim tempExpr As String

    ' 곱하기·나누기 등 유사기호 통일
    tempExpr = expr
    tempExpr = Replace(tempExpr, "×", "*")
    tempExpr = Replace(tempExpr, "x", "*")
    tempExpr = Replace(tempExpr, "X", "*")
    tempExpr = Replace(tempExpr, "÷", "/")

    ' 원의 경우 금액숫자만 남기고, 단위 등 제거
    tempExpr = Replace(tempExpr, "원", "")
    tempExpr = Replace(tempExpr, ",", "") '콤마 제거

    ' 안전: 허용된 문자(숫자, 연산자, 괄호, 소수점, 공백)만 남기기
    Dim re As Object: Set re = CreateObject("VBScript.RegExp")
    re.Global = True: re.Pattern = "[^0-9\+\-\*/\.\(\) ]"
    tempExpr = re.Replace(tempExpr, "")

    ' Application.Evaluate 사용
    On Error GoTo EvalErr
    EvaluateFormulaInText = Application.Evaluate(tempExpr)
    Exit Function

EvalErr:
    EvaluateFormulaInText = 0
End Function

