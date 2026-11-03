Attribute VB_Name = "modTableCalc"
Option Explicit
Public TC_LastColInfo As String   ' [v2.6] 열 자동 인식 결과 설명

' 계산 방향을 저장하는 전역 변수 (F,G열 계산식이 고정되어 이 변수의 역할은 없습니다.)
Public IsDefaultCalculation As Boolean

' 표 마지막 총합 행을 위한 변수
Private Const TOTAL_ROW_OFFSET As Long = 2 ' 데이터 마지막 행에서 총합이 들어갈 행까지의 오프셋 (2는 빈줄 + 총합 라인)
' MAX_COLUMN_WIDTH는 AutoFit 후 조절하는 최대 너비입니다. 내용에 따라 조절이 필요할 수 있습니다.
Private Const MAX_COLUMN_WIDTH As Double = 40 ' 열 너비 자동 조절 시 최대 너비 (조절 가능, 이보다 넓어지지 않음)

' 매크로 이름 변경: ProcessTableData -> 표계산
Sub 표계산(control As IRibbonControl)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim currentRangeStartRow As Long ' 부분합 계산 시작 행
    Dim originalCValue As Double    ' C열 원본 값 (숫자로 변환)
    Dim originalEValue As Double    ' E열 원본 값 (숫자로 변환)
    Dim calculatedEValue As Double  ' D열 계산식으로 얻은 E열 값
    Dim expressionD As String       ' D열의 계산식 문자열
    Dim cleanedExpressionD As String ' 정제된 D열 계산식
    Dim resultD As Variant          ' D열 계산 결과
    Dim c As Range                  ' 컬럼 변수 선언

    ' [수정] 클립보드를 먼저 확인 (비어 있으면 시트를 건드리지 않음)
    If Not TC_ClipboardHasText() Then
        MsgBox "클립보드에 표 데이터가 없습니다." & vbCrLf & vbCrLf & _
               "한글에서 계산할 표를 블록 지정해 복사(Ctrl+C)한 뒤 [표계산]을 누르세요.", vbExclamation, "표계산"
        Exit Sub
    End If
    ' [수정] 결과는 활성 시트를 지우지 않고 '표계산' 시트에 기록
    Set ws = TC_ResultSheet(ActiveWorkbook)
    ws.Activate

    IsDefaultCalculation = True

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' 기존 내용 초기화 (A1부터 시작)
    ws.Cells.ClearContents
    ws.Cells.ClearFormats

    ' 헤더 설정 및 글꼴 지정
    With ws
        .Range("A1").value = "No"
        .Range("B1").value = "계산1"
        .Range("C1").value = "결과1"
        .Range("D1").value = "계산2"
        .Range("E1").value = "결과2"
        .Range("F1").value = "비고1"
        .Range("G1").value = "비고2"
        
        .Range("A1:G1").Font.name = "맑은 고딕"
        .Range("A1:G1").Font.Size = 12
        .Range("A1:G1").Font.Bold = True
        .Range("A1:G1").HorizontalAlignment = xlCenter ' 헤더 가운데 정렬
    End With

    ' 한글에서 복사한 내용 B2 셀부터 자동 붙여넣기 (엑셀과 한글 매크로 팀 개선 방식)
    On Error GoTo PasteError
    ws.Paste Destination:=ws.Range("B2")
    On Error GoTo 0
    
    ' [추가] 한글 표의 두 줄짜리 칸은 붙여넣기 때 두 행으로 갈라지고 옆 칸이 세로 병합됨 → 병합 해제 후 이어지는 행을 앞 행에 합침
    TC_DetectColumns ws
    TC_NormalizeRows ws

    ' 현재 시트의 데이터가 실제로 어디까지 있는지 확인 (B열 기준으로)
    lastRow = ws.Cells(rows.Count, "B").End(xlUp).row

    If lastRow < 2 Then ' B2셀에도 데이터가 없는 경우
        MsgBox "클립보드에 처리할 데이터가 없거나 올바른 형식의 표가 아닙니다. 한글에서 표를 복사하여 다시 시도해주세요.", vbInformation
        Application.ScreenUpdating = True
        Application.DisplayAlerts = True
        Exit Sub
    End If
    
    ' 붙여넣기 된 모든 데이터 범위에 대해 글꼴 설정 및 굵게 해제
    ' HWP에서 복사된 데이터가 B열부터 시작된다고 가정
    Dim dataRange As Range
    ' B열부터 G열까지 붙여넣은 모든 데이터 범위
    Set dataRange = ws.Range(ws.Cells(2, "B"), ws.Cells(lastRow, "G"))
    With dataRange.Font
        .name = "맑은 고딕"
        .Size = 12 ' 크기를 12로 명시
        .Bold = False ' 붙여넣기된 내용의 굵게 서식 해제
    End With
    
    currentRangeStartRow = 2 ' 첫 부분합 범위 시작 행

    For i = 2 To lastRow ' B2부터 원래 데이터의 마지막 행까지 반복
        ' A열: No (순번)
        ws.Cells(i, "A").value = i - 1 ' 1부터 시작하는 순번
        ws.Cells(i, "A").Font.name = "맑은 고딕"
        ws.Cells(i, "A").Font.Size = 12
        ws.Cells(i, "A").HorizontalAlignment = xlCenter ' 순번 가운데 정렬

        ' C열, E열의 원본 값을 가져오기 (문자열)
        Dim cellCValue As String
        Dim cellEValue As String
        cellCValue = CStr(ws.Cells(i, "C").value)
        cellEValue = CStr(ws.Cells(i, "E").value)

        ' 빈칸 처리 (붙여넣기 시 빈칸은 이미 빈칸으로 유지됨)
        On Error Resume Next ' 숫자로 변환 불가 시 오류 방지
        originalCValue = 0
        If Not IsEmpty(cellCValue) And IsNumeric(cellCValue) Then
            originalCValue = CDbl(cellCValue) ' C열의 원본 1차 계산값 (검증 대상)
        End If
        Err.Clear

        originalEValue = 0
        If Not IsEmpty(cellEValue) And IsNumeric(cellEValue) Then
            originalEValue = CDbl(cellEValue) ' E열의 원본 2차 계산값 (검증 대상)
        End If
        Err.Clear
        On Error GoTo 0

        ' D열의 계산식 문자열
        expressionD = CStr(ws.Cells(i, "D").value)
        
        ' [수정] D열 계산식 정제: ×/÷/X/x → 연산자, 콤마·단위·항목명(한글 등) 제거 → 숫자와 연산자만 남김 (TC_CleanExpr)
        cleanedExpressionD = TC_CleanExpr(expressionD)

        ' D열 계산식 평가 및 E열에 계산된 값 넣기
        calculatedEValue = 0 ' 기본값 0으로 초기화
        On Error Resume Next ' 계산식 오류 시 처리
        If cleanedExpressionD <> "" Then ' 정제된 계산식이 비어있지 않은 경우에만 계산 시도
            resultD = Application.Evaluate(cleanedExpressionD)
            If Not IsError(resultD) And IsNumeric(resultD) Then ' 계산 성공 및 숫자일 경우
                calculatedEValue = CDbl(resultD)
            End If
        End If
        On Error GoTo 0 ' 에러 처리 비활성화 해제
        ' [추가] 산출기초는 원, 예산 칸은 천원인 표(단위: 천원) 자동 인식 → 계산값을 천원으로 환산
        If originalEValue <> 0 And calculatedEValue <> 0 Then
            If Abs(calculatedEValue / originalEValue) > 500 And Abs(calculatedEValue / originalEValue) < 2000 Then
                calculatedEValue = calculatedEValue / 1000
            End If
        End If

        ' B열에 '부분합' 텍스트가 명시적으로 있는 경우 부분합 행으로 간주 (사용자 원본 HWP 파일에서 '부분합' 텍스트가 있는 경우)
        If LCase(Trim(CStr(ws.Cells(i, "B").value))) = "부분합" Then
            ' 부분합 행 처리
            If i > currentRangeStartRow Then ' 첫 부분합 행이 아니면
                ' 부분합은 '현재 for 루프의 i 행'에 계산값을 채웁니다.
                ' 이 함수 안에서 B열의 "부분합" 텍스트를 제거하고 F, G열도 계산합니다.
                Call CalculateAndSetSubtotal(ws, currentRangeStartRow, i - 1, i)
            End If
            
            ' B열을 비우고 배경색만 남김 (사용자 요청 반영)
            ws.Cells(i, "B").value = ""
            ws.Cells(i, "B").Font.Bold = False ' 혹시 모를 굵게 처리 제거
            ws.Cells(i, "B").HorizontalAlignment = xlCenter
            ws.Range(ws.Cells(i, "A"), ws.Cells(i, "G")).Interior.color = RGB(220, 230, 241) ' 부분합 행 배경색 지정
            
            ' 부분합 행에서는 No와 F, G열은 공백으로
            ws.Cells(i, "A").value = ""
            ws.Cells(i, "F").value = ""
            ws.Cells(i, "G").value = ""

            currentRangeStartRow = i + 1 ' 다음 부분합 범위 시작 행 설정
        Else ' 일반 데이터 행 처리
            ' E열: 엑셀 계산값과 한글에서 복사해온 값 검증 및 표시
            ws.Cells(i, "E").value = calculatedEValue ' 계산된 값을 먼저 삽입
            If originalEValue <> calculatedEValue Then ' 한글 값과 엑셀 계산값이 다를 경우
                ws.Cells(i, "E").Font.color = vbRed ' 글자색 빨간색
            Else
                ws.Cells(i, "E").Font.color = vbBlack ' 일치하면 검은색
            End If
            ws.Cells(i, "E").Font.name = "맑은 고딕"
            ws.Cells(i, "E").Font.Size = 12
            ' 0은 표시하지 않음; 음수는 △기호로
            If calculatedEValue <> 0 Then
                If calculatedEValue < 0 Then
                    ws.Cells(i, "E").NumberFormat = "#,##0_ ;" & ChrW(9651) & "#,##0;;" ' △ (U+25B3) 기호 사용
                Else
                    ws.Cells(i, "E").NumberFormat = "#,##0_ ;[Red]-#,##0;;"
                End If
            Else
                ws.Cells(i, "E").NumberFormat = ";;;" ' 0일 때 아무것도 표시 안함
            End If


            ' C열도 음수는 △기호로 표시, 0은 표시 안함
            ws.Cells(i, "C").Font.name = "맑은 고딕"
            ws.Cells(i, "C").Font.Size = 12
            If originalCValue <> 0 Then
                If originalCValue < 0 Then
                    ws.Cells(i, "C").NumberFormat = "#,##0_ ;" & ChrW(9651) & "#,##0;;" ' △ (U+25B3) 기호 사용
                Else
                    ws.Cells(i, "C").NumberFormat = "#,##0_ ;[Red]-#,##0;;"
                End If
            Else
                ws.Cells(i, "C").NumberFormat = ";;;"
            End If


            ' F열: 비고1 (C2-E2로 고정)
            Dim valF As Double
            valF = originalCValue - calculatedEValue
            ws.Cells(i, "F").value = valF
            ws.Cells(i, "F").Font.name = "맑은 고딕"
            ws.Cells(i, "F").Font.Size = 12
            If valF <> 0 Then
                If valF < 0 Then
                    ws.Cells(i, "F").NumberFormat = "#,##0_ ;" & ChrW(9651) & "#,##0;;" ' △ (U+25B3) 기호 사용
                Else
                    ws.Cells(i, "F").NumberFormat = "#,##0_ ;[Red]-#,##0;;"
                End If
            Else
                ws.Cells(i, "F").NumberFormat = ";;;"
            End If


            ' G열: 비고2 (F2/E2*100으로 고정) - % 표시 제외, 음수 △기호 표시
            Dim valG As Double
            If calculatedEValue = 0 Then ' E열이 0이면 계산 불가
                ws.Cells(i, "G").value = "" ' 빈칸으로
            Else
                valG = (valF / calculatedEValue) * 100
                If valG <> 0 Then
                    ws.Cells(i, "G").value = valG ' % 기호 없이 순수 숫자 값으로
                    ' 음수 △기호 표시, 0은 표시 안함
                    If valG < 0 Then
                        ws.Cells(i, "G").NumberFormat = "#,##0.0_ ;" & ChrW(9651) & "#,##0.0;;" ' △ (U+25B3) 기호 사용
                    Else
                        ws.Cells(i, "G").NumberFormat = "#,##0.0_ ;[Red]-#,##0.0;;"
                    End If
                Else
                    ws.Cells(i, "G").NumberFormat = ";;;" ' 0일 때 아무것도 표시 안함
                End If
            End If
            ws.Cells(i, "G").Font.name = "맑은 고딕"
            ws.Cells(i, "G").Font.Size = 12
        End If ' End If for 부분합 행 vs 일반 데이터 행
    Next i

    ' 마지막 부분합 계산 (남은 데이터가 있다면)
    ' 이 시점의 lastRow는 원래 HWP에서 복사된 데이터의 마지막 행입니다.
    Dim lastSubtotalRow_for_subtotal_total As Long ' 부분합 총합 계산 시 제외할 마지막 부분합 행
    lastSubtotalRow_for_subtotal_total = 0 ' 초기화

    If LCase(Trim(CStr(ws.Cells(lastRow, "B").value))) <> "부분합" Then
         ' CalculateAndSetSubtotal은 마지막 부분합을 'lastRow + 1' 위치에 삽입합니다.
         Call CalculateAndSetSubtotal(ws, currentRangeStartRow, lastRow, lastRow + 1)
         lastSubtotalRow_for_subtotal_total = lastRow + 1 ' 마지막으로 추가된 부분합 행의 번호
         lastRow = lastRow + 1 ' 메인 처리 범위의 마지막 행을 업데이트
    End If
    
    ' 총합계 계산 (전체 표의 총합)
    ' lastRow는 현재 마지막 데이터 또는 부분합이 있는 행을 가리킵니다.
    Call CalculateAndSetGrandTotal(ws, lastRow, lastSubtotalRow_for_subtotal_total)

    ' 열 너비 자동 조정 및 최대 너비 제한 (데이터 처리 후 가장 마지막에 수행)
    For Each c In ws.UsedRange.Columns
        c.WrapText = False
        c.EntireColumn.AutoFit
        
        If c.ColumnWidth > MAX_COLUMN_WIDTH Then
            c.ColumnWidth = MAX_COLUMN_WIDTH
        End If
    Next c

    Application.DisplayAlerts = True ' 경고 메시지 표시 재활성화
    Application.ScreenUpdating = True ' 화면 업데이트 재활성화
    If Not HCM_Quiet Then MsgBox "한글 표 데이터 처리 및 계산이 완료되었습니다!" & IIf(Len(TC_LastColInfo) > 0, vbCrLf & vbCrLf & TC_LastColInfo, ""), vbInformation
    Exit Sub

PasteError:
    Application.DisplayAlerts = True ' 오류 발생 시 경고 메시지 활성화
    Application.ScreenUpdating = True
    MsgBox "클립보드에 엑셀에 붙여넣을 수 있는 표 형식의 데이터가 없습니다. 한글에서 표를 정확히 복사했는지 확인해주세요.", vbCritical
End Sub

' 계산 방향을 전환하는 매크로 (F,G열 계산식이 고정되었으므로 이 버튼은 기능 없음)
Sub ToggleCalculationDirection()
    MsgBox "F, G열의 계산식이 고정되어 있어 계산 방향 전환 기능은 동작하지 않습니다.", vbInformation
End Sub

' 부분합을 계산하여 해당 부분합 행에 입력하는 함수
Private Sub CalculateAndSetSubtotal(ByVal ws As Worksheet, ByVal startRow As Long, ByVal endRow As Long, ByVal subtotalDisplayRow As Long)
    Dim sumC As Double
    Dim sumE As Double
    Dim j As Long

    ' 범위 내의 C열과 E열 합계 계산
    For j = startRow To endRow
        ' 부분합/총합계 행이나 빈칸/비숫자 값은 합계에 포함하지 않습니다. (이중합산 방지)
        ' isSubtotalRow를 명시적으로 체크하는 것이 안전
        Dim isCurrentRowSubtotal As Boolean
        isCurrentRowSubtotal = (LCase(Trim(CStr(ws.Cells(j, "B").value))) = "") And _
                               (ws.Cells(j, "A").Interior.color = RGB(220, 230, 241)) ' B열이 공백이고 배경색이 있다면 부분합 행으로 간주

        If Not isCurrentRowSubtotal And Not IsEmpty(ws.Cells(j, "C").value) And IsNumeric(ws.Cells(j, "C").value) Then
            sumC = sumC + CDbl(ws.Cells(j, "C").value)
        End If
        If Not isCurrentRowSubtotal And Not IsEmpty(ws.Cells(j, "E").value) And IsNumeric(ws.Cells(j, "E").value) Then
            sumE = sumE + CDbl(ws.Cells(j, "E").value)
        End If
    Next j
    
    ' 부분합 행에 결과 설정 (B, C, E열만)
    ' A열 No 비움
    ws.Cells(subtotalDisplayRow, "A").value = ""
    ' B열은 공백으로 유지 (사용자 요청 반영)
    ws.Cells(subtotalDisplayRow, "B").value = ""
    ws.Cells(subtotalDisplayRow, "B").Font.Bold = False ' 굵게 처리 해제
    ws.Cells(subtotalDisplayRow, "B").HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(subtotalDisplayRow, "A"), ws.Cells(subtotalDisplayRow, "G")).Interior.color = RGB(220, 230, 241) ' 부분합 행 배경색 지정

    ws.Cells(subtotalDisplayRow, "C").value = sumC
    ws.Cells(subtotalDisplayRow, "E").value = sumE
    
    ws.Cells(subtotalDisplayRow, "C").Font.Bold = True
    If sumC <> 0 Then
        If sumC < 0 Then
            ws.Cells(subtotalDisplayRow, "C").NumberFormat = "#,##0_ ;" & ChrW(9651) & "#,##0;;"
        Else
            ws.Cells(subtotalDisplayRow, "C").NumberFormat = "#,##0_ ;[Red]-#,##0;;"
        End If
    Else
        ws.Cells(subtotalDisplayRow, "C").NumberFormat = ";;;" ' 0일 때 아무것도 표시 안함
    End If

    ws.Cells(subtotalDisplayRow, "E").Font.Bold = True
    If sumE <> 0 Then
        If sumE < 0 Then
            ws.Cells(subtotalDisplayRow, "E").NumberFormat = "#,##0_ ;" & ChrW(9651) & "#,##0;;"
        Else
            ws.Cells(subtotalDisplayRow, "E").NumberFormat = "#,##0_ ;[Red]-#,##0;;"
        End If
    Else
        ws.Cells(subtotalDisplayRow, "E").NumberFormat = ";;;" ' 0일 때 아무것도 표시 안함
    End If

    ' 부분합 행의 F, G열도 계산하여 채움 (사용자 요청 반영)
    Dim valF_sub As Double
    valF_sub = sumC - sumE
    ws.Cells(subtotalDisplayRow, "F").value = valF_sub
    ws.Cells(subtotalDisplayRow, "F").Font.Bold = True
    If valF_sub <> 0 Then
        If valF_sub < 0 Then
            ws.Cells(subtotalDisplayRow, "F").NumberFormat = "#,##0_ ;" & ChrW(9651) & "#,##0;;"
        Else
            ws.Cells(subtotalDisplayRow, "F").NumberFormat = "#,##0_ ;[Red]-#,##0;;"
        End If
    Else
        ws.Cells(subtotalDisplayRow, "F").NumberFormat = ";;;"
    End If

    Dim valG_sub As Double
    If sumE = 0 Then
        ws.Cells(subtotalDisplayRow, "G").value = ""
    Else
        valG_sub = (valF_sub / sumE) * 100
        ws.Cells(subtotalDisplayRow, "G").value = valG_sub
        ws.Cells(subtotalDisplayRow, "G").Font.Bold = True
        If valG_sub <> 0 Then
            If valG_sub < 0 Then
                ws.Cells(subtotalDisplayRow, "G").NumberFormat = "#,##0.0_ ;" & ChrW(9651) & "#,##0.0;;"
            Else
                ws.Cells(subtotalDisplayRow, "G").NumberFormat = "#,##0.0_ ;[Red]-#,##0.0;;"
            End If
        Else
            ws.Cells(subtotalDisplayRow, "G").NumberFormat = ";;;"
        End If
    End If

End Sub

' 전체 총합을 계산하여 표 마지막에 입력하는 함수
Private Sub CalculateAndSetGrandTotal(ByVal ws As Worksheet, ByVal effectiveLastRow As Long, ByVal lastSubtotalRow_to_exclude As Long)
    ' effectiveLastRow: 현재 시트에 있는 데이터+부분합의 실제 마지막 행
    ' lastSubtotalRow_to_exclude: 부분합 총합 계산 시 제외할 마지막 부분합 행 번호 (0이면 제외할 행 없음)

    Dim grandTotalC As Double
    Dim grandTotalE As Double
    Dim subtotalTotalC As Double ' 부분합들의 총합계 (C열)
    Dim subtotalTotalE As Double ' 부분합들의 총합계 (E열)
    Dim i As Long
    Dim totalRow As Long
    
    ' -- 총합계 계산 (B열이 공백이 아닌 순수 데이터만 합산) --
    grandTotalC = 0: grandTotalE = 0
    subtotalTotalC = 0: subtotalTotalE = 0 ' 부분합 총합 초기화

    For i = 2 To effectiveLastRow ' 현재 시트의 모든 데이터+부분합 행을 순회
        Dim isCurrentRowSubtotal As Boolean
        isCurrentRowSubtotal = (LCase(Trim(CStr(ws.Cells(i, "B").value))) = "") And _
                               (ws.Cells(i, "A").Interior.color = RGB(220, 230, 241)) ' B열이 공백이고 배경색이 있다면 부분합 행으로 간주

        If Not isCurrentRowSubtotal And Not IsEmpty(ws.Cells(i, "C").value) And IsNumeric(ws.Cells(i, "C").value) Then
            ' 현재 행이 부분합 행이 아니고, 값이 비어있지 않고 숫자이면 grandTotal에 합산
            grandTotalC = grandTotalC + CDbl(ws.Cells(i, "C").value)
        End If
        If Not isCurrentRowSubtotal And Not IsEmpty(ws.Cells(i, "E").value) And IsNumeric(ws.Cells(i, "E").value) Then
            grandTotalE = grandTotalE + CDbl(ws.Cells(i, "E").value)
        End If

        ' 부분합 총합 계산 (마지막 부분합 제외)
        If isCurrentRowSubtotal And Not IsEmpty(ws.Cells(i, "C").value) And IsNumeric(ws.Cells(i, "C").value) Then
            If i <> lastSubtotalRow_to_exclude Then ' 현재 부분합 행이 제외할 마지막 부분합 행이 아니라면
                subtotalTotalC = subtotalTotalC + CDbl(ws.Cells(i, "C").value)
            End If
        End If
        If isCurrentRowSubtotal And Not IsEmpty(ws.Cells(i, "E").value) And IsNumeric(ws.Cells(i, "E").value) Then
            If i <> lastSubtotalRow_to_exclude Then ' 현재 부분합 행이 제외할 마지막 부분합 행이 아니라면
                subtotalTotalE = subtotalTotalE + CDbl(ws.Cells(i, "E").value)
            End If
        End If
    Next i
    
    ' 총합계 행의 위치 계산
    totalRow = effectiveLastRow + TOTAL_ROW_OFFSET

    ' 총합계 표시
    With ws.Cells(totalRow, "A")
        .value = "총합계"
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Interior.color = RGB(255, 255, 153) ' 노란색 배경
    End With
    ws.Cells(totalRow, "C").value = grandTotalC
    ws.Cells(totalRow, "E").value = grandTotalE
    ' 총합계 비고1 (C - E) 계산
    Dim grandTotalF As Double
    grandTotalF = grandTotalC - grandTotalE
    ws.Cells(totalRow, "F").value = grandTotalF

    ws.Cells(totalRow, "C").Font.Bold = True
    If grandTotalC <> 0 Then
        If grandTotalC < 0 Then
            ws.Cells(totalRow, "C").NumberFormat = "#,##0_ ;" & ChrW(9651) & "#,##0;;"
        Else
            ws.Cells(totalRow, "C").NumberFormat = "#,##0_ ;[Red]-#,##0;;"
        End If
    Else
        ws.Cells(totalRow, "C").NumberFormat = ";;;"
    End If


    ws.Cells(totalRow, "E").Font.Bold = True
    If grandTotalE <> 0 Then
        If grandTotalE < 0 Then
            ws.Cells(totalRow, "E").NumberFormat = "#,##0_ ;" & ChrW(9651) & "#,##0;;"
        Else
            ws.Cells(totalRow, "E").NumberFormat = "#,##0_ ;[Red]-#,##0;;"
        End If
    Else
        ws.Cells(totalRow, "E").NumberFormat = ";;;"
    End If

    ws.Cells(totalRow, "F").Font.Bold = True
    If grandTotalF <> 0 Then
        If grandTotalF < 0 Then
            ws.Cells(totalRow, "F").NumberFormat = "#,##0_ ;" & ChrW(9651) & "#,##0;;"
        Else
            ws.Cells(totalRow, "F").NumberFormat = "#,##0_ ;[Red]-#,##0;;"
        End If
    Else
        ws.Cells(totalRow, "F").NumberFormat = ";;;"
    End If

    ' 총합계 비고2 (F/E * 100) - % 표시 제외, 음수 △기호 표시
    If grandTotalE = 0 Then
        ws.Cells(totalRow, "G").value = ""
    Else
        Dim grandTotalG As Double
        grandTotalG = (grandTotalF / grandTotalE) * 100
        If grandTotalG <> 0 Then
            ws.Cells(totalRow, "G").value = grandTotalG
            If grandTotalG < 0 Then
                ws.Cells(totalRow, "G").NumberFormat = "#,##0.0_ ;" & ChrW(9651) & "#,##0.0;;"
            Else
                ws.Cells(totalRow, "G").NumberFormat = "#,##0.0_ ;[Red]-#,##0.0;;"
            End If
        Else
            ws.Cells(totalRow, "G").NumberFormat = ";;;"
        End If
    End If
    ws.Cells(totalRow, "G").Font.Bold = True


    ' 부분합의 총합계 (따로 표시)
    With ws.Cells(totalRow + 1, "A")
        .value = "부분합 총합"
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Interior.color = RGB(255, 230, 204) ' 연한 오렌지 배경
    End With

    ws.Cells(totalRow + 1, "C").value = subtotalTotalC
    ws.Cells(totalRow + 1, "C").Font.Bold = True
    If subtotalTotalC <> 0 Then
        If subtotalTotalC < 0 Then
            ws.Cells(totalRow + 1, "C").NumberFormat = "#,##0_ ;" & ChrW(9651) & "#,##0;;"
        Else
            ws.Cells(totalRow + 1, "C").NumberFormat = "#,##0_ ;[Red]-#,##0;;"
        End If
    Else
        ws.Cells(totalRow + 1, "C").NumberFormat = ";;;"
    End If

    ws.Cells(totalRow + 1, "E").value = subtotalTotalE
    ws.Cells(totalRow + 1, "E").Font.Bold = True
    If subtotalTotalE <> 0 Then
        If subtotalTotalE < 0 Then
            ws.Cells(totalRow + 1, "E").NumberFormat = "#,##0_ ;" & ChrW(9651) & "#,##0;;"
        Else
            ws.Cells(totalRow + 1, "E").NumberFormat = "#,##0_ ;[Red]-#,##0;;"
        End If
    Else
        ws.Cells(totalRow + 1, "E").NumberFormat = ";;;"
    End If
End Sub

' 클립보드에 텍스트가 있는지 (htmlfile → 예비 DataObject)
Private Function TC_ClipboardHasText() As Boolean
    Dim html As Object, dobj As Object, T As String
    On Error Resume Next
    Set html = CreateObject("htmlfile")
    T = html.parentWindow.clipboardData.GetData("Text")
    If Len(T) = 0 Then
        Set dobj = CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
        dobj.GetFromClipboard
        T = dobj.GetText(1)
    End If
    TC_ClipboardHasText = (Len(Trim$(T)) > 0)
End Function

' 결과 시트: 있으면 재사용, 없으면 생성
Private Function TC_ResultSheet(ByVal wb As Workbook) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets("표계산")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.name = "표계산"
    End If
    Set TC_ResultSheet = ws
End Function

' 계산식 정제: ×/÷/X/x → 연산자, 콤마 제거, 숫자·연산자·괄호·소수점 외 문자(단위·항목명·줄바꿈)는 모두 제거
Private Function TC_CleanExpr(ByVal s As String) As String
    Dim i As Long, ch As String, r As String
    s = Replace(s, ",", "")
    s = Replace(s, ChrW(215), "*")
    s = Replace(s, ChrW(247), "/")
    s = Replace(s, "X", "*")
    s = Replace(s, "x", "*")
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If (ch >= "0" And ch <= "9") Or ch = "*" Or ch = "/" Or ch = "+" Or ch = "-" Or ch = "(" Or ch = ")" Or ch = "." Then
            r = r & ch
        End If
    Next i
    Do While Len(r) > 0
        ch = Right$(r, 1)
        If ch = "*" Or ch = "/" Or ch = "+" Or ch = "-" Or ch = "." Or ch = "(" Then r = Left$(r, Len(r) - 1) Else Exit Do
    Loop
    Do While Len(r) > 0
        ch = Left$(r, 1)
        If ch = "*" Or ch = "/" Or ch = "+" Or ch = "." Or ch = ")" Then r = Mid$(r, 2) Else Exit Do
    Loop
    TC_CleanExpr = r
End Function

' 붙여넣은 표 정규화: 병합 해제 → 결과1(C)·결과2(E)가 모두 빈 행은 앞 행의 계산식(B·D)에 이어 붙이고 삭제
Private Sub TC_NormalizeRows(ByVal ws As Worksheet)
    Dim lastRow As Long, i As Long, bTxt As String, dTxt As String
    On Error Resume Next
    ws.UsedRange.UnMerge
    ws.UsedRange.WrapText = False
    On Error GoTo 0
    lastRow = ws.Cells(ws.rows.Count, "B").End(xlUp).row
    If ws.Cells(ws.rows.Count, "D").End(xlUp).row > lastRow Then lastRow = ws.Cells(ws.rows.Count, "D").End(xlUp).row
    For i = lastRow To 3 Step -1
        If Len(Trim$(CStr(ws.Cells(i, "C").value))) = 0 And Len(Trim$(CStr(ws.Cells(i, "E").value))) = 0 Then
            bTxt = Trim$(CStr(ws.Cells(i, "B").value)): dTxt = Trim$(CStr(ws.Cells(i, "D").value))
            If Len(bTxt) > 0 Or Len(dTxt) > 0 Then
                If Len(bTxt) > 0 Then ws.Cells(i - 1, "B").value = Trim$(CStr(ws.Cells(i - 1, "B").value) & " " & bTxt)
                If Len(dTxt) > 0 Then ws.Cells(i - 1, "D").value = Trim$(CStr(ws.Cells(i - 1, "D").value) & " " & dTxt)
                ws.rows(i).Delete
            End If
        End If
    Next i
End Sub


' =================================================================================
' [v2.6] 붙여넣은 표에서 계산식 열과 금액 열을 자동으로 찾아 B~E(계산1·결과1·계산2·결과2)로 재배치
'   - 앞에 항목명 열, 뒤에 '올해-전년도 차이' 열이 있어도 됨 (버림)
'   - 계산식·금액 쌍이 1개뿐이면 D·E 는 비움. 머리글 행(금액 칸이 글자)은 제거
'   - 열 판별: 숫자(단위 포함) 60% 이상 → 결과 열, 연산자(×·*·÷·/·+·=)나 '원'이 든 글 30% 이상 → 계산식 열
' =================================================================================
Private Sub TC_DetectColumns(ByVal ws As Worksheet)
    Dim lastRow As Long, lastCol As Long, c As Long, r As Long, s As String
    Dim nNon() As Long, nNum() As Long, nExpr() As Long, cls() As String
    Dim p1 As Long, p2 As Long, vals As Variant, outv() As Variant, nR As Long, hdrRows As Long
    TC_LastColInfo = ""
    On Error Resume Next
    ws.UsedRange.UnMerge
    On Error GoTo 0
    lastRow = ws.UsedRange.row + ws.UsedRange.rows.Count - 1
    lastCol = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1
    If lastRow < 2 Or lastCol < 3 Then Exit Sub
    ReDim nNon(2 To lastCol): ReDim nNum(2 To lastCol): ReDim nExpr(2 To lastCol): ReDim cls(2 To lastCol)
    vals = ws.Range(ws.Cells(2, 2), ws.Cells(lastRow, lastCol)).Value2
    nR = UBound(vals, 1)
    For c = 2 To lastCol
        For r = 1 To nR
            If Not IsError(vals(r, c - 1)) Then
                s = Trim$(CStr(vals(r, c - 1)))
                If Len(s) > 0 Then
                    nNon(c) = nNon(c) + 1
                    If TC_IsNumText(s) Then
                        nNum(c) = nNum(c) + 1
                    ElseIf TC_LooksExpr(s) Then
                        nExpr(c) = nExpr(c) + 1
                    End If
                End If
            End If
        Next r
        If nNon(c) = 0 Then
            cls(c) = "empty"
        ElseIf nNum(c) >= nNon(c) * 0.6 Then
            cls(c) = "num"
        ElseIf nExpr(c) >= nNon(c) * 0.3 Then
            cls(c) = "expr"
        Else
            cls(c) = "text"
        End If
    Next c
    p1 = TC_FindPair(cls, 2, lastCol, False)
    If p1 = 0 Then p1 = TC_FindPair(cls, 2, lastCol, True)
    If p1 = 0 Then Exit Sub
    p2 = TC_FindPair(cls, p1 + 2, lastCol, False)
    If p2 = 0 Then p2 = TC_FindPair(cls, p1 + 2, lastCol, True)

    ReDim outv(1 To nR, 1 To 4)
    For r = 1 To nR
        outv(r, 1) = vals(r, p1 - 1): outv(r, 2) = vals(r, p1)
        If p2 > 0 Then outv(r, 3) = vals(r, p2 - 1): outv(r, 4) = vals(r, p2)
    Next r
    ws.Range(ws.Cells(2, 2), ws.Cells(lastRow, lastCol)).Clear
    ws.Range(ws.Cells(2, 2), ws.Cells(1 + nR, 5)).value = outv
    ' 머리글 행 제거: 결과1 칸이 숫자가 아닌 글자이고 계산1 칸도 계산식이 아니면 머리글
    Do While hdrRows < 3
        s = Trim$(CStr(ws.Cells(2, "C").value))
        If Len(s) > 0 And Not TC_IsNumText(s) And Not TC_LooksExpr(CStr(ws.Cells(2, "B").value)) Then
            ws.rows(2).Delete
            hdrRows = hdrRows + 1
        Else
            Exit Do
        End If
    Loop
    TC_LastColInfo = "열 자동 인식: 복사한 표의 " & (p1 - 1) & "·" & p1 & "번째 열 → 계산1·결과1"
    If p2 > 0 Then
        TC_LastColInfo = TC_LastColInfo & ",  " & (p2 - 1) & "·" & p2 & "번째 열 → 계산2·결과2"
    Else
        TC_LastColInfo = TC_LastColInfo & "  (두 번째 계산식·금액 쌍 없음)"
    End If
    If lastCol - 1 > IIf(p2 > 0, 4, 2) Then TC_LastColInfo = TC_LastColInfo & vbCrLf & "그 밖의 열(항목명 · 차이 등)은 계산에서 제외했습니다."
    If hdrRows > 0 Then TC_LastColInfo = TC_LastColInfo & vbCrLf & "머리글 " & hdrRows & "행 제거"
End Sub

' 계산식 열(expr, allowText 면 글자 열도) 바로 오른쪽이 숫자 열인 첫 위치
Private Function TC_FindPair(ByRef cls() As String, ByVal fromC As Long, ByVal lastCol As Long, ByVal allowText As Boolean) As Long
    Dim c As Long
    For c = fromC To lastCol - 1
        If (cls(c) = "expr" Or (allowText And cls(c) = "text")) And cls(c + 1) = "num" Then
            TC_FindPair = c
            Exit Function
        End If
    Next c
End Function

' "1,500", "1,500원", "△300", "-2,000천원" 처럼 숫자(+단위)만인 글자인지
Private Function TC_IsNumText(ByVal s As String) As Boolean
    Dim T As String, ch As String
    T = Replace(Replace(Replace(s, ",", ""), " ", ""), ChrW(9651), "-")
    T = Replace(T, ChrW(9650), "-")
    Do While Len(T) > 0
        ch = Right$(T, 1)
        If (ch >= "0" And ch <= "9") Or ch = "." Or ch = ")" Or ch = "%" Then Exit Do
        T = Left$(T, Len(T) - 1)
    Loop
    T = Replace(T, "%", "")
    If Left$(T, 1) = "(" And Right$(T, 1) = ")" Then T = "-" & Mid$(T, 2, Len(T) - 2)
    If Len(T) = 0 Then Exit Function
    TC_IsNumText = IsNumeric(T)
End Function

' 숫자가 있고 연산자(×·*·÷·/·+·=·x)나 '원' 이 함께 있으면 계산식으로 본다
Private Function TC_LooksExpr(ByVal s As String) As Boolean
    Dim i As Long, hasDigit As Boolean, ch As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch >= "0" And ch <= "9" Then hasDigit = True: Exit For
    Next i
    If Not hasDigit Then Exit Function
    If InStr(s, ChrW(215)) > 0 Or InStr(s, "*") > 0 Or InStr(s, ChrW(247)) > 0 Or InStr(s, "/") > 0 Or InStr(s, "+") > 0 Or InStr(s, "=") > 0 Then TC_LooksExpr = True: Exit Function
    If InStr(s, "x") > 0 Or InStr(s, "X") > 0 Then TC_LooksExpr = True: Exit Function
    If InStr(s, "원") > 0 Then TC_LooksExpr = True
End Function

' 빌드 검증용: 활성 시트 B2 부터 붙어 있는 표로 열 인식만 실행 → 인식 결과 문자열
Public Function TC_TestDetect(ByVal ws As Worksheet) As String
    TC_DetectColumns ws
    TC_TestDetect = TC_LastColInfo
End Function


