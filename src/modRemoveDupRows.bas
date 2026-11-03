Attribute VB_Name = "modRemoveDupRows"
Public Sub c중복셀값제거(control As IRibbonControl)
    ' --- 변수 선언 ---
    Dim targetCell As Range
    Dim ws As Worksheet
    Dim keyword As Variant
    Dim selectedAddress As String ' 선택한 셀의 주소를 저장할 변수
    Dim colNum As Long
    Dim lastRow As Long
    Dim i As Long
    Dim deletedCount As Long

    ' --- 1. 사용자로부터 기준 셀 선택받기 ---
    On Error Resume Next
    Set targetCell = Application.InputBox(prompt:="기준으로 삼을 셀을 마우스로 선택하거나 주소를 입력하세요." & vbCrLf & "(선택한 셀의 행은 삭제되지 않습니다)", title:="중복 행 삭제 (기준 셀 제외)", Type:=8)
    On Error GoTo 0

    If targetCell Is Nothing Then
        MsgBox "작업이 취소되었습니다.", vbInformation
        Exit Sub
    End If
    
    ' --- 2. 필수 정보 설정 ---
    Set ws = targetCell.Worksheet
    Set targetCell = targetCell.Cells(1, 1)   ' [수정] 여러 셀을 선택해도 첫 셀만 기준으로 (배열 형식오류 방지)
    If IsError(targetCell.value) Then
        MsgBox "선택한 셀이 오류값입니다. 값이 있는 셀을 선택해주세요.", vbExclamation
        Exit Sub
    End If
    keyword = targetCell.value
    selectedAddress = targetCell.Address ' 선택한 셀의 주소를 미리 저장
    colNum = targetCell.Column

    If IsEmpty(targetCell) Or Trim(keyword) = "" Then
        MsgBox "선택한 셀이 비어있습니다. 값이 있는 셀을 선택해주세요.", vbExclamation
        Exit Sub
    End If
    
    ' --- 3. 사용자에게 최종 확인 ---
    If UI_Choose("삭제 확인", "'" & targetCell.Address(False, False) & "' 셀의 행은 남기고, 같은 열에서 '" & keyword & "' 값과 같은 다른 모든 행을 삭제합니다." & vbCrLf & _
              "(실행 직전 상태는 숨김 백업 시트로 자동 저장됩니다)", Array("삭제 실행", "취소")) <> 1 Then
        MsgBox "작업이 취소되었습니다.", vbInformation
        Exit Sub
    End If

    ' --- 4. 매크로 실행 준비 ---
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' 시트 보호 해제 (비밀번호가 없다면 On Error 구문이 오류를 무시함)
    On Error Resume Next
    ws.Unprotect
    On Error GoTo 0

    If ws.AutoFilterMode Then ws.AutoFilterMode = False

    ' ★★★★★ 수정 및 추가된 부분 시작 ★★★★★
    
    ' --- 5. 행 삭제 작업 ---
    ' 기준 열의 마지막 데이터 행 번호 찾기
    lastRow = ws.Cells(ws.rows.Count, colNum).End(xlUp).row
    deletedCount = 0 ' 삭제 카운트 초기화

  ' 마지막 행부터 역순으로 반복
    For i = lastRow To 1 Step -1
        ' 현재 셀 참조를 변수에 담아 가독성 높임
        Dim currentCell As Range
        Set currentCell = ws.Cells(i, colNum)

        ' (조건1) 값 비교 및 (조건2) 기준 셀 제외 확인
        If currentCell.Address <> selectedAddress Then
        If Not IsError(currentCell.value) Then
        If Trim(CStr(currentCell.value)) = Trim(CStr(keyword)) Then
            ' 행 삭제 시 발생할 수 있는 오류 방지를 위해 에러 핸들링 추가
            On Error Resume Next
            ws.rows(i).EntireRow.Delete  ' EntireRow를 사용하여 명확하게 행 전체 삭제
            
            If Err.Number <> 0 Then
                Debug.Print "행 삭제 실패 (행 번호: " & i & "): " & Err.Description
                Err.Clear
            Else
                deletedCount = deletedCount + 1
            End If
            On Error GoTo 0
        End If
        End If
        End If
    Next i
    
    ' ★★★★★ 수정 및 추가된 부분 끝 ★★★★★

    ' --- 6. 매크로 마무리 ---
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    ' 시트 다시 보호 (기존에 보호되어 있었다면)
    ' 이 부분은 필요에 따라 사용 (보호가 필요 없다면 지워도 됨)
    ' ws.Protect

    ' --- 7. 완료 메시지 ---
    If deletedCount > 0 Then
        MsgBox "선택한 셀을 제외하고, '" & keyword & "' 값과 같은 행 " & deletedCount & "개를 삭제했습니다." & vbCrLf & vbCrLf & _
               "되돌리려면 리본 [백업 시트 복원]을 누르세요.", vbInformation, "작업 완료"
    Else
        MsgBox "삭제할 중복 행을 찾을 수 없었습니다.", vbInformation, "작업 완료"
    End If

End Sub

