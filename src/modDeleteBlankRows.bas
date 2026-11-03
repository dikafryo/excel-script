Attribute VB_Name = "modDeleteBlankRows"
Public Sub b빈칸삭제(control As IRibbonControl)
    ' --- 변수 선언 ---
    Dim targetRange As Range
    Dim ws As Worksheet
    Dim targetColNum As Long
    Dim startRow As Long
    Dim lastRow As Long
    Dim i As Long
    Dim userChoice As VbMsgBoxResult ' 사용자 선택(예, 아니요, 취소)을 저장할 변수
    Dim choiceMade As Boolean ' 사용자가 선택을 했는지 여부를 확인하는 변수
    Dim deleted As Long, kept As Long
    
    ' --- 1. 사용자로부터 기준 셀 선택받기 (마우스 선택 기능 추가) ---
    On Error Resume Next ' 사용자가 '취소'를 누를 경우를 대비
    Set targetRange = Application.InputBox(prompt:="기준으로 삼을 셀을 마우스로 선택하거나 주소를 입력하세요.", title:="빈 행 삭제 기준 선택", Type:=8)
    On Error GoTo 0 ' 오류 처리 원상 복구

    ' 사용자가 '취소'를 클릭했는지 확인
    If targetRange Is Nothing Then
        MsgBox "작업이 취소되었습니다.", vbInformation
        Exit Sub
    End If

    ' --- 2. 기준 정보 설정 ---
    Set ws = targetRange.Worksheet
    ' 선택된 영역 중 첫 번째 셀을 기준으로 설정
    Set targetRange = targetRange.Cells(1, 1)
    targetColNum = targetRange.Column
    startRow = targetRange.row
    lastRow = ws.Cells(ws.rows.Count, targetColNum).End(xlUp).row
    
    ' 삭제할 데이터가 없는 경우
    If lastRow < startRow Then
        MsgBox targetRange.Address(False, False) & " 셀 이후에는 데이터가 없어 작업을 진행하지 않습니다.", vbInformation
        Exit Sub
    End If
    
    ' --- 3. 빈 행 삭제 작업 ---
    Application.ScreenUpdating = False ' 화면 업데이트 중지로 속도 향상
    choiceMade = False ' 사용자 선택 초기화

    ' 마지막 행부터 역순으로 반복 (행 삭제 시 발생하는 오류 방지)
    For i = lastRow To startRow Step -1
        ' 기준 열의 셀이 비어있는지 확인
        ' [수정] 오류값(#N/A 등)은 빈칸이 아닌 것으로 취급 (Trim 형식 오류로 중단되던 문제)
        If Not IsError(ws.Cells(i, targetColNum).value) And Len(Trim$(CStr(ws.Cells(i, targetColNum).value))) = 0 Then
            ' 해당 행 전체에 데이터가 있는지 확인 (CountA: 비어있지 않은 셀의 개수)
            If Application.WorksheetFunction.CountA(ws.rows(i)) > 0 Then
                ' 행에 다른 데이터가 있고, 아직 사용자가 선택하지 않은 경우
                If Not choiceMade Then
                    userChoice = UI_ChoiceToMsg(UI_Choose("삭제 옵션 선택", "'" & ws.Cells(i, targetColNum).Address(False, False) & "' 셀이 비어 있지만 다른 열에 데이터가 있습니다. 어떻게 할까요?", Array("다른 데이터가 있어도 행 삭제", "완전히 빈 행만 삭제 (데이터 있는 행은 유지)")))
                    choiceMade = True
                End If

                Select Case userChoice
                    Case vbYes ' 1. 데이터가 있어도 삭제
                        ws.rows(i).Delete: deleted = deleted + 1
                    Case vbNo  ' 2. 데이터가 있는 행은 삭제 안함
                        kept = kept + 1
                    Case vbCancel ' 3. 취소
                        MsgBox "작업을 중단했습니다.", vbExclamation
                        GoTo Cleanup ' 매크로 종료 처리로 이동
                End Select
            Else
                ' 행 전체가 비어있는 경우, 묻지 않고 바로 삭제
                ws.rows(i).Delete: deleted = deleted + 1
            End If
        End If
    Next i

' --- 4. 매크로 마무리 ---
Cleanup:
    Application.ScreenUpdating = True ' 화면 업데이트 재개
    
    ' 취소 버튼을 누르지 않은 경우에만 완료 메시지 표시
    If Not (choiceMade And userChoice = vbCancel) Then
        MsgBox "빈 행 삭제 작업이 완료되었습니다." & vbCrLf & vbCrLf & _
               "삭제한 행: " & deleted & "개" & IIf(kept > 0, vbCrLf & "유지한 행(다른 열에 데이터): " & kept & "개", "") & vbCrLf & vbCrLf & _
               "되돌리려면 리본 [백업 시트 복원]을 누르세요.", vbInformation
    End If

End Sub


