Attribute VB_Name = "modMergeExcelFiles"
' [수정] 이미 열려 있는 통합문서는 그대로 쓰고 닫지 않는다 (사용자 작업 파일 보호)
Private mxOpenedByUs As Object
Private mxPresetMode As Long   ' [v2.6] 메뉴에서 미리 정한 방식: 0 물어봄 / 1 시트 이름별 / 2 모든 시트 한 시트
Private Function MX_OpenOrGet(ByVal p As String) As Workbook
    Dim W As Workbook
    If mxOpenedByUs Is Nothing Then Set mxOpenedByUs = CreateObject("Scripting.Dictionary")
    For Each W In Workbooks
        If StrComp(W.fullName, p, vbTextCompare) = 0 Then
            Set MX_OpenOrGet = W
            Exit Function
        End If
    Next W
    Set MX_OpenOrGet = Workbooks.Open(p, ReadOnly:=True, UpdateLinks:=0)
    mxOpenedByUs(LCase$(p)) = 1
End Function
Private Sub MX_CloseIfOpened(ByVal W As Workbook)
    If W Is Nothing Then Exit Sub
    If mxOpenedByUs Is Nothing Then Exit Sub
    If mxOpenedByUs.Exists(LCase$(W.fullName)) Then
        mxOpenedByUs.Remove LCase$(W.fullName)
        W.Close False
    End If
End Sub

' [v2.6] OFFICE파일합치기 메뉴 ⑤ 진입점. mode 1 = 시트 이름별로 통합(같은 이름 시트끼리 한 시트), 2 = 모든 시트를 한 시트(Combined)에 통합
Public Sub MX_Run(ByVal mode As Long)
    mxPresetMode = mode
    On Error Resume Next
    시트합치기 Nothing
    mxPresetMode = 0
    On Error GoTo 0
End Sub

' 빌드 검증용(모듈 컴파일 확인)
Public Function MX_SelfTest() As String
    MX_SelfTest = "MX ok " & SanitizeSheetName("a/b*c") & " " & ExtractSchoolWordFromFileName("2026_가상고_현황.xlsx")
End Function

Public Sub 시트합치기(control As IRibbonControl)

    '// ========================= 변수 선언부 =========================
    Dim fso As Object, folder As Object, file As Object
    Dim wb As Workbook, newWb As Workbook
    Dim ws As Worksheet, sumWs As Worksheet, combWs As Worksheet, destWs As Worksheet
    Dim dialog As FileDialog
    Dim folderPath As String, ext As String, fileName As String, filteredName As String, wsName As String
    Dim fileList As Collection
    Dim i As Long, j As Long, k As Long
    Dim pasteRow As Long, lastRow As Long, lastCol As Long
    Dim totalFiles As Long, processedFiles As Long, failedFiles As Long
    Dim MergeMode As Integer
    Dim wsDict As Object
    Dim mergeSheetName As String
    
    '// === 수정된 부분: 암호 보호 파일 감지를 위한 변수 추가 ===
    Dim protectedFiles As Collection
    Dim protectedFilesMsg As String
    Dim isProtected As Boolean
    
    '// ========================= 초기 설정 및 최적화 =========================
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    '// ==================== 사용자 입력 (폴더 및 통합 방식 선택) ====================
    Set dialog = Application.FileDialog(msoFileDialogFolderPicker)
    dialog.title = "합칠 엑셀 파일이 있는 폴더를 선택하세요"
    If dialog.Show = -1 Then
        folderPath = dialog.SelectedItems(1)
    Else
        MsgBox "폴더를 선택하지 않아 작업을 취소합니다.", vbExclamation
        GoTo Cleanup
    End If
    
    Select Case mxPresetMode
        Case 1: MergeMode = vbNo
        Case 2: MergeMode = vbYes
        Case Else
            MergeMode = UI_ChoiceToMsg(UI_Choose("데이터 통합 방식 선택", "폴더 안 엑셀 파일의 시트를 어떻게 합칠까요?", Array("모든 데이터를 한 시트(Combined)에 통합", "시트 이름별로 각각 다른 시트에 통합")))
    End Select

    If MergeMode = vbCancel Then GoTo Cleanup

    '// ==================== 지정 폴더 내 엑셀 파일 목록 수집 ====================
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(folderPath)
    Set fileList = New Collection
    
    For Each file In folder.files
        ext = LCase(fso.GetExtensionName(file.name))
        If (ext = "xlsx" Or ext = "xls" Or ext = "xlsm" Or ext = "xlsb") And Left(file.name, 2) <> "~$" Then
            fileList.Add file
        End If
    Next file
    
    totalFiles = fileList.Count
    If totalFiles = 0 Then
        MsgBox "선택한 폴더에 엑셀 파일이 없습니다!", vbExclamation
        GoTo Cleanup
    End If

    '// =========================================================================
    '// === 수정된 부분: 시트 보호(암호)가 걸린 파일이 있는지 미리 검사하고 알림 ===
    '// =========================================================================
    Set protectedFiles = New Collection
    For Each file In fileList
        On Error Resume Next ' 파일이 열리지 않는 경우 등 에러 대비
        Set wb = MX_OpenOrGet(file.path)
        If Err.Number <> 0 Then ' 파일 열기 실패 시 건너뛰기
            Err.Clear
            GoTo NextFileCheck
        End If
        On Error GoTo 0
        
        isProtected = False
        For Each ws In wb.Worksheets
            If ws.ProtectContents Then
                isProtected = True
                Exit For ' 보호된 시트를 하나라도 찾으면 더 이상 검사할 필요 없음
            End If
        Next ws
        
        If isProtected Then
            protectedFiles.Add file.name
        End If
        
        MX_CloseIfOpened wb
NextFileCheck:
    Next file

    '--- 검사 결과, 보호된 파일이 있으면 메시지 박스로 목록을 보여줌
    If protectedFiles.Count > 0 Then
        protectedFilesMsg = "다음 파일의 일부 시트가 암호로 보호되어 있습니다:" & vbCrLf & vbCrLf
        For i = 1 To protectedFiles.Count
            protectedFilesMsg = protectedFilesMsg & " - " & protectedFiles(i) & vbCrLf
        Next i
        protectedFilesMsg = protectedFilesMsg & vbCrLf & "시트에 암호가 있을경우 기존파일의 서식을 가져오지 않습니다."
        MsgBox protectedFilesMsg, vbInformation, "암호 보호 파일 알림"
    End If
    '// ======================= 사전 검사 및 알림 부분 끝 =======================


    '// ==================== 새 통합문서 및 '작업요약' 시트 준비 ====================
    Set newWb = Workbooks.Add
    
    On Error Resume Next
    Set sumWs = newWb.Sheets("작업요약")
    On Error GoTo 0
    If sumWs Is Nothing Then
        Set sumWs = newWb.Sheets.Add(before:=newWb.Sheets(1))
        sumWs.name = "작업요약"
    End If
    sumWs.Cells.Clear
    
    With sumWs
        .Range("A2").value = "선택 폴더:"
        .Range("B2").value = folderPath
        .Hyperlinks.Add anchor:=.Range("B2"), Address:=folderPath, TextToDisplay:=folderPath
        .Range("B2").Font.Underline = xlUnderlineStyleSingle
        .Range("B2").Font.color = vbBlue
        .Range("A3:C3").value = Array("No.", "파일명", "비고")
        With .Range("A3:C3")
            .Font.Bold = True
            .Interior.color = RGB(184, 204, 228)
            .Borders.color = RGB(0, 0, 0)
        End With
    End With

    '// ==================== 통합 방식에 따른 초기 시트 설정 ====================
    If MergeMode = vbYes Then
        Set combWs = newWb.Sheets.Add(After:=newWb.Sheets(newWb.Sheets.Count))
        combWs.name = "Combined"
        combWs.Cells.Clear
        With combWs.Range("A1:C1")
            .value = Array("No.", "파일명", "데이터")
            .Font.Bold = True
            .Interior.color = RGB(184, 204, 228)
        End With
        pasteRow = 2
    Else
        Set wsDict = CreateObject("Scripting.Dictionary")
    End If

    Application.StatusBar = "0% 준비중..."

    '// ==================== 메인 작업: 파일 반복 및 데이터 취합 ====================
    processedFiles = 0
    For i = 1 To totalFiles
        Set file = fileList(i)
        processedFiles = i
        
        Application.StatusBar = "진행률: " & Format(processedFiles / totalFiles, "0%") & " (" & processedFiles & "/" & totalFiles & ")"
        DoEvents
        
        With sumWs
            .Cells(3 + processedFiles, 1).value = processedFiles
            .Hyperlinks.Add anchor:=.Cells(3 + processedFiles, 2), Address:=file.path, TextToDisplay:=file.name
        End With
        
        filteredName = ExtractSchoolWordFromFileName(file.name)
        
        ' [수정] 열리지 않는 파일(잠김·손상)은 건너뛰고 작업요약에 표시 (전체 작업 중단 방지)
        Set wb = Nothing
        On Error Resume Next
        Set wb = MX_OpenOrGet(file.path)
        On Error GoTo 0
        If wb Is Nothing Then
            failedFiles = failedFiles + 1
            sumWs.Cells(3 + processedFiles, 3).value = "열기 실패(건너뜀)"
            sumWs.Cells(3 + processedFiles, 3).Font.color = vbRed
            GoTo NextFileMain
        End If

        For Each ws In wb.Worksheets
            lastRow = 0
            On Error Resume Next
            lastRow = ws.UsedRange.rows.Count
            On Error GoTo 0
            
            If lastRow > 0 Then
                If MergeMode = vbYes Then
                    '--- [A] "Combined" 시트에 모든 데이터 통합
                    With combWs
                        .Cells(pasteRow, 1).Resize(lastRow, 1).value = processedFiles
                        
                        For j = 0 To lastRow - 1
                            .Hyperlinks.Add anchor:=.Cells(pasteRow + j, 2), Address:=file.path, TextToDisplay:=filteredName
                        Next j
                        
                        '// === 수정된 부분: 시트 보호 여부에 따라 복사 방식 변경 ===
                        If ws.ProtectContents Then
                            '--- 보호된 시트: 서식 없이 값만 복사 ---
                            .Cells(pasteRow, 3).Resize(ws.UsedRange.rows.Count, ws.UsedRange.Columns.Count).value = ws.UsedRange.value
                        Else
                            '--- 일반 시트: 서식 포함하여 복사 ---
                            ws.UsedRange.Copy .Cells(pasteRow, 3)
                        End If
                    End With
                    pasteRow = pasteRow + lastRow
                Else
                    '--- [B] 시트 이름별로 별도 시트에 통합
                    
                    '--- 원본 시트 이름을 가져와 엑셀 규칙에 맞게 정제
                    mergeSheetName = SanitizeSheetName(ws.name)
                    
                    '--- 정제된 이름이 Dictionary에 없으면 새 시트 생성
                    If Not wsDict.Exists(mergeSheetName) Then
                        Set destWs = newWb.Sheets.Add(After:=newWb.Sheets(newWb.Sheets.Count))
                        
                        '--- 시트 이름 지정 (오류 발생 시 중복되지 않게 숫자 추가)
                        On Error Resume Next
                        destWs.name = mergeSheetName
                        If Err.Number <> 0 Then
                            Err.Clear
                            k = 1
                            Do
                                On Error Resume Next
                                Dim tempName As String
                                tempName = Left(mergeSheetName, 31 - (Len(CStr(k)) + 2)) & "_" & k
                                destWs.name = tempName
                                If Err.Number = 0 Then Exit Do
                                k = k + 1
                                On Error GoTo 0 ' 루프 내 에러 핸들링 초기화
                            Loop
                        End If
                        On Error GoTo 0
                        
                        '--- 최종 할당된 시트 이름으로 변수 업데이트
                        mergeSheetName = destWs.name
                        
                        wsDict.Add mergeSheetName, destWs ' Dictionary에 새 시트 등록
                        
                        With destWs.Range("A1:C1")
                           .value = Array("No.", "파일명", "데이터")
                           .Font.Bold = True
                           .Interior.color = RGB(184, 204, 228)
                        End With
                        
                        wsDict.Add mergeSheetName & "_p", 2 ' 붙여넣을 행 번호 등록
                    End If
                    
                    Set destWs = wsDict(mergeSheetName)
                    pasteRow = wsDict(mergeSheetName & "_p")
                    
                    With destWs
                        .Cells(pasteRow, 1).Resize(lastRow, 1).value = processedFiles
                        
                        For j = 0 To lastRow - 1
                            .Hyperlinks.Add anchor:=.Cells(pasteRow + j, 2), Address:=file.path, TextToDisplay:=filteredName
                        Next j
                        
                        '// === 수정된 부분: 시트 보호 여부에 따라 복사 방식 변경 ===
                        If ws.ProtectContents Then
                            '--- 보호된 시트: 서식 없이 값만 복사 ---
                            .Cells(pasteRow, 3).Resize(ws.UsedRange.rows.Count, ws.UsedRange.Columns.Count).value = ws.UsedRange.value
                        Else
                            '--- 일반 시트: 서식 포함하여 복사 ---
                            ws.UsedRange.Copy .Cells(pasteRow, 3)
                        End If
                    End With
                    
                    wsDict(mergeSheetName & "_p") = pasteRow + lastRow
                End If
            End If
        Next ws
        
        MX_CloseIfOpened wb
        Application.CutCopyMode = False '클립보드 정리
NextFileMain:
    Next i

    '// ==================== 마무리 작업: 서식 지정 및 UI 정리 ====================
    With sumWs
        Dim lastSummary As Long
        lastSummary = .Cells(.rows.Count, 1).End(xlUp).row
        If lastSummary > 3 Then
            For j = 4 To lastSummary
                If (j Mod 2 = 0) Then
                    .Range("A" & j & ":B" & j).Interior.color = RGB(220, 220, 220)
                End If
                If .Range("A" & j).value <> "" Then
                    .Range("A" & j & ":B" & j).Borders.LineStyle = xlContinuous
                End If
            Next j
        End If
        .Columns("A:C").AutoFit
    End With
    
    Application.StatusBar = "완료!"
    '// 2.7h: 완료 메시지 전에 화면 갱신을 켠다(꺼진 채 MsgBox 가 뜨면 엑셀 창이 검게 보임)
    Application.ScreenUpdating = True
    DoEvents

    If failedFiles = 0 Then
        MsgBox "작업이 성공적으로 완료되었습니다!" & vbCrLf & "총 " & totalFiles & "개 파일의 데이터를 취합했습니다.", vbInformation
    Else
        MsgBox "작업이 완료되었습니다." & vbCrLf & "취합: " & (totalFiles - failedFiles) & "개 / 열기 실패(건너뜀): " & failedFiles & "개" & vbCrLf & vbCrLf & _
               "실패한 파일은 '작업요약' 시트 비고란을 확인하세요.", vbExclamation
    End If

Cleanup:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
End Sub

'// =================================================================================
'// 함수 이름: SanitizeSheetName
'// 설? ? ? 명: 엑셀 시트 이름으로 사용할 수 없는 문자 제거 및 길이 조절
'// =================================================================================
Function SanitizeSheetName(ByVal sheetName As String) As String
    Dim invalidChars As String
    Dim i As Long
    Dim tempName As String
    
    '--- 엑셀 시트 이름에 사용할 수 없는 문자 목록
    invalidChars = "[]*/\?:" & "'"
    tempName = sheetName
    
    '--- 1. 유효하지 않은 문자를 밑줄(_)로 변경
    For i = 1 To Len(invalidChars)
        tempName = Replace(tempName, Mid(invalidChars, i, 1), "_")
    Next i
    
    '--- 2. 이름 길이가 31자를 넘으면 자르기
    If Len(tempName) > 31 Then
        tempName = Left(tempName, 31)
    End If
    
    '--- 3. 이름 앞/뒤 공백 제거
    tempName = Trim(tempName)
    
    '--- 4. 만약 이름이 비어있다면 기본 이름 부여
    If tempName = "" Then
        tempName = "Sheet1"
    End If
    
    SanitizeSheetName = tempName
End Function


'// =================================================================================
'// 함수 이름: ExtractSchoolWordFromFileName
'// 설? ? ? 명: 파일명에서 '고', '중', '초', '학교' 등의 키워드가 포함된 부분을 추출
'// =================================================================================
Function ExtractSchoolWordFromFileName(str As String) As String
    Dim temp As String, Result As String
    Dim i As Long
    Dim arr As Variant, piece As Variant
    Dim keywords As Variant
    keywords = Array("고", "중", "초", "학교")

    temp = str
    For i = 1 To Len(temp)
        If Not (Mid(temp, i, 1) Like "[가-힣a-zA-Z0-9]") Then
            Mid(temp, i, 1) = "_"
        End If
    Next i
    
    arr = Split(temp, "_")
    
    Result = ""
    For Each piece In arr
        If piece <> "" Then
            For i = LBound(keywords) To UBound(keywords)
                If InStr(piece, keywords(i)) > 0 Then
                    If InStr(Result, piece) = 0 Then
                        If Result = "" Then
                            Result = piece
                        Else
                            Result = Result & "," & piece
                        End If
                    End If
                    Exit For
                End If
            Next i
        End If
    Next piece
    
    If Result = "" Then Result = "-"
    
    ExtractSchoolWordFromFileName = Result
End Function


