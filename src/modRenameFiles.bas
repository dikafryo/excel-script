Attribute VB_Name = "modRenameFiles"
Option Explicit

' 전역 변수
Public userSelectionType As Integer
Private Const DUP_COL As Long = 16     ' 중복단어 목록 시작 열(P)

' =================================================================================================
' 1: 폴더/파일 목록 가져오기 (숨김폴더, .tmp 제외 반영)
' =================================================================================================
Sub p001_폴더선택(control As IRibbonControl)
    Dim fso As Object, sourceFolder As Object, subFolder As Object, fileItem As Object
    Dim folderPath As String, sheetName As String
    Dim wb As Workbook, ws As Worksheet
    Dim i As Long
    Dim fileNameWithoutExt As String, fileExt As String

    Set wb = ActiveWorkbook
    If wb Is Nothing Then
        MsgBox "작업을 실행할 엑셀 파일이 없습니다.", vbExclamation
        Exit Sub
    End If

    userSelectionType = 0
    userSelectionType = UI_Choose("작업 유형 선택", "이름을 바꿀 대상을 고르세요.", Array("폴더", "파일", "폴더 및 파일"))

    If userSelectionType = 0 Then
        MsgBox "작업이 취소되었습니다.", vbInformation
        Exit Sub
    End If

    With Application.FileDialog(msoFileDialogFolderPicker)
        .title = "기준 폴더를 선택하세요."
        If .Show <> -1 Then
            MsgBox "폴더를 선택하지 않아 작업을 취소합니다.", vbExclamation
            Exit Sub
        End If
        folderPath = .SelectedItems(1)
    End With

    Application.ScreenUpdating = False

    Select Case userSelectionType
        Case 1: sheetName = "폴더"
        Case 2: sheetName = "파일"
        Case 3: sheetName = "폴더및파일"
    End Select

    Application.DisplayAlerts = False
    On Error Resume Next
    wb.Sheets(sheetName).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    Set ws = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    ws.name = sheetName

    Call CreateMacroButtons(ws)
    Call SetupHeaders(ws)

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set sourceFolder = fso.GetFolder(folderPath)
    ws.Range("A2").value = sourceFolder.name
    ws.Hyperlinks.Add anchor:=ws.Range("B2"), Address:=folderPath, TextToDisplay:=folderPath

    i = 4

    ' 폴더 목록 가져오기
    If userSelectionType = 1 Or userSelectionType = 3 Then
        For Each subFolder In sourceFolder.SubFolders
            ' 숨김 폴더 제외 (vbHidden = 2)
            If (subFolder.Attributes And vbHidden) = 0 Then
                ws.Cells(i, "B").value = "  " & subFolder.name
                ws.Cells(i, "C").value = "폴더"
                i = i + 1
            End If
        Next subFolder
    End If

    ' 파일 목록 가져오기
    If userSelectionType = 2 Or userSelectionType = 3 Then
        For Each fileItem In sourceFolder.files
            ' .tmp 파일 및 엑셀 임시파일(~$) 제외
            If InStr(1, fileItem.name, ".tmp", vbTextCompare) = 0 And Left(fileItem.name, 2) <> "~$" Then
                fileExt = fso.GetExtensionName(fileItem.path)
                fileNameWithoutExt = fso.GetBaseName(fileItem.path)
                ws.Cells(i, "B").value = "  " & fileNameWithoutExt
                ws.Cells(i, "C").value = fileExt
                i = i + 1
            End If
        Next fileItem
    End If

    Call ApplyDataFormatting(ws)
    Application.ScreenUpdating = True
    MsgBox "'" & sheetName & "' 시트에 목록 생성이 완료되었습니다.", vbInformation
End Sub

Public Sub p001_목록생성()
    p001_폴더선택 Nothing
End Sub

' =================================================================================================
' 2: 단어 추출 및 이름 조합
' =================================================================================================
Public Sub p002_이름가져오기()
    Dim wb As Workbook, ws As Worksheet, part As Variant
    Dim lastRow As Long, i As Long, j As Long, lastCol As Long
    Dim originalName As String, nameParts() As String, keyword As String

    Set wb = ActiveWorkbook
    If wb Is Nothing Then Exit Sub
    Set ws = wb.ActiveSheet
    If Not IsValidSheet(ws) Then Exit Sub

    Application.ScreenUpdating = False

    lastRow = ws.Cells(ws.rows.Count, "B").End(xlUp).row
    If lastRow < 4 Then
        Application.ScreenUpdating = True
        MsgBox "목록이 없습니다. 먼저 [1 목록 생성]을 누르세요.", vbExclamation
        Exit Sub
    End If
    lastCol = ws.Cells(2, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < 6 Then lastCol = 10
    If Application.WorksheetFunction.CountA(ws.Range(ws.Cells(3, 6), ws.Cells(3, lastCol))) = 0 Then
        Application.ScreenUpdating = True
        MsgBox "3행 F열부터(단어입력1~) 찾을 단어를 입력한 뒤 [2 단어 추출]을 누르세요." & vbCrLf & _
               "예: 보고서, 2026 → 이름에서 그 단어가 든 부분만 뽑아 D열에 조합합니다.", vbInformation
        Exit Sub
    End If
    For i = 4 To lastRow
        originalName = Trim(ws.Cells(i, "B").value)
        nameParts = Split(ReplaceCharsWithSpace(originalName), " ")

        For j = 6 To lastCol
            keyword = ws.Cells(3, j).value
            If keyword <> "" Then
                ws.Cells(i, j).ClearContents
                For Each part In nameParts
                    If InStr(1, part, keyword, vbTextCompare) > 0 Then
                        ws.Cells(i, j).value = part
                        Exit For
                    End If
                Next part
            End If
        Next j
        Call UpdateCombinedName(ws, i)
    Next i

    For j = 6 To lastCol
        If ws.Cells(3, j).value <> "" Then ws.Cells(3, j).Interior.color = RGB(144, 238, 144)
    Next j

    Call ApplyDataFormatting(ws)
    Application.ScreenUpdating = True
End Sub

' =================================================================================================
' 3: 단어입력 추가
' =================================================================================================
Public Sub p_단어입력추가()
    Dim ws As Worksheet, lastCol As Long

    Set ws = ActiveSheet
    If Not IsValidSheet(ws) Then Exit Sub

    lastCol = ws.Cells(2, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < 5 Then lastCol = 5
    
    lastCol = lastCol + 1
    ws.Cells(2, lastCol).value = "단어입력" & (lastCol - 5)
    ws.Columns(lastCol).NumberFormat = "@"
    ws.Cells(3, lastCol).Interior.ColorIndex = xlNone
    
    With ws.Range(ws.Cells(3, 1), ws.Cells(3, lastCol))
        .Font.Bold = True
        .Interior.color = RGB(173, 216, 230)
        .Borders.LineStyle = xlContinuous
    End With
    ws.Cells(3, lastCol).Interior.ColorIndex = xlNone

    Call ApplyDataFormatting(ws)
End Sub

' =================================================================================================
' 4: 이름 수동 설정
' =================================================================================================
Public Sub p003_이름새로설정()
    Dim wb As Workbook, ws As Worksheet
    Dim lastRow As Long, i As Long, j As Long, lastCol As Long
    Dim keyword As String

    Set wb = ActiveWorkbook
    If wb Is Nothing Then Exit Sub
    Set ws = wb.ActiveSheet
    If Not IsValidSheet(ws) Then Exit Sub

    Application.ScreenUpdating = False

    lastRow = ws.Cells(ws.rows.Count, "B").End(xlUp).row
    If lastRow < 4 Then
        Application.ScreenUpdating = True
        Exit Sub
    End If
    lastCol = ws.Cells(2, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < 6 Then lastCol = 10

    For j = 6 To lastCol
        keyword = ws.Cells(3, j).value
        If keyword <> "" Then
            ws.Cells(3, j).Interior.color = RGB(255, 255, 224)
            If Application.WorksheetFunction.CountA(ws.Range(ws.Cells(4, j), ws.Cells(lastRow, j))) = 0 Then
                ws.Range(ws.Cells(4, j), ws.Cells(lastRow, j)).value = keyword
            End If
        End If
    Next j

    For i = 4 To lastRow
        Call UpdateCombinedName(ws, i)
    Next i

    Call ApplyDataFormatting(ws)
    Application.ScreenUpdating = True
End Sub

' =================================================================================================
' 5: 이름 변경 (수정본 - 핵심 버그 수정)
' =================================================================================================
Public Sub p004_이름변경()
    Dim wb As Workbook, ws As Worksheet, fso As Object
    Dim folderPath As String, oldName As String, newName As String, fileExt As String
    Dim oldPath As String, newPath As String
    Dim lastRow As Long, i As Long, successCount As Long, failCount As Long
    Dim skipCount As Long, finalNewName As String
    Dim invalidChars As String, c As String, k As Integer
    Dim foundFile As String, isFolder As Boolean

    Set wb = ActiveWorkbook
    If wb Is Nothing Then Exit Sub
    Set ws = wb.ActiveSheet
    If Not IsValidSheet(ws) Then Exit Sub
    Set fso = CreateObject("Scripting.FileSystemObject")

    On Error Resume Next
    If ws.Range("B2").Hyperlinks.Count > 0 Then
        folderPath = ws.Range("B2").Hyperlinks(1).Address
    Else
        folderPath = Trim(ws.Range("B2").value)
    End If
    On Error GoTo 0

    If folderPath = "" Or Not fso.FolderExists(folderPath) Then
        MsgBox "기준 폴더 경로가 유효하지 않습니다.", vbCritical
        Exit Sub
    End If

    lastRow = ws.Cells(ws.rows.Count, "B").End(xlUp).row
    If lastRow < 4 Then
        Application.ScreenUpdating = True
        Exit Sub
    End If

    If UI_Choose("이름 변경 확인", lastRow - 3 & "개 항목의 이름을 D열(이름 변경) 값으로 실제 변경합니다. 시작할까요?" & vbCrLf & "(결과는 E열에 표시됩니다)", Array("이름 변경 시작", "취소")) <> 1 Then Exit Sub

    Application.ScreenUpdating = False
    successCount = 0: failCount = 0: skipCount = 0
    invalidChars = "\/:*?""<>|"

    For i = 4 To lastRow
        ws.Cells(i, "E").ClearContents
        ws.Cells(i, "E").Interior.ColorIndex = xlNone

        ' B열 원본 이름 (앞 공백 2칸 제거)
        oldName = Trim(ws.Cells(i, "B").value)
        newName = Trim(ws.Cells(i, "D").value)
        fileExt = Trim(ws.Cells(i, "C").value)
        isFolder = (LCase(fileExt) = "폴더")

        ' [검증 1] 새 이름 비어있음
        If newName = "" Then
            ws.Cells(i, "E").value = "실패: 새 이름 없음"
            ws.Cells(i, "E").Interior.color = RGB(255, 200, 200)
            failCount = failCount + 1
            GoTo NextItem
        End If

        ' [검증 2] 새 이름에 금지문자 포함
        For k = 1 To Len(invalidChars)
            c = Mid(invalidChars, k, 1)
            If InStr(newName, c) > 0 Then
                ws.Cells(i, "E").value = "실패: 금지문자 '" & c & "' 포함"
                ws.Cells(i, "E").Interior.color = RGB(255, 200, 200)
                failCount = failCount + 1
                GoTo NextItem
            End If
        Next k

        ' [핵심 수정] 실제 디스크에서 원본 파일/폴더 찾기
        foundFile = FindActualFileName(fso, folderPath, oldName, fileExt, isFolder)

        If foundFile = "" Then
            ws.Cells(i, "E").value = "실패: 원본 없음 또는 비슷한 이름이 여러 개"
            ws.Cells(i, "E").Interior.color = RGB(255, 200, 200)
            failCount = failCount + 1
            GoTo NextItem
        End If

        ' 확장자 처리
        If Not isFolder And fileExt <> "" Then
            finalNewName = newName & "." & fileExt
        Else
            finalNewName = newName
        End If

        If Not UI_ValidFolderName(finalNewName) Then
            ws.Cells(i, "E").value = "실패: 사용할 수 없는 파일 이름"
            ws.Cells(i, "E").Interior.color = RGB(255, 200, 200)
            failCount = failCount + 1
            GoTo NextItem
        End If
        oldPath = folderPath & "\" & foundFile
        newPath = folderPath & "\" & finalNewName

        ' [검증 3] 동일 이름이면 건너뛰기
        If StrComp(oldPath, newPath, vbTextCompare) = 0 Then
            ws.Cells(i, "E").value = "건너뜀: 동일 이름"
            ws.Cells(i, "E").Interior.color = RGB(255, 255, 200)
            skipCount = skipCount + 1
            GoTo NextItem
        End If

        ' [검증 4] 대상 이름이 이미 존재 (대소문자 변경 제외)
        If isFolder Then
            If fso.FolderExists(newPath) And StrComp(oldPath, newPath, vbTextCompare) <> 0 Then
                ws.Cells(i, "E").value = "실패: 동일 이름 폴더 이미 존재"
                ws.Cells(i, "E").Interior.color = RGB(255, 200, 200)
                failCount = failCount + 1
                GoTo NextItem
            End If
        Else
            If fso.FileExists(newPath) And StrComp(oldPath, newPath, vbTextCompare) <> 0 Then
                ws.Cells(i, "E").value = "실패: 동일 이름 파일 이미 존재"
                ws.Cells(i, "E").Interior.color = RGB(255, 200, 200)
                failCount = failCount + 1
                GoTo NextItem
            End If
        End If

        ' 실제 이름 변경
        On Error Resume Next
        Name oldPath As newPath
        If Err.Number = 0 Then
            ws.Cells(i, "E").value = "성공"
            ws.Cells(i, "E").Interior.color = RGB(200, 255, 200)
            ' B열에 새 이름 반영 (공백 2칸 유지)
            ws.Cells(i, "B").value = "  " & newName
            successCount = successCount + 1
        Else
            ws.Cells(i, "E").value = "실패: " & Err.Description
            ws.Cells(i, "E").Interior.color = RGB(255, 200, 200)
            failCount = failCount + 1
        End If
        Err.Clear
        On Error GoTo 0
NextItem:
    Next i
    Application.ScreenUpdating = True
    MsgBox "성공: " & successCount & "개 / 실패: " & failCount & "개 / 건너뜀: " & skipCount & "개", vbInformation
End Sub

' =================================================================================================
' 5: 중복단어 삭제 (2단계)
'   1회차: B열 이름들을 빈칸·구분문자로 나눠 2개 이상 이름에 공통으로 들어 있는 단어를 P~R열에 정리
'          (P 단어 / Q 포함 이름 수 / R 삭제 여부 'O')  → 남길 단어는 R열의 O를 지운다
'   2회차: R열이 O인 단어를 각 이름에서 제거하고, 한 이름 안에서 반복되는 단어도 1개만 남겨 D열(이름 변경)에 기록
'          → [6 이름 변경]으로 실제 적용
' =================================================================================================
Public Sub p005_중복단어삭제()
    Dim ws As Worksheet, lastRow As Long, i As Long, r As Long
    Dim dict As Object, seen As Object, parts() As String, W As Variant, keys As Variant
    Dim n As Long, selCount As Long, changed As Long, newName As String, origName As String, ext As String
    Set ws = ActiveSheet
    If Not IsValidSheet(ws) Then Exit Sub
    lastRow = ws.Cells(ws.rows.Count, "B").End(xlUp).row
    If lastRow < 4 Then
        MsgBox "목록이 없습니다. 먼저 [1 목록 생성]을 누르세요.", vbExclamation, "중복단어 삭제"
        Exit Sub
    End If

    ' ---- 2회차: 목록이 있고 적용 전이면 삭제 적용 (O 단어 삭제 + 반복 단어 정리) ----
    If ws.Cells(3, DUP_COL).value = "중복 단어" And ws.Cells(4, DUP_COL).value <> "" And ws.Cells(2, DUP_COL).value <> "적용됨" Then
        DupApply ws, lastRow, False
        Exit Sub
    End If

    ' ---- 1회차: 공통 단어 목록 만들기 ----
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare
    For i = 4 To lastRow
        parts = Split(ReplaceCharsWithSpace(Trim$(CStr(ws.Cells(i, "B").value))), " ")
        Set seen = CreateObject("Scripting.Dictionary")
        seen.CompareMode = vbTextCompare
        For Each W In parts
            If Len(W) > 0 And Not seen.Exists(CStr(W)) Then
                seen.Add CStr(W), 1
                If dict.Exists(CStr(W)) Then dict(CStr(W)) = dict(CStr(W)) + 1 Else dict.Add CStr(W), 1
            End If
        Next W
    Next i

    Application.ScreenUpdating = False
    ws.Range(ws.Cells(2, DUP_COL), ws.Cells(ws.rows.Count, DUP_COL + 2)).Clear
    ws.Cells(2, DUP_COL).value = ""
    ws.Cells(3, DUP_COL).value = "중복 단어"
    ws.Cells(3, DUP_COL + 1).value = "포함 이름 수"
    ws.Cells(3, DUP_COL + 2).value = "삭제(O)"
    keys = SortDictKeysByValue(dict)
    r = 4
    For i = LBound(keys) To UBound(keys)
        If dict(keys(i)) >= 2 Then
            ws.Cells(r, DUP_COL).value = keys(i)
            ws.Cells(r, DUP_COL + 1).value = dict(keys(i))
            ws.Cells(r, DUP_COL + 2).value = "O"
            r = r + 1: n = n + 1
        End If
    Next i
    With ws.Range(ws.Cells(3, DUP_COL), ws.Cells(3, DUP_COL + 2))
        .Font.Bold = True
        .Interior.color = RGB(255, 230, 153)
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
    End With
    If n > 0 Then
        With ws.Range(ws.Cells(4, DUP_COL), ws.Cells(3 + n, DUP_COL + 2))
            .Borders.LineStyle = xlContinuous
            .Font.name = "맑은 고딕"
        End With
        ws.Range(ws.Cells(4, DUP_COL + 1), ws.Cells(3 + n, DUP_COL + 2)).HorizontalAlignment = xlCenter
    End If
    ws.Columns(DUP_COL).ColumnWidth = 18
    ws.Columns(DUP_COL + 1).ColumnWidth = 11
    ws.Columns(DUP_COL + 2).ColumnWidth = 8
    If n > 0 Then
        CreateDupApplyButton ws
    Else
        On Error Resume Next
        ws.Shapes("btnDupApply").Delete
        On Error GoTo 0
    End If
    Application.ScreenUpdating = True
    If n = 0 Then
        MsgBox "2개 이상의 이름에 공통으로 들어 있는 단어가 없습니다." & vbCrLf & _
               "(한 이름 안의 반복 단어 정리는 [5 중복단어 삭제]를 한 번 더 누르세요)", vbInformation, "중복단어 삭제"
        ws.Cells(4, DUP_COL).value = "(없음)"
        ws.Cells(4, DUP_COL + 2).value = ""
    Else
        MsgBox "여러 이름에 공통으로 들어 있는 단어 " & n & "개를 P~R열에 정리했습니다 (많이 나온 순)." & vbCrLf & vbCrLf & _
               "· 지울 단어는 R열 'O' 그대로 두고, 남길 단어는 R열의 O를 지우세요." & vbCrLf & _
               "· 그다음 [5 중복단어 삭제]를 다시 누르면 D열(이름 변경)에 정리된 이름이 들어갑니다.", vbInformation, "중복단어 삭제"
    End If
End Sub


' [v2.6] 표 위 버튼 [O 삭제 항목만 삭제] : R열이 'O'인 단어만 이름에서 지운다 (한 이름 안의 반복 단어는 그대로)
'        기존처럼 [5 중복단어 삭제]를 한 번 더 누르면 O 단어 삭제 + 반복 단어 정리를 함께 한다.
Public Sub p005_중복단어적용()
    Dim ws As Worksheet, lastRow As Long
    Set ws = ActiveSheet
    If Not IsValidSheet(ws) Then Exit Sub
    lastRow = ws.Cells(ws.rows.Count, "B").End(xlUp).row
    If ws.Cells(3, DUP_COL).value <> "중복 단어" Or ws.Cells(4, DUP_COL).value = "" Or ws.Cells(4, DUP_COL).value = "(없음)" Then
        MsgBox "중복 단어 목록이 없습니다. 먼저 [5 중복단어 삭제]를 눌러 목록을 만드세요.", vbExclamation, "중복단어 삭제"
        Exit Sub
    End If
    If lastRow < 4 Then Exit Sub
    DupApply ws, lastRow, True
End Sub

' 2회차(적용): R열 'O' 단어를 각 이름에서 제거하고 D열(이름 변경)에 기록. onlyMarked=False 면 한 이름 안의 반복 단어도 1개만 남김
Private Sub DupApply(ByVal ws As Worksheet, ByVal lastRow As Long, ByVal onlyMarked As Boolean)
    Dim dict As Object, seen As Object, parts() As String, W As Variant
    Dim i As Long, r As Long, selCount As Long, changed As Long, newName As String, origName As String
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare
    r = 4
    Do While ws.Cells(r, DUP_COL).value <> ""
        If UCase$(Trim$(CStr(ws.Cells(r, DUP_COL + 2).value))) = "O" Then
            dict(CStr(ws.Cells(r, DUP_COL).value)) = 1
            selCount = selCount + 1
        End If
        r = r + 1
    Loop
    If selCount = 0 Then
        If onlyMarked Then
            MsgBox "R열에 'O'로 표시된 단어가 없습니다. 지울 단어의 R열에 O 를 적은 뒤 다시 누르세요.", vbExclamation, "중복단어 삭제"
            Exit Sub
        End If
        If UI_Choose("중복단어 삭제", "R열에 'O'로 표시된 단어가 없습니다." & vbCrLf & "한 이름 안에서 반복되는 단어만 정리할까요?", Array("반복 단어만 정리", "취소")) <> 1 Then Exit Sub
    End If
    Application.ScreenUpdating = False
    For i = 4 To lastRow
        origName = Trim$(CStr(ws.Cells(i, "B").value))
        parts = Split(ReplaceCharsWithSpace(origName), " ")
        Set seen = CreateObject("Scripting.Dictionary")
        seen.CompareMode = vbTextCompare
        newName = ""
        For Each W In parts
            If Len(W) > 0 Then
                If Not dict.Exists(CStr(W)) And (onlyMarked Or Not seen.Exists(CStr(W))) Then
                    If Not seen.Exists(CStr(W)) Then seen.Add CStr(W), 1
                    newName = newName & IIf(newName = "", "", " ") & W
                End If
            End If
        Next W
        If newName = "" Then newName = origName          ' 전부 지워지면 원래 이름 유지
        ws.Cells(i, "D").value = newName
        If StrComp(newName, ReplaceCharsWithSpace(origName), vbTextCompare) <> 0 Then
            changed = changed + 1
            ws.Cells(i, "E").value = "중복단어 정리 → 이름 변경 대기"
            ws.Cells(i, "E").Interior.color = RGB(255, 242, 204)
        Else
            ws.Cells(i, "E").value = ""
            ws.Cells(i, "E").Interior.ColorIndex = xlNone
        End If
    Next i
    ws.Cells(2, DUP_COL).value = "적용됨"
    ws.Cells(2, DUP_COL).Font.color = RGB(0, 128, 0)
    Call ApplyDataFormatting(ws)
    Application.ScreenUpdating = True
    MsgBox "선택 단어 " & selCount & "개를 지" & IIf(onlyMarked, "워", "우고 반복 단어를 정리해") & " D열(이름 변경)에 넣었습니다." & vbCrLf & _
           "바뀐 이름: " & changed & "개 / 전체 " & (lastRow - 3) & "개" & vbCrLf & vbCrLf & _
           "D열을 확인·수정한 뒤 [6 이름 변경]을 누르면 실제로 바뀝니다." & vbCrLf & _
           "(다시 [5 중복단어 삭제]를 누르면 목록을 새로 만듭니다)", vbInformation, "중복단어 삭제"
End Sub

' 중복 단어 표(P~R) 위에 [O 삭제 항목만 삭제] 버튼 만들기 (1행, 기존 버튼 6개와 겹치지 않게)
Private Sub CreateDupApplyButton(ByVal ws As Worksheet)
    Dim s As Shape, L As Single
    On Error Resume Next
    ws.Shapes("btnDupApply").Delete
    On Error GoTo 0
    L = ws.Columns(DUP_COL).Left
    If L < 5 + 6 * 116 + 10 Then L = 5 + 6 * 116 + 10
    Set s = ws.Shapes.AddShape(msoShapeRoundedRectangle, L, 6, 150, 30)
    s.name = "btnDupApply"
    UI_StyleButton s, "O 삭제 항목만 삭제", "p005_중복단어적용", False, RGB(255, 149, 0), RGB(255, 255, 255)
End Sub

' Dictionary 키를 값 내림차순으로 정렬
Private Function SortDictKeysByValue(ByVal dict As Object) As Variant
    Dim keys As Variant, vals As Variant, i As Long, j As Long, tk As Variant, tv As Variant
    keys = dict.keys: vals = dict.items
    For i = LBound(keys) To UBound(keys) - 1
        For j = i + 1 To UBound(keys)
            If vals(i) < vals(j) Then
                tv = vals(i): vals(i) = vals(j): vals(j) = tv
                tk = keys(i): keys(i) = keys(j): keys(j) = tk
            End If
        Next j
    Next i
    SortDictKeysByValue = keys
End Function

' =================================================================================================
' 보조 프로시저들
' =================================================================================================

' [신규] 실제 디스크에서 원본 파일/폴더 찾기
' B열 이름은 ReplaceCharsWithSpace로 변형될 수 있으므로 실제 파일명과 매칭 필요
Private Function FindActualFileName(fso As Object, folderPath As String, oldName As String, fileExt As String, isFolder As Boolean) As String
    Dim folder As Object, entries As Object, item As Object
    Dim candidate As String, cleanOld As String, matchCount As Long
    Set folder = fso.GetFolder(folderPath)
    If isFolder Then
        Set entries = folder.SubFolders
    Else
        Set entries = folder.files
    End If
    cleanOld = ReplaceCharsWithSpace(oldName)
    For Each item In entries
        If isFolder Then
            candidate = item.name
        Else
            If StrComp(fso.GetExtensionName(item.path), fileExt, vbTextCompare) <> 0 Then GoTo NextExact
            candidate = fso.GetBaseName(item.path)
        End If
        If StrComp(candidate, oldName, vbTextCompare) = 0 Then
            FindActualFileName = item.name
            Exit Function
        End If
NextExact:
    Next item
    For Each item In entries
        If isFolder Then
            candidate = item.name
        Else
            If StrComp(fso.GetExtensionName(item.path), fileExt, vbTextCompare) <> 0 Then GoTo NextCandidate
            candidate = fso.GetBaseName(item.path)
        End If
        If StrComp(ReplaceCharsWithSpace(candidate), cleanOld, vbTextCompare) = 0 Then
            matchCount = matchCount + 1
            FindActualFileName = item.name
        End If
NextCandidate:
    Next item
    If matchCount <> 1 Then FindActualFileName = ""
End Function

Private Sub CreateMacroButtons(ws As Worksheet)
    Dim s As Shape, i As Long, names As Variant, captions As Variant, actions As Variant
    names = Array("btnFolderSelect", "btnApplyWords", "btnAddWord", "btnResetName", "btnDupWords", "btnRename")
    captions = Array("1 목록 생성", "2 단어 추출", "3 입력 추가", "4 이름 조합", "5 중복단어 삭제", "6 이름 변경")
    actions = Array("p001_목록생성", "p002_이름가져오기", "p_단어입력추가", "p003_이름새로설정", "p005_중복단어삭제", "p004_이름변경")
    ws.rows(1).RowHeight = 42
    For i = 0 To 5
        On Error Resume Next
        ws.Shapes(CStr(names(i))).Delete
        On Error GoTo 0
        Set s = ws.Shapes.AddShape(msoShapeRoundedRectangle, 5 + i * 116, 6, 108, 30)
        s.name = CStr(names(i))
        UI_StyleButton s, CStr(captions(i)), CStr(actions(i)), i = 5
    Next i
End Sub

Private Sub SetupHeaders(ws As Worksheet)
    ws.Columns("D:Z").NumberFormat = "@"

    ws.Range("A2").value = "폴더명"
    ws.Range("F2").value = "단어입력1"
    ws.Range("G2").value = "단어입력2"
    ws.Range("H2").value = "단어입력3"
    ws.Range("I2").value = "단어입력4"
    ws.Range("J2").value = "단어입력5"

    ws.Range("A3").value = "No."
    ws.Range("B3").value = "이름"
    ws.Range("C3").value = "정보"
    ws.Range("D3").value = "이름 변경"
    ws.Range("E3").value = "결과"

    With ws.Range("A3:J3")
        .Font.Bold = True
        .Interior.color = RGB(173, 216, 230)
        .Borders.LineStyle = xlContinuous
    End With

    ws.Range("F3:J3").Interior.ColorIndex = xlNone
End Sub

Private Sub ApplyDataFormatting(ws As Worksheet)
    Dim lastRow As Long, lastCol As Long, r As Long
    lastRow = ws.Cells(ws.rows.Count, "B").End(xlUp).row
    lastCol = ws.Cells(2, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < 5 Then lastCol = 5

    If lastRow >= 4 Then
        For r = 4 To lastRow
            If r Mod 2 <> 0 Then
                ws.Range(ws.Cells(r, 1), ws.Cells(r, lastCol)).Interior.color = RGB(245, 248, 252)
            Else
                ws.Range(ws.Cells(r, 1), ws.Cells(r, lastCol)).Interior.ColorIndex = xlNone
            End If
        Next r
        ws.Range("A4").value = 1
        If lastRow > 4 Then ws.Range("A4:A" & lastRow).DataSeries Rowcol:=xlColumns, Type:=xlLinear, Step:=1   ' [수정] 항목 1개면 DataSeries 오류
    End If
    ws.Range(ws.Columns(1), ws.Columns(lastCol)).AutoFit
End Sub

Private Sub UpdateCombinedName(ws As Worksheet, rowNum As Long)
    Dim combinedName As String
    Dim j As Long, lastCol As Long

    lastCol = ws.Cells(2, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < 6 Then Exit Sub

    combinedName = ""
    For j = 6 To lastCol
        If ws.Cells(rowNum, j).value <> "" Then
            combinedName = combinedName & ws.Cells(rowNum, j).value & " "
        End If
    Next j

    ws.Cells(rowNum, "D").value = Trim(combinedName)
End Sub

Private Function ReplaceCharsWithSpace(text As String) As String
    Dim CleanText As String, specialChars As String, i As Long
    specialChars = "_-.,()[]{}!@#$%^&+="
    CleanText = text
    For i = 1 To Len(specialChars)
        CleanText = Replace(CleanText, Mid(specialChars, i, 1), " ")
    Next i
    Do While InStr(CleanText, "  ") > 0
        CleanText = Replace(CleanText, "  ", " ")
    Loop
    ReplaceCharsWithSpace = Trim(CleanText)
End Function

Private Function IsValidSheet(ws As Worksheet) As Boolean
    If ws.name = "폴더" Or ws.name = "파일" Or ws.name = "폴더및파일" Then
        IsValidSheet = True
    Else
        MsgBox "이 매크로는 '폴더', '파일', '폴더및파일' 시트에서 실행해야 합니다.", vbExclamation
        IsValidSheet = False
    End If
End Function

