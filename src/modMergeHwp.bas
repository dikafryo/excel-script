Attribute VB_Name = "modMergeHwp"
Option Explicit

' ===== 사용 설정(필요시 조정) =====
Const SAVE_AS_HWPX As Boolean = False          ' (구형 진입점용) True면 .hwpx로 저장
Const INSERT_PAGE_BETWEEN As Boolean = False   ' 파일 사이 강제 페이지 나누기 삽입 여부(빈 페이지 예방 차원에서 False 권장)
Const KEEP_SECTION As Long = 0                 ' 0=구역 유지 끔(빈페이지 예방), 1=구역 유지 켬(레이아웃 보존용)
' ===================================
Public MH_TestFolder As String   ' 검증용: 지정되면 폴더 대화상자 생략
' [v2.6] MH_Run(mode) : 1 = 모든 글꼴·서식·여백 통합(HWPX, 구역·서식 유지 안 함 → 빈 페이지 없음)
'                       2 = 파일마다 구역을 나눠 글꼴·서식·여백 구분(HWPX, 구역·서식 유지)
'                       0 = 구형 동작(.hwp, 위 상수)

' 메인(구형 진입점): 폴더 선택 → HWP/HWPX 병합 → 시트 기록 → 열기 여부 확인
Public Sub MergeHwpInFolder(control As IRibbonControl)
    MH_Run 0
End Sub

Public Sub MH_Run(ByVal mode As Long)
    Dim folderPath As String, keepSec As Long, keepFmt As Long, asHwpx As Boolean, t0 As Single
    If Len(MH_TestFolder) > 0 Then
        folderPath = MH_TestFolder
    ElseIf Not PickFolder(folderPath) Then
        Exit Sub
    End If
    t0 = Timer
    Select Case mode
        Case 1: keepSec = 0: keepFmt = 0: asHwpx = True
        Case 2: keepSec = 1: keepFmt = 1: asHwpx = True
        Case Else: keepSec = KEEP_SECTION: keepFmt = 1: asHwpx = SAVE_AS_HWPX
    End Select

    Dim hwp As Object
    Set hwp = GetHwpObject()
    If hwp Is Nothing Then
        MsgBox "한글(HWP) 자동화 개체를 만들 수 없습니다." & vbCrLf & _
               "PC에 '한글'이 설치되어 있는지 확인해 주세요.", vbExclamation
        Exit Sub
    End If

    ' 한글 창 가급적 숨김
    On Error Resume Next
    hwp.XHwpWindows.item(0).Visible = False
    On Error GoTo 0

    ' 외부 경로 보안 모듈 등록(삽입 차단 방지)
    On Error Resume Next
    hwp.RegisterModule "FilePathCheckDLL", "FilePathCheckerModule"
    On Error GoTo 0

    ' 새 문서
    hwp.HAction.run "FileNew"

    ' 폴더 내 .hwp / .hwpx 수집 → 이름 오름차순 정렬
    Dim files() As String
    files = ListHwpFiles(folderPath)
    If Not IsArrInitialized(files) Then
        MsgBox "선택한 폴더에 HWP/HWPX 파일이 없습니다.", vbInformation
        Call CloseHwpApp(hwp)
        Exit Sub
    End If
    QuickSortStrings files, LBound(files), UBound(files)

    ' 순차 병합 (페이지 나누기 삽입 안 함 → 빈 페이지 최소화. 구역 유지 모드는 파일마다 새 페이지)
    Dim i As Long, failList As String, failCount As Long
    For i = LBound(files) To UBound(files)
        Application.StatusBar = "한글 병합 중 " & (i - LBound(files) + 1) & "/" & (UBound(files) - LBound(files) + 1) & " : " & FileNameOnly(files(i))
        DoEvents
        If i > LBound(files) And INSERT_PAGE_BETWEEN Then
            On Error Resume Next
            hwp.HAction.run "BreakPage"
            On Error GoTo 0
        End If
        If Not InsertFileToHwp(hwp, files(i), keepSec, keepFmt) Then
            failCount = failCount + 1
            failList = failList & vbCrLf & " - " & FileNameOnly(files(i))
        End If
    Next
    If mode = 1 Then MH_TrimTrailingEmpty hwp
    Application.StatusBar = False
    Dim savePath As String
    If asHwpx Then
        savePath = folderPath & "\병합_" & Format(Now, "yyyymmdd_hhnnss") & ".hwpx"
    Else
        savePath = folderPath & "\병합_" & Format(Now, "yyyymmdd_hhnnss") & ".hwp"
    End If

    ' 저장
    If Not SaveAsHwp(hwp, savePath, asHwpx) Then
        MsgBox "저장에 실패했습니다. 다른 경로나 파일명을 사용해 보세요.", vbExclamation
        Call CloseHwpApp(hwp)
        Exit Sub
    End If
    ' 저장 성공 여부는 실제 파일 존재로 확인
    If dir(savePath) = "" Then
        MsgBox "병합 파일이 만들어지지 않았습니다." & vbCrLf & "폴더 쓰기 권한과 한글 프로그램 상태를 확인하세요.", vbExclamation
        Call CloseHwpApp(hwp)
        Exit Sub
    End If
    ' 요약은 활성 시트를 덮어쓰지 않고 '한글병합' 시트에 기록
    Dim ws As Worksheet
    Set ws = MH_SummarySheet(ActiveWorkbook)
    WriteSummaryToSheet folderPath, savePath, files, ws
    ws.Activate
    CloseHwpApp hwp
    Dim ans As Long, okCount As Long, how As String
    okCount = UBound(files) - LBound(files) + 1 - failCount
    Select Case mode
        Case 1: how = vbCrLf & "방식: 글꼴·서식·여백을 첫 문서 기준으로 통합, 구역·쪽 나눔 없이 이어 붙임(빈 페이지 없음)"
        Case 2: how = vbCrLf & "방식: 파일마다 구역을 나눠 원래 글꼴·서식·여백 유지(각 파일이 새 페이지에서 시작)"
    End Select
    ans = UI_Choose("한글 병합 완료", "파일 " & okCount & "개를 합쳤습니다. (" & Format$(Timer - t0, "0.0") & "초)" & IIf(failCount > 0, vbCrLf & "삽입 실패 " & failCount & "개:" & failList, "") & how & vbCrLf & vbCrLf & _
                    "저장 위치: " & savePath & vbCrLf & "(합친 파일 목록은 '" & ws.name & "' 시트에 있습니다)", Array("병합된 한글 파일 열기", "닫기"))
    If ans = 1 Then OpenByShell savePath
End Sub

' 문서 끝의 빈 문단(빈 페이지 원인)을 정리한다. 실패해도 병합에는 영향 없음
Private Sub MH_TrimTrailingEmpty(ByVal hwp As Object)
    Dim k As Long
    On Error Resume Next
    hwp.HAction.run "MoveDocEnd"
    For k = 1 To 30
        hwp.HAction.run "MoveSelPrevChar"
        If Len(Trim$(CStr(hwp.GetTextFile("TEXT", "saveblock")))) > 0 Then
            hwp.HAction.run "Cancel"
            hwp.HAction.run "MoveDocEnd"
            Exit For
        End If
        hwp.HAction.run "Delete"
    Next k
    On Error GoTo 0
End Sub

' 요약 시트: 있으면 재사용, 없으면 새로 만든다
Private Function MH_SummarySheet(ByVal wb As Workbook) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets("한글병합")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.name = "한글병합"
        ws.Range("A2").value = "폴더"
        ws.Range("A3").value = "병합 파일"
        ws.Range("A2:A3").Font.Bold = True
    End If
    Set MH_SummarySheet = ws
End Function

' 시트 기록: B2=폴더 링크, B3=병합 결과 링크, A4/B4 헤더, A5부터 번호/파일명 링크, 끝에 A1 선택
Sub WriteSummaryToSheet(ByVal folderPath As String, ByVal mergedPath As String, _
                        ByRef files() As String, ByVal ws As Worksheet)
    Dim r As Long, i As Long
    Dim mergedName As String, fname As String
    Dim lastRow As Long
    lastRow = ws.rows.Count

    Application.ScreenUpdating = False

    ' 헤더 부근 병합 해제(A3에 "No." 표기되는 문제 예방)
    On Error Resume Next
    ws.Range("A3:B5").UnMerge
    ws.rows("3:5").hidden = False
    On Error GoTo 0

    ' 기존 하이퍼링크/내용 정리
    On Error Resume Next
    ws.Range("B2:B3").Hyperlinks.Delete
    ws.Range("A5:B" & lastRow).Hyperlinks.Delete
    ws.Range("A5:B" & lastRow).Clear
    On Error GoTo 0

    ' 헤더
    ws.Range("A4").value = "No."
    ws.Range("B4").value = "파일명"
    ws.Range("A4:B4").Font.Bold = True

    ' B2: 폴더 경로 + 링크
    ws.Range("B2").value = folderPath
    ws.Hyperlinks.Add anchor:=ws.Range("B2"), Address:=folderPath, TextToDisplay:=folderPath

    ' B3: 병합된 파일 링크(표시문구 = 파일명)
    mergedName = FileNameOnly(mergedPath) ' 확장자 포함 이름
    ws.Range("B3").value = mergedName
    ws.Hyperlinks.Add anchor:=ws.Range("B3"), Address:=mergedPath, TextToDisplay:=mergedName

    ' 목록: A5부터 번호, B5부터 파일명(확장자 제외) + 파일 링크
    r = 5
    For i = LBound(files) To UBound(files)
        fname = FileNameNoExt(files(i))
        ws.Cells(r, "A").value = r - 4   ' 1부터 증가
        ws.Cells(r, "B").value = fname
        ws.Hyperlinks.Add anchor:=ws.Cells(r, "B"), Address:=files(i), TextToDisplay:=fname
        r = r + 1
    Next i

    ' 보기 좋게 자동 맞춤
    ws.Columns("A:B").AutoFit

    ' 요청: 목록 시트는 A1 선택
    ws.Range("A1").Select

    Application.ScreenUpdating = True
End Sub

' 폴더 선택 대화상자
Function PickFolder(ByRef folderPath As String) As Boolean
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    With fd
        .title = "병합할 한글 파일이 있는 폴더를 선택하세요"
        If .Show <> -1 Then PickFolder = False: Exit Function
        folderPath = .SelectedItems(1)
    End With
    PickFolder = True
End Function

' 한글 OLE 개체 얻기 (버전/환경별 대체 시도 포함)
Function GetHwpObject() As Object
    Dim hwp As Object
    On Error Resume Next
    Set hwp = CreateObject("HwpFrame.HwpObject")
    If hwp Is Nothing Then Set hwp = CreateObject("HWPFrame.HwpObject")
    If hwp Is Nothing Then Set hwp = CreateObject("Hnc.HwpObject")
    On Error GoTo 0
    Set GetHwpObject = hwp
End Function

' .hwp/.hwpx 목록을 배열로 반환
Function ListHwpFiles(ByVal folderPath As String) As String()
    Dim arr() As String
    Dim n As Long: n = 0
    Dim f As String

    ' .hwp
    ' [수정] 이전 실행의 병합 결과(병합_*.hwp)와 임시파일(~$)은 제외 - 병합본이 다시 합쳐지는 것 방지
    f = dir(folderPath & "\*.hwp")
    Do While Len(f) > 0
        If Left$(f, 3) <> "병합_" And Left$(f, 2) <> "~$" And LCase$(Right$(f, 4)) = ".hwp" Then
            n = n + 1
            AddToArray arr, n, folderPath & "\" & f
        End If
        f = dir()
    Loop
    f = dir(folderPath & "\*.hwpx")
    Do While Len(f) > 0
        If Left$(f, 3) <> "병합_" And Left$(f, 2) <> "~$" Then
            n = n + 1
            AddToArray arr, n, folderPath & "\" & f
        End If
        f = dir()
    Loop
    If n = 0 Then Exit Function ' 초기화되지 않은 배열 반환
    ListHwpFiles = arr
End Function

' 동적 배열에 항목 추가
Private Sub AddToArray(ByRef arr() As String, ByVal Index As Long, ByVal value As String)
    If (Not Not arr) = 0 Then
        ReDim arr(1 To 1)
    End If
    If UBound(arr) < Index Then
        ReDim Preserve arr(1 To Index)
    End If
    arr(Index) = value
End Sub

' 문자열 배열 퀵소트 (대소문자 무시, 이름 오름차순)
Sub QuickSortStrings(ByRef a() As String, ByVal first As Long, ByVal last As Long)
    Dim i As Long, j As Long
    Dim pivot As String, temp As String
    i = first: j = last
    pivot = a((first + last) \ 2)

    Do While i <= j
        Do While StrComp(a(i), pivot, vbTextCompare) < 0
            i = i + 1
        Loop
        Do While StrComp(a(j), pivot, vbTextCompare) > 0
            j = j - 1
        Loop
        If i <= j Then
            temp = a(i): a(i) = a(j): a(j) = temp
            i = i + 1: j = j - 1
        End If
    Loop

    If first < j Then QuickSortStrings a, first, j
    If i < last Then QuickSortStrings a, i, last
End Sub

' 파일 삽입(문서 구조 보존 옵션)
Function InsertFileToHwp(ByVal hwp As Object, ByVal filePath As String, ByVal keepSection As Long, Optional ByVal keepFmt As Long = 1) As Boolean
    On Error GoTo EH
    hwp.HAction.GetDefault "InsertFile", hwp.HParameterSet.HInsertFile.HSet
    hwp.HParameterSet.HInsertFile.fileName = filePath

    ' 빈 페이지 예방: 구역 유지 끔 권장
    hwp.HParameterSet.HInsertFile.keepSection = keepSection
    hwp.HParameterSet.HInsertFile.KeepCharshape = keepFmt
    hwp.HParameterSet.HInsertFile.KeepParashape = keepFmt
    hwp.HParameterSet.HInsertFile.KeepStyle = keepFmt

    hwp.HAction.Execute "InsertFile", hwp.HParameterSet.HInsertFile.HSet
    InsertFileToHwp = True
    Exit Function
EH:
    InsertFileToHwp = False
End Function

' 저장(확장자에 맞춰 저장되며, 기본은 .hwp)
Function SaveAsHwp(ByVal hwp As Object, ByVal savePath As String, Optional ByVal asHwpx As Boolean = False) As Boolean
    On Error GoTo EH
    If asHwpx Then
        ' [v2.6] HWPX 저장: SaveAs(경로, "HWPX") 우선, 안 되면 HFileOpenSave.Format
        On Error Resume Next
        If (hwp.SaveAs(savePath, "HWPX", "") <> 0) Then
            If dir(savePath) <> "" Then SaveAsHwp = True: Exit Function
        End If
        Err.Clear
        hwp.HAction.GetDefault "FileSaveAs_S", hwp.HParameterSet.HFileOpenSave.HSet
        hwp.HParameterSet.HFileOpenSave.fileName = savePath
        hwp.HParameterSet.HFileOpenSave.Format = "HWPX"
        hwp.HAction.Execute "FileSaveAs_S", hwp.HParameterSet.HFileOpenSave.HSet
        SaveAsHwp = (dir(savePath) <> "")
        Exit Function
    End If
    hwp.HAction.GetDefault "FileSaveAs_S", hwp.HParameterSet.HFileOpenSave.HSet
    hwp.HParameterSet.HFileOpenSave.fileName = savePath
    hwp.HAction.Execute "FileSaveAs_S", hwp.HParameterSet.HFileOpenSave.HSet
    SaveAsHwp = True
    Exit Function
EH:
    On Error Resume Next
    hwp.HAction.GetDefault "FileSaveAs", hwp.HParameterSet.HFileOpenSave.HSet
    hwp.HParameterSet.HFileOpenSave.fileName = savePath
    hwp.HAction.Execute "FileSaveAs", hwp.HParameterSet.HFileOpenSave.HSet
    SaveAsHwp = True
End Function

' 한글 자동화 종료(창/문서 잔상 방지)
Sub CloseHwpApp(ByRef hwp As Object)
    On Error Resume Next
    hwp.HAction.run "FileClose" ' 열려 있는 문서 닫기
    hwp.HAction.run "FileExit"  ' 한글 종료
    Set hwp = Nothing
    On Error GoTo 0
End Sub

' 기본 프로그램으로 파일 열기(Yes일 때만)
Sub OpenByShell(ByVal targetPath As String)
    Dim cmd As String
    cmd = "cmd /c start """" """ & targetPath & """"
    Shell cmd, vbHide
End Sub

' 유틸: 배열 초기화 여부
Function IsArrInitialized(ByRef arr() As String) As Boolean
    On Error Resume Next
    IsArrInitialized = (Not Not arr) <> 0
    On Error GoTo 0
End Function

' 유틸: 파일명만 (확장자 포함)
Function FileNameOnly(ByVal fullPath As String) As String
    FileNameOnly = Mid(fullPath, InStrRev(fullPath, "\") + 1)
End Function

' 유틸: 파일명 (확장자 제외)
Function FileNameNoExt(ByVal fullPath As String) As String
    Dim nm As String, p As Long
    nm = Mid(fullPath, InStrRev(fullPath, "\") + 1)
    p = InStrRev(nm, ".")
    If p > 0 Then
        FileNameNoExt = Left(nm, p - 1)
    Else
        FileNameNoExt = nm
    End If
End Function

