Attribute VB_Name = "modFolderClassify"
Option Explicit

' =================================================================================================
'  모듈명: modFileClassifier (폴더/파일 분류 및 이동 자동화)
' =================================================================================================
'  사전 준비: VBA 편집기 > 도구 > 참조 에서 아래 2개 항목을 반드시 체크해주세요.
'  1. Microsoft Scripting Runtime
'  2. Microsoft VBScript Regular Expressions 5.5
' =================================================================================================
'  기능 요약:
'  1. 메뉴: 기준폴더 설정 / 폴더 / 파일 / 폴더및파일 (기준폴더는 현재 문서의 '기준폴더' 시트로 관리, 파일 저장으로 공유)
'     분석 키워드 15개(I~K열)·기준 키워드 15개(M~O열): 빈칸·구분문자로 나눈 단어 빈도, 연도·날짜·숫자 제외
'  2. 키워드 입력 없음 → 폴더/파일명 단어 빈도 분석으로 항목별 분류 폴더 자동 추천
'  3. 파일 내부 텍스트 검색 (Excel, Word, 텍스트파일)
'  4. 미분류 항목: 빈도 기반 자동 추천 키워드 생성
'  5. 분류_결과 시트에 표 작성 + 1행 이동 버튼
'  6. 이동 버튼 → E열 키워드명 폴더 자동 생성 → 파일/폴더 이동
'  7. 동일 파일명 충돌 시 _1, _2 등 자동 번호 부여
'  8. 대용량(1000+) 처리 최적화 + StatusBar 진행률 표시
'  9. 이동 전 미리보기 팝업 + 최종 승인 후 실행
'  10. 이동 완료 후 성공/실패 리포트 업데이트 (열: A NO·B 폴더명·C 파일명·D 확장자·E 이동 폴더·F 이름 변경(안)·G 이동결과·H 분석 추천·I 추천 근거)
' =================================================================================================

Private fso_FC As Object  ' FileSystemObject (모듈 전역)
Private Const KW_MAX As Long = 15                        ' 키워드 표 개수
Private Const BASE_SHEET As String = "기준폴더"
Private Const BASE_KW_PER_FOLDER As Long = 8
Private Const FC_GREEN As Long = 5482804      ' RGB(52,168,83) 정리 실행 버튼
Private Const FC_ORANGE As Long = 39167       ' RGB(255,153,0) 날짜·이름변경 버튼

' =================================================================================================
'  [메인] 분류 분석 실행
' =================================================================================================
Public Sub 폴더분석정리(control As IRibbonControl)
    ' --- 1. 메뉴: 기준폴더 설정 / 폴더 / 파일 / 폴더및파일 ---
    Dim analysisMode As Long, analysisType As String, basePath As String, baseInfo As String
    If fso_FC Is Nothing Then Set fso_FC = CreateObject("Scripting.FileSystemObject")
    Do
        basePath = FC_BasePath()
        If basePath = "" Then
            baseInfo = "기준폴더: (미설정)  ← 1번으로 설정하면 '기준폴더' 시트가 만들어지며, 그 파일을 저장해 관리합니다."
        Else
            baseInfo = "기준폴더(현재 문서의 '기준폴더' 시트): " & basePath
        End If
        analysisMode = UI_Choose("폴더/파일 정리", baseInfo, _
                       Array("기준폴더 설정  (정리해 넣을 기준폴더 지정 → '기준폴더' 시트 생성)", _
                             "선택폴더의 폴더정리", "선택폴더의 파일정리", "선택폴더의 폴더 및 파일정리"))
        If analysisMode = 0 Then Exit Sub
        If analysisMode = 1 Then
            FC_SetBasePathDialog          ' 기준폴더 시트가 만들어지며, 분석은 그 시트의 1행 버튼으로 실행
            Exit Sub
        Else
            Exit Do
        End If
    Loop
    Select Case analysisMode
        Case 2: analysisType = "폴더"
        Case 3: analysisType = "파일"
        Case 4: analysisType = "폴더및파일"
    End Select
    FC_RunAnalysis analysisType
End Sub

' 기준폴더 시트 1행 버튼용 진입점 (메뉴 없이 바로 분석)
Public Sub FC_분석_폴더()
    FC_RunAnalysis "폴더"
End Sub
Public Sub FC_분석_파일()
    FC_RunAnalysis "파일"
End Sub
Public Sub FC_분석_폴더및파일()
    FC_RunAnalysis "폴더및파일"
End Sub

' 분석 본체: 분석 폴더 선택 → 분류_결과 시트 작성 → 키워드 표 → 기준폴더 추천
Public Sub FC_RunAnalysis(ByVal analysisType As String, Optional ByVal presetPath As String = "")
    Dim basePath As String
    If fso_FC Is Nothing Then Set fso_FC = CreateObject("Scripting.FileSystemObject")
    ' --- 2. 분류 키워드: 입력받지 않는다. 폴더/파일 이름의 단어 빈도를 분석해 항목별로 자동 추천한다 ---
    Dim kwArr() As String
    ReDim kwArr(0 To 0)
    kwArr(0) = ""

    ' --- 3. 폴더 선택 (presetPath 가 있으면 대화상자 생략: 자동 검증용) ---
    Dim folderPath As String
    folderPath = presetPath
    If folderPath = "" Then
        With Application.FileDialog(msoFileDialogFolderPicker)
            .title = "분석할 폴더를 선택하세요"
            If .Show = -1 Then
                folderPath = .SelectedItems(1)
            Else
                MsgBox "폴더를 선택하지 않아 작업을 취소합니다.", vbExclamation
                Exit Sub
            End If
        End With
    End If
    
    ' --- 4. 초기화 ---
    If fso_FC Is Nothing Then Set fso_FC = CreateObject("Scripting.FileSystemObject")
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "분석 준비 중..."
    
    ' --- 5. 시트 생성/초기화 ---
    Dim ws As Worksheet
    Set ws = FC_GetOrCreateSheet("분류_결과")
    ws.Cells.Clear
    ws.Cells.Hyperlinks.Delete
    
    ' 1행: 이동 실행 버튼 생성
    FC_CreateMoveButton ws
    
    ' 2행: 경로 표시 + 하이퍼링크
    ws.Range("A2").value = ""
    ws.Range("B2").value = folderPath
    ws.Hyperlinks.Add anchor:=ws.Range("B2"), Address:=folderPath, TextToDisplay:=folderPath
    ws.Range("B2").Font.color = vbBlue
    ws.Range("B2").Font.Underline = xlUnderlineStyleSingle
    ws.Range("E2").value = "분석유형: " & analysisType
    ws.Range("F2").value = "분류기준: 이름 단어 빈도 분석(자동 추천)"
    
    ' 3행: 헤더
    ws.Range("A3").value = "NO"
    ws.Range("B3").value = "폴더명"
    ws.Range("C3").value = "파일명"
    ws.Range("D3").value = "확장자"
    ws.Range("E3").value = "이동 폴더 (수정 가능)"
    ws.Range("F3").value = "이름 변경(안)"
    ws.Range("G3").value = "이동결과"
    ws.Range("H3").value = "분석 추천(참고)"
    ws.Range("I3").value = "추천 근거"
    FC_StyleHeader ws.Range("A3:I3")
    
    ' --- 6. 분석 실행 ---
    Dim mainFolder As Object
    Set mainFolder = fso_FC.GetFolder(folderPath)
    
    Dim rowNum As Long: rowNum = 3  ' 데이터 시작행 (4행부터)
    Dim itemNo As Long: itemNo = 0
    
    ' 전체 단어 빈도 수집용 Dictionary
    Dim globalWordFreq As Object
    Set globalWordFreq = CreateObject("Scripting.Dictionary")
    globalWordFreq.CompareMode = vbTextCompare
    
    Select Case analysisType
        Case "폴더"
            FC_AnalyzeFolders ws, mainFolder, kwArr, rowNum, itemNo, globalWordFreq
        Case "파일"
            FC_AnalyzeFiles ws, mainFolder, kwArr, rowNum, itemNo, globalWordFreq
        Case "폴더및파일"
            FC_AnalyzeFoldersAndFiles ws, mainFolder, kwArr, rowNum, itemNo, globalWordFreq
    End Select
    
    ' --- 7. 미분류 항목에 추천 키워드 생성 (빈도 기반) ---
    FC_FillRecommendations ws, rowNum, globalWordFreq, kwArr
    
    ' --- 7-1. 분석 키워드 표 (I~K열, 빈칸·구분문자 기준 단어 빈도 15개, 연도·날짜·숫자 제외) ---
    FC_WriteKeywordTable ws, globalWordFreq, "K", "분석 키워드", "분석폴더:", folderPath

    ' --- 7-2. 기준폴더가 설정되어 있으면 기준 키워드(M~O열) 추천 + E열 재분류 ---
    basePath = FC_BasePath()
    If basePath <> "" Then
        If fso_FC.FolderExists(basePath) Then FC_ApplyBaseKeywords ws, basePath, False
    End If

    ' --- 8. 서식 적용 ---
    FC_ApplyFormatting ws, rowNum
    
    ' --- 9. 정리 ---
    FC_LayoutResult ws, analysisType
    
    ws.Activate
    ws.Range("A4").Select
    ActiveWindow.FreezePanes = True
    ws.Range("A1").Select
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
    
    If Not Application.Visible Then Exit Sub
    MsgBox "분류 분석이 완료되었습니다." & vbCrLf & vbCrLf & _
           "총 " & itemNo & "개 항목 분석, 키워드 15개(I~K열) 정리" & vbCrLf & _
           IIf(basePath <> "", "기준폴더 키워드(M~O열) 기준으로 E열 추천 완료" & vbCrLf, "") & vbCrLf & _
           "E열 분류(폴더명)를 확인/수정한 뒤" & vbCrLf & _
           "  · [분류 확인 후 이동] : 분석 폴더 안에 E열 이름 폴더를 만들어 이동" & vbCrLf & _
           "  · [기준폴더 이동] : 기준폴더 안 E열 이름 폴더로 이동" & vbCrLf & _
           "(E열 '미분류' 항목은 이동하지 않습니다)", _
           vbInformation, "분석 완료"
End Sub



' 2.7k: [폴더 다시 선택] ? 결과 시트 1행 버튼. 같은 분석 유형으로 다른 폴더를 골라 다시 분석한다
Public Sub FC_폴더다시선택()
    Dim ws As Worksheet, T As String
    Set ws = FC_ResultSheet()
    T = "폴더 및 파일"
    If Not ws Is Nothing Then
        T = Trim$(Replace(CStr(ws.Range("E2").value), "분석유형:", ""))
        If T = "" Then T = "폴더 및 파일"
    End If
    FC_RunAnalysis T
End Sub

' =================================================================================================
'  [메인] 파일 이동 실행 (1행 버튼에 할당)
' =================================================================================================
Public Sub FC_파일이동실행()
    Dim ws As Worksheet, basePath As String
    Set ws = FC_ResultSheet()
    If ws Is Nothing Then Exit Sub
    basePath = FC_SourcePath(ws)
    If basePath = "" Then Exit Sub
    ' 분석 폴더 안에 E열 이름의 폴더를 만들어 이동
    FC_MoveAll ws, basePath, basePath, "분류 확인 후 이동", False
End Sub


' =================================================================================================
'  [기준폴더] 설정: 레지스트리에 저장하지 않고 현재 통합 문서의 '기준폴더' 시트(B2)로만 관리한다.
'   → 시트가 든 파일을 저장해 두면 그 파일로 관리·공유(다른 사용자도 동일하게 사용). 매크로 실행마다 문서 기준으로 새로 읽음
' =================================================================================================
Public Function FC_BasePath() As String
    Dim ws As Worksheet, p As String
    If fso_FC Is Nothing Then Set fso_FC = CreateObject("Scripting.FileSystemObject")
    If ActiveWorkbook Is Nothing Then Exit Function
    On Error Resume Next
    Set ws = ActiveWorkbook.Sheets(BASE_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    p = Trim$(CStr(ws.Range("B2").value))
    On Error Resume Next
    If ws.Range("B2").Hyperlinks.Count > 0 Then p = ws.Range("B2").Hyperlinks(1).Address
    On Error GoTo 0
    If p <> "" Then
        If Not fso_FC.FolderExists(p) Then p = ""
    End If
    FC_BasePath = p
End Function

' 기준폴더 경로 직접 설정(대화상자 없음): 현재 문서에 '기준폴더' 시트를 만든다 - 자동화·검증용
Public Sub FC_SetBasePath(ByVal basePath As String)
    FC_BuildBaseSheet basePath
End Sub

' 기준폴더 선택 대화상자 → 저장. 반환: 선택한 경로("" = 취소)
Private Function FC_SetBasePathDialog() As String
    Dim basePath As String, freq As Object, keys As Variant, i As Long, n As Long, msg As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .title = "기준폴더를 선택하세요 (정리해 넣을 폴더 - 안의 폴더·파일 이름에서 분류 키워드를 뽑습니다)"
        If FC_BasePath() <> "" Then .InitialFileName = FC_BasePath() & "\"
        If .Show <> -1 Then Exit Function
        basePath = .SelectedItems(1)
    End With
    FC_BuildBaseSheet basePath
    Set freq = FC_BaseFolderFreq(basePath)
    keys = FC_TopKeywords(freq, KW_MAX)
    For i = LBound(keys) To UBound(keys)
        If Len(keys(i)) > 0 Then n = n + 1: msg = msg & IIf(msg = "", "", ", ") & keys(i)
    Next i
    MsgBox "기준폴더를 설정했습니다." & vbCrLf & basePath & vbCrLf & vbCrLf & _
           "※ 기준폴더는 이 통합 문서의 '기준폴더' 시트로 관리됩니다. 파일로 저장해 두면 다음에도(다른 사용자도) 그대로 사용할 수 있습니다." & vbCrLf & _
           "'기준폴더' 시트에 폴더별 키워드를 정리했습니다 (C열 키워드는 자유롭게 수정 가능)." & vbCrLf & _
           "전체 키워드 " & n & "개(빈도순): " & IIf(msg = "", "(없음)", msg) & vbCrLf & vbCrLf & _
           "1행 [폴더 선택]·[파일 선택]·[폴더 및 파일 선택]으로 바로 분석하면 이 표로 E열 이동 폴더를 추천합니다.", vbInformation, "기준폴더 설정"
    FC_SetBasePathDialog = basePath
End Function

' =================================================================================================
'  [기준폴더 시트] 기준폴더의 위치와 폴더별 키워드를 시트로 정리 (사용자가 키워드를 수정 가능)
'   1행: [폴더 선택] [파일 선택] [폴더 및 파일 선택] 버튼   2행: 기준폴더 경로(클릭 → 탐색기)
'   3행: 표머리   4행~: 폴더명 | 키워드(쉼표 구분) | 파일 수
' =================================================================================================
Public Sub FC_BuildBaseSheet(ByVal basePath As String)
    Dim ws As Worksheet, fo As Object, sf As Object, f As Object, r As Long, freq As Object, keys As Variant
    Dim i As Long, kwList As String, n As Long, fileCnt As Long, s As Shape, nm As Variant
    If fso_FC Is Nothing Then Set fso_FC = CreateObject("Scripting.FileSystemObject")
    Set ws = FC_GetOrCreateSheet(BASE_SHEET)
    Application.ScreenUpdating = False
    ws.Cells.Clear
    ws.Cells.Hyperlinks.Delete
    On Error Resume Next
    For Each nm In Array("btnBaseRunFolder", "btnBaseRunFile", "btnBaseRunBoth", "btnBaseRefresh", "noteBase")
        ws.Shapes(CStr(nm)).Delete
    Next nm
    On Error GoTo 0

    ' 1행 버튼
    ws.rows(1).RowHeight = 134
    Set s = ws.Shapes.AddShape(msoShapeRoundedRectangle, 6, 6, 150, 30)
    s.name = "btnBaseRefresh": UI_StyleButton s, "기준폴더 다시 읽기", "FC_기준폴더설정", False
    Set s = ws.Shapes.AddShape(msoShapeRoundedRectangle, 6, 44, 170, 30)
    s.name = "btnBaseRunFolder": UI_StyleButton s, "선택폴더의 폴더정리", "FC_분석_폴더", True, FC_GREEN, RGB(255, 255, 255)
    Set s = ws.Shapes.AddShape(msoShapeRoundedRectangle, 184, 44, 170, 30)
    s.name = "btnBaseRunFile": UI_StyleButton s, "선택폴더의 파일정리", "FC_분석_파일", True, FC_GREEN, RGB(255, 255, 255)
    Set s = ws.Shapes.AddShape(msoShapeRoundedRectangle, 362, 44, 210, 30)
    s.name = "btnBaseRunBoth": UI_StyleButton s, "선택폴더의 폴더 및 파일정리", "FC_분석_폴더및파일", True, FC_GREEN, RGB(255, 255, 255)

    ' 2행 경로 (클릭 → 탐색기)
    ws.Range("A2").value = "기준폴더:"
    ws.Range("A2").Font.Bold = True
    ws.Range("B2").value = basePath
    ws.Hyperlinks.Add anchor:=ws.Range("B2"), Address:=basePath, TextToDisplay:=basePath, ScreenTip:="클릭하면 탐색기에서 엽니다"
    FC_AddNote ws, "noteBase", "① 초록 버튼으로 정리할 폴더를 고르면(폴더만 / 파일만 / 폴더 및 파일) '분류_결과' 시트가 만들어집니다." & vbCrLf & _
                              "② 아래 표의 키워드1~15(셀마다 1개)는 자유롭게 고쳐 쓰세요. 이름에 맞는 키워드가 가장 많은 폴더를 E열 이동 폴더로 추천합니다.", 6, 80, 900, 48
    ws.Range("A2:D2").Font.name = "맑은 고딕"
    ws.Range("A2:D2").Font.Size = 10

    ' 3행 표머리
    ws.Range("A3:B3").value = Array("No", "폴더명")
    For i = 1 To KW_MAX
        ws.Cells(3, 2 + i).value = "키워드" & i
    Next i
    ws.Cells(3, 3 + KW_MAX).value = "파일 수"
    FC_StyleHeader ws.Range(ws.Cells(3, 1), ws.Cells(3, 3 + KW_MAX))

    ' 4행~ 폴더별 키워드
    r = 3
    Set fo = fso_FC.GetFolder(basePath)
    For Each sf In fo.SubFolders
        If Not FC_SkipItem(sf, True) Then
            r = r + 1
            Set freq = CreateObject("Scripting.Dictionary")
            freq.CompareMode = vbTextCompare
            FC_CollectWordFreq sf.name, freq
            FC_CollectWordFreq sf.name, freq            ' 폴더 이름 단어 가중치 2
            fileCnt = 0
            On Error Resume Next
            For Each f In sf.files
                If Not FC_SkipItem(f, False) Then
                    fileCnt = fileCnt + 1
                    FC_CollectWordFreq fso_FC.GetBaseName(f.name), freq
                End If
            Next f
            On Error GoTo 0
            keys = FC_TopKeywords(freq, KW_MAX)
            ws.Cells(r, "A").value = r - 3
            ws.Cells(r, "B").value = sf.name
            ws.Hyperlinks.Add anchor:=ws.Cells(r, "B"), Address:=sf.path, TextToDisplay:=sf.name
            kwList = ""
            For i = LBound(keys) To UBound(keys)
                If Len(keys(i)) > 0 Then ws.Cells(r, 2 + i).value = keys(i): kwList = "x"
            Next i
            If kwList = "" Then ws.Cells(r, "C").value = sf.name        ' 키워드가 없으면 폴더명 자체
            ws.Cells(r, 3 + KW_MAX).value = fileCnt
            n = n + 1
        End If
    Next sf
    ws.Range(ws.Columns(3), ws.Columns(2 + KW_MAX)).NumberFormat = "@"
    If n > 0 Then
        With ws.Range(ws.Cells(4, 1), ws.Cells(r, 3 + KW_MAX))
            .Font.name = "맑은 고딕"
            .Font.Size = 11
            .Font.color = RGB(28, 28, 30)
            .Borders.LineStyle = xlContinuous
            .Borders.color = RGB(220, 220, 224)
            .VerticalAlignment = xlCenter
            .RowHeight = 22
        End With
        ws.Range("A4:A" & r).HorizontalAlignment = xlCenter
        ws.Range("A4:A" & r).Font.color = RGB(142, 142, 147)
        ws.Range(ws.Cells(4, 3 + KW_MAX), ws.Cells(r, 3 + KW_MAX)).HorizontalAlignment = xlCenter
        ws.Range("B4:B" & r).Font.Bold = True
        ws.Range(ws.Cells(4, 3), ws.Cells(r, 2 + KW_MAX)).Interior.color = RGB(255, 251, 235)   ' 수정 가능 칸 표시
    Else
        ws.Range("B4").value = "(하위 폴더 없음 - 기준폴더 안에 분류 폴더를 만든 뒤 [기준폴더 다시 읽기])"
    End If
    ws.Columns("A").ColumnWidth = 5
    ws.Columns("B").ColumnWidth = 30
    ws.Range(ws.Columns(3), ws.Columns(2 + KW_MAX)).ColumnWidth = 12
    ws.Columns(3 + KW_MAX).ColumnWidth = 8
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Range("A4").Select
    ActiveWindow.FreezePanes = True
    Application.ScreenUpdating = True
End Sub

' 기준폴더 시트의 표 읽기 → 폴더명 배열·키워드 배열(각 폴더의 키워드는 "," 로 나눈 목록). 없으면 False
Private Function FC_ReadBaseMapping(ByRef folders() As String, ByRef kws() As String, ByRef sheetPath As String) As Boolean
    Dim ws As Worksheet, lastRow As Long, r As Long, n As Long, c As Long
    On Error Resume Next
    Set ws = ActiveWorkbook.Sheets(BASE_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    sheetPath = Trim$(CStr(ws.Range("B2").value))
    On Error Resume Next
    If ws.Range("B2").Hyperlinks.Count > 0 Then sheetPath = ws.Range("B2").Hyperlinks(1).Address
    On Error GoTo 0
    lastRow = ws.Cells(ws.rows.Count, "B").End(xlUp).row
    If lastRow < 4 Then Exit Function
    ReDim folders(1 To lastRow - 3): ReDim kws(1 To lastRow - 3)
    For r = 4 To lastRow
        If Trim$(CStr(ws.Cells(r, "B").value)) <> "" And Left$(Trim$(CStr(ws.Cells(r, "B").value)), 1) <> "(" Then
            n = n + 1
            folders(n) = Trim$(CStr(ws.Cells(r, "B").value))
            kws(n) = ""
            For c = 3 To 2 + KW_MAX
                If Trim$(CStr(ws.Cells(r, c).value)) <> "" Then kws(n) = kws(n) & IIf(kws(n) = "", "", ",") & Trim$(CStr(ws.Cells(r, c).value))
            Next c
            If kws(n) = "" Then kws(n) = folders(n)
        End If
    Next r
    If n = 0 Then Exit Function
    ReDim Preserve folders(1 To n): ReDim Preserve kws(1 To n)
    FC_ReadBaseMapping = True
End Function

' 이름과 기준폴더 표를 대조: 키워드가 가장 많이 맞는 폴더(같으면 위쪽 행 우선). 반환: 폴더명("" = 없음), hitKw 에 맞은 키워드
Private Function FC_MatchBaseFolder(ByVal nm As String, ByRef folders() As String, ByRef kws() As String, ByRef hitKw As String) As String
    Dim i As Long, parts() As String, k As Long, cnt As Long, best As Long, bestKw As String, kw As String
    For i = LBound(folders) To UBound(folders)
        parts = Split(kws(i), ",")
        cnt = 0: bestKw = ""
        For k = LBound(parts) To UBound(parts)
            kw = Trim$(parts(k))
            If Len(kw) > 0 Then
                If InStr(1, nm, kw, vbTextCompare) > 0 Then
                    cnt = cnt + 1
                    If bestKw = "" Then bestKw = kw
                End If
            End If
        Next k
        If cnt > best Then
            best = cnt: FC_MatchBaseFolder = folders(i): hitKw = bestKw
        End If
    Next i
End Function

' 기준폴더의 하위 폴더 이름 + 하위 폴더 안 파일 + 바로 아래 파일 이름의 단어 빈도
Private Function FC_BaseFolderFreq(ByVal basePath As String) As Object
    Dim freq As Object, fo As Object, sf As Object, f As Object
    Set freq = CreateObject("Scripting.Dictionary")
    freq.CompareMode = vbTextCompare
    If fso_FC Is Nothing Then Set fso_FC = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next
    Set fo = fso_FC.GetFolder(basePath)
    For Each sf In fo.SubFolders
        If Not FC_SkipItem(sf, True) Then
            FC_CollectWordFreq sf.name, freq
            FC_CollectWordFreq sf.name, freq            ' 폴더 이름은 가중치 2 (분류 폴더일 가능성이 높음)
            For Each f In sf.files
                If Not FC_SkipItem(f, False) Then FC_CollectWordFreq fso_FC.GetBaseName(f.name), freq
            Next f
        End If
    Next sf
    For Each f In fo.files
        If Not FC_SkipItem(f, False) Then FC_CollectWordFreq fso_FC.GetBaseName(f.name), freq
    Next f
    On Error GoTo 0
    Set FC_BaseFolderFreq = freq
End Function

' 빈도 상위 n개 키워드 (숫자·연도·날짜 제외, 폴더명으로 쓸 수 있는 것만)
Private Function FC_TopKeywords(ByVal freq As Object, ByVal maxN As Long) As Variant
    Dim keys As Variant, out() As String, i As Long, n As Long
    ReDim out(1 To maxN)
    If freq.Count > 0 Then
        keys = FC_SortDictByValue(freq)
        For i = LBound(keys) To UBound(keys)
            If n >= maxN Then Exit For
            If Not FC_IsNumberLike(CStr(keys(i))) And UI_ValidFolderName(CStr(keys(i))) Then
                n = n + 1
                out(n) = CStr(keys(i))
            End If
        Next i
    End If
    FC_TopKeywords = out
End Function

' 키워드 표 쓰기: colStart 열부터 3열(순위·키워드·빈도), 2행에 경로
Private Sub FC_WriteKeywordTable(ByVal ws As Worksheet, ByVal freq As Object, ByVal colStart As String, _
                                 ByVal title As String, ByVal pathLabel As String, ByVal pathValue As String)
    Dim c1 As Long, keys As Variant, i As Long, n As Long
    c1 = ws.Columns(colStart).Column
    ws.Range(ws.Cells(2, c1), ws.Cells(60, c1 + 2)).Hyperlinks.Delete
    ws.Range(ws.Cells(2, c1), ws.Cells(60, c1 + 2)).Clear
    ws.Cells(2, c1).value = pathLabel
    ws.Cells(2, c1 + 1).value = pathValue
    If pathValue <> "" Then ws.Hyperlinks.Add anchor:=ws.Cells(2, c1 + 1), Address:=pathValue, TextToDisplay:=pathValue
    ws.Cells(3, c1).value = "순위": ws.Cells(3, c1 + 1).value = title: ws.Cells(3, c1 + 2).value = "빈도"
    keys = FC_TopKeywords(freq, KW_MAX)
    For i = LBound(keys) To UBound(keys)
        If Len(keys(i)) > 0 Then
            n = n + 1
            ws.Cells(3 + n, c1).value = n
            ws.Cells(3 + n, c1 + 1).value = keys(i)
            ws.Cells(3 + n, c1 + 2).value = freq(keys(i))
        End If
    Next i
    With ws.Range(ws.Cells(3, c1), ws.Cells(3, c1 + 2))
        .Font.Bold = True
        .Interior.color = RGB(242, 242, 247)
        .HorizontalAlignment = xlCenter
    End With
    If n > 0 Then
        With ws.Range(ws.Cells(3, c1), ws.Cells(3 + n, c1 + 2))
            .Font.name = "맑은 고딕"
            .Font.Size = 11
            .Borders.LineStyle = xlContinuous
        End With
        ws.Range(ws.Cells(4, c1), ws.Cells(3 + n, c1)).HorizontalAlignment = xlCenter
    End If
    ws.Cells(2, c1).Font.Bold = True
    ws.Columns(c1).ColumnWidth = 6
    ws.Columns(c1 + 1).ColumnWidth = 22
    ws.Columns(c1 + 2).ColumnWidth = 7
End Sub

' =================================================================================================
'  [메인] 기준폴더 설정 (1행 버튼): 기준폴더 선택·저장 후 현재 결과 시트의 E열을 기준 키워드로 재추천
' =================================================================================================
Public Sub FC_기준폴더설정()
    Dim ws As Worksheet, basePath As String
    If fso_FC Is Nothing Then Set fso_FC = CreateObject("Scripting.FileSystemObject")
    basePath = FC_SetBasePathDialog()
    If basePath = "" Then Exit Sub
    On Error Resume Next
    Set ws = ActiveWorkbook.Sheets("분류_결과")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    If ws.Cells(ws.rows.Count, "A").End(xlUp).row < 4 Then Exit Sub
    Application.ScreenUpdating = False
    FC_ApplyBaseKeywords ws, basePath, True
    Application.ScreenUpdating = True
End Sub

' 이전 이름 호환
Public Sub FC_기준폴더만들기()
    FC_기준폴더설정
End Sub

' 기준 키워드(M~O열) 작성 + E열 재분류: 이름에 든 기준 키워드 중 빈도 최고 → E열, 없으면 '미분류'(H열에 분석 추천 보관)
Public Sub FC_ApplyBaseKeywords(ByVal ws As Worksheet, ByVal basePath As String, ByVal showMsg As Boolean)
    Dim freq As Object, keys As Variant, kw() As String, n As Long, i As Long, r As Long, k As Long, lastRow As Long
    Dim nm As String, hit As String, dest As String, hitCount As Long, noneCount As Long, prev As String
    Dim folders() As String, kws() As String, sheetPath As String, useMap As Boolean, Target As String
    Set freq = FC_BaseFolderFreq(basePath)
    FC_WriteKeywordTable ws, freq, "O", "기준 키워드", "기준폴더:", basePath
    keys = FC_TopKeywords(freq, KW_MAX)
    ReDim kw(1 To KW_MAX)
    For i = LBound(keys) To UBound(keys)
        If Len(keys(i)) > 0 Then n = n + 1: kw(n) = keys(i)
    Next i
    ' 기준폴더 시트(폴더별 키워드 표)가 있으면 그 표를 우선 사용. 없거나 다른 기준폴더면 새로 만든다
    useMap = FC_ReadBaseMapping(folders, kws, sheetPath)
    If useMap Then useMap = (StrComp(sheetPath, basePath, vbTextCompare) = 0)
    If Not useMap Then
        FC_BuildBaseSheet basePath
        ws.Activate
        useMap = FC_ReadBaseMapping(folders, kws, sheetPath)
    End If

    ws.Range("H3").value = "분석 추천(참고)"
    FC_StyleHeader ws.Range("H3")
    lastRow = ws.Cells(ws.rows.Count, "A").End(xlUp).row
    For r = 4 To lastRow
        If Left$(CStr(ws.Cells(r, "G").value), 2) = "성공" Then GoTo NextRow   ' 이미 이동된 항목
        nm = Trim$(CStr(ws.Cells(r, "C").value))
        If nm <> "" Then nm = fso_FC.GetBaseName(nm) Else nm = Trim$(CStr(ws.Cells(r, "B").value))
        prev = Trim$(CStr(ws.Cells(r, "E").value))
        If prev <> "" And prev <> "미분류" And ws.Cells(r, "H").value = "" Then ws.Cells(r, "H").value = prev
        hit = "": Target = ""
        If useMap Then
            Target = FC_MatchBaseFolder(nm, folders, kws, hit)
            If Target <> "" And StrComp(Target, Trim$(CStr(ws.Cells(r, "B").value)), vbTextCompare) = 0 And ws.Cells(r, "D").value = "폴더" Then Target = ""  ' 자기 자신
        End If
        If Target = "" Then
            For k = 1 To n
                If InStr(1, nm, kw(k), vbTextCompare) > 0 Then hit = kw(k): Exit For
            Next k
            If hit <> "" Then Target = fso_FC.GetFileName(FC_ResolveDest(basePath, hit, True))
        End If
        If Target <> "" Then
            ws.Cells(r, "E").value = Target
            ws.Cells(r, "E").Interior.color = RGB(223, 245, 228)
            If fso_FC.FolderExists(fso_FC.BuildPath(basePath, Target)) Then
                ws.Cells(r, "I").value = "키워드 '" & hit & "' → 기준폴더의 '" & Target & "' 폴더로"
            Else
                ws.Cells(r, "I").value = "키워드 '" & hit & "' → 기준폴더에 '" & Target & "' 폴더 새로 만듦"
            End If
            hitCount = hitCount + 1
        Else
            ws.Cells(r, "E").value = "미분류"
            ws.Cells(r, "E").Interior.color = RGB(255, 244, 214)
            ws.Cells(r, "I").value = "기준폴더 키워드 일치 없음 (H열 분석 추천 참고)"
            noneCount = noneCount + 1
        End If
        ws.Cells(r, "G").value = ""
        ws.Cells(r, "G").Interior.ColorIndex = xlNone
NextRow:
    Next r
    ws.Range("F2").value = "분류기준: '기준폴더' 시트의 폴더별 키워드 표 (보조: O~Q열 기준 키워드 / K~M열 분석 키워드)"
    If ws.Columns("I").ColumnWidth > 60 Then ws.Columns("I").ColumnWidth = 60
    If showMsg Then
        MsgBox "'기준폴더' 시트의 폴더별 키워드 표로 E열을 추천했습니다." & vbCrLf & vbCrLf & _
               "E열 분류: 키워드 일치 " & hitCount & "개 / 미분류 " & noneCount & "개" & vbCrLf & vbCrLf & _
               "E열을 확인·수정한 뒤 1행 [기준폴더 이동]을 누르세요.", vbInformation, "기준폴더 설정"
    End If
End Sub


' =================================================================================================
'  [메인] 기준폴더 이동 (1행 버튼): 저장된 기준폴더 안 E열 키워드 이름의 폴더로 이동, 미분류는 이동없음
' =================================================================================================
Public Sub FC_기준폴더이동()
    Dim ws As Worksheet, srcPath As String, basePath As String
    Set ws = FC_ResultSheet()
    If ws Is Nothing Then Exit Sub
    srcPath = FC_SourcePath(ws)
    If srcPath = "" Then Exit Sub
    basePath = FC_BasePath()
    If basePath = "" Then
        MsgBox "이 통합 문서에 '기준폴더' 시트(기준폴더 경로)가 없습니다." & vbCrLf & _
               "먼저 1행 [기준폴더 설정]으로 기준폴더를 선택하세요. (저장해 둔 기준폴더 파일이 있으면 그 파일을 연 뒤 실행)", vbExclamation, "기준폴더 이동"
        Exit Sub
    End If
    If Trim$(CStr(ws.Range("P2").value)) = "" Or StrComp(Trim$(CStr(ws.Range("P2").value)), basePath, vbTextCompare) <> 0 Then
        FC_ApplyBaseKeywords ws, basePath, False      ' 시트에 기준 키워드가 없거나 다른 기준폴더면 재추천
    End If
    FC_MoveAll ws, srcPath, basePath, "기준폴더 이동", True
End Sub


' 분류_결과 시트 (없으면 안내 후 Nothing)
Private Function FC_ResultSheet() As Worksheet
    Dim ws As Worksheet
    If fso_FC Is Nothing Then Set fso_FC = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next
    Set ws = ActiveWorkbook.Sheets("분류_결과")
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "'분류_결과' 시트를 찾을 수 없습니다." & vbCrLf & _
               "먼저 리본 [폴더/파일 정리]를 실행해 주세요.", vbExclamation
    End If
    Set FC_ResultSheet = ws
End Function

' 분석 폴더 경로(B2, 하이퍼링크 우선). 없거나 잘못되면 안내 후 ""
Private Function FC_SourcePath(ByVal ws As Worksheet) As String
    Dim p As String
    p = Trim$(CStr(ws.Range("B2").value))
    On Error Resume Next
    If ws.Range("B2").Hyperlinks.Count > 0 Then p = ws.Range("B2").Hyperlinks(1).Address
    On Error GoTo 0
    If p = "" Or Not fso_FC.FolderExists(p) Then
        MsgBox "분석 경로가 유효하지 않습니다. (B2셀 확인)", vbCritical
        p = ""
    End If
    FC_SourcePath = p
End Function

' 공통 이동: srcBase(분석 폴더)의 항목을 destBase\E열이름 폴더로 이동. 미분류는 '이동없음'
Private Sub FC_MoveAll(ByVal ws As Worksheet, ByVal srcBase As String, ByVal destBase As String, ByVal title As String, ByVal useExisting As Boolean)
    Dim lastRow As Long, r As Long, moveCount As Long, noneCount As Long
    Dim folderSet As Object, movedFolders As Object, fKey As Variant, previewMsg As String
    Dim eVal As String, keyword As String, destFolder As String, destPath As String, sourcePath As String
    Dim folderName As String, fileName As String, extName As String, isFolder As Boolean
    Dim successCount As Long, failCount As Long, skipCount As Long, processed As Long

    lastRow = ws.Cells(ws.rows.Count, "A").End(xlUp).row
    If lastRow < 4 Then
        MsgBox "이동할 데이터가 없습니다.", vbInformation, title
        Exit Sub
    End If

    ' --- 미리보기 ---
    Set folderSet = CreateObject("Scripting.Dictionary")
    folderSet.CompareMode = vbTextCompare
    For r = 4 To lastRow
        If Left$(CStr(ws.Cells(r, "G").value), 2) = "성공" Then GoTo NextPreview
        eVal = Trim$(CStr(ws.Cells(r, "E").value))
        If eVal = "" Or Left$(eVal, 3) = "미분류" Then
            noneCount = noneCount + 1
        Else
            If Not UI_ValidFolderName(eVal) Then
                MsgBox r & "행 E열 분류명을 확인하세요: " & eVal, vbExclamation, title
                Exit Sub
            End If
            moveCount = moveCount + 1
            eVal = fso_FC.GetFileName(FC_ResolveDest(destBase, eVal, useExisting))
            If Not folderSet.Exists(eVal) Then folderSet.Add eVal, 0
            folderSet(eVal) = folderSet(eVal) + 1
        End If
NextPreview:
    Next r

    ws.Cells(3, "G").value = "이동결과"
    ws.Cells(3, "G").Font.Bold = True
    ws.Cells(3, "G").Interior.color = RGB(242, 242, 247)

    If moveCount = 0 Then
        FC_MarkNoMove ws, lastRow
        MsgBox "이동할 항목이 없습니다." & vbCrLf & _
               "E열이 '미분류'인 항목 " & noneCount & "개는 이동하지 않습니다(G열 '이동없음').", vbInformation, title
        Exit Sub
    End If

    previewMsg = "대상 위치: " & destBase & vbCrLf & vbCrLf & "폴더별 이동 항목 수:" & vbCrLf
    For Each fKey In folderSet.keys
        previewMsg = previewMsg & "  ▶ " & CStr(fKey) & " → " & folderSet(fKey) & "개" & vbCrLf
    Next fKey
    previewMsg = previewMsg & vbCrLf & "총 " & moveCount & "개 이동, 미분류 " & noneCount & "개는 이동없음" & vbCrLf & _
                 "※ 파일 이동은 자동 백업되지 않습니다."
    If UI_Choose(title & " 최종 확인", previewMsg, Array("이동 실행", "취소")) <> 1 Then Exit Sub

    ' --- 이동 ---
    Application.ScreenUpdating = False
    Set movedFolders = CreateObject("Scripting.Dictionary")
    movedFolders.CompareMode = vbTextCompare
    For r = 4 To lastRow
        If Left$(CStr(ws.Cells(r, "G").value), 2) = "성공" Then GoTo NextMove
        keyword = Trim$(CStr(ws.Cells(r, "E").value))
        folderName = Trim$(CStr(ws.Cells(r, "B").value))
        fileName = Trim$(CStr(ws.Cells(r, "C").value))
        extName = Trim$(CStr(ws.Cells(r, "D").value))

        If keyword = "" Or Left$(keyword, 3) = "미분류" Then
            FC_SetResult ws, r, "이동없음 (미분류)", RGB(230, 230, 230)
            skipCount = skipCount + 1
            GoTo NextMove
        End If
        If extName <> "폴더" And folderName <> "" Then
            If movedFolders.Exists(folderName) Then
                FC_SetResult ws, r, "폴더와 함께 이동됨", RGB(230, 230, 230)
                skipCount = skipCount + 1
                GoTo NextMove
            End If
        End If

        processed = processed + 1
        Application.StatusBar = title & " 진행률: " & Format(processed / moveCount, "0%") & " (" & processed & "/" & moveCount & ")"
        DoEvents

        destFolder = FC_ResolveDest(destBase, keyword, useExisting)
        If extName = "폴더" Then
            isFolder = True
            sourcePath = fso_FC.BuildPath(srcBase, folderName)
            destPath = fso_FC.BuildPath(destFolder, folderName)
        Else
            isFolder = False
            If folderName <> "" Then
                sourcePath = fso_FC.BuildPath(fso_FC.BuildPath(srcBase, folderName), fileName)
            Else
                sourcePath = fso_FC.BuildPath(srcBase, fileName)
            End If
            destPath = fso_FC.BuildPath(destFolder, fileName)
        End If

        ' 자기 자신·하위 폴더로 이동하거나 이미 그 폴더에 있는 경우
        If isFolder Then
            If StrComp(sourcePath, destFolder, vbTextCompare) = 0 Or _
               StrComp(Left$(destFolder, Len(sourcePath) + 1), sourcePath & Application.PathSeparator, vbTextCompare) = 0 Then
                FC_SetResult ws, r, "건너뜀: 자기 폴더(또는 하위)로 이동 불가", RGB(255, 235, 156)
                skipCount = skipCount + 1
                GoTo NextMove
            End If
        ElseIf StrComp(fso_FC.GetParentFolderName(sourcePath), destFolder, vbTextCompare) = 0 Then
            FC_SetResult ws, r, "이동없음 (이미 해당 폴더)", RGB(230, 230, 230)
            skipCount = skipCount + 1
            GoTo NextMove
        End If

        On Error Resume Next
        If Not fso_FC.FolderExists(destFolder) Then fso_FC.CreateFolder destFolder
        If Err.Number <> 0 Then
            FC_SetResult ws, r, "실패: 폴더생성오류 (" & Err.Description & ")", RGB(255, 200, 200)
            failCount = failCount + 1
            Err.Clear
            On Error GoTo 0
            GoTo NextMove
        End If
        destPath = FC_GetUniquePath(destPath, isFolder)
        If isFolder Then
            If fso_FC.FolderExists(sourcePath) Then fso_FC.MoveFolder sourcePath, destPath Else Err.Raise 53, , "원본 폴더를 찾을 수 없습니다."
        Else
            If fso_FC.FileExists(sourcePath) Then fso_FC.MoveFile sourcePath, destPath Else Err.Raise 53, , "원본 파일을 찾을 수 없습니다."
        End If
        If Err.Number = 0 Then
            FC_SetResult ws, r, "성공 → " & fso_FC.GetFileName(destFolder) & "\" & fso_FC.GetFileName(destPath), RGB(200, 255, 200)
            successCount = successCount + 1
            If isFolder Then movedFolders(folderName) = True
        Else
            FC_SetResult ws, r, "실패: " & Err.Description, RGB(255, 200, 200)
            failCount = failCount + 1
            Err.Clear
        End If
        On Error GoTo 0
NextMove:
    Next r

    ws.Columns("G").AutoFit
    If ws.Columns("I").ColumnWidth > 60 Then ws.Columns("I").ColumnWidth = 60
    If ws.Columns("G").ColumnWidth > 45 Then ws.Columns("G").ColumnWidth = 45
    Application.ScreenUpdating = True
    Application.StatusBar = False

    MsgBox title & " 완료!" & vbCrLf & vbCrLf & _
           " - 성공: " & successCount & "개" & vbCrLf & _
           " - 이동없음(미분류 등): " & skipCount & "개" & vbCrLf & _
           " - 실패: " & failCount & "개" & vbCrLf & vbCrLf & _
           "G열에서 상세 결과를 확인하세요.", vbInformation, title
End Sub

' 이동 대상 폴더: destBase\키워드. useExisting 이면 같은 이름이 없을 때 키워드를 포함한 기존 하위 폴더(이름이 가장 짧은 것)를 쓴다
Private Function FC_ResolveDest(ByVal destBase As String, ByVal keyword As String, ByVal useExisting As Boolean) As String
    Dim p As String, sf As Object, best As String
    p = fso_FC.BuildPath(destBase, keyword)
    FC_ResolveDest = p
    If Not useExisting Then Exit Function
    If fso_FC.FolderExists(p) Then Exit Function
    On Error Resume Next
    For Each sf In fso_FC.GetFolder(destBase).SubFolders
        If InStr(1, sf.name, keyword, vbTextCompare) > 0 Then
            If best = "" Or Len(sf.name) < Len(best) Then best = sf.name
        End If
    Next sf
    On Error GoTo 0
    If best <> "" Then FC_ResolveDest = fso_FC.BuildPath(destBase, best)
End Function

' 분석 제외 항목: 임시(~$)·desktop.ini·Thumbs.db, 숨김(파일·폴더), 시스템 파일
' (폴더는 사용자 지정 아이콘만 있어도 시스템 속성이 붙으므로 숨김·$ 이름만 제외)
Private Function FC_SkipItem(ByVal o As Object, ByVal isFolder As Boolean) As Boolean
    Dim nm As String, att As Long
    On Error Resume Next
    nm = LCase$(o.name)
    att = o.Attributes
    On Error GoTo 0
    If Left$(nm, 2) = "~$" Or Left$(nm, 1) = "$" Or nm = "desktop.ini" Or nm = "thumbs.db" Then FC_SkipItem = True: Exit Function
    If isFolder Then
        FC_SkipItem = ((att And 2) <> 0)
    Else
        FC_SkipItem = ((att And 6) <> 0)
    End If
End Function

' 미분류 행에 '이동없음' 표시 (이동할 항목이 하나도 없을 때)
Private Sub FC_MarkNoMove(ByVal ws As Worksheet, ByVal lastRow As Long)
    Dim r As Long, eVal As String
    For r = 4 To lastRow
        If Left$(CStr(ws.Cells(r, "G").value), 2) <> "성공" Then
            eVal = Trim$(CStr(ws.Cells(r, "E").value))
            If eVal = "" Or Left$(eVal, 3) = "미분류" Then FC_SetResult ws, r, "이동없음 (미분류)", RGB(230, 230, 230)
        End If
    Next r
End Sub

Private Sub FC_SetResult(ByVal ws As Worksheet, ByVal r As Long, ByVal msg As String, ByVal clr As Long)
    ws.Cells(r, "G").value = msg
    ws.Cells(r, "G").Interior.color = clr
    ws.Cells(r, "G").Font.name = "맑은 고딕"
End Sub


' =================================================================================================
'  분석 모드별 서브루틴
' =================================================================================================

' 모드1: 폴더만 분석
Private Sub FC_AnalyzeFolders(ws As Worksheet, mainFolder As Object, kwArr() As String, _
                               ByRef rowNum As Long, ByRef itemNo As Long, globalFreq As Object)
    Dim subFolder As Object
    Dim totalCount As Long: totalCount = mainFolder.SubFolders.Count
    Dim processed As Long: processed = 0
    
    For Each subFolder In mainFolder.SubFolders
        If FC_SkipItem(subFolder, True) Then GoTo NextFolderA
        processed = processed + 1
        If totalCount > 0 Then
            Application.StatusBar = "폴더 분석 중... " & Format(processed / totalCount, "0%") & _
                                    " (" & processed & "/" & totalCount & ")"
        End If
        DoEvents
        
        rowNum = rowNum + 1
        itemNo = itemNo + 1
        
        ws.Cells(rowNum, "A").value = itemNo
        ws.Cells(rowNum, "B").value = subFolder.name
        ws.Cells(rowNum, "C").value = ""  ' 폴더일 경우 공란
        ws.Cells(rowNum, "D").value = "폴더"
        
        ' 키워드 매칭
        Dim matched As String
        matched = FC_MatchKeyword(subFolder.name, kwArr)
        ws.Cells(rowNum, "E").value = matched
        
        ' 빈도 수집
        FC_CollectWordFreq subFolder.name, globalFreq
NextFolderA:
    Next subFolder
End Sub

' 모드2: 파일만 분석 (최상위 폴더의 파일)
Private Sub FC_AnalyzeFiles(ws As Worksheet, mainFolder As Object, kwArr() As String, _
                             ByRef rowNum As Long, ByRef itemNo As Long, globalFreq As Object)
    Dim file As Object
    Dim totalCount As Long: totalCount = mainFolder.files.Count
    Dim processed As Long: processed = 0
    
    For Each file In mainFolder.files
        If FC_SkipItem(file, False) Then GoTo NextFile1
        
        processed = processed + 1
        If totalCount > 0 Then
            Application.StatusBar = "파일 분석 중... " & Format(processed / totalCount, "0%") & _
                                    " (" & processed & "/" & totalCount & ")"
        End If
        DoEvents
        
        rowNum = rowNum + 1
        itemNo = itemNo + 1
        
        ws.Cells(rowNum, "A").value = itemNo
        ws.Cells(rowNum, "B").value = ""  ' 최상위이므로 폴더명 없음
        ws.Cells(rowNum, "C").value = file.name
        ws.Cells(rowNum, "D").value = fso_FC.GetExtensionName(file.name)
        
        ' 키워드 매칭 (파일명 + 내부 텍스트)
        Dim matched As String
        matched = FC_MatchKeyword(file.name, kwArr)
        
        ' 파일명에서 매칭 안 되면 내부 텍스트 검색
        If matched = "" And FC_HasKeywords(kwArr) Then
            Dim innerText As String
            innerText = FC_ExtractFileText(file.path)
            If innerText <> "" Then
                matched = FC_MatchKeyword(innerText, kwArr)
                If matched <> "" Then matched = matched & " (내부검색)"
            End If
        End If
        
        ws.Cells(rowNum, "E").value = matched
        
        ' 빈도 수집
        FC_CollectWordFreq fso_FC.GetBaseName(file.name), globalFreq
        
NextFile1:
    Next file
End Sub

' 모드3: 폴더 및 파일 모두 분석
Private Sub FC_AnalyzeFoldersAndFiles(ws As Worksheet, mainFolder As Object, kwArr() As String, _
                                      ByRef rowNum As Long, ByRef itemNo As Long, globalFreq As Object)
    Dim subFolder As Object, file As Object
    Dim totalFolders As Long: totalFolders = mainFolder.SubFolders.Count
    Dim totalFiles As Long: totalFiles = mainFolder.files.Count
    Dim totalCount As Long: totalCount = totalFolders + totalFiles
    Dim processed As Long: processed = 0
    
    ' 최상위 파일 먼저
    For Each file In mainFolder.files
        If FC_SkipItem(file, False) Then GoTo NextFile2
        
        processed = processed + 1
        If totalCount > 0 Then
            Application.StatusBar = "분석 중... " & Format(processed / totalCount, "0%") & _
                                    " (" & processed & "/" & totalCount & ")"
        End If
        DoEvents
        
        rowNum = rowNum + 1
        itemNo = itemNo + 1
        
        ws.Cells(rowNum, "A").value = itemNo
        ws.Cells(rowNum, "B").value = ""
        ws.Cells(rowNum, "C").value = file.name
        ws.Cells(rowNum, "D").value = fso_FC.GetExtensionName(file.name)
        
        Dim matchedF As String
        matchedF = FC_MatchKeyword(file.name, kwArr)
        If matchedF = "" And FC_HasKeywords(kwArr) Then
            Dim innerTextF As String
            innerTextF = FC_ExtractFileText(file.path)
            If innerTextF <> "" Then
                matchedF = FC_MatchKeyword(innerTextF, kwArr)
                If matchedF <> "" Then matchedF = matchedF & " (내부검색)"
            End If
        End If
        ws.Cells(rowNum, "E").value = matchedF
        FC_CollectWordFreq fso_FC.GetBaseName(file.name), globalFreq
        
NextFile2:
    Next file
    
    ' 하위 폴더 및 그 안의 파일
    For Each subFolder In mainFolder.SubFolders
        If FC_SkipItem(subFolder, True) Then GoTo NextFolderB
        processed = processed + 1
        If totalCount > 0 Then
            Application.StatusBar = "분석 중... " & Format(processed / totalCount, "0%") & _
                                    " (" & processed & "/" & totalCount & ")"
        End If
        DoEvents
        
        ' 폴더 행
        rowNum = rowNum + 1
        itemNo = itemNo + 1
        
        ws.Cells(rowNum, "A").value = itemNo
        ws.Cells(rowNum, "B").value = subFolder.name
        ws.Cells(rowNum, "C").value = ""
        ws.Cells(rowNum, "D").value = "폴더"
        
        Dim matchedD As String
        matchedD = FC_MatchKeyword(subFolder.name, kwArr)
        ws.Cells(rowNum, "E").value = matchedD
        FC_CollectWordFreq subFolder.name, globalFreq
        
        ' 하위 파일들
        On Error Resume Next
        For Each file In subFolder.files
            If FC_SkipItem(file, False) Then GoTo NextFile3
            
            rowNum = rowNum + 1
            itemNo = itemNo + 1
            
            ws.Cells(rowNum, "A").value = itemNo
            ws.Cells(rowNum, "B").value = subFolder.name
            ws.Cells(rowNum, "C").value = file.name
                ws.Cells(rowNum, "D").value = fso_FC.GetExtensionName(file.name)
            
            Dim matchedSF As String
            matchedSF = FC_MatchKeyword(file.name, kwArr)
            If matchedSF = "" And FC_HasKeywords(kwArr) Then
                Dim innerTextSF As String
                innerTextSF = FC_ExtractFileText(file.path)
                If innerTextSF <> "" Then
                    matchedSF = FC_MatchKeyword(innerTextSF, kwArr)
                    If matchedSF <> "" Then matchedSF = matchedSF & " (내부검색)"
                End If
            End If
            ws.Cells(rowNum, "E").value = matchedSF
            FC_CollectWordFreq fso_FC.GetBaseName(file.name), globalFreq
            
NextFile3:
        Next file
        On Error GoTo 0
NextFolderB:
    Next subFolder
End Sub


' =================================================================================================
'  키워드 매칭 함수
' =================================================================================================
Private Function FC_HasKeywords(kwArr() As String) As Boolean
    Dim i As Long
    On Error Resume Next
    For i = LBound(kwArr) To UBound(kwArr)
        If Len(Trim$(kwArr(i))) > 0 Then FC_HasKeywords = True: Exit Function
    Next i
End Function

Private Function FC_MatchKeyword(ByVal text As String, kwArr() As String) As String
    Dim i As Long
    Dim matchedList As String: matchedList = ""
    
    For i = LBound(kwArr) To UBound(kwArr)
        If Len(kwArr(i)) > 0 Then
            If InStr(1, text, kwArr(i), vbTextCompare) > 0 Then
                If matchedList = "" Then
                    matchedList = kwArr(i)
                Else
                    matchedList = matchedList & ", " & kwArr(i)
                End If
            End If
        End If
    Next i
    
    FC_MatchKeyword = matchedList
End Function


' =================================================================================================
'  파일 내부 텍스트 추출
' =================================================================================================
Private Function FC_ExtractFileText(ByVal filePath As String) As String
    On Error GoTo ErrHandler
    
    Dim ext As String
    ext = LCase(fso_FC.GetExtensionName(filePath))
    
    Select Case ext
        Case "xlsx", "xls", "xlsm", "xlsb"
            FC_ExtractFileText = FC_ExtractExcelText(filePath)
        Case "docx", "doc"
            FC_ExtractFileText = FC_ExtractWordText(filePath)
        Case "txt", "csv", "log"
            FC_ExtractFileText = FC_ReadTextFile(filePath)
        Case Else
            FC_ExtractFileText = ""
    End Select
    Exit Function
    
ErrHandler:
    FC_ExtractFileText = ""
End Function

' Excel 파일 내부 텍스트 추출
Private Function FC_ExtractExcelText(ByVal filePath As String) As String
    On Error GoTo ErrHandler
    
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim Result As String: Result = ""
    Dim maxChars As Long: maxChars = 5000  ' 성능 고려 최대 문자수
    
    Set wb = Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=0)
    
    For Each ws In wb.Worksheets
        If Len(Result) > maxChars Then Exit For
        
        Dim usedRng As Range
        Set usedRng = ws.UsedRange
        
        If Not usedRng Is Nothing Then
            Dim cell As Range
            For Each cell In usedRng
                If Len(Result) > maxChars Then Exit For
                If Not IsEmpty(cell.value) Then
                    Result = Result & " " & CStr(cell.value)
                End If
            Next cell
        End If
    Next ws
    
    wb.Close False
    FC_ExtractExcelText = Result
    Exit Function
    
ErrHandler:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    On Error GoTo 0
    FC_ExtractExcelText = ""
End Function

' Word 파일 내부 텍스트 추출
Private Function FC_ExtractWordText(ByVal filePath As String) As String
    On Error GoTo ErrHandler
    
    Dim wdApp As Object
    Dim wdDoc As Object
    Dim Result As String
    
    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = False
    wdApp.DisplayAlerts = 0  ' wdAlertsNone
    
    Set wdDoc = wdApp.Documents.Open(filePath, ReadOnly:=True, AddToRecentFiles:=False)
    
    ' 최대 5000자까지만 추출 (성능 고려)
    If Len(wdDoc.content.text) > 5000 Then
        Result = Left(wdDoc.content.text, 5000)
    Else
        Result = wdDoc.content.text
    End If
    
    wdDoc.Close False
    wdApp.Quit
    
    Set wdDoc = Nothing
    Set wdApp = Nothing
    
    FC_ExtractWordText = Result
    Exit Function
    
ErrHandler:
    On Error Resume Next
    If Not wdDoc Is Nothing Then wdDoc.Close False
    If Not wdApp Is Nothing Then wdApp.Quit
    Set wdDoc = Nothing
    Set wdApp = Nothing
    On Error GoTo 0
    FC_ExtractWordText = ""
End Function

' 텍스트 파일 읽기
Private Function FC_ReadTextFile(ByVal filePath As String) As String
    On Error GoTo ErrHandler
    
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2  ' adTypeText
    stm.Charset = "utf-8"
    stm.Open
    stm.LoadFromFile filePath
    
    Dim Result As String
    Result = stm.ReadText
    stm.Close
    
    If Len(Result) > 5000 Then Result = Left(Result, 5000)
    
    FC_ReadTextFile = Result
    Exit Function
    
ErrHandler:
    FC_ReadTextFile = ""
End Function


' =================================================================================================
'  빈도 분석 및 추천 키워드 생성
' =================================================================================================

' 숫자·연도·날짜류 단어 판정 (키워드에서 제외): 숫자만, 2026년/2026년도/9월/25일/1차/2회/3기/1분기/4학년/2학기, 20260925 등
Public Function FC_IsNumberLike(ByVal W As String) As Boolean
    Dim re As Object
    If Len(W) = 0 Then FC_IsNumberLike = True: Exit Function
    If Not W Like "*[!0-9]*" Then FC_IsNumberLike = True: Exit Function
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "^\d+(년도|년|월|일|차|회|기|호|번|학년|학기|분기|주차|주|시|분|초|명|개|건|차시|p|P|장|쪽|차년도|기준)$"
    FC_IsNumberLike = re.Test(W)
End Function

' 단어 빈도 수집
Private Sub FC_CollectWordFreq(ByVal text As String, dict As Object)
    Dim cleaned As String
    cleaned = FC_CleanText(text)
    
    Dim words() As String
    words = Split(cleaned, " ")
    
    Dim W As Variant
    For Each W In words
        If Len(W) >= 2 And Not FC_IsNumberLike(CStr(W)) Then  ' 2글자 이상, 숫자·연도·날짜 제외
            If dict.Exists(CStr(W)) Then
                dict(CStr(W)) = dict(CStr(W)) + 1
            Else
                dict.Add CStr(W), 1
            End If
        End If
    Next W
End Sub

' 텍스트 정제 (특수문자 제거)
Private Function FC_CleanText(ByVal text As String) As String
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    
    With regEx
        .Global = True
        .IgnoreCase = True
        ' 한글, 영문, 숫자만 남기고 나머지는 공백으로
        .Pattern = "[^가-힣a-zA-Z0-9]"
        text = .Replace(text, " ")
        ' 숫자만으로 된 단어 제거 (날짜, 번호, 연도 등)
        .Pattern = "(^|\s)\d+(?=\s|$)"
        text = .Replace(text, " ")
        ' 연속 공백 정리
        .Pattern = "\s+"
        text = .Replace(text, " ")
    End With
    
    FC_CleanText = Trim(text)
End Function

' 미분류 항목에 추천 키워드 채우기
Private Sub FC_FillRecommendations(ws As Worksheet, lastRow As Long, globalFreq As Object, kwArr() As String)
    ' 전체 빈도 상위 5개 단어를 D2에 표시하고, E열이 비어 있는 행마다
    ' 그 항목 이름에 들어 있는 단어 중 전체에서 가장 자주 나온 단어(2회 이상)를 분류 폴더명으로 추천한다.
    Dim i As Long, r As Long
    Dim topWords As String, topCount As Long
    Dim SortedKeys As Variant
    Dim nm As String, words() As String, W As Variant, best As String, bestN As Long
    If globalFreq.Count = 0 Then Exit Sub

    For i = LBound(kwArr) To UBound(kwArr)
        If Len(kwArr(i)) > 0 Then
            If globalFreq.Exists(kwArr(i)) Then globalFreq.Remove kwArr(i)
        End If
    Next i

    SortedKeys = FC_SortDictByValue(globalFreq)
    For i = LBound(SortedKeys) To UBound(SortedKeys)
        If topCount >= 5 Then Exit For
        If globalFreq(SortedKeys(i)) >= 2 Then
            If topWords = "" Then topWords = CStr(SortedKeys(i)) Else topWords = topWords & ", " & CStr(SortedKeys(i))
            topCount = topCount + 1
        End If
    Next i
    If topWords = "" Then topWords = "추천없음"
    ws.Range("F2").value = "자동 추천 키워드(빈도순): " & topWords

    For r = 4 To lastRow
        If Trim(CStr(ws.Cells(r, "E").value)) = "" Then
            nm = Trim(CStr(ws.Cells(r, "C").value))
            If nm <> "" Then nm = fso_FC.GetBaseName(nm)   ' 확장자 제외
            If nm = "" Then nm = Trim(CStr(ws.Cells(r, "B").value))
            best = "": bestN = 1
            words = Split(FC_CleanText(nm), " ")
            For Each W In words
                If Len(W) >= 2 And Not FC_IsNumberLike(CStr(W)) Then
                    If globalFreq.Exists(CStr(W)) Then
                        If globalFreq(CStr(W)) > bestN Then
                            best = CStr(W): bestN = globalFreq(CStr(W))
                        End If
                    End If
                End If
            Next W
            If best <> "" And UI_ValidFolderName(best) Then
                ws.Cells(r, "E").value = best
                ws.Cells(r, "I").value = "이름에 '" & best & "' 포함 (전체 " & bestN & "회)"
                ws.Cells(r, "E").Interior.color = RGB(223, 245, 228)  ' 연한 녹색 = 자동 추천
            Else
                ws.Cells(r, "E").value = "미분류"
                ws.Cells(r, "I").value = "이름 공통 단어 없음 (전체 추천: " & topWords & ")"
                ws.Cells(r, "E").Interior.color = RGB(255, 244, 214)  ' 연한 파란색
            End If
        End If
    Next r
End Sub


' =================================================================================================
'  유틸리티 함수
' =================================================================================================

' Dictionary를 값 기준 내림차순 정렬
Private Function FC_SortDictByValue(dict As Object) As Variant
    Dim keys As Variant, values As Variant
    keys = dict.keys
    values = dict.items
    
    Dim i As Long, j As Long
    Dim tempKey As Variant, tempValue As Variant
    
    For i = LBound(keys) To UBound(keys) - 1
        For j = i + 1 To UBound(keys)
            If values(i) < values(j) Then  ' 내림차순
                tempValue = values(i): values(i) = values(j): values(j) = tempValue
                tempKey = keys(i): keys(i) = keys(j): keys(j) = tempKey
            End If
        Next j
    Next i
    
    FC_SortDictByValue = keys
End Function

' 시트 생성 또는 가져오기
Private Function FC_GetOrCreateSheet(sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ActiveWorkbook.Sheets(sheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count))
        ws.name = sheetName
    End If
    
    Set FC_GetOrCreateSheet = ws
End Function

' 동일 파일명 충돌 방지 (고유 경로 반환)
Private Function FC_GetUniquePath(ByVal originalPath As String, isFolder As Boolean) As String
    If isFolder Then
        If Not fso_FC.FolderExists(originalPath) Then
            FC_GetUniquePath = originalPath
            Exit Function
        End If
        
        Dim fCount As Long: fCount = 1
        Dim newFolderPath As String
        Do
            newFolderPath = originalPath & "_" & fCount
            If Not fso_FC.FolderExists(newFolderPath) Then
                FC_GetUniquePath = newFolderPath
                Exit Function
            End If
            fCount = fCount + 1
        Loop
    Else
        If Not fso_FC.FileExists(originalPath) Then
            FC_GetUniquePath = originalPath
            Exit Function
        End If
        
        Dim baseName As String
        Dim extPart As String
        Dim parentDir As String
        
        parentDir = fso_FC.GetParentFolderName(originalPath)
        baseName = fso_FC.GetBaseName(originalPath)
        extPart = fso_FC.GetExtensionName(originalPath)
        
        Dim cnt As Long: cnt = 1
        Dim newPath As String
        Do
            If extPart <> "" Then
                newPath = fso_FC.BuildPath(parentDir, baseName & "_" & cnt & "." & extPart)
            Else
                newPath = fso_FC.BuildPath(parentDir, baseName & "_" & cnt)
            End If
            
            If Not fso_FC.FileExists(newPath) Then
                FC_GetUniquePath = newPath
                Exit Function
            End If
            cnt = cnt + 1
        Loop
    End If
End Function

' 1행 버튼(왼쪽 정렬)
'   1줄(파랑): [분류 확인 후 이동] [기준폴더 설정] [기준폴더 이동]
'   2줄(주황): [년월일 키워드 변경] [키워드<>월일 변경] [년월일<>키워드 변경] [미분류 → 오늘날짜 추천]
'   3줄      : [이름변경 실행] [이름변경 후 기준폴더 이동]  + 설명 도형
Private Sub FC_CreateMoveButton(ws As Worksheet)
    Dim s As Shape, nm As Variant, note As String
    On Error Resume Next
    For Each nm In Array("btnMoveFiles", "btnBaseFolder", "btnBaseMove", "btnDateKw", "btnKwMD", "btnYmdKw", "btnTodayKw", "btnRenameRun", "btnRenameMove", "btnKwDateMove", "btnReselect", "noteResult")
        ws.Shapes(CStr(nm)).Delete
    Next nm
    On Error GoTo 0
    ws.rows(1).RowHeight = 176
    Set s = ws.Shapes.AddShape(msoShapeRoundedRectangle, 6, 6, 150, 30)
    s.name = "btnMoveFiles": UI_StyleButton s, "분류 확인 후 이동", "FC_파일이동실행", True
    Set s = ws.Shapes.AddShape(msoShapeRoundedRectangle, 164, 6, 130, 30)
    s.name = "btnBaseFolder": UI_StyleButton s, "기준폴더 설정", "FC_기준폴더설정", False
    Set s = ws.Shapes.AddShape(msoShapeRoundedRectangle, 302, 6, 130, 30)
    s.name = "btnBaseMove": UI_StyleButton s, "기준폴더 이동", "FC_기준폴더이동", True
    Set s = ws.Shapes.AddShape(msoShapeRoundedRectangle, 440, 6, 130, 30)
    s.name = "btnReselect": UI_StyleButton s, "폴더 다시 선택", "FC_폴더다시선택", False   ' 2.7k

    Set s = ws.Shapes.AddShape(msoShapeRoundedRectangle, 6, 44, 150, 30)
    s.name = "btnDateKw": UI_StyleButton s, "년월일 키워드 변경", "FC_날짜키워드변경", True, FC_ORANGE, RGB(255, 255, 255)
    Set s = ws.Shapes.AddShape(msoShapeRoundedRectangle, 164, 44, 160, 30)
    s.name = "btnKwMD": UI_StyleButton s, "키워드<>월일 변경", "FC_키워드월일변경", True, FC_ORANGE, RGB(255, 255, 255)
    Set s = ws.Shapes.AddShape(msoShapeRoundedRectangle, 332, 44, 170, 30)
    s.name = "btnYmdKw": UI_StyleButton s, "년월일<>키워드 변경", "FC_년월일키워드변경", True, FC_ORANGE, RGB(255, 255, 255)
    Set s = ws.Shapes.AddShape(msoShapeRoundedRectangle, 510, 44, 190, 30)
    s.name = "btnTodayKw": UI_StyleButton s, "미분류 → 오늘날짜 추천", "FC_미분류오늘날짜추천", True, FC_ORANGE, RGB(255, 255, 255)

    Set s = ws.Shapes.AddShape(msoShapeRoundedRectangle, 6, 82, 150, 30)
    s.name = "btnRenameRun": UI_StyleButton s, "이름변경 실행", "FC_이름변경실행", False, RGB(255, 236, 204), RGB(160, 82, 0)
    Set s = ws.Shapes.AddShape(msoShapeRoundedRectangle, 164, 82, 220, 30)
    s.name = "btnRenameMove": UI_StyleButton s, "이름변경 후 기준폴더 이동", "FC_이름변경후이동", False, RGB(255, 236, 204), RGB(160, 82, 0)

    note = "파랑: E열 '이동 폴더' 확인 → [분류 확인 후 이동](분석 폴더 안) / [기준폴더 이동](기준폴더 안). '미분류'는 이동하지 않음" & vbCrLf & _
           "주황: 이름 속 날짜 정규화 / 키워드<>월일·년월일<>키워드는 누를 때마다 형식 순환(1~4/4) → F열 '이름 변경(안)' / 미분류는 오늘 날짜 폴더 추천" & vbCrLf & _
           "실행: [이름변경 실행] F열대로 실제 이름 변경  ·  [이름변경 후 기준폴더 이동] F열 반영 후 이동    기준폴더: " & IIf(FC_BasePath() = "", "(미설정)", FC_BasePath())
    FC_AddNote ws, "noteResult", note, 6, 120, 1040, 50
    ws.Range("A1").value = ""
End Sub

' =================================================================================================
'  [날짜 키워드] 이름 속 날짜 토큰 → YYYYMMDD 정규화.  20260926 / 2026.09.26 / 2026-09-26 / 260926 / 0926(올해)
'   반환: 바뀐 이름, found = 찾은 날짜(YYYYMMDD, 없으면 "")
' =================================================================================================
Public Function FC_NormalizeDate(ByVal nm As String, ByRef found As String) As String
    Dim re As Object, m As Object, y As String, mo As String, d As String, tok As String
    found = ""
    FC_NormalizeDate = nm
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    ' 1) 2026.09.26 / 2026-09-26 / 2026_09_26 / 20260926
    re.Pattern = "(^|[^0-9])((19|20)\d{2})[.\-_/ ]?(0[1-9]|1[0-2])[.\-_/ ]?(0[1-9]|[12][0-9]|3[01])(?![0-9])"
    If re.Test(nm) Then
        Set m = re.Execute(nm)(0)
        y = m.SubMatches(1): mo = m.SubMatches(3): d = m.SubMatches(4)
        found = y & mo & d
        tok = Mid$(m.value, Len(m.SubMatches(0)) + 1)
        FC_NormalizeDate = Replace(nm, tok, found, 1, 1)
        Exit Function
    End If
    ' 2) 260926 (YYMMDD)
    re.Pattern = "(^|[^0-9])(\d{2})(0[1-9]|1[0-2])(0[1-9]|[12][0-9]|3[01])(?![0-9])"
    If re.Test(nm) Then
        Set m = re.Execute(nm)(0)
        found = "20" & m.SubMatches(1) & m.SubMatches(2) & m.SubMatches(3)
        tok = Mid$(m.value, Len(m.SubMatches(0)) + 1)
        FC_NormalizeDate = Replace(nm, tok, found, 1, 1)
        Exit Function
    End If
    ' 3) 0926 (MMDD → 올해)
    re.Pattern = "(^|[^0-9])(0[1-9]|1[0-2])(0[1-9]|[12][0-9]|3[01])(?![0-9])"
    If re.Test(nm) Then
        Set m = re.Execute(nm)(0)
        found = Format$(Date, "yyyy") & m.SubMatches(1) & m.SubMatches(2)
        tok = Mid$(m.value, Len(m.SubMatches(0)) + 1)
        FC_NormalizeDate = Replace(nm, tok, found, 1, 1)
    End If
End Function

' 행의 항목 이름(폴더명 또는 파일 기본명)과 확장자
Private Sub FC_RowName(ByVal ws As Worksheet, ByVal r As Long, ByRef nm As String, ByRef ext As String, ByRef isFolder As Boolean)
    isFolder = (CStr(ws.Cells(r, "D").value) = "폴더")
    If isFolder Then
        nm = Trim$(CStr(ws.Cells(r, "B").value)): ext = ""
    Else
        nm = fso_FC.GetBaseName(Trim$(CStr(ws.Cells(r, "C").value)))
        ext = fso_FC.GetExtensionName(Trim$(CStr(ws.Cells(r, "C").value)))
    End If
End Sub

' [주황1] 년월일 키워드 변경: 이름 속 날짜를 YYYYMMDD 로 바꾼 이름을 F열 '이름 변경(안)'에, E열이 날짜 키워드면 E열도 정규화
Public Sub FC_날짜키워드변경()
    Dim ws As Worksheet, lastRow As Long, r As Long, nm As String, ext As String, isFolder As Boolean
    Dim newNm As String, found As String, cnt As Long, eVal As String, eNew As String, f2 As String
    Set ws = FC_ResultSheet()
    If ws Is Nothing Then Exit Sub
    lastRow = ws.Cells(ws.rows.Count, "A").End(xlUp).row
    Application.ScreenUpdating = False
    For r = 4 To lastRow
        If Left$(CStr(ws.Cells(r, "G").value), 2) = "성공" Then GoTo NextRow
        FC_RowName ws, r, nm, ext, isFolder
        If nm = "" Then GoTo NextRow
        newNm = FC_NormalizeDate(nm, found)
        eVal = Trim$(CStr(ws.Cells(r, "E").value))
        If eVal <> "" And eVal <> "미분류" Then
            eNew = FC_NormalizeDate(eVal, f2)
            If f2 <> "" And eNew <> eVal Then ws.Cells(r, "E").value = eNew
        End If
        If found <> "" And newNm <> nm Then
            ws.Cells(r, "F").value = newNm
            ws.Cells(r, "F").Interior.color = RGB(255, 244, 214)
            cnt = cnt + 1
        End If
NextRow:
    Next r
    Application.ScreenUpdating = True
    MsgBox "이름 속 날짜를 년월일(YYYYMMDD)로 바꾼 이름 " & cnt & "개를 F열 '이름 변경(안)'에 넣었습니다." & vbCrLf & _
           "확인·수정 후 [이름변경 실행]을 누르면 실제 이름이 바뀝니다.", vbInformation, "년월일 키워드 변경"
End Sub

' [주황2-a] 키워드<>월일 변경 (클릭할 때마다 형식 순환, 월일 4자리):
'   1회 '키워드 월일' → 2회 '월일 키워드' → 3회 '월_일 키워드' → 4회 '키워드 월_일'  → F열 '이름 변경(안)'
Public Sub FC_키워드월일변경()
    FC_CycleNameFormat "MD"
End Sub

' [주황2-b] 년월일<>키워드 변경 (클릭할 때마다 형식 순환):
'   1회 '260926 키워드' → 2회 '키워드-260926' → 3회 '20260926 키워드' → 4회 '키워드-20260926'  → F열
Public Sub FC_년월일키워드변경()
    FC_CycleNameFormat "YMD"
End Sub

' 형식 순환 공통. 상태(현재 몇 번째 형식인지)는 H2(월일) / I2(년월일) 셀에 보관
Private Sub FC_CycleNameFormat(ByVal mode As String)
    Dim ws As Worksheet, lastRow As Long, r As Long, nm As String, ext As String, isFolder As Boolean
    Dim kw As String, found As String, cnt As Long, idx As Long, stateCell As String, newNm As String, lbl As String, btn As String, st As String
    Set ws = FC_ResultSheet()
    If ws Is Nothing Then Exit Sub
    If mode = "MD" Then
        stateCell = "H2": btn = "btnKwMD"
    Else
        stateCell = "I2": btn = "btnYmdKw"
    End If
    st = CStr(ws.Range(stateCell).value)
    If InStr(st, ":") > 0 Then idx = val(Mid$(st, InStr(st, ":") + 1)) Else idx = 0
    idx = (idx Mod 4) + 1
    lastRow = ws.Cells(ws.rows.Count, "A").End(xlUp).row
    Application.ScreenUpdating = False
    For r = 4 To lastRow
        If Left$(CStr(ws.Cells(r, "G").value), 2) = "성공" Then GoTo NextRow
        kw = FC_RowKeyword(ws, r)
        If kw = "" Then GoTo NextRow
        FC_RowName ws, r, nm, ext, isFolder
        If nm = "" Then GoTo NextRow
        FC_NormalizeDate nm, found
        If found = "" Then found = Format$(Date, "yyyymmdd")
        newNm = FC_FormatName(mode, idx, kw, found)
        ws.Cells(r, "F").value = newNm
        ws.Cells(r, "F").Interior.color = RGB(255, 244, 214)
        cnt = cnt + 1
NextRow:
    Next r
    lbl = FC_FormatLabel(mode, idx)
    ws.Range(stateCell).value = IIf(mode = "MD", "월일형식:", "년월일형식:") & idx
    ws.Range(stateCell).Font.color = RGB(142, 142, 147)
    ws.Range(stateCell).Font.Size = 9
    On Error Resume Next
    ws.Shapes(btn).TextFrame2.TextRange.text = IIf(mode = "MD", "키워드<>월일 ", "년월일<>키워드 ") & idx & "/4"
    On Error GoTo 0
    Application.ScreenUpdating = True
    Application.StatusBar = "이름 형식 " & idx & "/4: " & lbl & "  (" & cnt & "개 → F열 '이름 변경(안)')  다시 누르면 다음 형식"
End Sub

' 행의 키워드: E열 이동 폴더(미분류·날짜가 아니면) → 없으면 H열 분석 추천 → 없으면 ""
Private Function FC_RowKeyword(ByVal ws As Worksheet, ByVal r As Long) As String
    Dim e As String
    e = Trim$(CStr(ws.Cells(r, "E").value))
    If e <> "" And e <> "미분류" And Not FC_IsNumberLike(e) Then FC_RowKeyword = e: Exit Function
    e = Trim$(CStr(ws.Cells(r, "H").value))
    If e <> "" And Not FC_IsNumberLike(e) Then FC_RowKeyword = e: Exit Function
    ' 마지막 수단: 이름에서 날짜 토큰을 뺀 나머지
    Dim nm As String, ext As String, isFolder As Boolean, found As String
    FC_RowName ws, r, nm, ext, isFolder
    nm = FC_NormalizeDate(nm, found)
    If found <> "" Then nm = Replace(nm, found, "")
    nm = Trim$(nm)
    Do While InStr(nm, "  ") > 0: nm = Replace(nm, "  ", " "): Loop
    If nm <> "" And Not FC_IsNumberLike(nm) Then FC_RowKeyword = nm
End Function

' 형식별 이름 만들기. ymd = YYYYMMDD
Private Function FC_FormatName(ByVal mode As String, ByVal idx As Long, ByVal kw As String, ByVal ymd As String) As String
    Dim md As String, m_d As String
    md = Right$(ymd, 4): m_d = Mid$(ymd, 5, 2) & "_" & Right$(ymd, 2)
    If mode = "MD" Then
        Select Case idx
            Case 1: FC_FormatName = kw & " " & md
            Case 2: FC_FormatName = md & " " & kw
            Case 3: FC_FormatName = m_d & " " & kw
            Case Else: FC_FormatName = kw & " " & m_d
        End Select
    Else
        Select Case idx
            Case 1: FC_FormatName = Right$(ymd, 6) & " " & kw
            Case 2: FC_FormatName = kw & "-" & Right$(ymd, 6)
            Case 3: FC_FormatName = ymd & " " & kw
            Case Else: FC_FormatName = kw & "-" & ymd
        End Select
    End If
End Function

Private Function FC_FormatLabel(ByVal mode As String, ByVal idx As Long) As String
    If mode = "MD" Then
        FC_FormatLabel = Choose(idx, "키워드 월일", "월일 키워드", "월_일 키워드", "키워드 월_일")
    Else
        FC_FormatLabel = Choose(idx, "년월일(6자리) 키워드", "키워드-년월일(6자리)", "년월일(8자리) 키워드", "키워드-년월일(8자리)")
    End If
End Function

' [주황3] 미분류 → 오늘날짜 추천: E열이 '미분류'인 행에 오늘 날짜(YYYYMMDD) 폴더를 추천
Public Sub FC_미분류오늘날짜추천()
    Dim ws As Worksheet, lastRow As Long, r As Long, cnt As Long, today As String
    Set ws = FC_ResultSheet()
    If ws Is Nothing Then Exit Sub
    today = Format$(Date, "yyyymmdd")
    lastRow = ws.Cells(ws.rows.Count, "A").End(xlUp).row
    For r = 4 To lastRow
        If Trim$(CStr(ws.Cells(r, "E").value)) = "미분류" Then
            ws.Cells(r, "E").value = today
            ws.Cells(r, "E").Interior.color = RGB(255, 236, 204)
            ws.Cells(r, "I").value = "미분류 → 오늘 날짜(" & today & ") 폴더 추천"
            cnt = cnt + 1
        End If
    Next r
    MsgBox "미분류 " & cnt & "개 항목의 이동 폴더를 오늘 날짜 '" & today & "' 로 추천했습니다." & vbCrLf & _
           "E열을 확인한 뒤 [분류 확인 후 이동] 또는 [기준폴더 이동]을 누르세요.", vbInformation, "미분류 → 오늘날짜 추천"
End Sub

' [주황4] 이름변경 실행: F열 '이름 변경(안)'대로 실제 폴더/파일 이름을 바꾼다 (분석 폴더 안)
Public Sub FC_이름변경실행()
    Dim ws As Worksheet
    Set ws = FC_ResultSheet()
    If ws Is Nothing Then Exit Sub
    FC_ApplyRenames ws, True
End Sub

' [주황5] 이름변경 후 기준폴더 이동: F열대로 이름을 바꾼 뒤(없으면 그대로) 기준폴더로 이동
Public Sub FC_이름변경후이동()
    Dim ws As Worksheet, basePath As String, r As Long, lastRow As Long, cnt As Long
    Set ws = FC_ResultSheet()
    If ws Is Nothing Then Exit Sub
    basePath = FC_BasePath()
    If basePath = "" Then
        MsgBox "'기준폴더' 시트(기준폴더 경로)가 없습니다. 먼저 [기준폴더 설정]을 하세요.", vbExclamation, "이름변경 후 기준폴더 이동"
        Exit Sub
    End If
    lastRow = ws.Cells(ws.rows.Count, "A").End(xlUp).row
    For r = 4 To lastRow
        If Trim$(CStr(ws.Cells(r, "F").value)) <> "" Then cnt = cnt + 1
    Next r
    If cnt > 0 Then
        If UI_Choose("이름변경 후 기준폴더 이동", "F열 '이름 변경(안)' " & cnt & "개를 실제 이름에 반영한 뒤, E열 이동 폴더대로 기준폴더로 이동합니다.", Array("이름 변경 후 이동 실행", "취소")) <> 1 Then Exit Sub
        FC_ApplyRenames ws, False
    End If
    FC_기준폴더이동
End Sub

' F열대로 이름 변경. 반환: 성공 개수 (결과는 G열)
Public Function FC_ApplyRenames(ByVal ws As Worksheet, ByVal showMsg As Boolean) As Long
    Dim srcBase As String, lastRow As Long, r As Long, k As Long, nm As String, ext As String, isFolder As Boolean
    Dim newNm As String, oldPath As String, newPath As String, okCnt As Long, failCnt As Long, total As Long, oldFolder As String
    srcBase = FC_SourcePath(ws)
    If srcBase = "" Then Exit Function
    lastRow = ws.Cells(ws.rows.Count, "A").End(xlUp).row
    For r = 4 To lastRow
        If Trim$(CStr(ws.Cells(r, "F").value)) <> "" Then total = total + 1
    Next r
    If total = 0 Then
        If showMsg Then MsgBox "F열 '이름 변경(안)'이 비어 있습니다." & vbCrLf & "[년월일 키워드 변경] 등으로 안을 만들거나 직접 입력하세요.", vbInformation, "이름변경 실행"
        Exit Function
    End If
    If showMsg Then
        If UI_Choose("이름변경 실행", total & "개 항목의 실제 이름을 F열대로 변경합니다. (분석 폴더: " & srcBase & ")", Array("이름 변경 실행", "취소")) <> 1 Then Exit Function
    End If
    Application.ScreenUpdating = False
    For r = 4 To lastRow
        newNm = Trim$(CStr(ws.Cells(r, "F").value))
        If newNm = "" Then GoTo NextRow
        If Left$(CStr(ws.Cells(r, "G").value), 2) = "성공" Then GoTo NextRow
        FC_RowName ws, r, nm, ext, isFolder
        If Not UI_ValidFolderName(newNm) Then
            FC_SetResult ws, r, "이름변경 실패: 사용할 수 없는 이름", RGB(255, 200, 200): failCnt = failCnt + 1: GoTo NextRow
        End If
        If isFolder Then
            oldPath = fso_FC.BuildPath(srcBase, nm)
            newPath = fso_FC.BuildPath(srcBase, newNm)
        Else
            oldFolder = Trim$(CStr(ws.Cells(r, "B").value))
            If oldFolder <> "" Then
                oldPath = fso_FC.BuildPath(fso_FC.BuildPath(srcBase, oldFolder), nm & IIf(ext <> "", "." & ext, ""))
                newPath = fso_FC.BuildPath(fso_FC.BuildPath(srcBase, oldFolder), newNm & IIf(ext <> "", "." & ext, ""))
            Else
                oldPath = fso_FC.BuildPath(srcBase, nm & IIf(ext <> "", "." & ext, ""))
                newPath = fso_FC.BuildPath(srcBase, newNm & IIf(ext <> "", "." & ext, ""))
            End If
        End If
        If StrComp(oldPath, newPath, vbTextCompare) = 0 Then
            ws.Cells(r, "F").value = "": GoTo NextRow
        End If
        On Error Resume Next
        If isFolder Then
            If fso_FC.FolderExists(newPath) Then Err.Raise 58, , "같은 이름의 폴더가 이미 있습니다"
            fso_FC.MoveFolder oldPath, newPath
        Else
            If fso_FC.FileExists(newPath) Then Err.Raise 58, , "같은 이름의 파일이 이미 있습니다"
            fso_FC.MoveFile oldPath, newPath
        End If
        If Err.Number = 0 Then
            okCnt = okCnt + 1
            If isFolder Then
                ws.Cells(r, "B").value = newNm
                For k = 4 To lastRow      ' 폴더 안 파일 행의 폴더명 갱신
                    If k <> r And CStr(ws.Cells(k, "D").value) <> "폴더" And StrComp(CStr(ws.Cells(k, "B").value), nm, vbTextCompare) = 0 Then ws.Cells(k, "B").value = newNm
                Next k
            Else
                ws.Cells(r, "C").value = newNm & IIf(ext <> "", "." & ext, "")
            End If
            ws.Cells(r, "F").value = ""
            ws.Cells(r, "F").Interior.ColorIndex = xlNone
            If Left$(CStr(ws.Cells(r, "G").value), 2) <> "성공" Then FC_SetResult ws, r, "이름변경 완료 → " & newNm, RGB(223, 245, 228)
        Else
            FC_SetResult ws, r, "이름변경 실패: " & Err.Description, RGB(255, 200, 200)
            failCnt = failCnt + 1
            Err.Clear
        End If
        On Error GoTo 0
NextRow:
    Next r
    Application.ScreenUpdating = True
    FC_ApplyRenames = okCnt
    If showMsg Then MsgBox "이름 변경 완료: 성공 " & okCnt & "개 / 실패 " & failCnt & "개" & vbCrLf & "(G열에서 결과 확인)", vbInformation, "이름변경 실행"
End Function

' 설명 도형(애플 스타일: 연회색 바탕, 테두리 없음, 왼쪽 정렬)
Public Sub FC_AddNote(ByVal ws As Worksheet, ByVal shapeName As String, ByVal text As String, _
                      ByVal x As Single, ByVal y As Single, ByVal W As Single, ByVal h As Single)
    Dim s As Shape
    On Error Resume Next
    ws.Shapes(shapeName).Delete
    On Error GoTo 0
    Set s = ws.Shapes.AddShape(msoShapeRoundedRectangle, x, y, W, h)
    With s
        .name = shapeName
        .Adjustments(1) = 0.15
        .Placement = xlFreeFloating
        .line.Visible = msoFalse
        .Shadow.Visible = msoFalse
        .Fill.Solid
        .Fill.ForeColor.RGB = RGB(242, 242, 247)
        With .TextFrame2
            .AutoSize = msoAutoSizeNone
            .WordWrap = msoTrue
            .MarginLeft = 10: .MarginRight = 10: .MarginTop = 4: .MarginBottom = 4
            .VerticalAnchor = msoAnchorMiddle
            With .TextRange
                .text = text
                .Font.name = "맑은 고딕"
                .Font.NameFarEast = "맑은 고딕"
                .Font.Size = 9.5
                .Font.Bold = msoFalse
                .Font.Fill.ForeColor.RGB = RGB(60, 60, 67)
                .ParagraphFormat.Alignment = msoAlignLeft
            End With
        End With
    End With
End Sub

' 표머리 스타일(애플): 연회색 바탕, 진한 글자, 아래 테두리
Public Sub FC_StyleHeader(ByVal rng As Range)
    With rng
        .Font.name = "맑은 고딕"
        .Font.Size = 11
        .Font.Bold = True
        .Font.color = RGB(28, 28, 30)
        .Interior.color = RGB(242, 242, 247)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 24
        .Borders.LineStyle = xlContinuous
        .Borders.color = RGB(209, 209, 214)
        .Borders(xlEdgeBottom).weight = xlMedium
        .Borders(xlEdgeBottom).color = RGB(174, 174, 178)
    End With
End Sub

' 분석 유형별 열 배치: 폴더=파일명·확장자 숨김 / 파일=폴더명·확장자 숨김 / 폴더및파일=모두 표시
Private Sub FC_LayoutResult(ByVal ws As Worksheet, ByVal analysisType As String)
    ws.Columns("A").ColumnWidth = 6
    ws.Columns("B").ColumnWidth = 46
    ws.Columns("C").ColumnWidth = 52
    ws.Columns("D").ColumnWidth = 9
    ws.Columns("E").ColumnWidth = 26      ' 이동 폴더
    ws.Columns("F").ColumnWidth = 36      ' 이름 변경(안)
    ws.Columns("G").ColumnWidth = 30      ' 이동결과
    ws.Columns("H").ColumnWidth = 16      ' 분석 추천(참고)
    ws.Columns("I").ColumnWidth = 46      ' 추천 근거
    ws.Columns("B:I").hidden = False
    Select Case analysisType
        Case "폴더": ws.Columns("C:D").hidden = True
        Case "파일"
            ws.Columns("B").hidden = True: ws.Columns("D").hidden = True
            ws.Range("C2").value = ws.Range("B2").value
            On Error Resume Next
            ws.Hyperlinks.Add anchor:=ws.Range("C2"), Address:=CStr(ws.Range("B2").value), TextToDisplay:=CStr(ws.Range("B2").value)
            On Error GoTo 0
    End Select
    ws.Range("A2:G2").Font.name = "맑은 고딕"
    ws.Range("A2:G2").Font.Size = 10
    ws.Range("A2:G2").Font.color = RGB(99, 99, 102)
    ws.Range("A2").Font.Bold = True
    ws.rows(2).RowHeight = 20
End Sub

' 서식 적용(애플 스타일): 얇은 연회색 테두리, 홀수 행 연한 파랑 / 짝수 행 흰색, 폴더=맑은 고딕 굵게 · 파일=굴림
Private Sub FC_ApplyFormatting(ws As Worksheet, lastRow As Long)
    Dim dataRange As Range, r As Long
    If lastRow < 4 Then Exit Sub
    Set dataRange = ws.Range("A4:I" & lastRow)
    With dataRange
        .Font.name = "맑은 고딕"
        .Font.Size = 11
        .Font.Bold = False
        .Font.color = RGB(28, 28, 30)
        .Interior.ColorIndex = xlNone
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.color = RGB(220, 220, 224)
        .RowHeight = 21
    End With
    ws.Range("A4:A" & lastRow).HorizontalAlignment = xlCenter
    ws.Range("D4:D" & lastRow).HorizontalAlignment = xlCenter
    ws.Range("A4:A" & lastRow).Font.color = RGB(142, 142, 147)
    For r = 4 To lastRow
        If (r - 3) Mod 2 = 1 Then ws.Range("A" & r & ":I" & r).Interior.color = RGB(239, 245, 255)   ' 홀수 행 연한 파랑
        If ws.Cells(r, "D").value = "폴더" Then
            ws.Range("B" & r & ":C" & r).Font.name = "맑은 고딕"
            ws.Range("B" & r & ":C" & r).Font.Bold = True
        Else
            ws.Range("B" & r & ":C" & r).Font.name = "굴림"
        End If
    Next r
End Sub

