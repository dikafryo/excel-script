Attribute VB_Name = "modMergePdf"

Option Explicit

'// =============================================================================
'// PDF 합치기 + 목차 생성 모듈 (MS Word 활용) v2.1
'// =============================================================================
'// [기능]
'//   - 폴더 내 PDF를 MS Word로 병합 후 PDF 내보내기
'//   - 목차: 파일명(확장자 제거) + 점선리더 + 페이지번호, 클릭 시 해당 페이지로 이동(하이퍼링크 + 북마크)
'//   - [v2.6] MP_Run(tocMode) : 0 물어봄 / 1 목차 포함 / 2 목차 없음
'//            MP_MergeFiles(...) : 파일 목록을 받아 합치는 공용 코어 (hwp→pdf 하나로 합치기에서도 사용)
'// [필요 환경] MS Office 2013+ (Word의 PDF 변환 기능)
'// =============================================================================

'// ─── Word 상수 (Late Binding) ───────────────────────────────────────────────
Private Const WD_PAGE_BREAK         As Long = 7
Private Const WD_ALIGN_CENTER       As Long = 1
Private Const WD_ALIGN_LEFT         As Long = 0
Private Const WD_ALIGN_TAB_RIGHT    As Long = 2
Private Const WD_TAB_LEADER_DOTS    As Long = 1
Private Const WD_EXPORT_PDF         As Long = 17
Private Const WD_CREATE_BOOKMARKS   As Long = 2
Private Const WD_FIELD_EMPTY        As Long = -1
Private Const WD_STORY              As Long = 6
Private Const WD_COLLAPSE_END       As Long = 0
Private Const WD_ALERTS_NONE        As Long = 0
Private mLastErr As String


'// =============================================================================
'//  Main : PDF 합치기 + 목차 생성 (구형 진입점: 목차 여부를 물어봄)
'// =============================================================================
Public Sub eMergePdfsInFolder(control As IRibbonControl)
    MP_Run 0
End Sub

Public Sub MP_Run(ByVal tocMode As Long)
    Dim folderPath  As String
    Dim fso         As Object
    Dim pdfPaths()  As String
    Dim pdfNames()  As String
    Dim fileCount   As Long
    Dim wantTOC     As Boolean
    Dim savePath    As String
    Dim startTime   As Single
    Dim tocChoice   As Long

    startTime = Timer

    '// ── 1. 폴더 선택 ──
    With Application.FileDialog(4)
        .title = "PDF 파일이 있는 폴더를 선택하세요"
        If .Show <> -1 Then Exit Sub
        folderPath = .SelectedItems(1)
    End With
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"

    '// ── 2. PDF 파일 수집 및 정렬 ──
    Set fso = CreateObject("Scripting.FileSystemObject")
    fileCount = CollectPDFs(fso, folderPath, pdfPaths, pdfNames)

    If fileCount = 0 Then
        MsgBox "폴더에 PDF 파일이 없습니다.", vbExclamation, "알림"
        Exit Sub
    End If

    '// ── 3. 목차 생성 여부 ──
    Select Case tocMode
        Case 1: wantTOC = True
        Case 2: wantTOC = False
        Case Else
            tocChoice = UI_Choose("PDF 합치기", "PDF " & fileCount & "개를 이름순으로 하나로 합칩니다. 앞에 목차 페이지를 넣을까요?" & vbCrLf & _
                                  "(MS Word 로 변환해 합치므로 스캔본·복잡한 서식은 모양이 달라질 수 있습니다)", Array("목차 페이지 포함", "목차 없이 합치기", "취소"))
            If tocChoice = 0 Or tocChoice = 3 Then Exit Sub
            wantTOC = (tocChoice = 1)
    End Select

    If wantTOC Then
        savePath = folderPath & "병합_목차_" & Format(Now, "YYYYMMDD_HHmmss") & ".pdf"
    Else
        savePath = folderPath & "병합_" & Format(Now, "YYYYMMDD_HHmmss") & ".pdf"
    End If

    If Not MP_MergeFiles(fso, pdfPaths, pdfNames, fileCount, wantTOC, savePath) Then
        MsgBox "오류가 발생했습니다:" & vbCrLf & vbCrLf & mLastErr & vbCrLf & vbCrLf & _
               "MS Word가 정상 설치되어 있는지 확인하세요.", vbCritical, "오류"
        Exit Sub
    End If

    '// ── 엑셀 결과 기록 ──
    Call WriteResultSheet(folderPath, savePath, pdfNames, pdfPaths, fileCount, fso)

    Dim elapsed As Single
    elapsed = Timer - startTime

    If UI_Choose("PDF 합치기 완료", fileCount & "개의 PDF 를 합쳤습니다. (" & Format(elapsed, "0.0") & "초)" & vbCrLf & vbCrLf & _
                 "저장 위치: " & savePath & vbCrLf & "(파일 목록은 'PDF 병합 결과' 시트에 있습니다)", Array("합친 PDF 열기", "닫기")) = 1 Then
        Shell "explorer """ & savePath & """", vbNormalFocus
    End If
End Sub


'// =============================================================================
'//  코어 : PDF 목록(1~fileCount) → Word 로 합쳐 savePath 에 PDF 저장. 실패 시 False (사유 mLastErr)
'// =============================================================================
Public Function MP_MergeFiles(ByVal fso As Object, ByRef pdfPaths() As String, ByRef pdfNames() As String, ByVal fileCount As Long, _
                              ByVal wantTOC As Boolean, ByVal savePath As String) As Boolean
    On Error GoTo ErrHandler
    Dim bmNames()   As String
    Dim wdApp       As Object
    Dim wdDoc       As Object
    Dim wordWasOpen As Boolean
    Dim i           As Long
    Dim endPos      As Long
    Dim rng         As Object

    mLastErr = ""
    '// ── 4. MS Word 실행 ──
    Application.StatusBar = "MS Word를 시작하는 중..."
    wordWasOpen = False
    On Error Resume Next
    Set wdApp = GetObject(, "Word.Application")
    If Not wdApp Is Nothing Then wordWasOpen = True
    On Error GoTo ErrHandler

    If wdApp Is Nothing Then
        Set wdApp = CreateObject("Word.Application")
        wordWasOpen = False
    End If

    '// [성능 핵심] 화면갱신 + 자동페이지계산 OFF
    If Not wordWasOpen Then wdApp.Visible = False   ' 사용자가 쓰던 Word 창을 숨기지 않는다
    wdApp.DisplayAlerts = WD_ALERTS_NONE
    wdApp.ScreenUpdating = False
    wdApp.options.Pagination = False

    '// ── 5. 새 Word 문서 생성 ──
    Set wdDoc = wdApp.Documents.Add

    With wdDoc.PageSetup
        .TopMargin = wdApp.CentimetersToPoints(2)
        .BottomMargin = wdApp.CentimetersToPoints(2)
        .LeftMargin = wdApp.CentimetersToPoints(2)
        .RightMargin = wdApp.CentimetersToPoints(2)
    End With

    '// ── 6. PDF 순차 삽입 + 북마크 ──
    ReDim bmNames(1 To fileCount)

    For i = 1 To fileCount
        If i Mod 5 = 1 Or i = fileCount Then
            Application.StatusBar = "PDF 삽입 중 (" & i & "/" & fileCount & ")  " & pdfNames(i)
            DoEvents
        End If

        bmNames(i) = "BM_" & Format(i, "000")
        endPos = wdDoc.content.End - 1

        If i > 1 Then
            Set rng = wdDoc.Range(endPos, endPos)
            rng.InsertBreak WD_PAGE_BREAK
            endPos = wdDoc.content.End - 1
        End If

        Set rng = wdDoc.Range(endPos, endPos)
        wdDoc.Bookmarks.Add bmNames(i), rng

        On Error Resume Next
        Set rng = wdDoc.Range(endPos, endPos)
        rng.InsertFile fileName:=pdfPaths(i), ConfirmConversions:=False

        If Err.Number <> 0 Then
            Err.Clear
            Dim tempDoc As Object
            Set tempDoc = wdApp.Documents.Open( _
                fileName:=pdfPaths(i), _
                ConfirmConversions:=False, _
                ReadOnly:=True)
            If Not tempDoc Is Nothing Then
                tempDoc.content.Copy
                Set rng = wdDoc.Range(wdDoc.content.End - 1, wdDoc.content.End - 1)
                rng.Paste
                tempDoc.Close SaveChanges:=False
                Set tempDoc = Nothing
            End If
        End If
        On Error GoTo ErrHandler
    Next i

    '// ── 7. 목차 생성 ──
    If wantTOC Then
        Application.StatusBar = "목차를 생성하는 중..."
        Call BuildTOCPage(wdApp, wdDoc, pdfNames, bmNames, fileCount)
    End If

    '// ── 8. [최종 1회] 페이지계산 ON → Repaginate → 필드갱신 → PDF내보내기 ──
    Application.StatusBar = "PDF로 저장하는 중..."

    wdApp.options.Pagination = True
    wdDoc.Repaginate

    Dim fld As Object
    For Each fld In wdDoc.Fields
        fld.Update
    Next fld

    wdDoc.ExportAsFixedFormat _
        OutputFileName:=savePath, _
        ExportFormat:=WD_EXPORT_PDF, _
        OpenAfterExport:=False, _
        OptimizeFor:=0, _
        CreateBookmarks:=WD_CREATE_BOOKMARKS

    '// ── 9. Word 정리 ──
    wdApp.ScreenUpdating = True
    wdDoc.Close SaveChanges:=False
    Set wdDoc = Nothing

    If Not wordWasOpen Then
        wdApp.Quit SaveChanges:=False
        Set wdApp = Nothing
    End If

    Application.StatusBar = False
    MP_MergeFiles = fso.FileExists(savePath)
    If Not MP_MergeFiles Then mLastErr = "PDF 파일이 만들어지지 않았습니다."
    Exit Function

ErrHandler:
    mLastErr = "Error " & Err.Number & ": " & Err.Description
    Application.StatusBar = False
    On Error Resume Next
    If Not wdApp Is Nothing Then
        wdApp.ScreenUpdating = True
        wdApp.options.Pagination = True
    End If
    If Not wdDoc Is Nothing Then wdDoc.Close SaveChanges:=False
    If Not wdApp Is Nothing And Not wordWasOpen Then wdApp.Quit SaveChanges:=False
    Set wdDoc = Nothing: Set wdApp = Nothing
    On Error GoTo 0
    MP_MergeFiles = False
End Function


'// =============================================================================
'//  PDF 파일 수집 + 파일명 기준 오름차순 정렬
'// =============================================================================
Private Function CollectPDFs(fso As Object, folderPath As String, _
                             ByRef pdfPaths() As String, _
                             ByRef pdfNames() As String) As Long
    Dim folder As Object, file As Object
    Dim cnt As Long, i As Long, j As Long
    Dim tmp As String

    Set folder = fso.GetFolder(folderPath)

    '// 개수 세기
    cnt = 0
    For Each file In folder.files
        If LCase(fso.GetExtensionName(file.name)) = "pdf" And Left$(file.name, 3) <> "병합_" Then cnt = cnt + 1   ' 이전 병합본 제외
    Next
    If cnt = 0 Then CollectPDFs = 0: Exit Function

    ReDim pdfPaths(1 To cnt)
    ReDim pdfNames(1 To cnt)

    '// 배열 담기
    i = 0
    For Each file In folder.files
        If LCase(fso.GetExtensionName(file.name)) = "pdf" And Left$(file.name, 3) <> "병합_" Then
            i = i + 1
            pdfPaths(i) = file.path
            pdfNames(i) = fso.GetBaseName(file.name)   '// 확장자 제거
        End If
    Next

    '// 버블정렬 (파일명 오름차순)
    For i = 1 To cnt - 1
        For j = 1 To cnt - i
            If StrComp(pdfNames(j), pdfNames(j + 1), vbTextCompare) > 0 Then
                tmp = pdfPaths(j): pdfPaths(j) = pdfPaths(j + 1): pdfPaths(j + 1) = tmp
                tmp = pdfNames(j): pdfNames(j) = pdfNames(j + 1): pdfNames(j + 1) = tmp
            End If
        Next
    Next

    CollectPDFs = cnt
End Function


'// =============================================================================
'//  ★ 목차 페이지 생성 (문서 맨 앞)
'//
'//  [글꼴 고정] 전체 맑은 고딕 12pt
'//
'//  [구조]
'//     목  차                          ← 맑은 고딕 12pt Bold
'//    ──────────────────────
'//     1. 파일명A .............. 2     ← 맑은 고딕 12pt
'//     2. 파일명B .............. 5
'//    ──────── (페이지 나누기) ────────
'//
'//  ★ [성능] Repaginate/Fields.Update 없음 → Main에서 최종 1회
'// =============================================================================
Private Sub BuildTOCPage(wdApp As Object, wdDoc As Object, _
                         pdfNames() As String, bmNames() As String, _
                         fileCount As Long)
    Dim i As Long
    Dim pageWidth As Single

    '// 탭 위치 계산
    With wdDoc.PageSetup
        pageWidth = .pageWidth - .LeftMargin - .RightMargin
    End With

    '// ─── 커서를 문서 맨 앞으로 ───
    wdApp.Selection.HomeKey WD_STORY

    '// ???????????????????????????????????????????????????????
    '//  목차 제목: 맑은 고딕 12pt Bold 가운데
    '// ???????????????????????????????????????????????????????
    With wdApp.Selection
        .Font.name = "맑은 고딕"
        .Font.Size = 12
        .Font.Bold = True
        .Font.color = 0
        .Font.Underline = 0
        .ParagraphFormat.Alignment = WD_ALIGN_CENTER
        .ParagraphFormat.SpaceBefore = 40
        .ParagraphFormat.SpaceAfter = 10
        .TypeText "목  차"
        .TypeParagraph
    End With

    '// ???????????????????????????????????????????????????????
    '//  구분선
    '// ???????????????????????????????????????????????????????
    With wdApp.Selection
        .Font.name = "맑은 고딕"
        .Font.Size = 2
        .Font.Bold = False
        .ParagraphFormat.SpaceBefore = 0
        .ParagraphFormat.SpaceAfter = 16
        .ParagraphFormat.Alignment = WD_ALIGN_CENTER
        .TypeText String(80, "_")
        .TypeParagraph
    End With

    '// ???????????????????????????????????????????????????????
    '//  목차 항목 서식: 맑은 고딕 12pt + 점선 리더
    '// ???????????????????????????????????????????????????????
    With wdApp.Selection
        .Font.name = "맑은 고딕"
        .Font.Size = 12
        .Font.Bold = False
        .Font.color = 0
        .Font.Underline = 0
        .ParagraphFormat.Alignment = WD_ALIGN_LEFT
        .ParagraphFormat.SpaceBefore = 3
        .ParagraphFormat.SpaceAfter = 3

        '// 오른쪽 탭 + 점선 리더
        .ParagraphFormat.TabStops.clearAll
        .ParagraphFormat.TabStops.Add _
            Position:=pageWidth, _
            Alignment:=WD_ALIGN_TAB_RIGHT, _
            Leader:=WD_TAB_LEADER_DOTS
    End With

    '// ???????????????????????????????????????????????????????
    '//  목차 항목 입력
    '// ???????????????????????????????????????????????????????
    For i = 1 To fileCount
        Dim entryText As String
        entryText = Format(i, "0") & ". " & pdfNames(i)

        wdApp.Selection.TypeText entryText
        wdApp.Selection.TypeText vbTab

        '// PAGEREF 필드 (자동 페이지 번호)
        Dim oField As Object
        Set oField = wdDoc.Fields.Add( _
            Range:=wdApp.Selection.Range, _
            Type:=WD_FIELD_EMPTY, _
            text:="PAGEREF " & bmNames(i), _
            PreserveFormatting:=False)

        '// 필드 뒤로 커서 이동
        On Error Resume Next
        wdApp.Selection.Start = oField.Result.End
        wdApp.Selection.Collapse WD_COLLAPSE_END
        On Error GoTo 0

        wdApp.Selection.TypeParagraph
    Next i

    '// ???????????????????????????????????????????????????????
    '//  목차 뒤 페이지 나누기
    '// ???????????????????????????????????????????????????????
    wdApp.Selection.InsertBreak WD_PAGE_BREAK
    '// ★ 첫 파일 북마크 재지정(2.7): BM_001 은 문서 맨 앞(0)에 있어 목차가 앞에 들어오면 목차 쪽을 가리킴(첫 항목 쪽번호·링크가 1쪽 오류) → 목차 뒤 첫 내용 위치로
    On Error Resume Next
    wdDoc.Bookmarks.Add bmNames(1), wdDoc.Range(wdApp.Selection.Start, wdApp.Selection.Start)
    On Error GoTo 0

    '// ???????????????????????????????????????????????????????
    '//  하이퍼링크 추가 (클릭 → 해당 PDF 페이지 이동)
    '// ???????????????????????????????????????????????????????
    '//  단락: 1=제목, 2=구분선, 3~=목차항목
    Dim tocStartPara As Long
    tocStartPara = 3

    For i = 1 To fileCount
        Dim parIdx As Long
        parIdx = tocStartPara + (i - 1)

        If parIdx <= wdDoc.Paragraphs.Count Then
            Dim paraRange As Object
            Set paraRange = wdDoc.Paragraphs(parIdx).Range

            Dim paraText As String
            paraText = paraRange.text

            Dim tabPos As Long
            tabPos = InStr(paraText, vbTab)

            If tabPos > 1 Then
                Dim linkRng As Object
                Set linkRng = wdDoc.Range(paraRange.Start, paraRange.Start + tabPos - 1)

                wdDoc.Hyperlinks.Add _
                    anchor:=linkRng, _
                    Address:="", _
                    SubAddress:=bmNames(i), _
                    TextToDisplay:=""

                '// ★ 하이퍼링크 서식 고정: 맑은 고딕 12pt, 진파란, 밑줄 없음
                On Error Resume Next
                linkRng.Font.name = "맑은 고딕"
                linkRng.Font.Size = 12
                linkRng.Font.color = RGB(0, 51, 153)
                linkRng.Font.Underline = 0
                On Error GoTo 0
            End If
        End If
    Next i
End Sub


'// =============================================================================
'//  엑셀 시트에 결과 기록
'// =============================================================================
Private Sub WriteResultSheet(folderPath As String, savePath As String, _
                             pdfNames() As String, pdfPaths() As String, _
                             fileCount As Long, fso As Object)
    Dim wb As Workbook, ws As Worksheet
    Dim i As Long

    Set wb = ActiveWorkbook
    If wb Is Nothing Then Exit Sub

    On Error Resume Next
    Set ws = wb.Sheets("PDF 병합 결과")
    If ws Is Nothing Then
        Set ws = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        ws.name = "PDF 병합 결과"
    End If
    On Error GoTo 0

    ws.Cells.ClearContents
    ws.Activate

    ws.Range("A2").value = "통합된 파일명"
    ws.Range("A3").value = "No."
    ws.Range("B3").value = "파일명 (목차 순서)"

    With ws.Range("A2:B3")
        .Font.Bold = True
        .Interior.color = RGB(220, 230, 241)
    End With

    ws.Hyperlinks.Add anchor:=ws.Cells(2, "B"), _
                       Address:=savePath, _
                       TextToDisplay:=fso.GetFileName(savePath)

    For i = 1 To fileCount
        ws.Cells(i + 3, "A").value = i
        ws.Hyperlinks.Add anchor:=ws.Cells(i + 3, "B"), _
                           Address:=pdfPaths(i), _
                           TextToDisplay:=pdfNames(i) & ".pdf"
    Next i

    ws.Columns("A:B").AutoFit
End Sub

