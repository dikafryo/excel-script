Attribute VB_Name = "modOfficeMerge"
Option Explicit
' 폴더선택 office파일합치기: 폴더 안 한글(hwp/hwpx)·PDF·Word·Excel·PowerPoint·이미지(jpg/png/bmp/gif) 를
' 이름순으로 하나의 PDF 로 합친다. Word 문서 하나에 차례로 넣은 뒤 PDF 로 내보내는 방식.
'  - docx/doc, 이미지 : 원본 그대로 삽입
'  - hwp/hwpx        : 한글로 PDF 변환 후 삽입
'  - xls/xlsx/xlsm   : 엑셀로 PDF 변환 후 삽입
'  - ppt/pptx        : 슬라이드를 PNG 로 내보내 그림으로 삽입
'  - pdf             : Word 의 PDF 변환으로 삽입 (스캔본·복잡한 서식은 모양이 달라질 수 있음)
' 결과: 같은 폴더의 office병합_YYYYMMDD_HHMMSS.pdf + 'office병합 결과' 시트

Private Const WD_PAGE_BREAK As Long = 7
Private Const WD_EXPORT_PDF As Long = 17
Private Const WD_ALERTS_NONE As Long = 0
Private Const OUT_PREFIX As String = "office병합_"

Public Sub OfficeMergeToPdf(control As IRibbonControl)
    Dim fso As Object, folderPath As String, paths() As String, names() As String, kinds() As String
    Dim n As Long, i As Long, wdApp As Object, wdDoc As Object, wordWasOpen As Boolean
    Dim hwp As Object, ppApp As Object, tmpDir As String, savePath As String, t0 As Single
    Dim results() As String, okCount As Long, failCount As Long, rng As Object, endPos As Long

    BX_Trace = "OM:folder"
    With Application.FileDialog(msoFileDialogFolderPicker)
        .title = "PDF 하나로 합칠 office 파일들이 있는 폴더를 선택하세요"
        If .Show <> -1 Then Exit Sub
        folderPath = .SelectedItems(1)
    End With
    BX_Trace = "OM:collect"
    Set fso = CreateObject("Scripting.FileSystemObject")
    n = OM_Collect(fso, folderPath, paths, names, kinds)
    If n = 0 Then
        MsgBox "선택한 폴더에 합칠 파일(hwp/hwpx/pdf/doc/docx/xls/xlsx/ppt/pptx/jpg/png)이 없습니다.", vbExclamation, "office파일합치기"
        Exit Sub
    End If
    If UI_Choose("office파일합치기", "파일 " & n & "개를 이름순으로 PDF 하나로 합칩니다." & vbCrLf & _
                 "(한글·엑셀·PPT는 변환을 거치므로 파일 수에 따라 시간이 걸립니다)", Array("합치기 시작", "취소")) <> 1 Then Exit Sub

    BX_Trace = "OM:tmp"
    t0 = Timer
    tmpDir = fso.BuildPath(Environ$("TEMP"), "officemerge_" & Format$(Now, "hhnnss"))
    fso.CreateFolder tmpDir
    ReDim results(1 To n)

    On Error GoTo Fail
    BX_Trace = "OM:word"
    Application.StatusBar = "MS Word 준비 중..."
    On Error Resume Next
    Set wdApp = GetObject(, "Word.Application")
    On Error GoTo Fail
    wordWasOpen = Not (wdApp Is Nothing)
    If wdApp Is Nothing Then Set wdApp = CreateObject("Word.Application")
    If Not wordWasOpen Then wdApp.Visible = False
    wdApp.DisplayAlerts = WD_ALERTS_NONE
    Set wdDoc = wdApp.Documents.Add
    With wdDoc.PageSetup
        .TopMargin = wdApp.CentimetersToPoints(1.5): .BottomMargin = wdApp.CentimetersToPoints(1.5)
        .LeftMargin = wdApp.CentimetersToPoints(1.5): .RightMargin = wdApp.CentimetersToPoints(1.5)
    End With

    For i = 1 To n
        BX_Trace = "OM:file " & i & "/" & n & " " & names(i)
        Application.StatusBar = "office 파일 합치는 중 " & i & "/" & n & " : " & names(i)
        DoEvents
        endPos = wdDoc.content.End - 1
        If i > 1 Then
            Set rng = wdDoc.Range(endPos, endPos)
            rng.InsertBreak WD_PAGE_BREAK
            endPos = wdDoc.content.End - 1
        End If
        Set rng = wdDoc.Range(endPos, endPos)
        results(i) = OM_InsertOne(wdApp, wdDoc, rng, paths(i), kinds(i), fso, tmpDir, hwp, ppApp)
        If Left$(results(i), 2) = "성공" Then okCount = okCount + 1 Else failCount = failCount + 1
    Next i

    BX_Trace = "OM:export"
    Application.StatusBar = "PDF 로 저장하는 중..."
    savePath = fso.BuildPath(folderPath, OUT_PREFIX & Format$(Now, "yyyymmdd_hhnnss") & ".pdf")
    wdDoc.ExportAsFixedFormat OutputFileName:=savePath, ExportFormat:=WD_EXPORT_PDF, OpenAfterExport:=False
    wdDoc.Close SaveChanges:=False
    Set wdDoc = Nothing
    If Not wordWasOpen Then wdApp.Quit SaveChanges:=False
    Set wdApp = Nothing
    OM_Cleanup hwp, ppApp, fso, tmpDir
    Application.StatusBar = False

    BX_Trace = "OM:result"
    OM_WriteResult folderPath, savePath, names, paths, kinds, results, n, fso
    If UI_Choose("office파일합치기 완료", "성공 " & okCount & "개 / 실패 " & failCount & "개 (" & Format$(Timer - t0, "0.0") & "초)" & vbCrLf & vbCrLf & _
                 "저장 위치: " & savePath & vbCrLf & "(파일별 결과는 'office병합 결과' 시트에 있습니다)", Array("합친 PDF 열기", "닫기")) = 1 Then
        Shell "explorer """ & savePath & """", vbNormalFocus
    End If
    Exit Sub
Fail:
    Dim em As String: em = Err.Description & " @" & BX_Trace
    On Error Resume Next
    Application.StatusBar = False
    If Not wdDoc Is Nothing Then wdDoc.Close SaveChanges:=False
    If Not wdApp Is Nothing And Not wordWasOpen Then wdApp.Quit SaveChanges:=False
    OM_Cleanup hwp, ppApp, fso, tmpDir
    MsgBox "office 파일 합치기 중 오류가 났습니다." & vbCrLf & vbCrLf & em & vbCrLf & vbCrLf & _
           "MS Word 가 설치되어 있어야 하며, 한글·PowerPoint 파일은 해당 프로그램이 필요합니다.", vbCritical, "office파일합치기"
End Sub

' 대상 파일 수집 (이름순 정렬, 이전 병합본·임시파일 제외). kinds: doc/img/pdf/hwp/xls/ppt
Private Function OM_Collect(fso As Object, ByVal folderPath As String, ByRef paths() As String, ByRef names() As String, ByRef kinds() As String) As Long
    Dim f As Object, ext As String, k As String, cnt As Long, i As Long, j As Long, T As String
    For Each f In fso.GetFolder(folderPath).files
        k = OM_Kind(LCase$(fso.GetExtensionName(f.name)))
        If Len(k) > 0 And Left$(f.name, 2) <> "~$" And Left$(f.name, Len(OUT_PREFIX)) <> OUT_PREFIX And Left$(f.name, 3) <> "병합_" Then
            cnt = cnt + 1
            ReDim Preserve paths(1 To cnt): ReDim Preserve names(1 To cnt): ReDim Preserve kinds(1 To cnt)
            paths(cnt) = f.path: names(cnt) = f.name: kinds(cnt) = k
        End If
    Next f
    For i = 1 To cnt - 1
        For j = 1 To cnt - i
            If StrComp(names(j), names(j + 1), vbTextCompare) > 0 Then
                T = paths(j): paths(j) = paths(j + 1): paths(j + 1) = T
                T = names(j): names(j) = names(j + 1): names(j + 1) = T
                T = kinds(j): kinds(j) = kinds(j + 1): kinds(j + 1) = T
            End If
        Next j
    Next i
    OM_Collect = cnt
End Function

Private Function OM_Kind(ByVal ext As String) As String
    Select Case ext
        Case "doc", "docx", "rtf": OM_Kind = "doc"
        Case "jpg", "jpeg", "png", "bmp", "gif", "tif", "tiff": OM_Kind = "img"
        Case "pdf": OM_Kind = "pdf"
        Case "hwp", "hwpx": OM_Kind = "hwp"
        Case "xls", "xlsx", "xlsm": OM_Kind = "xls"
        Case "ppt", "pptx": OM_Kind = "ppt"
    End Select
End Function

' 파일 1개를 Word 문서 끝(rng)에 삽입. 반환: "성공" 또는 "실패: 이유"
Private Function OM_InsertOne(wdApp As Object, wdDoc As Object, rng As Object, ByVal p As String, ByVal kind As String, _
                              fso As Object, ByVal tmpDir As String, ByRef hwp As Object, ByRef ppApp As Object) As String
    Dim tmpPdf As String, k As Long, sld As Object, pres As Object, png As String, wb As Object, tmpCount As Long
    On Error GoTo EH
    Select Case kind
        Case "doc"
            rng.InsertFile fileName:=p, ConfirmConversions:=False
            OM_InsertOne = "성공"
        Case "img"
            OM_AddPicture wdApp, wdDoc, rng, p
            OM_InsertOne = "성공"
        Case "pdf"
            OM_InsertOne = OM_InsertPdf(wdApp, wdDoc, rng, p)
        Case "hwp"
            If hwp Is Nothing Then
                On Error Resume Next
                Set hwp = CreateObject("HWPFrame.HwpObject")
                If Not hwp Is Nothing Then
                    hwp.RegisterModule "FilePathCheckDLL", "FilePathCheckerModule"
                    hwp.XHwpWindows.item(0).Visible = False
                End If
                On Error GoTo EH
            End If
            If hwp Is Nothing Then OM_InsertOne = "실패: 한글(HWP) 프로그램 없음": Exit Function
            tmpPdf = fso.BuildPath(tmpDir, fso.GetBaseName(p) & "_" & Format$(Timer * 100, "0") & ".pdf")
            If (hwp.Open(p, "", "forceopen:true") <> 0) Then
                If (hwp.SaveAs(tmpPdf, "PDF", "") <> 0) And fso.FileExists(tmpPdf) Then
                    OM_InsertOne = OM_InsertPdf(wdApp, wdDoc, rng, tmpPdf)
                Else
                    OM_InsertOne = "실패: 한글 PDF 변환 실패"
                End If
                hwp.Clear 1
            Else
                OM_InsertOne = "실패: 한글 파일 열기 실패"
            End If
        Case "xls"
            tmpPdf = fso.BuildPath(tmpDir, fso.GetBaseName(p) & "_" & Format$(Timer * 100, "0") & ".pdf")
            Set wb = Application.Workbooks.Open(p, ReadOnly:=True, UpdateLinks:=0)
            wb.ExportAsFixedFormat Type:=xlTypePDF, fileName:=tmpPdf, OpenAfterPublish:=False
            wb.Close SaveChanges:=False
            OM_InsertOne = OM_InsertPdf(wdApp, wdDoc, rng, tmpPdf)
        Case "ppt"
            If ppApp Is Nothing Then Set ppApp = CreateObject("PowerPoint.Application")
            Set pres = ppApp.Presentations.Open(p, True, False, False)   ' ReadOnly, Untitled, WithWindow
            For k = 1 To pres.Slides.Count
                png = fso.BuildPath(tmpDir, fso.GetBaseName(p) & "_s" & Format$(k, "000") & ".png")
                pres.Slides(k).Export png, "PNG", 1600
                If k > 1 Then
                    Set rng = wdDoc.Range(wdDoc.content.End - 1, wdDoc.content.End - 1)
                    rng.InsertBreak WD_PAGE_BREAK
                End If
                Set rng = wdDoc.Range(wdDoc.content.End - 1, wdDoc.content.End - 1)
                OM_AddPicture wdApp, wdDoc, rng, png
            Next k
            pres.Close
            OM_InsertOne = "성공 (" & k - 1 & "장)"
        Case Else
            OM_InsertOne = "실패: 지원하지 않는 형식"
    End Select
    Exit Function
EH:
    OM_InsertOne = "실패: " & Err.Description
    On Error Resume Next
    If Not pres Is Nothing Then pres.Close
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
End Function

' PDF 삽입: InsertFile 이 안 되면 문서로 열어 복사/붙여넣기
Private Function OM_InsertPdf(wdApp As Object, wdDoc As Object, rng As Object, ByVal p As String) As String
    Dim tmpDoc As Object
    On Error Resume Next
    Err.Clear
    rng.InsertFile fileName:=p, ConfirmConversions:=False
    If Err.Number = 0 Then OM_InsertPdf = "성공": Exit Function
    Err.Clear
    Set tmpDoc = wdApp.Documents.Open(fileName:=p, ConfirmConversions:=False, ReadOnly:=True, Visible:=False)
    If tmpDoc Is Nothing Or Err.Number <> 0 Then OM_InsertPdf = "실패: PDF 변환 불가 (" & Err.Description & ")": Exit Function
    tmpDoc.content.Copy
    Set rng = wdDoc.Range(wdDoc.content.End - 1, wdDoc.content.End - 1)
    rng.Paste
    tmpDoc.Close SaveChanges:=False
    If Err.Number = 0 Then OM_InsertPdf = "성공" Else OM_InsertPdf = "실패: " & Err.Description
End Function

' 그림을 페이지 폭에 맞춰 삽입
Private Sub OM_AddPicture(wdApp As Object, wdDoc As Object, rng As Object, ByVal p As String)
    Dim shp As Object, maxW As Single, maxH As Single, r As Single
    With wdDoc.PageSetup
        maxW = .pageWidth - .LeftMargin - .RightMargin
        maxH = .PageHeight - .TopMargin - .BottomMargin - 6
    End With
    Set shp = rng.InlineShapes.AddPicture(fileName:=p, LinkToFile:=False, SaveWithDocument:=True, Range:=rng)
    shp.LockAspectRatio = True
    r = 1
    If shp.Width > maxW Then r = maxW / shp.Width
    If shp.Height * r > maxH Then r = maxH / shp.Height
    If r < 1 Then shp.Width = shp.Width * r
End Sub

Private Sub OM_Cleanup(ByRef hwp As Object, ByRef ppApp As Object, fso As Object, ByVal tmpDir As String)
    On Error Resume Next
    If Not hwp Is Nothing Then hwp.Quit: Set hwp = Nothing
    If Not ppApp Is Nothing Then
        If ppApp.Presentations.Count = 0 Then ppApp.Quit
        Set ppApp = Nothing
    End If
    If Len(tmpDir) > 0 Then If fso.FolderExists(tmpDir) Then fso.DeleteFolder tmpDir, True
End Sub

Private Sub OM_WriteResult(ByVal folderPath As String, ByVal savePath As String, names() As String, paths() As String, kinds() As String, results() As String, ByVal n As Long, fso As Object)
    Dim ws As Worksheet, i As Long, r As Long
    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets("office병합 결과")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ActiveWorkbook.Worksheets.Add(After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count))
        ws.name = "office병합 결과"
    End If
    ws.Cells.Clear
    ws.Range("A2").value = "폴더": ws.Hyperlinks.Add anchor:=ws.Range("B2"), Address:=folderPath, TextToDisplay:=folderPath
    ws.Range("A3").value = "합친 PDF": ws.Hyperlinks.Add anchor:=ws.Range("B3"), Address:=savePath, TextToDisplay:=fso.GetFileName(savePath)
    ws.Range("A5").value = "No.": ws.Range("B5").value = "파일명 (합친 순서)": ws.Range("C5").value = "유형": ws.Range("D5").value = "결과"
    With ws.Range("A5:D5")
        .Font.Bold = True: .Interior.color = RGB(220, 230, 241): .HorizontalAlignment = xlCenter
    End With
    ws.Range("A2:A3").Font.Bold = True
    r = 6
    For i = 1 To n
        ws.Cells(r, 1).value = i
        ws.Hyperlinks.Add anchor:=ws.Cells(r, 2), Address:=paths(i), TextToDisplay:=names(i)
        ws.Cells(r, 3).value = kinds(i)
        ws.Cells(r, 4).value = results(i)
        If Left$(results(i), 2) <> "성공" Then ws.Cells(r, 4).Font.color = RGB(255, 0, 0)
        r = r + 1
    Next i
    ws.Range(ws.Cells(5, 1), ws.Cells(r - 1, 4)).Borders.LineStyle = xlContinuous
    ws.Columns("A:D").AutoFit
    ws.Activate
    ws.Range("A1").Select
End Sub

