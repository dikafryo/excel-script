Attribute VB_Name = "modHwpToPdf"
Option Explicit
' 폴더 안 HWP/HWPX 를 PDF 로 일괄 변환하고 '변환 결과' 시트에 기록한다.
' [개선] 한글 보안모듈 등록+창 숨김(파일마다 승인창·한글창 뜨던 문제), 파일별 오류 처리(한 파일 오류로 중단·한글 잔류 방지),
'        임시파일(~$) 제외, 진행률, 성공/실패 집계. 리본에서 쓰지 않는 Acrobat 병합 Sub 삭제(PDF 병합은 modMergePdf).
' [v2.6] HP_Run(mode): 1 각각의 PDF / 2 하나의 PDF(목차·링크) / 3 하나의 PDF(파일만) - 2·3은 임시 폴더에 변환 후 modMergePdf.MP_MergeFiles 로 합침

Public HP_TestFolder As String   ' 검증용: 지정되면 폴더 대화상자 생략

Sub dHWP_to_PDF_Converter(control As IRibbonControl)
    HP_Run 1
End Sub

Public Sub HP_Run(ByVal mode As Long)
    Dim fso As Object, ws As Worksheet, folderPath As String, outFolder As String, total As Long
    Dim pdfPaths() As String, pdfNames() As String, okCount As Long, failCount As Long, savePath As String, t0 As Single, i As Long

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Len(HP_TestFolder) > 0 Then
        folderPath = HP_TestFolder
    Else
        With Application.FileDialog(msoFileDialogFolderPicker)
            .title = "HWP/HWPX 파일이 있는 폴더를 선택하세요"
            If .Show <> -1 Then Exit Sub
            folderPath = .SelectedItems(1)
        End With
    End If
    t0 = Timer
    If mode = 1 Then
        outFolder = folderPath
    Else
        outFolder = fso.BuildPath(folderPath, "_pdf임시_" & Format$(Now, "hhnnss"))
    End If

    Set ws = HP_ResultSheet(ActiveWorkbook)
    total = HP_ConvertFolder(fso, folderPath, outFolder, ws, pdfPaths, pdfNames, okCount, failCount)
    If total = 0 Then
        MsgBox "선택한 폴더에 HWP/HWPX 파일이 없습니다.", vbExclamation, "hwp to pdf"
        Exit Sub
    End If
    If total < 0 Then Exit Sub   ' 한글 없음 등 (메시지는 안에서)

    If mode = 1 Then
        ws.Activate
        MsgBox "PDF 변환이 끝났습니다.  (" & Format$(Timer - t0, "0.0") & "초)" & vbCrLf & vbCrLf & _
               "성공: " & okCount & "개 / 실패: " & failCount & "개" & vbCrLf & _
               "PDF 는 원본과 같은 폴더에 저장되었고, '변환 결과' 시트에서 파일을 열 수 있습니다.", _
               IIf(failCount > 0, vbExclamation, vbInformation), "hwp to pdf"
        Exit Sub
    End If

    ' 2·3: 변환된 PDF 를 하나로 합치기
    If okCount = 0 Then
        MsgBox "PDF 로 변환된 파일이 없어 합치지 못했습니다. '변환 결과' 시트를 확인하세요.", vbExclamation, "hwp to pdf"
        On Error Resume Next: fso.DeleteFolder outFolder, True: On Error GoTo 0
        ws.Activate
        Exit Sub
    End If
    If mode = 2 Then
        savePath = fso.BuildPath(folderPath, "병합_목차_" & Format$(Now, "yyyymmdd_hhnnss") & ".pdf")
    Else
        savePath = fso.BuildPath(folderPath, "병합_" & Format$(Now, "yyyymmdd_hhnnss") & ".pdf")
    End If
    If MP_MergeFiles(fso, pdfPaths, pdfNames, okCount, (mode = 2), savePath) Then
        ws.Range("A1").value = "합친 PDF"
        ws.Hyperlinks.Add anchor:=ws.Range("B1"), Address:=savePath, TextToDisplay:=fso.GetFileName(savePath)
        ws.Range("A1").Font.Bold = True
        ws.Columns("A:D").AutoFit
        ws.Activate
        On Error Resume Next
        fso.DeleteFolder outFolder, True
        On Error GoTo 0
        If UI_Choose("hwp to pdf 완료", "한글 파일 " & okCount & "개를 PDF 로 변환해 하나로 합쳤습니다. (" & Format$(Timer - t0, "0.0") & "초" & IIf(failCount > 0, ", 변환 실패 " & failCount & "개", "") & ")" & vbCrLf & vbCrLf & _
                     "저장 위치: " & savePath & vbCrLf & "(파일별 결과는 '변환 결과' 시트에 있습니다)", Array("합친 PDF 열기", "닫기")) = 1 Then
            Shell "explorer """ & savePath & """", vbNormalFocus
        End If
    Else
        ws.Activate
        MsgBox "PDF 합치기에 실패했습니다. 변환된 PDF 는 " & outFolder & " 에 남겨 두었습니다.", vbExclamation, "hwp to pdf"
    End If
End Sub

' 결과 시트: 있으면 재사용, 없으면 생성
Private Function HP_ResultSheet(ByVal wb As Workbook) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Sheets("변환 결과")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        ws.name = "변환 결과"
    End If
    Set HP_ResultSheet = ws
End Function

' 폴더 안 hwp/hwpx → outFolder 에 PDF 변환. 반환: 대상 파일 수(0=없음, -1=한글 실행 불가). pdfPaths/pdfNames 는 성공한 것만(1~okCount)
Public Function HP_ConvertFolder(ByVal fso As Object, ByVal folderPath As String, ByVal outFolder As String, ByVal ws As Worksheet, _
                                 ByRef pdfPaths() As String, ByRef pdfNames() As String, ByRef okCount As Long, ByRef failCount As Long) As Long
    Dim folder As Object, file As Object, objHwp As Object
    Dim filePath As String, pdfPath As String, ext As String
    Dim row As Long, total As Long, Done As Long, opened As Boolean
    Dim names() As String, paths() As String, i As Long, j As Long, T As String

    Set folder = fso.GetFolder(folderPath)
    For Each file In folder.files
        ext = LCase$(fso.GetExtensionName(file.name))
        If (ext = "hwp" Or ext = "hwpx") And Left$(file.name, 2) <> "~$" And Left$(file.name, 3) <> "병합_" Then
            total = total + 1
            ReDim Preserve names(1 To total): ReDim Preserve paths(1 To total)
            names(total) = file.name: paths(total) = file.path
        End If
    Next file
    If total = 0 Then Exit Function
    ' 이름순 정렬
    For i = 1 To total - 1
        For j = 1 To total - i
            If StrComp(names(j), names(j + 1), vbTextCompare) > 0 Then
                T = names(j): names(j) = names(j + 1): names(j + 1) = T
                T = paths(j): paths(j) = paths(j + 1): paths(j + 1) = T
            End If
        Next j
    Next i

    On Error Resume Next
    Set objHwp = CreateObject("HWPFrame.HwpObject")
    On Error GoTo 0
    If objHwp Is Nothing Then
        MsgBox "한글 프로그램을 실행할 수 없습니다. 한글(HWP)이 설치되어 있는지 확인하세요.", vbCritical, "hwp to pdf"
        HP_ConvertFolder = -1
        Exit Function
    End If
    On Error Resume Next
    objHwp.RegisterModule "FilePathCheckDLL", "FilePathCheckerModule"
    objHwp.XHwpWindows.item(0).Visible = False
    If Not fso.FolderExists(outFolder) Then fso.CreateFolder outFolder
    On Error GoTo 0

    ws.Cells.Clear
    ws.Range("A2").value = "폴더경로"
    ws.Hyperlinks.Add anchor:=ws.Range("B2"), Address:=folderPath, TextToDisplay:=folderPath
    ws.Range("A3").value = "No.": ws.Range("B3").value = "원본 파일명 (HWP)"
    ws.Range("C3").value = "변환 결과": ws.Range("D3").value = "변환된 파일 (PDF)"
    With ws.Range("A2:D3")
        .Font.Bold = True
        .Interior.color = RGB(220, 230, 241)
        .HorizontalAlignment = xlCenter
    End With
    row = 4
    ReDim pdfPaths(1 To total): ReDim pdfNames(1 To total)

    Application.ScreenUpdating = False
    For i = 1 To total
        Done = Done + 1
        Application.StatusBar = "PDF 변환 중 " & Done & "/" & total & " : " & names(i)
        DoEvents
        filePath = paths(i)
        pdfPath = fso.BuildPath(outFolder, fso.GetBaseName(names(i)) & ".pdf")
        If fso.FileExists(pdfPath) And outFolder <> folderPath Then pdfPath = fso.BuildPath(outFolder, fso.GetBaseName(names(i)) & "_" & Format$(i, "000") & ".pdf")
        ws.Cells(row, "A").value = row - 3
        ws.Hyperlinks.Add anchor:=ws.Cells(row, "B"), Address:=filePath, TextToDisplay:=names(i)

        On Error Resume Next
        Err.Clear
        opened = (objHwp.Open(filePath, "", "forceopen:true") <> 0)
        If Err.Number <> 0 Then opened = False
        If opened Then
            Err.Clear
            If (objHwp.SaveAs(pdfPath, "PDF", "") <> 0) And Err.Number = 0 And fso.FileExists(pdfPath) Then
                ws.Cells(row, "C").value = "성공"
                ws.Cells(row, "C").Font.color = RGB(0, 176, 80)
                ws.Hyperlinks.Add anchor:=ws.Cells(row, "D"), Address:=pdfPath, TextToDisplay:=fso.GetFileName(pdfPath)
                okCount = okCount + 1
                pdfPaths(okCount) = pdfPath
                pdfNames(okCount) = fso.GetBaseName(names(i))
            Else
                ws.Cells(row, "C").value = "변환 실패" & IIf(Err.Number <> 0, ": " & Err.Description, "")
                ws.Cells(row, "C").Font.color = RGB(255, 0, 0)
                failCount = failCount + 1
            End If
            objHwp.Clear 1
        Else
            ws.Cells(row, "C").value = "파일 열기 실패" & IIf(Err.Number <> 0, ": " & Err.Description, "")
            ws.Cells(row, "C").Font.color = RGB(255, 0, 0)
            failCount = failCount + 1
        End If
        On Error GoTo 0
        row = row + 1
    Next i

    On Error Resume Next
    objHwp.Quit
    Set objHwp = Nothing
    On Error GoTo 0
    ws.Columns("A:D").AutoFit
    Application.StatusBar = False
    Application.ScreenUpdating = True
    HP_ConvertFolder = total
End Function

