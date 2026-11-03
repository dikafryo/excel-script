Attribute VB_Name = "modMemoImage"
Option Explicit
'=======================================
' Windows API 선언
'=======================================
#If VBA7 Then
    Private Declare PtrSafe Sub keybd_event Lib "user32" ( _
        ByVal bVk As Byte, ByVal bScan As Byte, _
        ByVal dwFlags As Long, ByVal dwExtraInfo As LongPtr)
    Private Declare PtrSafe Sub Sleep Lib "kernel32" ( _
        ByVal dwMilliseconds As Long)
    Private Declare PtrSafe Function IsClipboardFormatAvailable Lib "user32" ( _
        ByVal wFormat As Long) As Long
#Else
    Private Declare Sub keybd_event Lib "user32" ( _
        ByVal bVk As Byte, ByVal bScan As Byte, _
        ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
    Private Declare Sub Sleep Lib "kernel32" ( _
        ByVal dwMilliseconds As Long)
    Private Declare Function IsClipboardFormatAvailable Lib "user32" ( _
        ByVal wFormat As Long) As Long
#End If

'=======================================
' 상수 정의
'=======================================
Private Const VK_LWIN As Byte = &H5B
Private Const VK_SHIFT As Byte = &H10
Private Const VK_S As Byte = &H53
Private Const KEYEVENTF_KEYDOWN As Long = 0
Private Const KEYEVENTF_KEYUP As Long = &H2
Private Const CF_BITMAP As Long = 2

'=======================================
' PowerShell로 클립보드 → PNG 직접 저장 (차트 Export 우회)
'=======================================
Private Function SaveClipboardToPNG(ByVal savePath As String) As Boolean
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    
    ' PowerShell 명령: 클립보드 비트맵을 PNG로 직접 저장
    Dim psCmd As String
    psCmd = "powershell -ExecutionPolicy Bypass -Command """ & _
            "Add-Type -AssemblyName System.Windows.Forms;" & _
            "$img = [System.Windows.Forms.Clipboard]::GetImage();" & _
            "if($img -ne $null){" & _
            "$img.Save('" & savePath & "', [System.Drawing.Imaging.ImageFormat]::Png);" & _
            "$img.Dispose()}" & _
            """"
    
    ' 동기 실행 (완료될 때까지 대기), 창 숨김
    wsh.run psCmd, 0, True
    
    Set wsh = Nothing
    
    ' 저장 확인
    SaveClipboardToPNG = (dir(savePath) <> "")
End Function

'=======================================
' PowerShell로 클립보드 이미지 크기 조회
'=======================================
Private Function GetClipboardImageSize(ByRef outWidth As Long, ByRef outHeight As Long) As Boolean
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    
    ' 임시 파일에 크기 정보 기록
    Dim tmpFile As String
    tmpFile = Environ("TEMP") & "\clipsize_" & Format(Now, "HHmmss") & ".txt"
    
    Dim psCmd As String
    psCmd = "powershell -ExecutionPolicy Bypass -Command """ & _
            "Add-Type -AssemblyName System.Windows.Forms;" & _
            "$img = [System.Windows.Forms.Clipboard]::GetImage();" & _
            "if($img -ne $null){" & _
            "[System.IO.File]::WriteAllText('" & tmpFile & "', " & _
            "($img.Width.ToString() + ',' + $img.Height.ToString()));" & _
            "$img.Dispose()}" & _
            """"
    
    wsh.run psCmd, 0, True
    Set wsh = Nothing
    
    GetClipboardImageSize = False
    
    If dir(tmpFile) <> "" Then
        Dim ff As Integer
        ff = FreeFile
        Dim sData As String
        Open tmpFile For Input As #ff
        Line Input #ff, sData
        Close #ff
        kill tmpFile
        
        Dim parts() As String
        parts = Split(sData, ",")
        If UBound(parts) >= 1 Then
            outWidth = CLng(parts(0))
            outHeight = CLng(parts(1))
            GetClipboardImageSize = True
        End If
    End If
End Function

'=======================================
' 메인 매크로 (리본 콜백)
'=======================================
Sub 메모에이미지넣기(control As IRibbonControl)
    ' 리본 클릭(콜백) 안에서는 셀 선택 창(Application.InputBox Type:=8)이 바로 취소되므로, 콜백이 끝난 뒤 실행되도록 예약한다
    If control Is Nothing Then
        MemoImage_Run
    Else
        Application.OnTime Now, "'" & ThisWorkbook.name & "'!MemoImage_Run"
    End If
End Sub

Public Sub MemoImage_Run()
    
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    '-------------------------------------------------------
    ' 1단계: Win + Shift + S 로 스크린 캡처 실행
    '-------------------------------------------------------
    ' [수정] 이미 복사해 둔 이미지가 있으면 그것을 쓸지 물어보고, 없을 때만 화면 캡처(Win+Shift+S)를 띄운다
    Dim choice As Long, needCapture As Boolean
    needCapture = True
    If IsClipboardFormatAvailable(CF_BITMAP) <> 0 Then
        choice = UI_Choose("메모배경넣기", "클립보드에 복사된 이미지가 있습니다. 어떤 이미지를 메모 배경으로 넣을까요?", _
                           Array("클립보드의 이미지 사용", "새로 화면 캡처 (Win+Shift+S)"))
        If choice = 0 Then Exit Sub
        needCapture = (choice = 2)
    End If
    If needCapture Then
        keybd_event VK_LWIN, 0, KEYEVENTF_KEYDOWN, 0
        keybd_event VK_SHIFT, 0, KEYEVENTF_KEYDOWN, 0
        keybd_event VK_S, 0, KEYEVENTF_KEYDOWN, 0
        Sleep 50
        keybd_event VK_S, 0, KEYEVENTF_KEYUP, 0
        keybd_event VK_SHIFT, 0, KEYEVENTF_KEYUP, 0
        keybd_event VK_LWIN, 0, KEYEVENTF_KEYUP, 0
        MsgBox "화면 영역을 드래그하여 캡처한 후" & vbNewLine & _
               "이 창에서 [확인] 버튼을 누르세요.", vbInformation, "스크린 캡처 대기"
        Sleep 500
    End If
    If IsClipboardFormatAvailable(CF_BITMAP) = 0 Then
        MsgBox "클립보드에 이미지가 없습니다." & vbNewLine & _
               "캡처가 정상적으로 되었는지 확인해주세요.", _
               vbExclamation, "오류"
        Exit Sub
    End If
    
    '-------------------------------------------------------
    ' 2단계: 클립보드 이미지 크기 조회
    '-------------------------------------------------------
    Dim pxWidth As Long, pxHeight As Long
    If Not GetClipboardImageSize(pxWidth, pxHeight) Then
        MsgBox "클립보드 이미지 크기를 읽을 수 없습니다.", vbExclamation, "오류"
        Exit Sub
    End If
    
    '-------------------------------------------------------
    ' 3단계: PowerShell로 클립보드 → PNG 직접 저장
    '        ★ 핵심: Chart.Export 대신 .NET으로 원본 그대로 저장
    '-------------------------------------------------------
    ' 임시폴더에 저장 (메모 배경용)
    Dim tempPath As String
    tempPath = Environ("TEMP") & "\ExcelCapture_" & Format(Now, "yyyyMMdd_HHmmss") & ".png"
    
    If Not SaveClipboardToPNG(tempPath) Then
        MsgBox "이미지 저장에 실패했습니다.", vbExclamation, "오류"
        Exit Sub
    End If
    
    ' 바탕화면에도 저장 (선택사항 - 불필요하면 아래 3줄 삭제)
    ' (바탕화면에 사본을 저장하던 동작 제거 - 사용할 때마다 바탕화면에 PNG 가 쌓이던 문제)
    
    '-------------------------------------------------------
    ' 4단계: 셀 선택 다이얼로그 표시
    '-------------------------------------------------------
    Dim rng As Range
    On Error Resume Next
    AppActivate Application.caption          ' 외부 프로세스(PowerShell) 실행 뒤 포커스를 엑셀로
    ActiveWindow.Activate
    Application.GoTo ActiveCell
    DoEvents
    Set rng = Application.InputBox( _
        prompt:="메모(그림 배경)를 추가할 셀을 선택하세요.", _
        title:="셀 선택", _
        Default:=ActiveCell.Address, _
        Type:=8)
    On Error GoTo 0
    
    If rng Is Nothing Then
        MsgBox "셀 선택이 취소되었습니다.", vbExclamation, "취소"
        On Error Resume Next
        kill tempPath
        On Error GoTo 0
        Exit Sub
    End If
    
    Set rng = rng.Cells(1, 1)
    
    '-------------------------------------------------------
    ' 5단계: 메모 추가 → 그림 배경 설정
    '-------------------------------------------------------
    If Not rng.Comment Is Nothing Then
        rng.Comment.Delete
    End If
    
    rng.AddComment text:=" "
    
    ' 픽셀 → 포인트 변환 (96dpi 기준)
    Dim ptWidth As Single, ptHeight As Single
    ptWidth = pxWidth * 72# / 96#
    ptHeight = pxHeight * 72# / 96#
    
    With rng.Comment.Shape
        ' 메모 크기를 이미지 비율에 맞춤
        .Width = ptWidth
        .Height = ptHeight
        
        ' 최소/최대 크기 제한
        If .Width < 150 Then
            Dim scaleUp As Single
            scaleUp = 150 / .Width
            .Width = 150
            .Height = .Height * scaleUp
        End If
        If .Height < 100 Then
            Dim scaleUp2 As Single
            scaleUp2 = 100 / .Height
            .Height = 100
            .Width = .Width * scaleUp2
        End If
        If .Width > 600 Then
            Dim scaleDown As Single
            scaleDown = 600 / .Width
            .Width = 600
            .Height = .Height * scaleDown
        End If
        
        ' 그림 배경 설정
        .Fill.Visible = msoTrue
        .Fill.UserPicture tempPath
        
        ' 메모 텍스트 숨김 처리
        .TextFrame.Characters.Font.color = RGB(255, 255, 255)
        .TextFrame.Characters.Font.Size = 1
    End With
    
    rng.Comment.Visible = False
    
    '-------------------------------------------------------
    ' 6단계: 임시파일 삭제 (메모에 이미 임베드됨)
    '-------------------------------------------------------
    On Error Resume Next
    kill tempPath
    On Error GoTo 0
    
    '-------------------------------------------------------
    ' 7단계: 완료 메시지
    '-------------------------------------------------------
    rng.Select
    
    Dim msg As String
    msg = "완료!" & vbNewLine & vbNewLine & _
          "  메모 셀: " & rng.Address & vbNewLine & _
          "  이미지 크기: " & pxWidth & " x " & pxHeight & " px" & vbNewLine
    
    
    msg = msg & vbNewLine & "※ 셀에 마우스를 올리면 캡처 이미지가 표시됩니다."
    
    MsgBox msg, vbInformation, "완료"

End Sub


