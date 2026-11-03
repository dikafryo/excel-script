Attribute VB_Name = "modMapping"
Option Explicit
'=====================================================================
' modMapping : 부서담당자 시트 / 단위사업카드 시트 (사용자 입력 → 적용)
'=====================================================================
'부서담당자 시트 열
Private Const S_NAME As Long = 1   '성명
Private Const S_TEAM As Long = 2   '팀/부서
Private Const S_POS As Long = 3    '직위
Private Const S_NOTE As Long = 4   '비고
Private Const S_CNT As Long = 5    '담당 사업수(자동)
Private Const S_HDR As Long = 4    '헤더 행
'사업별 담당자 변경 블록 (부서담당자 시트 오른쪽)
Private Const O_TEAM As Long = 7   '팀명
Private Const O_OWN As Long = 8    '담당자(현재)
Private Const O_NM As Long = 9     '사업명
Private Const O_NEW As Long = 10   '▶ 담당자 변경
Private Const O_KEY As Long = 11   '키(숨김)
Public Const SH_OWNCHG As String = "_담당자변경"
Private gOwnOld As Object

'단위사업카드 시트 열
Private Const C_NO As Long = 1
Private Const C_SEB As Long = 2    '세부사업
Private Const C_NM As Long = 3     '사업항목(요구서)
Private Const C_BNM As Long = 4    '장표 사업항목명(자동)
Private Const C_CARD As Long = 5   '단위과제카드(자동)
Private Const C_OWN As Long = 6    '담당자(자동)
Private Const C_MATCH As Long = 7  '매칭
Private Const C_HY As Long = 8     '예산현액
Private Const C_OWN2 As Long = 9   '▶담당자(수정)
Private Const C_CARD2 As Long = 10 '▶단위과제카드(수정)
Private Const C_BSET As Long = 11  '▶장표 사업항목 직접지정
Private Const C_KEY As Long = 12   '키(숨김)
Private Const C_HDR As Long = 4

'---------------------------------------------------------------- 읽기
Public Function ReadStaffTeams() As Object
    Dim d As Object, ws As Worksheet, r As Long, lr As Long, nm As String
    Set d = CreateObject("Scripting.Dictionary")
    If Not SheetExists(SH_STAFF) Then Set ReadStaffTeams = d: Exit Function
    Set ws = BW().Worksheets(SH_STAFF)
    lr = lastRow(ws, S_NAME)
    For r = S_HDR + 1 To lr
        nm = CleanText(ws.Cells(r, S_NAME).Value2)
        If nm <> "" Then d(nm) = CleanText(ws.Cells(r, S_TEAM).Value2)
    Next r
    Set ReadStaffTeams = d
End Function

Public Function ReadCardOverrides() As Object
    Dim d As Object, ws As Worksheet, r As Long, lr As Long, k As String
    Set d = CreateObject("Scripting.Dictionary")
    If Not SheetExists(SH_CARD) Then Set ReadCardOverrides = d: Exit Function
    Set ws = BW().Worksheets(SH_CARD)
    lr = lastRow(ws, C_NM)
    For r = C_HDR + 1 To lr
        k = CleanText(ws.Cells(r, C_KEY).Value2)
        If k = "" Then k = MakeKey(CleanText(ws.Cells(r, C_SEB).Value2), CleanText(ws.Cells(r, C_NM).Value2))
        If k <> "|" Then
            d(k) = Array(CleanText(ws.Cells(r, C_OWN2).Value2), CleanText(ws.Cells(r, C_CARD2).Value2), CleanText(ws.Cells(r, C_BSET).Value2))
        End If
    Next r
    '사업별 담당자 변경(_담당자변경 시트) 이 있으면 담당자수정 우선 적용
    Dim wsC As Worksheet, v As Variant
    If SheetExists(SH_OWNCHG) Then
        Set wsC = BW().Worksheets(SH_OWNCHG)
        For r = 2 To lastRow(wsC, 1)
            k = CleanText(wsC.Cells(r, 1).Value2)
            If k <> "" And CleanText(wsC.Cells(r, 2).Value2) <> "" Then
                If d.Exists(k) Then v = d(k) Else v = Array("", "", "")
                v(0) = CleanText(wsC.Cells(r, 2).Value2)
                d(k) = v
            End If
        Next r
    End If
    Set ReadCardOverrides = d
End Function

'사업별 담당자 변경 블록 (부서담당자 시트 G열~) 작성. 입력값(J열)은 키 기준으로 보존
Private Sub WriteOwnerChangeBlock(ws As Worksheet)
    Dim keys As Collection, i As Long, it As clsProject, r As Long, old As Object, lr As Long, k As String
    Set old = gOwnOld
    If old Is Nothing Then Set old = CreateObject("Scripting.Dictionary")
    ws.Range(ws.Columns(O_TEAM), ws.Columns(O_KEY)).Clear
    ws.Cells(S_HDR - 1, O_TEAM).Value2 = "사업별 담당자 변경  (J열에 새 담당자 입력 → [담당자 변경 적용] → 모든 시트·슬라이서 반영. 여러 명은 쉼표로)"
    ws.Cells(S_HDR - 1, O_TEAM).Font.Bold = True: ws.Cells(S_HDR - 1, O_TEAM).Font.name = FONT_NM: ws.Cells(S_HDR - 1, O_TEAM).Font.color = ClrNavy()
    ws.Cells(S_HDR, O_TEAM).Resize(1, 5).Value2 = Array("팀명", "담당자(현재)", "사업명", "▶ 담당자 변경", "키")
    StyleHeader ws.Cells(S_HDR, O_TEAM).Resize(1, 5)
    r = S_HDR
    If ModelReady Then
        Set keys = SortedKeys()
        For i = 1 To keys.Count
            Set it = gItems(keys(i))
            r = r + 1
            ws.Cells(r, O_TEAM).Value2 = IIf(it.팀 = "", "(팀 미지정)", it.팀)
            ws.Cells(r, O_OWN).Value2 = IIf(it.담당자 = "", "(미지정)", it.담당자)
            ws.Cells(r, O_NM).Value2 = it.사업항목
            If old.Exists(it.key) Then ws.Cells(r, O_NEW).Value2 = old(it.key)
            ws.Cells(r, O_KEY).Value2 = it.key
        Next i
    End If
    If r > S_HDR Then
        StyleBody ws.Range(ws.Cells(S_HDR + 1, O_TEAM), ws.Cells(r, O_KEY))
        ws.Range(ws.Cells(S_HDR + 1, O_NEW), ws.Cells(r, O_NEW)).Interior.color = RGB(255, 255, 224)
    End If
    ws.Columns(O_TEAM).ColumnWidth = 14: ws.Columns(O_OWN).ColumnWidth = 14: ws.Columns(O_NM).ColumnWidth = 36
    ws.Columns(O_NEW).ColumnWidth = 16: ws.Columns(O_KEY).hidden = True
    AddButton ws, "btnOwnApply", "▶ 담당자 변경 적용", "담당자변경적용", ws.Cells(1, O_TEAM).Left, 5, 130, 26, ClrOrange()
End Sub

'버튼: J열 입력값을 _담당자변경(숨김) 에 저장 → 전체 새로고침
Public Sub 담당자변경적용()
    Dim ws As Worksheet, wsC As Worksheet, r As Long, lr As Long, k As String, nm As String, n As Long, wsK As Worksheet, rr As Long
    If Not SheetExists(SH_STAFF) Then Exit Sub
    Set ws = BW().Worksheets(SH_STAFF)
    Set wsC = GetSheet(SH_OWNCHG, True)
    wsC.Cells(1, 1).Resize(1, 3).Value2 = Array("키", "담당자", "사업명")
    lr = lastRow(ws, O_KEY)
    For r = S_HDR + 1 To lr
        k = CleanText(ws.Cells(r, O_KEY).Value2): nm = CleanText(ws.Cells(r, O_NEW).Value2)
        If k <> "" And nm <> "" Then
            n = n + 1
            wsC.Cells(n + 1, 1).Value2 = k: wsC.Cells(n + 1, 2).Value2 = nm: wsC.Cells(n + 1, 3).Value2 = ws.Cells(r, O_NM).Value2
        End If
    Next r
    wsC.Visible = xlSheetVeryHidden
    '단위사업카드 시트가 있으면 ▶담당자(수정) 열에도 동기화
    If SheetExists(SH_CARD) Then
        Set wsK = BW().Worksheets(SH_CARD)
        For rr = C_HDR + 1 To lastRow(wsK, C_NM)
            k = CleanText(wsK.Cells(rr, C_KEY).Value2)
            For r = S_HDR + 1 To lr
                If CleanText(ws.Cells(r, O_KEY).Value2) = k Then wsK.Cells(rr, C_OWN2).Value2 = CleanText(ws.Cells(r, O_NEW).Value2): Exit For
            Next r
        Next rr
    End If
    새로고침 True
    If SheetExists(SH_STAFF) Then BW().Worksheets(SH_STAFF).Activate
    MsgBox "사업별 담당자 변경 " & n & "건을 모든 시트에 반영했습니다.", vbInformation
End Sub

'---------------------------------------------------------------- 부서담당자 시트
Public Sub RefreshStaffSheet(Optional ByVal activateIt As Boolean = False)
    Dim ws As Worksheet, old As Object, r As Long, lr As Long, nm As String, names As Collection, i As Long
    Dim cnt As Object, k As Variant, it As clsProject, arr As Variant, j As Long, isNew As Boolean
    Set old = CreateObject("Scripting.Dictionary")   '성명 -> Array(팀, 직위, 비고)
    isNew = Not SheetExists(SH_STAFF)
    Set ws = GetSheet(SH_STAFF)
    If Not isNew Then
        lr = lastRow(ws, S_NAME)
        For r = S_HDR + 1 To lr
            nm = CleanText(ws.Cells(r, S_NAME).Value2)
            If nm <> "" Then old(nm) = Array(CleanText(ws.Cells(r, S_TEAM).Value2), CleanText(ws.Cells(r, S_POS).Value2), CleanText(ws.Cells(r, S_NOTE).Value2))
        Next r
    End If
    '담당 사업수
    Set cnt = CreateObject("Scripting.Dictionary")
    If ModelReady Then
        For Each k In gItems.keys
            Set it = gItems(k)
            arr = Split(it.담당자, ",")
            For j = LBound(arr) To UBound(arr)
                nm = CleanText(arr(j))
                If nm <> "" Then cnt(nm) = NzD(cnt(nm)) + 1
            Next j
        Next k
    End If

    Set gOwnOld = CreateObject("Scripting.Dictionary")
    If Not isNew Then
        For r = S_HDR + 1 To lastRow(ws, O_KEY)
            If CleanText(ws.Cells(r, O_KEY).Value2) <> "" Then gOwnOld(CleanText(ws.Cells(r, O_KEY).Value2)) = CleanText(ws.Cells(r, O_NEW).Value2)
        Next r
    End If
    ClearSheet ws
    ws.Tab.color = ClrOrange()
    ws.Range("A1").Value2 = "부서 및 업무담당자"
    ws.Range("A1").Font.Size = 16: ws.Range("A1").Font.Bold = True: ws.Range("A1").Font.name = FONT_NM
    ws.Range("A2").Value2 = "성명은 요구서 담당자 열에서 자동 추가됩니다. 팀/부서·직위를 입력한 뒤 [적용] 버튼을 누르면 메인 보드에 반영됩니다. 이름은 직접 추가/삭제해도 됩니다."
    ws.Range("A2").Font.Size = 9: ws.Range("A2").Font.color = ClrGray()
    ws.Cells(S_HDR, S_NAME).Resize(1, 5).Value2 = Array("성명", "팀/부서 (입력)", "직위 (입력)", "비고 (입력)", "담당 사업수(자동)")
    StyleHeader ws.Cells(S_HDR, S_NAME).Resize(1, 5)

    '기존 순서 유지 + 신규 추가
    r = S_HDR
    For Each k In old.keys
        r = r + 1
        arr = old(k)
        ws.Cells(r, S_NAME).Value2 = k
        ws.Cells(r, S_TEAM).Value2 = arr(0): ws.Cells(r, S_POS).Value2 = arr(1): ws.Cells(r, S_NOTE).Value2 = arr(2)
        ws.Cells(r, S_CNT).Value2 = NzD(cnt(k))
    Next k
    Set names = StaffNames()
    For i = 1 To names.Count
        If Not old.Exists(names(i)) Then
            r = r + 1
            ws.Cells(r, S_NAME).Value2 = names(i)
            ws.Cells(r, S_CNT).Value2 = NzD(cnt(names(i)))
            ws.Cells(r, S_TEAM).Interior.color = RGB(255, 242, 204)   '신규(팀 미입력) 표시
        End If
    Next i
    If r > S_HDR Then
        StyleBody ws.Range(ws.Cells(S_HDR + 1, 1), ws.Cells(r, 5))
        ws.Range(ws.Cells(S_HDR + 1, S_TEAM), ws.Cells(r, S_NOTE)).Interior.color = RGB(255, 255, 224)
    End If
    ws.Columns(S_NAME).ColumnWidth = 14: ws.Columns(S_TEAM).ColumnWidth = 22
    ws.Columns(S_POS).ColumnWidth = 12: ws.Columns(S_NOTE).ColumnWidth = 30: ws.Columns(S_CNT).ColumnWidth = 14
    ws.rows(1).RowHeight = 28
    AddHomeButton ws, 250, 5
    AddButton ws, "btnApply", "▶ 적용 (메인 새로고침)", "적용", 318, 5, 140, 26, ClrGreen()
    AddButton ws, "btnImport", "부서명단 파일", "부서명단파일선택", 464, 5, 100, 26, ClrBlue()
    WriteOwnerChangeBlock ws
    If activateIt Then
        ws.Activate
        ws.Range("A1").Select
        ActiveWindow.FreezePanes = False
        ws.Cells(S_HDR + 1, 1).Select
        ActiveWindow.FreezePanes = True
    End If
End Sub

'---------------------------------------------------------------- 단위사업카드 시트
Public Sub RefreshCardSheet(Optional ByVal activateIt As Boolean = False)
    Dim ws As Worksheet, ovr As Object, keys As Collection, i As Long, r As Long, it As clsProject, v As Variant
    Set ovr = ReadCardOverrides()
    Set ws = GetSheet(SH_CARD, True)
    ws.Tab.color = ClrOrange()
    ws.Range("A1").Value2 = "단위사업카드 (사업항목 ↔ 단위과제카드 ↔ 담당자)"
    ws.Range("A1").Font.Size = 16: ws.Range("A1").Font.Bold = True: ws.Range("A1").Font.name = FONT_NM
    ws.Range("A2").Value2 = "자동 열은 요구서·장표에서 채워집니다. 노란 ▶열에 입력 후 [적용]을 누르면 메인·대시보드에 반영됩니다. " & _
                            "'장표 사업항목 직접지정'은 요구서 사업명과 장표 사업항목명이 달라 매칭이 안 될 때 장표의 사업항목명1을 그대로 입력합니다."
    ws.Range("A2").Font.Size = 9: ws.Range("A2").Font.color = ClrGray()
    ws.Cells(C_HDR, 1).Resize(1, C_KEY).Value2 = Array("No", "세부사업", "사업항목(요구서)", "장표 사업항목명(자동)", "단위과제카드(자동)", _
        "담당자(자동)", "매칭", "예산현액", "▶담당자(수정)", "▶단위과제카드(수정)", "▶장표 사업항목 직접지정", "키")
    StyleHeader ws.Cells(C_HDR, 1).Resize(1, C_KEY)

    r = C_HDR
    If ModelReady Then
        Set keys = SortedKeys()
        For i = 1 To keys.Count
            Set it = gItems(keys(i))
            r = r + 1
            ws.Cells(r, C_NO).Value2 = r - C_HDR
            ws.Cells(r, C_SEB).Value2 = it.세부사업
            ws.Cells(r, C_NM).Value2 = it.사업항목
            ws.Cells(r, C_BNM).Value2 = it.장표명
            ws.Cells(r, C_CARD).Value2 = it.카드자동
            ws.Cells(r, C_OWN).Value2 = it.담당자자동
            ws.Cells(r, C_MATCH).Value2 = it.매칭
            ws.Cells(r, C_HY).Value2 = it.예산현액
            ws.Cells(r, C_OWN2).Value2 = it.담당자수정
            ws.Cells(r, C_CARD2).Value2 = it.카드수정
            ws.Cells(r, C_BSET).Value2 = it.장표지정
            ws.Cells(r, C_KEY).Value2 = it.key
            If it.매칭 = "장표없음" Or it.매칭 = "장표만" Or it.매칭 = "유사" Then ws.Cells(r, C_MATCH).Interior.color = RGB(255, 199, 206)
        Next i
    End If
    '사용자 입력 행(모델에 없는 키) 보존
    If ModelReady Then
        For Each v In ovr.keys
            If Not gItems.Exists(v) Then
                r = r + 1
                ws.Cells(r, C_NO).Value2 = r - C_HDR
                ws.Cells(r, C_NM).Value2 = "(이전 항목)"
                ws.Cells(r, C_OWN2).Value2 = ovr(v)(0): ws.Cells(r, C_CARD2).Value2 = ovr(v)(1): ws.Cells(r, C_BSET).Value2 = ovr(v)(2)
                ws.Cells(r, C_KEY).Value2 = v
            End If
        Next v
    End If

    If r > C_HDR Then
        StyleBody ws.Range(ws.Cells(C_HDR + 1, 1), ws.Cells(r, C_KEY))
        ws.Range(ws.Cells(C_HDR + 1, C_HY), ws.Cells(r, C_HY)).NumberFormat = "#,##0"
        ws.Range(ws.Cells(C_HDR + 1, C_OWN2), ws.Cells(r, C_BSET)).Interior.color = RGB(255, 255, 224)
        ws.Range(ws.Cells(C_HDR, 1), ws.Cells(r, C_KEY)).AutoFilter
    End If
    ws.Columns(C_NO).ColumnWidth = 5: ws.Columns(C_SEB).ColumnWidth = 18: ws.Columns(C_NM).ColumnWidth = 34
    ws.Columns(C_BNM).ColumnWidth = 30: ws.Columns(C_CARD).ColumnWidth = 34: ws.Columns(C_OWN).ColumnWidth = 14
    ws.Columns(C_MATCH).ColumnWidth = 9: ws.Columns(C_HY).ColumnWidth = 16: ws.Columns(C_OWN2).ColumnWidth = 14
    ws.Columns(C_CARD2).ColumnWidth = 30: ws.Columns(C_BSET).ColumnWidth = 30
    ws.Columns(C_KEY).hidden = True
    ws.rows(1).RowHeight = 28
    AddButton ws, "btnApply", "▶ 적용 (메인 새로고침)", "적용", 640, 5, 150, 26, ClrGreen()
    AddHomeButton ws, 800, 5
    If activateIt Then
        ws.Activate
        ws.Range("A1").Select
        ActiveWindow.FreezePanes = False
        ws.Cells(C_HDR + 1, 4).Select
        ActiveWindow.FreezePanes = True
    End If
End Sub

'---------------------------------------------------------------- 버튼용 진입점
Public Sub 부서담당자()
    If Not ModelReady Then
        On Error Resume Next
SpeedOn:         BuildModel: SpeedOff
        On Error GoTo 0
    End If
    SpeedOn "부서담당자 시트 생성 중..."
    RefreshStaffSheet True
    SpeedOff
End Sub

Public Sub 단위사업카드()
    If Not ModelReady Then
        On Error Resume Next
SpeedOn:         BuildModel: SpeedOff
        On Error GoTo 0
    End If
    SpeedOn "단위사업카드 시트 생성 중..."
    RefreshCardSheet True
    SpeedOff
End Sub

'적용: 입력값 반영 → 전체 새로고침 → 메인으로
Public Sub 적용()
    새로고침 True
    메인으로
End Sub
