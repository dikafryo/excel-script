Attribute VB_Name = "modOfficeMenu"
Option Explicit
' =====================================================================================
'  modOfficeMenu : 리본 [OFFICE파일합치기] 통합 메뉴 (tag "OfficeMergeMenu")  [v2.6]
'   1 폴더 안 한글(HWP·HWPX) → HWPX 하나로     ├ 추가선택 1 모든 글꼴·서식·여백 통합 / 2 파일마다 구역 나눠 구분   (modMergeHwp.MH_Run)
'   2 폴더 안 한글(HWP·HWPX) → PDF             ├ 추가선택 1 각각 / 2 하나로(목차·링크) / 3 하나로(파일만)        (modHwpToPdf.HP_Run)
'   3 폴더 안 PDF 합치기                        ├ 추가선택 1 목차·링크 / 2 파일만                                 (modMergePdf.MP_Run)
'   4 폴더 안 Office 파일 전부 PDF 하나로        (한글·PDF·Word·Excel·PPT·이미지, modOfficeMerge.OfficeMergeToPdf)
' =====================================================================================
Private Const APP_TITLE As String = "OFFICE파일합치기"

Public Sub OfficeMergeMenu(control As IRibbonControl)
    Dim choice As Long, sub1 As Long
    choice = UI_Choose(APP_TITLE, _
        "폴더 안 파일을 합치거나 변환합니다. (결과는 같은 폴더에 저장, 목록은 시트에 기록. ⑥은 현재 문서의 시트)" & vbCrLf & _
        "※ ①② 는 한글(HWP) 프로그램, ②의 하나로 합치기·③·④ 는 MS Office(Word) 가 설치되어 있어야 합니다.", _
        Array("한글(HWP·HWPX) 파일들 → HWPX 하나로 합치기" & vbCrLf & "(빈 페이지 없이)", _
              "한글(HWP·HWPX) 파일들 → PDF 로 변환", _
              "PDF 파일들 → PDF 하나로 합치기", _
              "Office 파일 전부 → PDF 하나로" & vbCrLf & "(한글·PDF·Word·Excel·PPT·이미지)", _
              "엑셀 파일들 합치기" & vbCrLf & "(폴더 안 엑셀 파일의 시트를 이름별 또는 한 시트로)", _
              "이 문서의 시트들 → 하나로 합치기" & vbCrLf & "(시트 여러 개 선택 = 선택 시트만, 1개 = 전체)"))
    Select Case choice
        Case 1
            sub1 = UI_Choose(APP_TITLE & " - 한글 합치기", "합칠 때 서식을 어떻게 할까요?  (한글(HWP) 프로그램 필요)", _
                Array("모든 글꼴 · 서식 · 여백을 하나로 통합   (첫 문서 기준, 빈 페이지 없음)", _
                      "파일마다 구역을 나눠 글꼴 · 서식 · 여백을 구분   (각 파일이 새 페이지에서 시작)"))
            If sub1 = 0 Then Exit Sub
            MH_Run sub1
        Case 2
            sub1 = UI_Choose(APP_TITLE & " - PDF 변환", "변환 결과를 어떻게 만들까요?  (한글(HWP) 필요 · ②③ 하나로 합치기는 MS Office(Word) 도 필요)", _
                Array("각각의 PDF 파일로 변환   (파일마다 같은 이름 .pdf)", _
                      "하나의 PDF 로 합치기   (1페이지 목차 · 링크 추가)", _
                      "하나의 PDF 로 합치기   (파일만 이어 붙이기)"))
            If sub1 = 0 Then Exit Sub
            HP_Run sub1
        Case 3
            sub1 = UI_Choose(APP_TITLE & " - PDF 합치기", "합친 PDF 앞에 목차 페이지를 넣을까요?  (MS Office(Word) 필요)", _
                Array("목차 페이지 · 링크 추가", "파일만 이어 붙이기"))
            If sub1 = 0 Then Exit Sub
            MP_Run sub1
        Case 4
            OfficeMergeToPdf Nothing
        Case 5
            sub1 = UI_Choose(APP_TITLE & " - 엑셀 파일 합치기", "폴더 안 엑셀 파일의 시트를 어떻게 통합할까요?", _
                Array("시트 이름별로 통합   (같은 이름의 시트끼리 각각 한 시트로)", _
                      "모든 시트 통합   (전부 한 시트(Combined)에 이어 붙임)"))
            If sub1 = 0 Then Exit Sub
            MX_Run sub1
        Case 6
            MergeSheetsToCollection Nothing
    End Select
End Sub
