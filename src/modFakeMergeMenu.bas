Attribute VB_Name = "modFakeMergeMenu"
Option Explicit
' =====================================================================================
'  modFakeMergeMenu : 리본 [셀정리하기 ▶ 병합효과] 통합 메뉴 (tag "FakeMergeMenu")  [v2.6]
'   1 병합효과 적용 (modFakeMerge.Fake병합)   2 병합효과 취소 (modFakeMergeUndo.병합취소)
'   폼에 전·후 모양을 미리보기로 보여 준다: 값은 그대로 두고 글자만 숨겨 계산식이 계속 동작함을 설명
' =====================================================================================

Public Sub FakeMergeMenu(control As IRibbonControl)
    Dim choice As Long, dummy As Long
    choice = UI_ChooseEx("병합효과", _
        "셀을 실제로 병합하지 않고, 같은 값이 이어지는 구간의 글자를 숨겨 병합한 것처럼 보이게 합니다." & vbCrLf & _
        "값은 셀마다 그대로 남아 있어 정렬·필터·SUMIF 같은 계산식이 그대로 동작합니다.", _
        Array("병합효과 적용   (같은 값 구간 → 가운데 한 칸만 표시)", _
              "병합효과 취소   (숨긴 글자 · 테두리 원래대로)"), _
        Array("merge:after", "merge:before"), Empty, 0, dummy)
    Select Case choice
        Case 1: Fake병합 Nothing
        Case 2: 병합취소 Nothing
    End Select
End Sub
