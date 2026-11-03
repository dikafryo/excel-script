VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'=====================================================================
' clsProject : 사업항목(1레벨) 단위 데이터 객체
'=====================================================================
Public key As String            '세부사업|사업항목 정규화 키
Public 세부사업 As String
Public 사업항목 As String        '표시명 (요구서 우선)
Public 장표명 As String          '장표(102-B) 사업항목명1
Public 장표키 As String
Public 정책사업 As String
Public 단위사업 As String
Public 요구액 As Double          '요구서 예산요구액 (원)
Public 전년도액 As Double        '요구서 전년도예산액 (원)
Public 담당자자동 As String      '요구서 AC열에서 추출
Public 담당자수정 As String      '단위사업카드 시트에서 사용자 수정
Public 팀 As String
Public 카드자동 As String        '장표 단위업무(과제카드) 자동
Public 카드수정 As String
Public 장표지정 As String        '사용자가 직접 지정한 장표 사업항목명
Public 예산현액 As Double
Public 원인행위 As Double
Public 집행액 As Double          '지출 - 일상경비교부 + 일상경비지급
Public 본예산현액 As Double      '배정/재배정(예산)
Public 이월현액 As Double        '명시/사고/계속비
Public 매칭 As String
Public 순번 As Long
'--- V8
Public 차수 As String            '이 금액이 나온 차수 라벨 (본예산 / 1차 추경 / 2차 추경 ...)
Public 차수번호 As Long          '0 = 본예산, 1.. = 추경 추가 순서
Public 산출합 As Double          '산출내역(세목코드 행) 금액 합 (원). 세목코드 열이 없으면 -1
Public 변경 As String            '추경 파일에서: 추가 / 변경 / 동일
Public 장표예산 As Double        '장표 예산금액(기준값, 원) = 집행예산구분 '예산' 행 예산현액
Public 산출식합 As Double        'V8: 최신 산출내역(산출식 계산) 합 (원, 세목 행마다 금액 단위로 반올림)
Public 산출식허용 As Double      'V8: 끝수 허용폭 (원) = 천원 단위로 떨어지지 않는 세목 행 수 x 1천원
Public 산출식상태 As Long        'V8: 0 = 계산됨, 1 = 산출식을 읽지 못한 세목 행 있음, 2 = 산출내역 열 없음
Public 산출식전합 As Double      'V8: 다른 해(본예산: 작년) / 기정 산출내역 계산 합 (원)
Public 산출식전허용 As Double
Public 산출식전상태 As Long      '0 = 계산됨, 1 = 읽지 못함, 2 = 열 없음 또는 문구 없음
Public 산출식전종류 As String    '"작년" (본예산 파일) / "기정" (추경 파일)

Public Property Get 담당자() As String
    If Trim(담당자수정) <> "" Then 담당자 = Trim(담당자수정) Else 담당자 = 담당자자동
End Property

Public Property Get 과제카드() As String
    If Trim(카드수정) <> "" Then 과제카드 = Trim(카드수정) Else 과제카드 = 카드자동
End Property

Public Property Get 잔액() As Double
    잔액 = 예산현액 - 집행액
End Property

Public Property Get 집행률() As Double
    If 예산현액 <> 0 Then 집행률 = 집행액 / 예산현액 Else 집행률 = 0
End Property

Public Property Get 상태() As String
    상태 = StatusText(예산현액, 집행률, 매칭)
End Property

