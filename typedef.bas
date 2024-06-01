Attribute VB_Name = "typedef"
Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' #define
Public Const MAX_ACT    As Integer  = 6' 최대 활동의 수
Public Const MAX_N_CF   As Integer  = 4' 최대 CF의 갯수 (개발비를 최대로 나누어 받는 횟수)
Public Const W_INFO	As Integer = 12 ' 출력할 가로의 크기
Public Const H_INFO As Integer = 8 ' 출력할 세로의 크기

Public Const RND_HR_H = 20	' 고급 인력이 필요할 확율
Public Const RND_HR_M = 70	' 중급 인력이 필요할 확율
' #define end
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


' 활동의 정보를 담는 구조체
Type ACTIVITY_
    duration    As Integer  ' 활동의 기간
    start       As Integer  ' 활동의 시작
    end         As Integer  ' 활동의 끝
    hr_H        As Integer  ' 필요한 고급 인력 수
    hr_M        As Integer  ' 필요한 중급 인력 수
    hr_L        As Integer  ' 필요한 초급 인력 수
End Type





''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' 이하 삭제 예정
Type odl_typeProject
    num                 As Double   ' 프로젝트의 번호
    period              As Double   ' 프로젝트의 총 기간
    orderDate           As Integer  '발주일
    possibleDate        As Integer  ' 시작가능일
    startDate           As Integer  ' 시작일
    profit              As Integer  ' 수익 (HR종속)
    experience          As Integer  ' 경험 (0 : 무경험 1:유경험)
    successPercentage   As Integer  ' 성공확율
    CF                  As Integer  ' Cash Flower
    N_CF                As Integer  ' Number of CF (비용 지급 횟수 1cf 40%, 2cf 30%, 3cf 40% 등)
End Type


Type Occurrence
    Opt         As Double   ' 낙관적인 전망
    ML          As Double   ' 일반적인 전망
    Pess        As Double   ' 비관적인 전망
End Type

'' 이하 삭제 예정 끝
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Public Function max(x, y As Variant) As Variant
    max = IIf(x > y, x, y)
End Function
