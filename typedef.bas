Attribute VB_Name = "typedef"
Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' #define
Public Const MAX_ACT    As Integer	= 6	 ' 최대 활동의 수
Public Const MAX_N_CF   As Integer  = 4	 ' 최대 CF의 갯수 (개발비를 최대로 나누어 받는 횟수)
Public Const W_INFO		As Integer 	= 12 ' 출력할 가로의 크기
Public Const H_INFO 	As Integer 	= 8  ' 출력할 세로의 크기

Public Const RND_HR_H = 20	' 고급 인력이 필요할 확율
Public Const RND_HR_M = 70	' 중급 인력이 필요할 확율
' #define end
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' 활동의 정보를 담는 구조체
Type Activity
    ActivityType    As Integer  ' 1-분석설계/2-구현/3-단테/4-통테/5-유지보수
    Duration        As Integer  ' 활동의 기간
    StartDate       As Integer  ' 활동의 시작
    EndDate         As Integer  ' 활동의 끝
    HighSkill       As Integer  ' 필요한 고급 인력 수
    MidSkill        As Integer  ' 필요한 중급 인력 수
    LowSkill        As Integer  ' 필요한 초급 인력 수
End Type




Public Function max(x, y As Variant) As Variant
    max = IIf(x > y, x, y)
End Function
