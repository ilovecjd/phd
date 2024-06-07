Attribute VB_Name = "typedef"
Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' #define
Public Const MAX_ACT    As Integer	= 4	 ' 최대 활동의 수
Public Const MAX_N_CF   As Integer  = 4	 ' 최대 CF의 갯수 (개발비를 최대로 나누어 받는 횟수)
Public Const W_INFO		As Integer 	= 12 ' 출력할 가로의 크기
Public Const H_INFO 	As Integer 	= 8  ' 출력할 세로의 크기

Public Const RND_HR_H = 20	' 고급 인력이 필요할 확율
Public Const RND_HR_M = 70	' 중급 인력이 필요할 확율

' 1: 2~4 / 2:5~12 3:13~26 4:27~52 5:53~80
Public Const MAX_PRJ_TYPE = 5	' 프로젝트 기간별로 타입을 구분한다.
Public RND_PRJ_TYPE1 = 20	' 1번 타입일 확율 1:  2~4 주
Public RND_PRJ_TYPE2 = 70	' 2번 타입일 확율 2:  5~12주
Public RND_PRJ_TYPE3 = 20	' 3번 타입일 확율 3: 13~26주
Public RND_PRJ_TYPE4 = 70	' 4번 타입일 확율 4: 27~52주
Public RND_PRJ_TYPE5 = 20	' 5번 타입일 확율 5: 53~80주


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


' 프로젝트 생성의 정보를 담는 구조체
Type EnvProject
	Probabilty As Double
	MinDuration As Integer
	MaxDuration As Integer
	NumPartten	As Integer
End Type

' 활동생성의 정보를 담는 구조체
Type EnvActivity
    OccurActivityType    As Integer  ' 1-분석설계/2-구현/3-단테/4-통테/5-유지보수
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
