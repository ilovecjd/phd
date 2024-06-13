Attribute VB_Name = "typedef"

Option Explicit
Option Base 1

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' #define global variable

' sheet name
Public Const PARAMETER_SHEET_NAME	= "GenDBoard"
Public Const DBOARD_SHEET_NAME 		= "dashboard"
Public Const PROJECT_SHEET_NAME 	= "project"
Public Const ACTIVITY_SHEET_NAME 	= "activity_struct"

' 주요 테이블의 제목
Public Const ORDER_PROJECT_TITLE	= "발주 프로젝트 현황"


Public Const P_TYPE_INTERNAL = 1
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 프로그램 동작을 위한 기본 정보들. 
' prologue함수가 parameter 시트에서 읽어 온다.
Public SimulTerm   As Integer  ' 시뮬레이션을 동작 시킬 기간(주)
Public avgProjects	As Double  ' 주당 발생하는 평균 발주 프로젝트 수

Public Hr_Init_H   As Integer  ' 최초에 보유한 고급 인력
Public Hr_Init_M   As Integer  ' 최초에 보유한 중급 인력
Public Hr_Init_L   As Integer  ' 최초에 보유한 초급 인력
Public Hr_LeadTime As Integer  ' 인력 충원에 걸리는 시간

Public Cash_Init   As Integer  ' 최초 보유 현금
Public Problem     As Integer  ' 프로젝트 생성 개수 (= 문제의 개수) / MakePrj 함수의 인자

''''''''''''''''''''
Public Const MAX_ACT    As Integer	= 4	 ' 최대 활동의 수
Public Const MAX_N_CF   As Integer  = 4	 ' 최대 CF의 갯수 (개발비를 최대로 나누어 받는 횟수)
Public Const W_INFO		As Integer 	= 12 ' 출력할 가로의 크기
Public Const H_INFO 	As Integer 	= 8  ' 출력할 세로의 크기

Public Const RND_HR_H = 20	' 고급 인력이 필요할 확율
Public Const RND_HR_M = 70	' 중급 인력이 필요할 확율

' 1: 2~4 / 2:5~12 3:13~26 4:27~52 5:53~80
Public Const MAX_PRJ_TYPE = 5	' 프로젝트 기간별로 타입을 구분한다.
Public Const RND_PRJ_TYPE1 = 20	' 1번 타입일 확율 1:  2~4 주
Public Const RND_PRJ_TYPE2 = 70	' 2번 타입일 확율 2:  5~12주
Public Const RND_PRJ_TYPE3 = 20	' 3번 타입일 확율 3: 13~26주
Public Const RND_PRJ_TYPE4 = 70	' 4번 타입일 확율 4: 27~52주
Public Const RND_PRJ_TYPE5 = 20	' 5번 타입일 확율 5: 53~80주


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

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Global Variable




' utility functions


' desc      : 프로그램 시작을 위한 기본적인 값들을 설정한다.
' return    : none
Sub Prologue()

	Dim rng 	As Range
	Set rng		= ThisWorkbook.Sheets(PARAMETER_SHEET_NAME).Range("b:c")

	SimulTerm 	= GetVariableValue(rng, "SimulTerm")
	avgProjects = GetVariableValue(rng, "avgProjects")
	Hr_Init_H 	= GetVariableValue(rng, "Hr_Init_H")
	Hr_Init_M 	= GetVariableValue(rng, "Hr_Init_M")
	Hr_Init_L 	= GetVariableValue(rng, "Hr_Init_L")
	Hr_LeadTime = GetVariableValue(rng, "Hr_LeadTime")
	Cash_Init 	= GetVariableValue(rng, "Cash_Init")
	Problem 	= GetVariableValue(rng, "ProblemCnt")     ' MakePrj 함수의 인자

	' 속도 향상을 위해서
	Application.ScreenUpdating = False
	Application.Calculation = xlCalculationManual
	Application.EnableEvents = False
	ActiveSheet.DisplayPageBreaks = False

End Sub

Public Function Epilogue()

	Application.ScreenUpdating = True    
	Application.Calculation = xlAutomatic
	Application.EnableEvents = True   
	' 이 항목은 굳이 다시 켜지 말자. ActiveSheet.DisplayPageBreaks = True

End Function


'' 주어진 Range 에서 해당 스트링의 다음열의 값을 가져온다
Public Function GetVariableValue(rng As Range, variableName As String) As Variant
    Dim dataArray As Variant
    Dim matchIndex As Variant

    ' 범위를 배열로 변환
    dataArray = rng.Value

    ' 변수 이름이 있는 위치를 찾기
    matchIndex = Application.Match(variableName, Application.Index(dataArray, 0, 1), 0)
    
    ' 변수 이름이 있는 경우 값 반환
    If Not IsError(matchIndex) Then
        GetVariableValue = dataArray(matchIndex, 2)
    Else
        GetVariableValue = "Variable not found" 'song ==> 예외 처리는 나중에 하자.
    End If

End Function


' lambda(평균 발생률)를 인자로 받아 포아송 분포를 따르는 랜덤 값을 반환합니다.
Public Function PoissonRandom(lambda As Double) As Integer
    Dim L As Double
    Dim p As Double
    Dim k As Integer
    
    L = Exp(-lambda)
    p = 1
    k = 0
    
    Do
        k = k + 1
        p = p * Rnd()
    Loop While p > L
    
    PoissonRandom = k - 1
End Function
