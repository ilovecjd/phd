Attribute VB_Name = "typedef"

Option Explicit
Option Base 1

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Define Global Variable

' sheet name
Public Const PARAMETER_SHEET_NAME	= "GenDBoard"
Public Const DBOARD_SHEET_NAME 		= "dashboard"
Public Const PROJECT_SHEET_NAME 	= "project"
Public Const ACTIVITY_SHEET_NAME 	= "activity_struct"

' 주요 테이블의 제목
Public Const ORDER_PROJECT_TITLE	= "발주 프로젝트 현황"

Public Const P_TYPE_EXTERNAL = 0 ' 외부(발주)프로젝트
Public Const P_TYPE_INTERNAL = 1 ' 내부 프로젝트

Private gExcelInitialized 	As Boolean	' 전역 변수들이 초기화 되었는지 확인하는 플래그. 초기화 되면 1
Private gTableInitialized 	As Boolean	' 전역 테이블이 초기화 되었는지 확인하는 플래그. 초기화 되면 1
Private gTotalProjectNum	As Integer	' 발생한 프로젝트의 총 갯수 (누계)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' ' 프로그램 동작을 위한 기본 정보들. 
Private EnvVar				As EnvExcel
Private OrderTable()		As Variant 		' 발주된 프로젝트들을 관리하는 테이블
Private ProjectInfoTable()	As clsProject	' 모든 프로제트들을 담고 있는 테이블


Public PrintDurationTable()	As Variant 		' 사용하기 편하게 모든 월을 넣어 놓는다. 


''''''''''''''''''''
' 프로젝트 생성과 관련된 상수들
Public Const MAX_ACT    	As Integer	= 4	 ' 최대 활동의 수
Public Const MAX_N_CF   	As Integer  = 4	 ' 최대 CF의 갯수 (개발비를 최대로 나누어 받는 횟수)
Public Const W_INFO			As Integer 	= 12 ' 출력할 가로의 크기
Public Const H_INFO 		As Integer 	= 8  ' 출력할 세로의 크기

Public Const RND_HR_H = 20	' 고급 인력이 필요할 확율
Public Const RND_HR_M = 70	' 중급 인력이 필요할 확율

' 1: 2~4 / 2:5~12 3:13~26 4:27~52 5:53~80
Public Const MAX_PRJ_TYPE 	= 5	' 프로젝트 기간별로 타입을 구분한다.
Public Const RND_PRJ_TYPE1 	= 20	' 1번 타입일 확율 1:  2~4 주
Public Const RND_PRJ_TYPE2 	= 70	' 2번 타입일 확율 2:  5~12주
Public Const RND_PRJ_TYPE3 	= 20	' 3번 타입일 확율 3: 13~26주
Public Const RND_PRJ_TYPE4 	= 70	' 4번 타입일 확율 4: 27~52주
Public Const RND_PRJ_TYPE5 	= 20	' 5번 타입일 확율 5: 53~80주


' #define end
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'WorkBook 전체의 전역변수를 담을 구조체
Type EnvExcel

	ExcelInitialize 		As Boolean	' 전역 변수들이 초기화 되었는지 확인하는 플래그. 초기화 되면 1
	TableInitialize 		As Boolean	' 전역 테이블이 초기화 되었는지 확인하는 플래그. 초기화 되면 1

	SimulationDuration		As Integer  ' 시뮬레이션을 동작 시킬 기간(주)
	AvgProjects				As Double  	' 주당 발생하는 평균 발주 프로젝트 수
	Hr_Init_H   			As Integer  ' 최초에 보유한 고급 인력
	Hr_Init_M   			As Integer  ' 최초에 보유한 중급 인력
	Hr_Init_L   			As Integer  ' 최초에 보유한 초급 인력
	Hr_LeadTime 			As Integer  ' 인력 충원에 걸리는 시간
	Cash_Init   			As Integer  ' 최초 보유 현금
	Problem     			As Integer  ' 프로젝트 생성 개수 (= 문제의 개수) / MakePrj 함수의 인자

End Type


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

'' song ==> 사용 하지 않음
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



' Public functions
Public Property Get ExcelInitialized() As Boolean
	ExcelInitialized = gExcelInitialized
End Property

Public Property Let ExcelInitialized(value  As Boolean) 
	gExcelInitialized = value
End Property


Public Property Get TableInitialized() As Boolean
	TableInitialized = gTableInitialized
End Property

Public Property Let TableInitialized(value As Boolean) 
	gTableInitialized = value
End Property


Public Property Get GetOrderTable() As Variant
	GetOrderTable = OrderTable
End Property

Public Property Get GetProjectInfoTable() As Variant
	GetProjectInfoTable = ProjectInfoTable
End Property



' utility functions


' desc      : 프로그램 시작을 위한 기본적인 값들을 설정한다. 모든 프로시저들이 시작시 호출 하여야 한다.
' return    : none
Sub Prologue()

	
	If gExcelInitialized = 0 Then		' 전역 변수들이 초기화 되었는지 확인하는 플래그. 초기화 되면 1

		Dim rng 	As Range
		Set rng		= ThisWorkbook.Sheets(PARAMETER_SHEET_NAME).Range("b:c")

		EnvVar.SimulationDuration	= GetVariableValue(rng, "SimulTerm")
		EnvVar.AvgProjects 			= GetVariableValue(rng, "avgProjects")
		EnvVar.Hr_Init_H 			= GetVariableValue(rng, "Hr_Init_H")
		EnvVar.Hr_Init_M 			= GetVariableValue(rng, "Hr_Init_M")
		EnvVar.Hr_Init_L 			= GetVariableValue(rng, "Hr_Init_L")
		EnvVar.Hr_LeadTime 			= GetVariableValue(rng, "Hr_LeadTime")
		EnvVar.Cash_Init 			= GetVariableValue(rng, "Cash_Init")
		EnvVar.Problem 				= GetVariableValue(rng, "ProblemCnt")
		
		gExcelInitialized = 1		' 전역 변수들이 초기화 되었는지 확인하는 플래그. 초기화 되면 1

	End If

	If gTableInitialized = 0 Then ' Table 들이 만들어 졌는가?

		ReDim OrderTable(2,EnvVar.SimulationDuration)

		If CreateOrderTable() = False Then			
			MsgBox "CreateOrderTable Error", vbExclamation 			
			Exit Sub
		End If


		ReDim ProjectInfoTable(2, gTotalProjectNum)

		If CreateProjects() = False Then			
			MsgBox "CreateProjects Error", vbExclamation 			
			Exit Sub
		End If

		Dim i As Integer

		ReDim PrintDurationTable(1, EnvVar.SimulationDuration)

		For i = 1 to (EnvVar.SimulationDuration )
			PrintDurationTable(1, i) = i
		Next

		gTableInitialized = 1

	End If

	' 속도 향상을 위해서
	' Application.ScreenUpdating = False
	' Application.Calculation = xlCalculationManual
	' Application.EnableEvents = False
	' ActiveSheet.DisplayPageBreaks = False

End Sub

Private Function CreateOrderTable() As Boolean

	Dim week 			As Integer 
	Dim projectCount	As Integer
	Dim sum 			As Integer

	CreateOrderTable = True

	If gExcelInitialized = 0 Then
		CreateOrderTable = False
		MsgBox "EnvVars is not set", vbExclamation 
		Exit Function		
	End If

	If gTableInitialized = 0 Then ' Table 들이 만들어 졌는가?

		ReDim OrderTable(2,EnvVar.SimulationDuration)

		For week = 1 To EnvVar.SimulationDuration			
			projectCount 		= PoissonRandom(EnvVar.AvgProjects) ' 이번주 발생하는 프로젝트 갯수
			OrderTable(1,week)	= sum
			OrderTable(2,week)	= projectCount

			' 이번주 까지 발생한 프로젝트 갯수. 다음주에 기록된다. ==> 이전주까지 발생한 프로젝트 갯수후위연산. vba에서 do while 문법 모름... ㅎㅎ
			sum 	= sum + projectCount			
		Next

		gTotalProjectNum = sum
		gTableInitialized = 1

	End If

End Function

Private Function CreateProjects() As Boolean

	Dim week			As Integer
	Dim id 				As Integer
	Dim startPrjNum		As Integer
	Dim endPrjNum		As Integer
	Dim preTotal		As Integer		
	Dim tempPrj 		As clsProject	

	CreateProjects = True

	If gTotalProjectNum <= 0 Then
		MsgBox "gTotalProjectNum is 0", vbExclamation 
		CreateProjects = False
		Exit Function
	End If

	'프로젝트들을 생성한다. 
	ReDim ProjectInfoTable(gTotalProjectNum)

	For week = 1 to EnvVar.SimulationDuration
		
		preTotal 	= OrderTable(1,week)			' 이전 기간 까지 발생한 프로젝트 누적 갯수
		startPrjNum	= preTotal + 1 					' 이번 기간 시작프로젝트 번호
		endPrjNum 	= OrderTable(2,week) + preTotal	' 이번 기간 마지막 프로젝트 번호
		
		If startPrjNum = 0 Then
			GoTo Continue 
		End If

		If startPrjNum > endPrjNum Then
			GoTo Continue 
		End If	

		' 이번 주에 발생한 프로젝트들을 생성한다.
		For id = startPrjNum to endPrjNum ' 
			Set tempPrj 	= New clsProject
			Call tempPrj.Init(P_TYPE_EXTERNAL, id, PROJECT_SHEET_NAME,week) 
			Set ProjectInfoTable(id) = tempPrj
			'Call tempPrj.PrintInfo()
		Next

		Continue: 

	Next
	
End Function

Public Function Epilogue()

	' Application.ScreenUpdating = True    
	' Application.Calculation = xlAutomatic
	' Application.EnableEvents = True   
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
