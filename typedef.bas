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
Public gTotalProjectNum	As Integer	' 발생한 프로젝트의 총 갯수 (누계)

Public gWsGenDBoard			As Worksheet	' 워크시트들을 전역으로 미리 구해 놓는다.
Public gWsDashboard			As Worksheet
Public gWsProject			As Worksheet
Public gWsActivity_Struct	As Worksheet
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' ' 프로그램 동작을 위한 기본 정보들. 
Public gExcelEnv			As EnvExcel
Public gOrderTable()		As Variant 		' 발주된 프로젝트들을 관리하는 테이블
Public gProjectTable()	As clsProject	' 모든 프로제트들을 담고 있는 테이블


Public gPrintDurationTable()	As Variant 		' 사용하기 편하게 모든 월을 넣어 놓는다. 


''''''''''''''''''''
' 프로젝트 생성과 관련된 상수들
Public Const MAX_ACT    	As Integer	= 4	 ' 최대 활동의 수
Public Const MAX_N_CF   	As Integer  = 3	 ' 최대 CF의 갯수 (개발비를 최대로 나누어 받는 횟수)
Public Const W_INFO			As Integer 	= 16 ' 출력할 가로의 크기
Public Const H_INFO 		As Integer 	= 8  ' 출력할 세로의 크기

Public Const RND_HR_H = 20	' 고급 인력이 필요할 확율
Public Const RND_HR_M = 70	' 중급 인력이 필요할 확율

' 1: 2~4 / 2:5~12 3:13~26 4:27~52 5:53~80
Public Const MAX_PRJ_TYPE 	As Integer	= 5		' 프로젝트 기간별로 타입을 구분한다.
Public Const RND_PRJ_TYPE1 	As Integer	= 20	' 1번 타입일 확율 1:  2~4 주
Public Const RND_PRJ_TYPE2 	As Integer	= 70	' 2번 타입일 확율 2:  5~12주
Public Const RND_PRJ_TYPE3 	As Integer	= 20	' 3번 타입일 확율 3: 13~26주
Public Const RND_PRJ_TYPE4 	As Integer	= 70	' 4번 타입일 확율 4: 27~52주
Public Const RND_PRJ_TYPE5 	As Integer	= 20	' 5번 타입일 확율 5: 53~80주

''''''''''''''''''''
'' 출력과 로드를 위한 상수들
Public Const ORDER_TABLE_INDEX 		As Long	= 1		' 
Public Const DONG_TABLE_INDEX 		As Long	= 6		' 
Public Const PROJECT_TABLE_INDEX 	As Long	= 3		' 

' #define end
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'WorkBook 전체의 전역변수를 담을 구조체
Type EnvExcel
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


' Public functions
Public Property Get GetExcelEnv() As EnvExcel
	GetExcelEnv 		= gExcelEnv
End Property

Public Property Get GetExcelInitialized() As Boolean
	GetExcelInitialized = gExcelInitialized
End Property

Public Property Let LetExcelInitialized(value  As Boolean) 
	gExcelInitialized = value
End Property


Public Property Get GetTableInitialized() As Boolean
	GetTableInitialized = gTableInitialized
End Property

Public Property Let LetTableInitialized(value As Boolean) 
	gTableInitialized = value
End Property


Public Property Get GetTotalProjectNum() As Integer
	GetTotalProjectNum = gTotalProjectNum
End Property

Public Property Get GetOrderTable() As Variant
	GetOrderTable = gOrderTable
End Property

Public Property Get GetProjectTable() As Variant
	GetProjectTable = gProjectTable
End Property



' utility functions

' desc      : 프로그램 시작을 위한 기본적인 값들을 설정한다. 모든 프로시저들이 시작시 호출 하여야 한다.
' return    : none
Sub Prologue(TableInit As Integer )
On Error GoTo ErrorHandler

	Dim i As Integer
	
	If gExcelInitialized = 0 Then		' 전역 변수들이 초기화 되었는지 확인하는 플래그. 초기화 되면 1
		' 한번만 하면 되는 것들은 여기에
		Call LoadExcelEnv()

		ReDim gPrintDurationTable(1, gExcelEnv.SimulationDuration)	
		For i = 1 to (gExcelEnv.SimulationDuration )
			gPrintDurationTable(1, i) = i
		Next

		gExcelInitialized = 1		' 전역 변수들이 초기화 되었는지 확인하는 플래그. 초기화 되면 1
	End If

	' 테이블들은 새로 생성하거나 기존것을 로드하거나. 
	' 예외 처리는 하지말고 사용자가 조심해서 사용하도록 하자.
	gTableInitialized = TableInit
	If gTableInitialized = 0 Then ' Table 들이 만들어지지 않았으면 테이블 생성
		Call BuildTables()		
	Else
		Call LoadTablesFromExcel() ' 만들어져 있으면 기존의 엑셀 시트에서 값들을 로드
	End If 

	gTableInitialized = 1  'Prologue()를 호출하기전에 설정하고 호출한다.

	' 속도 향상을 위해서
	' Application.ScreenUpdating = False
	' Application.Calculation = xlCalculationManual
	' Application.EnableEvents = False
	' ActiveSheet.DisplayPageBreaks = False

	Exit Sub

ErrorHandler:
    Call HandleError("Prologue", Err.Description)

End Sub

Sub BuildTables()

	Call CreateOrderTable()	
	Call CreateProjects()

End Sub

Sub LoadTablesFromExcel()

	Call LoadOrderTable()	
	Call LoadProjects()

End Sub

Private Function LoadOrderTable() As Boolean

	ReDim gOrderTable(2,gExcelEnv.SimulationDuration)

	Dim startIndex As Long	
	startIndex = ORDER_TABLE_INDEX
	startIndex = startIndex + 2

	With gWsDashboard
		gOrderTable = .Range(.Cells(startIndex,2),.Cells(startIndex+1, gExcelEnv.SimulationDuration+1)).Value		
	End With

	gTotalProjectNum = gOrderTable(1,gExcelEnv.SimulationDuration) + gOrderTable(2,gExcelEnv.SimulationDuration)

End Function

Private Function LoadProjects() As Boolean

	Dim prjID 		As Integer
	Dim startRow	As Long
	Dim endRow 		As Long
	Dim prjInfo 	As Variant
	Dim iTemp 		As Integer '
	Dim tempPrj 	As clsProject

	For prjID = 1 to  gTotalProjectNum

		tempPrj = New clsProject
		startRow = PROJECT_TABLE_INDEX + (prjID-1) * H_INFO + 1
		endRow = startRow + H_INFO - 1

		With gWsProject
			prjInfo = .Range(.Cells(startRow,1),Cells(endRow,W_INFO)).Value
		End With


		Dim i As Integer
		Dim j As Integer
		Dim k As Integer

		i= 1 : j = 1
		tempPrj.ProjectType 		= prjInfo(i,j) : j = j + 1
		tempPrj.ProjectNum			= prjInfo(i,j) : j = j + 1
		tempPrj.OrderDate			= prjInfo(i,j) : j = j + 1
		tempPrj.PossibleStartDate	= prjInfo(i,j) : j = j + 1
		tempPrj.ProjectDuration		= prjInfo(i,j) : j = j + 1
		tempPrj.StartDate			= prjInfo(i,j) : j = j + 1
		tempPrj.Profit				= prjInfo(i,j) : j = j + 1
		tempPrj.Experience			= prjInfo(i,j) : j = j + 1
		tempPrj.SuccessProbability	= prjInfo(i,j) : j = j + 1
		tempPrj.NumCashFlows		= MAX_N_CF
		For k = 1 To MAX_N_CF
			tempPrj.CashFlows(k)	= prjInfo(i,j) : j = j + 1			
		Next		
		tempPrj.FirstPayment 		= prjInfo(i,j) : j = j + 1
		tempPrj.MiddlePayment 		= prjInfo(i,j) : j = j + 1
		tempPrj.FinalPayment 		= prjInfo(i,j) : j = j + 1


		i = 2 : j = 1		
		tempPrj.NumActivities		= prjInfo(i,j)
		
		j = 10 ' 여기는 늘 조심하자	
		tempPrj.FirstPaymentMonth 	= prjInfo(i,j) : j = j + 1
		tempPrj.MiddlePaymentMonth	= prjInfo(i,j) : j = j + 1
		tempPrj.FinalPaymentMonth 	= prjInfo(i,j) : j = j + 1
		
		Dim tempAct 	As Activity
		For i = 3 To (tempPrj.NumActivities + i - 1)
			j 						= 2
			tempAct.Duration		= prjInfo(i,j) : j = j + 1			
			tempAct.StartDate		= prjInfo(i,j) : j = j + 1			
			tempAct.EndDate			= prjInfo(i,j) : j = j + 1			
			tempAct.HighSkill		= prjInfo(i,j) : j = j + 1			
			tempAct.MidSkill		= prjInfo(i,j) : j = j + 1			
			tempAct.LowSkill		= prjInfo(i,j) : j = j + 1		
			tempPrj.Activities(i-2)	= tempAct
		Next
		
	Next
	
	
End Function

Sub LoadExcelEnv() ' 엑셀 워크북 전체에서 공동으로 사용하는 환경 변수 로드

	' 자주 사용하는 시트는 전역으로 가지고 있자. (속도 향상을 위해)
	Set gWsGenDBoard 		= ThisWorkbook.Sheets(PARAMETER_SHEET_NAME)
	Set gWsDashboard 		= ThisWorkbook.Sheets(DBOARD_SHEET_NAME)
	Set gWsProject 			= ThisWorkbook.Sheets(PROJECT_SHEET_NAME)
	Set gWsActivity_Struct	= ThisWorkbook.Sheets(ACTIVITY_SHEET_NAME)

	' 엑셀 전역 환경 변수들을 가져온다.
	Dim rng 	As Range
	Set rng		= gWsGenDBoard.Range("b:c")

	gExcelEnv.SimulationDuration	= GetVariableValue(rng, "SimulTerm")
	gExcelEnv.AvgProjects 			= GetVariableValue(rng, "avgProjects")
	gExcelEnv.Hr_Init_H 			= GetVariableValue(rng, "Hr_Init_H")
	gExcelEnv.Hr_Init_M 			= GetVariableValue(rng, "Hr_Init_M")
	gExcelEnv.Hr_Init_L 			= GetVariableValue(rng, "Hr_Init_L")
	gExcelEnv.Hr_LeadTime 			= GetVariableValue(rng, "Hr_LeadTime")
	gExcelEnv.Cash_Init 			= GetVariableValue(rng, "Cash_Init")
	gExcelEnv.Problem 				= GetVariableValue(rng, "ProblemCnt")

End Sub


' 기간동안의 모든 발주 프로젝트를 미리 구해서 넣어놓는다.
Private Function CreateOrderTable() 

	Dim week 			As Integer 
	Dim projectCount	As Integer
	Dim sum 			As Integer
		
	ReDim gOrderTable(2,gExcelEnv.SimulationDuration)

	For week = 1 To gExcelEnv.SimulationDuration			
		projectCount 		= PoissonRandom(gExcelEnv.AvgProjects) ' 이번주 발생하는 프로젝트 갯수
		gOrderTable(1,week)	= sum
		gOrderTable(2,week)	= projectCount

		' 이번주 까지 발생한 프로젝트 갯수. 다음주에 기록된다. ==> 이전주까지 발생한 프로젝트 갯수후위연산. vba에서 do while 문법 모름... ㅎㅎ
		sum 	= sum + projectCount			
	Next

	gTotalProjectNum = sum
	gTableInitialized = 1
	
End Function

Private Function CreateProjects()

	Dim week			As Integer
	Dim id 				As Integer
	Dim startPrjNum		As Integer
	Dim endPrjNum		As Integer
	Dim preTotal		As Integer		
	Dim tempPrj 		As clsProject	

	If gTotalProjectNum <= 0 Then
		MsgBox "gTotalProjectNum is 0", vbExclamation 		
		Exit Function
	End If

	'프로젝트들을 생성한다. 
	ReDim gProjectTable(gTotalProjectNum)

	For week = 1 to gExcelEnv.SimulationDuration
		
		preTotal 	= gOrderTable(1,week)			' 이전 기간 까지 발생한 프로젝트 누적 갯수
		startPrjNum	= preTotal + 1 					' 이번 기간 시작프로젝트 번호
		endPrjNum 	= gOrderTable(2,week) + preTotal	' 이번 기간 마지막 프로젝트 번호
		
		If startPrjNum = 0 Then
			GoTo Continue 
		End If

		If startPrjNum > endPrjNum Then
			GoTo Continue 
		End If	

		' 이번 주에 발생한 프로젝트들을 생성한다.
		For id = startPrjNum to endPrjNum ' 
			Set tempPrj 	= New clsProject
			Call tempPrj.Init(P_TYPE_EXTERNAL, id, week) 
			Set gProjectTable(id) = tempPrj
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

Sub PrintArrayWithLine(ws As Worksheet, startRow As Long, startCol As Long, dataArray As Variant)

    Dim startRange As Range
    Dim endRange As Range
    Dim numRows As Long
    Dim numCols As Long
    Dim i As Long
    
    Set startRange = ws.Cells(startRow, startCol) ' 시작 셀 설정
    
    ' 배열의 차원 확인
    Dim dimensions As Integer
    dimensions = GetArrayDimensions(dataArray)
    
    If dimensions = 1 Then
        ' 1차원 배열 처리
        numRows = UBound(dataArray) - LBound(dataArray) + 1
        numCols = 1 ' 1차원 배열이므로 열의 수는 1
        
        Set endRange = startRange.Resize(numRows, numCols) ' 출력할 범위 설정
        
        ' 1차원 배열을 2차원 범위에 출력
        For i = 1 To numRows
            endRange.Cells(i, 1).Value = dataArray(i)
        Next i
        
    ElseIf dimensions = 2 Then
        ' 2차원 배열 처리
        numRows = UBound(dataArray, 1) - LBound(dataArray, 1) + 1
        numCols = UBound(dataArray, 2) - LBound(dataArray, 2) + 1
        
        Set endRange = startRange.Resize(numRows, numCols) ' 출력할 범위 설정
        endRange.Value = dataArray ' 배열을 시트에 출력
    End If
    
    ' 테두리 그리기
    With endRange.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With

End Sub

' 배열의 차원을 구하는 함수
Function GetArrayDimensions(arr As Variant) As Integer

    Dim dimCount As Integer
    Dim currentDim As Integer
    
    On Error GoTo ErrHandler
    dimCount = 0
    currentDim = 0
    
    Do While True
        currentDim = currentDim + 1
        ' 배열의 각 차원을 확인
        Dim temp As Long
        temp = LBound(arr, currentDim)
        dimCount = currentDim
    Loop
    
ErrHandler:
    If Err.Number <> 0 Then
        GetArrayDimensions = dimCount
    End If
    On Error GoTo 0
End Function




Function PrintProjectHeader()

	Call ClearSheet(gWsProject)			'시트의 모든 내용을 지우고 셀 병합 해제

	Dim arrHeader As Variant
    Dim strHeader As String

	' 첫 번째 줄 헤더
    strHeader = "타입,순번,발주일,시작가능,기간,시작,수익,경험,성공%,지급횟수,CF1%,CF2%,CF3%,선금,중도금,잔금"
    arrHeader = Split(strHeader, ",")
    arrHeader = ConvertToBase1(arrHeader)
	arrHeader = ConvertTo1xN(arrHeader)
	Call PrintArrayWithLine(gWsProject, 2, 1,arrHeader)

    
    ' 두 번째 줄 헤더    
    strHeader = ",Dur,start,end,HR_H,HR_M,HR_L,,,,mon_cf1,mon_cf2,mon_cf3"
	arrHeader = Split(strHeader, ",")
    arrHeader = ConvertToBase1(arrHeader)
	arrHeader = ConvertTo1xN(arrHeader)
	Call PrintArrayWithLine(gWsProject, 3, 1,arrHeader)
	
End Function




Function PrintProjectAll()

	Dim temp As clsProject
	Dim i As Integer

	For i = 1 To gTotalProjectNum
		Set temp = gProjectTable(i)
		Call temp.PrintInfo()
	Next
	
End Function


' 0 기반 배열을 1 기반 배열로 변환하는 함수
Function ConvertToBase1(arr As Variant) As Variant
    Dim i As Integer
    Dim newArr() As Variant
    
    ReDim newArr(1 To UBound(arr) - LBound(arr) + 1)
    For i = LBound(arr) To UBound(arr)
        newArr(i - LBound(arr) + 1) = arr(i)
    Next i
    
    ConvertToBase1 = newArr
End Function

Function ConvertTo1xN(arr As Variant) As Variant
    Dim i As Integer
    Dim newArr() As Variant
    Dim numCols As Integer
    
    numCols = UBound(arr) - LBound(arr) + 1
    ReDim newArr(1 To 1, 1 To numCols)
    
    For i = LBound(arr) To UBound(arr)
        newArr(1, i - LBound(arr) + 1) = arr(i)
    Next i
    
    ConvertTo1xN = newArr
End Function

Function PrintDashboard()	' 생성된 대시보드를 시트에 출력한다

	On Error GoTo ErrorHandler

	Call ClearSheet(gWsDashboard)			'시트의 모든 내용을 지우고 셀 병합 해제

	Dim arrHeader As Variant
    arrHeader = Array("월", "누계", "발주")

	Call PrintArrayWithLine(gWsDashboard, 2, 1,arrHeader)		' 세로항목을 적고
	Call PrintArrayWithLine(gWsDashboard, 2, 2,gPrintDurationTable)	'기간을 적고	
	Call PrintArrayWithLine(gWsDashboard, 3, 2,gOrderTable)		' 내용을 적는다.

	' Set myArray = GetPrintHeaderTable
	' PrintArrayWithLine(ws, 1, 1,myArray)	

	Exit Function

	' Set myArray = GetProjectInfoTable
	' PrintArrayWithLine(DBOARD_SHEET_NAME, 2, 2,myArray)	

	ErrorHandler:
		Call HandleError("PrintDashboard", Err.Description)

End Function


Function ClearSheet(ws As Worksheet)

	With ws
		Dim endRow As Long ' 마지막행
        Dim endCol As Long ' 마지막열
        endRow = .UsedRange.Rows.Count + .UsedRange.Row - 1
        endCol = .UsedRange.Columns.Count + .UsedRange.Column - 1

        ' 엑셀 파일의 셀들을 정리한다.
        .Range(.Cells(1, 1), .Cells(endRow, endCol)).UnMerge
        .Range(.Cells(1, 1), .Cells(endRow, endCol)).Clear
        .Range(.Cells(1,1),.Cells(endRow,endCol)).ClearContents

	End With
	
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



' On Error GoTo ErrorHandler
' ErrorHandler:
'     Call HandleError("ExampleFunction", Err.Description)


Sub HandleError(funcName As String, errMsg As String)
    MsgBox "Error in Sub " & funcName & ": " & errMsg, vbExclamation
End Sub