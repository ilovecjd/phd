VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCompany"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit
Option Base 1

Dim QOrder			As clsQueue		' 발주된 프로젝트들
Dim QDoing 			As clsQueue		' 진행중인 프로젝트들
Dim QDone			As clsQueue		' 완료된 프로젝트들
Dim QDeffer			As clsQueue		' 보류된 프로젝트들

Dim OrderTable()	As Variant 		' 발주된 프로젝트들을 관리하는 테이블
Dim DoingTable()	As Integer		' 진행중인 프로젝트들을 관리하는 테이블
Dim DoneTable()		As Integer		' 완료된 프로젝트들을 관리하는 테이블
Dim DefferTable()	As Integer		' 보류된 프로젝트들을 관리하는 테이블


Private m_SimulationDuration As Integer ' 시뮬레이션을 수행하는 기간
Property Let SimulationDuration(value As Long)  
	m_SimulationDuration = value
End Property

' OrderTable의 Range 전달받아서 복사한다.
Public Function SetOrderTable(rng As Range)

	If m_SimulationDuration = 0 Then
		MsgBox "m_SimulationDuration Error " , " m_SimulationDuration Error", vbExclamation
		Exit Function
	End If

	ReDim OrderTable(2, m_SimulationDuration)

	OrderTable = rng.value

End Function 


'' 이번 기간에 결정할 일들. 프로젝트의 신규진행, 멈춤, 인원증감 결정
Public Function Decision(week As Integer)

	Dim projects As Collection

	' 완료프로젝트 검사	

	' 진행프로젝트 업데이트
	' 신규 프로젝트 선택
	
	
End Function

' All Q init and QToDo setting, 
Public Function Init()

	Set QOrder	= New clsQueue
	Set QDoing	= New clsQueue
	Set QDone	= New clsQueue

	Dim tempPrj As clsProject
	    
	' 전역 변수 설정
    ' 기간별 선택 가능프로젝트 확인
	Dim ws 				As Worksheet	'프로젝트 정보를 적을 시트
	Dim weeks 			As Integer		' 시뮬레이션 기간
	Dim i 				As Integer
	Dim j 				As Integer
	Dim endPrjNum		As Integer
	Dim preTotal		As Integer
	Dim prjNUm			As Integer	
	Dim totalPrjNum		As Integer	

	totalPrjNum   		= OrderTable(1,m_SimulationDuration)

	For i = 1 to (m_SimulationDuration) 
		weeks 		= i 'ws.Cells(2,i ) ' 주
		endPrjNum 	= OrderTable(1,i)' ws.Cells(3,i ) ' 이번달까지 누계는 마지막 번호

		If i = 1 Then ' 처음이면 초기값을
			preTotal 	= 0' 이전누계	
		Else
			preTotal 	= OrderTable(1,i-1)' ws.Cells(3,i-1 )' 이전누계				
		End If

		prjNUm = preTotal + 1 ' 시작프로젝트 번호

		If prjNUm = 0 Then
			GoTo Continue 
		End If

		If prjNUm > endPrjNum Then
			GoTo Continue 
		End If	

		For j = prjNUm to endPrjNum ' 
			Set tempPrj = New clsProject			
			Call tempPrj.Init(0, j, "project",weeks) 
			Call tempPrj.PrintInfo()
			Call QOrder.Enqueue(tempPrj)
		Next

		j = QOrder.GetCount

		Debug.Print "Q Count" & j
		Debug.Print 
		
		Continue: 

	Next
  

End Function

Public Function LoadTable()
	
End Function