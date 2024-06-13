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
Dim TableSizeY		As Integer		' 테이블들의 크기 ( 들어갈 프로젝트 갯수)


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


	' 1. 완료프로젝트 검사 및 진행프로젝트 업데이트
	' 지난주에 진행중인 프로젝트중 완료된 것이 있는가?
	' 완료되지 않은 프로젝트들만 이번주로 이관
	Call CheckLastWeek(week)

	' 2. 신규 프로젝트 선택 및 진행프로젝트 업데이트
	'  이번주에 발주된 프로젝트중 시작할 것이 있으면 이번주 진행 프로젝트로 기록
	Call SelectNewProject(week)

	' 
	
	
End Function

Private Function SelectNewProject(week As Integer)

	Dim project As clsProject
	Dim cnt As Integer
	Dim i  As Integer
	Dim id As Integer


	If week = 0 Then ' song ==> 현재까지 코드에서는 0 번째 주 라는 개념은 없다.
		Exit Function
	End If
	
	cnt = DoingTable(1,week) 
	If cnt =0 Then	' song ==> 지난주에 진행중이던 프로젝트가 없다.
		Exit Function
	End If

	For i = 1 to cnt
		id = DoingTable(i +1, week-1)
		project = QDoing.SearchByID(id)
		
		If project.ProjectType = P_TYPE_INTERNAL	Then ' 내부프로젝트이면

		Else
			If week = (project.StartDate + project.ProjectDuration) Then ' 지난번에 끝났으면
				Call QDoing.RemoveByID(project.ProjectNum )' q에서 꺼내고 
				Call RemoveFromTable(DoingTable,week, project.ProjectNum)' doing table에서 삭제
				Call AddToTable(DoingTable,week, project.ProjectNum)' done table에 업데이트			
			End If
		End If
	Next

End Function


Private Function CheckLastWeek(week As Integer)

	Dim project As clsProject
	Dim cnt As Integer
	Dim i  As Integer
	Dim id As Integer


	If week = 1 Then ' song ==> 현재까지 코드에서는 처음 시작에는 지난주가 없다
		Exit Function
	End If
	
	cnt = DoingTable(1,week-1) 
	If cnt =0 Then	' song ==> 지난주에 진행중이던 프로젝트가 없다.
		Exit Function
	End If

	For i = 1 to cnt
		id = DoingTable(i +1, week-1)
		project = QDoing.SearchByID(id)
		
		If project.ProjectType = P_TYPE_INTERNAL	Then ' 내부프로젝트이면

		Else
			If week = (project.StartDate + project.ProjectDuration) Then ' 지난번에 끝났으면
				Call QDoing.RemoveByID(project.ProjectNum )' q에서 꺼내고 
				Call RemoveFromTable(DoingTable,week, project.ProjectNum)' doing table에서 삭제
				Call AddToTable(DoingTable,week, project.ProjectNum)' done table에 업데이트			
			End If
		End If
	Next

End Function


' All Q init and QToDo setting, 
Public Function Init()	

	Set QOrder	= New clsQueue
	Set QDoing	= New clsQueue
	Set QDone	= New clsQueue

	TableSizeY = 6 'song ==> 추후 프로젝스 갯수에 따라 동적으로 변경 필요
	ReDim DoingTable(TableSizeY,	m_SimulationDuration)
	ReDim DoneTable(TableSizeY, 	m_SimulationDuration)
	ReDim DefferTable(TableSizeY, 	m_SimulationDuration)


	Dim tempPrj As clsProject
	    
	' 전역 변수 설정
    ' 기간별 선택 가능프로젝트 확인	
	Dim weeks 			As Integer		' 시뮬레이션 기간
	Dim i 				As Integer
	Dim j 				As Integer
	Dim endPrjNum		As Integer
	Dim preTotal		As Integer
	Dim prjNUm			As Integer	
	Dim totalPrjNum		As Integer	

	totalPrjNum   		= OrderTable(1,m_SimulationDuration)

	For i = 1 to m_SimulationDuration 
		
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
			'Init(initType As Integer, initNum As Integer, shtName As String, ODate As Integer)	
			Call tempPrj.Init(0, j, PROJECT_SHEET_NAME,i) 
			Call tempPrj.PrintInfo()
			Call QOrder.Enqueue(tempPrj)
		Next

		j = QOrder.GetCount

		Debug.Print "Q Count" & j
		Debug.Print 
		
		Continue: 

	Next
  

End Function


'' 프로젝트 관리 테이블에 프로젝트를 넣는 범용적인 함수 (Doing, Done, Deffer )
Sub AddToTable(prjTable As Variant, week As Integer, projectID As Integer)
    Dim sumRow As Integer
    Dim lastEmptyRow As Integer

    ' 배열의 크기 확인
    If UBound(prjTable, 2) < week Or week < 1 Or week > m_SimulationDuration Then
        MsgBox "Week is out of range", vbExclamation
        Exit Sub
    End If

    ' 합계 행 설정
    sumRow = 2

    ' 비어있는 마지막 행 찾기 (합계를 통해 결정)
    lastEmptyRow = prjTable(sumRow, week) + 3
    If lastEmptyRow > 2 + m_MaxIDsPerWeek Then
        MsgBox "No empty slot available in week " & week, vbExclamation
        Exit Sub
    End If

    ' 프로젝트 ID 추가
    prjTable(lastEmptyRow, week) = projectID

    ' 합계 및 프로젝트 수 증가
    prjTable(sumRow, week) = prjTable(sumRow, week) + 1
    prjTable(3, week) = prjTable(3, week) + 1
End Sub

Sub RemoveFromTable(prjTable As Variant, week As Integer, projectID As Integer)
    Dim i As Integer
    Dim sumRow As Integer
    Dim found As Boolean
    Dim lastRow As Integer

    ' 배열의 크기 확인
    If UBound(prjTable, 2) < week Or week < 1 Or week > m_SimulationDuration Then
        MsgBox "Week is out of range", vbExclamation
        Exit Sub
    End If

    ' 합계 행 설정
    sumRow = 2
    found = False
    lastRow = 2 + m_MaxIDsPerWeek

    ' 해당 주에서 projectID 찾기
    For i = 3 To lastRow
        If prjTable(i, week) = projectID Then
            found = True
            Exit For
        End If
    Next i

    If found Then
        ' ID 삭제 및 아래 행들 이동
        For i = i To lastRow - 1
            prjTable(i, week) = prjTable(i + 1, week)
        Next i
        prjTable(lastRow, week) = Empty

        ' 합계 및 프로젝트 수 감소
        prjTable(sumRow, week) = prjTable(sumRow, week) - 1
        prjTable(3, week) = prjTable(3, week) - 1
    Else
        MsgBox "Project ID " & projectID & " not found in week " & week, vbExclamation
    End If
End Sub
