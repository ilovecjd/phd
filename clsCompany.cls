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

Private comDoingTable()     As Integer  ' 진행중인 프로젝트들의 ID를 관리하는 테이블
Private comDoneTable()      As Integer  ' 완료된 프로젝트들의 ID를 관리하는 테이블
Private comDefferTable()    As Integer  ' 보류된 프로젝트들의 ID를 관리하는 테이블



Public comDoingTableSize As Integer		' comDoneTable의 현재크기 song ==> 동적으로 수정하자.
Public comDoneTableSize As Integer		' comDoneTable의 현재크기 song ==> 동적으로 수정하자.
Public comDefferTableSize As Integer		' comDoneTable의 현재크기 song ==> 동적으로 수정하자.

Private comProjectTable() As clsProject

Public comID As Integer ' 회사 이름

' Public Property Get PropertyDoingTable()  As Variant
'     PropertyDoingTable = comDoingTable
' End Property

Public Property Get PropertyDoingTable() As Integer()
    PropertyDoingTable = comDoingTable
End Property

Public Property Let PropertyDoingTable(Value() As Integer)
    comDoingTable = Value
End Property

Public Property Get PropertyDoneTable() As Integer()
    PropertyDoneTable = comDoneTable
End Property

Public Property Let PropertyDoneTable(Value() As Integer)
    comDoneTable = Value
End Property

Public Property Get PropertyDefferTable() As Integer()
    PropertyDefferTable = comDefferTable
End Property

Public Property Let PropertyDefferTable(Value() As Integer)
    comDefferTable = Value
End Property



' 생성자와 소멸자
Private Sub Class_Initialize()

	comDoingTableSize	= 10
	comDoneTableSize	= 10
	comDefferTableSize	= 10

	ReDim comDoingTable(1 To comDoingTableSize, 1 To gExcelEnv.SimulationDuration)		' 진행중인 프로젝트들의 ID를 관리하는 테이블
	ReDim comDoneTable(1 To comDoneTableSize, 1 To gExcelEnv.SimulationDuration)		' 완료된 프로젝트들의 ID를 관리하는 테이블
	ReDim comDefferTable(1 To comDefferTableSize, 1 To gExcelEnv.SimulationDuration)	' 보류된 프로젝트들의 ID를 관리하는 테이블

	comProjectTable = gProjectTable

End Sub

Private Sub Class_Terminate()
	' 종료할때 버퍼의 크기를 확인해서 가까운 값으로 초기값을 결정하자.
	Debug.Print "Class is being terminated" comDoingTableSize, comDoneTableSize,comDefferTableSize
End Sub



'' 이번 기간에 결정할 일들. 프로젝트의 신규진행, 멈춤, 인원증감 결정
Public Function Decision(week As Integer)

	' 1. 완료프로젝트 검사 및 진행프로젝트 업데이트
	' 지난주에 진행중인 프로젝트중 완료된 것이 있는가?
	' 완료되지 않은 프로젝트들만 이번주로 이관
	Call CheckLastWeek(week)

	' 2. 신규 프로젝트 선택 및 진행프로젝트 업데이트
	'  이번주에 발주된 프로젝트중 시작할 것이 있으면 이번주 진행 프로젝트로 기록
	Call SelectNewProject(week)

	' Call comPrintDashboard()
	' 	
	
End Function

Private Function comPrintDashboard()
	
	On Error GoTo ErrorHandler

	'Call ClearSheet(gWsDashboard)			'시트의 모든 내용을 지우고 셀 병합 해제

	Dim startRow 	As Long
	Dim arrHeader 	As Variant
    arrHeader = Array("월", "누계", "prjNum")

	startRow = 6
	Call PrintArrayWithLine(gWsDashboard, startRow + 1, 1,arrHeader)		' 세로항목을 적고
	Call PrintArrayWithLine(gWsDashboard, startRow + 1, 2,gPrintDurationTable)	'기간을 적고	
	Call PrintArrayWithLine(gWsDashboard, startRow + 2, 2,comDoingTable)		' 내용을 적는다.

	startRow = startRow + comDoingTableSize + 2
	Call PrintArrayWithLine(gWsDashboard, startRow + 1, 1,arrHeader)		' 세로항목을 적고
	Call PrintArrayWithLine(gWsDashboard, startRow + 1, 2,gPrintDurationTable)	'기간을 적고	
	Call PrintArrayWithLine(gWsDashboard, startRow + 2, 2,comDoneTable)		' 내용을 적는다.

	startRow = startRow + comDoneTableSize + 2
	Call PrintArrayWithLine(gWsDashboard, startRow + 1, 1,arrHeader)		' 세로항목을 적고
	Call PrintArrayWithLine(gWsDashboard, startRow + 1, 2,gPrintDurationTable)	'기간을 적고	
	Call PrintArrayWithLine(gWsDashboard, startRow + 2, 2,comDefferTable)		' 내용을 적는다.


	Exit Function

ErrorHandler:
	Call HandleError("comPrintDashboard", Err.Description)
	
End Function

Private Function SelectNewProject(week As Integer)

	Dim cProject			As clsProject
	Dim cnt 			As Integer	
	Dim id 				As Integer

	Dim startProjectNum As Integer
	Dim endProjectNum 	As Integer


	If week = 0 Then ' song ==> 현재까지 코드에서는 0 번째 주 라는 개념은 없다.
		Exit Function
	End If
	
	cnt = gOrderTable(2,week) 
	If cnt =0 Then	' 이번주에 발주된 프로젝트가 없다.
		Exit Function
	End If

	startProjectNum	= gOrderTable(1,week) + 1   ' 이번기간에 처음 프로젝트
	endProjectNum 	= gOrderTable(1,week) +  cnt' 이번기간의 마지막 프로젝트

	' song ==> 일단 모든 프로젝트를 다 하는것으로 하자.
	For id = startProjectNum to endProjectNum

		Set cProject = comProjectTable(id)
		
		If cProject.ProjectType = P_TYPE_INTERNAL	Then ' 내부프로젝트이면

		' song ==> 이번주에 시작할지 말지 결정하는 루틴이 필요함. 우선은 모두 진행으로
		Else	
			cProject.StartDate	 = week
			Call AddToTable(comDoingTable,comDoingTableSize, week, cProject.ProjectNum)' Doing table에 추가			
		End If
	Next

End Function

' 지난주에 완료된 프로젝트들이 있는지 확인한다.
Private Function CheckLastWeek(week As Integer)

	Dim cProject As clsProject
	Dim cnt 	As Integer
	Dim i  		As Integer
	Dim id 		As Integer


	If week = 1 Then ' song ==> 현재까지 코드에서는 처음 시작에는 지난주가 없다
		Exit Function
	End If
	
	cnt = comDoingTable(1,week-1) ' 지난주에 진행 중이던 프로젝트의 갯수
	If cnt =0 Then	' song ==> 지난주에 진행중이던 프로젝트가 없다.
		Exit Function
	End If

	' 지난 기간의 정보를 이번기간에 복사하고 복사된 정보에서 지난 기간에 완료된 프로젝트는 삭제한다.
	Call CopyLastToCurrent(comDoingTable, week, cnt)

	For i = 1 to cnt

		id = comDoingTable(i + 1, week)

		If id = 0 Then
			Exit For
		End If

		Set cProject = comProjectTable(id)
		
		If cProject.ProjectType = P_TYPE_INTERNAL	Then ' 내부프로젝트이면
			Debug.Print "내부 프로젝트"
		Else
			If week = (cProject.StartDate + cProject.ProjectDuration-1) Then ' 지난번에 끝났으면				
				Call RemoveFromTable(comDoingTable,comDoingTableSize, week, cProject.ProjectNum)' doing table에서 삭제
				i = i - 1 ' 이번 행을 삭제하고 아래것을 가져왔으므로 이번 행부터 이어서 검사해야 한다. 
			End If
		End If
	Next

End Function

' 복사할 테이블, 현재기간, 이전기간의 프로젝트 갯수
Private Function CopyLastToCurrent(prjTable As Variant, ByVal currentWeek As Integer, ByVal lastProjectCnt As Integer)

	' song ==> 예외 처리 필요
	Dim i As Integer
	
	For i = 1 To (lastProjectCnt + 1)
		prjTable(i, currentWeek) = prjTable(i, currentWeek-1)
	Next
	
End Function

' All Q init and QToDo setting, 
Public Sub Init()	

	Debug.Print "아직은 비어 있는 함수"

End Sub

Public Function Resize2DArray(arr As Variant, newRows As Integer, newCols As Integer)

    Dim oldRows As Integer
    Dim oldCols As Integer
    Dim i As Integer
    Dim j As Integer
    
    oldRows = UBound(arr, 1)
    oldCols = UBound(arr, 2)
    
    Dim tempArray() As Variant
    ReDim tempArray(1 To newRows, 1 To newCols)
    
    For i = 1 To oldRows
        For j = 1 To oldCols
            tempArray(i, j) = arr(i, j)
        Next j
    Next i
    
	ReDim arr(1 To newRows, 1 To newCols)

	  For i = 1 To oldRows
        For j = 1 To oldCols
            arr(i, j) = tempArray(i, j)
        Next j
    Next i
    
End Function


Public Function AddToTable(prjTable As Variant, tableSize As Integer, week As Integer, projectID As Integer) As Integer

    Dim currentRow As Integer
    
    ' 현재 주(week)의 첫 번째 행 값(값이 들어 있는 행의 갯수)을 가져옴
    currentRow = prjTable(1, week) + 2
    
    ' 만약 현재 행의 값이 tableSize를 초과하면 배열을 다시 할당 
	' currentRow는 ID 갯수 , tableSize는 currentRow + sum 을 기록하는 추가 행
    If currentRow = (tableSize-1) Then
        tableSize = tableSize + 1
        Call Resize2DArray(prjTable, tableSize,  gExcelEnv.SimulationDuration)
    End If
    
    ' 프로젝트 ID를 prjTable의 적절한 위치에 추가
    prjTable(currentRow, week) = projectID
    
    ' 첫 번째 행의 값을 1 증가시켜 행의 갯수를 업데이트
    prjTable(1, week) = prjTable(1, week) + 1
    
    ' 최종 변경된 행의 크기를 반환
    AddToTable = tableSize
End Function

' '' 프로젝트 관리 테이블에 프로젝트를 넣는 범용적인 함수 (Doing, Done, Deffer )
' Sub AddToTable(prjTable As Variant, week As Integer, projectID As Integer)

'     Dim sumRow As Integer
'     Dim lastEmptyRow As Integer

'     ' 배열의 크기 확인
'     If UBound(prjTable, 2) < week Or week < 1 Or week > m_SimulationDuration Then
'         MsgBox "Week is out of range", vbExclamation
'         Exit Sub
'     End If

'     ' 합계 행 설정
'     sumRow = 1

'     ' 비어있는 마지막 행 찾기 (합계를 통해 결정)
'     lastEmptyRow = prjTable(sumRow, week) + 2
'     If lastEmptyRow > (sumRow + m_MaxIDsPerWeek) Then
'         MsgBox "No empty slot available in week " & week, vbExclamation
'         Exit Sub
'     End If

'     ' 프로젝트 ID 추가
'     prjTable(lastEmptyRow, week) = projectID

'     ' 합계 및 프로젝트 수 증가
'     prjTable(sumRow, week) = prjTable(sumRow, week) + 1
 
' End Sub

Sub RemoveFromTable(prjTable As Variant, tableSize As Integer, ByVal week As Integer, ByVal projectID As Integer)

    Dim i 		As Integer
    Dim sumRow 	As Integer ' sum 이 들어있는 위치
	Dim idCount	As Integer	' id들의 갯수
    Dim found 	As Boolean
    	
    sumRow = 1
    found = False
	idCount = prjTable(sumRow,week)

    ' 배열의 크기 확인
    If UBound(prjTable, 2) < week Or week < 1 Or week > gExcelEnv.SimulationDuration Then
        MsgBox "Week is out of range", vbExclamation
        Exit Sub
    End If
	' song ==> tableSize의 크기도 검증하자.
	
    ' 해당 주에서 projectID 찾기
    For i = (sumRow + 1) To (sumRow + idCount)
        If prjTable(i, week) = projectID Then
            found = True
            Exit For
        End If
    Next i

	' ID를 발견하면 삭제 및 아래 행들 이동
    If found Then

        For i = i To tableSize

			If i = tableSize Then 'ID가 마지막 행에 들어 있으면 
				prjTable(i, week) = Empty
			Else
				prjTable(i, week) = prjTable(i + 1, week)
			End If

        Next i

        ' 합계 및 프로젝트 수 감소
        prjTable(sumRow, week) = prjTable(sumRow, week) - 1
        
    Else
        MsgBox "Project ID " & projectID & " not found in week " & week, vbExclamation
    End If

End Sub
