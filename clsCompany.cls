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

Private comDoingTable()			As Integer		' 진행중인 프로젝트들의 ID를 관리하는 테이블
Private comDoneTable()			As Integer		' 완료된 프로젝트들의 ID를 관리하는 테이블
Private comDefferTable()		As Integer		' 보류된 프로젝트들의 ID를 관리하는 테이블

Public comDoingTableSize As Integer		' comDoneTable의 현재크기 song ==> 동적으로 수정하자.
Public comDoneTableSize As Integer		' comDoneTable의 현재크기 song ==> 동적으로 수정하자.
Public comDefferTableSize As Integer		' comDoneTable의 현재크기 song ==> 동적으로 수정하자.

Public comID As Integer ' 회사 이름

' Public Property Get PropertyDoingTable()  As Variant
'     PropertyDoingTable = comDoingTable
' End Property

' 생성자와 소멸자
Private Sub Class_Initialize()

	comDoingTableSize	= 10
	comDoneTableSize	= 10
	comDefferTableSize	= 10

	ReDim comDoingTable(1 To comDoingTableSize)		' 진행중인 프로젝트들의 ID를 관리하는 테이블
	ReDim comDoneTable(1 To comDoneTableSize)		' 완료된 프로젝트들의 ID를 관리하는 테이블
	ReDim comDefferTable(1 To comDefferTableSize)		' 보류된 프로젝트들의 ID를 관리하는 테이블

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
	' 	
	
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
	
	cnt = OrderTable(2,week) 
	If cnt =0 Then	' 이번주에 발주된 프로젝트가 없다.
		Exit Function
	End If

	startProjectNum	= gOrderTable(1,week) + 1   ' 이번기간에 처음 프로젝트
	endProjectNum 	= gOrderTable(1,week) +  cnt' 이번기간의 마지막 프로젝트

	' song ==> 일단 모든 프로젝트를 다 하는것으로 하자.
	For id = startProjectNum to endProjectNum

		Set cProject = ProjectInfoTable(id)
		
		If cProject.ProjectType = P_TYPE_INTERNAL	Then ' 내부프로젝트이면

		' song ==> 이번주에 시작할지 말지 결정하는 루틴이 필요함. 우선은 모두 진행으로
		Else	
			cProject.StartDate	 = week
			Call AddToTable(DoingTable,week, cProject.ProjectNum)' Doing table에 추가			
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
	
	cnt = DoingTable(1,week-1) ' 지난주에 진행 중이던 프로젝트의 갯수
	If cnt =0 Then	' song ==> 지난주에 진행중이던 프로젝트가 없다.
		Exit Function
	End If

	' 지난 기간의 정보를 이번기간에 복사하고 복사된 정보에서 지난 기간에 완료된 프로젝트는 삭제한다.
	Call CopyLastToCurrent(DoingTable, week, cnt)

	For i = 1 to cnt
		id = DoingTable(i + 1, week)
		Set cProject = ProjectInfoTable(id)
		
		If project.ProjectType = P_TYPE_INTERNAL	Then ' 내부프로젝트이면
			Debug.Print "내부 프로젝트"
		Else
			If week = (project.StartDate + project.ProjectDuration-1) Then ' 지난번에 끝났으면				
				Call RemoveFromTable(DoingTable,week, project.ProjectNum)' doing table에서 삭제				
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
Public Function Init()	

	  
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
    sumRow = 1

    ' 비어있는 마지막 행 찾기 (합계를 통해 결정)
    lastEmptyRow = prjTable(sumRow, week) + 2
    If lastEmptyRow > (sumRow + m_MaxIDsPerWeek) Then
        MsgBox "No empty slot available in week " & week, vbExclamation
        Exit Sub
    End If

    ' 프로젝트 ID 추가
    prjTable(lastEmptyRow, week) = projectID

    ' 합계 및 프로젝트 수 증가
    prjTable(sumRow, week) = prjTable(sumRow, week) + 1
 
End Sub

Sub RemoveFromTable(prjTable As Variant, ByVal week As Integer, ByVal projectID As Integer)

    Dim i 		As Integer
    Dim sumRow 	As Integer
    Dim found 	As Boolean
    Dim lastRow As Integer

    ' 배열의 크기 확인
    If UBound(prjTable, 2) < week Or week < 1 Or week > m_SimulationDuration Then
        MsgBox "Week is out of range", vbExclamation
        Exit Sub
    End If

    ' 합계 행 설정
    sumRow = 1
    found = False
    lastRow = sumRow + m_MaxIDsPerWeek

    ' 해당 주에서 projectID 찾기
    For i = (sumRow + 1) To lastRow
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
        
    Else
        MsgBox "Project ID " & projectID & " not found in week " & week, vbExclamation
    End If
End Sub
