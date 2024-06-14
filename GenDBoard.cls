Option Explicit
Option Base 1

' 데시보드를 생성한다. 
Private Sub GenDBoard_Click()

	ExcelInitialized = 0	' 새로운 프로젝트들을 생성하기 위해 초기화 플래그 변경
	TableInitialized = 0	' 새로운 프로젝트들을 생성하기 위해 초기화 플래그 변경

    Call Prologue()          ' 파라메터를 참조해서 기본적인 데이터를 준비한다.
    ' Call CreateDashboard()     ' 대시보드를 생성하고 전체 프로젝트의 갯수를 구한다.
	' Call Epilogue()

End Sub

Function CreateProjects()

	Dim i 				As Integer
	Dim id 				As Integer
	Dim startPrjNum		As Integer
	Dim endPrjNum		As Integer
	Dim preTotal		As Integer	
	Dim totalPrjNum		As Integer	
	Dim tempPrj 		As clsProject	
	Dim envVars			As EnvExcel
	Dim bIsInit			as Boolean


	totalPrjNum   		= OrderTable(1,gSimulationDuration) + OrderTable(2,gSimulationDuration)

	'envVars = 
	'프로젝트들을 생성한다. 
	For i = 1 to m_SimulationDuration 
		
		preTotal 	= OrderTable(1,i)			' 이전 기간 까지 발생한 프로젝트 누적 갯수
		startPrjNum	= preTotal + 1 				' 이번 기간 시작프로젝트 번호
		endPrjNum 	= OrderTable(2,i) + preTotal' 이번 기간 마지막 프로젝트 번호
		
		If startPrjNum = 0 Then
			GoTo Continue 
		End If

		If startPrjNum > endPrjNum Then
			GoTo Continue 
		End If	

		' 이번 주에 발생한 프로젝트들을 생성한다.
		For id = startPrjNum to endPrjNum ' 
			Set tempPrj = New clsProject		
			'Init(initType As Integer, initNum As Integer, shtName As String, ODate As Integer)	
			Call tempPrj.Init(P_TYPE_EXTERNAL, id, PROJECT_SHEET_NAME,i) 
			Set ProjectInfoTable(id) = tempPrj
			Call tempPrj.PrintInfo()
		Next

		Continue: 

	Next
	
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Desc     : dashboard 시트를 생성하고 값을 기록한다.
'' @        : 없음
'' Return   : 생성해야할(기간동안 발생한) 프로젝트의 갯수
Function CreateDashboard()

    Dim 	dashArr()       As Variant
    Dim 	posY            As Integer ' row 를 증가시킬 인덱스로 사용한다.
    ReDim 	dashArr(1 To 100, 1 To SimulTerm + 1)

    Dim 	i               As Integer
    Dim 	sum             As Integer ' 발생한 모든 프로젝트 갯수의 누적치
    Dim 	prjCnt          As Integer ' 매달 발생하는 프로젝트 갯수

    posY 	= 1
    dashArr(posY, 1) = ORDER_PROJECT_TITLE:	 	posY = posY + 1
    dashArr(posY, 1) = "월":					posY = posY + 1
    dashArr(posY, 1) = "누계":					posY = posY + 1
    dashArr(posY, 1) = "발주":					posY = posY + 1

    sum 	= 0

    For i = 1 To SimulTerm
        
        Dim projectsWeek  As Integer
    
        ' 이번주 발생하는 프로젝트 갯수
        prjCnt 	= PoissonRandom(gAvgProjects) 
        posY 	= 2
        dashArr(posY, i + 1) = i:       posY = posY + 1
        dashArr(posY, i + 1) = sum:     posY = posY + 1
        dashArr(posY, i + 1) = prjCnt:  posY = posY + 1

		' 이번주 까지 발생한 프로젝트 갯수. 다음주에 기록된다. ==> 이전주까지 발생한 프로젝트 갯수
		sum 	= sum + prjCnt
        
    Next
	'song ==> 발주 프로젝트 전체 크기를 기록해 두자.

    ' posY 				= 11
    ' dashArr(posY, 1) 	= "진행 프로젝트 현황":    posY = posY + 1
    ' For i = 1 To SimulTerm
    '     dashArr(12, i + 1) = i
    ' Next
    ' dashArr(posY, 1) = "월":		posY = posY + 1
    ' dashArr(posY, 1) = "합계":		posY = posY + 1
    ' dashArr(posY, 1) = "prjNum":	posY = posY + 1

    
    ' posY = 21
    ' dashArr(posY, 1) = "보류 프로젝트 현황": posY = posY + 1
    ' For i = 1 To SimulTerm
    '     dashArr(22, i + 1) = i
    ' Next
    ' dashArr(posY, 1) = "월":		posY = posY + 1
    ' dashArr(posY, 1) = "합계": 		posY = posY + 1
    ' dashArr(posY, 1) = "prjNum":	posY = posY + 1
    

    ' posY = 31
    ' dashArr(posY, 1) = "완료 프로젝트 현황": posY = posY + 1
    ' For i = 1 To SimulTerm
    '     dashArr(posY, i + 1) = i
    ' Next
    ' dashArr(posY, 1) = "월":		posY = posY + 1
    ' dashArr(posY, 1) = "합계":		posY = posY + 1
    ' dashArr(posY, 1) = "prjNum":	posY = posY + 1

    Call PrintDashboard(dashArr)

    MakeDashboard = sum

End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Desc     : 생성된 프로젝트 배열을 엑셀에 출력하는 함수.
'' @prjArr  : 프로젝트 정보를 담아서 엑셀에 출력할 배열
'' Return   : 없음
Function PrintDashboard(Arr() As Variant)

    Dim sheetPrj        As Worksheet
    Set sheetPrj = Worksheets(DBOARD_SHEET_NAME)
    With sheetPrj
        Dim endRow As Long ' 마지막행
        Dim endCol As Long ' 마지막열
        endRow = .UsedRange.Rows.Count + .UsedRange.Row - 1
        endCol = .UsedRange.Columns.Count + .UsedRange.Column - 1

        ' 엑셀 파일의 셀들을 정리한다.
        .Range(.Cells(1, 1), .Cells(endRow, endCol)).UnMerge
        .Range(.Cells(1, 1), .Cells(endRow, endCol)).Clear
        '.Range(.Cells(1,1),.Cells(endRow,endCol)).ClearContents


        '''''''''''''''''''''''''''''''''''''
        ' 배열의 크기를 구해서 배열을 넣고 테두리를 그린다.
        endRow = UBound(Arr, 1) - LBound(Arr, 1) + 1
        endCol = UBound(Arr, 2) - LBound(Arr, 2) + 1
        .Range(.Cells(1, 1), .Cells(endRow, endCol)).value = Arr

        Dim RNG1 As Range
        ' 테두리그리기
        Set RNG1 = .Range(.Cells(2, 1), .Cells(4, endCol))
        RNG1.Borders.LineStyle = 1
        RNG1.Borders.Weight = xlThin
        RNG1.Borders.ColorIndex = 1

        ' ' 테두리그리기
        ' Set RNG1 = .Range(.Cells(12, 1), .Cells(19, endCol))
        ' RNG1.Borders.LineStyle = 1
        ' RNG1.Borders.Weight = xlThin
        ' RNG1.Borders.ColorIndex = 1

        ' ' 테두리그리기
        ' Set RNG1 = .Range(.Cells(22, 1), .Cells(29, endCol))
        ' RNG1.Borders.LineStyle = 1
        ' RNG1.Borders.Weight = xlThin
        ' RNG1.Borders.ColorIndex = 1

        ' ' 테두리그리기
        ' Set RNG1 = .Range(.Cells(32, 1), .Cells(39, endCol))
        ' RNG1.Borders.LineStyle = 1
        ' RNG1.Borders.Weight = xlThin
        ' RNG1.Borders.ColorIndex = 1

    End With
    
End Function


