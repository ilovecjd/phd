Option Explicit
Option Base 1

' 대시보드를 생성한다. 
Private Sub GenDBoard_Click()

	ExcelInitialized = 0	' 새로운 프로젝트들을 생성하기 위해 초기화 플래그 변경
	TableInitialized = 0	' 새로운 프로젝트들을 생성하기 위해 초기화 플래그 변경

    Call Prologue()          ' 전체 파라메터 로드, 대시보드를 생성, 프로젝트 생성
	Call PrintDashboard()	' 생성된 대시보드를 시트에 출력한다
	Call PrintProject()		' 생성된 프로젝트를 시트에 출력한다.

	' Call CreateDashboard()     ' 대시보드를 생성하고 전체 프로젝트의 갯수를 구한다.
	' Call Epilogue()

End Sub

Function PrintProject()

	Dim ws 			As Worksheet
	
	On Error GoTo ErrorHandler

	Set ws 			= ThisWorkbook.Sheets(PROJECT_SHEET_NAME)

	Call ClearSheet(ws)			'시트의 모든 내용을 지우고 셀 병합 해제


	Dim arrHeader As Variant
    Dim strHeader As String

	' 첫 번째 줄 헤더
    strHeader = "타입,순번,발주일,시작가능,기간,시작,수익,경험,성공%,CF1%,CF2%,CF3%"
    arrHeader = Split(strHeader, ",")
    arrHeader = ConvertToBase1(arrHeader)
	arrHeader = ConvertTo1xN(arrHeader)
	Call PrintArrayWithLine(ws, 2, 1,arrHeader)
    
    ' 두 번째 줄 헤더    
    strHeader = ",Dur,start,end,HR_H,HR_M,HR_L,,,mon_cf1,mon_cf2,mon_cf3"
	arrHeader = Split(strHeader, ",")
    arrHeader = ConvertToBase1(arrHeader)
	arrHeader = ConvertTo1xN(arrHeader)
	Call PrintArrayWithLine(ws, 3, 1,arrHeader)


	' Call PrintArrayWithLine(ws, 2, 2,PrintDurationTable)	'기간을 적고
	
	' Call PrintArrayWithLine(ws, 3, 2,GetOrderTable)		' 내용을 적는다.

	' Set myArray = GetPrintHeaderTable
	' PrintArrayWithLine(ws, 1, 1,myArray)	

	Exit Function

	' Set myArray = GetProjectInfoTable
	' PrintArrayWithLine(DBOARD_SHEET_NAME, 2, 2,myArray)	

ErrorHandler:
	MsgBox "An error occurred: " & Err.Description, vbExclamation
	
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

	Dim ws 			As Worksheet
	
	On Error GoTo ErrorHandler

	Set ws 			= ThisWorkbook.Sheets(DBOARD_SHEET_NAME)

	Call ClearSheet(ws)			'시트의 모든 내용을 지우고 셀 병합 해제

	Dim arrHeader As Variant
    arrHeader = Array("월", "누계", "발주")
	Call PrintArrayWithLine(ws, 2, 1,arrHeader)		' 세로항목을 적고

	Call PrintArrayWithLine(ws, 2, 2,PrintDurationTable)	'기간을 적고
	
	Call PrintArrayWithLine(ws, 3, 2,GetOrderTable)		' 내용을 적는다.

	' Set myArray = GetPrintHeaderTable
	' PrintArrayWithLine(ws, 1, 1,myArray)	

	Exit Function

	' Set myArray = GetProjectInfoTable
	' PrintArrayWithLine(DBOARD_SHEET_NAME, 2, 2,myArray)	

	ErrorHandler:
		MsgBox "An error occurred: " & Err.Description, vbExclamation

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


