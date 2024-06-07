Option Explicit
Option Base 1

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 프로그램 동작을 위한 기본 정보들
Dim DataColumn 	As Integer
Dim SimulTerm   As Integer  ' 시뮬레이션을 동작 시킬 기간(주)
Dim avgProjects	As Double  ' 주당 발생하는 평균 발주 프로젝트 수

Dim Hr_Init_H   As Integer  ' 최초에 보유한 고급 인력
Dim Hr_Init_M   As Integer  ' 최초에 보유한 중급 인력
Dim Hr_Init_L   As Integer  ' 최초에 보유한 초급 인력
Dim Hr_LeadTime As Integer  ' 인력 충원에 걸리는 시간

Dim Cash_Init   As Integer  ' 최초 보유 현금
Dim Problem     As Integer  ' 프로젝트 생성 개수 (= 문제의 개수) / MakePrj 함수의 인자

Private Sub GenDBoard_Click()
    
    Call Prologue          ' 파라메터를 참조해서 기본적인 데이터를 준비한다.
    Call MakeDashboard()     ' 대시보드를 생성하고 전체 프로젝트의 갯수를 구한다.

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' desc      : 프로그램 시작을 위한 기본적인 값들을 설정한다.
'' return    : none
Sub Prologue()

    Dim i As Integer

    i 			= 7           ' Global 변수가 시작 하는 행 번호
    DataColumn	= 3  ' Global 변수가 시작 하는 열 번호
                
    SimulTerm 	= GetDataAndSetNext(i, DataColumn)
    avgProjects = GetDataAndSetNext(i, DataColumn)
   
    Hr_Init_H 	= GetDataAndSetNext(i, DataColumn)
    Hr_Init_M 	= GetDataAndSetNext(i, DataColumn)
    Hr_Init_L 	= GetDataAndSetNext(i, DataColumn)
    Hr_LeadTime = GetDataAndSetNext(i, DataColumn)
    Cash_Init 	= GetDataAndSetNext(i, DataColumn)
 
    Problem 	= GetDataAndSetNext(i, DataColumn)     ' MakePrj 함수의 인자

End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Desc     : dashboard 시트를 생성하고 값을 기록한다.
'' @        : 없음
'' Return   : 생성해야할(기간동안 발생한) 프로젝트의 갯수
Function MakeDashboard()

    Dim 	dashArr()       As Variant
    Dim 	posY            As Integer ' row 를 증가시킬 인덱스로 사용한다.
    ReDim 	dashArr(1 To 100, 1 To SimulTerm + 1)

    Dim 	i               As Integer
    Dim 	sum             As Integer ' 발생한 모든 프로젝트 갯수의 누적치
    Dim 	prjCnt          As Integer ' 매달 발생하는 프로젝트 갯수

    posY 	= 1
    dashArr(posY, 1) = "발주 프로젝트 현황":	 posY = posY + 1
    dashArr(posY, 1) = "월":					posY = posY + 1
    dashArr(posY, 1) = "누계":					posY = posY + 1
    dashArr(posY, 1) = "발주":					posY = posY + 1

    sum 	= 0

    For i = 1 To SimulTerm
        
        Dim projectsWeek  As Integer
    
        ' 이번주 발생하는 프로젝트 갯수
        prjCnt 	= PoissonRandom(avgProjects)
        sum 	= sum + prjCnt

        posY 	= 2
        dashArr(posY, i + 1) = i:       posY = posY + 1
        dashArr(posY, i + 1) = sum:     posY = posY + 1
        dashArr(posY, i + 1) = prjCnt:  posY = posY + 1
        
    Next

    posY 				= 11
    dashArr(posY, 1) 	= "진행 프로젝트 현황":    posY = posY + 1
    For i = 1 To SimulTerm
        dashArr(12, i + 1) = i
    Next
    dashArr(posY, 1) = "월":		posY = posY + 1
    dashArr(posY, 1) = "합계":		posY = posY + 1
    dashArr(posY, 1) = "prjNum":	posY = posY + 1

    
    posY = 21
    dashArr(posY, 1) = "보류 프로젝트 현황": posY = posY + 1
    For i = 1 To SimulTerm
        dashArr(22, i + 1) = i
    Next
    dashArr(posY, 1) = "월":		posY = posY + 1
    dashArr(posY, 1) = "합계": 		posY = posY + 1
    dashArr(posY, 1) = "prjNum":	posY = posY + 1
    

    posY = 31
    dashArr(posY, 1) = "완료 프로젝트 현황": posY = posY + 1
    For i = 1 To SimulTerm
        dashArr(posY, i + 1) = i
    Next
    dashArr(posY, 1) = "월":		posY = posY + 1
    dashArr(posY, 1) = "합계":		posY = posY + 1
    dashArr(posY, 1) = "prjNum":	posY = posY + 1

    Call PrintDashboard(dashArr)

    MakeDashboard = sum

End Function

' lambda(평균 발생률)를 인자로 받아 포아송 분포를 따르는 랜덤 값을 반환합니다.
Function PoissonRandom(lambda As Double) As Integer
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


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Desc     : 셀 i,j 의 값을 리턴하고 인덱스 i 값을 다음 인덱스로 증가시킨다.
'' @i , j   : 값을 가져올 셀의 위치
'' Return   : 셀(i,j)의 값 / i (증가된  i 값)
Function GetDataAndSetNext(i As Integer, j As Integer)

    GetDataAndSetNext = Cells(i, j)
    i = i + 1

End Function


Sub Epilogue()

End Sub



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Desc     : 생성된 프로젝트 배열을 엑셀에 출력하는 함수.
'' @prjArr  : 프로젝트 정보를 담아서 엑셀에 출력할 배열
'' Return   : 없음
Function PrintDashboard(Arr() As Variant)

    Dim sheetPrj        As Worksheet
    Set sheetPrj = Worksheets("dashboard")
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
        .Range(.Cells(1, 1), .Cells(endRow, endCol)).Value = Arr

        Dim RNG1 As Range
        ' 테두리그리기
        Set RNG1 = .Range(.Cells(2, 1), .Cells(4, endCol))
        RNG1.Borders.LineStyle = 1
        RNG1.Borders.Weight = xlThin
        RNG1.Borders.ColorIndex = 1

        ' 테두리그리기
        Set RNG1 = .Range(.Cells(12, 1), .Cells(19, endCol))
        RNG1.Borders.LineStyle = 1
        RNG1.Borders.Weight = xlThin
        RNG1.Borders.ColorIndex = 1

        ' 테두리그리기
        Set RNG1 = .Range(.Cells(22, 1), .Cells(29, endCol))
        RNG1.Borders.LineStyle = 1
        RNG1.Borders.Weight = xlThin
        RNG1.Borders.ColorIndex = 1

        ' 테두리그리기
        Set RNG1 = .Range(.Cells(32, 1), .Cells(39, endCol))
        RNG1.Borders.LineStyle = 1
        RNG1.Borders.Weight = xlThin
        RNG1.Borders.ColorIndex = 1

    End With
    
End Function


