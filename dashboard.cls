Option Explicit
Option Base 1

'' 시트내의 값을 이용해서 프로젝트를 가상을 진행 시켜본다.
Private Sub CommandButton1_Click()

    ' 전역 변수 설정
    ' 월별 선택 가능프로젝트 확인
    'GenerateCombinations
    'CirculateMonth ' 월별 진행 프로젝트 현황 표를 생성한다.
    'Call BuildPrj(10)

	Dim i 				As Integer	
	Dim Company As clsCompany

	Call Prologue()

	Dim ws 		As Worksheet
	Dim rng 	As Range	
	Set ws		= ThisWorkbook.Sheets(DBOARD_SHEET_NAME)
	Set rng		= ws.Range(ws.Cells(3,2),ws.Cells(4,SimulTerm+1))


	Set Company = New clsCompany
	' song ==> 회사 ID를 넣어 주자. (같은 조건에서 여러 회사를 운영)
	Company.SimulationDuration = SimulTerm
	Call Company.SetOrderTable(rng)	' 
	Call Company.Init()	' 

	For i = 1 to 10 'song ==> SimulTerm 으로 수정
		Call Company.Decision(i)	' i번째 기간에 결정해야 할 일들
	Next

	Call Epilogue()
	
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' 프로젝트의 기본 정보를 생성한다.
'' @prjArr : 프로젝트 정보를 담아서 엑셀에 출력할 배열
'' @prjNum : 프로젝트 번호
Function buildPrjInfo(prjArr() As Variant, prjNum As Integer)
   
    ' 첫 행을 적어준다. 
    Dim activityCnt As Integer    
    Dim iRow As Integer ' 프로젝트가 위치할 배열의 첫 행 번호

    iRow = (prjNum -1)* SIZE_OF_PROJECT + 2'song 예외 처리 필요 (iRow가 1(제목열) 보다 작은경우)

    prjArr(iRow,1) = "순번"
    prjArr(iRow,2) = prjNum

    prjArr(iRow+1,1) = "기간" 
    prjArr(iRow+1,2) = Int(Rnd() * 100 Mod 12) + 1 ' 기간은 12개월 이내

    prjArr(iRow+2,1) = "ActivityCnt" 
    activityCnt = Int(Rnd() * 100 Mod 6) + 1 ' song 비율은 추후 수정
    prjArr(iRow+2,2) = activityCnt

End Function

Function BuildPrj(prjCnt As Integer)

    Dim project As typeProject
    Dim arrPrj() As typeProject
    ReDim rrPrj(1 to prjCnt)

    Dim i As Integer
    For i = 1 to prjCnt
        project = arrPrj(i)
        project.num = i
        period  = 
        orderDate
        possibleDate
        startDate
        profit
        experience
        successPercentage
        CF
        N_CF
        
    Next
    
End Function


Sub GenerateCombinations()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("dashboard") ' 적절한 시트 이름으로 변경
    
    Dim lastCol         As Integer
    lastCol = ws.Cells(2, ws.Columns.Count).End(xlToLeft).Column ' 마지막 열 찾기
    
    Dim periods()       As Integer
    Dim accumulate()    As Integer ' 누적된 발주프로젝의 갯수를 넣어 둔다.
    Dim counts()        As Integer
    Dim i               As Integer
    
    ' 기간과 갯수를 배열에 저장
    ReDim periods(1 To lastCol) 'song lastCol은 전역 변수에서 가져오는걸로 수정하자
    ReDim accumulate(1 To lastCol) 
    ReDim counts(1 To lastCol) 
    
    ' 데이터는 2번째 열부터 읽어옴
    For i = 1 To lastCol - 1 ' 헤더 부분은 빼고 읽어 온다
        periods(i)      = ws.Cells(2, i+1).value
        accumulate(i)   = ws.Cells(3, i+1).value
        counts(i)       = ws.Cells(4, i+1).value
    Next i
    
    ' 모든 조합 생성  
    Dim ePrjNum As Integer ' 마지막 프로젝트 번호  
    ePrjNum = accumulate(1)

    Dim result()          As Variant
    ReDim result(1 To lastCol)
    For i = 1 to ePrjNum
        ' 1번컬럼은 외부 루프에서 돌고 2번째 컬럼부터 본다
        result(1) = i
        GenerateCombinationRecursive periods, accumulate, counts, result, 2, lastCol, i 
    Next
End Sub
    
    
' 1st loop : 1 to count(col) ==> 외부에서 정리하고 재귀호출 함수는 2번 col 부터 
' other loop : acc(col) to acc(col) + (count(col)-1) ==> 프로젝트 번호
'              
' exit condition : count(col) = prjNum    
Sub GenerateCombinationRecursive(periods()      As Integer, _
                                accumulate()    As Integer, _
                                counts()        As Integer, _
                                result()        As Variant, _
                                col             As Integer, _
                                lastCol         As Integer, _
                                prjNum          As Integer)
    
    ' exit condition : count(col) = prjNum
    Dim i As Integer
    Dim sPrjNum As Integer ' 시작 프로젝트 번호
    Dim ePrjNum As Integer ' 마지막 프로젝트 번호

    If col > (lastCol - 1)Then ' check exit condition
        ' 결과 출력
        For i = 1 To lastCol
            Debug.Print result(i);
        Next i
        Debug.Print
        Exit Sub
    End If
    
    ' song 누적이나 발생한 프로젝트 갯수가 0 인경우 예외처리 필요    
    If 0 = counts(col) Then
        result(col) = 0 ' 프로젝트 번호
        col = col + 1
        GenerateCombinationRecursive periods, accumulate, counts, result, col + 1, lastCol, i        
    End If
    
    sPrjNum = accumulate(col-1) + 1 '이전 누적 다음 번호가 이번에 발생한 프로젝트 시작 순번
    ePrjNum = accumulate(col) ' 현재 누적은 이번 발생한 프로젝트의 마지막 순번 + (counts(col)-1) ' 

    For i = sPrjNum to ePrjNum
        result(col) = i ' 프로젝트 번호
        GenerateCombinationRecursive periods, accumulate, counts, result, col + 1, lastCol, i        
    Next

End Sub


' 루프를 월단위로 돌아야 한다. ㅠㅠ
' 월별로 프로젝트를 확인하고 
' 프로젝트마다 선택여부 설정
'1.	이번달 여유인력 계산
'2.	이번달 여유 자금 계산 ( 이월 + 수입 – 지출)
'3.	진행중인 프로젝트들 중 완료된 프로젝트가 있는지 검사
'4.	이번달 신규 진행할 프로젝트 선정
'   1)	신규 진행에 인원이 부족하지는 않은가?
'   2)	

' Sub CirculateMonth()
' Sub UpdateProject
' Sub UpdateHR
' Sub UpdateCash
' Sub SelectProject


' 
    'loop 1 to 36 개월
        'loop sPrjNum  to ePrjNum  이번달 시작 프로젝트 부터 이번달 마지막 프로젝트까지
            ' UpdateProject
            ' UpdateHR
            ' UpdateCash
            ' SelectProject

Sub CirculateMonth()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("dashboard") ' 적절한 시트 이름으로 변경

    Dim lastCol         As Integer
    lastCol = ws.Cells(2, ws.Columns.Count).End(xlToLeft).Column ' 마지막 열 찾기
    
    Dim periods()       As Integer
    Dim accumulate()    As Integer ' 누적된 발주프로젝의 갯수를 넣어 둔다.
    Dim counts()        As Integer
    
    ' 기간과 갯수를 배열에 저장
    ReDim periods(0 To lastCol) 'song lastCol은 전역 변수에서 가져오는걸로 수정하자
    ReDim accumulate(0 To lastCol) 
    ReDim counts(0 To lastCol) 
    
    Dim i               As Integer
    ' 데이터는 2번째 열부터 읽어옴
    For i = 1 To lastCol ' 헤더 부분은 빼고 읽어 온다
        periods(i)      = ws.Cells(2, i+1).Value
        accumulate(i)   = ws.Cells(3, i+1).Value
        counts(i)       = ws.Cells(4, i+1).Value
    Next i
    
    Dim currentMonth As Integer
    Dim sPrjNum As Integer ' 시작 프로젝트 번호
    Dim ePrjNum As Integer ' 마지막 프로젝트 번호
    Dim prjCount As Integer 

    periods(0)      = 0
    accumulate(0)   = 0
    counts(0)       = 0

    For i = 1 To lastCol ' 모든 기간 동안 (36개월)
        currentMonth    = periods(i)
        If i = 1 Then
            sPrjNum         = 1
        Else
            sPrjNum         = accumulate(i-1) + 1 '이전 누적 다음 번호가 이번에 발생한 프로젝트 시작 순번    
        End If
        'sPrjNum         = accumulate(i) + 1 '이전 누적 다음 번호가 이번에 발생한 프로젝트 시작 순번
        ePrjNum         = accumulate(i) ' 현재 누적은 이번 발생한 프로젝트의 마지막 순번 + (counts(col)-1) '
        prjCount        = counts(i)

        ' Debug.Print currentMonth;
        ' Debug.Print sPrjNum;
        ' Debug.Print ePrjNum;
        ' Debug.Print prjCount;
        ' Debug.Print 

        Dim j               As Integer
        Dim prjNum          As Integer
        Dim sumOfrunning    As Integer

        sumOfrunning = 0

        For j = 1 to prjCount

            Dim rand As Integer

            prjNum = sPrjNum + j - 1
            Debug.Print prjNum

            rand = Int(Rnd()*100)

            If rand > 30 Then
                ws.Cells(14+sumOfrunning,i +1) = prjNum
                sumOfrunning = sumOfrunning + 1
            End If

            
        Next

        ws.Cells(13,i+1 ) = sumOfrunning


    Next

End Sub

' 이번달 진행중인 프로젝트가 있는가?
' 이번달 진행중인 프로젝트중 완료된 프로젝트가 있는가?
' 완료된 프로젝트가 있으면 인원 테이블, 보유현금 업데이트 
Sub UpdateProject(month As Integer)


    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("dashboard")

    Dim sumOfrunning As Integer
    Dim i As Integer
    
    sumOfrunning = ws.Cells(13,month +1)

    If sumOfrunning = 0 Then

    End If

    

End Sub    


