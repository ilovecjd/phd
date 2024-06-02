VERSION 1.0 CLASS
BEGIN
MultiUse = -1  'True
END
Attribute VB_Name = "clsProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 프로젝트 클래스. 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'' 출력에 필요한 환경 변수들
Public strSheet             	As String   ' 출력할 시트 이름
Public widthOfInfo 				As Integer 	' 출력할 시트의 가로 크기
Public heightOfInfo 			As Integer 	' 출력할 시트의 세로 크기
Private arrPrjInfo( 1 to H_INFO , 1 to W_INFO)	As Variant ' 출력할 해더를 기록하는 배열

'' 프로젝트 속성
Public  prjType             	As Integer  ' 프로젝트 타입 (0: 외부 / 1: 내부)
Public  num                 	As Integer  ' 프로젝트의 번호
Public  monthsWork          	As Integer  ' 프로젝트의 총 기간

Public  orderDate           	As Integer  ' 발주일
Public  possibleDate        	As Integer  ' 시작가능일
Public  startDate           	As Integer  ' 시작일 (0: 미시작)

Public  profit              	As Double  	' 총 기대 수익 (HR종속)
Public  experience          	As Integer  ' 경험 (0 : 무경험 1:유경험)
Public  successProb         	As Integer  ' 성공확율  

'' Activity
Public  num_Act             	As Integer  ' Number of total activity
Private arrActivity(1 to MAX_ACT)  As ACTIVITY_ ' 활동에 관한 정보를 기록하는 배열

'' cash flower
Public  num_CF              	As Integer  ' Number of CF (비용 지급 횟수 1cf 40%, 2cf 30%, 3cf 40% 등)
private arrCF(1 to MAX_N_CF)	As Integer  ' 용역비를 받는 비율을 기록하는 배열. 가변적으로 하지 말자.. 귀찮다.
Public firstPayment 			As Long		' 선금 액수
Public middlePayment 			As Long		' 2차지급 액수
Public finalPayment 			As Long		' 3차지급 액수
Public firstMonth 				As Integer	' 선금 지급월
Public middleMonth				As Integer	' 2차 지급월
Public finalMonth 				As Integer	' 3차 지급월


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Desc 		: Project 클래스의 초기화
' initType 	: 프로젝트 타입 (0: 외부 / 1: 내부)
' initNum 	: 프로젝트의 관리 번호, 프로젝트마다 유니크함
' sheetName : 출력할 시트의 이름
' ODate 	: 프로젝트가 발생한 달
Public Sub Init(initType As Integer, initNum As Integer, sheetName As String, ODate As Integer)

	' 상수 값 설정
	widthOfInfo 	= W_INFO
	heightOfInfo	= H_INFO

	' 전달받은 값들로 초기화
	prjType       	= initType
	num           	= initNum
	strSheet      	= sheetName

	' 1 번 프로젝트 생성시 프로젝트 시트의 헤더를 만든다.
	If initNum = 1 Then 
		Call PrintHeader
	End If

	' 프로젝트 생성을 위해서 필요한 값들을 생성한다.
	' monthsWork    	= Int(Rnd() * 100 Mod 12) + 1 ' 기간은 12개월 이내 ==> activity 에 영향을 주고 받을까 고민중
	orderDate     	= ODate    ' 발주월
	possibleDate  	= ODate + Int(Rnd() * 100 Mod 4) ' 시작가능일 song ==> 프로젝트 규모에 따라서 조금 크거나 작게
	startDate     	= 0   ' 시작일 (0: 미시작) 
	profit        	= 100 ' 총 기대 수익 (HR종속 MakePrjHRand 함수에서 생성) 
	experience    	= GenerateWithProbability(95) ' 경험 (0 : 무경험 1:유경험)
	successProb   	= 100 ' 성공확율    song ==> 
	num_CF        	= MAX_N_CF  ' Number of CF (비용 지급 횟수 1cf 40%, 2cf 30%, 3cf 40% 등) song ==> 규모에 따라서 변경 1 ~ 3, 1의 경우 시작 또는 끝중 하나
	num_Act       	= Int(Rnd() * 100 Mod MAX_ACT) + 1 ' song 비율은 추후 수정' Number of total activity
	
	Call MakePrjActivity() 			' Activity와 상호 선후 관계 생성, 프로젝트 전체 기간 갱신
	Call MakePrjHRandProfit() 		' Activity별 투입 인력 생성, 프로젝트 전체 기대 수익 생성 및 기록
	Call CalculatePaymentSchedule()	' 대금 지급 조건 생성	
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Desc : Activity 생성, 활동 선후 관계, 프로젝트 전체 기간 갱신
Public Function MakePrjActivity()

	' 1. 활동의 기간을 생성한다
	Dim i As Integer
	Dim totalDur As Integer ' 프로젝트의 기간 (가장 늦게 끝나는 활동의 종료시점이 프로젝트의 기간이 된다)
	totalDur = 0

	For i = 1 to num_Act  ' 1 <= num_Act =< 6		

		Dim temp    As Integer
		Dim current As Integer
		Dim result  As Integer

		Dim j       As Integer
		Dim k       As Integer

		arrActivity(i).duration = Int(Rnd() * 100 Mod 4) + 1 'duration song 비율은 추후 수정
	
		arrActivity(i).start  = 0
		arrActivity(i).end    = arrActivity(i).duration
		'totalDur = arrActivity(i).end

		If i = 1 Then     ' 첫 활동은 선후관계 없이 시작
		
		Else ' 2. 활동이 2개 이상이면 선후 관계를 구한다.	i 가 1보다 크면 위와 선후관계 비교

			For j = 1 to (i-1)

				k = Int(Rnd()*100) 

				If k > 50 Then ' 연관관계 있음 song ==> 상수로 변경하자  ' 연관관계있는 것의 완료일과 나의 시작일중 큰것으로 교체
					temp                  = arrActivity(j).end' 위에것의 완료일
					current               = arrActivity(i).start' 나의 시작일
					result                = max(temp,current)
					arrActivity(i).start  = result ' 시작일 교체
					arrActivity(i).end    = arrActivity(i).start + arrActivity(i).duration'완료일 교체                
				
				Else ' 다른 활동과 연관 없음. ***song 지금은 첫달부터 시작하는 걸로되어 있음. 수정 필요
                    'arrActivity(i).start   = 0 ' 시작일 교체
                    'arrActivity(i).end 		= arrActivity(i).start + arrActivity(i).duration'완료일 교체              
                
				End If			

			Next                

		End If

		totalDur = max(totalDur , arrActivity(i).end)

	Next

	'3. 프로젝트의 전체 기간을 업데이트 한다.
	monthsWork = totalDur

End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Desc : Activity별 투입 인력 생성, 프로젝트 전체 기대 수익 생성 및 기록
Function MakePrjHRandProfit()
        
    Dim i As Integer
	Dim j As Integer
	Dim high As Integer, mid As Integer, low As Integer
	high = 0 : mid = 0 : low = 0 ' 등급별 M/M
  
	'song ==> 여러명이 필요한 경우로 수정하자.
    For i = 1 to num_Act ' i 는 프로젝트의 총 activity 갯수

        j = Int(Rnd() *100)

        If (0< j) and (j <= RND_HR_H) Then
           arrActivity(i).hr_H  = 1		   

        ElseIf (RND_HR_H < j) and (j <= RND_HR_M) Then
            arrActivity(i).hr_M  = 1			
        Else
            arrActivity(i).hr_L   = 1			
        End If        
    Next

	For i = 1 to num_Act ' i 는 프로젝트의 총 activity 갯수
		high	= high + arrActivity(i).hr_H * arrActivity(i).duration
		mid		= mid + arrActivity(i).hr_M * arrActivity(i).duration
		low		= low + arrActivity(i).hr_L * arrActivity(i).duration
    Next

	' 프로젝트별 수익은 int 형을 사용, 합산은 double형 사용
    profit = CalculateTotalLaborCost(high, mid, low)

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Desc : 등급별 투입인력계산 및 프로젝트의 수익 생성
Public Function CalculateTotalLaborCost(highCount As Integer, midCount As Integer, lowCount As Integer) As Integer
    Dim highLaborCost As Integer
    Dim midLaborCost As Integer
    Dim lowLaborCost As Integer
    
    ' 상급 인건비 계산
    highLaborCost = CalculateLaborCost("H") * highCount
    
    ' 중급 인건비 계산
    midLaborCost = CalculateLaborCost("M") * midCount
    
    ' 하급 인건비 계산
    lowLaborCost = CalculateLaborCost("L") * lowCount
    
    ' 총 인건비 계산
    CalculateTotalLaborCost = highLaborCost + midLaborCost + lowLaborCost
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Desc : 등급별 투입인력에 따른 수익 계산
Private Function CalculateLaborCost(grade As String) As Integer
    Dim directLaborCost	As Integer
    Dim overheadCost	As Integer
    Dim technicalFee	As Integer
    Dim totalLaborCost	As Integer
    
    ' 등급을 대문자로 변환
    grade = UCase(grade)
    
    ' 등급에 따른 직접인건비 설정
    Select Case grade
        Case "H"
            directLaborCost = 500
        Case "M"
            directLaborCost = 390
        Case "L"
            directLaborCost = 250
        Case Else
            MsgBox "잘못된 등급입니다. 'H', 'M', 'L' 중 하나를 입력하세요."
            Exit Function
    End Select
    
    ' 제경비 계산 (직접인건비의 60%)
    overheadCost = directLaborCost * 0.6
    
    ' 기술료 계산 (직접인건비와 제경비의 합의 20%)
    technicalFee = (directLaborCost + overheadCost) * 0.2
    
    ' 총 인건비 계산
    totalLaborCost = directLaborCost + overheadCost + technicalFee
    
    ' 결과 반환
    CalculateLaborCost = totalLaborCost
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Desc : 대금 지급 조건 생성, 6개월 미만, 1년 미만, 1년 이상일때 각각 다른 조건으로 지급조건 생성
Public Function CalculatePaymentSchedule()

	Dim paymentType 	As Integer
	Dim paymentRatio 	As Integer
	Dim totalPayments 	As Integer
	' Dim paymentSchedule As String

	firstMonth = 1 ' 중복이지만 남겨 놓자.  선금은 언제나 첫달

	Randomize ' 난수 발생기 초기화

	' 1번 조건: 일한 개월수가 6개월 미만일 때
	If monthsWork < 6 Then
		paymentType = Int(3 * Rnd + 1) ' 3:7, 4:6, 5:5 중 하나 선택

		Select Case paymentType
			Case 1
				firstPayment = profit * 0.3 : arrCF(1) = 30 : arrCF(2) = 70
			Case 2
				firstPayment = profit * 0.4 : arrCF(1) = 40 : arrCF(2) = 60
			Case 3
				firstPayment = profit * 0.5 : arrCF(1) = 50	: arrCF(2) = 50
		End Select

		finalPayment = profit - firstPayment
		totalPayments = 2        
		middleMonth = monthsWork

	' 2번 조건: 일한 개월수가 6개월 이상 12개월 미만일 때
	ElseIf monthsWork < 12 Then
		paymentType = Int(10 * Rnd + 1) ' 1-3:선금,잔금, 4-10:선금,중도금,잔금

		If paymentType <= 3 Then
			paymentRatio = Int(3 * Rnd + 1) ' 1:3:7, 2:4:6, 3:5:5 중 하나 선택
			Select Case paymentRatio
				Case 1
					firstPayment = profit * 0.3	: arrCF(1) = 30 : arrCF(2) = 70					
				Case 2
					firstPayment = profit * 0.4	: arrCF(1) = 40 : arrCF(2) = 60
				Case 3
					firstPayment = profit * 0.5 : arrCF(1) = 50 : arrCF(2) = 50
			End Select
			finalPayment = profit - firstPayment
			totalPayments = 2
			middleMonth = monthsWork

		Else
			paymentRatio = Int(10 * Rnd + 1) ' 1-6:3:3:4, 7-10:3:4:3
			If paymentRatio <= 6 Then
				firstPayment 	= profit * 0.3 
				middlePayment	= profit * 0.3 
				arrCF(1) 		= 30
				arrCF(2) 		= 30
				arrCF(3) 		= 40
			Else
				firstPayment	= profit * 0.3
				middlePayment 	= profit * 0.4				
				arrCF(1) 		= 30
				arrCF(2) 		= 40
				arrCF(3) 		= 30				
			End If

			finalPayment = profit - firstPayment - middlePayment

			totalPayments 	= 3
			middleMonth = WorksheetFunction.Ceiling(monthsWork / 2, 1)
			finalMonth = monthsWork			

		End If

	' 3번 조건: 일한 개월수가 12개월 이상일 때
	Else
		firstPayment 	= profit * 0.3
		middlePayment 	= profit * 0.4
		finalPayment 	= finalPayment = profit - firstPayment - middlePayment

		arrCF(1) 		= 30
		arrCF(2) 		= 40
		arrCF(3) 		= 30
		
		totalPayments 	= 3
		middleMonth 	= WorksheetFunction.Ceiling(monthsWork / 2, 1)
		finalMonth 		= monthsWork
	End If

	num_CF = totalPayments ' 지급 횟수

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Desc : 주어진 확율에 따라서 1또는 0을 반납하는 함수
Function GenerateWithProbability(probability As Integer) As Integer
    Dim randomProb As Double
    
    ' 난수 발생기 초기화
    Randomize
    
    ' 0에서 1 사이의 랜덤 확률 값 생성
    randomProb = Rnd
    
    ' 주어진 확률에 따라 1을 반환
    If randomProb <= probability / 100 Then
        GenerateWithProbability = 1
    Else
        GenerateWithProbability = 0
    End If
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Desc : 프로젝트들을 출력할 엑셀 시트에 헤더를 생성한다.
Public Function PrintHeader()

	Dim MyArray() As String, strHeader As String	

	Dim sheetPrj        As Worksheet
    Set sheetPrj    = Worksheets(strSheet)    

    With sheetPrj ' song ==> 출력전에 버퍼의 크기와 widthOfInfo 를 한번 비교해 봐도 좋을듯 하다
		strHeader = "타입 순번 발주일 시작가능 기간 시작 수익 경험 성공% CF1% CF2% CF3%"
		MyArray = Split(strHeader)
        .Range(.Cells(1 , 1),.Cells(1, widthOfInfo )).Value = MyArray

		strHeader = " Dur start end HR_H HR_M HR_L   mon_cf1 mon_cf2 mon_cf2"
		MyArray = Split(strHeader)
        .Range(.Cells(2 , 1),.Cells(2, widthOfInfo )).Value = MyArray
    End With 

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Desc : 프로젝트의 모든 정보를 엑셀에 출력한다.
Public Function PrintInfo()

	Dim i 		As Integer
	Dim posX 	As Integer ' 출력할 가로 위치
	Dim posY 	As Integer ' 출력할 세로 위치

	' 기본정보 출력
	posX = 1 : posY = 1 : 	arrPrjInfo(posY, posX) = prjType		' 프로젝트 타입 (0: 외부 / 1: 내부)
	posX = posX +1 :  		arrPrjInfo(posY, posX) = num 			' 순번
	posX = posX +1 : 		arrPrjInfo(posY, posX) = orderDate 		' 발주일
	posX = posX +1 : 		arrPrjInfo(posY, posX) = possibleDate	' 시작가능일
	posX = posX +1 : 		arrPrjInfo(posY, posX) = monthsWork 	' 기간
	posX = posX +1 : 		arrPrjInfo(posY, posX) = startDate 		' 시작일
	posX = posX +1 : 		arrPrjInfo(posY, posX) = profit 		' 수익
	posX = posX +1 : 		arrPrjInfo(posY, posX) = experience 	' 경험유무
	posX = posX +1 : 		arrPrjInfo(posY, posX) = successProb 	' 성공%
	' posX = posX +1 : 		arrPrjInfo(posY, posX) = arrCF(1)	 	' cf1%
	' posX = posX +1 : 		arrPrjInfo(posY, posX) = arrCF(2)	 	' cf1%
	' posX = posX +1 : 		arrPrjInfo(posY, posX) = arrCF(3)	 	' cf1%
	For i = 1 to num_CF 
		posX = posX +1 : 	arrPrjInfo(posY, posX) = arrCF(i)
	Next

	posX =1 : posY = 2 : 	arrPrjInfo(posY, posX) = "ActivityCnt"
	posX = posX +1 : 		arrPrjInfo(posY, posX) = num_Act
	posX = posX +8 : 		arrPrjInfo(posY, posX) = firstMonth
	posX = posX +1 : 		arrPrjInfo(posY, posX) = middleMonth
	posX = posX +1 : 		arrPrjInfo(posY, posX) = finalMonth
	

			
    ' activity & HR information 출력
    For i = 1 to MAX_ACT

        Dim strAct As String, strNum As String

        strNum	= Format(i,"#0")
        strAct 	= "Activity" & strNum

		posX =1 : posY = posY +1 : 	arrPrjInfo(posY, posX) = strAct

        posX = posX +1 :	arrPrjInfo(posY, posX) = arrActivity(i).duration
		posX = posX +1 :	arrPrjInfo(posY, posX) = arrActivity(i).start
		posX = posX +1 :	arrPrjInfo(posY, posX) = arrActivity(i).end
		posX = posX +1 :	arrPrjInfo(posY, posX) = arrActivity(i).hr_H
		posX = posX +1 :	arrPrjInfo(posY, posX) = arrActivity(i).hr_M
		posX = posX +1 :	arrPrjInfo(posY, posX) = arrActivity(i).hr_L

    Next

	Dim sheetPrj        As Worksheet
    Set sheetPrj    = Worksheets(strSheet)

	posY = (num-1) * heightOfInfo + 3' 출력할 세로 위치

    With sheetPrj		
		Dim RNG1 As Range        
		Set RNG1 = .Range(.Cells(posY , 1),.Cells(posY + heightOfInfo - 1 , widthOfInfo ))
		RNG1.Value = arrPrjInfo				
		RNG1.BorderAround LineStyle:=1, Weight:=xlThin, ColorIndex:=1 	'테두리 그리기		
	End With

End Function

