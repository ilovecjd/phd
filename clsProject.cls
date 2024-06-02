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

' 출력에 필요한 환경 변수들
Public SheetName			As String ' 출력할 시트 이름
Public InfoWidth 			As Integer ' 출력할 시트의 가로 크기
Public InfoHeight 			As Integer ' 출력할 시트의 세로 크기
Private ProjectInfo(1 To H_INFO, 1 To W_INFO) As Variant ' 출력할 헤더를 기록하는 배열

' 프로젝트 속성
Public ProjectType			As Integer ' 프로젝트 타입 (0: 외부 / 1: 내부)
Public ProjectNum 			As Integer ' 프로젝트의 번호
Public ProjectDuration		As Integer ' 프로젝트의 총 기간

Public OrderDate 			As Integer ' 발주일
Public PossibleStartDate 	As Integer ' 시작 가능일
Public StartDate 			As Integer ' 시작일 (0: 미시작)

Public Profit 				As Double ' 총 기대 수익 (HR 종속)
Public Experience 			As Integer ' 경험 (0: 무경험 1: 유경험)
Public SuccessProbability 	As Integer ' 성공 확률

' 활동
Public NumActivities 		As Integer ' 총 활동 수
Private Activities(1 To MAX_ACT) As Activity ' 활동에 관한 정보를 기록하는 배열

' 현금 흐름
Public NumCashFlows 		As Integer ' 비용 지급 횟수
Private CashFlows(1 To MAX_N_CF) As Integer ' 용역비를 받는 비율을 기록하는 배열
Public FirstPayment 		As Long ' 선금 액수
Public MiddlePayment 		As Long ' 2차 지급 액수
Public FinalPayment 		As Long ' 3차 지급 액수
Public FirstPaymentMonth 	As Integer ' 선금 지급 월
Public MiddlePaymentMonth	As Integer ' 2차 지급 월
Public FinalPaymentMonth 	As Integer ' 3차 지급 월

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Desc     : Project 클래스의 초기화
' initType : 프로젝트 타입 (0: 외부 / 1: 내부)
' initNum  : 프로젝트의 관리 번호, 프로젝트마다 유니크함
' shtName: 출력할 시트의 이름
' ODate    : 프로젝트가 발생한 달
Public Sub Init(initType As Integer, initNum As Integer, shtName As String, ODate As Integer)
    ' 상수 값 설정
    InfoWidth 	= W_INFO
    InfoHeight 	= H_INFO

    ' 전달받은 값들로 초기화
    ProjectType = initType
    ProjectNum 	= initNum
    SheetName 	= shtName

    ' 1번 프로젝트 생성 시 프로젝트 시트의 헤더를 만든다.
    If initNum = 1 Then
        Call PrintHeader
    End If

    ' 프로젝트 생성을 위해 필요한 값들을 생성한다.
    OrderDate 			= ODate ' 발주월
    PossibleStartDate 	= ODate + Int(Rnd() * 4) ' 시작 가능일
    StartDate 			= 0 ' 시작일 (0: 미시작)
    Profit 				= 100 ' 총 기대 수익 (HR 종속 MakeProjectHRandProfit 함수에서 생성)
    Experience 			= GenerateWithProbability(95) ' 경험 (0: 무경험 1: 유경험)
    SuccessProbability	= 100 ' 성공 확률
    NumCashFlows 		= MAX_N_CF ' 비용 지급 횟수 (규모에 따라 변경 가능)

    Call CreateActivities ' 활동과 상호 선후 관계 생성, 프로젝트 전체 기간 갱신
    Call CalculateHRandProfit ' 활동별 투입 인력 생성, 프로젝트 전체 기대 수익 생성 및 기록
    Call CalculatePaymentSchedule ' 대금 지급 조건 생성
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Desc : 활동 생성, 활동 선후 관계, 프로젝트 전체 기간 갱신
Public Function CreateActivities()
    Randomize ' 난수 발생기 초기화

    ' 1. 활동의 갯수 생성
    NumActivities = Int(Rnd() * MAX_ACT) + 1

    ' 2. 활동의 기간을 생성한다
    Dim i As Integer
    Dim totalDuration As Integer ' 프로젝트의 기간 (가장 늦게 끝나는 활동의 종료시점이 프로젝트의 기간이 된다)
    totalDuration = Int(3 * Rnd + 1)

    For i = 1 To NumActivities
        Dim temp 	As Integer
        Dim current	As Integer
        Dim result 	As Integer
        Dim j 		As Integer
        Dim k 		As Integer

        Activities(i).Duration 	= Int(Rnd() * 4) + 1 ' 활동 기간

        Activities(i).StartDate	= 0
        Activities(i).EndDate 	= Activities(i).Duration

        If i > 1 Then
            For j = 1 To (i - 1)
                k = Int(Rnd() * 100)
                If k > 50 Then
                    temp 					= Activities(j).EndDate
                    current 				= Activities(i).StartDate
                    result 					= max(temp, current)
                    Activities(i).StartDate = result
                    Activities(i).EndDate 	= Activities(i).StartDate + Activities(i).Duration
                End If
            Next j
        End If

        totalDuration = max(totalDuration, Activities(i).EndDate)
    Next i

    ' 3. 프로젝트의 전체 기간을 업데이트 한다.
    ProjectDuration = totalDuration
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Desc : 활동별 투입 인력 생성, 프로젝트 전체 기대 수익 생성 및 기록
Function CalculateHRandProfit()
    Dim i As Integer
    Dim j As Integer
    Dim high As Integer, mid As Integer, low As Integer
    high = 0: mid = 0: low = 0 ' 등급별 M/M

    For i = 1 To NumActivities
        j = Int(Rnd() * 100)
        If (0 < j) And (j <= RND_HR_H) Then
            Activities(i).HighSkill	= 1
        ElseIf (RND_HR_H < j) And (j <= RND_HR_M) Then
            Activities(i).MidSkill 	= 1
        Else
            Activities(i).LowSkill 	= 1
        End If
    Next i

    For i = 1 To NumActivities
        high	= high + Activities(i).HighSkill * Activities(i).Duration
        mid		= mid  + Activities(i).MidSkill  * Activities(i).Duration
        low 	= low  + Activities(i).LowSkill  * Activities(i).Duration
    Next i

    Profit = CalculateTotalLaborCost(high, mid, low)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Desc : 등급별 투입인력 계산 및 프로젝트의 수익 생성
Public Function CalculateTotalLaborCost(highCount As Integer, midCount As Integer, lowCount As Integer) As Integer
    Dim highLaborCost	As Integer
    Dim midLaborCost 	As Integer
    Dim lowLaborCost 	As Integer

    highLaborCost	= CalculateLaborCost("H") * highCount
    midLaborCost 	= CalculateLaborCost("M") * midCount
    lowLaborCost 	= CalculateLaborCost("L") * lowCount

    CalculateTotalLaborCost = highLaborCost + midLaborCost + lowLaborCost
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Desc : 등급별 투입인력에 따른 수익 계산
Private Function CalculateLaborCost(grade As String) As Integer
    Dim directLaborCost	As Integer
    Dim overheadCost 	As Integer
    Dim technicalFee 	As Integer
    Dim totalLaborCost 	As Integer

    grade = UCase(grade)

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

    overheadCost 	= directLaborCost * 0.6
    technicalFee 	= (directLaborCost + overheadCost) * 0.2
    totalLaborCost	= directLaborCost + overheadCost + technicalFee

    CalculateLaborCost = totalLaborCost
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Desc : 대금 지급 조건 생성
Public Function CalculatePaymentSchedule()
    Dim paymentType 	As Integer
    Dim paymentRatio 	As Integer
    Dim totalPayments 	As Integer

    FirstPaymentMonth = 1

    Randomize ' 난수 발생기 초기화

    ' 1번 조건: 일한 개월수가 6개월 미만일 때
    If ProjectDuration < 6 Then
        paymentType = Int(3 * Rnd + 1)

        Select Case paymentType
            Case 1
                FirstPayment = Profit * 0.3: CashFlows(1) = 30: CashFlows(2) = 70
            Case 2
                FirstPayment = Profit * 0.4: CashFlows(1) = 40: CashFlows(2) = 60
            Case 3
                FirstPayment = Profit * 0.5: CashFlows(1) = 50: CashFlows(2) = 50
        End Select

        FinalPayment = Profit - FirstPayment
        totalPayments = 2
        MiddlePaymentMonth = ProjectDuration

    ' 2번 조건: 일한 개월수가 6개월 이상 12개월 미만일 때
    ElseIf ProjectDuration < 12 Then
        paymentType = Int(10 * Rnd + 1)

        If paymentType <= 3 Then
            paymentRatio = Int(3 * Rnd + 1)
            Select Case paymentRatio
                Case 1
                    FirstPayment = Profit * 0.3: CashFlows(1) = 30: CashFlows(2) = 70
                Case 2
                    FirstPayment = Profit * 0.4: CashFlows(1) = 40: CashFlows(2) = 60
                Case 3
                    FirstPayment = Profit * 0.5: CashFlows(1) = 50: CashFlows(2) = 50
            End Select
            FinalPayment = Profit - FirstPayment
            totalPayments = 2
            MiddlePaymentMonth = ProjectDuration

        Else
            paymentRatio 		= Int(10 * Rnd + 1)
            If paymentRatio <= 6 Then
                FirstPayment 	= Profit * 0.3
                MiddlePayment 	= Profit * 0.3
                CashFlows(1) 	= 30
                CashFlows(2) 	= 30
                CashFlows(3) 	= 40
            Else
                FirstPayment	= Profit * 0.3
                MiddlePayment 	= Profit * 0.4
                CashFlows(1) 	= 30
                CashFlows(2) 	= 40
                CashFlows(3) 	= 30
            End If

            FinalPayment 		= Profit - FirstPayment - MiddlePayment
            totalPayments 		= 3
            MiddlePaymentMonth	= WorksheetFunction.Ceiling(ProjectDuration / 2, 1)
            FinalPaymentMonth 	= ProjectDuration
        End If

    ' 3번 조건: 일한 개월수가 12개월 이상일 때
    Else
        FirstPayment 		= Profit * 0.3
        MiddlePayment		= Profit * 0.4
        FinalPayment 		= Profit - FirstPayment - MiddlePayment

        CashFlows(1) 		= 30
        CashFlows(2) 		= 40
        CashFlows(3) 		= 30

        totalPayments 		= 3
        MiddlePaymentMonth	= WorksheetFunction.Ceiling(ProjectDuration / 2, 1)
        FinalPaymentMonth 	= ProjectDuration
    End If

    NumCashFlows = totalPayments ' 지급 횟수

    ' 지급 일정 출력
    Debug.Print "총 지급 횟수: " & NumCashFlows
    Debug.Print "지급 일정: "
    If NumCashFlows = 2 Then
        Debug.Print FirstPaymentMonth & ", 선금, " & FirstPayment
        Debug.Print MiddlePaymentMonth & ", 잔금, " & FinalPayment
    ElseIf NumCashFlows = 3 Then
        Debug.Print FirstPaymentMonth & ", 선금, " & FirstPayment
        Debug.Print MiddlePaymentMonth & ", 중도금, " & MiddlePayment
        Debug.Print FinalPaymentMonth & ", 잔금, " & FinalPayment
    End If
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Desc : 주어진 확률에 따라서 1 또는 0을 반환하는 함수
Function GenerateWithProbability(probability As Integer) As Integer
    Dim randomProb As Double
    Randomize
    randomProb = Rnd
    If randomProb <= probability / 100 Then
        GenerateWithProbability = 1
    Else
        GenerateWithProbability = 0
    End If
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Desc : 프로젝트들을 출력할 엑셀 시트에 헤더를 생성한다.
Public Function PrintHeader()

    Dim MyArray() 	As String, strHeader As String
    Dim sheetPrj 	As Worksheet

    Set sheetPrj 	= Worksheets(SheetName)
    With sheetPrj
        strHeader	= "타입 순번 발주일 시작가능 기간 시작 수익 경험 성공% CF1% CF2% CF3%"
        MyArray 	= Split(strHeader)
        .Range(.Cells(1, 1), .Cells(1, InfoWidth)).Value = MyArray

        strHeader 	= "Dur start end HR_H HR_M HR_L mon_cf1 mon_cf2 mon_cf3"
        MyArray 	= Split(strHeader)
        .Range(.Cells(2, 1), .Cells(2, InfoWidth)).Value = MyArray
    End With

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Desc : 프로젝트의 모든 정보를 엑셀에 출력한다.
Public Function PrintInfo()
    Dim i 		As Integer
    Dim posX	As Integer ' 출력할 가로 위치
    Dim posY 	As Integer ' 출력할 세로 위치

    ' 기본 정보 출력
    posX = 1: posY = 1: ProjectInfo(posY, posX) = ProjectType ' 프로젝트 타입 (0: 외부 / 1: 내부)
    posX = posX + 1: ProjectInfo(posY, posX) = ProjectNum ' 순번
    posX = posX + 1: ProjectInfo(posY, posX) = OrderDate ' 발주일
    posX = posX + 1: ProjectInfo(posY, posX) = PossibleStartDate ' 시작 가능일
    posX = posX + 1: ProjectInfo(posY, posX) = ProjectDuration ' 기간
    posX = posX + 1: ProjectInfo(posY, posX) = StartDate ' 시작일
    posX = posX + 1: ProjectInfo(posY, posX) = Profit ' 수익
    posX = posX + 1: ProjectInfo(posY, posX) = Experience ' 경험 유무
    posX = posX + 1: ProjectInfo(posY, posX) = SuccessProbability ' 성공 확률

    For i = 1 To NumCashFlows
        posX = posX + 1: ProjectInfo(posY, posX) = CashFlows(i)
    Next i

    posX = 1: posY = 2: ProjectInfo(posY, posX) = "ActivityCnt"
    posX = posX + 1: ProjectInfo(posY, posX) = NumActivities
    posX = posX + 8: ProjectInfo(posY, posX) = FirstPaymentMonth
    posX = posX + 1: ProjectInfo(posY, posX) = MiddlePaymentMonth
    posX = posX + 1: ProjectInfo(posY, posX) = FinalPaymentMonth

    ' 활동 및 인력 정보 출력
    For i = 1 To MAX_ACT
        Dim strAct As String, strNum As String
        strNum = Format(i, "#0")
        strAct = "Activity" & strNum

        posX = 1: posY = posY + 1: ProjectInfo(posY, posX) = strAct
        posX = posX + 1: ProjectInfo(posY, posX) = Activities(i).Duration
        posX = posX + 1: ProjectInfo(posY, posX) = Activities(i).StartDate
        posX = posX + 1: ProjectInfo(posY, posX) = Activities(i).EndDate
        posX = posX + 1: ProjectInfo(posY, posX) = Activities(i).HighSkill
        posX = posX + 1: ProjectInfo(posY, posX) = Activities(i).MidSkill
        posX = posX + 1: ProjectInfo(posY, posX) = Activities(i).LowSkill
    Next i

    Dim sheetPrj As Worksheet

    Set sheetPrj = Worksheets(SheetName)

    posY = (ProjectNum - 1) * InfoHeight + 3

    With sheetPrj
        Dim RNG1 As Range
        Set RNG1 = .Range(.Cells(posY, 1), .Cells(posY + InfoHeight - 1, InfoWidth))
        RNG1.Value = ProjectInfo
        RNG1.BorderAround LineStyle:=1, Weight:=xlThin, ColorIndex:=1 ' 테두리 그리기
    End With

End Function
