Option Explicit
Option Base 1

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 프로그램 동작을 위한 기본 정보들
Dim DataColumn  As Integer  ' 데이터가 들어있는 열 번호

Dim DataIndex   As Integer  ' 데이터가 시작되는 행 번호
Dim LastIndex   As Integer  ' 데이터가 끝나는 행 번호
Dim SimulTerm   As Integer  ' 시뮬레이션을 동작 시킬 기간(월)
Dim PrjDataSize As Integer  ' 프로젝트 구성 요소들의 갯수

Dim Hr_Init_H   As Integer  ' 최초에 보유한 고급 인력
Dim Hr_Init_M   As Integer  ' 최초에 보유한 중급 인력
Dim Hr_Init_L   As Integer  ' 최초에 보유한 초급 인력
Dim Hr_LeadTime As Integer  ' 인력 충원에 걸리는 시간

Dim Cash_Init   As Integer  ' 최초 보유 현금

Dim PrjMaxCnt   As Integer  ' 월간 최대 발생 프로젝트(년평균)
Dim Opt         As Double   ' 낙관적인 전망치 / 각각 독립적으로 3가지 상황을 가정
Dim ML          As Double   ' 일반적인 전망치 / 각각 독립적으로 3가지 상황을 가정
Dim Pess        As Double   ' 비관적인 전망치 / 각각 독립적으로 3가지 상황을 가정

Dim Problem     As Integer  ' 프로젝트 생성 개수 (= 문제의 개수) / MakePrj 함수의 인자

Dim SIZE_OF_PROJECT As Integer ' 프로젝트 구조체의 크기

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 입력 받을 데이터들
Dim ExtPrjType      As Integer
Dim ExtPrjActCntLB  As Integer
Dim ExtPrjActCntUB  As Integer
Dim ExtActDurLB     As Integer
Dim ExtActDurUB     As Integer
Dim ExtPrjModeCntLB As Integer
Dim ExtPrjModeCntUB As Integer


Private Sub Generation_Click()
    
    Call Prologue()        ' 파라메터를 참조해서 기본적인 데이터를 준비한다.    

    Dim         prjCnt As Integer
    prjCnt =    MakeDashboard()  ' 대시보드를 생성하고 전체 프로젝트의 갯수를 구한다.

    Call BuildPrj(prjCnt)        ' 프로젝트의 갯수만큼 프로젝트를 생성한다.


End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' desc      : 프로그램 시작을 위한 기본적인 값들을 설정한다.
'' return    : none
Sub Prologue()

    Dim i As Integer

    i           = 7 ' Global 변수가 시작 하는 행 번호
    DataColumn  = 3 ' Global 변수가 시작 하는 열 번호
                
    DataIndex   = GetDataAndSetNext(i, DataColumn)
    LastIndex   = GetDataAndSetNext(i, DataColumn)
    SimulTerm   = GetDataAndSetNext(i, DataColumn)
    PrjDataSize = GetDataAndSetNext(i, DataColumn)

    Hr_Init_H   = GetDataAndSetNext(i, DataColumn)
    Hr_Init_M   = GetDataAndSetNext(i, DataColumn)
    Hr_Init_L   = GetDataAndSetNext(i, DataColumn)
    Hr_LeadTime = GetDataAndSetNext(i, DataColumn)
    Cash_Init   = GetDataAndSetNext(i, DataColumn)

    PrjMaxCnt   = GetDataAndSetNext(i, DataColumn)
    Opt         = GetDataAndSetNext(i, DataColumn)
    ML          = GetDataAndSetNext(i, DataColumn)
    Pess        = GetDataAndSetNext(i, DataColumn)

    Problem     = GetDataAndSetNext(i, DataColumn) ' MakePrj 함수의 인자

    SIZE_OF_PROJECT = 9 ' 번호, 기간, 활동수, 활동6가지
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Desc     : 생성된 프로젝트 배열을 엑셀에 출력하는 함수.
'' @prjArr  : 프로젝트 정보를 담아서 엑셀에 출력할 배열
'' Return   : 없음
Function PrintDashboard(Arr() As Variant)

    Dim sheetPrj        As Worksheet
    Set sheetPrj    = Worksheets("dashboard")    
    With sheetPrj   
        Dim endRow As Long ' 마지막행
        Dim endCol As Long ' 마지막열
        endRow = .UsedRange.Rows.Count + .UsedRange.Row - 1
        endCol = .UsedRange.Columns.Count + .UsedRange.Column - 1

        ' 엑셀 파일의 셀들을 정리한다.    
        .Range(.Cells(1,1),.Cells(endRow,endCol)).UnMerge
        .Range(.Cells(1,1),.Cells(endRow,endCol)).Clear
        '.Range(.Cells(1,1),.Cells(endRow,endCol)).ClearContents 
      
        .Range(.Cells(1 , 1),.Cells(100, 50 )).Value = Arr

        Dim RNG1 As Range
        ' 테두리그리기
        Set RNG1 = .Range(.Cells(2,1),.Cells(4,37))
        RNG1.Borders.LineStyle = 1
        RNG1.Borders.Weight = xlThin
        RNG1.Borders.ColorIndex = 1

        ' 테두리그리기
        Set RNG1 = .Range(.Cells(12,1),.Cells(19,37))
        RNG1.Borders.LineStyle = 1
        RNG1.Borders.Weight = xlThin
        RNG1.Borders.ColorIndex = 1

        ' 테두리그리기
        Set RNG1 = .Range(.Cells(22,1),.Cells(29,37))
        RNG1.Borders.LineStyle = 1
        RNG1.Borders.Weight = xlThin
        RNG1.Borders.ColorIndex = 1

        ' 테두리그리기   
        Set RNG1 = .Range(.Cells(32,1),.Cells(39,37))
        RNG1.Borders.LineStyle = 1
        RNG1.Borders.Weight = xlThin
        RNG1.Borders.ColorIndex = 1

    End With 
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Desc     : dashboard 시트를 생성하고 값을 기록한다.
'' @        : 없음
'' Return   : 생성해야할(기간동안 발생한) 프로젝트의 갯수
Function MakeDashboard()
    Dim dashArr() As Variant
    Redim dashArr(1 to 100 , 1 to 50) ' song 추후 정상적인 값으로 파라메터 값을 수정하자

    Dim i       As Integer
    Dim sum     As Integer ' 발생한 모든 프로젝트 갯수의 누적치
    Dim prjCnt  As Integer ' 매달 발생하는 프로젝트 갯수

    dashArr(1,1) = "발주 프로젝트 현황"
    dashArr(2,1) = "월"
    dashArr(3,1) = "누계"
    dashArr(4,1) = "발주"

    sum = 0

    For i = 1 to 36

        prjCnt = Int(Rnd() * 100 Mod 4)' song 비율은 추후 수정
        sum = sum + prjCnt

        dashArr(2,i+1) = i
        dashArr(3,i+1) = sum
        dashArr(4,i+1) = prjCnt
        
    Next

    dashArr(11,1) = "진행 프로젝트 현황"
    dashArr(12,1) = "월"
    dashArr(13,1) = "합계"
    dashArr(14,1) = "prjNum"

    For i = 1 to 36
        dashArr(12,i+1) = i
    Next

    dashArr(21,1) = "보류 프로젝트 현황"
    dashArr(22,1) = "월"
    dashArr(23,1) = "합계"
    dashArr(24,1) = "prjNum"
    For i = 1 to 36
        dashArr(22,i+1) = i
    Next

    dashArr(31,1) = "완료 프로젝트 현황"
    dashArr(32,1) = "월"
    dashArr(33,1) = "합계"
    dashArr(34,1) = "prjNum"
    For i = 1 to 36
        dashArr(32,i+1) = i
    Next

    Call PrintDashboard(dashArr)

    MakeDashboard = sum

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Desc     : 셀 i,j 의 값을 리턴하고 인덱스 i 값을 다음 인덱스로 증가시킨다.
'' @i , j   : 값을 가져올 셀의 위치
'' Return   : 셀(i,j)의 값 / i (증가된  i 값)
Function GetDataAndSetNext(i As Integer, j As Integer)

    GetDataAndSetNext = Cells(i, j)
    i = i + 1

End Function

' desc      : 프로젝트를 위한 변수들을 준비한다.
'' prjCnt   : 생성해야할 프로젝트의 갯수
' return    : none
Sub BuildPrj(prjCnt As Integer)

    Dim i As Integer
    i               = DataIndex
    
    ExtPrjType      = GetDataAndSetNext(i, DataColumn)

    ExtPrjActCntLB  = GetDataAndSetNext(i, DataColumn)
    ExtPrjActCntUB  = GetDataAndSetNext(i, DataColumn)

    ExtActDurLB     = GetDataAndSetNext(i, DataColumn)
    ExtActDurUB     = GetDataAndSetNext(i, DataColumn)

    ExtPrjModeCntLB = GetDataAndSetNext(i, DataColumn)
    ExtPrjModeCntUB = GetDataAndSetNext(i, DataColumn)

    Dim array_sizeX As Integer ' 배열의 열의 갯수
    Dim array_sizeY As Integer ' 배열의 행의 갯수
    Dim prjArr()    As Variant ' 프로젝트를 담을 버퍼

    array_sizeX     = 10                ' Dur	start	end	HR_H	HR_M	HR_L

    ' 헤더 크기 + 프로젝트 갯수 * 프로젝트 구조체의 크기 ( 순번,기간,ActivityCnt,Activity1,Activity2,Activity3,Activity4,Activity5,Activity6)
    array_sizeY     = 1 + prjCnt * 9   

    Redim prjArr(1 to array_sizeY, 1 to array_sizeX) 'song 사이즈 추후 확정
    prjArr(1,1) = prjCnt    '생성될 프로젝트이 총 갯수를 적는다

    '1. 1열의 내용을 채운다.
    Call MakePrjRow(prjArr)

    '2. 프로젝트를 갯수만큼 생성한다.
    For i = 1 to prjCnt
        Call MakePrjInfo(prjArr,i)
        Call MakePrjActivity(prjArr,i)
        Call MakePrjHR(prjArr,i)
    Next
    
    '3. project 시트에 prjArr를 출력한다.
    Call PrintPrj(prjArr)
    
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' 프로젝트를 엑셀에 표시할때 사용할 첫줄(헤더)을 버퍼에 생성한다.
'' @prjArr : 프로젝트 정보를 담아서 엑셀에 출력할 배열
Function MakePrjRow(prjArr() As Variant)
   
    ' 첫 열을 적어준다. song 맨 위에만 적어 주면 될듯 하다.
    Dim i As Integer
    i = 1 ' 1열에 기록한다.
    prjArr(i,2) = "Dur"
    prjArr(i,3) = "start"
    prjArr(i,4) = "end"
    prjArr(i,5) = "HR_H"
    prjArr(i,6) = "HR_M"
    prjArr(i,7) = "HR_L"
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' 프로젝트의 기본 정보를 생성한다.
'' @prjArr : 프로젝트 정보를 담아서 엑셀에 출력할 배열
'' @prjNum : 프로젝트 번호
Function MakePrjInfo(prjArr() As Variant, prjNum As Integer)
   
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

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' 프로젝트를 엑셀에 표시할때 사용할 첫줄(헤더)을 버퍼에 생성한다.
'' @prjArr : 프로젝트 정보를 담아서 엑셀에 출력할 배열
Function MakePrjActivity(prjArr() As Variant, prjNum As Integer)
    
    Dim activityCnt As Integer        
    Dim iRow As Integer ' 프로젝트가 위치할 배열의 첫 행 번호
    iRow = (prjNum -1)* SIZE_OF_PROJECT + 2'song 예외 처리 필요 (iRow가 1(제목열) 보다 작은경우)

    activityCnt = prjArr(iRow + 2, 2)

    Dim i As Integer
    Dim iActLocation As Integer ' Activity 가 시작되는 위치

    iActLocation = iRow + 2

    '1. 활동별 duration 구한다.
   

    '1. 2열을 채운다.    
    For i = 1 to 6 ' Activity 1 ~6  song 6은 maxActSize 정도로 수정 하자

        Dim strAct As String
        Dim strNum As String

        strNum    = Format(i,"#0")
        strAct = "Activity" & strNum  
        prjArr(iActLocation + i ,1) = strAct
        
    Next

    For i = 1 to activityCnt '
        prjArr(iActLocation + i , 2) = Int(Rnd() * 100 Mod 4) + 1 'duration song 비율은 추후 수정        
        
        If i = 1 Then ' 첫 활동은 선후관계 없이 시작
            prjArr(iActLocation+ i , 3) = 0 ' start
            prjArr(iActLocation+ i , 4) = prjArr(iActLocation+ i , 2) ' end = duration
        End If
    Next

    '3. 활동이 2개 이상이면 선후 관계를 구한다.
    For i = 1 to activityCnt ' i 는 프로젝트의 총 activity 갯수
        Dim temp As Integer
        Dim current As Integer
        Dim result As Integer

        Dim j  As Integer
        Dim k As Integer


        If i = 1 Then ' i가 1이면 통과
            
        Else ' i 가 1보다 크면 위와 선후관계 비교
            prjArr(iActLocation+ i,3)   = 0

            For j = 1 to (i-1)

                k = Int(Rnd()*100) 
                If k > 50 Then ' 연관관계 있음
                    ' 연관관계있는 것의 완료일과 나의 시작일중 큰것으로 교체
                    temp            = prjArr(iActLocation+ j, 4)' 위에것의 완료일
                    current         = prjArr(iActLocation+ i, 3)' 나의 시작일
                    result          = max(temp,current)
                    prjArr(iActLocation+ i,3)   = result ' 시작일 교체
                    prjArr(iActLocation+ i,4)   = prjArr(iActLocation+ i,2) + prjArr(iActLocation+ i,3)'완료일 교체
                
                Else ' 다른 활동과 연관 없음. ***song 일단 첫달부터 시작하는 걸로
                    prjArr(iActLocation+ i,3)   = 0 ' 시작일 교체
                    prjArr(iActLocation+ i,4)   = prjArr(iActLocation+ i,2) + prjArr(iActLocation+ i,3)'완료일 교체
                
                End If
            Next                
            
        End If
        
    Next


     '4. 총 활동 기간을 구한다.
    Dim test As Integer        
    Dim testresult As Integer

    For i = 1 to activityCnt '
        test = prjArr(iActLocation + i,2) 
        testresult = test + testresult        
    Next
    prjArr(iRow + 1 ,2) = testresult

    
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' 프로젝트를 엑셀에 표시할때 사용할 첫줄(헤더)을 버퍼에 생성한다.
'' @prjArr : 프로젝트 정보를 담아서 엑셀에 출력할 배열
Function MakePrjHR(prjArr() As Variant, prjNum As Integer)
    
    Dim activityCnt As Integer        
    Dim iRow As Integer ' 프로젝트가 위치할 배열의 첫 행 번호    

    iRow = (prjNum -1)* SIZE_OF_PROJECT + 2'song 예외 처리 필요 (iRow가 1(제목열) 보다 작은경우)
    activityCnt = prjArr(iRow + 2, 2)

    Dim i As Integer
    Dim iActLocation As Integer ' Activity 가 시작되는 위치

    iActLocation = iRow + 2

    Dim rndH As Integer ' 고급 인력이 렌덤하게 선택 될 확율은 0 to rndH
    Dim rndM As Integer ' 중급 인력이 렌덤하게 선택 될 확율은 rndH to rndM    
    rndH = 20
    rndM = 70
   
    Dim k As Integer

    ' 0 을 넣을까 말까? 0 넣는 코드
    For i = 1 to activityCnt ' i 는 프로젝트의 총 activity 갯수
        prjArr(iActLocation+i,5)    = 0
        prjArr(iActLocation + i,6)  = 0
        prjArr(iActLocation+i,7)    = 0
    Next


    For i = 1 to activityCnt ' i 는 프로젝트의 총 activity 갯수

        k = Int(Rnd() *100)

        If (0< k) and (k <= rndH) Then
            prjArr(iActLocation+i,5)   = 1
        ElseIf (rndH < k) and (k <= rndM) Then
            prjArr(iActLocation + i,6)  = 1
        Else
            prjArr(iActLocation+i,7)   = 1
        End If        
     Next
    
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Desc : 생성된 프로젝트 배열을 엑셀에 출력하는 함수.
'' prjArr : 프로젝트 정보를 담아서 엑셀에 출력할 배열
Function PrintPrj(prjArr() As Variant)

    Dim array_sizeY As Integer ' 배열의 행의 갯수
    Dim prjCnt      As Integer    
    prjCnt          = prjArr(1,1)     '생성될 프로젝트의 총 갯수
    
    ' 헤더 크기 + 프로젝트 갯수 * 프로젝트 구조체의 크기 ( 순번,기간,ActivityCnt,Activity1,Activity2,Activity3,Activity4,Activity5,Activity6)
    array_sizeY     = 1 + prjCnt * 9 'song 전역으로 변경할 변수   

    Dim sheetPrj        As Worksheet
    Set sheetPrj    = Worksheets("project")    
    With sheetPrj   
        Dim endRow As Long ' 마지막행
        Dim endCol As Long ' 마지막열
        endRow = .UsedRange.Rows.Count + .UsedRange.Row - 1
        endCol = .UsedRange.Columns.Count + .UsedRange.Column - 1

        ' 엑셀 파일의 셀들을 정리한다.    
        .Range(.Cells(1,1),.Cells(endRow,endCol)).UnMerge
        .Range(.Cells(1,1),.Cells(endRow,endCol)).Clear
        '.Range(.Cells(1,1),.Cells(endRow,endCol)).ClearContents 
      
        .Range(.Cells(1 , 1),.Cells(array_sizeY, 10 )).Value = prjArr

    End With 
    
End Function



Public Function max(x, y As Variant) As Variant
    max = IIf(x > y, x, y)
End Function



Sub Epilogue()

End Sub


Sub old_MakePrj(sCnt As Integer)

    ' 인덱스 변수들
    Dim s As Integer,  i As Integer,  j As Integer,  k As Integer

    ' 환경 변수들
    Dim posX            As Integer  ' 출력 시작 x
    Dim posY            As Integer  ' 출력 시작 y 
    Dim headY           As Integer  ' 헤더의 줄 수

    Dim sum             As Integer  ' 매월 수주한 프로젝트의 합
    Dim sheetPrj        As Worksheet

    Dim strHead(1 to 3) As String   ' 발생 확율 타이틀
    Dim p(1 to 3)       As Variant  ' 발생 확율
    Dim rangeEnd        As Integer

    strHead(1)  = "Opt" : strHead(2)    = "ML"  : strHead(3)    = "Pess"
    p(1)        = 0.6   : p(2)          = 0.25  : p(3)          = 0.175 
    
    posX        = 2 ' 언제나 2보다 커야한다. (개월수를 표시할 공간을 위해서)
    posY        = 2
    headY       = 3 ' 헤더는 3줄이 들어간다.
    rangeEnd    = posY + headY + SimulTerm - 1 ' 배열 복사할 Range의 끝 위치
 
    ' 사용된 모든 범위에서 마지막셀을 찾아서 병합을 풀고 내용을 지운다.    
    'With ThisWorkBook.WorkSheets("project")
    Set sheetPrj    = Worksheets("project")    
    With sheetPrj   
        Dim endRow As Long ' 마지막행
        Dim endCol As Long ' 마지막열
        endRow = .UsedRange.Rows.Count + .UsedRange.Row - 1
        endCol = .UsedRange.Columns.Count + .UsedRange.Column - 1
    
        .Range(.Cells(1,1),.Cells(endRow,endCol)).UnMerge
        .Range(.Cells(1,1),.Cells(endRow,endCol)).Clear
        '.Range(.Cells(1,1),.Cells(endRow,endCol)).ClearContents 

        For i = 1 to sCnt
            .Range(.Cells(posY    , (i-1) * 3 + posX),.Cells(posY    ,i * 3 +posX-1)).Merge
            .Range(.Cells(posY    , (i-1) * 3 + posX),.Cells(posY    ,i * 3 +posX-1)).Value = i
            .Range(.Cells(posY + 1, (i-1) * 3 + posX),.Cells(posY + 1,i * 3 +posX-1)).Value = strHead
            .Range(.Cells(posY + 2, (i-1) * 3 + posX),.Cells(posY + 2,i * 3 +posX-1)).Value = p
        Next

        Dim term() As Integer
        Redim term(1 to SimulTerm, 1 to 1)

        For i = 1 to SimulTerm
            term(i,1) = i
        Next
        .Range(.Cells(posY+headY , posX-1),.Cells(rangeEnd, posX-1 )).Value = term

    End With 

    Dim     tempArr()   As  Integer
    Redim   tempArr(1 to SimulTerm, 1 to 3)

    ' 필요한 프로젝트 시나리오 수 만큼 생성
    For s = 1 To sCnt

        ' 시뮬레이션 하고자 하는 기간 만큼 생성
        For i = 1 To SimulTerm

            ' 3가지 전망에 관해서 발생할 프로젝트 숫자를 생성
            For j = 1 To 3                
                sum = 0

                ' 매달 발생 예상 프로젝트 수
                For k = 1 To PrjMaxCnt
                    If Rnd < p(j) Then sum = sum + 1
                Next k

                tempArr( i , j) = sum

            Next j

        Next i

        ' 생성된 시나리오를 복사해서 옮긴다.
        With sheetPrj
            .Range(.Cells(posY+headY , (s-1) * 3 + posX),.Cells(rangeEnd, s * 3 + posX - 1)).Value = tempArr
        End With

    Next s

End Sub