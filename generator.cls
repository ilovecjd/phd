Option Explicit


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 프로그램 동작을 위한 기본 정보들
Dim Opt         As Double   ' 낙관적인 전망
Dim ML          As Double   ' 일반적인 전망
Dim Pess        As Double   ' 비관적인 전망
Dim DataColumn  As Integer  ' 데이터가 들어있는 열 번호
Dim DataIndex   As Integer  ' 데이터가 시작되는 행 번호
Dim LastIndex   As Integer  ' 데이터가 끝나는 행 번호
Dim SimulTerm   As Integer  ' 시뮬레이션을 동작 시킬 기간(월)
Dim PrjDataSize As Integer  ' 프로젝트 구성 요소들의 갯수


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 입력 받을 데이터들
Dim ExtPrjType      As Integer
Dim ExtPrjActCntLB  As Integer
Dim ExtPrjActCntUB  As Integer
Dim ExtActDurLB     As Integer
Dim ExtActDurUB     As Integer
Dim ExtPrjModeCntLB As Integer
Dim ExtPrjModeCntUB As Integer


' desc      : 프로그램 시작을 위한 기본적인 값들을 설정한다.
' return    : none
Sub Prologue()

    Dim i As Integer

    i           = 7 ' Global 변수가 시작 하는 행 번호
    DataColumn  = 3 ' Global 변수가 시작 하는 열 번호
                
    DataIndex   = GetDataAndSetNext(i, DataColumn)
    LastIndex   = GetDataAndSetNext(i, DataColumn)
    SimulTerm   = GetDataAndSetNext(i, DataColumn)
    PrjDataSize = GetDataAndSetNext(i, DataColumn)

End Sub


Private Sub Generation_Click()

    Prologue
    BuildPrj

End Sub


' desc      : 셀 i,j 의 값을 리턴하고 인덱스 i 값을 다음 인덱스로 증가시킨다.
' return    : 셀 i,j 의 값  / 증가된  i 값
Function GetDataAndSetNext(i As Integer, j As Integer)

    GetDataAndSetNext = Cells(i, j)
    i = i + 1

End Function

' desc      : 프로젝트를 위한 변수들을 준비한다.
' return    : none
Sub BuildPrj()

    Dim i As Integer
    i               = DataIndex

    Opt             = GetDataAndSetNext(i, DataColumn)
    ML              = GetDataAndSetNext(i, DataColumn)
    Pess            = GetDataAndSetNext(i, DataColumn)
    
    ExtPrjType      = GetDataAndSetNext(i, DataColumn)

    ExtPrjActCntLB  = GetDataAndSetNext(i, DataColumn)
    ExtPrjActCntUB  = GetDataAndSetNext(i, DataColumn)

    ExtActDurLB     = GetDataAndSetNext(i, DataColumn)
    ExtActDurUB     = GetDataAndSetNext(i, DataColumn)

    ExtPrjModeCntLB = GetDataAndSetNext(i, DataColumn)
    ExtPrjModeCntUB = GetDataAndSetNext(i, DataColumn)

    ' project 시트에 프로젝트를 생성한다.
    MakePrj (100)
    
End Sub


' sCnt : 발생시킬 시나리오의 갯수
Sub MakePrj(sCnt As Integer)

    ' 인덱스 변수들
    Dim s As Integer,  i As Integer,  j As Integer,  k As Integer

    ' 환경 변수들
    Dim posX        As Integer  ' 출력 시작 x
    Dim posY        As Integer  ' 출력 시작 y 
    Dim headY       As Integer  ' 헤더의 줄 수

    Dim sum         As Integer  ' 매월 수주한 프로젝트의 합
    Dim sheetPrj    As Worksheet

    Dim strHead(1 to 3) As String   ' 발생 확율 타이틀
    Dim p(1 to 3)       As Variant  ' 발생 확율

    strHead(1)  = "Opt" : strHead(2)    = "ML"  : strHead(3)    = "Pess"
    p(1)        = 0.6   : p(2)          = 0.25  : p(3)          = 0.175 
    
    posX        = 2
    posY        = 2
    headY       = 3             ' 헤더는 3줄이 들어간다.
 
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
                For k = 1 To 4
                    If Rnd < p(j) Then sum = sum + 1
                Next k

                tempArr( i , j) = sum

            Next j

        Next i

        ' 생성된 시나리오를 복사해서 옮긴다.
        With sheetPrj
            .Range(.Cells(posY+headY , (s-1) * 3 + posX),.Cells(headY + SimulTerm, s * 3 + posX - 1)).Value = tempArr
        End With

    Next s

End Sub


Sub Epilogue()

End Sub



