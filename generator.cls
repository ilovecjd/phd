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


    Dim myPrj As New clsProject

    myPrj.index = 0


    'Prologue
    'BuildPrj

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
    MakePrj (1)
    
End Sub

' 프로젝트의 요소들을 prjCnt만큼 기록한다.
Sub MakePrj(prjCnt As Integer)

    Dim i           As Integer
    Dim posX        As Integer ' 출력 시작 x
    Dim posY        As Integer ' 출력 시작 y

    Dim SheetPrj    As Worksheet

    posX = 1
    posY = 1

    Set SheetPrj = Worksheets("project")

    ' 시뮬레이션 하고자 하는 기간을 기록한다.
    For i = 1 To SimulTerm
        SheetPrj.Cells(posY, posX + i).Value = i
    Next

    SheetPrj.Cells(posY + 1, 1).Value = "type"
    SheetPrj.Cells(posY + 2, 1).Value = "ActCnt"
    SheetPrj.Cells(posY + 3, 1).Value = "ActDur"
    SheetPrj.Cells(posY + 4, 1).Value = "ActMode"

End Sub


Sub Epilogue()

End Sub



