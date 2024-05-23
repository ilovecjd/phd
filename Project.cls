Option Explicit

Option Base 1

Private Sub CommandButton1_Click()

    Dim strTile
    strTile = Array("빈도","시나리오","월","프로젝트","기간","활동수", _ 
                    "활동","활동시작","활동끝","r1","r2","r3")

    Dim sheetActtmp As Worksheet 
    set sheetActtmp = Worksheets("act_temp")

    With sheetActtmp
        .Range(.Cells(1,1),.Cells(1,12)).Value  = strTile
    End With

    Dim freq        As Integer
    Dim scenario    As Integer
    Dim month       As Integer
    Dim project     As Integer
    Dim prjTime     As Integer
    Dim actCnt      As Integer    
    Dim actNum      As Integer
    Dim actSTime    As Integer
    Dim actETime    As Integer
    Dim r1          As Integer
    Dim r2          As Integer
    Dim r3          As Integer

    freq            = 1
    scenario        = 2
    month           = 3
    project         = 4
    prjTime         = 5
    actCnt          = 6
    actNum          = 7
    actSTime        = 8
    actETime        = 9
    r1              = 10
    r2              = 11
    r3              = 12

    Dim h As Integer
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer

    ' 참고 : 36개월 * 5(최대프로젝트) * 8(최대활동) = 1440 개의 활동
    ' 1. 단기프로젝트 - 전체 프로젝트의 60%로 구성, 최소 1개월 ~ 최대 6개월 이내
    ' 프로젝트 별 활동수를 구한다.

    ' 활동의 최대수는 4개로 한다.

    ' 활동수가 1이면 적은 기간으로 설정 (2개월 이내)

    ' 활동수가 2이면 확율적으로 겹치게 한다. (4개월 이내)
    ' 두 활동의 기간의 합은 최소한 전체 기간보다 커야한다.

    ' 활동수가 3 이면 (4 ~ 6개월)
    ' 1) 1번 활동은 시작으로 마지막 활동은 끝으로 설정한다
    ' 2) 1번과 2번 활동기간의 합은 전체 기간의 60%을 넘지 않는다. 
    ' 3) 중간의 활동은 확율적으로 1번이나 2번과 겹치는 기간을 갖는다.
    ' 
    ' 활동수가 4이면 (4 ~ 6개월)
    ' 1) 1번 활동은 시작으로 마지막 활동은 끝으로 설정한다
    ' 2) 1번과 2번 활동기간의 합은 전체 기간의 40%을 넘지 않는다. 
    ' 3) 중간의 활동은 확율적으로 1번이나 2번과 겹치는 기간을 갖는다.
    ' 4) 중간의 활동간에도 확율적으로 겹치는 기간을 갖는다.

    ' 활동별 duration 을 구한다.

    Dim problemSheetName As String
    Dim Arr()            As Variant

    Arr = Range(Cells(5,2),Cells(40,2))         ' Opt, ML, Pess 에 따라 선택 영역이 변하게 수정해야함.

    Dim fullArr(1 to 1440,1 to 12) As Variant   ' 1440 대신 범위 계산해서 넣어야함.
    Dim fullIndex               As Integer      ' fullArr배열을 순차적으로 찾아갈 변수

    fullIndex                   = 1         

    for i = 1 to 3 ' 문제갯수 만큼 시트를 만들고 시트에 기록한다.
        ' 시트 생성후 시트내 헤더 만들기
        for j = 1 to 3
            problemSheetName = MakeNewSheet(i,j) '001_Opt, 001_ML, 001_Pess, 생성한 시트이름을 리턴한다.


        Next
        

        ' = MakeActivity() ' 기간, 활동수 전달받아서 활동을 완성한다.
        'delete Array 'MakeActivity() 에서 받든 배열을 삭제한다. 

    Next

End Sub
 
Function MakeNewSheet(i As Integer, j As Integer)
    
    Dim strSheetName As String  ' 생성할 시트의 이름을 담을 문자열
    Dim strOption  

    strOption       = Array("Opt","ML","Press")    

    strSheetName    = Format(i,"#000")
    strSheetName    = strSheetName & "_" & strOption(j)
    
    Delete_Exist_Sheet(strSheetName)    ' 시트가 이미 존재하면 지운다.

    Sheets.Add(After:=Sheets(Sheets.Count)).Name = strSheetName '새로운 시트 생성
    
    ' 검증의 편리성을 위해서 일단 여기서 타이틀을 삽입한다.
    Dim strTile
    strTile = Array("빈도","시나리오","월","프로젝트","기간","활동수", _ 
                    "활동","활동시작","활동끝","r1","r2","r3")

    Dim sheetActtmp As Worksheet 
    set sheetActtmp = Worksheets(strSheetName)

    Dim arrSize As Integer
    arrSize = UBound(strTile) - Lbound(strTile) + 1

    With sheetActtmp
        .Range(.Cells(1,1),.Cells(1,arrSize)).Value  = strTile
    End With

    MakeNewSheet = strSheetName

End Function


Function Delete_Exist_Sheet(WorkSheet_Name As String)

    Dim Work_sheet As Worksheet

    For Each Work_sheet In ThisWorkbook.Worksheets

        If Work_sheet.Name = WorkSheet_Name Then        
            Application.DisplayAlerts   = False
            Worksheets(WorkSheet_Name).Delete
            Application.DisplayAlerts   = True
        End If

    Next

End Function




' 시작월과 종료월, 활동의 수가 주어지면
' 활동별 시작 끝 r1, r2, r3 를 생성하는 함수
Function MakeActivity(ByRef fullArr As Variant, fullIndex As Integer) 

    Dim i           As Integer  
    Dim j           As Integer
    
    Dim prjTime     As Integer ' 프로젝트 기간
    Dim actCnt      As Integer ' 활동수
    Dim month       As Integer ' 프로젝트 시작 달
    
    Dim actSTime    As Integer
    Dim actETime    As Integer
    
    prjTime = fullArr(fullIndex, 5)
    actCnt  = fullArr(fullIndex, 6)
    month   = fullArr(fullIndex, 3)
    
    If actCnt = 1 Then ' 활동이 1 이면 기간은 전체기간 으로 정해짐
        actSTime = month
        actETime = month + prjTime - 1    

    ElseIf actCnt = 2 Then ' 활동이 2개이면 시작과 끝으로 적당히 나눔
        If sheetPrj = 1 Then ' 1개월 짜리이면
            actSTime = month
            actETime = month + prjTime - 1    


    Select Case l
    Case 
    Case
    Case
    Case Else

    End Select
    MakeActivity = 0
End Function

Private Sub old_CommandButton1_Click()

    Dim strTile
    strTile = Array("빈도","시나리오","월","프로젝트","기간","활동수", _ 
                    "활동","활동시작","활동끝","r1","r2","r3")

    Dim sheetActtmp As Worksheet 
    set sheetActtmp = Worksheets("act_temp")

    With sheetActtmp
        .Range(.Cells(1,1),.Cells(1,12)).Value  = strTile
    End With

    Dim freq        As Integer
    Dim scenario    As Integer
    Dim month       As Integer
    Dim project     As Integer
    Dim prjTime     As Integer
    Dim actCnt      As Integer    
    Dim actNum      As Integer
    Dim actSTime    As Integer
    Dim actETime    As Integer
    Dim r1          As Integer
    Dim r2          As Integer
    Dim r3          As Integer

    freq            = 1
    scenario        = 2
    month           = 3
    project         = 4
    prjTime         = 5
    actCnt          = 6
    actNum        = 7
    actSTime        = 8
    actETime        = 9
    r1              = 10
    r2              = 11
    r3              = 12

    Dim h As Integer
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer

    ' 참고 : 36개월 * 5(최대프로젝트) * 8(최대활동) = 1440 개의 활동
    ' 1. 단기프로젝트 - 전체 프로젝트의 60%로 구성, 최소 1개월 ~ 최대 6개월 이내
    ' 프로젝트 별 활동수를 구한다.

    ' 활동의 최대수는 4개로 한다.

    ' 활동수가 1이면 적은 기간으로 설정 (2개월 이내)

    ' 활동수가 2이면 확율적으로 겹치게 한다. (4개월 이내)
    ' 두 활동의 기간의 합은 최소한 전체 기간보다 커야한다.

    ' 활동수가 3 이면 (4 ~ 6개월)
    ' 1) 1번 활동은 시작으로 마지막 활동은 끝으로 설정한다
    ' 2) 1번과 2번 활동기간의 합은 전체 기간의 60%을 넘지 않는다. 
    ' 3) 중간의 활동은 확율적으로 1번이나 2번과 겹치는 기간을 갖는다.
    ' 
    ' 활동수가 4이면 (4 ~ 6개월)
    ' 1) 1번 활동은 시작으로 마지막 활동은 끝으로 설정한다
    ' 2) 1번과 2번 활동기간의 합은 전체 기간의 40%을 넘지 않는다. 
    ' 3) 중간의 활동은 확율적으로 1번이나 2번과 겹치는 기간을 갖는다.
    ' 4) 중간의 활동간에도 확율적으로 겹치는 기간을 갖는다.

    ' 활동별 duration 을 구한다.

    ' Opt 1을 가지고 테스해 보자
    Dim Arr()                   As Variant
    Arr = Range(Cells(5,2),Cells(40,2))         ' Opt, ML, Pess 에 따라 선택 영역이 변하게 수정해야함.

    Dim fullArr(1 to 1440,1 to 12) As Variant   ' 1440 대신 범위 계산해서 넣어야함.
    Dim fullIndex               As Integer      ' fullArr배열을 순차적으로 찾아갈 변수
    fullIndex                   = 1             ' 

    h = 1                       ' frequency 번호 Opt=1, ML=2, Pess=3
    i = 1                       ' 시나리오번호 1 to ProblemCnt

    Dim prjTotalCont As Integer
    prjTotalCont = 0    ' 프로젝트 번호 
        
    For j = 1 To 36             ' 탐색할 총 개월수, 구해오게 변경해야함. 
        Dim kk                  As Integer
        Dim kkk                 As Integer
        Dim Cnt                 As Integer

        kk = Arr(j,1)
        
        For k=1 To kk           ' 프로젝트 루프
            prjTotalCont = prjTotalCont +1
            If Rnd < 0.2 Then           ' 장기프로젝트
                Cnt             = 6     ' 임시로 장기 프로젝트는 활동수 6개로 고정    
                kkk             = 10    ' 임시로 장기프로젝트 기간은 10개월로 고정

            Else '단기 프로젝트
                kkk             = (Rnd * 100) ' 단기 프로젝트 활동의 수는 1~4개 사이로
                kkk             = (kkk mod 4) + 1
                Cnt             = kkk

                kkk             = (Rnd * 100)   ' 단기 프로젝트 총 기간은 1~6개월 사이로
                kkk             = (kkk mod 6) + 1
                
            End If

            Dim l As Integer                
            For l = 1 To Cnt ' 활동의 갯수만큼 루프를 돈다.

                fullArr(fullIndex, freq)        = h
                fullArr(fullIndex, scenario)    = i ' 공란으로 남겨 놓으려면 "" 입력
                fullArr(fullIndex, month)       = j
                fullArr(fullIndex, project)     = prjTotalCont
                fullArr(fullIndex, prjTime)     = kkk
                fullArr(fullIndex, actCnt)      = Cnt
                fullArr(fullIndex, actNum)      = l ' 여기까지 작성함.
                fullArr(fullIndex, actSTime)    = l '
                fullArr(fullIndex, actETime)    = l '
                fullArr(fullIndex, r1)          = "" '
                fullArr(fullIndex, r2)          = l '
                fullArr(fullIndex, r3)          = l '

                fullIndex = fullIndex + 1
            Next
            
        Next
    Next

    With sheetActtmp
        .Range(.Cells(2,1),.Cells(1441,12)).Value = fullArr
    End With

End Sub

Private Sub old2_CommandButton1_Click()

    Dim strTile
    strTile = Array("빈도","시나리오","월","프로젝트","기간","활동수", _ 
                    "활동","활동시작","활동끝","r1","r2","r3")

    Dim sheetActtmp As Worksheet 
    set sheetActtmp = Worksheets("act_temp")

    With sheetActtmp
        .Range(.Cells(1,1),.Cells(1,12)).Value  = strTile
    End With

    Dim freq        As Integer
    Dim scenario    As Integer
    Dim month       As Integer
    Dim project     As Integer
    Dim prjTime     As Integer
    Dim actCnt      As Integer    
    Dim actNum      As Integer
    Dim actSTime    As Integer
    Dim actETime    As Integer
    Dim r1          As Integer
    Dim r2          As Integer
    Dim r3          As Integer

    freq            = 1
    scenario        = 2
    month           = 3
    project         = 4
    prjTime         = 5
    actCnt          = 6
    actNum          = 7
    actSTime        = 8
    actETime        = 9
    r1              = 10
    r2              = 11
    r3              = 12

    Dim h As Integer
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer

    ' 참고 : 36개월 * 5(최대프로젝트) * 8(최대활동) = 1440 개의 활동
    ' 1. 단기프로젝트 - 전체 프로젝트의 60%로 구성, 최소 1개월 ~ 최대 6개월 이내
    ' 프로젝트 별 활동수를 구한다.

    ' 활동의 최대수는 4개로 한다.

    ' 활동수가 1이면 적은 기간으로 설정 (2개월 이내)

    ' 활동수가 2이면 확율적으로 겹치게 한다. (4개월 이내)
    ' 두 활동의 기간의 합은 최소한 전체 기간보다 커야한다.

    ' 활동수가 3 이면 (4 ~ 6개월)
    ' 1) 1번 활동은 시작으로 마지막 활동은 끝으로 설정한다
    ' 2) 1번과 2번 활동기간의 합은 전체 기간의 60%을 넘지 않는다. 
    ' 3) 중간의 활동은 확율적으로 1번이나 2번과 겹치는 기간을 갖는다.
    ' 
    ' 활동수가 4이면 (4 ~ 6개월)
    ' 1) 1번 활동은 시작으로 마지막 활동은 끝으로 설정한다
    ' 2) 1번과 2번 활동기간의 합은 전체 기간의 40%을 넘지 않는다. 
    ' 3) 중간의 활동은 확율적으로 1번이나 2번과 겹치는 기간을 갖는다.
    ' 4) 중간의 활동간에도 확율적으로 겹치는 기간을 갖는다.

    ' 활동별 duration 을 구한다.

    Dim problemSheetName As String
    Dim Arr()            As Variant

    Arr = Range(Cells(5,2),Cells(40,2))         ' Opt, ML, Pess 에 따라 선택 영역이 변하게 수정해야함.

    Dim fullArr(1 to 1440,1 to 12) As Variant   ' 1440 대신 범위 계산해서 넣어야함.
    Dim fullIndex               As Integer      ' fullArr배열을 순차적으로 찾아갈 변수

    fullIndex                   = 1         

    for i = 1 to 3 ' 문제갯수 만큼 시트를 만들고 시트에 기록한다.
        ' 시트 생성후 시트내 헤더 만들기
        for j = 1 to 3
            problemSheetName = MakeNewSheet(i,j) '001_Opt, 001_ML, 001_Pess, 생성한 시트이름을 리턴한다.
            

        Next
        

        ' = MakeActivity() ' 기간, 활동수 전달받아서 활동을 완성한다.
        'delete Array 'MakeActivity() 에서 받든 배열을 삭제한다. 

    Next

End Sub
 

