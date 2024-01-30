Option Explicit

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
    Dim activity    As Integer
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
    activity        = 7
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

    '  1이면 적은 기간으로 설정 (2개월 이내)

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
    Arr = Range(Cells(5,2),Cells(40,2))

    Dim fullArr(1 to 1440,1 to 12) As Integer
    Dim fullIndex               As Integer
    fullIndex                   = 1

    h = 1                       ' frequency 번호 Opt, ML, Pess
    i = 1                       ' 시나리오번호
    For j = 1 To 36             ' 탐색할 총 개월수
        Dim kk                  As Integer
        Dim kkk                 As Integer
        Dim Cnt                 As Integer

        kk = Arr(j,1)

        For k=1 To kk           ' 프로젝트 루프

            If Rnd < 0.2 Then   ' 장기프로젝트
                Cnt           = 6      ' 임시로 장기 프로젝트는 활동수 6개로 고정    
                prjTime       = 10 ' 임시로 장기프로젝트 기간은 10개월로 고정

            Else '단기 프로젝트
                kkk             = (Rnd * 100) ' 프로젝트 갯수
                kkk             = (kkk mod 4) + 1
                Cnt             = kkk

                kkk             = (Rnd * 100)   ' 프로젝트 총 기간
                kkk             = (kkk mod 6) + 1
                
            End If

            Dim l As Integer                
            For l = 1 To Cnt ' 활동의 갯수만큼 루프를 돈다.
                ' Dim l2 As Integer
                ' If Cnt = 1 Then ' 활동이 1 이면 기간은 전체기간 으로 정해짐
                '     actSTime = j ' month
                '     actETime = j + kkk - 1'month + prjTime -1

                ' ElseIf Cnt = 2 Then ' 활동이 2개이면 시작과 끝으로 적당히 나눔
                '     Select Case
                '     Case 
                '     Case
                '     Case
                '     Case Else
                '     End Select
                
                


                fullArr(fullIndex, actCnt) = Cnt
                fullArr(fullIndex, activity) = l
                fullArr(fullIndex, project) = k
                fullArr(fullIndex, month) = j
                fullArr(fullIndex, scenario) = i
                fullArr(fullIndex, freq) = h
                fullArr(fullIndex, prjTime) = kkk

                fullIndex = fullIndex + 1
            Next
            
        Next
    Next



    With sheetActtmp
        .Range(.Cells(2,1),.Cells(1441,12)).Value = fullArr
    End With



End Sub
