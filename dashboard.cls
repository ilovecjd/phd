Option Explicit

'' 시트내의 값을 이용해서 프로젝트를 가상을 진행 시켜본다.
Private Sub CommandButton1_Click()

    ' 전역 변수 설정
    ' 월별 선택 가능프로젝트 확인
    ' 
    GenerateCombinations
End Sub



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
    
    
    For i = 1 To lastCol
        periods(i)      = ws.Cells(2, i).Value
        accumulate(i)   = ws.Cells(3, i).Value
        counts(i)       = ws.Cells(4, i).Value
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

    If col > lastCol Then ' check exit condition
        ' 결과 출력
        For i = 1 To lastCol
            Debug.Print result(i);
        Next i
        Debug.Print
        Exit Sub
    End If
    
    ' song 누적이나 발생한 프로젝트 갯수가 0 인경우 예외처리 필요
    sPrjNum = accumulate(col-1) + 1 '이전 누적 다음 번호가 이번에 발생한 프로젝트 시작 순번
    ePrjNum = accumulate(col) ' 현재 누적은 이번 발생한 프로젝트의 마지막 순번 + (counts(col)-1) ' 


    For i = sPrjNum to ePrjNum
        result(col) = i ' 프로젝트 번호
        GenerateCombinationRecursive periods, accumulate, counts, result, col + 1, lastCol, i        
    Next

    ' For i = 0 To counts(col)
    '     result(col) = prjNum 'periods(col) + i
    '     GenerateCombinationRecursive periods, accumulate, counts, result, col + 1, lastCol, i
    ' Next 

End Sub