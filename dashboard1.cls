Option Explicit

'' ��Ʈ���� ���� �̿��ؼ� ������Ʈ�� ������ ���� ���Ѻ���.
Private Sub CommandButton1_Click()

    ' ���� ���� ����
    ' ���� ���� ����������Ʈ Ȯ��
    'GenerateCombinations
    'CirculateMonth ' ���� ���� ������Ʈ ��Ȳ ǥ�� �����Ѵ�.
    'Call BuildPrj(10)

    Dim ws              As Worksheet
    Dim month           As Integer
    Dim i               As Integer
    Dim j               As Integer
    Dim endPrjNum       As Integer
    Dim preTotal        As Integer
    Dim prjNum          As Integer
    Dim arrAllPrj()     As clsProject
    Dim totalPrjNum As Integer

    Set ws = ThisWorkbook.Sheets("dashboard")              ' ������ ��Ʈ �̸����� ����
    totalPrjNum = ws.Cells(3, 37).Value
    ReDim arrAllPrj(1 To totalPrjNum)

    

    ' For i = 1 to totalPrjNum
    '   Set arrAllPrj(i) = New clsProject
    '   Call arrAllPrj(i).Init(0, mon, "project2",2)    'Init(initType As Integer, initNum As Integer, sheetName As String)
    '   Call arrAllPrj(i).PrintInfo()
    ' Next

    For i = 2 To 37
        month = ws.Cells(2, i)       ' ��
        endPrjNum = ws.Cells(3, i)   ' �̹��ޱ��� ����� ������ ��ȣ

        If i = 2 Then ' ó���̸� �ʱⰪ��
            preTotal = 0   ' ��������
        Else
            preTotal = ws.Cells(3, i - 1) ' ��������
        End If

        prjNum = preTotal + 1 ' ����������Ʈ ��ȣ

        If prjNum = 0 Then
            GoTo Continue
        End If

        If prjNum > endPrjNum Then
            GoTo Continue
        End If

        For j = prjNum To endPrjNum '
            Set arrAllPrj(j) = New clsProject
            Call arrAllPrj(j).Init(0, j, "project2", month)   'Init(initType As Integer, initNum As Integer, sheetName As String)
            Call arrAllPrj(j).PrintInfo
        Next


Continue:


    Next

    'Dim cProject As New clsProject
    ' Call cProject.Init(0, 1, "project2")    'Init(initType As Integer, initNum As Integer, sheetName As String)
    ' Call cProject.PrintInfo()

  
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' ������Ʈ�� �⺻ ������ �����Ѵ�.
'' @prjArr : ������Ʈ ������ ��Ƽ� ������ ����� �迭
'' @prjNum : ������Ʈ ��ȣ
Function buildPrjInfo(prjArr() As Variant, prjNum As Integer)
   
    ' ù ���� �����ش�.
    Dim activityCnt As Integer
    Dim iRow As Integer ' ������Ʈ�� ��ġ�� �迭�� ù �� ��ȣ

    iRow = (prjNum - 1) * SIZE_OF_PROJECT + 2 'song ���� ó�� �ʿ� (iRow�� 1(����) ���� �������)

    prjArr(iRow, 1) = "����"
    prjArr(iRow, 2) = prjNum

    prjArr(iRow + 1, 1) = "�Ⱓ"
    prjArr(iRow + 1, 2) = Int(Rnd() * 100 Mod 12) + 1 ' �Ⱓ�� 12���� �̳�

    prjArr(iRow + 2, 1) = "ActivityCnt"
    activityCnt = Int(Rnd() * 100 Mod 6) + 1 ' song ������ ���� ����
    prjArr(iRow + 2, 2) = activityCnt

End Function

Function BuildPrj(prjCnt As Integer)

    Dim project As typeProject
    Dim arrPrj() As typeProject
    ReDim rrPrj(1 To prjCnt)

    Dim i As Integer
    For i = 1 To prjCnt
        project = arrPrj(i)
        project.num = i
        period  =
        OrderDate
        possibleDate
        StartDate
        Profit
        Experience
        successPercentage
        CF
        N_CF
        
    Next


    
End Function
Sub GenerateCombinations()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("dashboard") ' ������ ��Ʈ �̸����� ����
    
    Dim lastCol         As Integer
    lastCol = ws.Cells(2, ws.Columns.Count).end(xlToLeft).Column ' ������ �� ã��
    
    Dim periods()       As Integer
    Dim accumulate()    As Integer ' ������ ������������ ������ �־� �д�.
    Dim counts()        As Integer
    Dim i               As Integer
    
    ' �Ⱓ�� ������ �迭�� ����
    ReDim periods(1 To lastCol) 'song lastCol�� ���� �������� �������°ɷ� ��������
    ReDim accumulate(1 To lastCol)
    ReDim counts(1 To lastCol)
    
    ' �����ʹ� 2��° ������ �о��
    For i = 1 To lastCol - 1 ' ��� �κ��� ���� �о� �´�
        periods(i) = ws.Cells(2, i + 1).Value
        accumulate(i) = ws.Cells(3, i + 1).Value
        counts(i) = ws.Cells(4, i + 1).Value
    Next i
    
    ' ��� ���� ����
    Dim ePrjNum As Integer ' ������ ������Ʈ ��ȣ
    ePrjNum = accumulate(1)

    Dim result()          As Variant
    ReDim result(1 To lastCol)
    For i = 1 To ePrjNum
        ' 1���÷��� �ܺ� �������� ���� 2��° �÷����� ����
        result(1) = i
        GenerateCombinationRecursive periods, accumulate, counts, result, 2, lastCol, i
    Next
End Sub
    
    
' 1st loop : 1 to count(col) ==> �ܺο��� �����ϰ� ���ȣ�� �Լ��� 2�� col ����
' other loop : acc(col) to acc(col) + (count(col)-1) ==> ������Ʈ ��ȣ
'
' exit condition : count(col) = prjNum
Sub GenerateCombinationRecursive(periods() As Integer, _
                                accumulate() As Integer, _
                                counts() As Integer, _
                                result() As Variant, _
                                col As Integer, _
                                lastCol As Integer, _
                                prjNum As Integer)
    
    ' exit condition : count(col) = prjNum
    Dim i As Integer
    Dim sPrjNum As Integer ' ���� ������Ʈ ��ȣ
    Dim ePrjNum As Integer ' ������ ������Ʈ ��ȣ

    If col > (lastCol - 1) Then ' check exit condition
        ' ��� ���
        For i = 1 To lastCol
            Debug.Print result(i);
        Next i
        Debug.Print
        Exit Sub
    End If
    
    ' song �����̳� �߻��� ������Ʈ ������ 0 �ΰ�� ����ó�� �ʿ�
    If 0 = counts(col) Then
        result(col) = 0 ' ������Ʈ ��ȣ
        col = col + 1
        GenerateCombinationRecursive periods, accumulate, counts, result, col + 1, lastCol, i
    End If
    
    sPrjNum = accumulate(col - 1) + 1 '���� ���� ���� ��ȣ�� �̹��� �߻��� ������Ʈ ���� ����
    ePrjNum = accumulate(col) ' ���� ������ �̹� �߻��� ������Ʈ�� ������ ���� + (counts(col)-1) '

    For i = sPrjNum To ePrjNum
        result(col) = i ' ������Ʈ ��ȣ
        GenerateCombinationRecursive periods, accumulate, counts, result, col + 1, lastCol, i
    Next

End Sub


' ������ �������� ���ƾ� �Ѵ�. �Ф�
' ������ ������Ʈ�� Ȯ���ϰ�
' ������Ʈ���� ���ÿ��� ����
'1. �̹��� �����η� ���
'2. �̹��� ���� �ڱ� ��� ( �̿� + ���� ? ����)
'3. �������� ������Ʈ�� �� �Ϸ�� ������Ʈ�� �ִ��� �˻�
'4. �̹��� �ű� ������ ������Ʈ ����
'   1)  �ű� ���࿡ �ο��� ���������� ������?
'   2)

' Sub CirculateMonth()
' Sub UpdateProject
' Sub UpdateHR
' Sub UpdateCash
' Sub SelectProject


'
    'loop 1 to 36 ����
        'loop sPrjNum  to ePrjNum  �̹��� ���� ������Ʈ ���� �̹��� ������ ������Ʈ����
            ' UpdateProject
            ' UpdateHR
            ' UpdateCash
            ' SelectProject

Sub CirculateMonth()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("dashboard") ' ������ ��Ʈ �̸����� ����

    Dim lastCol         As Integer
    lastCol = ws.Cells(2, ws.Columns.Count).end(xlToLeft).Column ' ������ �� ã��
    
    Dim periods()       As Integer
    Dim accumulate()    As Integer ' ������ ������������ ������ �־� �д�.
    Dim counts()        As Integer
    
    ' �Ⱓ�� ������ �迭�� ����
    ReDim periods(0 To lastCol) 'song lastCol�� ���� �������� �������°ɷ� ��������
    ReDim accumulate(0 To lastCol)
    ReDim counts(0 To lastCol)
    
    Dim i               As Integer
    ' �����ʹ� 2��° ������ �о��
    For i = 1 To lastCol ' ��� �κ��� ���� �о� �´�
        periods(i) = ws.Cells(2, i + 1).Value
        accumulate(i) = ws.Cells(3, i + 1).Value
        counts(i) = ws.Cells(4, i + 1).Value
    Next i
    
    Dim currentMonth As Integer
    Dim sPrjNum As Integer ' ���� ������Ʈ ��ȣ
    Dim ePrjNum As Integer ' ������ ������Ʈ ��ȣ
    Dim prjCount As Integer

    periods(0) = 0
    accumulate(0) = 0
    counts(0) = 0

    For i = 1 To lastCol ' ��� �Ⱓ ���� (36����)
        currentMonth = periods(i)
        If i = 1 Then
            sPrjNum = 1
        Else
            sPrjNum = accumulate(i - 1) + 1       '���� ���� ���� ��ȣ�� �̹��� �߻��� ������Ʈ ���� ����
        End If
        'sPrjNum         = accumulate(i) + 1 '���� ���� ���� ��ȣ�� �̹��� �߻��� ������Ʈ ���� ����
        ePrjNum = accumulate(i)         ' ���� ������ �̹� �߻��� ������Ʈ�� ������ ���� + (counts(col)-1) '
        prjCount = counts(i)

        ' Debug.Print currentMonth;
        ' Debug.Print sPrjNum;
        ' Debug.Print ePrjNum;
        ' Debug.Print prjCount;
        ' Debug.Print

        Dim j               As Integer
        Dim prjNum          As Integer
        Dim sumOfrunning    As Integer

        sumOfrunning = 0

        For j = 1 To prjCount

            Dim rand As Integer

            prjNum = sPrjNum + j - 1
            Debug.Print prjNum

            rand = Int(Rnd() * 100)

            If rand > 30 Then
                ws.Cells(14 + sumOfrunning, i + 1) = prjNum
                sumOfrunning = sumOfrunning + 1
            End If

            
        Next

        ws.Cells(13, i + 1) = sumOfrunning


    Next

End Sub

' �̹��� �������� ������Ʈ�� �ִ°�?
' �̹��� �������� ������Ʈ�� �Ϸ�� ������Ʈ�� �ִ°�?
' �Ϸ�� ������Ʈ�� ������ �ο� ���̺�, �������� ������Ʈ
Sub UpdateProject(month As Integer)


    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("dashboard")

    Dim sumOfrunning As Integer
    Dim i As Integer
    
    sumOfrunning = ws.Cells(13, month + 1)

    If sumOfrunning = 0 Then

    End If

    

End Sub



