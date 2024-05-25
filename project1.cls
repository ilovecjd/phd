Option Explicit

Option Base 1

Private Sub CommandButton1_Click()

    Dim strTile
    strTile = Array("��", "�ó�����", "��", "������Ʈ", "�Ⱓ", "Ȱ����", _
                    "Ȱ��", "Ȱ������", "Ȱ����", "r1", "r2", "r3")

    Dim sheetActtmp As Worksheet
    Set sheetActtmp = Worksheets("act_temp")

    With sheetActtmp
        .Range(.Cells(1, 1), .Cells(1, 12)).Value = strTile
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

    freq = 1
    scenario = 2
    month = 3
    project = 4
    prjTime = 5
    actCnt = 6
    actNum = 7
    actSTime = 8
    actETime = 9
    r1 = 10
    r2 = 11
    r3 = 12

    Dim h As Integer
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer

    ' ���� : 36���� * 5(�ִ�������Ʈ) * 8(�ִ�Ȱ��) = 1440 ���� Ȱ��
    ' 1. �ܱ�������Ʈ - ��ü ������Ʈ�� 60%�� ����, �ּ� 1���� ~ �ִ� 6���� �̳�
    ' ������Ʈ �� Ȱ������ ���Ѵ�.

    ' Ȱ���� �ִ���� 4���� �Ѵ�.

    ' Ȱ������ 1�̸� ���� �Ⱓ���� ���� (2���� �̳�)

    ' Ȱ������ 2�̸� Ȯ�������� ��ġ�� �Ѵ�. (4���� �̳�)
    ' �� Ȱ���� �Ⱓ�� ���� �ּ��� ��ü �Ⱓ���� Ŀ���Ѵ�.

    ' Ȱ������ 3 �̸� (4 ~ 6����)
    ' 1) 1�� Ȱ���� �������� ������ Ȱ���� ������ �����Ѵ�
    ' 2) 1���� 2�� Ȱ���Ⱓ�� ���� ��ü �Ⱓ�� 60%�� ���� �ʴ´�.
    ' 3) �߰��� Ȱ���� Ȯ�������� 1���̳� 2���� ��ġ�� �Ⱓ�� ���´�.
    '
    ' Ȱ������ 4�̸� (4 ~ 6����)
    ' 1) 1�� Ȱ���� �������� ������ Ȱ���� ������ �����Ѵ�
    ' 2) 1���� 2�� Ȱ���Ⱓ�� ���� ��ü �Ⱓ�� 40%�� ���� �ʴ´�.
    ' 3) �߰��� Ȱ���� Ȯ�������� 1���̳� 2���� ��ġ�� �Ⱓ�� ���´�.
    ' 4) �߰��� Ȱ�������� Ȯ�������� ��ġ�� �Ⱓ�� ���´�.

    ' Ȱ���� duration �� ���Ѵ�.

    Dim problemSheetName As String

    For i = 1 To 3 ' �������� ��ŭ ��Ʈ�� ����� ��Ʈ�� ����Ѵ�.

        ' ��Ʈ ������ ��Ʈ�� ��� �����
        For j = 1 To 3
            problemSheetName = MakeNewSheet(i, j) '001_Opt, 001_ML, 001_Pess, ������ ��Ʈ�̸��� �����Ѵ�.
        Next
        

        ' = MakeActivity() ' �Ⱓ, Ȱ���� ���޹޾Ƽ� Ȱ���� �ϼ��Ѵ�.
        'delete Array 'MakeActivity() ���� �޵� �迭�� �����Ѵ�.

    Next

End Sub
 
Function MakeNewSheet(i As Integer, j As Integer)
    
    Dim strSheetName As String  ' ������ ��Ʈ�� �̸��� ���� ���ڿ�
    Dim strOption

    strOption = Array("Opt", "ML", "Press")

    strSheetName = Format(i, "#000")
    strSheetName = strSheetName & "_" & strOption(j)
    
    Delete_Exist_Sheet (strSheetName)   ' ��Ʈ�� �̹� �����ϸ� �����.

    Sheets.Add(After:=Sheets(Sheets.Count)).Name = strSheetName '���ο� ��Ʈ ����
    
    '' ������ �������� ���ؼ� �ϴ� ���⼭ Ÿ��Ʋ�� �����Ѵ�.
    Dim strTile
    strTile = Array("��", "�ó�����", "��", "������Ʈ", "�Ⱓ", "Ȱ����", _
                    "Ȱ��", "Ȱ������", "Ȱ����", "r1", "r2", "r3")

    Dim sheetActtmp As Worksheet
    Set sheetActtmp = Worksheets(strSheetName)

    Dim arrSize As Integer
    arrSize = UBound(strTile) - LBound(strTile) + 1

    With sheetActtmp
        .Range(.Cells(1, 1), .Cells(1, arrSize)).Value = strTile
    End With

    MakeNewSheet = strSheetName

End Function


Function Delete_Exist_Sheet(WorkSheet_Name As String)

    Dim Work_sheet As Worksheet

    For Each Work_sheet In ThisWorkbook.Worksheets

        If Work_sheet.Name = WorkSheet_Name Then
            Application.DisplayAlerts = False
            Worksheets(WorkSheet_Name).Delete
            Application.DisplayAlerts = True
        End If

    Next

End Function




' ���ۿ��� �����, Ȱ���� ���� �־�����
' Ȱ���� ���� �� r1, r2, r3 �� �����ϴ� �Լ�
Function MakeActivity(ByRef fullArr As Variant, fullIndex As Integer)

    Dim i           As Integer
    Dim j           As Integer
    
    Dim prjTime     As Integer ' ������Ʈ �Ⱓ
    Dim actCnt      As Integer ' Ȱ����
    Dim month       As Integer ' ������Ʈ ���� ��
    
    Dim actSTime    As Integer
    Dim actETime    As Integer
    
    prjTime = fullArr(fullIndex, 5)
    actCnt = fullArr(fullIndex, 6)
    month = fullArr(fullIndex, 3)
    
    If actCnt = 1 Then ' Ȱ���� 1 �̸� �Ⱓ�� ��ü�Ⱓ ���� ������
        actSTime = month
        actETime = month + prjTime - 1

    ElseIf actCnt = 2 Then ' Ȱ���� 2���̸� ���۰� ������ ������ ����
        If sheetPrj = 1 Then ' 1���� ¥���̸�
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
    strTile = Array("��", "�ó�����", "��", "������Ʈ", "�Ⱓ", "Ȱ����", _
                    "Ȱ��", "Ȱ������", "Ȱ����", "r1", "r2", "r3")

    Dim sheetActtmp As Worksheet
    Set sheetActtmp = Worksheets("act_temp")

    With sheetActtmp
        .Range(.Cells(1, 1), .Cells(1, 12)).Value = strTile
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

    freq = 1
    scenario = 2
    month = 3
    project = 4
    prjTime = 5
    actCnt = 6
    actNum = 7
    actSTime = 8
    actETime = 9
    r1 = 10
    r2 = 11
    r3 = 12

    Dim h As Integer
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer

    ' ���� : 36���� * 5(�ִ�������Ʈ) * 8(�ִ�Ȱ��) = 1440 ���� Ȱ��
    ' 1. �ܱ�������Ʈ - ��ü ������Ʈ�� 60%�� ����, �ּ� 1���� ~ �ִ� 6���� �̳�
    ' ������Ʈ �� Ȱ������ ���Ѵ�.

    ' Ȱ���� �ִ���� 4���� �Ѵ�.

    ' Ȱ������ 1�̸� ���� �Ⱓ���� ���� (2���� �̳�)

    ' Ȱ������ 2�̸� Ȯ�������� ��ġ�� �Ѵ�. (4���� �̳�)
    ' �� Ȱ���� �Ⱓ�� ���� �ּ��� ��ü �Ⱓ���� Ŀ���Ѵ�.

    ' Ȱ������ 3 �̸� (4 ~ 6����)
    ' 1) 1�� Ȱ���� �������� ������ Ȱ���� ������ �����Ѵ�
    ' 2) 1���� 2�� Ȱ���Ⱓ�� ���� ��ü �Ⱓ�� 60%�� ���� �ʴ´�.
    ' 3) �߰��� Ȱ���� Ȯ�������� 1���̳� 2���� ��ġ�� �Ⱓ�� ���´�.
    '
    ' Ȱ������ 4�̸� (4 ~ 6����)
    ' 1) 1�� Ȱ���� �������� ������ Ȱ���� ������ �����Ѵ�
    ' 2) 1���� 2�� Ȱ���Ⱓ�� ���� ��ü �Ⱓ�� 40%�� ���� �ʴ´�.
    ' 3) �߰��� Ȱ���� Ȯ�������� 1���̳� 2���� ��ġ�� �Ⱓ�� ���´�.
    ' 4) �߰��� Ȱ�������� Ȯ�������� ��ġ�� �Ⱓ�� ���´�.

    ' Ȱ���� duration �� ���Ѵ�.

    ' Opt 1�� ������ �׽��� ����
    Dim Arr()                   As Variant
    Arr = Range(Cells(5, 2), Cells(40, 2))      ' Opt, ML, Pess �� ���� ���� ������ ���ϰ� �����ؾ���.

    Dim fullArr(1 To 1440, 1 To 12) As Variant  ' 1440 ��� ���� ����ؼ� �־����.
    Dim fullIndex               As Integer      ' fullArr�迭�� ���������� ã�ư� ����
    fullIndex = 1                               '

    h = 1                       ' frequency ��ȣ Opt=1, ML=2, Pess=3
    i = 1                       ' �ó�������ȣ 1 to ProblemCnt

    Dim prjTotalCont As Integer
    prjTotalCont = 0    ' ������Ʈ ��ȣ
        
    For j = 1 To 36             ' Ž���� �� ������, ���ؿ��� �����ؾ���.
        Dim kk                  As Integer
        Dim kkk                 As Integer
        Dim Cnt                 As Integer

        kk = Arr(j, 1)
        
        For k = 1 To kk         ' ������Ʈ ����
            prjTotalCont = prjTotalCont + 1
            If Rnd < 0.2 Then           ' ���������Ʈ
                Cnt = 6                 ' �ӽ÷� ��� ������Ʈ�� Ȱ���� 6���� ����
                kkk = 10                ' �ӽ÷� ���������Ʈ �Ⱓ�� 10������ ����

            Else '�ܱ� ������Ʈ
                kkk = (Rnd * 100)             ' �ܱ� ������Ʈ Ȱ���� ���� 1~4�� ���̷�
                kkk = (kkk Mod 4) + 1
                Cnt = kkk

                kkk = (Rnd * 100)               ' �ܱ� ������Ʈ �� �Ⱓ�� 1~6���� ���̷�
                kkk = (kkk Mod 6) + 1
                
            End If

            Dim l As Integer
            For l = 1 To Cnt ' Ȱ���� ������ŭ ������ ����.

                fullArr(fullIndex, freq) = h
                fullArr(fullIndex, scenario) = i    ' �������� ���� �������� "" �Է�
                fullArr(fullIndex, month) = j
                fullArr(fullIndex, project) = prjTotalCont
                fullArr(fullIndex, prjTime) = kkk
                fullArr(fullIndex, actCnt) = Cnt
                fullArr(fullIndex, actNum) = l      ' ������� �ۼ���.
                fullArr(fullIndex, actSTime) = l    '
                fullArr(fullIndex, actETime) = l    '
                fullArr(fullIndex, r1) = ""          '
                fullArr(fullIndex, r2) = l          '
                fullArr(fullIndex, r3) = l          '

                fullIndex = fullIndex + 1
            Next
            
        Next
    Next

    With sheetActtmp
        .Range(.Cells(2, 1), .Cells(1441, 12)).Value = fullArr
    End With

End Sub




