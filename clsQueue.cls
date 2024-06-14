VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsQueue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' clsQueue 클래스 모듈
Option Explicit
Option Base 1

Private Const MAX_QUEUE_SIZE As Integer = 255
Private Queue(1 To MAX_QUEUE_SIZE) As clsProject
Private Head As Integer
Private Tail As Integer
Private Count As Integer

' 클래스 초기화
Private Sub Class_Initialize()
    Head = 1
    Tail = 1
    Count = 0
End Sub

' 큐에 객체 추가
Public Sub Enqueue(item As clsProject)
    If Count < MAX_QUEUE_SIZE Then
        Set Queue(Tail) = item
        Tail = (Tail Mod MAX_QUEUE_SIZE) + 1
        Count = Count + 1
    Else
        MsgBox "Queue is full", vbExclamation
    End If
End Sub

' 큐에서 객체 제거 및 반환
Public Function Dequeue() As clsProject
    If Count = 0 Then
        Set Dequeue = Nothing
    Else
        Set Dequeue = Queue(Head)
        Head = (Head Mod MAX_QUEUE_SIZE) + 1
        Count = Count - 1
    End If
End Function

' 큐에서 특정 객체 검색
' Public Function SearchByID(id As Integer) As clsProject
'     Dim i As Integer
'     Dim index As Integer
    
'     For i = 0 To Count - 1
'         index = ((Head + i - 1) Mod MAX_QUEUE_SIZE) + 1
'         If Queue(index).ProjectNum = id Then
'             Set SearchByID = Queue(index)
'             Exit Function
'         End If
'     Next i
'     Set SearchByID = Nothing
' End Function


Public Function SearchByID(id As Integer) As clsProject
    Dim i As Integer
    Dim index As Integer
    
    ' 기본값 설정
    Set SearchByID = Nothing

    For i = 0 To Count - 1
        index = ((Head + i - 1) Mod MAX_QUEUE_SIZE) + 1
        If Not Queue(index) Is Nothing Then ' Queue(index)가 Nothing이 아닌지 확인
            If Queue(index).ProjectNum = id Then
                Set SearchByID = Queue(index)
                Exit Function
            End If
        End If
    Next i
End Function


Public Function RemoveAndReturnByID(id As Integer) As clsProject
    Dim i As Integer
    Dim index As Integer
    Dim found As Boolean
    
    found = False
    
    For i = 0 To Count - 1
        index = ((Head + i - 1) Mod MAX_QUEUE_SIZE) + 1
        If Queue(index).ProjectNum = id Then
            found = True
            Exit For
        End If
    Next i
    
    If found Then
        ' 저장된 객체를 반환
        Set RemoveAndReturnByID = Queue(index)
        
        ' 객체 제거 및 나머지 객체 이동
        For i = index To (Tail - 2) Mod MAX_QUEUE_SIZE + 1
            Set Queue(i) = Queue((i Mod MAX_QUEUE_SIZE) + 1)
        Next i
        Tail = Tail - 1
        If Tail < 1 Then Tail = MAX_QUEUE_SIZE
        Count = Count - 1
    Else
        Set RemoveAndReturnByID = Nothing
        MsgBox "Project with ID " & id & " not found in the queue", vbExclamation
    End If
End Function

Public Sub RemoveByID(id As Integer)
    Dim i As Integer
    Dim index As Integer
    Dim found As Boolean
    
    found = False
    
    For i = 0 To Count - 1
        index = ((Head + i - 1) Mod MAX_QUEUE_SIZE) + 1
        If Queue(index).ProjectNum = id Then
            found = True
            Exit For
        End If
    Next i
    
    If found Then
        ' 객체 제거 및 나머지 객체 이동
        For i = index To (Tail - 2) Mod MAX_QUEUE_SIZE + 1
            Set Queue(i) = Queue((i Mod MAX_QUEUE_SIZE) + 1)
        Next i
        Tail = Tail - 1
        If Tail < 1 Then Tail = MAX_QUEUE_SIZE
        Count = Count - 1
    Else
        MsgBox "Project with ID " & id & " not found in the queue", vbExclamation
    End If
End Sub

' 큐의 현재 크기 반환
Public Function GetCount() As Integer
    GetCount = Count
End Function