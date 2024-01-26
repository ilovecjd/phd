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
' 프로젝트 클래스

  Public Dim index       As Integer
  Public Dim type        As Integer
  Public Dim orderMonth  As Integer
  Public Dim startMontj  As Integer
  Public Dim status      As Integer
  Public Dim actCnt      As Integer
  
  Public Sub Init
      msgbox "Init"
  End Sub
