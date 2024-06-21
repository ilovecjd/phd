Option Explicit
Option Base 1

' 대시보드를 생성한다. 
Private Sub GenDBoard_Click()

	LetExcelInitialized = 0	' 새로운 프로젝트들을 생성하기 위해 초기화 플래그 변경
	LetTableInitialized = 0	' 새로운 프로젝트들을 생성하기 위해 초기화 플래그 변경

    Call Prologue(0)		' 전체 파라메터 로드-> 대시보드 생성 -> 프로젝트 생성
	Call PrintDashboard()	' 대시보드를 시트에 출력한다
	Call PrintProjectHeader()		' 프로젝트를 시트에 출력한다
	Call PrintProjectAll()
	
	' Call CreateDashboard()     ' 대시보드를 생성하고 전체 프로젝트의 갯수를 구한다.
	' Call Epilogue()

End Sub





