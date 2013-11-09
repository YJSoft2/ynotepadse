Attribute VB_Name = "modProcess"
'---------------------------------------------------------------------------------------
' Module    : modProcess
' DateTime  : 2013-04-03 13:36
' Author    : YJSoft
' Purpose   :
'---------------------------------------------------------------------------------------
Public UpdateSite As String
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Private Declare Function Process32First Lib "kernel32" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "kernel32" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Sub CloseHandle Lib "kernel32" (ByVal hPass As Long)
Private Const TH32CS_SNAPPROCESS As Long = &H2
Private Type PROCESSENTRY32
  dwSize As Long
  cntUsage As Long
  th32ProcessID As Long
  th32DefaultHeapID As Long
  th32ModuleID As Long
  cntThreads As Long
  th32ParentProcessID As Long
  pcPriClassBase As Long
  dwFlags As Long
  szExeFile As String * 260
End Type
Private Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Const PROCESS_TERMINATE As Long = (&H1)
Private Declare Function TerminateProcess Lib "kernel32.dll" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
'Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
Public ErrSkip As Boolean

Public Function GetPidByImage(ByVal image As String) As Long
  On Local Error GoTo ErrOut:
  Dim hSnapShot As Long
  Dim uProcess As PROCESSENTRY32
  Dim r As Long, l As Long
  
  hSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0&)
  If hSnapShot = 0 Then Exit Function
  uProcess.dwSize = Len(uProcess)
  r = Process32First(hSnapShot, uProcess)
  l = Len(image)
  If l = 0 Then Exit Function
  Do While r
    If LCase(Left(uProcess.szExeFile, l)) = LCase(image) Then
      GetPidByImage = uProcess.th32ProcessID
      Exit Do
    End If
    r = Process32Next(hSnapShot, uProcess)
  Loop
  Call CloseHandle(hSnapShot)
ErrOut:
End Function

Public Function KillPID(ByVal PID As Long) As Boolean
On Local Error Resume Next
Dim h As Long

If PID = 0 Then 'pid를 찾을 수 없었다!
    KillPID = False '찾을 수 없음
    Exit Function '종료
End If

h = OpenProcess(PROCESS_TERMINATE, False, PID)
TerminateProcess h, 0
CloseHandle h
KillPID = True '찾았다
Sleep 10
ErrOut:
End Function

Public Sub KillProcessByName(ByVal ProcessName As String)
Dim a As Boolean
Do
a = KillPID(GetPidByImage(ProcessName))
Loop While a '해당 이름으로 된 프로세스를 모두 강제 종료한다
End Sub

Public Sub KillProcessByPID(ByVal PID As Long)
KillPID PID
End Sub

