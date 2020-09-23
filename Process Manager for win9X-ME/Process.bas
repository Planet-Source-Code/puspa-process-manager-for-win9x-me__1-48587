Attribute VB_Name = "Control_Funct"
Public Declare Function CreateToolhelp32Snapshot Lib "kernel32.dll" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Public Declare Function Process32First Lib "kernel32.dll" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Public Declare Function Process32Next Lib "kernel32.dll" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Public Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)
Public Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function GetForegroundWindow Lib "user32" () As Long

Public Const TH32CS_SNAPPROCESS = &H2 'Initalize Api's and types and constants here
Public Const TH32CS_SNAPALL = &HF
Public Const HWND_TOPMOST = -1
Public Const SWP_SHOWWINDOW = &H40
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOSIZE = &H1
Public Const WM_CLOSE = &H10

Public Type PROCESSENTRY32
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

Public items1 As Long
Public items2 As Long
'Public xcodm As Long
'Public xcodr As Long
'Public ycodm As Long
'Public ycodr As Long

Public Function Kill_Process(process_id As Long, Optional uExitCode As Long = 0) As Boolean 'close the passed process
       Dim phandle As Long
       Const PROCESS_ALL_ACCESS = 0
       phandle = OpenProcess(PROCESS_ALL_ACCESS, False, process_id) 'first get complete handle of the process
       Kill_Process = CBool(TerminateProcess(phandle, uExitCode)) 'then with this close the process
       Call CloseHandle(phandle) 'always close the opened handle
End Function

Public Function GetFileName(FullPath As String) As String 'can use getfilename from filesystem but I code one mine own.
       On Error Resume Next
       Dim dta As String
       Dim ch As String
       Dim plen As Long
       Dim cnt As Integer
       plen = Len(FullPath)
       cnt = 0
       ch = Mid$(FullPath, plen, 1)
       While ch <> "\" And cnt < plen 'start from last and search for "\" character
            dta = ch & dta
            cnt = cnt + 1
            ch = Mid$(FullPath, plen - cnt, 1)
       Wend 'when found "\" loop exits
       GetFileName = dta 'Thus dta contains only Filename
End Function
