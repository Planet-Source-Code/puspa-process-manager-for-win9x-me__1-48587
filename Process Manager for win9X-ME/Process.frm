VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form maingui 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Process Manager"
   ClientHeight    =   3705
   ClientLeft      =   1725
   ClientTop       =   2760
   ClientWidth     =   7950
   Icon            =   "Process.frx":0000
   LinkTopic       =   "maingui"
   MaxButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   7950
   Begin VB.Timer killer 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   2520
      Top             =   3240
   End
   Begin VB.Timer update 
      Interval        =   1000
      Left            =   1080
      Top             =   3240
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Run New Process"
      Height          =   330
      Left            =   6390
      TabIndex        =   3
      Top             =   3330
      Width           =   1500
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Terminate"
      Height          =   330
      Left            =   4995
      TabIndex        =   2
      Top             =   3330
      Width           =   1410
   End
   Begin MSComctlLib.ListView info 
      Height          =   3255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   5741
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Process Name"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Process ID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Parent Process ID"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "No. Of Threads"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "No. of Usage"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "Module ID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "Default Heap ID"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   90
      TabIndex        =   0
      Top             =   3375
      Width           =   3015
   End
   Begin VB.Menu mnuaction 
      Caption         =   "&Action"
      Begin VB.Menu mnushow 
         Caption         =   "&Refresh"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnurun 
         Caption         =   "&Run"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuterminate 
         Caption         =   "&Terminate"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnukiller 
         Caption         =   "&Killer"
         Shortcut        =   ^K
      End
      Begin VB.Menu mnuexit 
         Caption         =   "&Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnusystem 
      Caption         =   "&System"
      Begin VB.Menu mnushutdown 
         Caption         =   "&Shut Down"
      End
      Begin VB.Menu mnurestart 
         Caption         =   "&Reboot"
      End
      Begin VB.Menu mnuforceclose 
         Caption         =   "&Force Shut Down"
      End
   End
   Begin VB.Menu mnuabout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "maingui"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' This is simple example of showing all running process in computer
' It can terminate running process and can open new process
' It is thus very useful in finding if any virus ,trojan wormbs are running in your computer
' If one is running it can give You full path of its location
' It also demonstrate use windows API and sendmessage api
'
' Feel FREE to use this code except for commercial purpose
' Just make sure You rate this coding and give me credit for this
'
' To use this coding correctly You should running Win9X or WinME
' It also runs in WinXP but can't close process , shutdown fetures disabled, don't give full path of selected process
'
' Its prime feature is to give XP like task manager for win 9X / ME
' You can use to set this application to top of all other application with setwindowpos api
'
' Coder: Puspa Raj Mahat
'
' "Suprise Everyone"
'

Public Sub display() 'looks for process currently running
    Dim hSnapshot As Long ' declare variables
    Dim processInfo As PROCESSENTRY32
    Dim success As Long
    Dim exeName As String
    Dim retval As Long
    Dim itm As ListItem
        
    hSnapshot = CreateToolhelp32Snapshot(TH32CS_SNAPALL, 0) 'get process snapshot handle
    processInfo.dwSize = Len(processInfo) 'set actual size for information
    success = Process32First(hSnapshot, processInfo) 'get first process info from snapshot handle
    If hSnapshot = -1 Then 'No process running
       Exit Sub
    End If
    items2 = 0 'It is used to refresh system only when new process is added or removed
    While success <> 0 'Count Total number of process currently running
          items2 = items2 + 1
          success = Process32Next(hSnapshot, processInfo)
    Wend
    retval = CloseHandle(hSnapshot) 'close handle
    If items1 <> items2 Then 'If new process present or removed display new information
       hSnapshot = CreateToolhelp32Snapshot(TH32CS_SNAPALL, 0) 'same old procedure to get latest process status
       processInfo.dwSize = Len(processInfo)
       success = Process32First(hSnapshot, processInfo)
       cnt = 1
       info.ListItems.Clear
       While success <> 0
             exeName = Left(processInfo.szExeFile, InStr(processInfo.szExeFile, vbNullChar) - 1)
             Set itm = info.ListItems.Add(cnt, , GetFileName(exeName)) 'Display result in listview
             itm.Tag = exeName
             itm.SubItems(1) = processInfo.th32ProcessID
             itm.SubItems(2) = processInfo.th32ParentProcessID
             itm.SubItems(3) = processInfo.cntThreads
             itm.SubItems(4) = processInfo.cntUsage
             itm.SubItems(5) = processInfo.th32ModuleID
             itm.SubItems(6) = processInfo.th32DefaultHeapID
             cnt = cnt + 1
             processInfo.dwSize = Len(processInfo)
             success = Process32Next(hSnapshot, processInfo)
       Wend
       retval = CloseHandle(hSnapshot)
       Label3.Caption = cnt - 1 & " Process Running" 'inform user with updated information
       items1 = items2 'assign new item count as old item count
     End If
End Sub

Private Sub Command2_Click() 'Terminates the selected process
       Dim ret As Boolean
       Dim retval As Long
       retval = MsgBox("Are You Sure To Terminate " & info.SelectedItem.Text, vbYesNo, "Process Terminator") 'confirm on termination
       If retval = vbYes Then 'if confirmed then
          ret = Kill_Process(CLng(info.SelectedItem.SubItems(1))) 'remove process
          If ret = True Then MsgBox info.SelectedItem.Text & " Terminated", vbInformation, "Process Manager" 'on success inform success
          If ret = False Then MsgBox "Process can't be Terminated", vbCritical, "Process Manager" 'on failure inform failure
          Call display 'Refresh display of listview with new process list
       End If
End Sub

Private Sub Command3_Click()
        Call mnurun_Click 'Run application wizard
End Sub

Private Sub Form_Load()
       ' You can register window to top most as task manager using setwindowpos api
       'ret = SetWindowPos(Me.hwnd, HWND_TOPMOST, 100, 100, 1, 1, SWP_SHOWWINDOW Or SWP_NOSIZE) ' set the window on top of all other windows
       items1 = 0 'set running process to zero
       run.Hide 'don't show run wizard
       Call display 'refresh the system
End Sub

Private Sub Form_Unload(Cancel As Integer)
        'SetWindowPos Me.hwnd, HWND_NOTOPMOST, Me.Left, Me.Top, 1, 1, SWP_SHOWWINDOW Or SWP_NOSIZE 'unregister window from topmost position
        ExitProcess (0)
End Sub


Private Sub info_DblClick()
            MsgBox info.SelectedItem.Tag, vbInformation, "Process Manager" ' Displays complete path of the running syste,
            ' only compatible in WIn9X and Win ME
End Sub

Private Sub killer_Timer() 'closes the foreground or current working window
       Dim winID As Long
       Dim ret As Long
       killer.Enabled = False
       winID = GetForegroundWindow() 'get foreground window handle
       ret = SendMessage(winID, WM_CLOSE, 0, 0) 'send message to foreground window to immediately close
End Sub

Private Sub mnuabout_Click() 'Just some Info
       MsgBox "[ Software ] : Process Manager" & vbCrLf & _
              "[ Version ] : 2.0" & vbCrLf & _
              "[ PlatForm ] : Windows 9 X / ME" & vbCrLf & vbCrLf & _
              "[ Coder ] : Puspa Raj Mahat" & vbCrLf & _
              "[ CopyRights ] : All Rights Reserved to Coder" & vbCrLf & vbCrLf & _
              "[ Description ] : Shows All HIDDEN Process Running" & vbCrLf & _
              "[ Note ] : For Full Path Double Click Process ", vbInformation, "Process Manager"
End Sub

Private Sub mnuexit_Click() 'exit Application
       ExitProcess (0)
End Sub

Private Sub mnuforceclose_Click() 'Shut down windows without waiting like restart
        ExitWindowsEx 4, 0
End Sub

Private Sub mnukiller_Click() 'enable killer timer for user to actvate window or application to be closed
        If killer.Enabled = False Then killer.Enabled = True
End Sub

Private Sub mnurestart_Click() 'Restart the Computer
        ExitWindowsEx 2, 0
End Sub

Private Sub mnurun_Click() 'Shows the New Application Run Wizard
        'SetWindowPos Me.hwnd, HWND_NOTOPMOST, Me.Left, Me.Top, 1, 1, SWP_SHOWWINDOW & SWP_NOSIZE 'unregister window from topmost position
        'SetWindowPos run.hwnd, HWND_TOPMOST, 100, 100, 1, 1, SWP_SHOWWINDOW & SWP_NOSIZE 'register Run wizard as top most window
        run.Show 'show the wizard
End Sub

Private Sub mnushow_Click() 'Refresh the current list manually
        Call display
End Sub

Private Sub mnushutdown_Click() 'Shut Down the Computer
            ExitWindowsEx 1, 0
End Sub

Private Sub mnuterminate_Click() 'Terminate the selected process in the listview
        Call Command2_Click
End Sub

Private Sub update_Timer() 'Continuously monitor the system running processes
        Update.Enabled = False
        Call display 'Checks for new process or old process removed
        Update.Enabled = True
End Sub
