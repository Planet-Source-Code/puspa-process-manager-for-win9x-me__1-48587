VERSION 5.00
Begin VB.Form run 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Run Program"
   ClientHeight    =   1170
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4245
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "run"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1170
   ScaleWidth      =   4245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox procpth 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   0
      TabIndex        =   3
      Top             =   360
      Width           =   4245
   End
   Begin VB.CommandButton cmdcan 
      Caption         =   "Cancel"
      Height          =   330
      Left            =   2745
      TabIndex        =   1
      Top             =   765
      Width           =   1455
   End
   Begin VB.CommandButton cmdrun 
      Caption         =   "Run Process"
      Default         =   -1  'True
      Height          =   330
      Left            =   1260
      TabIndex        =   0
      Top             =   765
      Width           =   1500
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Executiable Program Path"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   0
      TabIndex        =   2
      Top             =   90
      Width           =   2715
   End
End
Attribute VB_Name = "run"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdcan_Click() 'hide itself
        'SetWindowPos Me.hwnd, HWND_NOTOPMOST, Me.Left, Me.Top, 1, 1, SWP_SHOWWINDOW Or SWP_NOSIZE 'unregister window from topmost position
        'SetWindowPos maingui.hwnd, HWND_TOPMOST, maingui.Left, maingui.Top, 1, 1, SWP_SHOWWINDOW & SWP_NOSIZE
        Me.Hide
        maingui.Show
End Sub

Private Sub cmdrun_Click() 'Run selected application
        If procpth.Text <> "" Then 'check if button is falsely hit
           Dim rethandle As Double
           On Error GoTo shellerr
           rethandle = Shell(procpth.Text, vbNormalFocus) 'run app You can use ShellExecute Api to run application or other type of file with lot customizable parameters
           'For example of ShellExexute Api search for Inter Process Communication of mine in planet-source-code.com
           'SetWindowPos Me.hwnd, HWND_NOTOPMOST, Me.Left, Me.Top, 1, 1, SWP_SHOWWINDOW Or SWP_NOSIZE 'unregister window from topmost position
           'SetWindowPos maingui.hwnd, HWND_TOPMOST, 100, 100, 1, 1, SWP_SHOWWINDOW & SWP_NOSIZE
           Me.Hide ' hide itself
           Call maingui.Show  'show main window
           maingui.SetFocus
        Else: MsgBox "Enter/Select Executiable Program Path", vbCritical, "Process Manager" 'Inform to type process path
        End If
        Exit Sub
shellerr: MsgBox Err.Description, vbCritical, "Process Manager" ' Infom cause of not running  the process
End Sub

Private Sub Form_Load() 'initialize the combo box
        procpth.AddItem "Explorer.exe", 0
        procpth.AddItem "Command.com", 1
        procpth.AddItem "Regedit.exe", 2
        procpth.AddItem "Msconfig.exe", 3
        procpth.AddItem "Notepad.exe", 4
End Sub
