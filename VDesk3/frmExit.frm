VERSION 5.00
Begin VB.Form frmExit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exit VirtualDesktop"
   ClientHeight    =   2205
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5895
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmExit.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2205
   ScaleWidth      =   5895
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboOptions 
      Height          =   315
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1080
      Width           =   4695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Exit"
      Default         =   -1  'True
      Height          =   375
      Left            =   3360
      TabIndex        =   0
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmExit.frx":058A
      Height          =   735
      Left            =   960
      TabIndex        =   2
      Top             =   240
      Width           =   4695
   End
   Begin VB.Image Image1 
      Height          =   465
      Left            =   240
      Picture         =   "frmExit.frx":063C
      Top             =   240
      Width           =   465
   End
End
Attribute VB_Name = "frmExit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'//
'// frmExit
'// -----------------------------
'// The program's exit dialog
'//
'// Please comment and vote on PSC
'//

Private Sub Command1_Click()

    Dim desktopCounter As Integer
    Dim windowCounter As Integer

    Call frmSettings.publicCallDesktop
    DoEvents
    
    '// Choose the senario and act upon it, whether it be close windows on
    '// other desktops or move them to the main one
    If cboOptions.Text = "Exit and Move All Programs to the Main Desktop" Then
        desktopCounter = 0
        While desktopCounter < desktopCount
            windowCounter = 0
            While windowCounter < openWindowsCount(desktopCounter)
                RetVal = ShowWindow(openWindows(desktopCounter, windowCounter), SW_SHOW)
                windowCounter = windowCounter + 1
            Wend
            desktopCounter = desktopCounter + 1
        Wend
        Shell_NotifyIcon NIM_DELETE, NotifyIcon
        End
    ElseIf cboOptions.Text = "Exit and Close All Programs Not on the Main Desktop" Then
        desktopCounter = 1
        While desktopCounter < desktopCount
            windowCounter = 0
            While windowCounter < openWindowsCount(desktopCounter)
                RetVal = SendMessage(openWindows(desktopCounter, windowCounter), WM_CLOSE, 0, 0)
                windowCounter = windowCounter + 1
            Wend
            desktopCounter = desktopCounter + 1
        Wend
        Shell_NotifyIcon NIM_DELETE, NotifyIcon
        End
    End If
        
End Sub

Private Sub Command2_Click()

    '// Unload it if it is canceled
    Unload Me

End Sub

Private Sub Form_Load()

    '// When the form loads, fill the combo box and keep on top
    SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    cboOptions.AddItem "Exit and Move All Programs to the Main Desktop"
    cboOptions.AddItem "Exit and Close All Programs Not on the Main Desktop"
End Sub
