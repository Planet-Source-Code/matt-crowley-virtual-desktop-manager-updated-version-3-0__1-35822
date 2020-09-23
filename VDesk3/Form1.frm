VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5325
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4680
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "I have saved ALL my work - Lets give it a go!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2280
      Width           =   4455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit VWK"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2520
      Width           =   4455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "I have saved ALL my work - Lets give it a go!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   4455
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   4560
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   $"Form1.frx":0000
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   4455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   $"Form1.frx":0096
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   4455
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   4560
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "WARNING!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

Dim hidden(0 To 1024) As Long
Dim hiddencount As Integer


Sub CloseAllOpenProgrammes()

IsTask = WS_VISIBLE Or WS_BORDER

Dim hwCurr As Long
Dim intLen As Long
Dim strTitle As String

hiddencount = 0
hwCurr = GetWindow(Me.hwnd, GW_HWNDFIRST)

Do While hwCurr
    If hwCurr <> Me.hwnd And TaskWindow(hwCurr) Then
        intLen = GetWindowTextLength(hwCurr) + 1
        strTitle = Space$(intLen)
        intLen = GetWindowText(hwCurr, strTitle, intLen)
        If intLen > 0 Then
            If hwCurr <> Me.hwnd Then
                RetVal = ShowWindow(hwCurr, SW_HIDE) 'PostMessage(hwCurr, &H10, 0&, 0&)
                hidden(hiddencount) = hwCurr
                hiddencount = hiddencount + 1
            End If
        End If
    End If
    hwCurr = GetWindow(hwCurr, GW_HWNDNEXT)
Loop
End Sub

Sub ShowAllOpenProgrammes()
Dim ct As Integer
ct = 0
While ct < hiddencount
RetVal = ShowWindow(hidden(ct), SW_SHOW) 'PostMessage(hwCurr, &H10, 0&, 0&)
ct = ct + 1
Wend
End Sub


Private Sub Command1_Click()
    CloseAllOpenProgrammes
End Sub
Private Sub Command2_Click()
    MsgBox "Thanks! I hope you can use this in your development! If you want to contact me, do so on ICQ# 78105132 or email me at cdrmalton@yahoo.com - This source is devoted to planet source code!", vbInformation, "Exit VWK"
    End
End Sub

Private Sub Command3_Click()
ShowAllOpenProgrammes
End Sub

