VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOptions 
   Caption         =   "VirtualDesktop Settings"
   ClientHeight    =   4425
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7560
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4425
   ScaleWidth      =   7560
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fmeInfo 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1335
      Index           =   0
      Left            =   360
      TabIndex        =   6
      Top             =   2280
      Width           =   6855
      Begin VB.Frame fmeHotkey 
         Caption         =   " Hotkey Settings "
         Enabled         =   0   'False
         Height          =   855
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   6615
         Begin VB.ComboBox cboKey2 
            Height          =   315
            Left            =   4080
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   360
            Width           =   2295
         End
         Begin VB.ComboBox cboKey1 
            Height          =   315
            ItemData        =   "frmOptions.frx":058A
            Left            =   840
            List            =   "frmOptions.frx":059A
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   360
            Width           =   2055
         End
         Begin VB.Label lblKey2 
            Caption         =   "Key 2:"
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
            Left            =   3480
            TabIndex        =   13
            Top             =   405
            Width           =   735
         End
         Begin VB.Label lblKeyPlus 
            Caption         =   "+"
            Height          =   255
            Left            =   3120
            TabIndex        =   11
            Top             =   390
            Width           =   735
         End
         Begin VB.Label lblKey1 
            Caption         =   "Key 1:"
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
            Left            =   240
            TabIndex        =   10
            Top             =   405
            Width           =   735
         End
      End
      Begin VB.CheckBox chkHotkey 
         Caption         =   "This Desktop Has A Hotkey"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   0
         Width           =   5775
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   6000
      TabIndex        =   5
      Top             =   3840
      Width           =   1335
   End
   Begin MSComctlLib.TabStrip tspSettings 
      Height          =   1935
      Left            =   240
      TabIndex        =   4
      Top             =   1800
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   3413
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   " Desktop and Hotkey Settings "
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox cboDesktops 
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1320
      Width           =   4455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Desktop:"
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
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Desktop and Wallpaper Information"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   840
      TabIndex        =   1
      Top             =   445
      Width           =   4335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Amoeba VirtualDesktop Settings"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   145
      Width           =   4815
   End
   Begin VB.Image Image2 
      Height          =   585
      Left            =   120
      Picture         =   "frmOptions.frx":05B5
      Stretch         =   -1  'True
      Top             =   120
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   990
      Left            =   0
      Picture         =   "frmOptions.frx":0B3F
      Stretch         =   -1  'True
      Top             =   -120
      Width           =   7620
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboDesktops_Click()
    On Error Resume Next

    If getFromIni(App.Path + "\settings.ini", "hasHotkey", CStr(cboDesktops.ListIndex)) = "1" Then
        chkHotkey.Value = 1
        fmeHotkey.Enabled = True
        cboKey1.Enabled = True
        cboKey2.Enabled = True
        lblKey1.Enabled = True
        lblKey2.Enabled = True
        lblKeyPlus.Enabled = True
    Else
        chkHotkey.Value = 0
        fmeHotkey.Enabled = False
        cboKey1.Enabled = False
        cboKey2.Enabled = False
        lblKey1.Enabled = False
        lblKey2.Enabled = False
        lblKeyPlus.Enabled = False
    End If
    
    If InStr(1, getFromIni(App.Path + "\settings.ini", "Hotkey", CStr(cboDesktops.ListIndex)), ",", vbBinaryCompare) <> 0 Then
        cboKey1.Text = cboKey1.List(CInt(Split(getFromIni(App.Path + "\settings.ini", "Hotkey", CStr(cboDesktops.ListIndex)), ",", , vbBinaryCompare)(0)))
        cboKey2.Text = cboKey2.List(CInt(Split(getFromIni(App.Path + "\settings.ini", "Hotkey", CStr(cboDesktops.ListIndex)), ",", , vbBinaryCompare)(1)))
    End If
End Sub

Private Sub cboKey1_Click()
    Dim Counter As Integer
    Counter = 0
    If cboKey1.ListIndex = -1 Then Exit Sub
    If cboKey2.ListIndex = -1 Then Exit Sub
    While Counter < desktopCount
        If CStr(cboKey1.ListIndex) + "," + CStr(cboKey2.ListIndex) = getFromIni(App.Path + "\settings.ini", "Hotkey", CStr(Counter)) Then Exit Sub
        Counter = Counter + 1
    Wend
    WritePrivateProfileString "Hotkey", CStr(cboDesktops.ListIndex), CStr(cboKey1.ListIndex) + "," + CStr(cboKey2.ListIndex), App.Path + "\settings.ini"
    Call frmSettings.createDesktopMenus
End Sub

Private Sub cboKey2_Click()
    Dim Counter As Integer
    Counter = 0
    If cboKey1.ListIndex = -1 Then Exit Sub
    If cboKey2.ListIndex = -1 Then Exit Sub
    While Counter < desktopCount
        If CStr(cboKey1.ListIndex) + "," + CStr(cboKey2.ListIndex) = getFromIni(App.Path + "\settings.ini", "Hotkey", CStr(Counter)) Then Exit Sub
        Counter = Counter + 1
    Wend
    WritePrivateProfileString "Hotkey", CStr(cboDesktops.ListIndex), CStr(cboKey1.ListIndex) + "," + CStr(cboKey2.ListIndex), App.Path + "\settings.ini"
    Call frmSettings.createDesktopMenus
End Sub

Private Sub chkHotkey_Click()

    If chkHotkey.Value = 1 Then
        If cboDesktops.Text <> "" Then
            WritePrivateProfileString "hasHotkey", CStr(cboDesktops.ListIndex), "1", App.Path + "\settings.ini"
            Call frmSettings.createDesktopMenus
        End If
        fmeHotkey.Enabled = True
        cboKey1.Enabled = True
        cboKey2.Enabled = True
        lblKey1.Enabled = True
        lblKey2.Enabled = True
        lblKeyPlus.Enabled = True
    ElseIf chkHotkey.Value = 0 Then
        If cboDesktops.Text <> "" Then
            WritePrivateProfileString "hasHotkey", CStr(cboDesktops.ListIndex), "0", App.Path + "\settings.ini"
            Call frmSettings.createDesktopMenus
        End If
        fmeHotkey.Enabled = False
        cboKey1.Enabled = False
        cboKey2.Enabled = False
        lblKey1.Enabled = False
        lblKey2.Enabled = False
        lblKeyPlus.Enabled = False
    End If
    
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
        
    '// Prepare all the info for the HotKey keycode information
    '// and place the data into the comboboxes
    prepareHotKeyInfo
    Dim Counter As Integer
    Counter = 0
    While Counter < desktopCount
        cboDesktops.AddItem "Desktop " + CStr(Counter + 1)
        Counter = Counter + 1
    Wend
    Counter = 0
    While Counter < 51
        cboKey2.AddItem HotkeyNames(Counter)
        Counter = Counter + 1
    Wend
    
    '// SNEAK PEEK:
    '// For future use: Icons for each individual desktop
    '// copyFolder GetSpecialfolder(CSIDL_DESKTOP), "C:\desktop"
    

End Sub



Private Sub tspSettings_Click()
    
    fmeInfo(0).Visible = False
    fmeInfo(tspSettings.SelectedItem.Index - 1).Visible = True

End Sub



Sub hotkeyVisible(Optional isVisible As Boolean = False)

End Sub
