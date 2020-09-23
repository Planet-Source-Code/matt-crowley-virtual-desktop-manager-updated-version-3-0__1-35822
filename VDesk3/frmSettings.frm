VERSION 5.00
Object = "{307C5043-76B3-11CE-BF00-0080AD0EF894}#1.0#0"; "MsgHoo32.OCX"
Begin VB.Form frmSettings 
   Caption         =   "Amoeba VirtualDesktop Settings"
   ClientHeight    =   3450
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   5970
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSettings.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   3450
   ScaleWidth      =   5970
   StartUpPosition =   2  'CenterScreen
   Begin MsghookLib.Msghook Msghook 
      Left            =   0
      Top             =   0
      _Version        =   65536
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
   End
   Begin VB.Menu mnu_1 
      Caption         =   "Menu"
      Begin VB.Menu mnuDesktop 
         Caption         =   "Goto Desktop 1"
         Index           =   0
      End
      Begin VB.Menu mnuSeperatorA 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAddNewDesktop 
         Caption         =   "&Add New Desktop"
      End
      Begin VB.Menu mnuRemoveOldDesktop 
         Caption         =   "&Remove Old Desktop"
      End
      Begin VB.Menu mnuSeperatorB 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVirtualDesktopSettings 
         Caption         =   "VirtualDesktop Settings"
      End
      Begin VB.Menu mnuSeperatorC 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'//
'// frmSettings
'// -----------------------------
'// Manages the tray icon and menu
'//
'// Please comment and vote on PSC
'//

'//
'// VERSION TWO
'// BUILD 2 - Updated Hotkeys to CTRL + Number
 
'// VERSION THREE - Major Update
'// BUILD 1 - Thanks to Bob Shull his menu array code, saved some
'//           spaced and allowed for me being able to add these
'//           new features:
'//              -Add and Delete Desktops
'//              -Added frmOptions
'//              -Editible Hotkey Info
'//              -Dynamic menu loading
'//              -Dynamic Atom and Hotkey Registration
'//           *Special thanks to:
'//              -Bob Shull: Suggestions and contributed menu array code
'//               and some of the new ideas
'//              -Clint LaFever: Suggestions and ideas (also told me about
'//               a shareware version that does virtually the same thing for
'//               30 bucks)
'//              -Everyone who commented + those who told me compatibility info.
'//
'// WANT MORE INFO OR WANT TO JOIN THE VIRTUAL DESKTOP PROJECT?
'// Email me at CodeMonkey04@cs.com to Join in and contribute your
'// part.
'//
'// LITTLE KNOWN PROJECT INFO:
'// -Project started on June 12th
'// -Uploaded on June 13th
'// -Version 2 Posted June 12th
'// -Gathered new suggestions/code and released June 15th
'//
'// Matt Crowley
'// CodeMonkey04@cs.com
'// http://www.greenwave.org
'//



Private Sub Form_Load()

    On Error Resume Next
    
    '// PREPARE THE ARRAYS!
    prepareHotKeyInfo
    DoEvents
    
    '// Get the current windows background
    getBackground
    
    '// Hide this form
    Me.Hide
    
    '// Get all the desktops, write the menu information
    desktopCount = CLng(getFromIni(App.Path + "\settings.ini", "Desktop", "count"))
    createDesktopMenus
    
    '// Initialize the current and past desktops
    currentDesktop = 0
    pastDesktop = 0

    '// Call up the first desktop
    Call mnuDesktop_Click(0)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

    '// Check if both isMinimized and Right Click on mouse
    Dim Result As Long
    Dim Message As Long
    If Me.ScaleMode = vbPixels Then
        Message = x
    Else
        Message = x / Screen.TwipsPerPixelX
    End If
    If Message = WM_RBUTTONUP Then
        Result = SetForegroundWindow(Me.hWnd)
        Me.PopupMenu Me.mnu_1
    End If

End Sub

Private Sub Form_Resize()

    '// Hide the form if it is minimized
    If frmSettings.WindowState = vbMinimized Then
        frmSettings.Hide
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    '// Begin UnRegister Atom & Hotkey Process
    createDesktopMenus True
    
    '// Kill the SysTray icon
    Shell_NotifyIcon NIM_DELETE, NotifyIcon
    
End Sub

'// Build 3 Addition
'// Add UBound Desktop
Private Sub mnuAddNewDesktop_Click()
    If desktopCount + 1 <= 1024 Then
        WritePrivateProfileString "Desktop", "count", CStr(desktopCount + 1), App.Path + "\settings.ini"
        DoEvents
        desktopCount = CLng(getFromIni(App.Path + "\settings.ini", "Desktop", "count"))
        DoEvents
        createDesktopMenus
        DoEvents
        MsgBox "Desktop " + CStr(desktopCount) + " has been created successfully.", vbInformation + vbOKOnly, "Information"
    Else
        MsgBox "You cannot have over 1024 desktops.", vbInformation + vbOKOnly, "Information"
    End If
End Sub

'// Major edit for Build 3: Now up to 1024 desktops
'// All these menus deal with access to the 10 desktops, Desktop 1 being
'// the original
Private Sub mnuDesktop_Click(Index As Integer)
    Shell_NotifyIcon NIM_DELETE, NotifyIcon
    DoEvents
    switchDesktop currentDesktop, Index
    DoEvents
    With NotifyIcon
        .cbSize = Len(NotifyIcon)
        .hWnd = Me.hWnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallBackMessage = WM_MOUSEMOVE
        .hIcon = Me.Icon
        .szTip = "VirtualDesktop: Viewing Desktop " & CStr(Index + 1) & vbNullChar
    End With
    checkDesktopMenu Index
    Shell_NotifyIcon NIM_ADD, NotifyIcon
    DoEvents
End Sub

Private Sub mnuExit_Click()
    Load frmExit
    frmExit.Show
End Sub

'// Major edit for Build 3
'// Controls the HOTKEYS
Private Sub Msghook_Message(ByVal Msg As Long, ByVal wp As Long, ByVal lp As Long, Result As Long)

    Dim Counter As Integer
    While Counter < desktopCount
        If wp = MsgAtom(Counter) Then
            Call mnuDesktop_Click(Counter)
        End If
        Counter = Counter + 1
    Wend
    Msghook.InvokeWindowProc Msg, wp, lp
    
End Sub

'// Build 3 Addition
'// This sub manages the menu by first unloading all members
'// then reloading all desktop-related members with the
'// correct amount of desktops
'// ADDED:
'// Manages HotKey lookup information also
Sub createDesktopMenus(Optional unRegHotKeys As Boolean = False)
    
    '// THIS IS NEEDED! THERE IS BOUND TO BE AT LEAST
    '// ONE ERROR BECAUSE OF MENU and HOTKEY ATOM SETTINGS
    On Error Resume Next
    
    '// OK now down to the code
    Dim Counter As Integer
    Counter = 0
    While Counter <= mnuDesktop.UBound
        Unload mnuDesktop.Item(Counter)
        UnregisterHotKey Me.hWnd, MsgAtom(Counter)
        GlobalDeleteAtom MsgAtom(Counter)
        Counter = Counter + 1
    Wend
    If unRegHotKeys = True Then Exit Sub
    Counter = 0
    While Counter < desktopCount
        Load mnuDesktop.Item(Counter)
        If getFromIni(App.Path + "\settings.ini", "hasHotkey", CStr(Counter)) = "1" Then
            If InStr(1, getFromIni(App.Path + "\settings.ini", "Hotkey", CStr(cboDesktops.ListIndex)), ",", vbBinaryCompare) <> 0 Then
                MsgAtom(Counter) = GlobalAddAtom("V" + CStr(Counter))
                RegisterHotKey frmSettings.hWnd, MsgAtom(Counter), ModCodes(CLng(Split(getFromIni(App.Path + "\settings.ini", "Hotkey", CStr(Counter)), ",", , vbBinaryCompare)(0))), HotkeyCodes(CLng(Split(getFromIni(App.Path + "\settings.ini", "Hotkey", CStr(Counter)), ",", , vbBinaryCompare)(1)))
            End If
        End If
        mnuDesktop.Item(Counter).Caption = "Goto Desktop " + CStr(Counter + 1)
        mnuDesktop.Item(Counter).Checked = False
        Counter = Counter + 1
    Wend
    Msghook.HwndHook = frmSettings.hWnd
    Msghook.Message(WM_HOTKEY) = True
    checkDesktopMenu currentDesktop
End Sub

'// Build 3 Addition
'// Thanks to Bob Shull
'// Easier way to manage Desktop Menu Checking
Sub checkDesktopMenu(Index As Integer)
    Dim Counter As Integer
    Counter = 0
    While Counter <= mnuDesktop.UBound
        mnuDesktop.Item(Counter).Checked = False
        Counter = Counter + 1
    Wend
    mnuDesktop.Item(Index).Checked = True
End Sub

'// Build 3 Addition
'// Removes UBound desktop
Private Sub mnuRemoveOldDesktop_Click()
    If desktopCount - 1 >= 1 Then
        WritePrivateProfileString "Desktop", "count", CStr(desktopCount - 1), App.Path + "\settings.ini"
        DoEvents
        desktopCount = CLng(getFromIni(App.Path + "\settings.ini", "Desktop", "count"))
        DoEvents
        createDesktopMenus
        DoEvents
        MsgBox "Desktop " + CStr(desktopCount + 1) + " has been deleted successfully.", vbInformation + vbOKOnly, "Information"
    Else
        MsgBox "You cannot have less than 1 desktop.", vbInformation + vbOKOnly, "Information"
    End If
End Sub

'// Build 3 Addition
'// Access VirtualDesktop Settings
Private Sub mnuVirtualDesktopSettings_Click()
    frmOptions.Show
End Sub

Sub publicCallDesktop()
    Call mnuDesktop_Click(0)
End Sub
