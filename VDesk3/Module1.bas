Attribute VB_Name = "modWindowFunctions"
'//
'// modWindowFunctions
'// -----------------------------
'// Controls both functions to 'switch' desktops and control
'// the app's presence within the system tray
'//
'// Please comment and vote on PSC
'//

'// API Calls
Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As String, ByVal fuWinIni As Long) As Long
Public Declare Function RegisterHotKey Lib "user32" (ByVal hWnd As Long, ByVal id As Long, ByVal fsModifiers As Long, ByVal vk As Long) As Long
Public Declare Function GlobalAddAtom Lib "kernel32" Alias "GlobalAddAtomA" (ByVal lpString As String) As Integer
Public Declare Function GlobalDeleteAtom Lib "kernel32" (ByVal nAtom As Integer) As Integer
Public Declare Function UnregisterHotKey Lib "user32" (ByVal hWnd As Long, ByVal id As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Public Declare Function GetWindowWord Lib "user32" (ByVal hWnd As Long, ByVal wIndx As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal wIndx As Long) As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpSting As String, ByVal nMaxCount As Long) As Long
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

'// Public Types
Public Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uId As Long
    uFlags As Long
    uCallBackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

'// Public Constants
Public Const SPI_SETDESKWALLPAPER = 20
Public Const SPIF_UPDATEINIFILE = &H1
Public Const SPIF_SENDWININICHANGE = &H2
Public Const WM_HOTKEY = &H312
Public Const MOD_ALT = &H1
Public Const MOD_CONTROL = &H2
Public Const MOD_SHIFT = &H4
Public Const MOD_WIN = &H8
Public Const SWP_NOMOVE = 2
Public Const SWP_NOSIZE = 1
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDNEXT = 2
Public Const GWL_STYLE = (-16)
Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
Public Const SW_HIDE = 0
Public Const SW_MAXIMIZE = 3
Public Const SW_SHOW = 5
Public Const SW_MINIMIZE = 6
Public Const WM_CLOSE = &H10
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WS_VISIBLE = &H10000000
Public Const WS_BORDER = &H800000

'// Desktop Information Arrays (Kilodesktop, lol)
'// These hold all the hWnd information for up to 1024 desktops, 1024 windows each
'// (0 to 9) and the bax number of windows per desktop (1024)
Public openWindows(0 To 1023, 0 To 1023) As Long
Public openWindowsCount(0 To 1023) As Long
Public desktopCount As Long
Public currentDesktop As Integer
Public pastDesktop As Integer

'// Feature added to get any changed background info
Public openWindowsBackground(0 To 1023) As Long

'// Set var to type
Public NotifyIcon As NOTIFYICONDATA

'// Var to notifiy if it is a task
Public IsTask As Long

'// Switch Desktop
Public Function switchDesktop(fromDesktop As Integer, gotoDesktop As Integer)

    Dim hwndCurrentWindow As Long
    Dim intLen As Long
    Dim strWindowTitle As String
    Dim windowCounter As Integer
    
    '// Go through every window, check if it is a task, check if it is
    '// itself, then, if not, hide it
    IsTask = WS_VISIBLE Or WS_BORDER
    windowCounter = 0
    hwndCurrentWindow = GetWindow(frmSettings.hWnd, GW_HWNDFIRST)
    Do While hwndCurrentWindow
        If hwndCurrentWindow <> frmSettings.hWnd And TaskWindow(hwndCurrentWindow) Then
            intLen = GetWindowTextLength(hwndCurrentWindow) + 1
            strWindowTitle = Space$(intLen)
            intLen = GetWindowText(hwndCurrentWindow, strWindowTitle, intLen)
            If intLen > 0 Then
                If hwndCurrentWindow <> frmSettings.hWnd Then
                    RetVal = ShowWindow(hwndCurrentWindow, SW_HIDE)
                    openWindows(fromDesktop, windowCounter) = hwndCurrentWindow
                    windowCounter = windowCounter + 1
                End If
            End If
        End If
        hwndCurrentWindow = GetWindow(hwndCurrentWindow, GW_HWNDNEXT)
    Loop
    openWindowsCount(fromDesktop) = windowCounter
    
    '// write the background info to settings.ini
    WritePrivateProfileString "wallpaper", CStr(fromDesktop), getBackground, App.Path + "\settings.ini"

    '// Now, unhide the desktop that we want to have on top.  Go through
    '// the array information we collected from the last opening of this
    '// desktop.  By default, the array is blank, meaning no window will
    '// be opened if it the first time opening this desktop
    windowCounter = 0
    While windowCounter < openWindowsCount(gotoDesktop)
        RetVal = ShowWindow(openWindows(gotoDesktop, windowCounter), SW_SHOW)
        windowCounter = windowCounter + 1
    Wend
    
    '// Get the background settings for new desktop
    If Not getFromIni(App.Path + "\settings.ini", "wallpaper", CStr(gotoDesktop)) = "0" Then
        setBackground gotoDesktop, getFromIni(App.Path + "\settings.ini", "wallpaper", CStr(gotoDesktop))
    End If
    
    '// Move the current to past, the new desktop as current
    pastDesktop = fromDesktop
    currentDesktop = gotoDesktop
    
End Function

Function TaskWindow(hwCurr As Long) As Long
    
    '// Determine if this is a task window
    Dim lngStyle As Long
    lngStyle = GetWindowLong(hwCurr, GWL_STYLE)
    If (lngStyle And IsTask) = IsTask Then TaskWindow = True

End Function

'// The proceeding functions are related to getting and setting the
'// Windows background per desktop

Function setBackground(desktopNumber As Integer, fileName As String)

    SystemParametersInfo SPI_SETDESKWALLPAPER, 0&, fileName, SPIF_UPDATEINIFILE Or SPIF_SENDWININICHANGE

End Function

Function getBackground() As String 'desktopNumber As Integer)

    getBackground = getFromIni(GetWinDir + "\win.ini", "Desktop", "wallpaper")
    
End Function

Function getFromIni(fileName As String, Section As String, Keyword As String)

    Dim Msg, Success, x As String
    Dim Result As String * 128
    Success = GetPrivateProfileString(Section, Keyword, "", Result, Len(Result), fileName)
    If Left$(Result, 1) <> Chr$(0) Then
        x = Left$(Result, InStr(Result, Chr$(0)) - 1)
    Else: x = ""
    End If
    getFromIni = x
    
End Function


Public Function GetWinDir()

    Dim SetString As String, LenSetString As Integer
    SetString = Space(255)
    LenSetString = GetWindowsDirectory(SetString, 255)
    GetWinDir = Left(SetString, LenSetString)
    
End Function

Public Function copyFolder(Source As String, Dest As String)
    
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fld = fso.createfolder(Dest)
    If fso.folderexists(Source) Then
        If Not fso.folderexists(Dest) Then
            Set fld = fso.createfolder(Dest)
        End If
        fso.copyFolder Source, Dest, True
    Else
        MsgBox "Error. Unable to copy desktop information.", vbCritical, "FSO Error"
    End If

End Function
