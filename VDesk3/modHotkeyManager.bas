Attribute VB_Name = "modHotkeyManager"
Public MsgAtom(0 To 1023) As Long

Public ModNames(0 To 50) As String
Public ModCodes(0 To 50) As Long

Public HotkeyNames(0 To 50) As String
Public HotkeyCodes(0 To 50) As Long

Public Function prepareHotKeyInfo()
    Dim Counter As Integer
    Counter = 0
    While Counter < 51
        If Counter = 0 Then
            HotkeyNames(Counter) = "TAB"
            HotkeyCodes(Counter) = 9
        ElseIf Counter = 1 Then
            HotkeyNames(Counter) = "RETURN"
            HotkeyCodes(Counter) = 13
        ElseIf Counter = 2 Then
            HotkeyNames(Counter) = "PAGE UP"
            HotkeyCodes(Counter) = 33
        ElseIf Counter = 3 Then
            HotkeyNames(Counter) = "PAGE DOWN"
            HotkeyCodes(Counter) = 34
        ElseIf Counter = 4 Then
            HotkeyNames(Counter) = "INSERT"
            HotkeyCodes(Counter) = 45
        ElseIf Counter = 5 Then
            HotkeyNames(Counter) = "DELETE"
            HotkeyCodes(Counter) = 46
        ElseIf Counter > 5 And Counter < 16 Then
            HotkeyNames(Counter) = Chr(Counter + 42)
            HotkeyCodes(Counter) = Counter + 42
        ElseIf Counter > 15 And Counter < 39 Then
            HotkeyNames(Counter) = Chr(Counter + 49)
            HotkeyCodes(Counter) = Counter + 49
        ElseIf Counter > 38 And Counter < 48 Then
            HotkeyNames(Counter) = "F" & Chr(Counter + 10)
            HotkeyCodes(Counter) = Counter + 73
        ElseIf Counter = 48 Then
            HotkeyNames(Counter) = "F10"
            HotkeyCodes(Counter) = Counter + 73
        ElseIf Counter = 49 Then
            HotkeyNames(Counter) = "F11"
            HotkeyCodes(Counter) = Counter + 73
        ElseIf Counter = 50 Then
            HotkeyNames(Counter) = "F12"
            HotkeyCodes(Counter) = Counter + 73
        End If
        Counter = Counter + 1
    Wend
    Counter = 0
    While Counter < 4
        If Counter = 0 Then
            ModNames(Counter) = "ALT"
            ModCodes(Counter) = &H1
        ElseIf Counter = 1 Then
            ModNames(Counter) = "CTRL"
            ModCodes(Counter) = &H2
        ElseIf Counter = 2 Then
            ModNames(Counter) = "SHIFT"
            ModCodes(Counter) = &H4
        ElseIf Counter = 3 Then
            ModNames(Counter) = "WIN"
            ModCodes(Counter) = &H8
        End If
        Counter = Counter + 1
    Wend
        
End Function
