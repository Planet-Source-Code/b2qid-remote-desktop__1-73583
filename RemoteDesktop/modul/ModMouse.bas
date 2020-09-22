Attribute VB_Name = "ModMouse"
'<CSCC>
'--------------------------------------------------------------------------------
'    Component  : ModMouse
'    Project    : NetRemote
'    Author     : B2qid www.labsoft.web.id
'    Description: {ParamList}
'
'    Modified   : 11/12/2010 2:52:18 PM
'--------------------------------------------------------------------------------
'</CSCC>
Private Type RECT
    Left   As Long
    Top    As Long
    Right  As Long
    Bottom As Long
End Type
Private Type POINTAPI
    x As Long
    y As Long
End Type
Public Const MOUSEEVENTF_MOVE = &H1
Public Const MOUSEEVENTF_LEFTDOWN = &H2
Public Const MOUSEEVENTF_LEFTUP = &H4
Public Const MOUSEEVENTF_RIGHTDOWN = &H8
Public Const MOUSEEVENTF_RIGHTUP = &H10
Public Const MOUSEEVENTF_MIDDLEDOWN = &H20
Public Const MOUSEEVENTF_MIDDLEUP = &H40
Public Const MOUSEEVENTF_ABSOLUTE = &H8000
Const KEYEVENTF_EXTENDEDKEY = &H1
Const KEYEVENTF_KEYUP = &H2


'Declare the API-Functions
Private Declare Function GetCursorPos Lib "user32" (lppoint As POINTAPI) As Long
Private Declare Function SetCursorPos _
                Lib "user32" (ByVal x As Long, _
                              ByVal y As Long) As Long
Private Declare Sub mouse_event _
                Lib "user32" (ByVal dwFlags As Long, _
                              ByVal dX As Long, _
                              ByVal dy As Long, _
                              ByVal cButtons As Long, _
                              ByVal dwExtraInfo As Long)
Private Declare Sub keybd_event _
                Lib "user32" (ByVal bVk As Byte, _
                              ByVal bScan As Byte, _
                              ByVal dwFlags As Long, _
                              ByVal dwExtraInfo As Long)
Public Type PCURSORINFO
    cbSize As Long
    Flags As Long
    hCursor As Long
    ptScreenPos As POINTAPI
End Type
Public Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal hIcon As Long) As Long
Public Const CursorIconSize As Integer = 9
Public Declare Function GetCursorInfo Lib "user32.dll" (ByRef pci As PCURSORINFO) As Long



Public Sub GetMousePos(x As Long, y As Long)
    Dim lppoint As POINTAPI
    GetCursorPos lppoint
    x = lppoint.x
    y = lppoint.y
End Sub

Public Sub SetMousePos(x As Long, y As Long)
    SetCursorPos x, y
End Sub

Public Sub MouseDown(x As Long, y As Long, Button As Long)
    Dim dwFlags As Long
    dwFlags = 0
    ' dwFlags = MOUSEEVENTF_ABSOLUTE
    ' + MOUSEEVENTF_MOVE
    SetMousePos x, y

    If Button = 1 Then dwFlags = dwFlags Or MOUSEEVENTF_LEFTDOWN
    If Button = 2 Then dwFlags = dwFlags Or MOUSEEVENTF_RIGHTDOWN
    If Button = 3 Then dwFlags = dwFlags Or MOUSEEVENTF_MIDDLEDOWN
    mouse_event dwFlags, x, y, 0&, 0&
    Debug.Print "MouseDown@" & Button & ";" & dwFlags
End Sub

Public Sub MouseUp(x As Long, y As Long, Button As Long)
    Dim dwFlags As Long
    dwFlags = 0
    SetMousePos x, y

    ' dwFlags = MOUSEEVENTF_ABSOLUTE
    ' + MOUSEEVENTF_MOVE
    If Button = 1 Then dwFlags = dwFlags Or MOUSEEVENTF_LEFTUP
    If Button = 2 Then dwFlags = dwFlags Or MOUSEEVENTF_RIGHTUP
    If Button = 3 Then dwFlags = dwFlags Or MOUSEEVENTF_MIDDLEUP
    mouse_event dwFlags, x, y, 0&, 0&
    Debug.Print "MouseUp@" & Button & ";" & dwFlags
End Sub

Public Sub SetKeyDown(KeyAscii As Long, Shift As Integer)
    Dim byt As Byte
    byt = Asc(Chr$(KeyAscii))

    If Shift Then
        keybd_event byt, CByte(0), KEYEVENTF_EXTENDEDKEY, 0
    Else
        keybd_event byt, 0, 0, 0
    End If

End Sub

Public Sub SetKeyUp(KeyAscii As Long, Shift As Integer)
    Dim byt As Byte
    byt = Asc(Chr$(KeyAscii))

    If Shift Then
        keybd_event byt, CByte(0), KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0
    Else
        keybd_event byt, CByte(0), KEYEVENTF_KEYUP, 0
    End If

End Sub

Public Sub Setkeypress(KeyAscii As Byte, Shift As Integer)
    keybd_event CByte(KeyAscii), CByte(0), CLng(0), CLng(0)
    keybd_event CByte(KeyAscii), CByte(0), KEYEVENTF_KEYUP, CLng(0)
End Sub

Public Sub MouseClick(x As Long, y As Long, Button As Long)

    Select Case Button

        Case 1
            ' Removed MOUSEEVENTF_MOVE + MOUSEEVENTF_ABSOLUTE +
            mouse_event MOUSEEVENTF_ABSOLUTE + MOUSEEVENTF_LEFTDOWN + MOUSEEVENTF_LEFTUP, x, y, 0, 0

        Case 2
            mouse_event MOUSEEVENTF_ABSOLUTE + MOUSEEVENTF_RIGHTDOWN + MOUSEEVENTF_RIGHTUP, x, y, 0, 0

        Case 3
            mouse_event MOUSEEVENTF_ABSOLUTE + MOUSEEVENTF_MIDDLEUP, x, y, 0, 0
    End Select

End Sub


