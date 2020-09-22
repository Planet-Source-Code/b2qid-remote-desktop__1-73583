VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.ocx"
Begin VB.UserControl NRServer 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   InvisibleAtRuntime=   -1  'True
   Picture         =   "RDServer.ctx":0000
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "RDServer.ctx":06A9
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   480
      Top             =   720
   End
   Begin MSWinsockLib.Winsock listener 
      Left            =   2640
      Top             =   2040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   27977
   End
   Begin MSWinsockLib.Winsock client 
      Index           =   0
      Left            =   2160
      Top             =   2040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "NRServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Const FieldSeparator = vbVerticalTab ' Separator for fields...
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Dim cMonitors As clsMonitors
'Monitor class, contains information about a monitor
Dim cMonitor As clsMonitor


'Rectangle structure, for determining
'monitors at a given position
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type


        

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Declare Function GetCursor Lib "user32" () As Long
Private Declare Function GetCursorPos Lib "user32" (lppoint As POINTAPI) As Long

'Constants for the return value when finding a monitor
Const MONITOR_DEFAULTTONULL = &H0       'If the monitor is not found, return 0
Const MONITOR_DEFAULTTOPRIMARY = &H1    'If the monitor is not found, return the primary monitor
Const MONITOR_DEFAULTTONEAREST = &H2

Private dib As New cdibSection
'Public AllowedIps As String
Private m_bAllowControl As Boolean
Private m_LonPort As Long
Private m_LonMonitorCount As Long
Private m_LonSelectedMonitor As Long
Private m_LonClientCount As Long
Private m_LonMaxConection As Long
Dim osock() As Boolean
Private m_StrIP As String
Private xClientCount As Integer
Private b_sending As Boolean

Public Type tClient
    ClientIP        As String
    ClientPort      As Long
    Clientsock      As Byte
    ClientID        As Long
    User            As String
    AllowControl    As Boolean
End Type

Dim xClient() As tClient

Public Event OnClientConect(oClient As tClient)
Public Event OnRequest(oClient As tClient, Accept As Boolean)
Public Event OnClientDisconect(oClient As tClient)
Public Event GotError(ErrNo As Long, ErrDesc As String)
Public Event OnReceiveData(sData As String)

'Public Event OnInfo(Info As String)

Private b_aktif As Boolean
Private m_LonQuality As Long
Private m_bUseCursor As Boolean
'Private m_StrUserName As String
Private m_StrPassword As String

Public Property Get Password() As String
    Password = m_StrPassword
End Property

Public Property Let Password(ByVal StrValue As String)
    m_StrPassword = StrValue
    PropertyChanged "Password"
End Property

'Public Property Get UserName() As String
'    UserName = m_StrUserName
'End Property

'Public Property Let UserName(ByVal StrValue As String)
'    m_StrUserName = StrValue
'    PropertyChanged "UserName"
'End Property

Public Property Get ShowCursor() As Boolean
    ShowCursor = m_bUseCursor
End Property

Public Property Let ShowCursor(ByVal bValue As Boolean)
    m_bUseCursor = bValue
    PropertyChanged "ShowCursor"
End Property

Public Property Get Quality() As Long
    Quality = m_LonQuality
End Property

Public Property Let Quality(ByVal LonValue As Long)
    If LonValue > 100 Then LonValue = 100
    m_LonQuality = LonValue
    PropertyChanged "Quality"
End Property

Public Property Get IsActive() As Boolean
    IsActive = b_aktif
End Property

Public Property Get IsSending() As Boolean
    IsSending = Timer2.Enabled
End Property

Public Property Get IP() As String
    IP = listener.LocalIP
End Property

Public Property Get MaxConection() As Long
    MaxConection = m_LonMaxConection
End Property

Public Property Let MaxConection(ByVal LonValue As Long)
    m_LonMaxConection = LonValue
    PropertyChanged "MaxConection"
End Property

Public Property Get ClientCount() As Long
    ClientCount = client.Count - 1
End Property

Public Function GetClient(ClientIndex) As tClient
On Error Resume Next
    GetClient = xClient(ClientIndex)
End Function

Public Function SetAllowControl(ClientIndex As Long, Allow As Boolean)
    xClient(ClientIndex).AllowControl = Allow
End Function

Public Property Let SelectedMonitor(ByVal LonValue As Long)
    Dim Counter As Long
    Counter = cMonitors.Monitors.Count
    If LonValue > Counter Then LonValue = Counter
    m_LonSelectedMonitor = LonValue
End Property

Public Property Get MonitorCount() As Long
    cMonitors.Refresh
    m_LonMonitorCount = cMonitors.Monitors.Count
    MonitorCount = m_LonMonitorCount
End Property

Public Property Get Port() As Long
    Port = m_LonPort
End Property

Public Property Let Port(ByVal LonValue As Long)
    m_LonPort = LonValue
    PropertyChanged "Port"
End Property

Public Property Get AllowControl() As Boolean
    AllowControl = m_bAllowControl
End Property

Public Property Let AllowControl(ByVal bValue As Boolean)
    m_bAllowControl = bValue
    PropertyChanged "AllowControl"
End Property

Private Sub RedrawDesktop(Index As Integer)
        On Error GoTo RedrawDesktop_Err
        '</EhHeader>
100     b_sending = False
        'On Error GoTo errme
102     If client.Count < 2 Then
            Timer2.Enabled = False
            Exit Sub
        End If
    
        Dim t, t2, t3, TimeTook
        Dim DeskDC&
        Dim CurL, CurT, CurW, CurH As Long
    
104     With cMonitors.Monitors(m_LonSelectedMonitor)
106         CurL = .Left
108         CurT = .Top
110         CurW = .Width
112         CurH = .Height
        End With
        
    
        Dim newinterval
114     t = Timer
        Dim xwidth  As Long
        Dim xheight As Long

116     DeskDC& = GetDC(GetDesktopWindow()) ' Get's the DC of the desktop

118     xwidth = CurW 'Screen.Width / Screen.TwipsPerPixelX ' Gets the width and height of the screen (in pixels) //Old
120     xheight = CurH 'Screen.Height / Screen.TwipsPerPixelY //Old
    
122     t2 = Timer
124     dib.Create xwidth, xheight
        BitBlt dib.hdc, 0&, 0&, xwidth, xheight, DeskDC&, CurL, CurT, vbSrcCopy
        Dim Point As POINTAPI
        Dim pcin As PCURSORINFO
        Dim Ret
    
126     GetCursorPos Point
        If m_bUseCursor Then
128         pcin.hCursor = GetCursor
130         pcin.cbSize = Len(pcin)
132         'Ret = GetCursorInfo(pcin)
134         b_sending = True
136
138         DrawIcon dib.hdc, Point.x - CursorIconSize - CurL, Point.y - CursorIconSize - CurT, pcin.hCursor
            End If
        Dim bufsize As Long
140     bufsize = 512000
142     ReDim buffer(bufsize) As Byte ' Reserve 512k RAM
144     SaveJPGToPtr dib, VarPtr(buffer(0)), bufsize, m_LonQuality

146     t3 = Timer
        '
        Dim mousex As Long
        Dim mousey As Long
148     'GetMousePos mousex, mousey ' Gets the mouse position
        
        ' Sends the JPG data to the clients
150     SendDatatoClients xwidth & vbTab & xheight & vbTab & Point.x & vbTab & Point.y, StrConv(buffer(), vbUnicode), bufsize
        
152     dib.ClearUp ' Clears the DIB... Recover resources
154     ReDim buffer(0)
156     Erase buffer()

158     TimeTook = Timer - t
160     newinterval = 2 * (TimeTook) * 1000
        
        ' Adjust the timer to 2 times the rendering time of the desktop picture...
162     If newinterval > 30000 Then newinterval = 30000
        ' Cant take more than 30 seconds
        Debug.Print "Interval Took " & newinterval
164     Timer2.Interval = newinterval

        'Check is there is another client connected
        ' if no one connected then stop create dib
        If client.Count < 2 Then Timer2.Enabled = False

        Exit Sub
errme:
    
        'Timer1.Enabled = True
        '<EhFooter>
        Exit Sub

RedrawDesktop_Err:
       RaiseEvent GotError(Err.Number, Err.Description & "@ line" & Erl)
       Resume Next
End Sub

Private Sub client_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    'Timer1.Enabled = False
    Dim a$
    Dim Data() As String 'command||Username||Password||
    
    client(Index).GetData a$

    'If InStr(AllowedIps, client(Index).RemoteHostIP & vbTab) Then
    
    If Mid(a$, 1, 7) = "Request" Then
        Data = Split(a$, "||")
        xClient(Index).User = Data(1)
        
        'If m_StrUserName <> Data(1) Then 'Wrong User Disconect it
        '    client(Index).SendData "Acces||failed||Wrong User Name||"
        '    Exit Sub
        'End If
        
        If m_StrPassword <> Data(2) Then 'Wrong Password Disconect it
            client(Index).SendData "Acces||failed||Wrong Password||"
            Exit Sub
        End If
        
        Dim Accept As Boolean
        Accept = True
        RaiseEvent OnRequest(xClient(Index), Accept)
        If Not Accept Then
            client(Index).SendData "Acces||failed||Server Rejected Your request||"
            Exit Sub
        Else
            Timer2.Enabled = True
            client(Index).SendData "Acces||Permited||Server Accept Your request||"
        End If
        Exit Sub
    ElseIf a$ = "Connected" Then
        RaiseEvent OnClientConect(xClient(Index))
        Exit Sub
    ElseIf a$ = "Disconected" Then
        DisconectClient Index
        RaiseEvent OnClientDisconect(xClient(Index))
        Exit Sub
    ElseIf Mid(a$, 1, 4) = "Data" Then
        Data = Split(a$, "||")
        RaiseEvent OnReceiveData(Data(1))
        Exit Sub
    End If
    
    If m_bAllowControl Then
        If xClient(Index).AllowControl Then
            HandleInput a$, Index
        End If
    End If
    'Else
        ' Ignore packet
    'End If

    'Timer1.Enabled = True
End Sub

Private Sub DisconectClient(Index As Integer)
    On Error Resume Next
    client(Index).Close
    Unload client(Index)
End Sub

Sub Disconnect(ClientIdex As Long)
    On Error Resume Next
    client_Close (ClientIdex)
End Sub

Sub SendData(Data As String, Optional Clientsock As Long = 0)
    If Clientsock <> 0 Then
        If osock(Clientsock) Then
            If client(Clientsock).State = 7 Then ' Connected
                client(Clientsock).SendData "Data||" & Data
            End If
        End If
    Else
        Dim o As Long
        For o = 1 To m_LonMaxConection - 1
        If osock(o) Then
        If client(o).State = 7 Then ' Connected
            If client(o).Tag = "OK" Then ' It's free to send data
                'RaiseEvent OnInfo("Sending Screen to " & client(o).RemotePort)
                client(o).Tag = "SENDING" ' Marks as 'SENDING'. It will be OK once all data is transmited

                'If InStr(AllowedIps, client(o).RemoteHostIP & vbTab) Then
                    client(o).SendData "Data||" & Data
                'Else
                '    client(o).SendData header & vbTab & "0" & vbTab & datsize & JPGSeparator & Left(data, datsize)
                'End If

                ' Sends the datsize along with the data spaced by a null character
                ' The buffer can handle it :)
            End If
        Else
            Disconnect o
        End If
        End If
    Next

    End If
End Sub

Private Sub TerminateAll()
    Dim I As Long
    On Error GoTo errme
    For I = 1 To m_LonMaxConection - 1
        If osock(I) = True Then
            client(I).Close
            osock(I) = False
        End If
    Next
errme:
    
End Sub

Private Sub client_Close(Index As Integer)
    ' Client has disconnected... Yay... I don't have to send bytes there anymore! :)
    On Error Resume Next
    RaiseEvent OnClientDisconect(xClient(Index))
    osock(Index) = False
    xClient(Index).AllowControl = False
    xClient(Index).ClientID = 0
    xClient(Index).ClientIP = ""
    xClient(Index).ClientPort = 0
    xClient(Index).Clientsock = 0
    xClient(Index).User = ""
    DisconectClient Index
    If client.Count <= 1 Then
        Timer2.Enabled = False
    End If
End Sub

Private Sub client_SendComplete(Index As Integer)
    On Error Resume Next
    client(Index).Tag = "OK" ' I'm ready!
End Sub

Private Sub listener_ConnectionRequest(ByVal requestID As Long)
        '<EhHeader>
        On Error GoTo listener_ConnectionRequest_Err
        '</EhHeader>
        Dim iOpenSocket As Byte
        
100     iOpenSocket = FindSocket
        
102     If iOpenSocket > m_LonMaxConection Then
            Exit Sub
        End If
104     If iOpenSocket = 0 Then
            'AddToLog "User connecting but slots full..."
        Else
106         osock(iOpenSocket) = True
108         Load client(iOpenSocket)
            client(iOpenSocket).Close
            'Load SockFile(iOpenSocket)
110         client(iOpenSocket).Accept requestID
112         client(iOpenSocket).Tag = "OK"
            xClient(iOpenSocket).ClientID = requestID
            xClient(iOpenSocket).Clientsock = iOpenSocket
            xClient(iOpenSocket).ClientIP = listener.RemoteHostIP
            xClient(iOpenSocket).ClientPort = listener.RemotePort
            'Timer2.Enabled = True '// theres a client connected make sure u draw the desktops
            xClientCount = xClientCount + 1
            'RaiseEvent OnClientConect(xClient)
        End If
    
    
        '<EhFooter>
        Exit Sub

listener_ConnectionRequest_Err:
        RaiseEvent GotError(Err.Number, Err.Description)
        Resume Next

        'Err.Raise vbObjectError + 100, _
                  "RemoteDesktop.RDServer.listener_ConnectionRequest", _
                  "RDServer component failure"
        '</EhFooter>
End Sub

Private Sub Timer2_Timer()
    RedrawDesktop 0
End Sub

Private Sub UserControl_Initialize()
    Set cMonitor = New clsMonitor
    Set cMonitors = New clsMonitors
    WriteJPGLib
    m_LonSelectedMonitor = 1
    cMonitors.Refresh
End Sub

Private Sub UserControl_InitProperties()
    m_bAllowControl = True
    m_LonPort = 27977
    m_LonMonitorCount = 1
    m_LonMaxConection = 10
    m_LonQuality = 90
    m_bUseCursor = False
    'm_StrUserName = "User"
    m_StrPassword = "Password"
End Sub

Private Function FindSocket(Optional currentSocket As Byte) As Byte
        '<EhHeader>
        On Error GoTo FindSocket_Err
        '</EhHeader>
    

        Dim iTemp As Byte
100     If currentSocket = 0 Then
102         iTemp = 1
        Else
104         iTemp = currentSocket
        End If
        
106     For iTemp = iTemp To m_LonMaxConection
108         If currentSocket = 0 Then
110             If osock(iTemp) = False Then
112                 FindSocket = iTemp
                    Exit For
                End If
            Else
114             If osock(iTemp) = True Then
116                 FindSocket = iTemp
                    Exit For
                End If
            End If
118     Next iTemp
        Exit Function
120     Resume Next
        '<EhFooter>
        Exit Function

FindSocket_Err:
       
        Resume Next
        '</EhFooter>
End Function

Private Sub SendDatatoClients(header As String, Data As String, datsize As Long)
    Dim o As Long
    For o = 1 To m_LonMaxConection - 1
        If osock(o) Then
        If client(o).State = 7 Then ' Connected
            If client(o).Tag = "OK" Then ' It's free to send data
                'RaiseEvent OnInfo("Sending Screen to " & client(o).RemotePort)
                client(o).Tag = "SENDING" ' Marks as 'SENDING'. It will be OK once all data is transmited

                'If InStr(AllowedIps, client(o).RemoteHostIP & vbTab) Then
                    client(o).SendData header & vbTab & "1" & vbTab & datsize & JPGSeparator & Left(Data, datsize)
                'Else
                '    client(o).SendData header & vbTab & "0" & vbTab & datsize & JPGSeparator & Left(data, datsize)
                'End If

                ' Sends the datsize along with the data spaced by a null character
                ' The buffer can handle it :)
            End If
        Else
            Disconnect o
        End If
        End If
    Next

End Sub

Private Sub WriteJPGLib()
    On Error Resume Next
    Dim b
    Dim TargetDLL

    If modJPG.InstalledOK = False Then
        
    End If

End Sub

Private Function JPGSeparator()
    JPGSeparator = vbNullChar & "RDC_SEP" & vbNullChar
End Function

Private Sub HandleInput(buf As String, Index As Integer)
    Dim c() As String, e() As String
    Dim kpr
    
    c = Split(buf, vbTab)
    Dim o, cmx, cmy
    Dim gotmouse As Boolean
    Dim d() As String
    kpr = ""
    
    For o = LBound(c) To UBound(c)
        d = Split(c(o), "=")

        If UBound(d) < 1 Then Exit For
        e = Split(d(1) & ",,,,,", ",")  ' Make sure it has many args so the following cmds don't crash in case of under-buffer

        Select Case d(0)

            Case "MCL"
                ' Mouse Click
                ModMouse.MouseClick Val(e(0)), Val(e(1)), 1

            Case "MDBL"
                ' Mouse double click
                ModMouse.MouseClick Val(e(0)), Val(e(1)), 1

                ' The click event is also raised in double-click... so the following line
                ' isn't needed
                '        ModMouse.MouseClick Val(e(0)), Val(e(1)), 1
            Case "MDN"

                ' Mouse button pressed
                 ModMouse.MouseDown Val(e(2)), Val(e(3)), Val(e(0))
                 gotmouse = False ' No need to move the mouse to the pos
            Case "MUP"

                ' Mouse button depressed
                 ModMouse.MouseUp Val(e(2)), Val(e(3)), Val(e(0))
                 gotmouse = False
            Case "MMV"
                ' Mouse moved
                cmx = Val(e(2))
                cmy = Val(e(3))
                gotmouse = True

            Case "KDN"

                ' Key down
                ' SetKeyDown Val(e(0)), Val(e(1))
            Case "KUP"

                ' SetKeyUp Val(e(0)), Val(e(1))
            Case "KPR"
                '        kpr = kpr & Chr$(Val(e(0)))
                ' SetKeyDown Val(e(0)), Val(e(1))
                SendKeys Chr$(e(0))
                ' SetKeyDown Val(e(0)), Val(e(1))
                ' SetKeyUp Val(e(0)), Val(e(1))
        End Select

    Next

    If gotmouse = True Then
        ModMouse.SetMousePos CLng(cmx), CLng(cmy)
    End If

End Sub

Public Property Let Start(bValue As Boolean)
        '<EhHeader>
        On Error GoTo Start_Err
        '</EhHeader>
        Dim o As Integer
        ReDim osock(m_LonMaxConection) As Boolean
        ReDim xClient(m_LonMaxConection) As tClient
        xClientCount = 0
        b_aktif = bValue
        'Timer2.Enabled = bValue
100     If bValue Then
            
            If modJPG.InstalledOK = False Then
                Timer2.Enabled = Not bValue
                b_aktif = Not bValue
                MsgBox "IJL15.DLL isn't installed.", vbExclamation, "IJL15.DLL Required"
                Exit Property
            End If
            
104         listener.Close
106         listener.LocalPort = m_LonPort
108         listener.Listen
        Else
            Timer2.Enabled = bValue
112         For o = 1 To m_LonMaxConection
114             DisconectClient o
            Next
122         listener.Close
        End If
        '<EhFooter>
        Exit Property

Start_Err:
        listener.Close
        b_aktif = False
        RaiseEvent GotError(Err.Number, Err.Description)
        Resume Next
        'Err.Raise vbObjectError + 100, _
                  "RemoteDesktop.RDServer.Start", _
                  "RDServer component failure"
        
        '</EhFooter>
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_bAllowControl = PropBag.ReadProperty("AllowControl", True)
    m_LonPort = PropBag.ReadProperty("Port", 27977)
    m_LonMaxConection = PropBag.ReadProperty("MaxConection", 10)
    m_LonQuality = PropBag.ReadProperty("Quality", 90)
    m_bUseCursor = PropBag.ReadProperty("ShowCursor", False)
    'm_StrUserName = PropBag.ReadProperty("UserName", "User")
    m_StrPassword = PropBag.ReadProperty("Password", "Password")
End Sub

Private Sub UserControl_Resize()
    UserControl.Height = 480
    UserControl.Width = 480
End Sub

Private Sub UserControl_Terminate()
    TerminateAll
    Timer2.Enabled = False
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("AllowControl", m_bAllowControl, True)
    Call PropBag.WriteProperty("Port", m_LonPort, 27977)
    Call PropBag.WriteProperty("MaxConection", m_LonMaxConection, 10)
    Call PropBag.WriteProperty("Quality", m_LonQuality, 90)
    Call PropBag.WriteProperty("ShowCursor", m_bUseCursor, False)
    'Call PropBag.WriteProperty("UserName", m_StrUserName, "User")
    Call PropBag.WriteProperty("Password", m_StrPassword, "Password")
End Sub

Public Sub About()
    frmAbout.Show vbModal
End Sub
