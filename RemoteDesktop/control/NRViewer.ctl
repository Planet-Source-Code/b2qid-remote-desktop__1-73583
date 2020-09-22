VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.ocx"
Begin VB.UserControl NRViewer 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "NRViewer.ctx":0000
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "NRViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'<CSCC>
'--------------------------------------------------------------------------------
'    Component  : NRViewer
'    Project    : NetRemote
'    Author     : B2qid www.labsoft.web.id
'    Description: {ParamList}
'
'    Modified   : 11/12/2010 2:52:18 PM
'--------------------------------------------------------------------------------
'</CSCC>
Option Explicit
Private MousePressing As Integer
Private InControl As Boolean
Private mousex As Long, mousey As Long
Private mousesx As Long, mousesy As Long
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long
Private OutBuffer
'Private RemoteHeight As Long
'Private RemoteWidth As Long
Dim Buffer As String
Dim CurrentBufferSize As Long
Private m_bReadOnly As Boolean

Public Event OnConnect()
Public Event OnDisconnect()
Public Event OnClick()
Public Event OnReject(Description As String)
Public Event OnReceiveData(sData As String)

Private m_bScretch As Boolean
Dim mbw, mbh As Long
Private m_StrIP As String
Private IsSaving As Boolean
Private sFileName As String
Private m_OLEBackColor As OLE_COLOR

Private b_Connected As Boolean
Private m_LonRemoteWidth As Long
Private m_LonRemoteHeight As Long
Dim CanConnect          As Boolean
Private m_bShowCursor As Boolean
Private m_StrUserName As String
Private m_StrPassword As String
Private Zlib As New clsZLib

Public Property Get Password() As String
    Password = m_StrPassword
End Property

Public Property Let Password(ByVal StrValue As String)
    m_StrPassword = StrValue
    PropertyChanged "Password"
End Property

Public Property Get UserName() As String
    UserName = m_StrUserName
End Property

Public Property Let UserName(ByVal StrValue As String)
    m_StrUserName = StrValue
    PropertyChanged "UserName"
End Property


Public Property Get ShowCursor() As Boolean
    ShowCursor = m_bShowCursor
End Property

Public Property Let ShowCursor(ByVal bValue As Boolean)
    m_bShowCursor = bValue
    PropertyChanged "ShowCursor"
End Property

Public Property Get About()
    'frmAbout.Show vbModal
End Property

Public Property Get RemoteHeight() As Long
    RemoteHeight = m_LonRemoteHeight
End Property

Public Property Get RemoteWidth() As Long
    RemoteWidth = m_LonRemoteWidth
End Property


Public Property Get BackColor() As OLE_COLOR
    BackColor = m_OLEBackColor
End Property

Public Property Get Connected() As Boolean
    Connected = b_Connected
End Property

Public Property Let BackColor(ByVal OLEValue As OLE_COLOR)
    m_OLEBackColor = OLEValue
    UserControl.BackColor = m_OLEBackColor
    PropertyChanged "BackColor"
End Property

Public Property Get IP() As String
    IP = Winsock1.LocalIP
End Property

Public Property Get IsAktiv() As Boolean
    IsAktiv = b_Connected
End Property

Public Property Get Stretch() As Boolean
    Stretch = m_bScretch
End Property

Public Property Let Stretch(ByVal bValue As Boolean)
    m_bScretch = bValue
    PropertyChanged "Stretch"
End Property

Public Property Get ReadOnly() As Boolean
    ReadOnly = m_bReadOnly
End Property

Public Property Let ReadOnly(ByVal bValue As Boolean)
    m_bReadOnly = bValue
    InControl = Not bValue
    PropertyChanged "ReadOnly"
End Property

Private Sub UserControl_Click()
    RaiseEvent OnClick
    'SendData "MCL=" & mousex & "," & mousey & vbTab
End Sub

Private Sub UserControl_DblClick()
    On Error Resume Next
    SendData "MDBL=" & mousex & "," & mousey & vbTab
End Sub

Private Sub UserControl_Initialize()
    WriteJPGLib
    m_bScretch = True
    m_bReadOnly = True
    m_LonRemoteWidth = 0
    m_LonRemoteHeight = 0
    UserControl.BackColor = vbBlack
End Sub

Private Sub WriteJPGLib()
    On Error Resume Next
    Dim b
    Dim TargetDLL

    If modJPG.InstalledOK = False Then
        MsgBox "IJL15.DLL isn't installed.", vbExclamation, "IJL15.DLL Required"
        CanConnect = False
    End If

End Sub

Private Sub UserControl_InitProperties()
    m_bReadOnly = True
    InControl = Not m_bReadOnly
    m_bScretch = True
    m_OLEBackColor = vbBlack
    m_bShowCursor = False
    m_StrUserName = "User"
    m_StrPassword = "Password"
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    SendData "KDN=" & KeyCode & "," & Shift & vbTab
    'Debug.Print "Key down, keycode= " & KeyCode & "," & Shift

    KeyCode = 0: Shift = 0
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    SendData "KPR=" & KeyAscii & vbTab
    'Debug.Print "Key press, ascii=" & KeyAscii

    KeyAscii = 0
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    SendData "KUP=" & KeyCode & "," & Shift & vbTab
    'Debug.Print "Key up, keycode = " & KeyCode

    KeyCode = 0: Shift = 0 ' ditto
End Sub

Private Sub UserControl_MouseDown(Button As Integer, _
                                  Shift As Integer, _
                                  x As Single, _
                                  y As Single)
                                      
    On Error Resume Next
    mousex = UserControl.ScaleX(x, UserControl.ScaleMode, vbPixels) & ",": mousey = UserControl.ScaleY(y, UserControl.ScaleMode, vbPixels)
    If m_bScretch Then
        mousex = (UserControl.ScaleX(mbw * x, UserControl.ScaleMode, vbPixels) / UserControl.Width)
        mousey = (UserControl.ScaleY(mbh * y, UserControl.ScaleMode, vbPixels) / UserControl.Height)
    Else
        mousex = UserControl.ScaleX(m_LonRemoteWidth, UserControl.ScaleMode, vbPixels)
        mousey = UserControl.ScaleX(m_LonRemoteHeight, UserControl.ScaleMode, vbPixels)
    End If
    SendData "MDN=" & Button & "," & Shift & "," & mousex & "," & mousey & vbTab
    'Debug.Print "Mouse down on " & mousex & "," & mousey
    MousePressing = Button
End Sub

Private Sub UserControl_MouseMove(Button As Integer, _
                                  Shift As Integer, _
                                  x As Single, _
                                  y As Single)
    On Error Resume Next
    mousex = UserControl.ScaleX(mbw * x, UserControl.ScaleMode, vbPixels) & ",": mousey = UserControl.ScaleY(mbh * y, UserControl.ScaleMode, vbPixels)
    If m_bScretch Then
        mousex = (UserControl.ScaleX(mbw * x, UserControl.ScaleMode, vbPixels) / UserControl.Width)
        mousey = (UserControl.ScaleY(mbh * y, UserControl.ScaleMode, vbPixels) / UserControl.Height)
    Else
        mousex = UserControl.ScaleX(m_LonRemoteWidth, UserControl.ScaleMode, vbPixels)
        mousey = UserControl.ScaleX(m_LonRemoteHeight, UserControl.ScaleMode, vbPixels)
    End If
    SendData "MMV=" & Button & "," & Shift & "," & mousex & "," & mousey & vbTab
    'Debug.Print "Mouse@" & mousex & "," & mousey & " (" & Button & ")"
    
End Sub

Private Sub UserControl_MouseUp(Button As Integer, _
                                Shift As Integer, _
                                x As Single, _
                                y As Single)
    On Error Resume Next
    mousex = UserControl.ScaleX(x, UserControl.ScaleMode, vbPixels) & ",": mousey = UserControl.ScaleY(y, UserControl.ScaleMode, vbPixels)
    If m_bScretch Then
        mousex = (UserControl.ScaleX(mbw * x, UserControl.ScaleMode, vbPixels) / UserControl.Width)
        mousey = (UserControl.ScaleY(mbh * y, UserControl.ScaleMode, vbPixels) / UserControl.Height)
    Else
        mousex = UserControl.ScaleX(m_LonRemoteWidth, UserControl.ScaleMode, vbPixels)
        mousey = UserControl.ScaleX(m_LonRemoteHeight, UserControl.ScaleMode, vbPixels)
    End If
    SendData "MUP=" & Button & "," & Shift & "," & mousex & "," & mousey & vbTab
    'Debug.Print "Mouse up on " & mousex & "," & mousey & " (" & Button & ")"
End Sub

Private Sub SendData(Data As String)
    If Not m_bReadOnly Then
        If Winsock1.State = 7 Then
            Winsock1.SendData Data
        End If
    End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_bReadOnly = PropBag.ReadProperty("ReadOnly", True)
    m_bScretch = PropBag.ReadProperty("Stretch", True)
    m_OLEBackColor = PropBag.ReadProperty("BackColor", vbBlack)
    UserControl.BackColor = m_OLEBackColor
    m_bShowCursor = PropBag.ReadProperty("ShowCursor", False)
    m_StrUserName = PropBag.ReadProperty("UserName", "User")
    m_StrPassword = PropBag.ReadProperty("Password", "Password")
End Sub

Private Sub UserControl_Terminate()
    Me.DisConect
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("ReadOnly", m_bReadOnly, True)
    Call PropBag.WriteProperty("Stretch", m_bScretch, True)
    Call PropBag.WriteProperty("BackColor", m_OLEBackColor, vbBlack)
    Call PropBag.WriteProperty("ShowCursor", m_bShowCursor, False)
    Call PropBag.WriteProperty("UserName", m_StrUserName, "User")
    Call PropBag.WriteProperty("Password", m_StrPassword, "Password")
End Sub

Function Connect(RemoteHost, RemotePort)
        '<EhHeader>
        On Error GoTo Connect_Err
        '</EhHeader>
        If modJPG.InstalledOK = False Then
            MsgBox "IJL15.DLL isn't installed.", vbExclamation, "IJL15.DLL Required"
            Exit Function
        End If
        
        If Winsock1.State = 7 Then
            Exit Function
        End If
        b_Connected = True
100     Winsock1.Close
102     Winsock1.Connect RemoteHost, RemotePort
        '<EhFooter>
        Exit Function

Connect_Err:
        Err.Raise vbObjectError + 100, _
                  "RemoteDesktop.RDClient.Connect", _
                  "RDClient component failure"
        '</EhFooter>
End Function

Function DisConect()
    If Winsock1.State = 7 Then
        Winsock1.SendData "Disconected"
    End If
    b_Connected = False
    RaiseEvent OnDisconnect
    Winsock1.Close
    UserControl.Cls
    b_Connected = False
End Function

Private Sub Winsock1_Close()
    b_Connected = False
    RaiseEvent OnDisconnect
    b_Connected = False
    UserControl.Cls
End Sub

Private Sub Winsock1_Connect()
    b_Connected = True
    If Winsock1.State = 7 Then
        Winsock1.SendData "Request||" & m_StrUserName & "||" & m_StrPassword & "||"
    End If
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
        '<EhHeader>
        On Error GoTo Winsock1_DataArrival_Err
        '</EhHeader>
        Dim a As String
        Dim t As Long
100     Winsock1.GetData a

        Dim Data() As String
        
        If Mid(a, 1, 5) = "Acces" Then
            Data = Split(a, "||")
            If Data(1) = "failed" Then
                b_Connected = False
                Winsock1.Close
                UserControl.Cls
                RaiseEvent OnReject(Data(2))
            Else
                b_Connected = True
                RaiseEvent OnConnect
                Winsock1.SendData "Connected"
            End If
            Exit Sub
        ElseIf a = "Disconnect" Then
            Me.DisConect
        ElseIf Mid(a, 1, 4) = "Data" Then
            Data = Split(a$, "||")
            RaiseEvent OnReceiveData(Data(1))
            Exit Sub
        End If

102     If CurrentBufferSize = 0 Then
104         t = InStr(a, JPGSeparator)

106         If t <= 0 Then
                ' What ? No Separator and no size ?... Ignore packet
                Exit Sub
            End If

108         GetPicInfo Left(a, t - 1)
110         Buffer = Mid(a, t + Len(JPGSeparator))
        Else
112         Buffer = Buffer & a

114         If Len(Buffer) >= CurrentBufferSize Then
116             HandleBuffer
            End If
        End If
        
        'If Winsock1.State = 7 Then
            'Winsock1.SendData "NEXT_SCREEN"
        'End If
        
        '<EhFooter>
        Exit Sub

Winsock1_DataArrival_Err:

End Sub

Private Function JPGSeparator()
    JPGSeparator = vbNullChar & "RDC_SEP" & vbNullChar
End Function

Private Sub HandleBuffer()
        '<EhHeader>
        On Error GoTo HandleBuffer_Err
        '</EhHeader>
    'On Error GoTo ErrMe
        Dim imgdata(0) As String
        Dim t          As Double
        'Label1.Caption = "Rendering desktop from " & Winsock1.RemoteHost
100     t = Timer

        Dim dib As New cdibSection
        Dim jpgbuffer() As Byte
104     jpgbuffer() = StrConv(Buffer, vbFromUnicode)
        'Zlib.DecompressByte jpgbuffer
106     If LoadJPGFromPtr(dib, VarPtr(jpgbuffer(0)), CurrentBufferSize) Then
            
108         If m_bScretch Then
110             dib.PaintPicture UserControl.hdc, 0, 0, UserControl.ScaleWidth / Screen.TwipsPerPixelX, UserControl.ScaleHeight / Screen.TwipsPerPixelY
            Else
112             dib.PaintPicture2 UserControl.hdc
            End If
            
            If IsSaving Then
                SaveJPG dib, sFileName
            End If
            
            
            IsSaving = False
            sFileName = ""
            
114         dib.ClearUp

            If Winsock1.State = 7 Then
                Winsock1.SendData "ready||" & m_StrUserName & "||"
            End If
            
        End If

116     ReDim jpgbuffer(0)
118     Erase jpgbuffer()


120     CurrentBufferSize = 0
        Dim tt
122     Buffer = Mid(Buffer, CurrentBufferSize + 1)
124     tt = InStr(Buffer, JPGSeparator)

126     If tt <= 0 Then
            ' What ? No Separator and no size information ?... Clear the buffer... this shouldn't happen
128         Buffer = ""
            Exit Sub
        Else
130         GetPicInfo Left(Buffer, tt - 1)
132         Buffer = Mid(Buffer, tt + Len(JPGSeparator))
        End If
        'Winsock1.SendData "NEXT_SCREEN"
errme:
        '<EhFooter>
        Exit Sub

HandleBuffer_Err:
        Resume Next
        '</EhFooter>
End Sub

Sub SaveFrame(Filename As String)
    IsSaving = True
    sFileName = Filename
End Sub

Private Sub GetPicInfo(header)
        ' CurrentBufferSize = Val(Left(buffer, t - 1))
        '<EhHeader>
        On Error GoTo GetPicInfo_Err
        '</EhHeader>
        Dim n() As String
    
100     n = Split(header, vbTab)
102     UserControl.KeyPreview = True
104     If UBound(n) = 5 Then
106         'RemoteWidth = Val(n(0))
108         'RemoteHeight = Val(n(1))
            m_LonRemoteHeight = Val(n(1))
            m_LonRemoteWidth = Val(n(0))
110         mbw = Val(n(0)) * 15 + (UserControl.Width - UserControl.ScaleWidth)
112         mbh = Val(n(1)) * 15 + (UserControl.Height - UserControl.ScaleHeight)

114         If UserControl.Width < mbw Or UserControl.Height < mbh Then
                On Error Resume Next ' If maximized, please don't crash :)
                'Me.Move Me.Left, Me.Top, mbw, mbh
            End If
            mousesx = Val(n(2))
            mousesy = Val(n(3))
            'PutMousePointerAt Val(n(2)), Val(n(3))
            'InControl = (n(4) = "1")
116         InControl = Not m_bReadOnly
            
            'mousecursor.Visible = m_bReadOnly
118         CurrentBufferSize = Val(n(5))
        End If

        '<EhFooter>
        Exit Sub

GetPicInfo_Err:
        '</EhFooter>
End Sub

