VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'<CSCC>
'--------------------------------------------------------------------------------
'    Component  : clsUser
'    Project    : NetRemote
'    Author     : B2qid www.labsoft.web.id
'    Description: {ParamList}
'
'    Modified   : 11/12/2010 2:52:18 PM
'--------------------------------------------------------------------------------
'</CSCC>
Private m_StrClientIP As String
Private m_VarClientPort As Variant
Private m_BytClientsock As Byte
Private m_LonClientID As Long
Private m_Struser As String
Private m_bAllowControl As Boolean

Public Property Get AllowControl() As Boolean
    AllowControl = m_bAllowControl
End Property

Public Property Let AllowControl(ByVal bValue As Boolean)
    m_bAllowControl = bValue
End Property

Public Property Get User() As String
    User = m_Struser
End Property

Public Property Let User(ByVal StrValue As String)
    m_Struser = StrValue
End Property

Public Property Get ClientID() As Long
    ClientID = m_LonClientID
End Property

Public Property Let ClientID(ByVal LonValue As Long)
    m_LonClientID = LonValue
End Property

Public Property Get Clientsock() As Byte
    Clientsock = m_BytClientsock
End Property

Public Property Let Clientsock(ByVal BytValue As Byte)
    m_BytClientsock = BytValue
End Property

Public Property Get ClientPort() As Variant
    ClientPort = m_VarClientPort
End Property

Public Property Let ClientPort(ByVal VarValue As Variant)
    m_VarClientPort = VarValue
End Property

Public Property Get ClientIP() As String
    ClientIP = m_StrClientIP
End Property

Public Property Let ClientIP(ByVal StrValue As String)
    m_StrClientIP = StrValue
End Property
