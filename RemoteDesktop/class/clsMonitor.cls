VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'<CSCC>
'--------------------------------------------------------------------------------
'    Component  : clsMonitor
'    Project    : NetRemote
'    Author     : B2qid www.labsoft.web.id
'    Description: {ParamList}
'
'    Modified   : 11/12/2010 2:52:18 PM
'--------------------------------------------------------------------------------
'</CSCC>
Option Explicit

' --------------------------------------------------------------------------
'               Copyright (C) 1998 Microsoft Corporation                   '
' --------------------------------------------------------------------------
' You have a royalty-free right to use, modify, reproduce and distribute   '
' the Sample Application Files (and/or any modified version) in any way    '
' you find useful, provided that you agree that Microsoft has no warranty, '
' obligations or liability for any Sample Application Files.               '
' --------------------------------------------------------------------------
' Written by Mike Dixon (mikedix@microsoft.com)                            '
' --------------------------------------------------------------------------


'===================================================
'Monitor Class, Contains information about a monitor
'All values should be stored as pixels
'===================================================

Private m_Handle        As Long
Private m_Left          As Long
Private m_Right         As Long
Private m_Top           As Long
Private m_Bottom        As Long
Private m_WorkLeft      As Long
Private m_Workright     As Long
Private m_Worktop       As Long
Private m_Workbottom    As Long
Private m_Width         As Long
Private m_Height        As Long

Public Property Get Handle() As Long
    Handle = m_Handle
End Property

Public Property Let Handle(lHandle As Long)
    m_Handle = lHandle
End Property

Public Property Get Height() As Long
    Height = m_Height
End Property

Public Property Let Height(h As Long)
    m_Height = h
End Property

Public Property Get Width() As Long
    Width = m_Width
End Property

Public Property Let Width(w As Long)
    m_Width = w
End Property

Public Property Get Left() As Long
    Left = m_Left
End Property

Public Property Let Left(l As Long)
    m_Left = l
End Property

Public Property Get Right() As Long
    Right = m_Right
End Property

Public Property Let Right(r As Long)
    m_Right = r
End Property

Public Property Get Top() As Long
    Top = m_Top
End Property

Public Property Let Top(t As Long)
    m_Top = t
End Property

Public Property Get Bottom() As Long
    Bottom = m_Bottom
End Property

Public Property Let Bottom(b As Long)
    m_Bottom = b
End Property

Public Property Get WorkLeft() As Long
    WorkLeft = m_WorkLeft
End Property

Public Property Let WorkLeft(l As Long)
    m_WorkLeft = l
End Property

Public Property Get WorkRight() As Long
    WorkRight = m_Workright
End Property

Public Property Let WorkRight(r As Long)
    m_Workright = r
End Property

Public Property Get WorkTop() As Long
    WorkTop = m_Worktop
End Property

Public Property Let WorkTop(t As Long)
    m_Worktop = t
End Property

Public Property Get WorkBottom() As Long
    WorkBottom = m_Workbottom
End Property

Public Property Let WorkBottom(b As Long)
    m_Workbottom = b
End Property


