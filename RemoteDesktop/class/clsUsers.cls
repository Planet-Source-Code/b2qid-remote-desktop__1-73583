VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'<CSCC>
'--------------------------------------------------------------------------------
'    Component  : clsUsers
'    Project    : NetRemote
'    Author     : B2qid www.labsoft.web.id
'    Description: {ParamList}
'
'    Modified   : 11/12/2010 2:52:18 PM
'--------------------------------------------------------------------------------
'</CSCC>
' The private collection object
Private m_ColclsUsers As Collection


Private Sub Class_Initialize()

    Set m_ColclsUsers = New Collection

End Sub


' This sub adds a new clsUser item to the collection.
Sub Add(ClsItem As clsUser, Optional VarKey As Variant)

    'TODO: Initialize the new clsUser item's properties here

    m_ColclsUsers.Add ClsItem, VarKey

End Sub


' This sub remove an item from the collection.
Sub Remove(VarIndex As Variant)

    m_ColclsUsers.Remove VarIndex

End Sub


' This function returns a clsUser item from the collection. It's the default method.
Function Item(VarIndex As Variant) As clsUser
Attribute Item.VB_UserMemId = 0

    Set Item = m_ColclsUsers.Item(VarIndex)

End Function


' This property returns the number of items in the collection.
Property Get Count() As Long

    Count = m_ColclsUsers.Count

End Property


' This sub remove all items from the collection.
Sub Clear()

    Set m_ColclsUsers = New Collection

End Sub


' This function adds "For Each" enumeration support. Must have a -4 DispID.
Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4

    Set NewEnum = m_ColclsUsers.[_NewEnum]

End Function



