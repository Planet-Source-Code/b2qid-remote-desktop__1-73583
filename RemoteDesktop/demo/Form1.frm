VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{3CAD2590-BEE3-4F92-AF6A-446BAEE77F91}#3.1#0"; "NetRemote.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Server Demo"
   ClientHeight    =   4800
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   6150
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   6150
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1200
      TabIndex        =   11
      Text            =   "512000"
      Top             =   1320
      Width           =   1335
   End
   Begin NetRemote.NRServer NRServer1 
      Height          =   735
      Left            =   5280
      TabIndex        =   10
      Top             =   0
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1296
      ShowCursor      =   -1  'True
      BufferSize      =   128000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Client"
      Height          =   495
      Left            =   4800
      TabIndex        =   9
      Top             =   840
      Width           =   1215
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2895
      Left            =   120
      TabIndex        =   8
      Top             =   1680
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   5106
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "User"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "IP"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Port"
         Object.Width           =   1764
      EndProperty
   End
   Begin VB.TextBox txtMax 
      Height          =   285
      Left            =   1200
      TabIndex        =   4
      Text            =   "10"
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   360
      Left            =   2640
      TabIndex        =   3
      Top             =   720
      Width           =   990
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   360
      Left            =   2640
      TabIndex        =   2
      Top             =   240
      Width           =   990
   End
   Begin VB.TextBox txtPort 
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Text            =   "27977"
      Top             =   240
      Width           =   1335
   End
   Begin VB.TextBox txtQuality 
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      Text            =   "30"
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Buffer"
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label lblMaxClient 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Max Client"
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   600
      Width           =   735
   End
   Begin VB.Label lblPort 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Port"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   6
      Top             =   240
      Width           =   285
   End
   Begin VB.Label lblPort 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Quality"
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   960
      Width           =   480
   End
   Begin VB.Menu c_control 
      Caption         =   "Control"
      Visible         =   0   'False
      Begin VB.Menu control 
         Caption         =   "Allow Control"
         Index           =   0
      End
      Begin VB.Menu control 
         Caption         =   "Disconect"
         Index           =   1
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim UserIndex As Long
'Dim f() As Form2

Private Sub cmdStart_Click()
    NRServer1.Port = Val(txtPort.Text)
    NRServer1.MaxConection = Val(txtMax.Text)
    NRServer1.Quality = Val(txtQuality.Text)
    NRServer1.BufferSize = Val(Text1.Text)
    NRServer1.Start = True
    cmdStart.Enabled = Not NRServer1.IsActive
    cmdStop.Enabled = NRServer1.IsActive
    'If NRServer1.IsActive Then Form2.Show
End Sub

Private Sub cmdStop_Click()
    NRServer1.Start = False
    cmdStart.Enabled = Not NRServer1.IsActive
    cmdStop.Enabled = NRServer1.IsActive
    ListView1.ListItems.Clear
End Sub

Private Sub Command1_Click()
    Dim f As New Form2
    f.Connect
    f.Show
    'ReDim f(Val(txtMax)) As Form2
    'For i = 0 To Val(txtMax)
    '    Set f(i) = New Form2
    '    f(i).Connect
    '    f(i).Show
    'Next
End Sub

Private Sub control_Click(Index As Integer)
    Select Case Index
        Case 0
            control(Index).Checked = Not control(Index).Checked
            Call NRServer1.SetAllowControl(UserIndex, control(Index).Checked)
        Case 1
            NRServer1.Disconnect UserIndex
    End Select
End Sub

Private Sub Form_Load()
    NRServer1.About
End Sub

Private Sub Form_Unload(Cancel As Integer)
    NRServer1.Start = False
    Dim frm As Form
    For Each frm In Forms
        Unload frm
    Next
End Sub

Private Sub ListView1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If ListView1.ListItems.Count = 0 Then Exit Sub
    
    If Button = 1 Then Exit Sub
    
    UserIndex = Val(Mid(ListView1.SelectedItem.Key, 2))
    Dim Client As NetRemote.clsUser
    Set Client = NRServer1.GetClient(UserIndex)
    If Client Is Nothing Then Exit Sub
    
    If Client.AllowControl Then
        control(0).Checked = True
    Else
        control(0).Checked = False
    End If
    
    PopupMenu c_control
End Sub

Private Sub NRServer1_GotError(ErrNo As Long, ErrDesc As String)
    Debug.Print ErrDesc
End Sub


Private Sub NRServer1_OnClientConect(oClient As NetRemote.clsUser)
    Dim List As ListItem
    Set List = ListView1.ListItems.Add(, "C" & (oClient.Clientsock), oClient.Clientsock)
    List.SubItems(1) = oClient.User
    List.SubItems(2) = oClient.ClientIP
    List.SubItems(3) = oClient.ClientPort
End Sub

Private Sub NRServer1_OnClientDisconect(oClient As NetRemote.clsUser)
    Dim i As Integer
    For i = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(i).Key = "C" & (oClient.Clientsock) Then
            ListView1.ListItems.Remove (i)
            Exit Sub
        End If
    Next
End Sub

Private Sub NRServer1_OnRequest(oClient As NetRemote.clsUser, Accept As Boolean)
    Accept = True
    oClient.AllowControl = True
End Sub
