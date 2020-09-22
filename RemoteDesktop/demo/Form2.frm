VERSION 5.00
Object = "{3CAD2590-BEE3-4F92-AF6A-446BAEE77F91}#2.0#0"; "NetRemote.ocx"
Begin VB.Form Form2 
   Caption         =   "Client Demo"
   ClientHeight    =   2535
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2115
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   2535
   ScaleWidth      =   2115
   StartUpPosition =   3  'Windows Default
   Begin NetRemote.NRViewer NRViewer1 
      Height          =   4455
      Left            =   120
      TabIndex        =   11
      Top             =   1200
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   7858
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   3600
      TabIndex        =   10
      Top             =   600
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1080
      TabIndex        =   9
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Capture"
      Height          =   375
      Left            =   7320
      TabIndex        =   6
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox txtPort 
      Height          =   285
      Left            =   4320
      TabIndex        =   5
      Text            =   "27977"
      Top             =   120
      Width           =   735
   End
   Begin VB.CheckBox chkScretch 
      Caption         =   "Scretch"
      Height          =   255
      Left            =   1320
      TabIndex        =   4
      Top             =   120
      Width           =   975
   End
   Begin VB.CheckBox chkControl 
      Caption         =   "control"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox txt 
      Height          =   285
      Left            =   2520
      TabIndex        =   2
      Text            =   "192.168.1.2"
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton cmdConect 
      Caption         =   "Conect"
      Height          =   375
      Left            =   5160
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdDisconect 
      Caption         =   "Disconect"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6000
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Password"
      Height          =   255
      Left            =   2520
      TabIndex        =   8
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "User"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   600
      Width           =   735
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkControl_Click()
    NRViewer1.ReadOnly = chkControl.Value = 0
End Sub

Private Sub chkScretch_Click()
    NRViewer1.Stretch = chkScretch.Value = 1
End Sub

Private Sub cmdConect_Click()
    NRViewer1.UserName = Text1.Text
    NRViewer1.Password = Text2.Text
    'NRViewer1.LocalPort = Text3
    NRViewer1.Connect txt.Text, txtPort.Text
    
    NRViewer1.ReadOnly = chkControl.Value = 0
    NRViewer1.Stretch = chkScretch.Value = 1
    NRViewer1.ShowCursor = True 'Still Progress
    Button NRViewer1.IsAktiv
End Sub

Sub Connect()
    NRViewer1.UserName = Text1.Text
    NRViewer1.Password = Text2.Text
    'NRViewer1.LocalPort = Text3
    NRViewer1.Connect txt.Text, txtPort.Text
    
    NRViewer1.ReadOnly = chkControl.Value = 0
    NRViewer1.Stretch = chkScretch.Value = 1
    NRViewer1.ShowCursor = True 'Still Progress
    Button NRViewer1.IsAktiv
End Sub

Private Sub cmdDisconect_Click()
    NRViewer1.DisConect
    Button NRViewer1.IsAktiv
End Sub

Private Sub Command1_Click()
On Error Resume Next
    Kill App.Path & "\Cap.jpg"
    NRViewer1.SaveFrame (App.Path & "\Cap.jpg")
End Sub

Private Sub Form_Load()
    txt = NRViewer1.IP
    Text1.Text = NRViewer1.UserName
    Text2.Text = NRViewer1.Password
    Button False
    chkScretch.Value = IIf(NRViewer1.Stretch, 1, 0)
    chkControl.Value = IIf(NRViewer1.ReadOnly, 1, 0)
End Sub

Sub Button(En As Boolean)
    cmdConect.Enabled = Not En
    cmdDisconect.Enabled = En
End Sub

Private Sub Form_Resize()
    NRViewer1.Move NRViewer1.Left, NRViewer1.Top, Me.Width - NRViewer1.Left - 300, Me.Height - NRViewer1.Top - 600
End Sub

Private Sub Form_Unload(Cancel As Integer)
    NRViewer1.DisConect
End Sub

Private Sub NRViewer1_OnConnect()
    Button NRViewer1.IsAktiv
End Sub

Private Sub NRViewer1_OnDisconnect()
    Button NRViewer1.IsAktiv
End Sub

Private Sub NRViewer1_OnReject(Description As String)
    MsgBox "Server Has Reject Your Connection, Reason : " & Description
End Sub
