VERSION 5.00
Begin VB.Form PluginDemo 
   Caption         =   "Form1"
   ClientHeight    =   4005
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9120
   LinkTopic       =   "Form1"
   ScaleHeight     =   4005
   ScaleWidth      =   9120
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Run Plugin"
      Height          =   735
      Left            =   240
      TabIndex        =   8
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Caption         =   "Send Event"
      Height          =   1815
      Left            =   2520
      TabIndex        =   2
      Top             =   960
      Width           =   4695
      Begin VB.CommandButton Command2 
         Caption         =   "Send"
         Height          =   375
         Left            =   1200
         TabIndex        =   7
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox params 
         Height          =   285
         Left            =   2160
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   720
         Width           =   2175
      End
      Begin VB.TextBox EventName 
         Height          =   285
         Left            =   2160
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "Parameters"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Name:"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.ComboBox List 
      Height          =   315
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Config Plugin"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   1935
   End
End
Attribute VB_Name = "PluginDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim P As PluginCom

Private Sub Command1_Click()
  If List.ListIndex = -1 Then Exit Sub
  On Error Resume Next
  'Get the Form FrmConfig & show it
  Dim F As Form
  Set F = P.plug.ConfigForm(ID)
  F.Show
End Sub

Private Sub Command2_Click()
  If List.ListIndex = -1 Then Exit Sub
  P.plug.SendEvent ID, EventName.Text, params.Text
End Sub

Private Sub Command3_Click()
  If List.ListIndex = -1 Then Exit Sub
  P.plug.SendEvent ID, "StartFrm"
End Sub

Private Sub Form_Load()
  Set P = New PluginCom
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set P = Nothing
End Sub

Private Function ID() As Long
  Dim b() As String
  b = Split(List.Tag, ",")
  'Get the id out of the tag
  ID = b(List.ListIndex)
End Function
