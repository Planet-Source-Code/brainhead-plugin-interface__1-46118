VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4155
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5955
   LinkTopic       =   "Form1"
   ScaleHeight     =   4155
   ScaleWidth      =   5955
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "RightCall"
      Height          =   735
      Left            =   1320
      TabIndex        =   6
      Top             =   2760
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      Caption         =   "Send Event"
      Height          =   1815
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   4695
      Begin VB.TextBox EventName 
         Height          =   285
         Left            =   2160
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   360
         Width           =   2175
      End
      Begin VB.TextBox params 
         Height          =   285
         Left            =   2160
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   720
         Width           =   2175
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Send"
         Height          =   375
         Left            =   1200
         TabIndex        =   1
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Name:"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Parameters"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Width           =   1575
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  Dim N As String
  PluginCom.MainObject.RightCal N, 9
  MsgBox N
End Sub

Private Sub Command2_Click()
  PluginCom.SendEventToApp CStr(Me.EventName.Text), Me.params.Text
End Sub
