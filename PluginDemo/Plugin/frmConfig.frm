VERSION 5.00
Begin VB.Form frmConfig 
   Caption         =   "Form2"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Config"
      Height          =   1695
      Left            =   480
      TabIndex        =   0
      Top             =   720
      Width           =   3135
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  SendEventToPlugin 1, "Dada", "L"
  BroadCastMessage "HELLO EVRYBODY", "GG"
  SendEventToApp "Ready to close", "Give me some time..."
  MsgBox PluginCom.GetPluginData.Soort
  ClosePlugin
End Sub
