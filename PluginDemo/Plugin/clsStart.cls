VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'****
'**** Interface Host
'****

'aan te maken constanten, variablen, forms & subs:
'Constanten:
'- Info: info over plugin
'
'Variablen:
'- Soort: bekijk informatie programma
'
'Forms:
'- FrmConfig : Configuratievenster Plugin
'
'Subs:
'- RunEvent(EventName as string, Params as string) : Events te ontvangen van de Host
'
'
'
'opgelet, bij het laden van het programma staat de plugin standaard aan, gelieve de plugin terug uit te schakelen indien het niet nodig is da de plugin aanstaat

Public Function Infoplug() As String
  Infoplug = Info
End Function

Public Function PlugSoort() As String
  PlugSoort = Soort
End Function

Public Sub Init(MainForm As Object, PluginCom As Object, IDL As Long)
   Set MainObject = MainForm
   Set PluginInterface = PluginCom
   ID = IDL
End Sub

Public Sub RunEvents(EventName As String, params As String)
  RunEvent EventName, params
End Sub

Public Function Config() As Object
  Set Config = New frmConfig
End Function

Public Sub Class_Terminate()
  Set PluginInterface = Nothing
  Set MainObject = Nothing
  Dim f As Form
  For Each f In Forms
    Unload f
  Next f
End Sub
