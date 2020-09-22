VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PluginCom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents plug As plugin.Plugins
Attribute plug.VB_VarHelpID = -1

Private Sub Class_Initialize()
  Set plug = New plugin.Plugins
  'Where whil the plugins be?
  plug.Pluginpath = App.Path & "\Plugin\"
  'HoofdObject is the mainobject where the dll plugin has to comunicate with
  'Mostly this points to the object ME
  plug.Hoofdobject = Me
  'Search for plugins and Init them for later use.
  plug.SearchPlugins
End Sub

Private Sub Class_Terminate()
  Set plug = Nothing
End Sub

Private Sub plug_ErrorInPlugin(PlugInID As Long, Err As ErrObject, Procedure As String)
  MsgBox "Error in plugin: " & PlugInID & vbCrLf & Err.Description & "(" & Err.Number & ")" & vbCrLf & "Procedure: " & Procedure, vbOKOnly, "Error"
End Sub

Private Sub plug_EventFromPlugin(PlugInID As Long, EventName As String, params As String)
  MsgBox "Received An Event from the plugin: " & vbCrLf & "EventName: " & EventName & vbCrLf & "Parameters:" & params, vbOKOnly
End Sub

Private Sub plug_PluginFound(naam As String, ID As Long)
  PluginDemo.List.AddItem "Naam"
  'Save the id as a list in the tag because yoo need to know the ID of the plugin to use it
  PluginDemo.List.Tag = PluginDemo.List.Tag & CStr(ID) & ","
End Sub

Public Sub RightCal(ByRef name As String, P As Integer)
  name = InputBox("Give A word", "Test of MainApp", CStr(P))
End Sub
