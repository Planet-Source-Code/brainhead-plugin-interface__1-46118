Attribute VB_Name = "PluginCom"
Public MainObject As Object
Public PluginInterface As Object
Public ID As Long

Public Sub ClosePlugin()
  PluginInterface.SluitPlugin (ID)
End Sub

Public Sub BroadCastMessage(Message As String, Optional params As String)
  PluginInterface.BroadCastEvent Message, params
End Sub

Public Sub SendEventToApp(EventName As String, Optional params As String)
  PluginInterface.EventToApp ID, EventName, params
End Sub

Public Sub SendEventToPlugin(PlugID As Long, EventName As String, Optional params As String)
  PluginInterface.SendEvent PlugID, EventName, params
End Sub

Public Function GetPluginData() As Object
  Set GetPluginData = PluginInterface.PluginData(ID)
End Function
