Attribute VB_Name = "Module1"
Global Const Info = "Test"
Global Const Soort = "Background"

Public Sub RunEvent(EventName As String, params As String)
  If EventName = "StartFrm" Then
    Form1.Show
  Else
    MsgBox EventName & " " & params, vbOKOnly, "PLUGIN"
  End If
End Sub
