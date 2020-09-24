Attribute VB_Name = "modClient"
Option Explicit

Public Sub Connect()
    On Error Resume Next
    
    ' connect to the host...
    frmMain.wskClient.Close
    frmMain.wskClient.Connect frmMain.txtHost.Text, frmMain.txtPort.Text
    DoEvents
End Sub
