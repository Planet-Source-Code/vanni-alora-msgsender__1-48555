Attribute VB_Name = "modServer"
Option Explicit

Public Sub StartServer()
    On Error GoTo er
    
    ' set the server to for listening...
    frmMain.wskServer(0).Close
    frmMain.wskServer(0).Bind frmMain.txtPort.Text
    frmMain.wskServer(0).Listen
    Exit Sub
er:
    ' error handler...
    MsgBox Err.Description, , Err.Source
End Sub
