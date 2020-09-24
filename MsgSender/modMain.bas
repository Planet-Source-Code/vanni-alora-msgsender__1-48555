Attribute VB_Name = "modMain"
Option Explicit

' define some constant variables...
Public Const PORT As Long = 2323

Public Sub FormDefault()
    On Error Resume Next
    
    ' set the form to default...
    frmMain.txtHost.Text = frmMain.wskServer(0).LocalIP
    frmMain.txtPort.Text = PORT
    frmMain.txtMsg.Text = "Hello, World!"
    
    frmMain.txtHost.SetFocus
End Sub

