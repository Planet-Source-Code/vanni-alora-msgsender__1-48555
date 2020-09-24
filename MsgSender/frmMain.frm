VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Message Sender v1.0"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5415
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   5415
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSendMsg 
      Caption         =   "Send MSG"
      Default         =   -1  'True
      Height          =   375
      Left            =   3840
      TabIndex        =   6
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox txtMsg 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   1080
      Width           =   4695
   End
   Begin VB.TextBox txtPort 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2520
      TabIndex        =   3
      Text            =   "65535"
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox txtHost 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   360
      TabIndex        =   1
      Text            =   "127.0.0.1"
      Top             =   360
      Width           =   1935
   End
   Begin MSWinsockLib.Winsock wskClient 
      Left            =   4080
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock wskServer 
      Index           =   0
      Left            =   3600
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label3 
      Caption         =   "Message:"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Port:"
      Height          =   255
      Left            =   2520
      TabIndex        =   2
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "IP Address:"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'#Region " ReadMe First... "
    ' Author:   Vanni Alora
    ' E-Mail:   vanjo08@msn.com
    ' WebSite:  http://www.skyalpha.net

    ' Information provided by this program is intended solely to
    ' provide general guidance on matters of interest for the personal
    ' use of the user of this program, who accepts full responsibility
    ' for its use.

    ' This program provided "as-is", without any express or implied
    ' warranty in no event will the author be held liable for any
    ' damages arising from the use of it. You may copy and/or
    ' distribute this program (or any portion of it) in any way you
    ' may find it useful.

    ' The author may have retained certain copyrights to this code...
    ' please observe their request and the law by reviewing all
    ' copyright conditions at the above URL.
'#End Region

Option Explicit

Private Sub Form_Load()
    On Error Resume Next
    
    ' call the FormDefault method...
    modMain.FormDefault ' see modMain.bas
    
    ' call the StartListen method...
    modServer.StartServer ' see modServer.bas
End Sub

Private Sub cmdSendMsg_Click()
    On Error Resume Next
    
    ' connect to the server...
    modClient.Connect ' see modClient.bas
    DoEvents
    ' send the message with a unique send string...
    wskClient.SendData txtMsg.Text
    Exit Sub
er:
    'error handler...
    MsgBox Err.Description, , Err.Source
End Sub

Private Sub wskClient_Close()
    On Error Resume Next
    
    ' if we disconnect from the server...
    ' close the socket and unload the winsock object...
    wskClient.Close
    Unload wskClient
End Sub

Private Sub wskServer_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    On Error Resume Next
    
    Dim nIndex As Integer
    ' for any connection request, then do this...
    If Index = 0 Then
        nIndex = wskServer.UBound + 1
        ' load a new winsock object...
        Load wskServer(nIndex)
        ' accept the connection...
        wskServer(nIndex).Accept (requestID)
    End If
End Sub

Private Sub wskServer_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    On Error Resume Next
    
    Dim dataStr As String
    
    ' get the data...
    wskServer(Index).GetData dataStr, , bytesTotal
    
    ' display the message with the client's ip and time/date...
    MsgBox "Message from " & wskServer(Index).RemoteHostIP & " (" & Now & ")" & vbCrLf & vbCrLf & dataStr, , "Message"
    ' disconnect him...
    wskServer(Index).Close
    ' unload the winsock object...
    Unload wskServer(Index)
End Sub
