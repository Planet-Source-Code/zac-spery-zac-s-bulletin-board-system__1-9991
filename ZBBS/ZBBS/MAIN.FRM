VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form main 
   Caption         =   "bbs"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6795
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   6795
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer CheckLogOff 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   1935
      Top             =   2835
   End
   Begin VB.Timer TimeSendData 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1215
      Top             =   2745
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000C000&
      Height          =   2670
      Left            =   45
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   6720
   End
   Begin MSWinsockLib.Winsock socketListen 
      Left            =   405
      Top             =   2745
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock socket 
      Index           =   0
      Left            =   0
      Top             =   2745
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const NodeCount = 4
Dim Node(NodeCount) As Object

Private Sub Stat(din As String)
    Text1.Text = Text1.Text + din
End Sub
Private Sub StatLn(din As String)
    Text1.Text = Text1.Text + din + vbCrLf
End Sub

Sub InitializeBBS()
    Dim i As Integer
    Stat "Creating 1 system node: "
    socketListen.LocalPort = 23
    socketListen.Close
    StatLn "done"
    Stat "Creating " + CStr(NodeCount) + " usernodes: "
    For i = 1 To NodeCount
        Load socket(i)
        socket(i).Close
        socket(i).RemotePort = 0
        Set Node(i) = CreateObject("zbbsnode.class1")
    Next i
    StatLn "done"
End Sub

Private Sub CheckLogOff_Timer()
    Dim i As Integer, strTemp As String
    For i = 1 To NodeCount
        If Node(i).logoff Then
            StatLn "NODE" + CStr(i) + " Connection terminated: " + Node(i).logoffreason
            socket(i).Close
            Node(i).Reset
        End If
    Next i
End Sub

Private Sub Form_Load()
    Me.Show
    DoEvents
    StatLn "ZBBS Starting..."
    InitializeBBS
    socketListen.Listen
    TimeSendData.Enabled = True
    CheckLogOff.Enabled = True
    StatLn "ZBBS listening on port 23"
End Sub
Private Function GetNextNode() As Integer
    Dim i As Integer
    For i = 1 To NodeCount
        If socket(i).State = 0 Then
            GetNextNode = i
            Exit For
        End If
    Next i
End Function

Private Sub socket_Close(Index As Integer)
    StatLn "NODE" + CStr(Index) + " Connection terminated: User terminated"
    socket(Index).Close
    Node(Index).Reset
End Sub

Private Sub socket_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim strTemp As String
    socket(Index).GetData strTemp
    Node(Index).rx strTemp
    'Text1.Text = Text1.Text + strTemp
End Sub

Private Sub socket_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    socket(Index).Close
    StatLn "NODE" + CStr(Index) + " Error #" + CStr(Number) + " " + Description + ":Node#" + CStr(Index) + " shut down."
    Node(Index).Reset
End Sub

Private Sub socketListen_ConnectionRequest(ByVal requestID As Long)
    Dim openport As Integer, i As Integer
    
    Stat "Connection requested: "
    openport = GetNextNode
    If openport = 0 Then
        socket(0).Accept requestID
        For i = 1 To 10000
            DoEvents
        Next i
        socket(0).SendData "No more nodes available.  Try again later"
        For i = 1 To 10000
            DoEvents
        Next i
        
        socket(0).Close
        StatLn "refused"
    Else
        socket(openport).Accept requestID
        Node(openport).run (openport)
        StatLn "accepted on node " + CStr(openport)
    End If
End Sub

Private Sub socketListen_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    StatLn "System node ERROR #" + CStr(Number) + " " + Description
End Sub

Private Sub Text1_Change()
    Text1.SelStart = Len(Text1.Text)
    If Len(Text1.Text) > 64000 Then
        Text1.Text = Right$(Text1.Text, 32000)
    End If
End Sub

Private Sub TimeSendData_Timer()
    Dim i As Integer, strTemp As String
    For i = 1 To NodeCount
        If socket(i).State <> 0 Then
             strTemp = Node(i).tx
            If StrComp(strTemp, "", vbBinaryCompare) <> 0 Then
                socket(i).SendData strTemp
            End If
        End If
        'If Node(i).logoff Then
        '    StatLn "NODE" + CStr(i) + " Connection terminated: " + Node(i).logoffreason
        '    socket(i).Close
        '    Node(i).Reset
        'End If
    Next i
End Sub
