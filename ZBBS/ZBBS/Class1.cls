VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim WithEvents node1 As Node
Attribute node1.VB_VarHelpID = -1
Public LogOff As Boolean
Public LogOffReason As String


Private Sub Class_Initialize()
    Set node1 = New Node
End Sub

Private Sub Class_Terminate()
    Set node1 = Nothing
    End
End Sub
Public Sub reset()
    node1.LogOff = True
    node1.OutCache = ""
    node1.InCache = ""
    LogOff = False
    LogOffReason = ""
    node1.startup.Enabled = False
    node1.reset
End Sub


Public Sub Run(NodeNum As Integer)
    node1.NodeIndex = NodeNum
    node1.startup.Enabled = True
End Sub

Public Sub Rx(din As String)
    'do not call locally
    node1.SetCache din
    'node1.Text1.Text = node1.Text1.Text + "'" + din + "'"
End Sub
Public Function Tx() As String
    'do not call locally
    Tx = node1.OutCache
    'node1.Text1.Text = node1.Text1.Text + node1.OutCache
    node1.OutCache = ""
End Function


Private Sub node1_LogOff(reason As String)
    LogOffReason = reason
    LogOff = True
End Sub

Private Sub node1_LogOffUser(reason As String)
    LogOffReason = reason
    LogOff = True

End Sub
