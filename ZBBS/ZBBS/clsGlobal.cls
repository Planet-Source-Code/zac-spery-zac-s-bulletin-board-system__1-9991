VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGlobal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function GetUsers() As String
    Dim strTemp As String
    Dim i As Integer
    For i = 1 To 255
        If StrComp(users(i), "", vbBinaryCompare) <> 0 Then
            strTemp = strTemp + CStr(i) + vbTab + users(i) + vbTab
        End If
    Next i
    GetUsers = strTemp
End Function
Public Sub AddUser(username As String, node As Integer)
    cAddUser username, node
End Sub
Public Sub DeleteUser(username As String, node As Integer)
    cDeleteUser username, node
End Sub

Private Sub Class_Initialize()
    Dim i As Integer
    For i = 1 To 255
        users(i) = ""
    Next i
End Sub
