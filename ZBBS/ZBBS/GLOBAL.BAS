Attribute VB_Name = "global"
Option Explicit

Global users(255) As String
Sub cAddUser(username As String, node As Integer)
    'Dim i As Integer
    'For i = 1 To 255
    '    If StrComp(users(i), "", vbTextCompare) = 0 Then
            users(node) = username
    '        Exit For
    '    End If
    'Next i
End Sub
Sub cDeleteUser(username As String, node As Integer)
    'Dim i As Integer
    'For i = 1 To 255
    '    If StrComp(users(i), username, vbTextCompare) = 0 Then
            users(node) = ""
    '        Exit For
    '    End If
    'Next i
End Sub

