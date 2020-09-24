VERSION 5.00
Begin VB.Form Node 
   Caption         =   "1"
   ClientHeight    =   2895
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7170
   LinkTopic       =   "Form1"
   ScaleHeight     =   2895
   ScaleWidth      =   7170
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.TextBox Text1 
      Height          =   1815
      Left            =   90
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   945
      Width           =   6945
   End
   Begin VB.Timer CallBack 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   810
      Top             =   405
   End
   Begin VB.Timer startup 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   180
      Top             =   225
   End
End
Attribute VB_Name = "Node"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const DARKBLACK = "0;30" 'black
Private Const DARKRED = "0;31"
Private Const DARKGREEN = "0;32"
Private Const DARKYELLOW = "0;33"
Private Const DARKBLUE = "0;34"
Private Const DARKPURPLE = "0;35"
Private Const DARKCYAN = "0;36"
Private Const DARKGRAY = "0;37" 'dark gray

Private Const LIGHTBLACK = "1;30" 'light gray
Private Const LIGHTRED = "1;31"
Private Const LIGHTGREEN = "1;32"
Private Const LIGHTYELLOW = "1;33"
Private Const LIGHTBLUE = "1;34"
Private Const LIGHTPURPLE = "1;35"
Private Const LIGHTCYAN = "1;36"
Private Const LIGHTGRAY = "1;37" 'white

Private Const BACKBLACK = ";40m"
Private Const BACKRED = ";41m"
Private Const BACKGREEN = ";42m"
Private Const BACKYELLOW = ";43m"
Private Const BACKBLUE = ";44m"
Private Const BACKPURPLE = ";45m"
Private Const BACKCYAN = ";46m"
Private Const BACKGRAY = ";47m"

'Private C1 As String, c2 As String, c3 As String, c4 As String, c5 As String, c6 As String
'Private c7 As String, c8 As String, c9 As String
Private esc As String

Public InCache As String
Public InCacheInUse As Boolean
Public OutCache As String
Public OutCacheInUse As Boolean

Public LogOff As Boolean
Dim IsPaused As Boolean
Public NodeIndex As Integer


Dim Stack As StackType
Dim CurrMenu As String
Dim CurrSubmenu As String
Dim CurrPrompt As String
Dim CurrCallBack As String
Dim InCall As Boolean
Dim AcceptableKeys As String
Dim Ansi As Boolean

Dim zbbsObj As Object

Dim Con As New ADODB.Connection
Dim MsgCon As New ADODB.Connection
Dim dbRs As New ADODB.Recordset

Dim LoggedOn As Boolean
Dim TimeOfLastKey As Double

'poor coding variables
Dim bFirstTime As Boolean
Dim bTimeWarning As Boolean

Private Type StackType
    Menu As String
    submenu As String
    CallBack As String
End Type
Private Type MsgBaseType
    name As String
    filename As String
    flags As String
    groups As String
    public As Boolean
    private As Boolean
    network As Boolean
End Type
Dim CurrBase As MsgBaseType
Private Type MsgType
    from As String
    to As String
    fromid As String
    toid As String
    dateposted As String
    title As String
    deleted As Boolean
    msg_id As Long
    'msg_txt As String
End Type
Dim CurrMsg As MsgType
Dim CurrMsgList() As MsgType
Dim CurrMsgText As String

Private Type UserType
    userid As String
    alias As String
    real As String
    groups As String
    flags As String
    laston As String
    firston As String
    status As String
    password As String
    sex As String
    timeon As Integer
    timeused As Integer
    totaltime As Integer
    timeperday As Integer
    occupation As String
    reference As String
    sysopnote As String
    basetime As String
    basedate As String
End Type
Dim CurrUser As UserType

Private Type MenuType
    name As String
    index As String
    key As String
    description As String
    command As String
    hotkey As Boolean
    groups As String
    flags As String
    fallback As String
    prompt As String
    help_level As String
    help_file As String
    extra As String
    clear As Boolean
    add As Boolean
    delete As Boolean
End Type
Dim Menu() As MenuType


'Public Event FormSend(din As String)
Public Event LogOffUser(reason As String)
Public Sub reset()
    SaveUser
    bTimeWarning = False
    CallBack.Enabled = False
    CurrMenu = ""
    CurrSubmenu = ""
    CurrPrompt = ""
    CurrCallBack = ""
    InCall = False
    LogOff = False
    ReDim Menu(1)
    If dbRs.State <> 0 Then dbRs.Close
End Sub

Public Sub SetCache(din As String)
    InCache = InCache + din
End Sub

Private Function GetData() As String
    GetData = InCache
    InCache = ""
End Function

Private Sub SendData(din As String)
    OutCache = OutCache + din
    
End Sub

Private Sub Push()
    Stack.Menu = CurrMenu
    Stack.submenu = CurrSubmenu
    Stack.CallBack = CurrCallBack
End Sub
Private Sub Pop()
    CurrMenu = Stack.Menu
    CurrSubmenu = Stack.submenu
    CurrCallBack = Stack.CallBack
End Sub


Private Function RepeatChar(strChar As String, intTimes As Integer) As String
    Dim strTemp As String
    Dim i As Integer
    For i = 1 To intTimes
        strTemp = strTemp + strChar
    Next i
    RepeatChar = strTemp
End Function
Private Sub clr()
    If Ansi Then
        send Chr$(12)
    Else
        sendln ""
    End If
End Sub
Private Function LeftJustify(strTxt As String, strChar As String, intSpace As Integer) As String
    Dim strTemp As String
    Dim i As Integer
    'strTxt = filtercodes(strTxt)
    For i = 1 To intSpace - Len(filtercodes(strTxt))
        strTemp = strTemp + strChar
    Next i
    LeftJustify = strTemp + strTxt
End Function
Private Function FillSpace(str1 As String, strChar As String, intSpace As Integer) As String
    Dim strTemp As String
    Dim i As Integer
    strTemp = str1
    For i = 1 To intSpace - Len(filtercodes(str1))
        strTemp = strTemp + strChar
    Next i
    FillSpace = strTemp
End Function
Private Function filtercodes(strDin As String) As String
    Dim nStart As Long, nEnd As Long
    Dim strTemp As String
    nStart = 1
    nEnd = InStr(1, strDin, "|", vbBinaryCompare)
    Do Until nEnd = 0
        strTemp = strTemp + Mid$(strDin, nStart, nEnd - nStart)
        nStart = nEnd + 3
        nEnd = InStr(nStart, strDin, "|", vbBinaryCompare)
    Loop
    strTemp = strTemp + Mid$(strDin, nStart, Len(strDin) - nStart + 1)
    filtercodes = strTemp
End Function
Private Function Filt(strDin As String) As String
    Dim nStart As Long, nEnd As Long
    Dim strTemp As String, strCommand As String, strSubCommand As String
    Dim strTimeLeft As String, strHours As String, strMins As String
    nStart = 1
    nEnd = InStr(1, strDin, "|", vbBinaryCompare)
    Do Until nEnd = 0
        strTemp = strTemp + Mid$(strDin, nStart, nEnd - nStart)
        strCommand = Mid$(strDin, nEnd + 1, 2)
        strSubCommand = Right$(strCommand, 1)
        Select Case Left$(strCommand, 1)
            Case "U"
                Select Case strSubCommand
                    Case "T"
                        strTimeLeft = Format$(DateAdd("n", CurrUser.timeperday - (CurrUser.timeon + CurrUser.timeused), "00:00"), "hhmm")
                        strHours = Left$(strTimeLeft, 2)
                        strMins = Right$(strTimeLeft, 2)
                        If strHours <> "00" Then
                            strHours = Format$(strHours, "0") + " hrs "
                        Else
                            strHours = ""
                        End If
                        If strMins <> "00" Then
                            strMins = Format$(strMins, "0") + " min(s)"
                        Else
                            strMins = ""
                        End If
                        
                        If strHours = "00" And strMins = "00" Then
                            strTimeLeft = "0 mins"
                        Else
                            strTimeLeft = strHours + strMins
                        End If
                        strTemp = strTemp + strTimeLeft
                End Select
            Case "S"
                If Ansi Then
                    Select Case strSubCommand
                        Case "0"
                            strTemp = strTemp + esc + DARKGRAY + BACKBLACK
                        Case "1"
                            strTemp = strTemp + esc + LIGHTBLACK + BACKBLACK
                        Case "2"
                            strTemp = strTemp + esc + LIGHTGRAY + BACKBLACK
                        Case "3"
                            strTemp = strTemp + esc + DARKRED + BACKBLACK
                        Case "4"
                            strTemp = strTemp + esc + LIGHTRED + BACKBLACK
                        Case "5"
                            strTemp = strTemp + esc + DARKBLUE + BACKBLACK
                        Case "6"
                            strTemp = strTemp + esc + LIGHTBLUE + BACKBLACK
                        Case "7"
                            strTemp = strTemp + esc + DARKYELLOW + BACKBLACK
                        Case "8"
                            strTemp = strTemp + esc + LIGHTYELLOW + BACKBLACK
                        Case "9"
                            strTemp = strTemp + esc + DARKPURPLE + BACKBLACK
                        Case "A"
                            strTemp = strTemp + esc + LIGHTPURPLE + BACKBLACK
                        Case "B"
                            strTemp = strTemp + esc + DARKCYAN + BACKBLACK
                        Case "C"
                            strTemp = strTemp + esc + LIGHTCYAN + BACKBLACK
                        Case "D"
                            strTemp = strTemp + esc + DARKGREEN + BACKBLACK
                        Case "E"
                            strTemp = strTemp + esc + LIGHTGREEN + BACKBLACK
                        Case "F"
                            strTemp = strTemp + esc + LIGHTBLUE + BACKBLUE
                        Case "G"
                            strTemp = strTemp + esc + LIGHTCYAN + BACKBLUE
                        Case "H"
                            strTemp = strTemp + esc + LIGHTGREEN + BACKBLUE
                        
                    End Select
                Else
                    strTemp = strTemp
                End If
                    
            Case "C"
                Select Case strSubCommand
                    Case "R"
                        strTemp = strTemp + vbCrLf
                End Select
            Case Else
                Filt = strDin
                Exit Function
        End Select
        nStart = nEnd + 3
        nEnd = InStr(nStart, strDin, "|", vbBinaryCompare)
    Loop
    strTemp = strTemp + Mid$(strDin, nStart, Len(strDin) - nStart + 1)
    Filt = strTemp
End Function


Private Sub sendln(din As String)
    SendData Filt(din) + vbCrLf
End Sub
Private Sub send(din As String)
    If StrComp(din, Chr$(8), vbBinaryCompare) = 0 Then
        din = Chr$(8) + " " + Chr$(8)
    End If
    SendData Filt(din)
End Sub

Private Function anykey() As String
    anykey = GetData
End Function
Private Function OneKey(AcceptChars As String) As String
    Dim strTemp As String
    strTemp = GetData
    If InStr(1, AcceptChars, strTemp, vbTextCompare) > 0 Then
        If StrComp(strTemp, "", vbBinaryCompare) <> 0 Then
            OneKey = strTemp
            TimeOfLastKey = Timer
        End If
    End If
End Function
Private Sub pause()
    send "Press any key to continue"
    IsPaused = True
End Sub

Private Function prompt(AcceptChars As String, Length As Integer, echo As Boolean, echochar As String) As String
    Dim strTemp2 As String
    Dim strTemp As String
    Dim i As Integer
    Static CurrString As String
    strTemp2 = GetData
    For i = 1 To Len(strTemp2)
        strTemp = Mid$(strTemp2, i, 1)
        If StrComp(strTemp, "", vbBinaryCompare) <> 0 Then
            'Text1.Text = Text1.Text + "'" + strtemp + "'"
            TimeOfLastKey = Timer
            If StrComp(strTemp, Chr$(13), vbBinaryCompare) = 0 Then
                If StrComp(CurrString, "", vbBinaryCompare) = 0 Then
                    prompt = Chr$(13)
                Else
                    prompt = CurrString
                End If
                CurrString = ""
                Exit Function
            End If
            If InStr(1, AcceptChars + " " + Chr$(8), strTemp, vbTextCompare) <> 0 Then
                'Text1.Text = "acceptbale"
                If Len(CurrString) < 1 And StrComp(strTemp, Chr$(8), vbBinaryCompare) = 0 Then
                    'donothing
                ElseIf strTemp = Chr$(8) Then
                    CurrString = Left$(CurrString, Len(CurrString) - 1)
                    send Chr$(8)
                ElseIf Len(CurrString) < Length Then
                    If echo Then
                        If StrComp(echochar, "", vbBinaryCompare) <> 0 Then
                            send echochar
                            CurrString = CurrString + strTemp
                        Else
                            send strTemp
                            CurrString = CurrString + strTemp
                        End If
                    End If
                    
                End If
            End If
        End If
    
    Next i
End Function
Private Function GetRSVal(strField As String) As String
    If Not IsNull(dbRs.Fields(strField).Value) Then
        GetRSVal = CStr(Trim(dbRs.Fields(strField).Value))
    Else
        GetRSVal = ""
    End If
End Function
Private Sub LoadUser(strUserId As String)
    'Dim temprs As ADODB.Recordset
    dbRs.Open "select * from usr_man where usr_id = " + strUserId, Con, adOpenStatic, adLockBatchOptimistic
    'On Error Resume Next
    CurrUser.alias = GetRSVal("usr_man_alias")
    CurrUser.firston = GetRSVal("usr_man_first_on")
    CurrUser.laston = GetRSVal("usr_man_last_on")
    CurrUser.groups = GetRSVal("usr_man_usr_grp")
    CurrUser.flags = GetRSVal("usr_man_usr_flg")
    CurrUser.real = GetRSVal("usr_man_real")
    CurrUser.status = GetRSVal("usr_man_stat")
    CurrUser.sex = GetRSVal("usr_man_sex")
    CurrUser.totaltime = GetRSVal("usr_man_tot_tim")
    CurrUser.timeused = GetRSVal("usr_man_tim_left")
    CurrUser.timeperday = GetRSVal("usr_man_tim_per_day")
    CurrUser.sysopnote = GetRSVal("usr_man_sys_not")
    CurrUser.occupation = GetRSVal("usr_man_occ")
    CurrUser.reference = GetRSVal("usr_man_ref")
    If DateDiff("d", CurrUser.laston, CStr(Date) + " " + CStr(Time)) >= 1 Then
        CurrUser.timeused = 0
        dbRs.Fields("usr_man_tim_left").Value = 0
    End If
    dbRs.Fields("usr_man_last_on").Value = CDate(Date + Time)
    dbRs.UpdateBatch
    If CInt(CurrUser.timeused) >= CurrUser.timeperday Then
        'kill user
        DoEvents
    End If
    dbRs.Close
End Sub
Private Sub LoadMenu(strMenu As String)
    Dim strTemp As String
    Dim index As Integer
    If dbRs.State <> 0 Then dbRs.Close
    dbRs.Open "select * from mnu_man where mnu_nm = '" + strMenu + "' order by mnu_idx", Con, adOpenStatic, adLockReadOnly
    If dbRs.RecordCount = 0 Then Exit Sub
    ReDim Menu(dbRs.RecordCount - 1)
    dbRs.MoveFirst
    AcceptableKeys = ""
    Do Until dbRs.EOF
        index = CInt(dbRs.Fields("mnu_idx").Value)
        Menu(index).name = GetRSVal("mnu_nm")
        Menu(index).index = GetRSVal("mnu_idx")
        strTemp = GetRSVal("mnu_hot")
        If StrComp(strTemp, "y", vbTextCompare) = 0 Then
            Menu(index).hotkey = True
        Else
            Menu(index).hotkey = False
        End If
        Menu(index).description = GetRSVal("mnu_dsc")
        Menu(index).command = GetRSVal("mnu_cmd")
        Menu(index).key = GetRSVal("mnu_key")
        If index <> 0 And Menu(index).key <> "FIRSTCMD" Then
            AcceptableKeys = AcceptableKeys + Menu(index).key
        End If
        Menu(index).groups = GetRSVal("mnu_grp")
        Menu(index).flags = GetRSVal("mnu_flg")
        Menu(index).fallback = GetRSVal("mnu_fal_bck")
        Menu(index).prompt = GetRSVal("mnu_prompt")
        Menu(index).help_file = GetRSVal("mnu_hlp_fil")
        Menu(index).help_level = CInt(GetRSVal("mnu_hlp_lvl"))
        Menu(index).extra = GetRSVal("mnu_str")
        
        strTemp = GetRSVal("mnu_clr")
        If StrComp(strTemp, "y", vbTextCompare) = 0 Then
            Menu(index).clear = True
        Else
            Menu(index).clear = False
        End If
        dbRs.MoveNext
    Loop
    dbRs.Close
    CurrMenu = Menu(0).name
End Sub
Private Sub DisplayMenu(strMenu As String)
    Dim i As Integer, strTemp As String, i2 As Integer
    'Dim TempMenu() As MenuType
    If StrComp(strMenu, Menu(0).name, vbTextCompare) <> 0 Then
        LoadMenu (strMenu)
    End If
    For i = 1 To UBound(Menu)
        If Menu(i).key = "FIRSTCMD" Then
           Call RunCommand(Menu(i).command, i)
        End If
    Next i
    sendln ""
    If StrComp(Menu(0).help_file, "", vbBinaryCompare) <> 0 Then
        DisplayFile Menu(0).help_file
    Else
        If Menu(0).help_level <> 1 Then
            If Menu(0).clear Then clr
            sendln "|SG" + FillSpace(Menu(0).description, " ", 67) + "? for help "
            sendln "|SB" + RepeatChar("=", 78) + "|S2"
            i2 = 0
            For i = 1 To UBound(Menu)
                If StrComp(Menu(i).key, "FIRSTCMD", vbBinaryCompare) <> 0 _
                        And StrComp(Menu(i).command, "?", vbTextCompare) <> 0 Then
                    'strtemp = FillSpace("  ", " ", 2) + Menu(i).key
                    i2 = i2 + 1
                    strTemp = LeftJustify("|S8(|S4" + Menu(i).key, " ", 3) + "|S8)"
                    strTemp = "|S2" + FillSpace(strTemp, " ", 6) + Menu(i).description
                    If i2 Mod 2 <> 0 Then
                        send FillSpace(strTemp, " ", 37)
                    Else
                        sendln strTemp
                    End If
                End If
            Next i
            'send strTemp
            If i2 Mod 2 <> 0 Then
                sendln ""
            End If
            sendln "|SB" + RepeatChar("=", 78)
        End If
        'sendln ""
    End If
    If bFirstTime = True Then
        send Menu(0).prompt
        bFirstTime = False
    End If
End Sub
Private Sub DisplayFile(strFile As String)
    Dim strTemp As String
    Dim FileNumber As Integer
    FileNumber = FreeFile
    Open strFile For Input Access Read Shared As FileNumber
    strTemp = Input(FileLen(strFile), FileNumber)
    Close FileNumber
    send strTemp
End Sub
Private Sub LoadMessageBase(strBase As String)
    Dim i As Integer, strTemp As String
    If dbRs.State <> 0 Then dbRs.Close
    dbRs.Open "select * from msg_lst where msg_lst_nm = '" + strBase + "'", Con, adOpenStatic, adLockReadOnly
    'dbRs.MoveFirst
    CurrBase.filename = GetRSVal("msg_lst_fn")
    CurrBase.flags = GetRSVal("msg_lst_usr_flg")
    CurrBase.groups = GetRSVal("msg_lst_usr_grp")
    CurrBase.name = GetRSVal("msg_lst_nm")
    If GetRSVal("msg_lst_net") = "y" Then
        CurrBase.network = True
    Else
        CurrBase.network = False
    End If
    If GetRSVal("msg_lst_pri") = "y" Then
        CurrBase.private = True
    Else
        CurrBase.private = False
    End If
    If GetRSVal("msg_lst_pub") = "y" Then
        CurrBase.public = True
    Else
        CurrBase.public = False
    End If
    dbRs.Close
    If MsgCon.State <> 0 Then MsgCon.Close
    MsgCon.Open "PROVIDER=Microsoft.Jet.OLEDB.3.51;Data Source=c:\zbbs\msg\" + CurrBase.filename + ".mdb;"
    If CurrBase.private Then
        dbRs.Open "select * from msg_man where msg_man_del = 'n' And msg_man_frm_id = " + CStr(CurrUser.userid), MsgCon, adOpenStatic, adLockReadOnly
    Else
        dbRs.Open "select * from msg_man where msg_man_del = 'n'", MsgCon, adOpenStatic, adLockReadOnly
    End If
    dbRs.MoveFirst
    ReDim CurrMsgList(dbRs.RecordCount)
    i = 1
    Do Until dbRs.EOF
        CurrMsgList(i).msg_id = GetRSVal("msg_id")
        strTemp = GetRSVal("msg_man_del")
        If strTemp = "y" Then
            CurrMsgList(i).deleted = True
        Else
            CurrMsgList(i).deleted = False
        End If
        CurrMsgList(i).dateposted = GetRSVal("msg_man_dt")
        CurrMsgList(i).from = GetRSVal("msg_man_frm")
        CurrMsgList(i).to = GetRSVal("msg_man_to")
        CurrMsgList(i).fromid = GetRSVal("msg_man_frm_id")
        CurrMsgList(i).toid = GetRSVal("msg_man_to_id")
        CurrMsgList(i).title = GetRSVal("msg_man_title")
        i = i + 1
        dbRs.MoveNext
    Loop
    dbRs.Close
End Sub
Private Sub DisplayMessageList()
    Dim i As Integer, strTemp As String
    sendln " #   From             To                Title"
    sendln RepeatChar("=", 80)
    For i = 1 To UBound(CurrMsgList)
        'sendln CStr(i) + "    " + CurrMsgList(i).from + "            " + CurrMsgList(i).to + "           " + CurrMsgList(i).title
        strTemp = LeftJustify(CStr(i), " ", 2)
        strTemp = strTemp + "   " + CurrMsgList(i).from
        strTemp = FillSpace(strTemp, " ", 22) + CurrMsgList(i).to
        strTemp = FillSpace(strTemp, " ", 40) + CurrMsgList(i).title
        sendln strTemp
    Next i
    sendln RepeatChar("=", 80)
End Sub
Private Sub DisplayMessageBaseList()
    Dim i As Integer
    Dim strTemp As String
    If dbRs.State <> 0 Then dbRs.Close
    dbRs.Open "select * from msg_lst", Con, adOpenStatic, adLockReadOnly
    dbRs.MoveFirst
    clr
    i = 1
    sendln "Num" + RepeatChar(" ", 35) + "Base Name"
    sendln RepeatChar("=", 60)
    Do Until dbRs.EOF
        If GetRSVal("msg_lst_pri") <> "y" Then
            'strTemp = CStr(i) + RepeatChar(" ", 30) + GetRSVal("msg_lst_dsc")
            strTemp = LeftJustify(CStr(i), " ", 3)
            strTemp = FillSpace(strTemp, " ", 38) + GetRSVal("msg_lst_dsc")
            sendln strTemp
            i = i + 1
        End If
        dbRs.MoveNext
    Loop
    dbRs.Close
    sendln RepeatChar("=", 60)
    
End Sub
Private Sub DisplayMenus()
    Dim strTemp As String, strTemp2 As String
    Dim i As Integer, i2 As Integer
    If dbRs.State <> 0 Then dbRs.Close
    dbRs.Open "select mnu_nm from mnu_man where mnu_idx = 0", Con, adOpenStatic, adLockReadOnly
    dbRs.MoveFirst
    clr
    sendln "|SG      Name                     Name                     Name                  "
    sendln "|SB" + RepeatChar("=", 78)
    i = 1
    i2 = 1
    Do Until dbRs.EOF
        'strTemp = LeftJustify(CStr(i), " ", 3)
        strTemp = ""
        
        strTemp = FillSpace(strTemp, " ", 6) + GetRSVal("mnu_nm")
        strTemp2 = strTemp2 + FillSpace(strTemp, " ", 25)
        If i2 = 3 Then
            sendln strTemp2
            strTemp2 = ""
            i2 = 0
        End If
        i2 = i2 + 1
        i = i + 1
        dbRs.MoveNext
    Loop
    If strTemp2 <> "" Then
        sendln strTemp2
    End If
    dbRs.Close
    sendln "|SB" + RepeatChar("=", 78)
End Sub

Private Sub DisplayUsers()
    Dim temp As String, i As Integer
    Dim nStart As Long, nEnd As Long
    Dim strTemp As String, strTemp2 As String
    
    temp = zbbsObj.getusers
    sendln "|S2Users currently logged in:"
    sendln FillSpace("", "=", 78)
    nStart = 1
    nEnd = InStr(1, temp, vbTab, vbBinaryCompare)
    Do Until nEnd = 0
        'sendln LeftJustify(CStr(NodeIndex), " ", 4) + "   " + Mid$(temp, nStart, nEnd - nStart)
        strTemp = Mid$(temp, nStart, nEnd - nStart)
        nStart = nEnd + 1
        nEnd = InStr(nStart, temp, vbTab, vbBinaryCompare)
        strTemp2 = Mid$(temp, nStart, nEnd - nStart)
        sendln LeftJustify(strTemp, " ", 4) + "   " + strTemp2
        nStart = nEnd + 1
        nEnd = InStr(nStart, temp, vbTab, vbBinaryCompare)
        
    Loop
    sendln FillSpace("", "=", 78)
End Sub

Private Sub RunCommand(strCommand As String, FirstCmd As Integer)
    Dim i As Integer
    Dim menuindex As Integer
    Dim strExtra As String
    If FirstCmd = 0 Then
        For i = 1 To UBound(Menu)
            If StrComp(Menu(i).key, strCommand, vbTextCompare) = 0 Then
                menuindex = i
                Exit For
            End If
        Next i
    Else
        menuindex = FirstCmd
        i = 1
    End If
    If i <> 0 Then
        Select Case Menu(menuindex).command
            Case "AS"
                Call DisplayUserStats
            Case "NL"
                Call DisplayUsers
            Case "FD"
                'file display
                strExtra = Menu(menuindex).extra
                Call DisplayFile(strExtra)
                
            Case "MD"
                'display message base list
                Call DisplayMessageBaseList
                
            Case "ML"
                'load message room
                strExtra = Menu(menuindex).extra
                Call LoadMessageBase(strExtra)
                
            Case "MM"
                'display message list
                Call DisplayMessageList
                
            Case "#L"
                'display menus in menueditor
                Call DisplayMenus
            Case "#E"
                Push
                SetCallBack "editmenu", "", ""
                Exit Sub
            Case "#A"
                Push
                sendln ""
                send "|S2Name :"
                SetCallBack "editmenu", "addmenu", ""
                Exit Sub
            Case "#D"
                Push
                sendln ""
                send "|S2Which Menu :"
                SetCallBack "editmenu", "removemenu", ""
                Exit Sub
            Case "GM"
                'load menu
                strExtra = Menu(menuindex).extra
                Call DisplayMenu(strExtra)
                
            Case "GF"
                'load fallback menu
                strExtra = Menu(0).fallback
                Call DisplayMenu(strExtra)
                
            Case "?"
                'display current menu
                Call DisplayMenu(CurrMenu)
                
            Case "LG"
                'logoff
                Call BootUser("User logged off")
                Exit Sub
            Case Else
                sendln ""
                sendln "Unknown command"
                sendln ""
                CallBack.Enabled = True
        End Select
        If FirstCmd = 0 Then
            'sendln ""
            CallBack.Enabled = True
            send Menu(0).prompt
        End If
    Else
        CallBack.Enabled = True
        Exit Sub
    End If
        
End Sub

Private Sub CallBack_Menu()
    'Static strInput$
    Dim strTemp$
    If StrComp(Menu(0).name, CurrMenu, vbTextCompare) <> 0 Then
        DisplayMenu (CurrMenu)
        CallBack.Enabled = True
    Else
        If Menu(0).hotkey <> True Then
            strTemp = prompt("abcedefghijklmnopqrstuvwxyz0123456789 " + AcceptableKeys, 8, True, "")
        Else
            strTemp = OneKey(AcceptableKeys)
            If StrComp(strTemp, "", vbBinaryCompare) <> 0 Then
                sendln strTemp
                TimeOfLastKey = Timer
            End If
            
        End If
        If StrComp(strTemp, "", vbBinaryCompare) <> 0 Then
            Call RunCommand(strTemp, 0)
        Else
            CallBack.Enabled = True
        End If
        
    End If
End Sub

Private Sub CallBack_Login()
    Static strAlias$, strPassword$
    Dim strTemp$
    
    Select Case CurrSubmenu
        Case "askansi"
            strTemp = LCase(OneKey("yn" + Chr$(13)))
            If strTemp <> "" Then
                If strTemp = "y" Then
                    Ansi = True
                    sendln ""
                    sendln "|S4Ansi Color|S0 enabled."
                    sendln ""
                    
                Else
                    Ansi = False
                    sendln ""
                    sendln "Standard ASCII enabled."
                    sendln ""
                End If
                SetSubMenu "start"
                Exit Sub
            Else
                CallBack.Enabled = True
                Exit Sub
            End If
        Case "start"
            sendln "|SDNew users type '|SENEW|SD'"
            sendln ""
            send "|S0Username|S6: |S0"
            SetSubMenu "username"
            Exit Sub

        Case "username"
            strTemp = prompt("abcdefghijklmnopqrstuvwxyz0123456789 " + Chr$(13), 25, True, "")
            If Trim$(LCase(strTemp)) = "new" Then
                sendln ""
                sendln "That function is not yet supported."
                sendln ""
                SetSubMenu "start"
                Exit Sub
            End If
            If strTemp <> "" Then
                strAlias = strTemp
                sendln ""
                
                sendln ""
                send "|S0Password|S6: |S0"
                SetSubMenu "password"
                Exit Sub
            Else
                CallBack.Enabled = True
                Exit Sub
            End If
        Case "password"
            strTemp = prompt("abcdefghijklmnopqrstuvwxyz0123456789 " + Chr$(13), 25, True, "x")
            If strTemp <> "" Then
                strPassword = strTemp
                
                dbRs.Open "select usr_id, usr_psw from usr_lku where usr_psw = '" + strAlias + "'", Con, adOpenStatic, adLockReadOnly
                If dbRs.RecordCount = 0 Then
                    dbRs.Close
                    sendln ""
                    sendln "|S4Invalid login": sendln ""
                    send "|S0Username|S6: |S0"
                    SetSubMenu "username"
                    Exit Sub
                End If
                dbRs.MoveFirst
                CurrUser.userid = dbRs.Fields(0).Value
                CurrUser.password = dbRs.Fields(1).Value
                dbRs.Close
                If StrComp(CurrUser.password, strPassword, vbTextCompare) = 0 Then
                    LoggedOn = True
                    LoadUser CurrUser.userid
                    bTimeWarning = False
                    zbbsObj.adduser CurrUser.alias, NodeIndex
                    DisplayUserStats
                    pause
                    SetCallBack "menu", "main", ""
                Else
                    sendln ""
                    sendln "|S4Invalid login": sendln ""
                    send "|S0Username|S6: |S0"
                    SetSubMenu "username"
                    Exit Sub
                End If
            Else
                CallBack.Enabled = True
                Exit Sub
            End If
        
        Case Else
            sendln ""
            send "Ansi terminal emulation(y/N)?"
            SetSubMenu "askansi"
    End Select
End Sub
Private Sub DisplayUserStats()
    clr
    sendln "|S2[|S4Current User Statistics|S3]|S6"
    sendln "|S6" + RepeatChar("=", 79) + "|S3"
    sendln "|S2" + FillSpace("Username", " ", 35) + "|S7: |SC" + CurrUser.alias
    sendln "|S2" + FillSpace("Real Name", " ", 35) + "|S7: |SC" + CurrUser.real
    sendln "|S2" + FillSpace("Last on", " ", 35) + "|S7: |SC" + CurrUser.laston
    sendln "|S2" + FillSpace("First on", " ", 35) + "|S7: |SC" + CurrUser.firston
    sendln "|S2" + FillSpace("Time used", " ", 35) + "|S7: |SC" + CStr(CurrUser.timeused)
    sendln "|S2" + FillSpace("Time per day", " ", 35) + "|S7: |SC" + CStr(CurrUser.timeperday)
    sendln "|S2" + FillSpace("Total time used", " ", 35) + "|S7: |SC" + CStr(CurrUser.totaltime)
    sendln "|S2" + FillSpace("Sex", " ", 35) + "|S7: |SC" + CurrUser.sex
    sendln "|S6" + RepeatChar("=", 79)
    sendln ""

End Sub

Private Sub savemenu(emenu() As MenuType)
    Dim i As Integer
    If dbRs.State <> 0 Then dbRs.Close
    dbRs.Open "select * from mnu_man where mnu_nm = '" + emenu(1).name + "' order by mnu_idx", Con, adOpenStatic, adLockBatchOptimistic
    'For i = 1 To dbRs.RecordCount
    '    dbRs.delete
    'Next i
    'dbRs.MoveFirst
    Do Until dbRs.EOF
        dbRs.delete
        dbRs.MoveNext
    Loop
    
    For i = 1 To UBound(emenu)
        'If dbRs.EOF Then dbRs.AddNew
        dbRs.AddNew
        dbRs.Fields("mnu_nm") = emenu(1).name
        dbRs.Fields("mnu_idx") = CStr(i - 1) 'emenu(i).index
        dbRs.Fields("mnu_key") = emenu(i).key
        dbRs.Fields("mnu_dsc") = emenu(i).description
        dbRs.Fields("mnu_cmd") = emenu(i).command
        If emenu(i).hotkey = True Then
            dbRs.Fields("mnu_hot") = "y"
        Else
            dbRs.Fields("mnu_hot") = "n"
        End If
        dbRs.Fields("mnu_grp") = emenu(i).groups
        dbRs.Fields("mnu_flg") = emenu(i).flags
        dbRs.Fields("mnu_fal_bck") = emenu(i).fallback
        dbRs.Fields("mnu_hlp_fil") = emenu(i).help_file
        If emenu(i).help_level = "" Then emenu(i).help_level = "0"
        dbRs.Fields("mnu_hlp_lvl") = CByte(emenu(i).help_level)
        dbRs.Fields("mnu_str") = emenu(i).extra
        If emenu(i).clear = True Then
            dbRs.Fields("mnu_clr") = "y"
        
        Else
            dbRs.Fields("mnu_clr") = "n"
        End If
        
        'dbRs.Fields("mnu_clr") = emenu(i).clear
        dbRs.Fields("mnu_prompt") = emenu(i).prompt
        'dbRs.MoveNext
    Next i
    'dbRs.MoveFirst
    
    dbRs.UpdateBatch
    dbRs.Close
End Sub

Private Sub Callback_editmenu()
    Dim strTemp As String
    Static emenu() As MenuType
    Static bEditItems As Boolean
    Dim i As Integer, i2 As Integer
    Static ItemNum As Integer
    Static ItemInput As String
    Select Case CurrMenu
        Case "removemenu"
            strTemp = prompt("abcdefghijklmnopqrstuvwxyz0123456789", 20, True, "")
            If strTemp = Chr$(13) Then
                Pop
                Menu(0).name = ""
                bFirstTime = True
                CallBack.Enabled = True
                Exit Sub
            End If
            If strTemp <> "" Then
                If dbRs.State <> 0 Then dbRs.Close
                dbRs.Open "select * from mnu_man where mnu_nm = '" + strTemp + "'", Con, adOpenStatic, adLockBatchOptimistic
                If dbRs.RecordCount = 0 Then
                    sendln ""
                    sendln "Menu not found"
                    Pop
                    Menu(0).name = ""
                    bFirstTime = True
                    CallBack.Enabled = True
                    Exit Sub
                End If
                dbRs.MoveFirst
                Do Until dbRs.EOF
                    dbRs.delete
                    dbRs.MoveNext
                Loop
                'For i = 0 To dbRs.RecordCount - 1
                '    dbRs.Move i
                '    dbRs.Delete
                'Next i
                dbRs.UpdateBatch
                dbRs.Close
                Pop
                Menu(0).name = ""
                bFirstTime = True
                CallBack.Enabled = True
                Exit Sub
                
            End If
            CallBack.Enabled = True
            Exit Sub
        Case "addmenu"
            strTemp = prompt("abcdefghijklmnopqrstuvwxyz0123456789", 20, True, "")
            If strTemp = Chr$(13) Then
                Pop
                Menu(0).name = ""
                bFirstTime = True
                CallBack.Enabled = True
                Exit Sub
            End If
            If strTemp <> "" Then
                If dbRs.State <> 0 Then dbRs.Close
                strTemp = UCase(strTemp)
                dbRs.Open "select * from mnu_man where mnu_nm = '" + strTemp + "'", Con, adOpenDynamic, adLockPessimistic
                If dbRs.RecordCount > 0 Then
                    dbRs.Close
                    sendln ""
                    sendln "Name already in use"
                    Pop
                    Menu(0).name = ""
                    bFirstTime = True
                    CallBack.Enabled = True
                    Exit Sub
                End If
                ReDim emenu(3)
                emenu(1).clear = True
                emenu(1).name = strTemp
                emenu(1).command = ""
                emenu(1).description = "New Menu"
                emenu(1).extra = ""
                emenu(1).fallback = "MAIN"
                emenu(1).flags = ""
                emenu(1).groups = ""
                emenu(1).help_file = ""
                emenu(1).help_level = 3
                emenu(1).hotkey = True
                emenu(1).index = 0
                emenu(1).key = ""
                emenu(1).prompt = "New Menu :"
                emenu(2).key = "Q"
                emenu(2).description = "Quit"
                emenu(2).hotkey = True
                emenu(2).command = "GF"
                emenu(2).index = 1
                emenu(2).name = strTemp
                emenu(3).key = "?"
                emenu(3).description = "Help"
                emenu(3).hotkey = True
                emenu(3).command = "?"
                emenu(3).index = 2
                emenu(3).name = strTemp
                For i = 1 To 3
                    dbRs.AddNew
                    dbRs.Fields("mnu_nm") = emenu(1).name
                    dbRs.Fields("mnu_idx") = CStr(i - 1) 'emenu(i).index
                    dbRs.Fields("mnu_key") = emenu(i).key
                    dbRs.Fields("mnu_dsc") = emenu(i).description
                    dbRs.Fields("mnu_cmd") = emenu(i).command
                    If emenu(i).hotkey = True Then
                        dbRs.Fields("mnu_hot") = "y"
                    Else
                        dbRs.Fields("mnu_hot") = "n"
                    End If
                    dbRs.Fields("mnu_grp") = emenu(i).groups
                    dbRs.Fields("mnu_flg") = emenu(i).flags
                    dbRs.Fields("mnu_fal_bck") = emenu(i).fallback
                    dbRs.Fields("mnu_hlp_fil") = emenu(i).help_file
                    If emenu(i).help_level = "" Then emenu(i).help_level = "0"
                    dbRs.Fields("mnu_hlp_lvl") = CByte(emenu(i).help_level)
                    dbRs.Fields("mnu_str") = emenu(i).extra
                    If emenu(i).clear = True Then
                        dbRs.Fields("mnu_clr") = "y"
                    Else
                        dbRs.Fields("mnu_clr") = "n"
                    End If
                    dbRs.Fields("mnu_prompt") = emenu(i).prompt
                Next i
                dbRs.Update
                dbRs.Close
                Pop
                Menu(0).name = ""
                bFirstTime = True
                CallBack.Enabled = True
                Exit Sub
                
            End If
            CallBack.Enabled = True
            Exit Sub
        Case "selectmenu"
            strTemp = prompt("abcdefghijklmnopqrstuvwxyz0123456789", 20, True, "")
            If strTemp = Chr$(13) Then
                Pop
                Menu(0).name = ""
                bFirstTime = True
                CallBack.Enabled = True
                Exit Sub
            End If
            If strTemp <> "" Then
                If dbRs.State <> 0 Then dbRs.Close
                dbRs.Open "select * from mnu_man where mnu_nm = '" + strTemp + "' order by mnu_idx"
                If dbRs.RecordCount = 0 Then
                    dbRs.Close
                    sendln ""
                    sendln "menu not found"
                    SetMenu ""
                    Exit Sub
                End If
                ReDim emenu(dbRs.RecordCount)
                dbRs.MoveFirst
                emenu(1).name = GetRSVal("mnu_nm")
                strTemp = GetRSVal("mnu_clr")
                If strTemp = "y" Then
                    emenu(1).clear = True
                Else
                    emenu(1).clear = False
                End If
                
                'EMenu(1).command = GetRSVal("mnu_cmd")
                emenu(1).description = GetRSVal("mnu_dsc")
                emenu(1).extra = GetRSVal("mnu_str")
                emenu(1).fallback = GetRSVal("mnu_fal_bck")
                emenu(1).flags = GetRSVal("mnu_flg")
                emenu(1).groups = GetRSVal("mnu_grp")
                emenu(1).help_file = GetRSVal("mnu_hlp_fil")
                emenu(1).help_level = GetRSVal("mnu_hlp_lvl")
                strTemp = GetRSVal("mnu_hot")
                If strTemp = "y" Then
                    emenu(1).hotkey = True
                Else
                    emenu(1).hotkey = False
                End If
                emenu(1).index = GetRSVal("mnu_idx")
                emenu(1).key = GetRSVal("mnu_key")
                emenu(1).prompt = GetRSVal("mnu_prompt")
                dbRs.MoveNext
                i = 2
                Do Until dbRs.EOF
                    emenu(i).command = GetRSVal("mnu_cmd")
                    emenu(i).description = GetRSVal("mnu_dsc")
                    emenu(i).extra = GetRSVal("mnu_str")
                    emenu(i).flags = GetRSVal("mnu_flg")
                    emenu(i).groups = GetRSVal("mnu_grp")
                    emenu(i).key = GetRSVal("mnu_key")
                    i = i + 1
                    dbRs.MoveNext
                Loop
                SetMenu "HeadDisplay"
            
            Else
                CallBack.Enabled = True
                Exit Sub
            End If
        'Case "ListItems"
        '    clr
        '    sendln "|S2Num  Key       Descript            Cmd   Param           Grp           Flg"
        '    sendln "|SG" + FillSpace("", "=", 79)
        '    For i = 1 To UBound(emenu)
        '        strTemp = "|S2" + LeftJustify(CStr(i), " ", 2) _
        '            + "    " + FillSpace(emenu(i).key, " ", 9)
        '        strTemp = FillSpace(strTemp, " ", 15) + emenu(i).description
        '        strTemp = FillSpace(strTemp, " ", 34) + " " + emenu(i).command
        '        strTemp = FillSpace(strTemp, " ", 40) + " " + emenu(i).extra
        '        strTemp = FillSpace(strTemp, " ", 55) + " " + emenu(i).groups
        '        strTemp = FillSpace(strTemp, " ", 70) + " " + emenu(2).flags
        '        sendln strTemp
        '
        '    Next i
        '    sendln "|SG" + FillSpace("", "=", 79)
        '    sendln "|S2Enter number to edit, Q to quit, T to edit menu header, S to save menu"
        '    send "Command : "
        '    SetMenu "HeadCommand"
        '    Exit Sub
            
        Case "HeadDisplay"
            clr
            If bEditItems = False Then
                sendln "|SBMenu Name           : " + emenu(1).name
                sendln "|SG" + FillSpace("", "=", 79)
                sendln "|S2" + "(A) Menu Description    : " + emenu(1).description
                sendln "|S2" + "(B) Toggle hotkeys      : " + CStr(emenu(1).hotkey)
                sendln "|S2" + "(C) Help Level          : " + emenu(1).help_level
                sendln "|S2" + "(D) Help File           : " + emenu(1).help_file
                sendln "|S2" + "(E) Toggle Clear        : " + CStr(emenu(1).clear)
                sendln "|S2" + "(F) Fallback Menu       : " + emenu(1).fallback
                sendln "|S2" + "(G) Groups              : " + emenu(1).groups
                sendln "|S2" + "(H) Flags               : " + emenu(1).flags
                sendln "|S2" + "(I) Prompt              : " + vbCrLf + emenu(1).prompt
                sendln "|SG" + FillSpace("", "=", 79)
                sendln "|S2Enter letter to edit, Q to quit, T to edit menu items, S to save menu"
            Else
                sendln "|S2Num  Key       Descript                      Cmd   Param                            "
                sendln "|SG" + FillSpace("", "=", 79)
                For i = 2 To UBound(emenu)
                    If emenu(i).add Then
                        strTemp = "|S2A"
                    ElseIf emenu(i).delete Then
                        strTemp = "|S2D"
                    Else
                        strTemp = "|S2 "
                    End If
                    strTemp = strTemp + LeftJustify(CStr(i - 1), " ", 2) _
                        + "  " + FillSpace(emenu(i).key, " ", 7)
                    strTemp = FillSpace(strTemp, " ", 15) + emenu(i).description
                    strTemp = FillSpace(strTemp, " ", 44) + " " + emenu(i).command
                    strTemp = FillSpace(strTemp, " ", 50) + " " + emenu(i).extra
                    'strTemp = FillSpace(strTemp, " ", 55) + " " + emenu(i).groups
                    'strTemp = FillSpace(strTemp, " ", 70) + " " + emenu(2).flags
                    sendln strTemp
                    
                Next i
                sendln "|SG" + FillSpace("", "=", 79)
                sendln "|S2Enter number to edit, Q to quit, T to edit menu header, S to save menu"
                sendln "|S2A to Add or D to Delete items"
            End If
            send "Command : "
            SetMenu "HeadCommand"
            Exit Sub
        Case "EditItem"
            Select Case CurrSubmenu
                Case "ItemPrompt"
                    strTemp = prompt("abcdefghijklmnopqrstuvwxyz 0123456789()*&^%$#@!" + vbCr, 50, True, "")
                    If strTemp = "" Then
                        CallBack.Enabled = True
                        Exit Sub
                    End If
                    sendln ""
                    Select Case strTemp
                        Case "a"
                            sendln "New Description : "
                        Case "b"
                            sendln "New Key : "
                        Case "c"
                            sendln "New Command : "
                        Case "d"
                            sendln "New Extra : "
                        Case "e"
                            sendln "New Groups : "
                        Case "f"
                            sendln "New Extra : "
                    End Select
                    SetSubMenu LCase(strTemp)
                    Exit Sub
                Case "a"
                    strTemp = prompt("abcdefghijklmnopqrstuvwxyz 0123456789()*&^%$#@!" + vbCr, 50, True, "")
                    If strTemp = "" Then
                        CallBack.Enabled = True
                        Exit Sub
                    End If
                    If strTemp = vbCr Then
                        SetSubMenu ""
                        Exit Sub
                    End If
                
                    emenu(ItemNum).description = strTemp
                    SetSubMenu ""
                Case "b"
                    strTemp = prompt("abcdefghijklmnopqrstuvwxyz 0123456789()*&^%$#@!" + vbCr, 50, True, "")
                    If strTemp = "" Then
                        CallBack.Enabled = True
                        Exit Sub
                    End If
                    If strTemp = vbCr Then
                        SetSubMenu ""
                        Exit Sub
                    End If
                    emenu(ItemNum).key = strTemp
                    SetSubMenu ""
                Case "c"
                    strTemp = prompt("abcdefghijklmnopqrstuvwxyz 0123456789()*&^%$#@!" + vbCr, 50, True, "")
                    If strTemp = "" Then
                        CallBack.Enabled = True
                        Exit Sub
                    End If
                    If strTemp = vbCr Then
                        SetSubMenu ""
                        Exit Sub
                    End If
                    emenu(ItemNum).command = strTemp
                    SetSubMenu ""
                Case "d"
                    strTemp = prompt("abcdefghijklmnopqrstuvwxyz 0123456789()*&^%$#@!" + vbCr, 50, True, "")
                    If strTemp = "" Then
                        CallBack.Enabled = True
                        Exit Sub
                    End If
                    If strTemp = vbCr Then
                        SetSubMenu ""
                        Exit Sub
                    End If
                    emenu(ItemNum).extra = strTemp
                    SetSubMenu ""
                Case "e"
                    strTemp = prompt("abcdefghijklmnopqrstuvwxyz 0123456789()*&^%$#@!" + vbCr, 50, True, "")
                    If strTemp = "" Then
                        CallBack.Enabled = True
                        Exit Sub
                    End If
                    If strTemp = vbCr Then
                        SetSubMenu ""
                        Exit Sub
                    End If
                    emenu(ItemNum).groups = strTemp
                    SetSubMenu ""
                Case "f"
                    strTemp = prompt("abcdefghijklmnopqrstuvwxyz 0123456789()*&^%$#@!" + vbCr, 50, True, "")
                    If strTemp = "" Then
                        CallBack.Enabled = True
                        Exit Sub
                    End If
                    If strTemp = vbCr Then
                        SetSubMenu ""
                        Exit Sub
                    End If
                    emenu(ItemNum).flags = strTemp
                    SetSubMenu ""
                Case "q"
                    SetMenu "HeadDisplay"
                    Exit Sub
                    
                Case Else
                    clr
                    sendln "|SG" + FillSpace("", "=", 79)
                    sendln "|S2(A)Description      :" + emenu(ItemNum).description
                    sendln "|S2(B)Key              :" + emenu(ItemNum).key
                    sendln "|S2(C)Command          :" + emenu(ItemNum).command
                    sendln "|S2(D)Extra            :" + emenu(ItemNum).extra
                    sendln "|S2(E)Group            :" + emenu(ItemNum).groups
                    sendln "|S2(F)Flags            :" + emenu(ItemNum).flags
                    sendln "|SG" + FillSpace("", "=", 79)
                    sendln "|S2Enter letter to edit.  Enter Q when done."
                    send "Command : "
                    SetSubMenu "ItemPrompt"
            End Select
        Case "HeadCommand"
            Select Case CurrSubmenu
                Case "DeleteWhich"
                    strTemp = prompt("0123456789" + vbCr + Chr$(13), 2, True, "")
                    If strTemp <> "" Then
                        If strTemp = vbCr Then
                            SetSubMenu ""
                            Exit Sub
                        End If
                        If IsNumeric(strTemp) Then
                            If CInt(strTemp) > UBound(emenu) Then
                                SetSubMenu ""
                                Exit Sub
                            End If
                            i2 = CInt(strTemp) + 1
                            For i = i2 To UBound(emenu)
                                If i <> UBound(emenu) Then
                                    emenu(i).clear = emenu(i + 1).clear
                                    emenu(i).command = emenu(i + 1).command
                                    emenu(i).description = emenu(i + 1).description
                                    emenu(i).extra = emenu(i + 1).extra
                                    emenu(i).flags = emenu(i + 1).flags
                                    emenu(i).fallback = emenu(i + 1).fallback
                                    emenu(i).groups = emenu(i + 1).groups
                                    emenu(i).help_file = emenu(i + 1).help_file
                                    emenu(i).help_level = emenu(i + 1).help_level
                                    emenu(i).hotkey = emenu(i + 1).hotkey
                                    emenu(i).key = emenu(i + 1).key
                                    emenu(i).name = emenu(i + 1).name
                                    emenu(i).prompt = emenu(i + 1).prompt
                                    emenu(i).index = CStr(i2)
                                End If
                            Next i
                            ReDim Preserve emenu(UBound(emenu) - 1)
                            SetMenu "HeadDisplay"
                            SetSubMenu ""
                            Exit Sub
                            
                        End If
                    End If
                    CallBack.Enabled = True
                    Exit Sub
                Case "desc"
                    strTemp = prompt("abcdefghijklmnopqrstuvwxyz0123456789!@#$%^&*()[];:'"",.<>/?" + Chr$(13), 50, True, "")
                    If strTemp <> "" Then
                        If strTemp <> Chr$(13) Then
                            emenu(1).description = strTemp
                        End If
                        SetMenu "HeadDisplay"
                        SetSubMenu ""
                        Exit Sub
                    End If
                    CallBack.Enabled = True
                    Exit Sub
                Case "file"
                    strTemp = prompt("abcdefghijklmnopqrstuvwxyz0123456789" + Chr$(13), 50, True, "")
                    'strTemp = prompt("abcdefghijklmnopqrstuvwxyz0123456789!@#$%^&*()[];:'"",.<>/?" + Chr$(13), 50, True, "")
                    If strTemp <> "" Then
                        If strTemp <> Chr$(13) Then
                            emenu(1).help_file = strTemp
                        End If
                        SetMenu "HeadDisplay"
                        SetSubMenu ""
                        Exit Sub
                    End If
                    CallBack.Enabled = True
                    Exit Sub
                Case "prompt"
                    strTemp = prompt("abcdefghijklmnopqrstuvwxyz0123456789!@#$%^&*()[];:'"",.<>/?|" + Chr$(13), 50, True, "")
                    If strTemp <> "" Then
                        If strTemp <> Chr$(13) Then
                            emenu(1).prompt = strTemp
                        End If
                        SetMenu "HeadDisplay"
                        SetSubMenu ""
                        Exit Sub
                    End If
                    CallBack.Enabled = True
                    Exit Sub
                Case "hotkey"
                    strTemp = OneKey("yn" + Chr$(13))
                    If strTemp <> "" Then
                        If strTemp <> Chr$(13) Then
                            If StrComp(strTemp, "n", vbTextCompare) = 0 Then
                                emenu(1).hotkey = False
                            Else
                                emenu(1).hotkey = True
                            End If
                            SetMenu "HeadDisplay"
                            SetSubMenu ""
                            Exit Sub
                        Else
                            emenu(1).hotkey = True
                        End If
                    End If
                    CallBack.Enabled = True
                    Exit Sub
                Case "level"
                    strTemp = OneKey("123" + Chr$(13))
                    If strTemp <> "" Then
                        If strTemp <> Chr$(13) Then
                            emenu(1).help_level = strTemp
                        End If
                        SetMenu "HeadDisplay"
                        SetSubMenu ""
                        Exit Sub
                        
                    End If
                    CallBack.Enabled = True
                    Exit Sub
                Case "clear"
                    strTemp = OneKey("yn" + Chr$(13))
                    If strTemp <> "" Then
                        If StrComp(strTemp, "y", vbTextCompare) = 0 Then
                            emenu(1).clear = True
                        Else
                            emenu(1).clear = False
                        End If
                        SetMenu "HeadDisplay"
                        SetSubMenu ""
                        Exit Sub
                        
                    End If
                    CallBack.Enabled = True
                    Exit Sub
                Case "menu"
                    strTemp = prompt("abcdefghijklmnopqrstuvwxyz0123456789_" + Chr$(13), 20, True, "")
                    If strTemp <> "" Then
                        emenu(1).fallback = strTemp
                        SetMenu "HeadDisplay"
                        SetSubMenu ""
                        Exit Sub
                    End If
                    CallBack.Enabled = True
                    Exit Sub
                Case "group"
                    strTemp = prompt("abcdefghijklmnopqrstuvwxyz0123456789_" + Chr$(13), 20, True, "")
                    If strTemp <> "" Then
                        emenu(1).groups = strTemp
                        SetMenu "HeadDisplay"
                        SetSubMenu ""
                        Exit Sub
                    End If
                    CallBack.Enabled = True
                    Exit Sub
                Case "flag"
                    strTemp = prompt("abcdefghijklmnopqrstuvwxyz0123456789_" + Chr$(13), 20, True, "")
                    If strTemp <> "" Then
                        emenu(1).flags = strTemp
                        SetMenu "HeadDisplay"
                        SetSubMenu ""
                        Exit Sub
                        
                    End If
                    CallBack.Enabled = True
                    Exit Sub
                Case "save"
                    strTemp = OneKey("yn" + Chr$(13))
                    If strTemp <> "" Then
                        If StrComp(strTemp, "y", vbTextCompare) = 0 Then
                            send "|C4Saving..."
                            savemenu emenu()
                            sendln "|C2Done"
                            SetMenu "HeadDisplay"
                            SetSubMenu ""
                            Exit Sub
                        End If
                    End If
                    CallBack.Enabled = True
                    Exit Sub
                Case Else
                    If bEditItems = False Then
                        strTemp = OneKey("abcdefghiqts")
                        If strTemp <> "" Then
                            Select Case LCase(strTemp)
                                Case "a"
                                    sendln ""
                                    send "New Desc: "
                                    SetSubMenu "desc"
                                    Exit Sub
                                Case "b"
                                    sendln ""
                                    send "Use hotkeys(Y/n)"
                                    SetSubMenu "hotkey"
                                    Exit Sub
                                Case "c"
                                    sendln ""
                                    send "New helplevel: "
                                    SetSubMenu "level"
                                    Exit Sub
                                Case "d"
                                    sendln ""
                                    send "File name: "
                                    SetSubMenu "file"
                                    Exit Sub
                                Case "e"
                                    sendln ""
                                    send "Clear screen before menu(Y/n)"
                                    SetSubMenu "clear"
                                    Exit Sub
                                Case "f"
                                    sendln ""
                                    send "New Fallback Menu: "
                                    SetSubMenu "menu"
                                    Exit Sub
                                Case "g"
                                    sendln ""
                                    send "Groups: "
                                    SetSubMenu "group"
                                    Exit Sub
                                Case "h"
                                    sendln ""
                                    send "Flags: "
                                    SetSubMenu "flag"
                                    Exit Sub
                                Case "i"
                                    sendln ""
                                    sendln "New Prompt"
                                    SetSubMenu "prompt"
                                    send ":"
                                    Exit Sub
                                Case "s"
                                    sendln ""
                                    sendln ""
                                    send "Save?"
                                    SetSubMenu "save"
                                    Exit Sub
                               
                                Case "q"
                                    Pop
                                    bFirstTime = True
                                    DisplayMenu (CurrMenu)
                                    CallBack.Enabled = True
                                    Exit Sub
                                Case "t"
                                    bEditItems = True
                                    SetMenu "HeadDisplay"
                                    CallBack.Enabled = True
                                    Exit Sub
                            End Select
                        End If
                        CallBack.Enabled = True
                    Else
                        
                        strTemp = prompt("1234567890stqad" + Chr$(13), 2, True, "")
                        If strTemp <> "" Then
                            Select Case LCase(strTemp)
                                Case "a"
                                    sendln ""
                                    i = UBound(emenu) + 1
                                    ReDim Preserve emenu(i)
                                    emenu(i).command = "NEW"
                                    emenu(i).key = "NEW"
                                    emenu(i).description = "New Command"
                                    emenu(i).extra = ""
                                    emenu(i).flags = ""
                                    emenu(i).groups = ""
                                    SetMenu "HeadDisplay"
                                    CallBack.Enabled = True
                                    Exit Sub
                                Case "d"
                                    sendln ""
                                    send "Delete which : "
                                    SetSubMenu "DeleteWhich"
                                    Exit Sub
                                Case "s"
                                    sendln ""
                                    sendln ""
                                    send "Save?"
                                    SetSubMenu "save"
                                    Exit Sub
                               
                                Case "q"
                                    Pop
                                    bFirstTime = True
                                    DisplayMenu (CurrMenu)
                                    CallBack.Enabled = True
                                    Exit Sub
                                Case "t"
                                    bEditItems = False
                                    SetMenu "HeadDisplay"
                                    CallBack.Enabled = True
                                    Exit Sub
                                Case Else
                                    If IsNumeric(strTemp) Then
                                        If CInt(strTemp + 1) > UBound(emenu) Then
                                            sendln ""
                                            sendln "Unknown command"
                                            SetMenu "HeadDisplay"
                                            CallBack.Enabled = True
                                            Exit Sub
                                        Else
                                            SetMenu "EditItem"
                                            ItemNum = CInt(strTemp + 1)
                                            CallBack.Enabled = True
                                            Exit Sub
                                        End If
                                    Else
                                        sendln ""
                                        sendln "Unknown command"
                                        SetMenu "HeadDisplay"
                                        CallBack.Enabled = True
                                        Exit Sub
                                    End If
                            End Select
                        End If
                        CallBack.Enabled = True
                    End If
            End Select
        Case Else
            bEditItems = False
            sendln ""
            send "Which menu: "
            SetMenu "selectmenu"
    End Select
End Sub

Private Sub SetCallBack(strCallBack As String, Menu As String, submenu As String)

    CurrCallBack = strCallBack
    CurrMenu = Menu
    CurrSubmenu = submenu
    CallBack.Enabled = True
End Sub
Private Sub SetMenu(Menu As String)
    CurrMenu = Menu
    CallBack.Enabled = True
End Sub
Private Sub SetSubMenu(submenu As String)
    CurrSubmenu = submenu
    CallBack.Enabled = True
End Sub


Private Sub CallBack_Timer()
    Dim strTemp As String, KeyDiff As Double
    Dim i As Integer, TimeTot As Integer
    CallBack.Enabled = False
    KeyDiff = Timer - TimeOfLastKey
    If KeyDiff > 300 Then
        sendln ""
        sendln "Inactivity Timeout.  Logging off"
        For i = 1 To 1000
            DoEvents
        Next i
        Call BootUser("Inactivity Timeout")
        Exit Sub
    
    End If
    If LogOff Then
        Exit Sub
    End If
    CurrUser.timeon = DateDiff("n", CurrUser.basetime, Time)
    TimeTot = CStr(CurrUser.timeon + CurrUser.timeused)
    If TimeTot > CurrUser.timeperday Then
        sendln ""
        sendln "Time limit exceeded.  Logging off"
        Call BootUser("Time limit exceeded")
        For i = 1 To 1000
            DoEvents
        Next i
        
        Exit Sub
    End If
    If bTimeWarning = False Then
        If CurrUser.timeperday - CurrUser.timeon < 5 Then
            bTimeWarning = True
            sendln ""
            sendln "Less than five minutes left!"
        End If
    End If
    If IsPaused Then
        strTemp = anykey
        If StrComp(strTemp, "", vbBinaryCompare) = 0 Then
            CallBack.Enabled = True
            Exit Sub
        Else
            TimeOfLastKey = Timer
            IsPaused = False
            'CallBack.Enabled = True
        End If
    End If
    
    Select Case CurrCallBack
        Case "login": CallBack_Login
        Case "menu": CallBack_Menu
        Case "editmenu": Callback_editmenu
    End Select
    
End Sub
Private Sub BootUser(strReason As String)
    zbbsObj.deleteuser CurrUser.alias, NodeIndex
    LogOff = True
    CallBack.Enabled = False
    SaveUser
    
    RaiseEvent LogOffUser(strReason)
End Sub

Private Sub SaveUser()
    If LoggedOn = True Then
        LoggedOn = False
        If dbRs.State <> 0 Then dbRs.Close
        dbRs.Open "select * from usr_man where usr_id = " + CurrUser.userid, Con, adOpenDynamic, adLockOptimistic
        
        dbRs.Fields("usr_man_tim_left").Value = CurrUser.timeon + CurrUser.timeused
        dbRs.Fields("usr_man_tot_tim").Value = CStr(CLng(GetRSVal("usr_man_tot_tim")) + CLng(CurrUser.timeon))
        dbRs.Update
        dbRs.Close
    End If
End Sub

Private Sub login()
    If zbbsObj Is Nothing Then
        Set zbbsObj = CreateObject("zbbsglobal.clsglobal")
    End If
    bTimeWarning = True
    bFirstTime = True
    TimeOfLastKey = Timer
    CurrUser.basetime = Time
    CurrUser.basedate = Format(Date, "yyyy-mm-dd")
    CurrUser.timeperday = "3"
    CurrUser.timeused = 0
    CurrUser.timeon = 0
    SetCallBack "login", "login", ""
'    CurrCallBack = "login"
'    CurrMenu = "login"
'    CurrSubmenu = ""
'    CallBack.Enabled = True
End Sub



Private Sub startup_Timer()
    Dim strUserName As String, strPassword As String
    startup.Enabled = False
    Dim i As Integer
    esc = Chr$(27) + "["
    
    'RaiseEvent FormSend("ZBBS " + vbCrLf + "Connected")
    'Set Con = CreateObject("adodb.connection")
    'Set dbRs = CreateObject("adodb.recordset")
    If Con.State = 0 Then Con.Open "dsn=zbbsdb"
    ReDim Menu(1)
    SendData vbCrLf + Chr$(254) + " ZBBS alpha" + vbCrLf + "Connected to node #" + CStr(NodeIndex) + vbCrLf
    'For i = 30 To 50
    'SendData Chr$(27) + "[41m"
    'SendData Chr$(27) + "[" + CStr(i) + "m  " + CStr(i) + "This is a test"
    'Next i
    Call login
    
End Sub


