VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const Base64EmptyByte As Byte = 61
Private Const Base64EmptyWord As Integer = 61

Private mbytBase64EncodeByte(0 To 63) As Byte
Private mbytBase64EncodeWord(0 To 63) As Integer

Private Declare Function ArrPtr Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long '<-- VB6
Private Declare Sub PokeLng Lib "kernel32" Alias "RtlMoveMemory" (ByVal Addr As Long, Source As Long, Optional ByVal Bytes As Long = 4)

Private mstrSvrResponse As String
Private mlngResponseInternal As Long
Private mstrResponse As String
Private WithEvents mobjSck As Winsock
Attribute mobjSck.VB_VarHelpID = -1

Public Event SendMailBody(ByRef MessageInfo As String, ByRef Cancel As Boolean)
Public Event SendMailComplete()

Public Enum TextFormat

    vbTextPlain
    vbTextHtml
    
End Enum

Public Enum MailFormat

    vbMultipartAlternative
    vbMultipartMixed
    
End Enum

Private mobjNameSpace As Object
Private mobjApp As Object
Private mobjMAPISession As Object

Private Sub Base64EncodeByte(mInByte() As Byte, mOutByte() As Byte, Num As Integer)
    '------------------------------------------------------------------------------------------------------------------
    '功能：
    '------------------------------------------------------------------------------------------------------------------
    Dim tByte As Byte
    Dim i As Integer
    
    If Num = 1 Then
        mInByte(1) = 0
        mInByte(2) = 0
    ElseIf Num = 2 Then
        mInByte(2) = 0
    End If
    
    tByte = mInByte(0) And &HFC
    mOutByte(0) = tByte / 4
    tByte = ((mInByte(0) And &H3) * 16) + (mInByte(1) And &HF0) / 16
    mOutByte(1) = tByte
    tByte = ((mInByte(1) And &HF) * 4) + ((mInByte(2) And &HC0) / 64)
    mOutByte(2) = tByte
    tByte = (mInByte(2) And &H3F)
    mOutByte(3) = tByte
    
    For i = 0 To 3
        If mOutByte(i) >= 0 And mOutByte(i) <= 25 Then
            mOutByte(i) = mOutByte(i) + Asc("A")
        ElseIf mOutByte(i) >= 26 And mOutByte(i) <= 51 Then
            mOutByte(i) = mOutByte(i) - 26 + Asc("a")
        ElseIf mOutByte(i) >= 52 And mOutByte(i) <= 61 Then
            mOutByte(i) = mOutByte(i) - 52 + Asc("0")
        ElseIf mOutByte(i) = 62 Then
            mOutByte(i) = Asc("+")
        Else
            mOutByte(i) = Asc("/")
        
        End If
    Next i
    
    If Num = 1 Then
        mOutByte(2) = Asc("=")
        mOutByte(3) = Asc("=")
    ElseIf Num = 2 Then
        mOutByte(3) = Asc("=")
    End If
End Sub

Public Property Set WinSockObj(ByVal vData As Winsock)
    Set mobjSck = vData
End Property

Public Property Let Response(ByVal vData As String)
    mstrResponse = vData
End Property

Public Property Get Response() As String
    Response = mstrResponse
End Property

Public Property Let ResponseInternal(ByVal vData As Long)
    
    mlngResponseInternal = vData
    
End Property


Private Sub Base64Init()
    '------------------------------------------------------------------------------------------------------------------
    '功能：建立Base64码数组
    '------------------------------------------------------------------------------------------------------------------

  Const Chars64 As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
  Static i As Long
  Dim Code As Integer

  If i Then Exit Sub

  For i = 0 To 63
    Code = Asc(Mid$(Chars64, i + 1, 1))
    mbytBase64EncodeByte(i) = Code
    mbytBase64EncodeWord(i) = Code
  Next i
  
End Sub

Private Static Function Base64EncodeString(ByRef Text As String) As String
    '------------------------------------------------------------------------------------------------------------------
    '功能：Base64码转换函数
    '------------------------------------------------------------------------------------------------------------------

    Dim Chars() As Integer
    Dim SavePtr As Long
    Dim SADescrPtr As Long
    Dim DataPtr As Long
    Dim CountPtr As Long
    Dim TextLen As Long
    Dim i As Long
    Dim Chars64() As Integer
    Dim SavePtr64 As Long
    Dim SADescrPtr64 As Long
    Dim DataPtr64 As Long
    Dim CountPtr64 As Long
    Dim TextLen64 As Long
    Dim j As Long
    Dim b1 As Integer
    Dim b2 As Integer
    Dim b3 As Integer
    
    On Error GoTo errHand
    
    j = 0
    
    TextLen = Len(Text)
    If TextLen = 0 Then Exit Function
    TextLen64 = ((TextLen + 2) \ 3) * 4
    
    Base64EncodeString = Space$(TextLen64)
    
    If SavePtr = 0 Then
        ReDim Chars(1 To 1)
        
        SavePtr = VarPtr(Chars(1))
        PokeLng VarPtr(SADescrPtr), ByVal ArrPtr(Chars)
        DataPtr = SADescrPtr + 12
        CountPtr = SADescrPtr + 16
        
        ReDim Chars64(0 To 0)
        
        SavePtr64 = VarPtr(Chars64(0))
        PokeLng VarPtr(SADescrPtr64), ByVal ArrPtr(Chars64)
        DataPtr64 = SADescrPtr64 + 12
        CountPtr64 = SADescrPtr64 + 16
    End If

    PokeLng DataPtr, StrPtr(Text)
    PokeLng CountPtr, TextLen
    PokeLng DataPtr64, StrPtr(Base64EncodeString)
    PokeLng CountPtr64, TextLen64
    
    
    Call Base64Init

    '输入字符串转换为Base64码
    For i = 1 To TextLen - 2 Step 3
        b1 = Chars(i)
        b2 = Chars(i + 1)
        b3 = Chars(i + 2)
                
        Chars64(j) = mbytBase64EncodeWord(b1 \ &H4)
        Chars64(j + 1) = mbytBase64EncodeWord((b1 And &H3) * &H10 + b2 \ &H10)
        Chars64(j + 2) = mbytBase64EncodeWord((b2 And &HF) * &H4 + b3 \ &H40)
        Chars64(j + 3) = mbytBase64EncodeWord(b3 And &H3F)
        
        j = j + 4
    Next i

    '继续将未转换完的输入字符串转换为Base64码
    Select Case TextLen - i
    Case 0 '2 Bytes
        b1 = Chars(i)
        Chars64(j) = mbytBase64EncodeWord(b1 \ &H4)
        Chars64(j + 1) = mbytBase64EncodeByte((b1 And &H3) * &H10)
        Chars64(j + 2) = Base64EmptyWord
        Chars64(j + 3) = Base64EmptyWord
    Case 1 '1 Byte
        b1 = Chars(i)
        b2 = Chars(i + 1)
        
        Chars64(j) = mbytBase64EncodeWord(b1 \ &H4)
        Chars64(j + 1) = mbytBase64EncodeWord((b1 And &H3) * &H10 + b2 \ &H10)
        Chars64(j + 2) = mbytBase64EncodeWord((b2 And &HF) * &H4)
        Chars64(j + 3) = Base64EmptyWord
    End Select
    
    '返回转换成Base64码的字符串
    PokeLng DataPtr64, SavePtr64
    PokeLng CountPtr64, 1
    PokeLng DataPtr, SavePtr
    PokeLng CountPtr, 1
    
    Exit Function
    
errHand:
    MsgBox Err.Description
End Function

Private Function WaitResponse(ResponseCode As String, Optional ByVal Len2 As Long = 1) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能：检查是否收到SMTP服务器的返回代码
    '------------------------------------------------------------------------------------------------------------------
  
    Dim Start As Single
    Dim tmr As Single
    
    mstrSvrResponse = ""
    
    Start = Timer
    While Len(Response) = 0
        tmr = Timer - Start
        DoEvents
        If tmr > mlngResponseInternal Then
            MsgBox "SMTP服务错误，等待服务应答超时！", vbInformation, "邮件服务"
            Exit Function
        End If
        Wend
      
    While Left(Response, 3) <> ResponseCode
        tmr = Timer - Start
        DoEvents
        If tmr > Len2 Then
          
            MsgBox "SMTP服务错误，" & Response, vbInformation, "邮件服务"
            Exit Function
        End If
    Wend
    
    mstrSvrResponse = Response
    
    Response = "" ' Response清空

    WaitResponse = True
  
End Function

Private Function AuthLogin(ByVal UserName As String, ByVal UserPwd As String, ByVal Auths As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能：
    '------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHand
    
    '检查winsock的状态是否为关
    If mobjSck.State = sckClosed Then Exit Function
        
    '进行校验LOGIN
    Call mobjSck.SendData("AUTH LOGIN" + vbCrLf)
    
    If WaitResponse("334") = False Then Exit Function
    
    '输入用户名
    Call mobjSck.SendData(Base64EncodeString(UserName) + vbCrLf)
    If WaitResponse("334") = False Then Exit Function
    
    '输入用户口令
    Call mobjSck.SendData(Base64EncodeString(UserPwd) + vbCrLf)
        
    If WaitResponse("235") = False Then Exit Function
        
    AuthLogin = True
    
    Exit Function
    
errHand:
    
End Function

Private Function CheckAuthType() As String
    '------------------------------------------------------------------------------------------------------------------
    '功能：
    '------------------------------------------------------------------------------------------------------------------
    
    'HELO
'    Call mobjSck.SendData("HELO localhost" + vbCrLf)
'    Call mobjSck.SendData("EHLO 211.158.7.155" + vbCrLf)
    
    Call mobjSck.SendData("EHLO localhost" + vbCrLf)
    
    Call WaitResponse("250")
    
    If InStr(UCase(mstrSvrResponse), "250-AUTH") > 0 Then CheckAuthType = "AUTH"
    
End Function

Public Function SendHead(ByVal ToEmailAddress As String, ByVal FromName As String, ByVal FromEmailAddress As String, ByVal EmailSubject As String, Optional ByVal Mail As MailFormat = vbMultipartAlternative) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能：
    '------------------------------------------------------------------------------------------------------------------
    Dim strFirst As String
    Dim strSecond As String
    Dim strThird As String
    Dim strboundary As String
    Dim strContentType As String
    
    On Error GoTo errHand
    
    Select Case Mail
    Case MailFormat.vbMultipartMixed
        strContentType = "multipart/mixed"
    Case MailFormat.vbMultipartAlternative
        strContentType = "multipart/alternative"
    End Select
    
    strboundary = "----_=_NextPart_001_01C534D2.584225D6"
    
    strFirst = "mail from: " + FromEmailAddress + vbCrLf

    '收件人地址
    strSecond = "rcpt to:" + ToEmailAddress + vbCrLf
        
    strThird = _
            "MIME-Version: 1.0" & vbCrLf & _
            "To:" & ToEmailAddress & vbCrLf & _
            "From:" & FromName & vbCrLf & "Date:" & Format(Date, "Ddd") & ", " & Format(Date, "dd Mmm YYYY") & " " & Format(Time, "hh:mm:ss") & " +0800" & vbCrLf & _
            "Subject:" & EmailSubject & vbCrLf & _
            "Reply-To: " & FromEmailAddress & vbCrLf & _
            "Content-Type: " & strContentType & "; boundary=" & Chr(34) & strboundary & Chr(34) & vbCrLf
            
    '检查winsock的状态是否为关
    If mobjSck.State = sckClosed Then Exit Function
            
    Call mobjSck.SendData(strFirst)
        
    If WaitResponse("250") = False Then Exit Function
    
    Call mobjSck.SendData(strSecond)
    If WaitResponse("250") = False Then Exit Function
    
    Call mobjSck.SendData("data" + vbCrLf)
    If WaitResponse("354") = False Then Exit Function
    Call mobjSck.SendData(strThird + vbCrLf)

    SendHead = True
    
    Exit Function
    
errHand:
    
End Function

Public Function SendMessage(ByVal EmailMessage As String, Optional ByVal Text As TextFormat = vbTextPlain) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能：发送邮件
    '
    '
    '------------------------------------------------------------------------------------------------------------------
    Dim strboundary As String
    Dim strTextFormat As String
    
    On Error GoTo errHand
    
    Select Case Text
    Case TextFormat.vbTextPlain
        strTextFormat = "text/plain"
    Case TextFormat.vbTextHtml
        strTextFormat = "text/html"
    End Select

    strboundary = "----_=_NextPart_001_01C534D2.584225D6"
    
    mobjSck.SendData "--" & strboundary & vbCrLf & "content-type: " & strTextFormat & "; Charset=GB2312" & vbCrLf & vbCrLf
    mobjSck.SendData EmailMessage
    mobjSck.SendData vbCrLf & vbCrLf & "--" & strboundary & vbCrLf
    
    SendMessage = True
    
    Exit Function
    
errHand:
    
End Function

Public Function SendAttach(ByVal strAttachFile As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能：发送邮件附件
    '
    '
    '------------------------------------------------------------------------------------------------------------------
    Dim intFile As Integer
    Dim bytIn(3) As Byte
    Dim bytOut(4) As Byte
    Dim bytRead As Byte
    Dim intCount As Integer
    Dim intLineLen As Integer
    Dim strTmp As String
    Dim strboundary As String
    Dim strAttachName As String
    
    On Error GoTo errHand
    
    strboundary = "----_=_NextPart_001_01C534D2.584225D6"
    
    strAttachName = strAttachFile
    
    If Trim(strAttachName) <> "" Then

        For intCount = Len(strAttachName) To 1 Step -1

            If Mid(strAttachName, intCount, 1) = "\" Then
                strAttachName = Mid(strAttachName, intCount + 1)
                Exit For
            End If

        Next intCount

        mobjSck.SendData "--" & strboundary & vbCrLf
        mobjSck.SendData "Content-Type:application/octet-stream;" & vbCrLf & "name=" & Chr(34) & strAttachName & Chr(34) & vbCrLf
        mobjSck.SendData "Content-Disposition:attachment; filename=" & Chr(34) & strAttachName & Chr(34) & vbCrLf
        mobjSck.SendData "Content-Transfer-Encoding: base64" & vbCrLf & vbCrLf
    
        intFile = FreeFile()
        Open strAttachFile For Binary As #intFile
        
        While Not EOF(intFile)
            intCount = 0
            Do While intCount < 3
                Get #intFile, , bytRead
                If Not EOF(intFile) Then
                    bytIn(intCount) = bytRead
                    intCount = intCount + 1
                Else
                    Exit Do
                End If
            Loop
            
            Call Base64EncodeByte(bytIn, bytOut, intCount)
            
            For intCount = 0 To 3
                strTmp = strTmp & Chr(bytOut(intCount))
            Next
            
            intLineLen = intLineLen + 1
            
            If intLineLen * 4 > 70 Then
                intLineLen = 0
                strTmp = strTmp & vbCrLf
                
                '发送
                Call mobjSck.SendData(strTmp)

                strTmp = ""
            End If
            
        Wend
        
        If strTmp <> "" Then Call mobjSck.SendData(strTmp)
                
        Call mobjSck.SendData("==" & vbCrLf & vbCrLf)
        Call mobjSck.SendData("--" & strboundary & "--" & vbCrLf)
        
        Close intFile
           
    End If
    
    SendAttach = True
    
    Exit Function
    
errHand:
    
End Function

Public Function SendOver() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能：
    '------------------------------------------------------------------------------------------------------------------
    Dim strboundary As String
    
    On Error GoTo errHand
    
    '检查winsock的状态是否为关
    If mobjSck.State = sckClosed Then Exit Function
            
    strboundary = "----_=_NextPart_001_01C534D2.584225D6--"
    
    'mobjSck.SendData vbCrLf & vbCrLf & "--" & strboundary & vbCrLf
    Call mobjSck.SendData("." + vbCrLf)
    If WaitResponse("250", 50) = False Then Exit Function
    
    SendOver = True
    
    Exit Function
    
errHand:
    
End Function

Public Function OpenMailServer(ByVal MailServerName As String, ByVal EmialUsername As String, ByVal EmialPassword As String, Optional ByVal MailPort As Long = 25) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能：
    '------------------------------------------------------------------------------------------------------------------
    Dim strAuths As String
    
    On Error GoTo errHand
    
    '检查winsock的状态是否为关
    If mobjSck.State <> sckClosed Then mobjSck.Close
        
    mobjSck.LocalPort = 0               '用端口0来动态的建立连接
    
    mobjSck.Protocol = sckTCPProtocol   ' 设置协议为TCP
    mobjSck.RemoteHost = MailServerName     ' SMTP地址
    mobjSck.RemotePort = MailPort           ' SMTP端口
        
    mobjSck.Connect                     ' 开始连接
    
    If WaitResponse("220") = False Then Exit Function
                
    strAuths = CheckAuthType
    
    If Trim(strAuths) <> "" Then
        If AuthLogin(EmialUsername, EmialPassword, strAuths) = False Then Exit Function
    End If
    
    OpenMailServer = True
    
    Exit Function
    
errHand:
    
End Function

Public Function CloseMailServer() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能：
    '------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHand
    
    '检查winsock的状态是否为关
    If mobjSck.State = sckClosed Then Exit Function
            
    Call mobjSck.SendData("quit" + vbCrLf)
        
    If WaitResponse("221") = False Then Exit Function
    
    mobjSck.Close
    
    CloseMailServer = True
    
    Exit Function
    
errHand:
    
End Function


Private Sub Class_Initialize()
    mlngResponseInternal = 30
End Sub

Private Sub Class_Terminate()
    Set mobjSck = Nothing
End Sub

Private Sub mobjSck_DataArrival(ByVal bytesTotal As Long)
    '接收SMTP服务器的信息
    Dim strData As String
    
    On Error Resume Next
    
    mobjSck.GetData strData
    
    Response = strData
End Sub

'======================================================================================================================
'======================================================================================================================
'======================================================================================================================
'以下是利用现有的OutLook发送邮件，前提是工作站必须安装OutLook2000以上的版本
Public Function OpenOutLookMail(ByVal strProFile As String, ByVal strPassWord As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能：
    '------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHand
        
    Set mobjApp = Nothing
    Set mobjApp = CreateObject("Outlook.Application")
                
    If mobjApp Is Nothing Then Exit Function
    
    Set mobjNameSpace = Nothing
    Set mobjNameSpace = mobjApp.GetNamespace("MAPI")
    
    If mobjNameSpace Is Nothing Then Exit Function
    
    mobjNameSpace.Logon strProFile, strPassWord, False, True
    
    OpenOutLookMail = True
    
    Exit Function
errHand:
    
End Function

Public Function CloseOutLookMail(Optional ByVal blnQuit As Boolean = False) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能：
    '------------------------------------------------------------------------------------------------------------------
    On Error Resume Next
    
    mobjNameSpace.Logoff
    
    Set mobjNameSpace = Nothing
    Set mobjApp = Nothing
    
    If blnQuit Then mobjApp.Quit
    
    CloseOutLookMail = True
    
End Function

Public Function SendOutLookMail(ByVal strToAddr As String, ByVal strSubject As String, ByVal strBodyMsg As String, Optional ByVal strAttach As String = "") As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能：
    '------------------------------------------------------------------------------------------------------------------
    Dim objOutlookAttach As Object
    Dim objMail As Object
    
    On Error GoTo errHand
    
    Set objMail = mobjApp.CreateItem(0)
                
    With objMail
    
        .To = strToAddr
        .Subject = strSubject
        
        .BodyFormat = 2
        .HTMLBody = strBodyMsg
        
        If Len(Dir(strAttach)) > 0 Then
            Set objOutlookAttach = .Attachments.Add(strAttach)
        End If
        
        .Send
        
    End With
    
    Set objMail = Nothing
    Set objOutlookAttach = Nothing
    
    SendOutLookMail = True
    
    Exit Function
    
errHand:
    Set objMail = Nothing
    Set objOutlookAttach = Nothing
End Function

'======================================================================================================================
'======================================================================================================================
'======================================================================================================================
'以下是利用现有的OutLook Express发送邮件，OutLook Express是IE自带的
Public Function OpenOutLookExMail(Optional ByVal strUserName As String, Optional ByVal strPassWord As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能：
    '------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHand
    
    Set mobjMAPISession = Nothing
        
    Set mobjMAPISession = CreateObject("MSMAPI.MAPISession")
        
    With mobjMAPISession
    
        .DownLoadMail = False
        .LogonUI = False
        .NewSession = True
        .UserName = strUserName
        .Password = strPassWord
        
        .SignOn
        
    End With
    
    OpenOutLookExMail = True
    
    Exit Function
errHand:
    
End Function

Public Function CloseOutLookExMail() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能：
    '------------------------------------------------------------------------------------------------------------------
    On Error Resume Next
    
    mobjMAPISession.SignOff
    
    Set mobjMAPISession = Nothing
    
    CloseOutLookExMail = True
    
End Function

Public Function SendOutLookExMail(ByVal strToAddr As String, ByVal strSubject As String, ByVal strBodyMsg As String, Optional ByVal strAttach As String = "") As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能：
    '------------------------------------------------------------------------------------------------------------------
    Dim objMAPIMessages As Object
    
    On Error GoTo errHand
    
    Set objMAPIMessages = Nothing
    Set objMAPIMessages = CreateObject("MSMAPI.MAPIMessages")
        
    With objMAPIMessages
        
        .SessionID = mobjMAPISession.SessionID
        
        .Compose
        
        .RecipDisplayName = strToAddr
        
        .MsgSubject = strSubject
        .MsgNoteText = strBodyMsg
        
        .AttachmentType = 0
        .AttachmentPathName = strAttach

        .Send
        
    End With
    
    SendOutLookExMail = True
    
    Exit Function
    
errHand:

End Function

Public Function GetOracleMail(ByVal strToAddr As String, ByVal strSubject As String, ByVal strSenderAddr As String, ByVal strSenderSmtp As String, ByVal strUserName As String, ByVal strUserPsw As String, _
            ByVal strBodyMsg As String, Optional ByVal strAttachFile As String = "", Optional ByVal lngPort As Long = 25) As String
    
    Dim strAttachName As String
    Dim strSQL As String
    Dim intCount As Integer
    Dim strBase64File As String
    Dim intFile As Integer
    Dim bytIn(3) As Byte
    Dim bytOut(4) As Byte
    Dim bytRead As Byte
    
    Dim intLineLen As Integer
    Dim strTmp As String
    
    On Error GoTo errHand
    
    strSQL = "zl_SendMail('" & strToAddr & ",'" & strSubject & "','" & strSenderAddr & "','" & strSenderSmtp & "',"
    
    strSQL = strSQL & "'" & Base64EncodeString(strUserName) & "','" & Base64EncodeString(strUserPsw) & "',"
    
    'strBodyMsg = ReplaceAll(strBodyMsg, "'", "''")
    
    strSQL = strSQL & "'" & strBodyMsg & "',"
    
    strAttachName = strAttachFile
    
    If Trim(strAttachName) <> "" Then

        For intCount = Len(strAttachName) To 1 Step -1

            If Mid(strAttachName, intCount, 1) = "\" Then
                strAttachName = Mid(strAttachName, intCount + 1)
                Exit For
            End If

        Next intCount
        
        strSQL = strSQL & "'" & strAttachName & "',"
            
        intFile = FreeFile()
        Open strAttachFile For Binary As #intFile
        
        While Not EOF(intFile)
            intCount = 0
            Do While intCount < 3
                Get #intFile, , bytRead
                If Not EOF(intFile) Then
                    bytIn(intCount) = bytRead
                    intCount = intCount + 1
                Else
                    Exit Do
                End If
            Loop
            
            Call Base64EncodeByte(bytIn, bytOut, intCount)
            
            For intCount = 0 To 3
                strTmp = strTmp & Chr(bytOut(intCount))
            Next
            
            intLineLen = intLineLen + 1
            
            If intLineLen * 4 > 70 Then
                intLineLen = 0
                strTmp = strTmp & vbCrLf
                
                strBase64File = strBase64File & strTmp
                
                strTmp = ""
            End If
            
        Wend
        
        If strTmp <> "" Then strBase64File = strBase64File & strTmp
                        
        Close intFile
               
        strSQL = strSQL & "'" & strBase64File & "',"
    Else
        strSQL = strSQL & "NULL,NULL,"
    End If
    
    strSQL = strSQL & lngPort & ",1)"
    
    GetOracleMail = strSQL
    
    Exit Function
    
errHand:
    
End Function

