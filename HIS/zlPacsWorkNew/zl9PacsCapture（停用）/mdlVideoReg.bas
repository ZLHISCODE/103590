Attribute VB_Name = "mdlVideoReg"
Option Explicit


Public Const LOGIN_TYPE_视频设备 As String = "影像视频设备数量"


Public gint视频设备数量 As Integer

Public Function funVideoRegTime(frmParent As Form) As String
'功能：检索注册信息，有则返回注册时间
'参数： frmParent ---父窗体
'       str类型 ---'在注册码中使用的类型名称
'返回值：当前日期；无注册信息返回空
On Error GoTo err
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim strIP地址 As String         '注册的IP地址
    
    funVideoRegTime = ""
    
    If gint视频设备数量 <= -1 Then
        funVideoRegTime = Now
        Exit Function
    End If
    
    strIP地址 = funGetOneIP(frmParent)
    
    strSQL = "select 工作站 from zltools.zlclients where ip=[1] and 启用视频源=1"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "提取注册信息", strIP地址)
    
    If Not rsTemp.EOF Then funVideoRegTime = Now
    Exit Function
err:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
    err.Clear
End Function

Public Function FunLogIn(frmParent As Form, str类型 As String) As String
'功能：对程序进行注册，如果注册成功，则返回注册时间
'参数： frmParent ---父窗体
'       str类型 ---'在注册码中使用的类型名称
'返回值：注册成功注册日期；注册失败返回空

    Dim intNum As Integer
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim strIP地址 As String         '需要注册的IP地址
    
    On Error GoTo err
    
    strIP地址 = funGetOneIP(frmParent)
    
    '从注册码中提取授权的数量，-1--无限制；0--禁止；X（X>0）--按照数量控制
    intNum = gint视频设备数量
    
    'intNUM >0 ,则调用过程注册程序
    If intNum > 0 Then  '按数量限制
        strSQL = "Zl_影像操作记录_Update('" & strIP地址 & "','" & str类型 & "'," & intNum & ")"
        zlDatabase.ExecuteProcedure strSQL, "注册" & str类型
        '检查注册是否成功
        strSQL = "Select 启动时间,IP地址 from 影像操作记录 where  类型=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "提取启动时间", str类型)
        
        If rsTemp.RecordCount <= intNum Then
            rsTemp.Filter = "IP地址='" & strIP地址 & "'"
            If rsTemp.RecordCount = 1 Then  '注册成功
                FunLogIn = rsTemp!启动时间
                Exit Function
            End If
        End If
    ElseIf intNum = -1 Then     '无限制
        FunLogIn = Now
        Exit Function
    Else    '=0，或者其他值，禁止，不做任何处理，后面有提示
    
    End If
    
    '注册失败，可能是两个原因：
    '1、注册的数量超过了许可的数量，无法注册IP地址
    '2、直接通过SQL向表中添加了IP地址，导致表中的记录总数量超过了许可的数量
    Call MsgboxCus("打开的" & str类型 & "超过您购买的总数量（" & intNum & "）。请向软件供应商联系。", vbOKOnly, G_STR_HINT_TITLE)
    FunLogIn = ""
    
    Exit Function
err:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
    err.Clear
End Function

Public Function FunCheckRegInfo(frmParent As Form) As Boolean
'功能：检查是否存在注册的ip地址且启用了视频源
On Error GoTo errHandle
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim strIP地址 As String         '需要注册的IP地址
    
    FunCheckRegInfo = False
    
    strIP地址 = funGetOneIP(frmParent)
    
    strSQL = "select 工作站 from zltools.zlclients where ip=[1] and 启用视频源=1"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "提取注册信息", strIP地址)
    
    If rsTemp.EOF = False Then FunCheckRegInfo = True
    
Exit Function
errHandle:
End Function

Public Function FunCheckIp(frmParent As Form, str类型 As String) As Boolean
'功能：检查是否存在注册的ip地址
On Error GoTo errHandle
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim strIP地址 As String         '需要注册的IP地址
    
    FunCheckIp = False
    
    strIP地址 = funGetOneIP(frmParent)
    
    strSQL = "Select 启动时间 from 影像操作记录 where 类型=[2] and IP地址=[1] "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "提取启动时间", strIP地址, str类型)

    
    If rsTemp.EOF = False Then FunCheckIp = True
    
Exit Function
errHandle:
End Function

Public Function FunLogOut(frmParent As Form, str类型 As String, str启动时间 As String) As Boolean
'功能：退出程序的时候，检查程序是否合法注册过，避免有人通过触发器等手段定时删除“影像操作记录”表中的记录。
'参数： frmParent ---父窗体
'       str类型 ---'在注册码中使用的类型名称
'       str启动时间 --- 注册工作站时返回的时间
'返回值：合法注册True；非法启动的False
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim strIP地址 As String         '需要注册的IP地址
    Dim intNum As Integer
    
    On Error GoTo err
    strIP地址 = funGetOneIP(frmParent)
    
    '启动时间为空，表示注册失败，没有正常启动，因此退出的时候不再检测数据库
    If str启动时间 = "" Then
        FunLogOut = True
        Exit Function
    End If
    
    '从注册码中提取授权的数量，-1--无限制；0--禁止；X（X>0）--按照数量控制
    intNum = gint视频设备数量
    
    If intNum > 0 Then '按照数量控制
        strSQL = "Select 启动时间 from 影像操作记录 where IP地址=[1] and 类型=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "提取启动时间", strIP地址, str类型)
        If rsTemp.EOF = False Then
            FunLogOut = True
        Else
            '对比启动时间和数据库的时间，如果不是同一天，说明是前一天开启程序后注册信息被删除了，
            '这种情况认为是合法注册
            strSQL = "Select sysdate from dual"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "提取数据库时间")
            If Format(rsTemp!sysdate, "yyyy-mm-dd") <> Format(str启动时间, "yyyy-mm-dd") Then
                FunLogOut = True
            Else
                FunLogOut = False
            End If
        End If
    ElseIf intNum = -1 Then     '无限制
        FunLogOut = True
    Else    '=0，或者其他值，禁止
    
    End If
    If FunLogOut = False Then
        Call MsgboxCus("打开的" & str类型 & "超过您购买的总数量（" & intNum & "）。请向软件供应商联系。", vbOKOnly, G_STR_HINT_TITLE)
    End If
    Exit Function
err:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Function




Public Function getLicenseCount(strLicenseName As String) As Integer
'读取授权的数量
'参数： strLicenseName --- 授权名称
    Dim strLiceseCount As String
    
    On Error GoTo err
    
    strLiceseCount = zlRegInfo(strLicenseName)
    If strLiceseCount = "" Then '无限制
        getLicenseCount = -1
    ElseIf Val(strLiceseCount) > 0 Then '按照数量限制
        getLicenseCount = Val(strLiceseCount)
    Else '禁止
        getLicenseCount = 0
    End If
    
    Exit Function
err:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Function



Private Function funGetOneIP(frmParent As Form) As String
'------------------------------------------------
'功能：读取当前计算机的首个IP地址
'参数： frmParent  -- 父窗体
'返回：返回读取当前计算机的首个IP地址
'------------------------------------------------
    Dim strIP地址 As String
    
    On Error Resume Next
    
    strIP地址 = funcGetLocalIP(frmParent)
    If strIP地址 = "" Then
        funGetOneIP = "127.0.0.1"
    ElseIf InStr(strIP地址, ",") <> 0 Then
        funGetOneIP = Split(strIP地址, ",")(0)
    Else
        funGetOneIP = strIP地址
    End If
End Function




Private Function funcGetLocalIP(frmParent As Form) As String
'------------------------------------------------
'功能：提取当前计算机的IP地址串，用逗号分隔
'参数： frmParent  -- 父窗体
'返回：返回当前计算机的IP地址串，用逗号分隔
'------------------------------------------------
    Dim hostname As String * 256
    Dim hostent_addr As Long
    Dim host As HOSTENT
    Dim hostip_addr As Long
    Dim temp_ip_address() As Byte
    Dim i As Integer
    Dim ip_address As String
    Dim strLocalIPs As String

    '启动Socket
    Call SocketsInitialize(frmParent)

    If gethostname(hostname, 256) = SOCKET_ERROR Then
        MsgboxCus "Windows Sockets error " & Str(WSAGetLastError()), vbOKOnly, G_STR_HINT_TITLE
        Exit Function
    Else
        hostname = Trim$(hostname)
    End If

    hostent_addr = gethostbyname(hostname)

    If hostent_addr = 0 Then
        MsgboxCus "Winsock.dll is not responding.", vbOKOnly, G_STR_HINT_TITLE
        Exit Function
    End If

    RtlMoveMemory host, hostent_addr, LenB(host)
    RtlMoveMemory hostip_addr, host.hAddrList, 4

    ''''''''''''''''get all of the IP address if machine is  multi-homed

    Do
        ReDim temp_ip_address(1 To host.hLength)
        RtlMoveMemory temp_ip_address(1), hostip_addr, host.hLength

        For i = 1 To host.hLength
            ip_address = ip_address & temp_ip_address(i) & "."
        Next
        ip_address = Mid$(ip_address, 1, Len(ip_address) - 1)

        strLocalIPs = IIf(strLocalIPs = "", ip_address, strLocalIPs & "," & ip_address)

        ip_address = ""
        host.hAddrList = host.hAddrList + LenB(host.hAddrList)
        RtlMoveMemory hostip_addr, host.hAddrList, 4
     Loop While (hostip_addr <> 0)

    '清除Socket
    Call SocketsCleanup(frmParent)

    funcGetLocalIP = strLocalIPs
End Function




Private Sub SocketsInitialize(frmParent As Form)
'------------------------------------------------
'功能：初始化Socket
'参数： frmParent  -- 父窗体
'返回：无
'------------------------------------------------
    Dim WSAD As WSADATA
    Dim iReturn As Integer
    Dim sLowByte As String, sHighByte As String, sMsg As String

    iReturn = WSAStartup(WS_VERSION_REQD, WSAD)

    If iReturn <> 0 Then
        MsgboxCus "Winsock.dll is not responding.", vbOKOnly, G_STR_HINT_TITLE
        Exit Sub
    End If

    If lobyte(WSAD.wversion) < WS_VERSION_MAJOR Or (lobyte(WSAD.wversion) = _
        WS_VERSION_MAJOR And hibyte(WSAD.wversion) < WS_VERSION_MINOR) Then

        sHighByte = Trim$(Str$(hibyte(WSAD.wversion)))
        sLowByte = Trim$(Str$(lobyte(WSAD.wversion)))
        sMsg = "Windows Sockets version " & sLowByte & "." & sHighByte
        sMsg = sMsg & " is not supported by winsock.dll "
        MsgboxCus sMsg, vbOKOnly, G_STR_HINT_TITLE
        Exit Sub
    End If

    ''''''''''''''''iMaxSockets is not used in winsock 2. So the following check is only
    ''''''''''''''''necessary for winsock 1. If winsock 2 is requested,
    ''''''''''''''''the following check can be skipped.

    If WSAD.iMaxSockets < MIN_SOCKETS_REQD Then
        sMsg = "This application requires a minimum of "
        sMsg = sMsg & Trim$(Str$(MIN_SOCKETS_REQD)) & " supported sockets."
        MsgboxCus sMsg, vbOKOnly, G_STR_HINT_TITLE
        Exit Sub
    End If
End Sub



'
Private Sub SocketsCleanup(frmParent As Form)
'------------------------------------------------
'功能：清除Socket
'参数： frmParent  -- 父窗体
'返回：无
'------------------------------------------------
Dim lReturn As Long

    lReturn = WSACleanup()

    If lReturn <> 0 Then
        MsgboxCus "Socket error " & Trim$(Str$(lReturn)) & " occurred in Cleanup ", vbOKOnly, G_STR_HINT_TITLE
        Exit Sub
    End If
End Sub



'
Private Function hibyte(ByVal wParam As Integer)
    hibyte = wParam \ &H100 And &HFF&
End Function




Private Function lobyte(ByVal wParam As Integer)
    lobyte = wParam And &HFF&
End Function

