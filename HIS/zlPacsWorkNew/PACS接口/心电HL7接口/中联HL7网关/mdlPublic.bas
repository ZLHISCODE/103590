Attribute VB_Name = "mdlPublic"
Option Explicit

Public Declare Function InitCommonControls Lib "comctl32.dll" () As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

'---------------------------------------------------------------
'-注册表 API 声明...
'---------------------------------------------------------------
Public Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Public Declare Function RegCreateKeyEx Lib "advapi32" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByRef lpSecurityAttributes As SECURITY_ATTRIBUTES, ByRef phkResult As Long, ByRef lpdwDisposition As Long) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Public Declare Function SetActiveWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Const GWL_EXSTYLE = (-20)
Public Const WinStyle = &H40000
Public Const SWP_NOSIZE = &H1
Public Const SWP_SHOWWINDOW = &H40
Public Const HWND_TOPMOST = -1

'---------------------------------------------------------------
'- 注册表 Api 常数...
'---------------------------------------------------------------
' Reg Data Types...
Public Const REG_SZ = 1                         ' Unicode空终结字符串
Public Const REG_EXPAND_SZ = 2                  ' Unicode空终结字符串
Public Const REG_DWORD = 4                      ' 32-bit 数字

' 注册表创建类型值...
Public Const REG_OPTION_NON_VOLATILE = 0       ' 当系统重新启动时，关键字被保留

' 注册表关键字安全选项...
Public Const READ_CONTROL = &H20000
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_SET_VALUE = &H2
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const KEY_CREATE_LINK = &H20
Public Const KEY_READ = KEY_QUERY_VALUE + KEY_ENUMERATE_SUB_KEYS + KEY_NOTIFY + READ_CONTROL
Public Const KEY_WRITE = KEY_SET_VALUE + KEY_CREATE_SUB_KEY + READ_CONTROL
Public Const KEY_EXECUTE = KEY_READ
Public Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' 注册表关键字根类型...
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004

' 返回值...
Public Const ERROR_NONE = 0
Public Const ERROR_BADKEY = 2
Public Const ERROR_ACCESS_DENIED = 8
Public Const ERROR_SUCCESS = 0

'---------------------------------------------------------------
'- 注册表安全属性类型...
'---------------------------------------------------------------
Public Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Boolean
End Type

'---------------------------------------------------------------
'读取网卡的多个IP
'---------------------------------------------------------------
Private Const WS_VERSION_REQD = &H101
Private Const WS_VERSION_MAJOR = WS_VERSION_REQD \ &H100 And &HFF&
Private Const WS_VERSION_MINOR = WS_VERSION_REQD And &HFF&
Private Const MIN_SOCKETS_REQD = 1
Private Const SOCKET_ERROR = -1
Private Const WSADescription_Len = 256
Private Const WSASYS_Status_Len = 128

Private Type HOSTENT
    hName As Long
    hAliases As Long
    hAddrType As Integer
    hLength As Integer
    hAddrList As Long
End Type

Private Type WSADATA
    wversion As Integer
    wHighVersion As Integer
    szDescription(0 To WSADescription_Len) As Byte
    szSystemStatus(0 To WSASYS_Status_Len) As Byte
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpszVendorInfo As Long
End Type

Private Declare Function WSAGetLastError Lib "WSOCK32.DLL" () As Long
Private Declare Function WSAStartup Lib "WSOCK32.DLL" (ByVal wVersionRequired As Integer, lpWSAData As WSADATA) As Long
Private Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long

Private Declare Function gethostname Lib "WSOCK32.DLL" (ByVal hostname$, ByVal HostLen As Long) As Long
Private Declare Function gethostbyname Lib "WSOCK32.DLL" (ByVal hostname$) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (hpvDest As Any, ByVal hpvSource&, ByVal cbCopy&)

'延时计时
Public Declare Function timeGetTime Lib "winmm.dll" () As Long


'------------------------HL7网关使用的自定义类型-----------------------------
Public Type THL7Service
    lngID As Long                   '服务ID
    strIP As String                 '服务的IP地址
    strSendApp As String            '服务的发送程序名称
    strSendFacility As String       '服务的发送设备名称
    strReceiveApp As String         '服务的接收程序名称
    strReceiveFacility As String    '服务的接收设备名称
    lngPort As Long               '服务的端口号
    intServiceType As Integer       '服务的类别；1-接收；2-发送
    Started  As Boolean             '当前服务是否成功启动
End Type
Public HL7Services() As THL7Service    '存储应用于当前IP地址的HL7服务对


Public Function Nvl(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'功能：相当于Oracle的NVL，将Null值改成另外一个预设值
    Nvl = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

Public Function funcGetLocalIP() As String
'返回当前计算机的IP地址串，用逗号分隔
    Dim hostname As String * 256
    Dim hostent_addr As Long
    Dim host As HOSTENT
    Dim hostip_addr As Long
    Dim temp_ip_address() As Byte
    Dim i As Integer
    Dim ip_address As String
    Dim strLocalIPs As String

    '启动Socket
    Call SocketsInitialize

    If gethostname(hostname, 256) = SOCKET_ERROR Then
        MsgBox "Windows Sockets error " & Str(WSAGetLastError())
        Exit Function
    Else
        hostname = Trim$(hostname)
    End If

    hostent_addr = gethostbyname(hostname)

    If hostent_addr = 0 Then
        MsgBox "Winsock.dll is not responding."
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
    Call SocketsCleanup
    
    funcGetLocalIP = strLocalIPs
End Function

Private Sub SocketsInitialize()
    Dim WSAD As WSADATA
    Dim iReturn As Integer
    Dim sLowByte As String, sHighByte As String, sMsg As String

    iReturn = WSAStartup(WS_VERSION_REQD, WSAD)

    If iReturn <> 0 Then
        MsgBox "Winsock.dll is not responding."
        End
    End If

    If lobyte(WSAD.wversion) < WS_VERSION_MAJOR Or (lobyte(WSAD.wversion) = _
        WS_VERSION_MAJOR And hibyte(WSAD.wversion) < WS_VERSION_MINOR) Then

        sHighByte = Trim$(Str$(hibyte(WSAD.wversion)))
        sLowByte = Trim$(Str$(lobyte(WSAD.wversion)))
        sMsg = "Windows Sockets version " & sLowByte & "." & sHighByte
        sMsg = sMsg & " is not supported by winsock.dll "
        MsgBox sMsg
        End
    End If

    ''''''''''''''''iMaxSockets is not used in winsock 2. So the following check is only
    ''''''''''''''''necessary for winsock 1. If winsock 2 is requested,
    ''''''''''''''''the following check can be skipped.

    If WSAD.iMaxSockets < MIN_SOCKETS_REQD Then
        sMsg = "This application requires a minimum of "
        sMsg = sMsg & Trim$(Str$(MIN_SOCKETS_REQD)) & " supported sockets."
        MsgBox sMsg
        End
    End If
End Sub

Private Sub SocketsCleanup()
Dim lReturn As Long

    lReturn = WSACleanup()

    If lReturn <> 0 Then
        MsgBox "Socket error " & Trim$(Str$(lReturn)) & " occurred in Cleanup "
        End
    End If
End Sub

Private Function hibyte(ByVal wParam As Integer)
    hibyte = wParam \ &H100 And &HFF&
End Function

Private Function lobyte(ByVal wParam As Integer)
    lobyte = wParam And &HFF&
End Function


Public Function MsgInQueue(strMsg As String) As Boolean
'------------------------------------------------
'功能：消息入队，等候被处理
'参数： strMsg －－【IN】需要入队的消息内容，可能是完整消息，也可能是一个消息包
'返回：True－－入队成功，False－－入队失败
'-----------------------------------------------
    Dim Timer As Long
    Dim intCount As Integer
    
    MsgInQueue = False
    
    If gblnQueueBusy = True Then
        '如果队列忙，正在进行队列处理，则等待，如果等待超时，则将消息放到备用队列
        
        '记录处理日志
        Call WriteProcessLog("MsgInQueue", "消息未入队", "处理队列忙，消息未入队，消息前半段：" + Left(strMsg, 150), 3)
    Else
        '如果队列空闲，则进行入队处理，并且标记队列忙
        gblnQueueBusy = True
        
        On Error GoTo err
        
        '处理消息入队
        intCount = UBound(gstrMsgQueue) + 1
        ReDim Preserve gstrMsgQueue(intCount) As String
        gstrMsgQueue(intCount) = strMsg
        
        '记录处理日志
        Call WriteProcessLog("MsgInQueue", "消息入队", "接收到完整消息，该消息进入处理队列，消息前半段：" + Left(strMsg, 150), 3)
    
        '队列处理完成，标记队列为闲
        gblnQueueBusy = False
    End If
        
    MsgInQueue = True
    Exit Function
err:
    '出错就退出，暂时不做处理
    Call WriteLog(4003, err.Number, "MsgInQueue 出现错误，strMsg前半段 = " & Left(strMsg, 150) & "，错误描述是：" & err.Description)
    gblnQueueBusy = False
End Function

Public Function MsgOutQueue() As String
'------------------------------------------------
'功能：消息出队，调用过程负责处理出队的消息
'参数：
'返回：返回出队的消息内容
'-----------------------------------------------
    Dim iCount As Integer
    Dim strMsg As String
        
    MsgOutQueue = ""        '初始化为空消息
    
    If gblnQueueBusy = True Then
        '如果队列忙，正在进行队列处理，则等待，如果等待超时，退出出队操作
        
    Else
        '如果队列空闲，则进行出队处理，并且标记队列忙
        gblnQueueBusy = True
        
        On Error GoTo err
        
        '消息出队处理
        iCount = UBound(gstrMsgQueue)
        If iCount = 0 Then
            gblnQueueBusy = False
            Exit Function     '队列为空，不用出队
        End If
        
        '从队列中提取消息
        strMsg = gstrMsgQueue(gintQueueIndex)
        
        '处理队列指针
        gintQueueIndex = gintQueueIndex + 1
        If gintQueueIndex > iCount Then
            '如果当前取出来的是队列中的最后一个消息，则清空消息队列
            ReDim Preserve gstrMsgQueue(0) As String
            gintQueueIndex = 1
        End If
        
        MsgOutQueue = strMsg
        
        '出队处理完成，标记队列闲
        gblnQueueBusy = False
    End If
    
    Exit Function
err:
    Call WriteLog(4005, err.Number, "MsgOutQueue 出现错误，strMsg前半段 = " & Left(strMsg, 150) & "，错误描述是：" & err.Description)
    gblnQueueBusy = False
End Function

Public Function funGetAMessage(ByRef strMsg As String) As Boolean
'------------------------------------------------
'功能：从消息队列提取一个消息
'参数： strMsg -- 返回提取到的消息,空表示出错或者没有消息
'返回： True - 成功；False - 失败
'-----------------------------------------------
    
    funGetAMessage = False
    
    On Error GoTo err
    
    '从消息队列提取一个消息
    strMsg = MsgOutQueue
    
    '如果消息不为空，则处理这个消息
    If strMsg <> "" Then
        funGetAMessage = True
    End If
    
    Exit Function
err:
    '暂不处理
End Function


Public Function funMsgProcess() As Boolean
'------------------------------------------------
'功能：自动处理消息
'参数： 无
'返回：True -- 成功； False -- 失败
'-----------------------------------------------
    Dim strMsg As String
    
    funMsgProcess = False
    
    '如果已经进行消息处理了，则退出
    If gblnMsgProcessing = True Then Exit Function
    
    '设置消息处理标记，防止这个过程被多次调用
    gblnMsgProcessing = True
    
    On Error GoTo err
    
    '消息出队并处理消息
    While funGetAMessage(strMsg) = True
        '记录日志
        Call WriteProcessLog("funMsgProcess", "处理消息", "消息前半段 = " & Left(strMsg, 150), 2)

        '解析并处理消息
        Call funParseInMsg(strMsg)
    Wend
    
    '消息处理完成，退出程序
    gblnMsgProcessing = False
    
    funMsgProcess = True
    Exit Function
err:
    Call WriteLog(4006, err.Number, "funMsgProcess 出现错误，strMsg前半段 = " & Left(strMsg, 150) & "，错误描述是：" & err.Description)
    gblnMsgProcessing = False
End Function

Public Sub WriteProcessLog(logSubName As String, logTitle As String, logDesc As String, lngLogLevel As Long)
'------------------------------------------------
'功能：记录通讯日志
'参数： logSubName  --  产生日志的函数名
'       logTitle   --   日志名称
'       logDesc   --    日志内容
'       lngLogLevel --  日志级别，通过日志级别确定当前日志是否需要记录
'返回：无
'------------------------------------------------

    Dim strSQL As String
    
    On Error GoTo err
    
    '启动了记录日志，才记录当前的日志,判断日志级别，确定本次日志是否需要记录
    If gblnProcessLog And glngProcessLogLevel >= lngLogLevel Then
        If gcnAccess.State = adStateClosed Then Exit Sub
        
        '对日志内容中的单引号进行转义，否则保存到Access数据库会出错
        logDesc = Replace(logDesc, "'", "‘")
        
        strSQL = "Insert into HL7通讯日志 (通讯时间,通讯函数,记录标题,记录内容) " & _
            "Values( cDate('" & Date & " " & Time() & "'),'" & logSubName & "','" & logTitle & _
            "','" & logDesc & "')"
        gcnAccess.Execute strSQL
    End If
    Exit Sub
err:
    Call WriteLog(9001, err.Number, "WriteProcessLog 记录通讯日志出错" & ",logSubName=" & logSubName & "，logTitle=" & logTitle & "，错误描述是：" & err.Description)
End Sub

Public Sub WriteLog(ByVal ErrorType As Integer, ErrorNum As Long, ErrorDesc As String)
'-----------------------------------------------------------------------------
'功能:填写错误日志
'参数： ErrorType ----错误类型代码，保存图像错误100，WORKLIST和QR错误200，FTP错误300,funSplitSeriesUID错误1001,文件通讯错误4000
'       ErrorNum ----错误号
'       ErrorDesc ----错误描述
'返回值：无
'-----------------------------------------------------------------------------
    Dim strSQL As String
    On Error Resume Next
    If gcnAccess.State = adStateClosed Then Exit Sub
    
    '对日志内容中的单引号进行转义，否则保存到Access数据库会出错
    ErrorDesc = Replace(ErrorDesc, "'", "‘")
        
    strSQL = "Insert Into 错误日志(产生时间,错误类型,错误号,错误信息) " & _
        "Values(cDate('" & Date & " " & Time() & "')," & ErrorType & "," & ErrorNum & ",'" & ErrorDesc & "')"
    
    gcnAccess.Execute strSQL
End Sub

Public Sub WriteMessageLog(strMessageType As String, strMessage As String)
'------------------------------------------------
'功能：记录接收到并且处理成功的消息
'参数： strMessageType  --  消息处理类型
'       strMessage   --   消息内容
'返回：无
'------------------------------------------------

    Dim strSQL As String
    
    On Error Resume Next
    
    If gcnAccess.State = adStateClosed Then Exit Sub
    
    '对日志内容中的单引号进行转义，否则保存到Access数据库会出错
    strMessage = Replace(strMessage, "'", "‘")
    
    strSQL = "Insert into HL7消息记录 (通讯时间,消息类型,消息内容) " & _
        "Values( cDate('" & Date & " " & Time() & "'),'" & strMessageType & "','" & strMessage & "')"
    gcnAccess.Execute strSQL
    
End Sub

Public Function funMsgFullType(strMsg As String) As Long
'-----------------------------------------------------------------------------
'功能:检查消息完整的类型
'参数： strMsg ----消息原文
'返回值：0 -- 是完整消息；1 -- 是消息头；2 -- 是消息尾；3 -- 是消息中间段；4 -- 错误
'-----------------------------------------------------------------------------
    On Error GoTo err
    
    '检查消息的完整性，如果是完整消息，直接入队，如果不完整，则先放入临时队列等待后续消息
    '消息的完整性是有首尾标记的，首字符是chr(11)，消息结束符是chr(28)chr(13)
    If Left(strMsg, 1) = Chr(11) And InStr(strMsg, Chr(28) & Chr(13)) <> 0 Then
        funMsgFullType = 0  '完整消息
    ElseIf Left(strMsg, 1) = Chr(11) Then
        funMsgFullType = 1  '消息头
    ElseIf InStr(strMsg, Chr(28) & Chr(13)) <> 0 Then
        funMsgFullType = 2  '消息尾
    Else
        funMsgFullType = 3  '消息中间段
    End If
    
    Exit Function
err:
    funMsgFullType = 4  '错误
End Function

Public Function getMsgDefFromDB(strActionType As String) As THl7Messages
'-----------------------------------------------------------------------------
'功能:根据动作类型，从数据库中读取需要处理的HL7消息定义
'参数： strActionType ----动作类型
'返回值：返回组织好的HL7消息定义
'-----------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsMsg As ADODB.Recordset
    Dim rsSegment As ADODB.Recordset
    Dim arrHL7Msgs As THl7Messages
    Dim iMsg As Integer
    Dim iSeg As Integer
    Dim iField As Integer
    Dim arrLoadSegments() As String
    
    On Error GoTo err
    
    ReDim getMsgDefFromDB.arrMsgs(0)
    
    If strActionType <> HL7_MSG_SEND_NEW_ORDER And strActionType <> HL7_MSG_SEND_CANCEL_ORDER _
        And strActionType <> HL7_MSG_SEND_DEL_ORDER Then
        Exit Function
    End If
    
    '首先从“hl7消息定义”表中读取需要发送的HL7消息段组合
    strSQL = "Select a.ID,a.服务ID,a.动作类型,a.消息名称,a.消息类型,a.消息段组合,b.IP地址,b.端口号 " & _
                " From zlhis.hl7消息定义 a,zlhis.HL7服务配置 b Where a.服务ID = b.Id And 服务类型 = 2 and a.动作类型 =[1]"
    Set rsMsg = gzlDatabase.OpenSQLRecord(strSQL, "查找医嘱消息的定义", strActionType)
    
    strSQL = "Select 消息id, 消息段名称, 段内序号, 数据类型, 接收数据值, 发送数据值, 元素名称 " & _
             " From zlhis.Hl7消息段配置 a ,zlhis.hl7消息定义 b Where  b.动作类型 =[1] And a.消息id = b.Id order by 消息段名称,段内序号"
    Set rsSegment = gzlDatabase.OpenSQLRecord(strSQL, "查找医嘱消息的内容", strActionType)
    
    If rsMsg.EOF = False Then
        ReDim arrHL7Msgs.arrMsgs(rsMsg.RecordCount)
        
        rsMsg.MoveFirst
        For iMsg = 1 To rsMsg.RecordCount
            arrHL7Msgs.arrMsgs(iMsg).lngID = rsMsg!ID
            arrHL7Msgs.arrMsgs(iMsg).lngServiceID = rsMsg!服务ID
            arrHL7Msgs.arrMsgs(iMsg).strActionType = Nvl(rsMsg!动作类型)
            arrHL7Msgs.arrMsgs(iMsg).strMsgName = Nvl(rsMsg!消息名称)
            arrHL7Msgs.arrMsgs(iMsg).strMsgType = Nvl(rsMsg!消息类型)
            arrHL7Msgs.arrMsgs(iMsg).strMsgSegmentDef = Nvl(rsMsg!消息段组合)
            arrHL7Msgs.arrMsgs(iMsg).strIP = Nvl(rsMsg!IP地址)
            arrHL7Msgs.arrMsgs(iMsg).lngPort = Nvl(rsMsg!端口号, 0)
            arrHL7Msgs.arrMsgs(iMsg).blnSendOK = False
            
            arrLoadSegments = Split(arrHL7Msgs.arrMsgs(iMsg).strMsgSegmentDef, "|")
            ReDim arrHL7Msgs.arrMsgs(iMsg).arrSegments(UBound(arrLoadSegments) + 1)
            '填充每一个消息的段
            
            For iSeg = 0 To UBound(arrLoadSegments)
                rsSegment.Filter = "消息段名称 = '" & arrLoadSegments(iSeg) & "'"
                If rsSegment.EOF = False Then
                    With arrHL7Msgs.arrMsgs(iMsg).arrSegments(iSeg + 1)
                        .intNo = iSeg + 1
                        .strName = arrLoadSegments(iSeg)
                        ReDim .arrFields(rsSegment.RecordCount)
                        
                        
                        rsSegment.MoveFirst
                        For iField = 1 To rsSegment.RecordCount
                            .arrFields(iField).intNo = iField
                            .arrFields(iField).strDataType = Nvl(rsSegment!数据类型)
                            .arrFields(iField).strElementName = Nvl(rsSegment!元素名称)
                            .arrFields(iField).strRecDataDef = Nvl(rsSegment!接收数据值)
                            .arrFields(iField).strSendDataDef = Nvl(rsSegment!发送数据值)
                            rsSegment.MoveNext
                        Next iField
                    End With
                End If
            Next iSeg
            
            rsMsg.MoveNext
        Next iMsg
    End If
    
    getMsgDefFromDB = arrHL7Msgs
    Exit Function
err:

    Call WriteLog(2001, err.Number, "getMsgDefFromDB 出现错误，错误描述是：" & err.Description)
End Function

Public Function funfillMsgValue(thisMessages As THl7Messages, strWorkIDs As String) As Long
'-----------------------------------------------------------------------------
'功能:根据业务ID，填充HL7消息的内容
'参数： thisMessages -- HL7消息定义
'       strWorkIDs -- 业务ID串，使用“;”连接多个ID。发送医嘱时，是“医嘱ID;发送号”
'返回值：0-成功；1-医嘱消息，业务ID串是空，无法发送消息；2-失败
'-----------------------------------------------------------------------------
    Dim iMsg As Integer
    Dim iSeg As Integer
    Dim iField As Integer
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim strValue As String
    Dim strField As String
    Dim strFieldValue As String
    Dim strSysDate As String
    Dim strOneSegment As String
    Dim strOneMsg As String
    Dim arrWorkIDs() As String
    Dim strSpliterL As String
    Dim strSpliterR As String
    Dim lngPatientID As Long
    
    On Error GoTo err
    
    '消息为空，则退出
    If UBound(thisMessages.arrMsgs) = 0 Then Exit Function
    
    strSpliterL = "##L##"
    strSpliterR = "##R##"
    
    '从数据库中读取医嘱
    '判断是否医嘱消息，现在只处理医嘱消息，新医嘱，取消医嘱，删除医嘱
    
    For iMsg = 1 To UBound(thisMessages.arrMsgs)
        strOneMsg = ""
        
        If thisMessages.arrMsgs(iMsg).strActionType = HL7_MSG_SEND_NEW_ORDER _
            Or thisMessages.arrMsgs(iMsg).strActionType = HL7_MSG_SEND_CANCEL_ORDER Then
            
            '暂时没有删除医嘱需要发送 Or thisMessages.arrMsgs(iMsg).strActionType = HL7_MSG_SEND_DEL_ORDER Then
            
            '医嘱消息的业务ID串是：医嘱ID，发送号
            arrWorkIDs = Split(strWorkIDs, ";")
            If UBound(arrWorkIDs) <> 1 Then
                funfillMsgValue = 1     '医嘱消息，业务ID串是空
                Exit Function
            End If
            
            '提取数据库时间的时间戳，毫秒级2位
            strSQL = "Select to_char(current_timestamp,'YYYYMMDDHH24MISSFF2') as MSGControlID From dual"
            Set rsTemp = gzlDatabase.OpenSQLRecord(strSQL, "提取当前时间")
            strSysDate = rsTemp!MSGControlID
            
            '根据业务ID串(医嘱ID，发送号)读取需要发送的医嘱信息
            '病人医嘱记录.医嘱期效 ----- 0-长期;1-临时
            '病人医嘱记录.医嘱状态 --- --1-未生效的暂存医嘱；1-新开；2-校对疑问；3-已校对；4-已作废；5-已重整；6-已暂停；7-已启用；8-已停止；9-已确认停止
            '病人医嘱状态.操作类型 --- 1-新开；2-校对疑问；3-校对通过；4-作废；5-重整；6-暂停；7-启用；8-停止；9-确认停止；10-皮试结果,11-审核通过，12-审核未通过，13-实习医师停嘱后待审核
            '病人医嘱发送.执行状态 ---- 0-未执行;1-完全执行;2-拒绝执行;3-正在执行(今后可能分解为若干实际步骤)
            
            strSQL = "Select 病人ID from zlhis.病人医嘱记录 where id =[1]"
            Set rsTemp = gzlDatabase.OpenSQLRecord(strSQL, "查询医嘱基本信息", CLng(arrWorkIDs(0)), CLng(arrWorkIDs(1)))
            If rsTemp.RecordCount = 0 Then Exit Function
            
            lngPatientID = rsTemp!病人ID
                        
            
            '先判断医嘱状态是否跟当前消息状态一致，这里无法判断
            
            On Error Resume Next
            
            '填充每一段医嘱消息
            For iSeg = 1 To UBound(thisMessages.arrMsgs(iMsg).arrSegments)
                
                strOneSegment = thisMessages.arrMsgs(iMsg).arrSegments(iSeg).strName
                
                '填写消息段的每一个字段
                For iField = 1 To UBound(thisMessages.arrMsgs(iMsg).arrSegments(iSeg).arrFields)
                    With thisMessages.arrMsgs(iMsg).arrSegments(iSeg).arrFields(iField)
                        strValue = .strSendDataDef
                        
                        '解码返回字符串
                        Do While InStr(strValue, "[") <> 0
                            '返回字符串不符合规则的，直接退出循环，返回空
                            If InStr(strValue, "]") = 0 Or InStr(strValue, "]") < InStr(strValue, "[") Then
                                strValue = ""
                                Exit Do
                            End If
                        
                            strField = Mid(strValue, InStr(strValue, "[") + 1, InStr(strValue, "]") - InStr(strValue, "[") - 1)
                            
                            strFieldValue = ""
                            If strField = "当前时间" Then
                                strFieldValue = strSysDate
                            ElseIf strField = "消息类型" Then
                                strFieldValue = thisMessages.arrMsgs(iMsg).strMsgType
                            Else
                                strFieldValue = funGetFieldValueByFun(strField, lngPatientID, CLng(arrWorkIDs(0)), CLng(arrWorkIDs(1)), thisMessages.arrMsgs(iMsg).lngID)
                            End If
                            
                            '替换掉值串中可能出现的[]分隔符
                            strFieldValue = Replace(strFieldValue, "[", strSpliterL)
                            strFieldValue = Replace(strFieldValue, "]", strSpliterR)
                            
                            strValue = Replace(strValue, "[" & strField & "]", strFieldValue)
                        Loop
                        
                        '解析完之后，将原来的[]分隔符替换回去
                        strValue = Replace(strValue, strSpliterL, "[")
                        strValue = Replace(strValue, strSpliterR, "]")
                        
                        .strSendDataValue = strValue
                        strOneSegment = strOneSegment & "|" & .strSendDataValue
                    End With
                    
                Next iField
                
                thisMessages.arrMsgs(iMsg).arrSegments(iSeg).strText = strOneSegment
                strOneMsg = strOneMsg & thisMessages.arrMsgs(iMsg).arrSegments(iSeg).strText & Chr(13)
            Next iSeg
            
            thisMessages.arrMsgs(iMsg).strText = Chr(11) & strOneMsg & Chr(28) & Chr(13)
                
        End If
    Next iMsg
    
    Exit Function
err:
    Call WriteLog(2002, err.Number, "fillMsgValue 出现错误，错误描述是：" & err.Description)
End Function

Private Function funGetFieldValueByFun(strField As String, lngPatientID As Long, lngOrderID As Long, _
    lngSendNo As Long, lngMesageID As Long) As String
'-----------------------------------------------------------------------------
'功能:通过数据库的函数 HL7_Replace_Element_Value 读取对应的数据值
'参数： strField -- 字段名称
'       lngPatientID ---病人ID
'       lngOrderID --- 医嘱ID
'       lngSendNo --- 发送号
'       lngMesageID --- 消息ID
'返回值：字段对应的返回值
'-----------------------------------------------------------------------------
    On Error GoTo err
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "Select zlhis.b_hl7interface.HL7_Replace_Element_Value([1],[2],[3],[4],[5]) as 结果值 from dual "
    Set rsTemp = gzlDatabase.OpenSQLRecord(strSQL, "查询字段对应的值", strField, lngPatientID, lngOrderID, lngSendNo, lngMesageID)
    
    If rsTemp.RecordCount > 0 Then
        funGetFieldValueByFun = Nvl(rsTemp!结果值)
    End If
    
    Exit Function
err:
    Call WriteLog(2005, err.Number, "funGetFieldValueByFun 出现错误，错误描述是：" & err.Description)
End Function

Public Function funDuplicateMsg(thisMessages As THl7Messages, intTimes As Integer) As Long
'-----------------------------------------------------------------------------
'功能:根据发送次数，复制消息，针对长嘱的多次发送
'参数： thisMessages -- HL7消息
'       intTimes -- 发送次数
'返回值：0-成功；1-失败
'-----------------------------------------------------------------------------
    Dim iMsgCount As Integer
    Dim iMsg As Integer
    Dim iSegment As Integer
    Dim iField As Integer
    Dim iDuplicate As Integer
    
    
    On Error GoTo err
    
    If intTimes <= 1 Then Exit Function
    
    iMsgCount = UBound(thisMessages.arrMsgs)
    
    '消息为空，则退出
    If iMsgCount = 0 Then Exit Function
    
    ReDim Preserve thisMessages.arrMsgs(iMsgCount * intTimes)
    
    On Error Resume Next
    
    For iDuplicate = 1 To intTimes - 1
        For iMsg = 1 To iMsgCount
            thisMessages.arrMsgs(iDuplicate * iMsgCount + iMsg).blnSendOK = False
            thisMessages.arrMsgs(iDuplicate * iMsgCount + iMsg).lngID = thisMessages.arrMsgs(iMsg).lngID
            thisMessages.arrMsgs(iDuplicate * iMsgCount + iMsg).lngPort = thisMessages.arrMsgs(iMsg).lngPort
            thisMessages.arrMsgs(iDuplicate * iMsgCount + iMsg).lngServiceID = thisMessages.arrMsgs(iMsg).lngServiceID
            thisMessages.arrMsgs(iDuplicate * iMsgCount + iMsg).strActionType = thisMessages.arrMsgs(iMsg).strActionType
            thisMessages.arrMsgs(iDuplicate * iMsgCount + iMsg).strIP = thisMessages.arrMsgs(iMsg).strIP
            thisMessages.arrMsgs(iDuplicate * iMsgCount + iMsg).strMsgName = thisMessages.arrMsgs(iMsg).strMsgName
            thisMessages.arrMsgs(iDuplicate * iMsgCount + iMsg).strMsgSegmentDef = thisMessages.arrMsgs(iMsg).strMsgSegmentDef
            thisMessages.arrMsgs(iDuplicate * iMsgCount + iMsg).strMsgType = thisMessages.arrMsgs(iMsg).strMsgType
            thisMessages.arrMsgs(iDuplicate * iMsgCount + iMsg).strText = thisMessages.arrMsgs(iMsg).strText
            ReDim thisMessages.arrMsgs(iDuplicate * iMsgCount + iMsg).arrSegments(UBound(thisMessages.arrMsgs(iMsg).arrSegments))
            For iSegment = 1 To UBound(thisMessages.arrMsgs(iMsg).arrSegments)
                With thisMessages.arrMsgs(iDuplicate * iMsgCount + iMsg).arrSegments(iSegment)
                    .blnEnable = thisMessages.arrMsgs(iMsg).arrSegments(iSegment).blnEnable
                    .intNo = thisMessages.arrMsgs(iMsg).arrSegments(iSegment).intNo
                    .strName = thisMessages.arrMsgs(iMsg).arrSegments(iSegment).strName
                    .strText = thisMessages.arrMsgs(iMsg).arrSegments(iSegment).strText
                    ReDim .arrFields(UBound(thisMessages.arrMsgs(iMsg).arrSegments(iSegment).arrFields))
                    For iField = 1 To UBound(thisMessages.arrMsgs(iMsg).arrSegments(iSegment).arrFields)
                        .arrFields(iField).blnEnable = thisMessages.arrMsgs(iMsg).arrSegments(iSegment).arrFields(iField).blnEnable
                        .arrFields(iField).intNo = thisMessages.arrMsgs(iMsg).arrSegments(iSegment).arrFields(iField).intNo
                        .arrFields(iField).strDataType = thisMessages.arrMsgs(iMsg).arrSegments(iSegment).arrFields(iField).strDataType
                        .arrFields(iField).strElementName = thisMessages.arrMsgs(iMsg).arrSegments(iSegment).arrFields(iField).strElementName
                        .arrFields(iField).strRecDataDef = thisMessages.arrMsgs(iMsg).arrSegments(iSegment).arrFields(iField).strRecDataDef
                        .arrFields(iField).strRecDataValue = thisMessages.arrMsgs(iMsg).arrSegments(iSegment).arrFields(iField).strRecDataValue
                        .arrFields(iField).strSendDataDef = thisMessages.arrMsgs(iMsg).arrSegments(iSegment).arrFields(iField).strSendDataDef
                        .arrFields(iField).strSendDataValue = thisMessages.arrMsgs(iMsg).arrSegments(iSegment).arrFields(iField).strSendDataValue
                    Next iField
                    '在段复制完之后，针对多次发送医嘱的处理，修改PV1-19
                    If .strName = "PV1" Then
                        .arrFields(19).strSendDataDef = Val(thisMessages.arrMsgs(iMsg).arrSegments(iSegment).arrFields(19).strSendDataDef) + iDuplicate
                    End If
                End With
            Next iSegment
        Next iMsg
    Next iDuplicate
    
    Exit Function
err:
    Call WriteLog(2004, err.Number, "funDuplicateMsg 出现错误，错误描述是：" & err.Description)
    funDuplicateMsg = 1
End Function


Public Sub subNewLogFile()
'功能： 创建新的日志文件
'参数： 无
    
    Dim strDate As String
    
    On Error GoTo err
    
    '创建当前的日期时间标记
    strDate = Date & "-" & Hour(Time) & "-" & Minute(Time) & "-" & Second(Time)
    
    '复制日志文件之前，先关闭日志文件
    If gcnAccess.State <> adStateClosed Then gcnAccess.Close
    FileCopy gstrAccessName, gstrAccessPath & "-" & strDate & ".mdb"
        
    '重新连接数据库
    gcnAccess.Open
    '清空当前日志中的内容
    gcnAccess.Execute "delete from HL7通讯日志"
    gcnAccess.Execute "delete from HL7消息记录"
    gcnAccess.Execute "delete from 错误日志"
    
    
    '压缩数据库文件
    gcnAccess.Close
    DBEngine.CompactDatabase gstrAccessName, gstrAccessPath & "-zip.mdb"
    Kill gstrAccessName
    FileCopy gstrAccessPath & "-zip.mdb", gstrAccessName
    Kill gstrAccessPath & "-zip.mdb"
    gcnAccess.Open
    
    Exit Sub
err:
    Call WriteLog(1013, err.Number, "创建新日志出现错误，错误描述是：" & err.Description)
End Sub


Public Function funGetMessage(strData As String, strMessage As String, strRemain As String) As Boolean
'功能： 从strData中提取一个完整的HL7消息
'参数： strData     --- 源字符串
'       strMessage  --- 提取出来的第一个HL7消息
'       strRemain   --- 提取出HL7消息后剩下的字符串，以MSH段开头
'返回值：True 消息提取成功，剩余字符串为空或者是以MSH开头的下一个消息；False 消息提取不成功，无消息，或者提取出错
    
    '从strData的头开始提取
    '消息的完整性是有首尾标记的，首字符是chr(11)，消息结束符是chr(28)chr(13)
    
    funGetMessage = False
    
    On Error GoTo err
    
    strMessage = ""
    strRemain = ""
    
    If InStr(strData, Chr(11)) <> 0 And InStr(strData, Chr(28) & Chr(13)) <> 0 Then
        '正常处理，包含一个完整消息,提取完整消息，保留消息后面的部分内容。如果消息前面有半截内容，将忽略
        If InStr(strData, Chr(11)) < InStr(strData, Chr(28) & Chr(13)) Then
            strMessage = Mid(strData, InStr(strData, Chr(11)), InStr(strData, Chr(28) & Chr(13)) + 1)
            strRemain = Right(strData, Len(strData) - InStr(strData, Chr(28) & Chr(13)) - 1)
            funGetMessage = True
        End If
    End If
    
    Exit Function
err:
    funGetMessage = False
End Function

Public Function DynamicCreate(ByVal strclass As String, ByVal strCaption As String) As Object
'动态创建对象
    On Error Resume Next
    Set DynamicCreate = CreateObject(strclass)
   
    If err <> 0 Then
        MsgBox strCaption & "组件创建失败，请联系管理员检查是否正确安装!", vbInformation, "提示"
        Set DynamicCreate = Nothing
    End If
    err.Clear
End Function
