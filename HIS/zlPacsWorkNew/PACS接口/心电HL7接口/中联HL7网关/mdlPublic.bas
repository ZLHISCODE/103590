Attribute VB_Name = "mdlPublic"
Option Explicit

Public Declare Function InitCommonControls Lib "comctl32.dll" () As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

'---------------------------------------------------------------
'-ע��� API ����...
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
'- ע��� Api ����...
'---------------------------------------------------------------
' Reg Data Types...
Public Const REG_SZ = 1                         ' Unicode���ս��ַ���
Public Const REG_EXPAND_SZ = 2                  ' Unicode���ս��ַ���
Public Const REG_DWORD = 4                      ' 32-bit ����

' ע���������ֵ...
Public Const REG_OPTION_NON_VOLATILE = 0       ' ��ϵͳ��������ʱ���ؼ��ֱ�����

' ע���ؼ��ְ�ȫѡ��...
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
                     
' ע���ؼ��ָ�����...
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004

' ����ֵ...
Public Const ERROR_NONE = 0
Public Const ERROR_BADKEY = 2
Public Const ERROR_ACCESS_DENIED = 8
Public Const ERROR_SUCCESS = 0

'---------------------------------------------------------------
'- ע���ȫ��������...
'---------------------------------------------------------------
Public Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Boolean
End Type

'---------------------------------------------------------------
'��ȡ�����Ķ��IP
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

'��ʱ��ʱ
Public Declare Function timeGetTime Lib "winmm.dll" () As Long


'------------------------HL7����ʹ�õ��Զ�������-----------------------------
Public Type THL7Service
    lngID As Long                   '����ID
    strIP As String                 '�����IP��ַ
    strSendApp As String            '����ķ��ͳ�������
    strSendFacility As String       '����ķ����豸����
    strReceiveApp As String         '����Ľ��ճ�������
    strReceiveFacility As String    '����Ľ����豸����
    lngPort As Long               '����Ķ˿ں�
    intServiceType As Integer       '��������1-���գ�2-����
    Started  As Boolean             '��ǰ�����Ƿ�ɹ�����
End Type
Public HL7Services() As THL7Service    '�洢Ӧ���ڵ�ǰIP��ַ��HL7�����


Public Function Nvl(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'���ܣ��൱��Oracle��NVL����Nullֵ�ĳ�����һ��Ԥ��ֵ
    Nvl = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

Public Function funcGetLocalIP() As String
'���ص�ǰ�������IP��ַ�����ö��ŷָ�
    Dim hostname As String * 256
    Dim hostent_addr As Long
    Dim host As HOSTENT
    Dim hostip_addr As Long
    Dim temp_ip_address() As Byte
    Dim i As Integer
    Dim ip_address As String
    Dim strLocalIPs As String

    '����Socket
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

    '���Socket
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
'���ܣ���Ϣ��ӣ��Ⱥ򱻴���
'������ strMsg ������IN����Ҫ��ӵ���Ϣ���ݣ�������������Ϣ��Ҳ������һ����Ϣ��
'���أ�True������ӳɹ���False�������ʧ��
'-----------------------------------------------
    Dim Timer As Long
    Dim intCount As Integer
    
    MsgInQueue = False
    
    If gblnQueueBusy = True Then
        '�������æ�����ڽ��ж��д�����ȴ�������ȴ���ʱ������Ϣ�ŵ����ö���
        
        '��¼������־
        Call WriteProcessLog("MsgInQueue", "��Ϣδ���", "�������æ����Ϣδ��ӣ���Ϣǰ��Σ�" + Left(strMsg, 150), 3)
    Else
        '������п��У��������Ӵ������ұ�Ƕ���æ
        gblnQueueBusy = True
        
        On Error GoTo err
        
        '������Ϣ���
        intCount = UBound(gstrMsgQueue) + 1
        ReDim Preserve gstrMsgQueue(intCount) As String
        gstrMsgQueue(intCount) = strMsg
        
        '��¼������־
        Call WriteProcessLog("MsgInQueue", "��Ϣ���", "���յ�������Ϣ������Ϣ���봦����У���Ϣǰ��Σ�" + Left(strMsg, 150), 3)
    
        '���д�����ɣ���Ƕ���Ϊ��
        gblnQueueBusy = False
    End If
        
    MsgInQueue = True
    Exit Function
err:
    '������˳�����ʱ��������
    Call WriteLog(4003, err.Number, "MsgInQueue ���ִ���strMsgǰ��� = " & Left(strMsg, 150) & "�����������ǣ�" & err.Description)
    gblnQueueBusy = False
End Function

Public Function MsgOutQueue() As String
'------------------------------------------------
'���ܣ���Ϣ���ӣ����ù��̸�������ӵ���Ϣ
'������
'���أ����س��ӵ���Ϣ����
'-----------------------------------------------
    Dim iCount As Integer
    Dim strMsg As String
        
    MsgOutQueue = ""        '��ʼ��Ϊ����Ϣ
    
    If gblnQueueBusy = True Then
        '�������æ�����ڽ��ж��д�����ȴ�������ȴ���ʱ���˳����Ӳ���
        
    Else
        '������п��У�����г��Ӵ������ұ�Ƕ���æ
        gblnQueueBusy = True
        
        On Error GoTo err
        
        '��Ϣ���Ӵ���
        iCount = UBound(gstrMsgQueue)
        If iCount = 0 Then
            gblnQueueBusy = False
            Exit Function     '����Ϊ�գ����ó���
        End If
        
        '�Ӷ�������ȡ��Ϣ
        strMsg = gstrMsgQueue(gintQueueIndex)
        
        '�������ָ��
        gintQueueIndex = gintQueueIndex + 1
        If gintQueueIndex > iCount Then
            '�����ǰȡ�������Ƕ����е����һ����Ϣ���������Ϣ����
            ReDim Preserve gstrMsgQueue(0) As String
            gintQueueIndex = 1
        End If
        
        MsgOutQueue = strMsg
        
        '���Ӵ�����ɣ���Ƕ�����
        gblnQueueBusy = False
    End If
    
    Exit Function
err:
    Call WriteLog(4005, err.Number, "MsgOutQueue ���ִ���strMsgǰ��� = " & Left(strMsg, 150) & "�����������ǣ�" & err.Description)
    gblnQueueBusy = False
End Function

Public Function funGetAMessage(ByRef strMsg As String) As Boolean
'------------------------------------------------
'���ܣ�����Ϣ������ȡһ����Ϣ
'������ strMsg -- ������ȡ������Ϣ,�ձ�ʾ�������û����Ϣ
'���أ� True - �ɹ���False - ʧ��
'-----------------------------------------------
    
    funGetAMessage = False
    
    On Error GoTo err
    
    '����Ϣ������ȡһ����Ϣ
    strMsg = MsgOutQueue
    
    '�����Ϣ��Ϊ�գ����������Ϣ
    If strMsg <> "" Then
        funGetAMessage = True
    End If
    
    Exit Function
err:
    '�ݲ�����
End Function


Public Function funMsgProcess() As Boolean
'------------------------------------------------
'���ܣ��Զ�������Ϣ
'������ ��
'���أ�True -- �ɹ��� False -- ʧ��
'-----------------------------------------------
    Dim strMsg As String
    
    funMsgProcess = False
    
    '����Ѿ�������Ϣ�����ˣ����˳�
    If gblnMsgProcessing = True Then Exit Function
    
    '������Ϣ�����ǣ���ֹ������̱���ε���
    gblnMsgProcessing = True
    
    On Error GoTo err
    
    '��Ϣ���Ӳ�������Ϣ
    While funGetAMessage(strMsg) = True
        '��¼��־
        Call WriteProcessLog("funMsgProcess", "������Ϣ", "��Ϣǰ��� = " & Left(strMsg, 150), 2)

        '������������Ϣ
        Call funParseInMsg(strMsg)
    Wend
    
    '��Ϣ������ɣ��˳�����
    gblnMsgProcessing = False
    
    funMsgProcess = True
    Exit Function
err:
    Call WriteLog(4006, err.Number, "funMsgProcess ���ִ���strMsgǰ��� = " & Left(strMsg, 150) & "�����������ǣ�" & err.Description)
    gblnMsgProcessing = False
End Function

Public Sub WriteProcessLog(logSubName As String, logTitle As String, logDesc As String, lngLogLevel As Long)
'------------------------------------------------
'���ܣ���¼ͨѶ��־
'������ logSubName  --  ������־�ĺ�����
'       logTitle   --   ��־����
'       logDesc   --    ��־����
'       lngLogLevel --  ��־����ͨ����־����ȷ����ǰ��־�Ƿ���Ҫ��¼
'���أ���
'------------------------------------------------

    Dim strSQL As String
    
    On Error GoTo err
    
    '�����˼�¼��־���ż�¼��ǰ����־,�ж���־����ȷ��������־�Ƿ���Ҫ��¼
    If gblnProcessLog And glngProcessLogLevel >= lngLogLevel Then
        If gcnAccess.State = adStateClosed Then Exit Sub
        
        '����־�����еĵ����Ž���ת�壬���򱣴浽Access���ݿ�����
        logDesc = Replace(logDesc, "'", "��")
        
        strSQL = "Insert into HL7ͨѶ��־ (ͨѶʱ��,ͨѶ����,��¼����,��¼����) " & _
            "Values( cDate('" & Date & " " & Time() & "'),'" & logSubName & "','" & logTitle & _
            "','" & logDesc & "')"
        gcnAccess.Execute strSQL
    End If
    Exit Sub
err:
    Call WriteLog(9001, err.Number, "WriteProcessLog ��¼ͨѶ��־����" & ",logSubName=" & logSubName & "��logTitle=" & logTitle & "�����������ǣ�" & err.Description)
End Sub

Public Sub WriteLog(ByVal ErrorType As Integer, ErrorNum As Long, ErrorDesc As String)
'-----------------------------------------------------------------------------
'����:��д������־
'������ ErrorType ----�������ʹ��룬����ͼ�����100��WORKLIST��QR����200��FTP����300,funSplitSeriesUID����1001,�ļ�ͨѶ����4000
'       ErrorNum ----�����
'       ErrorDesc ----��������
'����ֵ����
'-----------------------------------------------------------------------------
    Dim strSQL As String
    On Error Resume Next
    If gcnAccess.State = adStateClosed Then Exit Sub
    
    '����־�����еĵ����Ž���ת�壬���򱣴浽Access���ݿ�����
    ErrorDesc = Replace(ErrorDesc, "'", "��")
        
    strSQL = "Insert Into ������־(����ʱ��,��������,�����,������Ϣ) " & _
        "Values(cDate('" & Date & " " & Time() & "')," & ErrorType & "," & ErrorNum & ",'" & ErrorDesc & "')"
    
    gcnAccess.Execute strSQL
End Sub

Public Sub WriteMessageLog(strMessageType As String, strMessage As String)
'------------------------------------------------
'���ܣ���¼���յ����Ҵ���ɹ�����Ϣ
'������ strMessageType  --  ��Ϣ��������
'       strMessage   --   ��Ϣ����
'���أ���
'------------------------------------------------

    Dim strSQL As String
    
    On Error Resume Next
    
    If gcnAccess.State = adStateClosed Then Exit Sub
    
    '����־�����еĵ����Ž���ת�壬���򱣴浽Access���ݿ�����
    strMessage = Replace(strMessage, "'", "��")
    
    strSQL = "Insert into HL7��Ϣ��¼ (ͨѶʱ��,��Ϣ����,��Ϣ����) " & _
        "Values( cDate('" & Date & " " & Time() & "'),'" & strMessageType & "','" & strMessage & "')"
    gcnAccess.Execute strSQL
    
End Sub

Public Function funMsgFullType(strMsg As String) As Long
'-----------------------------------------------------------------------------
'����:�����Ϣ����������
'������ strMsg ----��Ϣԭ��
'����ֵ��0 -- ��������Ϣ��1 -- ����Ϣͷ��2 -- ����Ϣβ��3 -- ����Ϣ�м�Σ�4 -- ����
'-----------------------------------------------------------------------------
    On Error GoTo err
    
    '�����Ϣ�������ԣ������������Ϣ��ֱ����ӣ���������������ȷ�����ʱ���еȴ�������Ϣ
    '��Ϣ��������������β��ǵģ����ַ���chr(11)����Ϣ��������chr(28)chr(13)
    If Left(strMsg, 1) = Chr(11) And InStr(strMsg, Chr(28) & Chr(13)) <> 0 Then
        funMsgFullType = 0  '������Ϣ
    ElseIf Left(strMsg, 1) = Chr(11) Then
        funMsgFullType = 1  '��Ϣͷ
    ElseIf InStr(strMsg, Chr(28) & Chr(13)) <> 0 Then
        funMsgFullType = 2  '��Ϣβ
    Else
        funMsgFullType = 3  '��Ϣ�м��
    End If
    
    Exit Function
err:
    funMsgFullType = 4  '����
End Function

Public Function getMsgDefFromDB(strActionType As String) As THl7Messages
'-----------------------------------------------------------------------------
'����:���ݶ������ͣ������ݿ��ж�ȡ��Ҫ�����HL7��Ϣ����
'������ strActionType ----��������
'����ֵ��������֯�õ�HL7��Ϣ����
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
    
    '���ȴӡ�hl7��Ϣ���塱���ж�ȡ��Ҫ���͵�HL7��Ϣ�����
    strSQL = "Select a.ID,a.����ID,a.��������,a.��Ϣ����,a.��Ϣ����,a.��Ϣ�����,b.IP��ַ,b.�˿ں� " & _
                " From zlhis.hl7��Ϣ���� a,zlhis.HL7�������� b Where a.����ID = b.Id And �������� = 2 and a.�������� =[1]"
    Set rsMsg = gzlDatabase.OpenSQLRecord(strSQL, "����ҽ����Ϣ�Ķ���", strActionType)
    
    strSQL = "Select ��Ϣid, ��Ϣ������, �������, ��������, ��������ֵ, ��������ֵ, Ԫ������ " & _
             " From zlhis.Hl7��Ϣ������ a ,zlhis.hl7��Ϣ���� b Where  b.�������� =[1] And a.��Ϣid = b.Id order by ��Ϣ������,�������"
    Set rsSegment = gzlDatabase.OpenSQLRecord(strSQL, "����ҽ����Ϣ������", strActionType)
    
    If rsMsg.EOF = False Then
        ReDim arrHL7Msgs.arrMsgs(rsMsg.RecordCount)
        
        rsMsg.MoveFirst
        For iMsg = 1 To rsMsg.RecordCount
            arrHL7Msgs.arrMsgs(iMsg).lngID = rsMsg!ID
            arrHL7Msgs.arrMsgs(iMsg).lngServiceID = rsMsg!����ID
            arrHL7Msgs.arrMsgs(iMsg).strActionType = Nvl(rsMsg!��������)
            arrHL7Msgs.arrMsgs(iMsg).strMsgName = Nvl(rsMsg!��Ϣ����)
            arrHL7Msgs.arrMsgs(iMsg).strMsgType = Nvl(rsMsg!��Ϣ����)
            arrHL7Msgs.arrMsgs(iMsg).strMsgSegmentDef = Nvl(rsMsg!��Ϣ�����)
            arrHL7Msgs.arrMsgs(iMsg).strIP = Nvl(rsMsg!IP��ַ)
            arrHL7Msgs.arrMsgs(iMsg).lngPort = Nvl(rsMsg!�˿ں�, 0)
            arrHL7Msgs.arrMsgs(iMsg).blnSendOK = False
            
            arrLoadSegments = Split(arrHL7Msgs.arrMsgs(iMsg).strMsgSegmentDef, "|")
            ReDim arrHL7Msgs.arrMsgs(iMsg).arrSegments(UBound(arrLoadSegments) + 1)
            '���ÿһ����Ϣ�Ķ�
            
            For iSeg = 0 To UBound(arrLoadSegments)
                rsSegment.Filter = "��Ϣ������ = '" & arrLoadSegments(iSeg) & "'"
                If rsSegment.EOF = False Then
                    With arrHL7Msgs.arrMsgs(iMsg).arrSegments(iSeg + 1)
                        .intNo = iSeg + 1
                        .strName = arrLoadSegments(iSeg)
                        ReDim .arrFields(rsSegment.RecordCount)
                        
                        
                        rsSegment.MoveFirst
                        For iField = 1 To rsSegment.RecordCount
                            .arrFields(iField).intNo = iField
                            .arrFields(iField).strDataType = Nvl(rsSegment!��������)
                            .arrFields(iField).strElementName = Nvl(rsSegment!Ԫ������)
                            .arrFields(iField).strRecDataDef = Nvl(rsSegment!��������ֵ)
                            .arrFields(iField).strSendDataDef = Nvl(rsSegment!��������ֵ)
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

    Call WriteLog(2001, err.Number, "getMsgDefFromDB ���ִ��󣬴��������ǣ�" & err.Description)
End Function

Public Function funfillMsgValue(thisMessages As THl7Messages, strWorkIDs As String) As Long
'-----------------------------------------------------------------------------
'����:����ҵ��ID�����HL7��Ϣ������
'������ thisMessages -- HL7��Ϣ����
'       strWorkIDs -- ҵ��ID����ʹ�á�;�����Ӷ��ID������ҽ��ʱ���ǡ�ҽ��ID;���ͺš�
'����ֵ��0-�ɹ���1-ҽ����Ϣ��ҵ��ID���ǿգ��޷�������Ϣ��2-ʧ��
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
    
    '��ϢΪ�գ����˳�
    If UBound(thisMessages.arrMsgs) = 0 Then Exit Function
    
    strSpliterL = "##L##"
    strSpliterR = "##R##"
    
    '�����ݿ��ж�ȡҽ��
    '�ж��Ƿ�ҽ����Ϣ������ֻ����ҽ����Ϣ����ҽ����ȡ��ҽ����ɾ��ҽ��
    
    For iMsg = 1 To UBound(thisMessages.arrMsgs)
        strOneMsg = ""
        
        If thisMessages.arrMsgs(iMsg).strActionType = HL7_MSG_SEND_NEW_ORDER _
            Or thisMessages.arrMsgs(iMsg).strActionType = HL7_MSG_SEND_CANCEL_ORDER Then
            
            '��ʱû��ɾ��ҽ����Ҫ���� Or thisMessages.arrMsgs(iMsg).strActionType = HL7_MSG_SEND_DEL_ORDER Then
            
            'ҽ����Ϣ��ҵ��ID���ǣ�ҽ��ID�����ͺ�
            arrWorkIDs = Split(strWorkIDs, ";")
            If UBound(arrWorkIDs) <> 1 Then
                funfillMsgValue = 1     'ҽ����Ϣ��ҵ��ID���ǿ�
                Exit Function
            End If
            
            '��ȡ���ݿ�ʱ���ʱ��������뼶2λ
            strSQL = "Select to_char(current_timestamp,'YYYYMMDDHH24MISSFF2') as MSGControlID From dual"
            Set rsTemp = gzlDatabase.OpenSQLRecord(strSQL, "��ȡ��ǰʱ��")
            strSysDate = rsTemp!MSGControlID
            
            '����ҵ��ID��(ҽ��ID�����ͺ�)��ȡ��Ҫ���͵�ҽ����Ϣ
            '����ҽ����¼.ҽ����Ч ----- 0-����;1-��ʱ
            '����ҽ����¼.ҽ��״̬ --- --1-δ��Ч���ݴ�ҽ����1-�¿���2-У�����ʣ�3-��У�ԣ�4-�����ϣ�5-��������6-����ͣ��7-�����ã�8-��ֹͣ��9-��ȷ��ֹͣ
            '����ҽ��״̬.�������� --- 1-�¿���2-У�����ʣ�3-У��ͨ����4-���ϣ�5-������6-��ͣ��7-���ã�8-ֹͣ��9-ȷ��ֹͣ��10-Ƥ�Խ��,11-���ͨ����12-���δͨ����13-ʵϰҽʦͣ��������
            '����ҽ������.ִ��״̬ ---- 0-δִ��;1-��ȫִ��;2-�ܾ�ִ��;3-����ִ��(�����ֽܷ�Ϊ����ʵ�ʲ���)
            
            strSQL = "Select ����ID from zlhis.����ҽ����¼ where id =[1]"
            Set rsTemp = gzlDatabase.OpenSQLRecord(strSQL, "��ѯҽ��������Ϣ", CLng(arrWorkIDs(0)), CLng(arrWorkIDs(1)))
            If rsTemp.RecordCount = 0 Then Exit Function
            
            lngPatientID = rsTemp!����ID
                        
            
            '���ж�ҽ��״̬�Ƿ����ǰ��Ϣ״̬һ�£������޷��ж�
            
            On Error Resume Next
            
            '���ÿһ��ҽ����Ϣ
            For iSeg = 1 To UBound(thisMessages.arrMsgs(iMsg).arrSegments)
                
                strOneSegment = thisMessages.arrMsgs(iMsg).arrSegments(iSeg).strName
                
                '��д��Ϣ�ε�ÿһ���ֶ�
                For iField = 1 To UBound(thisMessages.arrMsgs(iMsg).arrSegments(iSeg).arrFields)
                    With thisMessages.arrMsgs(iMsg).arrSegments(iSeg).arrFields(iField)
                        strValue = .strSendDataDef
                        
                        '���뷵���ַ���
                        Do While InStr(strValue, "[") <> 0
                            '�����ַ��������Ϲ���ģ�ֱ���˳�ѭ�������ؿ�
                            If InStr(strValue, "]") = 0 Or InStr(strValue, "]") < InStr(strValue, "[") Then
                                strValue = ""
                                Exit Do
                            End If
                        
                            strField = Mid(strValue, InStr(strValue, "[") + 1, InStr(strValue, "]") - InStr(strValue, "[") - 1)
                            
                            strFieldValue = ""
                            If strField = "��ǰʱ��" Then
                                strFieldValue = strSysDate
                            ElseIf strField = "��Ϣ����" Then
                                strFieldValue = thisMessages.arrMsgs(iMsg).strMsgType
                            Else
                                strFieldValue = funGetFieldValueByFun(strField, lngPatientID, CLng(arrWorkIDs(0)), CLng(arrWorkIDs(1)), thisMessages.arrMsgs(iMsg).lngID)
                            End If
                            
                            '�滻��ֵ���п��ܳ��ֵ�[]�ָ���
                            strFieldValue = Replace(strFieldValue, "[", strSpliterL)
                            strFieldValue = Replace(strFieldValue, "]", strSpliterR)
                            
                            strValue = Replace(strValue, "[" & strField & "]", strFieldValue)
                        Loop
                        
                        '������֮�󣬽�ԭ����[]�ָ����滻��ȥ
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
    Call WriteLog(2002, err.Number, "fillMsgValue ���ִ��󣬴��������ǣ�" & err.Description)
End Function

Private Function funGetFieldValueByFun(strField As String, lngPatientID As Long, lngOrderID As Long, _
    lngSendNo As Long, lngMesageID As Long) As String
'-----------------------------------------------------------------------------
'����:ͨ�����ݿ�ĺ��� HL7_Replace_Element_Value ��ȡ��Ӧ������ֵ
'������ strField -- �ֶ�����
'       lngPatientID ---����ID
'       lngOrderID --- ҽ��ID
'       lngSendNo --- ���ͺ�
'       lngMesageID --- ��ϢID
'����ֵ���ֶζ�Ӧ�ķ���ֵ
'-----------------------------------------------------------------------------
    On Error GoTo err
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "Select zlhis.b_hl7interface.HL7_Replace_Element_Value([1],[2],[3],[4],[5]) as ���ֵ from dual "
    Set rsTemp = gzlDatabase.OpenSQLRecord(strSQL, "��ѯ�ֶζ�Ӧ��ֵ", strField, lngPatientID, lngOrderID, lngSendNo, lngMesageID)
    
    If rsTemp.RecordCount > 0 Then
        funGetFieldValueByFun = Nvl(rsTemp!���ֵ)
    End If
    
    Exit Function
err:
    Call WriteLog(2005, err.Number, "funGetFieldValueByFun ���ִ��󣬴��������ǣ�" & err.Description)
End Function

Public Function funDuplicateMsg(thisMessages As THl7Messages, intTimes As Integer) As Long
'-----------------------------------------------------------------------------
'����:���ݷ��ʹ�����������Ϣ����Գ����Ķ�η���
'������ thisMessages -- HL7��Ϣ
'       intTimes -- ���ʹ���
'����ֵ��0-�ɹ���1-ʧ��
'-----------------------------------------------------------------------------
    Dim iMsgCount As Integer
    Dim iMsg As Integer
    Dim iSegment As Integer
    Dim iField As Integer
    Dim iDuplicate As Integer
    
    
    On Error GoTo err
    
    If intTimes <= 1 Then Exit Function
    
    iMsgCount = UBound(thisMessages.arrMsgs)
    
    '��ϢΪ�գ����˳�
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
                    '�ڶθ�����֮����Զ�η���ҽ���Ĵ����޸�PV1-19
                    If .strName = "PV1" Then
                        .arrFields(19).strSendDataDef = Val(thisMessages.arrMsgs(iMsg).arrSegments(iSegment).arrFields(19).strSendDataDef) + iDuplicate
                    End If
                End With
            Next iSegment
        Next iMsg
    Next iDuplicate
    
    Exit Function
err:
    Call WriteLog(2004, err.Number, "funDuplicateMsg ���ִ��󣬴��������ǣ�" & err.Description)
    funDuplicateMsg = 1
End Function


Public Sub subNewLogFile()
'���ܣ� �����µ���־�ļ�
'������ ��
    
    Dim strDate As String
    
    On Error GoTo err
    
    '������ǰ������ʱ����
    strDate = Date & "-" & Hour(Time) & "-" & Minute(Time) & "-" & Second(Time)
    
    '������־�ļ�֮ǰ���ȹر���־�ļ�
    If gcnAccess.State <> adStateClosed Then gcnAccess.Close
    FileCopy gstrAccessName, gstrAccessPath & "-" & strDate & ".mdb"
        
    '�����������ݿ�
    gcnAccess.Open
    '��յ�ǰ��־�е�����
    gcnAccess.Execute "delete from HL7ͨѶ��־"
    gcnAccess.Execute "delete from HL7��Ϣ��¼"
    gcnAccess.Execute "delete from ������־"
    
    
    'ѹ�����ݿ��ļ�
    gcnAccess.Close
    DBEngine.CompactDatabase gstrAccessName, gstrAccessPath & "-zip.mdb"
    Kill gstrAccessName
    FileCopy gstrAccessPath & "-zip.mdb", gstrAccessName
    Kill gstrAccessPath & "-zip.mdb"
    gcnAccess.Open
    
    Exit Sub
err:
    Call WriteLog(1013, err.Number, "��������־���ִ��󣬴��������ǣ�" & err.Description)
End Sub


Public Function funGetMessage(strData As String, strMessage As String, strRemain As String) As Boolean
'���ܣ� ��strData����ȡһ��������HL7��Ϣ
'������ strData     --- Դ�ַ���
'       strMessage  --- ��ȡ�����ĵ�һ��HL7��Ϣ
'       strRemain   --- ��ȡ��HL7��Ϣ��ʣ�µ��ַ�������MSH�ο�ͷ
'����ֵ��True ��Ϣ��ȡ�ɹ���ʣ���ַ���Ϊ�ջ�������MSH��ͷ����һ����Ϣ��False ��Ϣ��ȡ���ɹ�������Ϣ��������ȡ����
    
    '��strData��ͷ��ʼ��ȡ
    '��Ϣ��������������β��ǵģ����ַ���chr(11)����Ϣ��������chr(28)chr(13)
    
    funGetMessage = False
    
    On Error GoTo err
    
    strMessage = ""
    strRemain = ""
    
    If InStr(strData, Chr(11)) <> 0 And InStr(strData, Chr(28) & Chr(13)) <> 0 Then
        '������������һ��������Ϣ,��ȡ������Ϣ��������Ϣ����Ĳ������ݡ������Ϣǰ���а�����ݣ�������
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
'��̬��������
    On Error Resume Next
    Set DynamicCreate = CreateObject(strclass)
   
    If err <> 0 Then
        MsgBox strCaption & "�������ʧ�ܣ�����ϵ����Ա����Ƿ���ȷ��װ!", vbInformation, "��ʾ"
        Set DynamicCreate = Nothing
    End If
    err.Clear
End Function
