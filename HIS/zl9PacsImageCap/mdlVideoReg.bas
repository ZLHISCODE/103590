Attribute VB_Name = "mdlVideoReg"
Option Explicit


Public Const LOGIN_TYPE_��Ƶ�豸 As String = "Ӱ����Ƶ�豸����"


Public gint��Ƶ�豸���� As Integer

Public Function funVideoRegTime(frmParent As Form) As String
'���ܣ�����ע����Ϣ�����򷵻�ע��ʱ��
'������ frmParent ---������
'       str���� ---'��ע������ʹ�õ���������
'����ֵ����ǰ���ڣ���ע����Ϣ���ؿ�
On Error GoTo err
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim strIP��ַ As String         'ע���IP��ַ
    
    funVideoRegTime = ""
    
    If gint��Ƶ�豸���� <= -1 Then
        funVideoRegTime = Now
        Exit Function
    End If
    
    strIP��ַ = funGetOneIP(frmParent)
    
    strSQL = "select ����վ from zltools.zlclients where ip=[1] and ������ƵԴ=1"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡע����Ϣ", strIP��ַ)
    
    If Not rsTemp.EOF Then funVideoRegTime = Now
    Exit Function
err:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
    err.Clear
End Function

Public Function FunLogIn(frmParent As Form, str���� As String) As String
'���ܣ��Գ������ע�ᣬ���ע��ɹ����򷵻�ע��ʱ��
'������ frmParent ---������
'       str���� ---'��ע������ʹ�õ���������
'����ֵ��ע��ɹ�ע�����ڣ�ע��ʧ�ܷ��ؿ�

    Dim intNum As Integer
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim strIP��ַ As String         '��Ҫע���IP��ַ
    
    On Error GoTo err
    
    strIP��ַ = funGetOneIP(frmParent)
    
    '��ע��������ȡ��Ȩ��������-1--�����ƣ�0--��ֹ��X��X>0��--������������
    intNum = gint��Ƶ�豸����
    
    'intNUM >0 ,����ù���ע�����
    If intNum > 0 Then  '����������
        strSQL = "Zl_Ӱ�������¼_Update('" & strIP��ַ & "','" & str���� & "'," & intNum & ")"
        zlDatabase.ExecuteProcedure strSQL, "ע��" & str����
        '���ע���Ƿ�ɹ�
        strSQL = "Select ����ʱ��,IP��ַ from Ӱ�������¼ where  ����=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ����ʱ��", str����)
        
        If rsTemp.RecordCount <= intNum Then
            rsTemp.Filter = "IP��ַ='" & strIP��ַ & "'"
            If rsTemp.RecordCount = 1 Then  'ע��ɹ�
                FunLogIn = rsTemp!����ʱ��
                Exit Function
            End If
        End If
    ElseIf intNum = -1 Then     '������
        FunLogIn = Now
        Exit Function
    Else    '=0����������ֵ����ֹ�������κδ�����������ʾ
    
    End If
    
    'ע��ʧ�ܣ�����������ԭ��
    '1��ע���������������ɵ��������޷�ע��IP��ַ
    '2��ֱ��ͨ��SQL����������IP��ַ�����±��еļ�¼��������������ɵ�����
    Call MsgboxCus("�򿪵�" & str���� & "�������������������" & intNum & "�������������Ӧ����ϵ��", vbOKOnly, G_STR_HINT_TITLE)
    FunLogIn = ""
    
    Exit Function
err:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
    err.Clear
End Function

Public Function FunCheckRegInfo(frmParent As Form) As Boolean
'���ܣ�����Ƿ����ע���ip��ַ����������ƵԴ
On Error GoTo errHandle
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim strIP��ַ As String         '��Ҫע���IP��ַ
    
    FunCheckRegInfo = False
    
    strIP��ַ = funGetOneIP(frmParent)
    
    strSQL = "select ����վ from zltools.zlclients where ip=[1] and ������ƵԴ=1"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡע����Ϣ", strIP��ַ)
    
    If rsTemp.EOF = False Then FunCheckRegInfo = True
    
Exit Function
errHandle:
End Function

Public Function FunCheckIp(frmParent As Form, str���� As String) As Boolean
'���ܣ�����Ƿ����ע���ip��ַ
On Error GoTo errHandle
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim strIP��ַ As String         '��Ҫע���IP��ַ
    
    FunCheckIp = False
    
    strIP��ַ = funGetOneIP(frmParent)
    
    strSQL = "Select ����ʱ�� from Ӱ�������¼ where ����=[2] and IP��ַ=[1] "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ����ʱ��", strIP��ַ, str����)

    
    If rsTemp.EOF = False Then FunCheckIp = True
    
Exit Function
errHandle:
End Function

Public Function FunLogOut(frmParent As Form, str���� As String, str����ʱ�� As String) As Boolean
'���ܣ��˳������ʱ�򣬼������Ƿ�Ϸ�ע�������������ͨ�����������ֶζ�ʱɾ����Ӱ�������¼�����еļ�¼��
'������ frmParent ---������
'       str���� ---'��ע������ʹ�õ���������
'       str����ʱ�� --- ע�Ṥ��վʱ���ص�ʱ��
'����ֵ���Ϸ�ע��True���Ƿ�������False
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim strIP��ַ As String         '��Ҫע���IP��ַ
    Dim intNum As Integer
    
    On Error GoTo err
    strIP��ַ = funGetOneIP(frmParent)
    
    '����ʱ��Ϊ�գ���ʾע��ʧ�ܣ�û����������������˳���ʱ���ټ�����ݿ�
    If str����ʱ�� = "" Then
        FunLogOut = True
        Exit Function
    End If
    
    '��ע��������ȡ��Ȩ��������-1--�����ƣ�0--��ֹ��X��X>0��--������������
    intNum = gint��Ƶ�豸����
    
    If intNum > 0 Then '������������
        strSQL = "Select ����ʱ�� from Ӱ�������¼ where IP��ַ=[1] and ����=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ����ʱ��", strIP��ַ, str����)
        If rsTemp.EOF = False Then
            FunLogOut = True
        Else
            '�Ա�����ʱ������ݿ��ʱ�䣬�������ͬһ�죬˵����ǰһ�쿪�������ע����Ϣ��ɾ���ˣ�
            '���������Ϊ�ǺϷ�ע��
            strSQL = "Select sysdate from dual"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ���ݿ�ʱ��")
            If Format(rsTemp!sysdate, "yyyy-mm-dd") <> Format(str����ʱ��, "yyyy-mm-dd") Then
                FunLogOut = True
            Else
                FunLogOut = False
            End If
        End If
    ElseIf intNum = -1 Then     '������
        FunLogOut = True
    Else    '=0����������ֵ����ֹ
    
    End If
    If FunLogOut = False Then
        Call MsgboxCus("�򿪵�" & str���� & "�������������������" & intNum & "�������������Ӧ����ϵ��", vbOKOnly, G_STR_HINT_TITLE)
    End If
    Exit Function
err:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Function




Public Function getLicenseCount(strLicenseName As String) As Integer
'��ȡ��Ȩ������
'������ strLicenseName --- ��Ȩ����
    Dim strLiceseCount As String
    
    On Error GoTo err
    
    strLiceseCount = zlRegInfo(strLicenseName)
    If strLiceseCount = "" Then '������
        getLicenseCount = -1
    ElseIf Val(strLiceseCount) > 0 Then '������������
        getLicenseCount = Val(strLiceseCount)
    Else '��ֹ
        getLicenseCount = 0
    End If
    
    Exit Function
err:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Function



Private Function funGetOneIP(frmParent As Form) As String
'------------------------------------------------
'���ܣ���ȡ��ǰ��������׸�IP��ַ
'������ frmParent  -- ������
'���أ����ض�ȡ��ǰ��������׸�IP��ַ
'------------------------------------------------
    Dim strIP��ַ As String
    
    On Error Resume Next
    
    strIP��ַ = funcGetLocalIP(frmParent)
    If strIP��ַ = "" Then
        funGetOneIP = "127.0.0.1"
    ElseIf InStr(strIP��ַ, ",") <> 0 Then
        funGetOneIP = Split(strIP��ַ, ",")(0)
    Else
        funGetOneIP = strIP��ַ
    End If
End Function




Private Function funcGetLocalIP(frmParent As Form) As String
'------------------------------------------------
'���ܣ���ȡ��ǰ�������IP��ַ�����ö��ŷָ�
'������ frmParent  -- ������
'���أ����ص�ǰ�������IP��ַ�����ö��ŷָ�
'------------------------------------------------
    Dim hostname As String * 256
    Dim hostent_addr As Long
    Dim host As HOSTENT
    Dim hostip_addr As Long
    Dim temp_ip_address() As Byte
    Dim i As Integer
    Dim ip_address As String
    Dim strLocalIPs As String

    '����Socket
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

    '���Socket
    Call SocketsCleanup(frmParent)

    funcGetLocalIP = strLocalIPs
End Function




Private Sub SocketsInitialize(frmParent As Form)
'------------------------------------------------
'���ܣ���ʼ��Socket
'������ frmParent  -- ������
'���أ���
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
'���ܣ����Socket
'������ frmParent  -- ������
'���أ���
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

