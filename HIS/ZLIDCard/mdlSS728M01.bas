Attribute VB_Name = "mdlSS728M01"
Option Explicit
Private mblnReadCard        As Boolean

Public Enum led_Color
    ��ɫ = 0
    ��ɫ = 1
    ��ɫ = 2
End Enum

Public Enum led_Status
    �� = 1 '01H������
    �� = 2 '02H����
    ��˸ = 3 '03H����˸��
End Enum

Public Enum Net_Ip
    ����IP = 1
    ����IP = 2
End Enum

Public Type typ_SS728M01
    '2.3.3ȡ����
    ss_id_query_name        As String
    '2.3.4ȡ�Ա����
    ss_id_query_sexid       As String
    '2.3.5ȡ�Ա�����
    ss_id_query_sex         As String
    '2.3.6ȡ�������
    ss_id_query_folkid      As String
    '2.3.7ȡ��������
    ss_id_query_folk        As String
    '2.3.8ȡ��������
    ss_id_query_birth       As String
    '2.3.9ȡסַ
    ss_id_query_address     As String
    '2.3.10ȡ��ݺ���
    ss_id_query_number      As String
    '2.3.11ȡǩ������
    ss_id_query_organ       As String
    '2.3.12ȡ��Ч������ʼ����
    ss_id_query_beginterm   As String
    '2.3.13ȡ��Ч���޽�ֹ����
    ss_id_query_endterm     As String
    '2.3.14ȡ��Ƭ
    ss_id_query_photofile   As String
    '2.3.15 ȡ����סַ
    ss_id_query_newaddr     As String
End Type

Public SS728M01 As typ_SS728M01

'====================================================================================================================================================
'2.2 ͨ�ú���
'====================================================================================================================================================

'2.2.1��ȡ�ӿڿ���Ϣ
'long __stdcall ss_lib_version(char* pszVerInfo);
Public Declare Function ss_lib_version Lib "SS728M01.dll" _
    (ByVal pszVerInfo As String) As Long

'2.2.2�򿪱����ն�
'long __stdcall ss_dev_open(char* szDevCom);
Public Declare Function ss_dev_open Lib "SS728M01.dll" _
    (ByVal szDevCom As String) As Long

'2.2.3�رձ����ն�
'long __stdcall ss_dev_close(long icdev);
Public Declare Function ss_dev_close Lib "SS728M01.dll" _
    (ByVal icdev As Long) As Long

'2.2.4ע�������ն�
'long __stdcall ss_dev_login(char* szDevIP);
Public Declare Function ss_dev_login Lib "SS728M01.dll" _
    (ByVal szDevIP As String) As Long
    
'2.2.5ע�������ն�
'long __stdcall ss_dev_logout(long icdev);
Public Declare Function ss_dev_logout Lib "SS728M01.dll" _
    (ByVal icdev As Long) As Long

'2.2.6��ȡ�ն˰汾
'long __stdcall ss_dev_version(long icdev, char* pszVerInfo);
Public Declare Function ss_dev_version Lib "SS728M01.dll" _
    (ByVal icdev As Long, ss_dev_version As String) As Long

'2.2.7������
'long __stdcall ss_dev_beep(long icdev, unsigned short _Amount, unsigned short _Msec);
Public Declare Function ss_dev_beep Lib "SS728M01.dll" _
    (ByVal icdev As Long, ByVal Amount As Long, ByVal Msec As Long) As Long
'2.2.8ָʾ��
'long __stdcall ss_dev_led(long icdev, unsigned char _Color, unsigned char _Status, unsigned short _Amount, unsigned short _Msec);
'icdev [in]���豸��ʶ���� _Color [in]��ָʾ����ɫ������00H����ɫ����01H����ɫ����02H����ɫ���� _Status [in]��ָʾ��״̬������01H��������02H���𣩡�03H����˸����
'_Amount [in]����˸������ _Msec [in]����˸ʱ������λ100���롣
Public Declare Function ss_dev_led Lib "SS728M01.dll" _
    (ByVal icdev As Long, ByVal Color As led_Color, ByVal Status As led_Status, ByVal Amount As Long, ByVal Msec As Long) As Long
'2.2.9��ȡ�������
'long __stdcall ss_dev_getnet(long icdev, unsigned char _Type, char* pszParam);
Public Declare Function ss_dev_getnet Lib "SS728M01.dll" _
    (ByVal icdev As Long, ByVal ipType As Net_Ip, ByVal pszParam As String) As Long
'2.2.10 �����������
'long __stdcall ss_dev_setnet(long icdev, unsigned char _Type, char* pszParam);
Public Declare Function ss_dev_setnet Lib "SS728M01.dll" _
    (ByVal icdev As Long, ByVal ipType As Net_Ip, ByVal pszParam As String) As Long
    
'====================================================================================================================================================
'2.3 �ڶ����������֤����
'====================================================================================================================================================

'2.3.1Ѱ��
'long __stdcall ss_id_find_card(long icdev);
Public Declare Function ss_id_find_card Lib "SS728M01.dll" _
    (ByVal icdev As Long) As Long

'2.3.2��ȡ����Ϣ
'long __stdcall ss_id_read_card(long icdev);
Public Declare Function ss_id_read_card Lib "SS728M01.dll" _
    (ByVal icdev As Long) As Long

'2.3.3ȡ����
'long __stdcall ss_id_query_name(long icdev, char* pszText);
Public Declare Function ss_id_query_name Lib "SS728M01.dll" _
    (ByVal icdev As Long, ByVal pszText As String) As Long

'2.3.4ȡ�Ա����
'long __stdcall ss_id_query_sexid(long icdev, char* pszText);
Public Declare Function ss_id_query_sexid Lib "SS728M01.dll" _
    (ByVal icdev As Long, ByVal pszText As String) As Long
    
'2.3.5ȡ�Ա�����
'long __stdcall ss_id_query_sex(long icdev, char* pszText);
Public Declare Function ss_id_query_sex Lib "SS728M01.dll" _
    (ByVal icdev As Long, ByVal pszText As String) As Long
    
'2.3.6ȡ�������
'long __stdcall ss_id_query_folkid(long icdev, char* pszText);
Public Declare Function ss_id_query_folkid Lib "SS728M01.dll" _
    (ByVal icdev As Long, ByVal pszText As String) As Long

'2.3.7ȡ��������
'long __stdcall ss_id_query_folk(long icdev, char* pszText);
Public Declare Function ss_id_query_folk Lib "SS728M01.dll" _
    (ByVal icdev As Long, ByVal pszText As String) As Long
    
'2.3.8ȡ��������
'long __stdcall ss_id_query_birth(long icdev, char* pszText);
Public Declare Function ss_id_query_birth Lib "SS728M01.dll" _
    (ByVal icdev As Long, ByVal pszText As String) As Long

'2.3.9ȡסַ
'long __stdcall ss_id_query_address(long icdev, char* pszText);
Public Declare Function ss_id_query_address Lib "SS728M01.dll" _
    (ByVal icdev As Long, ByVal pszText As String) As Long

'2.3.10ȡ��ݺ���
'long __stdcall ss_id_query_number(long icdev, char* pszText);
Public Declare Function ss_id_query_number Lib "SS728M01.dll" _
    (ByVal icdev As Long, ByVal pszText As String) As Long

'2.3.11ȡǩ������
'long __stdcall ss_id_query_organ(long icdev, char* pszText);
Public Declare Function ss_id_query_organ Lib "SS728M01.dll" _
    (ByVal icdev As Long, ByVal pszText As String) As Long

'2.3.12ȡ��Ч������ʼ����
'long __stdcall ss_id_query_beginterm(long icdev, char* pszText);
Public Declare Function ss_id_query_beginterm Lib "SS728M01.dll" _
    (ByVal icdev As Long, ByVal pszText As String) As Long

'2.3.13ȡ��Ч���޽�ֹ����
'long __stdcall ss_id_query_endterm(long icdev, char* pszText);
Public Declare Function ss_id_query_endterm Lib "SS728M01.dll" _
    (ByVal icdev As Long, ByVal pszText As String) As Long

'2.3.14ȡ��Ƭ
'long __stdcall ss_id_query_photofile(long icdev, unsigned char _Format, char* szImagePath);
Public Declare Function ss_id_query_photofile Lib "SS728M01.dll" _
    (ByVal icdev As Long, ByVal typeFile As Long, ByVal szImagePath As String) As Long

'2.3.15 ȡ����סַ
'long __stdcall ss_id_query_newaddr(long icdev, char* pszText);
Public Declare Function ss_id_query_newaddr Lib "SS728M01.dll" _
    (ByVal icdev As Long, ByVal pszText As String) As Long

'2.3.16 ��տ���Ϣ
'long __stdcall ss_id_free_card(long icdev);
Public Declare Function ss_id_free_card Lib "SS728M01.dll" _
    (ByVal icdev As Long) As Long

Public Property Let ReadCardSS728M01(ByVal vNewValue As Boolean)
    mblnReadCard = vNewValue
End Property

Public Function ReadSS728M01() As Boolean
    Dim bytReadType     As Byte
    Dim strPort         As String
    Dim strServerIP     As String
    Dim lngReturn       As Long
    Dim pszVerInfo      As String
    Dim szDevCom        As String
    Dim icdev           As Long
    Dim szDevIP         As String
    Dim Amount          As Long
    Dim Msec            As Long
    Dim pszParam        As String
    Dim strSFZH         As String
    Dim pszText         As String
On Error GoTo ErrH
    bytReadType = IIf(UCase(GetSetting("ZLSOFT", "����ȫ��\IDCard\SS728M01", "LinkType", "Local")) = "NET", 2, 1)
    strPort = UCase(GetSetting("ZLSOFT", "����ȫ��\IDCard\SS728M01", "Port", "AUTO"))
    strServerIP = GetSetting("ZLSOFT", "����ȫ��\IDCard\SS728M01", "NetIp", "192.168.31.169")
    
    '���豸
    If bytReadType = 1 Then
        '2.2.2�򿪱����ն�
        'long __stdcall ss_dev_open(char* szDevCom);
        szDevCom = "AUTO"
        lngReturn = ss_dev_open(szDevCom)
        If lngReturn < 0 Then
            If ss_error(lngReturn, False) Then
                Exit Function
            End If
        End If
    ElseIf bytReadType = 2 Then
        '2.2.4ע�������ն�
        szDevIP = strServerIP
        lngReturn = ss_dev_login(szDevIP)
        If lngReturn < 0 Then
            If ss_error(lngReturn, False) Then
                Exit Function
            End If
        End If
    End If
    
    icdev = lngReturn
    ''2.2.7������
    Amount = 1
    Msec = 2
    lngReturn = ss_dev_beep(icdev, Amount, Msec)
    If ss_error(lngReturn, False) Then
        Exit Function
    End If
        
    'Ѱ��
    mblnReadCard = True
    
    If mblnReadCard Then
        'һֱ����
        Do While True
            DoEvents
            If Not mblnReadCard Then
                '2.2.8ָʾ��
                'long __stdcall ss_dev_led(long icdev, unsigned char _Color, unsigned char _Status, unsigned short _Amount, unsigned short _Msec);
                'icdev [in]���豸��ʶ���� _Color [in]��ָʾ����ɫ������00H����ɫ����01H����ɫ����02H����ɫ���� _Status [in]��ָʾ��״̬������01H��������02H���𣩡�03H����˸���� _Amount [in]����˸������ _Msec [in]����˸ʱ������λ100���롣
                Amount = 1
                Msec = 2
                lngReturn = ss_dev_led(icdev, ��ɫ, ��, Amount, Msec)
                If ss_error(lngReturn, False) Then
                    Exit Function
                End If
                If bytReadType = 1 Then
                    '2.2.3�رձ����ն�
                    'long __stdcall ss_dev_close(long icdev);
                    lngReturn = ss_dev_close(icdev)
                    Call ss_error(lngReturn, False)
                    Exit Function
                Else
                    '2.2.5ע�������ն�
                    lngReturn = ss_dev_logout(icdev)
                    Call ss_error(lngReturn, False)
                    Exit Function
                End If
            End If
            '2.2.8ָʾ��
            'long __stdcall ss_dev_led(long icdev, unsigned char _Color, unsigned char _Status, unsigned short _Amount, unsigned short _Msec);
            'icdev [in]���豸��ʶ���� _Color [in]��ָʾ����ɫ������00H����ɫ����01H����ɫ����02H����ɫ���� _Status [in]��ָʾ��״̬������01H��������02H���𣩡�03H����˸���� _Amount [in]����˸������ _Msec [in]����˸ʱ������λ100���롣
            Amount = 1
            Msec = 1
            lngReturn = ss_dev_led(icdev, ��ɫ, ��˸, Amount, Msec)
            If ss_error(lngReturn, False) Then
                Exit Function
            End If
            '2.3.1Ѱ��
            'long __stdcall ss_id_find_card(long icdev);
            lngReturn = ss_id_find_card(icdev)
            If lngReturn = 0 Then
                'Ѱ���ɹ�
                Exit Do
            Else
                'Ѱ��ʧ�ܣ�����������ѭ��Ѱ��
                '2.2.8ָʾ��
                'long __stdcall ss_dev_led(long icdev, unsigned char _Color, unsigned char _Status, unsigned short _Amount, unsigned short _Msec);
                'icdev [in]���豸��ʶ���� _Color [in]��ָʾ����ɫ������00H����ɫ����01H����ɫ����02H����ɫ���� _Status [in]��ָʾ��״̬������01H��������02H���𣩡�03H����˸���� _Amount [in]����˸������ _Msec [in]����˸ʱ������λ100���롣
                Amount = 1
                Msec = 2
                lngReturn = ss_dev_led(icdev, ��ɫ, ��, Amount, Msec)
                If ss_error(lngReturn, False) Then
                    Exit Function
                End If
            End If
        Loop
    Else
        'ֻ��һ��
        '2.2.8ָʾ��
        'long __stdcall ss_dev_led(long icdev, unsigned char _Color, unsigned char _Status, unsigned short _Amount, unsigned short _Msec);
        'icdev [in]���豸��ʶ���� _Color [in]��ָʾ����ɫ������00H����ɫ����01H����ɫ����02H����ɫ���� _Status [in]��ָʾ��״̬������01H��������02H���𣩡�03H����˸���� _Amount [in]����˸������ _Msec [in]����˸ʱ������λ100���롣
        Amount = 1
        Msec = 1
        lngReturn = ss_dev_led(icdev, ��ɫ, ��˸, Amount, Msec)
        If ss_error(lngReturn, False) Then
            Exit Function
        End If
        '2.3.1Ѱ��
        'long __stdcall ss_id_find_card(long icdev);
        lngReturn = ss_id_find_card(icdev)
        If lngReturn = 0 Then
            'Ѱ���ɹ�
        Else
            '2.2.8ָʾ��
            'long __stdcall ss_dev_led(long icdev, unsigned char _Color, unsigned char _Status, unsigned short _Amount, unsigned short _Msec);
            'icdev [in]���豸��ʶ���� _Color [in]��ָʾ����ɫ������00H����ɫ����01H����ɫ����02H����ɫ���� _Status [in]��ָʾ��״̬������01H��������02H���𣩡�03H����˸���� _Amount [in]����˸������ _Msec [in]����˸ʱ������λ100���롣
            Amount = 1
            Msec = 2
            lngReturn = ss_dev_led(icdev, ��ɫ, ��, Amount, Msec)
            If ss_error(lngReturn, False) Then
                Exit Function
            End If
            'Ѱ��ʧ��
            If bytReadType = 1 Then
                '2.2.3�رձ����ն�
                'long __stdcall ss_dev_close(long icdev);
                lngReturn = ss_dev_close(icdev)
                Call ss_error(lngReturn, False)
                Exit Function
            Else
                '2.2.5ע�������ն�
                lngReturn = ss_dev_logout(icdev)
                Call ss_error(lngReturn, False)
                Exit Function
            End If
        End If
    End If
    
    '2.3.2��ȡ����Ϣ
    'long __stdcall ss_id_read_card(long icdev);
    lngReturn = ss_id_read_card(icdev)
    DoEvents
    If ss_error(lngReturn) Then
        If bytReadType = 1 Then
            '2.2.3�رձ����ն�
            'long __stdcall ss_dev_close(long icdev);
            lngReturn = ss_dev_close(icdev)
            Call ss_error(lngReturn, False)
            Exit Function
        Else
            '2.2.5ע�������ն�
            lngReturn = ss_dev_logout(icdev)
            Call ss_error(lngReturn, False)
            Exit Function
        End If
    End If
    '2.2.7������[�����ɹ�]
    Amount = 2
    Msec = 1
    lngReturn = ss_dev_beep(icdev, Amount, Msec)
    If ss_error(lngReturn, False) Then
        Exit Function
    End If
    
    '2.3.3ȡ����
    pszText = Space(30)
    lngReturn = ss_id_query_name(icdev, pszText)
    If ss_error(lngReturn, False) Then
        Exit Function
    End If
    SS728M01.ss_id_query_name = TruncZero(pszText)
    
    '2.3.4ȡ�Ա����
    pszText = Space(1)
    lngReturn = ss_id_query_sexid(icdev, pszText)
    If ss_error(lngReturn, False) Then
        Exit Function
    End If
    SS728M01.ss_id_query_sexid = TruncZero(pszText)
     
    '2.3.5ȡ�Ա�����
    pszText = Space(2)
    lngReturn = ss_id_query_sex(icdev, pszText)
    If ss_error(lngReturn, False) Then
        Exit Function
    End If
    SS728M01.ss_id_query_sex = TruncZero(pszText)
    '2.3.6ȡ�������
    pszText = Space(2)
    lngReturn = ss_id_query_folkid(icdev, pszText)
    If ss_error(lngReturn, False) Then
        Exit Function
    End If
    SS728M01.ss_id_query_folkid = TruncZero(pszText)
    '2.3.7ȡ��������
    pszText = Space(10)
    lngReturn = ss_id_query_folk(icdev, pszText)
    If ss_error(lngReturn, False) Then
        Exit Function
    End If
    SS728M01.ss_id_query_folk = TruncZero(pszText)
    '2.3.8ȡ��������
    pszText = Space(8)
    lngReturn = ss_id_query_birth(icdev, pszText)
    If ss_error(lngReturn, False) Then
        Exit Function
    End If
    SS728M01.ss_id_query_birth = TruncZero(pszText)
    '2.3.9ȡסַ
    pszText = Space(70)
    lngReturn = ss_id_query_address(icdev, pszText)
    If ss_error(lngReturn, False) Then
        Exit Function
    End If
    SS728M01.ss_id_query_address = TruncZero(pszText)
    '2.3.10ȡ��ݺ���
    pszText = Space(18)
    lngReturn = ss_id_query_number(icdev, pszText)
    If ss_error(lngReturn, False) Then
        Exit Function
    End If
    SS728M01.ss_id_query_number = TruncZero(pszText)
    '2.3.11ȡǩ������
    pszText = Space(30)
    lngReturn = ss_id_query_organ(icdev, pszText)
    If ss_error(lngReturn, False) Then
        Exit Function
    End If
    SS728M01.ss_id_query_organ = TruncZero(pszText)
    '2.3.12ȡ��Ч������ʼ����
    pszText = Space(8)
    lngReturn = ss_id_query_beginterm(icdev, pszText)
    If ss_error(lngReturn, False) Then
        Exit Function
    End If
    SS728M01.ss_id_query_beginterm = TruncZero(pszText)
    '2.3.13ȡ��Ч���޽�ֹ����
    pszText = Space(8)
    lngReturn = ss_id_query_endterm(icdev, pszText)
    If ss_error(lngReturn, False) Then
        Exit Function
    End If
    SS728M01.ss_id_query_endterm = TruncZero(pszText)
    'ȡ��Ƭ
    '2.3.14ȡ��Ƭ
    lngReturn = ss_id_query_photofile(icdev, 2, "c:\" & SS728M01.ss_id_query_number)
    If ss_error(lngReturn, False) Then
        Exit Function
    End If
    SS728M01.ss_id_query_photofile = "c:\" & SS728M01.ss_id_query_number & ".bmp"
    
    '2.3.15 ȡ����סַ
    pszText = Space(70)
    lngReturn = ss_id_query_newaddr(icdev, pszText)
    If ss_error(lngReturn, False) Then
        Exit Function
    End If
    SS728M01.ss_id_query_newaddr = TruncZero(pszText)
    
    '2.3.16 ��տ���Ϣ
    'long __stdcall ss_id_free_card(long icdev);
    lngReturn = ss_id_free_card(icdev)
    If ss_error(lngReturn, False) Then
        Exit Function
    End If
        
    '2.2.8ָʾ��
    'long __stdcall ss_dev_led(long icdev, unsigned char _Color, unsigned char _Status, unsigned short _Amount, unsigned short _Msec);
    'icdev [in]���豸��ʶ���� _Color [in]��ָʾ����ɫ������00H����ɫ����01H����ɫ����02H����ɫ���� _Status [in]��ָʾ��״̬������01H��������02H���𣩡�03H����˸���� _Amount [in]����˸������ _Msec [in]����˸ʱ������λ100���롣
    Amount = 1
    Msec = 2
    lngReturn = ss_dev_led(icdev, ��ɫ, ��, Amount, Msec)
    If ss_error(lngReturn, False) Then
        Exit Function
    End If
    If bytReadType = 1 Then
        '2.2.3�رձ����ն�
        'long __stdcall ss_dev_close(long icdev);
        lngReturn = ss_dev_close(icdev)
        If ss_error(lngReturn, False) Then
            Exit Function
        End If
    Else
        '2.2.5ע�������ն�
        lngReturn = ss_dev_logout(icdev)
        If ss_error(lngReturn, False) Then
            Exit Function
        End If
    End If
    ReadSS728M01 = True
    Exit Function
ErrH:
    Err.Clear
    Exit Function
End Function

Public Function ClosSS728M01() As Boolean
    Dim bytReadType     As Byte
    Dim strPort         As String
    Dim strServerIP     As String
    Dim lngReturn       As Long
    Dim pszVerInfo      As String
    Dim szDevCom        As String
    Dim icdev           As Long
    Dim szDevIP         As String
    Dim Amount          As Long
    Dim Msec            As Long
    Dim pszParam        As String
    Dim strSFZH         As String
    Dim pszText         As String
On Error GoTo ErrH
    bytReadType = IIf(UCase(GetSetting("ZLSOFT", "����ȫ��\IDCard\SS728M01", "LinkType", "Local")) = "NET", 2, 1)
    strPort = UCase(GetSetting("ZLSOFT", "����ȫ��\IDCard\SS728M01", "Port", "AUTO"))
    strServerIP = GetSetting("ZLSOFT", "����ȫ��\IDCard\SS728M01", "Port", "192.168.31.169")
    '���豸
    If bytReadType = 1 Then
        '2.2.2�򿪱����ն�
        'long __stdcall ss_dev_open(char* szDevCom);
        szDevCom = "AUTO"
        lngReturn = ss_dev_open(szDevCom)
        If lngReturn < 0 Then
            If ss_error(lngReturn, False) Then
                Exit Function
            End If
        End If
    ElseIf bytReadType = 2 Then
        '2.2.4ע�������ն�
        szDevIP = strServerIP
        lngReturn = ss_dev_login(szDevIP)
        If lngReturn < 0 Then
            If ss_error(lngReturn, False) Then
                Exit Function
            End If
        End If
    End If
    '2.2.8ָʾ��
    'long __stdcall ss_dev_led(long icdev, unsigned char _Color, unsigned char _Status, unsigned short _Amount, unsigned short _Msec);
    'icdev [in]���豸��ʶ���� _Color [in]��ָʾ����ɫ������00H����ɫ����01H����ɫ����02H����ɫ���� _Status [in]��ָʾ��״̬������01H��������02H���𣩡�03H����˸���� _Amount [in]����˸������ _Msec [in]����˸ʱ������λ100���롣
    Amount = 1
    Msec = 2
    lngReturn = ss_dev_led(icdev, ��ɫ, ��, Amount, Msec)
    If ss_error(lngReturn, False) Then
        Exit Function
    End If
    If bytReadType = 1 Then
        '2.2.3�رձ����ն�
        'long __stdcall ss_dev_close(long icdev);
        lngReturn = ss_dev_close(icdev)
        If ss_error(lngReturn, False) Then
            Exit Function
        End If
    Else
        '2.2.5ע�������ն�
        lngReturn = ss_dev_logout(icdev)
        If ss_error(lngReturn, False) Then
            Exit Function
        End If
    End If
    Exit Function
ErrH:
    Err.Clear
    Exit Function
End Function

'���ݴ�����жϸ�����ʾ��
Public Function ss_error(lngErrNum As Long, Optional blnShowErr As Boolean = True) As Boolean
    Dim strErrMsg           As String
On Error GoTo ErrH
    Select Case lngErrNum
        Case 0 '�����ɹ�
            ss_error = False
            Exit Function
        Case -1 '���������Ӵ�
            strErrMsg = "���������Ӵ�"
        Case -2 'δ��������
            strErrMsg = "δ��������"
        Case -3 '���������
            strErrMsg = "���������"
        Case -4 'ͨѶ����
            strErrMsg = "ͨѶ����"
        Case -11 '�޿�
            strErrMsg = "�޿�"
        Case -12 '��Ƭ���Ͳ���
            strErrMsg = "��Ƭ���Ͳ���"
        Case -13 '�п�δ�ϵ�
            strErrMsg = "�п�δ�ϵ�"
        Case -14 '��Ƭ��Ӧ��
            strErrMsg = "��Ƭ��Ӧ��"
        Case -21 '�����ⲿdllʧ��
            strErrMsg = "�����ⲿdllʧ��"
        Case -31 '��Ƭ����ʧ��
            strErrMsg = "��Ƭ����ʧ��"
        Case -100 'ִ��ʧ�ܻ�δ֪����
            strErrMsg = "ִ��ʧ�ܻ�δ֪����"
        Case Else
            strErrMsg = CStr(lngErrNum) & "δ֪�Ĵ�����Ϣ"
    End Select
    If blnShowErr Then MsgBox strErrMsg, vbExclamation, "��˼����������ʾ"
    ss_error = True
    Exit Function
ErrH:
    MsgBox Err.Description, vbCritical, "ϵͳ��Ϣ"
    Err.Clear
    ss_error = True
End Function
 
