Attribute VB_Name = "mdlPACSWork"
Option Explicit
Public SplashObj As New frmSplash
Public gstrStation As String                '������վ����
Public gstrSystems As String
Public gstr��λ���� As String
'-----------------------------------------
'�����롢ע���롢�������������ע���������
Public gstrRegCode As String
Public gstrPublish As String
Public gstrParseRegCode As String
Public gstrParsePublish As String
'-----------------------------------------



Public gcnOracle As New ADODB.Connection        '�������ݿ����ӣ��ر�ע�⣺��������Ϊ�µ�ʵ��
Public gstrPrivs As String                   '��ǰ�û����еĵ�ǰģ��Ĺ���
Public gstrSysName As String                'ϵͳ����
Public gstrVersion As String                'ϵͳ�汾
Public gstrAviPath As String                'AVI�ļ��Ĵ��Ŀ¼
Public glngModul As Long
Public glngSys As Long
Public gstrIme As String                    '�Ƿ��Զ��������뷨
Public gstrDbUser As String                 '��ǰ���ݿ��û�
Public glngUserId As Long                   '��ǰ�û�id
Public gstrUserCode As String               '��ǰ�û�����
Public gstrUserName As String               '��ǰ�û�����
Public gstrUserAbbr As String               '��ǰ�û�����

Public glngDeptId As Long                   '��ǰ�û�����id
Public gstrDeptCode As String               '��ǰ�û����ű���
Public gstrDeptName As String               '��ǰ�û���������


Public gstrUnitName As String '�û���λ����
Public gfrmMain As Object

Public gstrSQL As String
Public glngTXTProc As Long
Public gbln�Ӱ�Ӽ� As Boolean
Public grsDuty As ADODB.Recordset '���ҽ��ְ��
Public grsSysPars As ADODB.Recordset
Public gbytCardNOLen As Byte

'ϵͳ����
Public gbytDec As Byte '���ý���С����λ��
Public gstrDec As String '��С��λ������ĸ�ʽ����,��"0.0000"

Public gobjKernel As New zlCISKernel.clsCISKernel 'ҽ������
Public gobjRichEPR As New zlRichEPR.cRichEPR

Public gbytCardLen As Byte '���￨�ų���
Public gblnCardHide As Boolean '���￨��������ʾ
Public gstrCardMask As String  '���￨�������ĸǰ׺:AA|BB|CC...
Public gint�Һ����� As Integer '�Һŵ���Ч����

'�б���ɫ����
Public gdblColor�ѵǼ� As Double
Public gdblColor�ѱ��� As Double
Public gdblColor�Ѽ�� As Double
Public gdblColor�ѱ��� As Double
Public gdblColor����� As Double
Public gdblColor����� As Double
Public gdblColor������ As Double
Public gdblColor������ As Double
Public gdblColor����� As Double
Public gdblColor�Ѿܾ� As Double


Public gConnectedShardDir() As String   '�Ѿ����ӹ��Ĺ���Ŀ¼���豸������

'---------------------------�豸�������ƣ�ע��-------------------------------
Public Const LOGIN_TYPE_��Ƶ�豸 As String = "Ӱ����Ƶ�豸����"
Public Const LOGIN_TYPE_��Ƭ��ӡ�� As String = "Ӱ��Ƭ��ӡ������"
Public Const LOGIN_TYPE_DICOM�豸 As String = "Ӱ��DICOM�豸����"
Public gint��Ƶ�豸���� As Integer
Public gint��Ƭ��ӡ������ As Integer
Public gintDICOM�豸���� As Integer


'---------------------------------------------------------------
'- ע��� Api ����...
'---------------------------------------------------------------
' Reg Data Types...
Const REG_SZ = 1                         ' Unicode���ս��ַ���
Const REG_EXPAND_SZ = 2                  ' Unicode���ս��ַ���
Const REG_DWORD = 4                      ' 32-bit ����

' ע���������ֵ...
Const REG_OPTION_NON_VOLATILE = 0       ' ��ϵͳ��������ʱ���ؼ��ֱ�����

' ע���ؼ��ְ�ȫѡ��...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_READ = KEY_QUERY_VALUE + KEY_ENUMERATE_SUB_KEYS + KEY_NOTIFY + READ_CONTROL
Const KEY_WRITE = KEY_SET_VALUE + KEY_CREATE_SUB_KEY + READ_CONTROL
Const KEY_EXECUTE = KEY_READ
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' ע���ؼ��ָ�����...
Public Const HKEY_LOCAL_MACHINE = &H80000002
Const HKEY_CLASSES_ROOT = &H80000000
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_USERS = &H80000003
Const HKEY_PERFORMANCE_DATA = &H80000004

' ����ֵ...
Const ERROR_NONE = 0
Const ERROR_BADKEY = 2
Const ERROR_ACCESS_DENIED = 8
Const ERROR_SUCCESS = 0

'---------------------------------------------------------------
'- ע���ȫ��������...
'---------------------------------------------------------------
Public Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Boolean
End Type




Public mrsDeptParas As ADODB.Recordset '���Ʋ�������
'-----------------------------------------------------------
Public Type TYPE_USER_INFO
    ID As Long
    ����ID As Long
    ��� As String
    ���� As String
    ���� As String
    �û��� As String
End Type
Public UserInfo As TYPE_USER_INFO

'��ȡ�����Ķ��IP
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

Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long
Private Declare Function WSAGetLastError Lib "WSOCK32.DLL" () As Long
Private Declare Function WSAStartup Lib "WSOCK32.DLL" (ByVal wVersionRequired As Integer, lpWSAData As WSADATA) As Long
Private Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long

Private Declare Function gethostname Lib "WSOCK32.DLL" (ByVal hostname$, ByVal HostLen As Long) As Long
Private Declare Function gethostbyname Lib "WSOCK32.DLL" (ByVal hostname$) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (hpvDest As Any, ByVal hpvSource&, ByVal cbCopy&)

'---------------------------------------------------------------
'-ע��� API ����...
'---------------------------------------------------------------
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByRef lpSecurityAttributes As SECURITY_ATTRIBUTES, ByRef phkResult As Long, ByRef lpdwDisposition As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long

Public Declare Function SetActiveWindow Lib "user32" (ByVal Hwnd As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal Hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Const GWL_EXSTYLE = (-20)
Public Const WinStyle = &H40000
Public Const SWP_NOSIZE = &H1
Public Const SWP_SHOWWINDOW = &H40
Public Const HWND_TOPMOST = -1


Public Sub Main()
    Dim lngReturn As Long
    Dim StrUnitName As String
    Dim BlnShowFlash As Boolean
    Dim strCode As String, IntCount As Integer, StrStyle As String
    Dim rsMenu As ADODB.Recordset, StrHaveSys As String
    
    
    If App.PrevInstance Then
        MsgBox "Ӱ����շ����Ѿ������������ٴ����С�", vbInformation, "����"
        Exit Sub
    End If
    
    
    'Ϊʵ��XP�������ʾ����ǰ����ִ�иú���
    Call InitCommonControls

    
    BlnShowFlash = False
    Load SplashObj
    '��ע����л�ȡ�û�ע�������Ϣ,����û���λ���Ʋ�Ϊ��,����ʾ���ִ���
    StrUnitName = GetSetting("ZLSOFT", "ע����Ϣ", "��λ����", "")
    gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "��ʾ", "")
    If StrUnitName <> "" Then
        With SplashObj
            '��������Ҫ����
            Call ApplyOEM_Picture(.ImgIndicate, "Picture")
            Call ApplyOEM_Picture(.imgPic, "PictureB")
            .Show
            .lblGrant = StrUnitName
            StrUnitName = GetSetting("ZLSOFT", "ע����Ϣ", "������", "")
            If Trim(StrUnitName) = "" Then
                .Label3.Visible = False
                .lbl������.Visible = False
            Else
                .lbl������.Caption = ""
                For IntCount = 0 To UBound(Split(StrUnitName, ";"))
                    .lbl������.Caption = .lbl������.Caption & Split(StrUnitName, ";")(IntCount) & vbCrLf
                Next
            End If
            .LblProductName = GetSetting("ZLSOFT", "ע����Ϣ", "��Ʒȫ��", "")
            .lbl����֧���� = GetSetting("ZLSOFT", "ע����Ϣ", "����֧����", "")
        End With
        
        BlnShowFlash = True
        DoEvents
    End If
    
    gstrStation = Space(200)
    lngReturn = GetComputerName(gstrStation, 200)
    gstrStation = Trim(gstrStation)
    If Len(gstrStation) > 1 Then
        gstrStation = Left(gstrStation, Len(gstrStation) - 1)
    Else
        gstrStation = "..."
    End If
    
    '�û�ע��
    frmUserLogin.Show 1
    If gcnOracle.State <> adStateOpen Then
        Unload frmUserLogin
        Unload SplashObj
        
        Set gcnOracle = Nothing
        Exit Sub
    End If
    
    '��ʼ����������
    InitCommon gcnOracle
    If RegCheck = False Then
        Unload SplashObj
        
        Set gcnOracle = Nothing
        Exit Sub
    End If
    
    '�����������Ч��Ϊ�ջ�Ϊ"-"�������˳�
    gstrParsePublish = zlRegInfo("��Ʒ����")
    gstrParseRegCode = zlRegInfo("��λ����", , -1)
    
    gstrSysName = gstrParsePublish & "���"
    SaveSetting "ZLSOFT", "ע����Ϣ", "��ʾ", gstrSysName
    SaveSetting "ZLSOFT", "ע����Ϣ", UCase("gstrSysName"), gstrSysName
    gstrVersion = App.Major & "." & App.Minor & "." & App.Revision
    SaveSetting "ZLSOFT", "ע����Ϣ", UCase("gstrVersion"), gstrVersion
    gstrAviPath = App.Path & "\�����ļ�"
    SaveSetting "ZLSOFT", "ע����Ϣ", UCase("gstrAviPath"), gstrAviPath
    
    With SplashObj
        If BlnShowFlash = False Then
            .lblGrant = gstrParseRegCode
            .lbl����֧����.Caption = zlRegInfo("����֧����", , -1)
            .LblProductName = zlRegInfo("��Ʒ����")
            
            strCode = zlRegInfo("��Ʒ������", , -1)
            .lbl������.Caption = ""
            For IntCount = 0 To UBound(Split(strCode, ";"))
                .lbl������.Caption = .lbl������.Caption & Split(strCode, ";")(IntCount) & vbCrLf
            Next
            Call ApplyOEM_Picture(.ImgIndicate, "Picture")
            .Show
            BlnShowFlash = True
        End If
        DoEvents
    End With
    
    '���û�ע�������Ϣд��ע���,���´�����ʱ��ʾ
    SaveSetting "ZLSOFT", "ע����Ϣ", "��λ����", gstrParseRegCode
    SaveSetting "ZLSOFT", "ע����Ϣ", "��Ʒȫ��", zlRegInfo("��Ʒ����")
    SaveSetting "ZLSOFT", "ע����Ϣ", "��Ʒ����", zlRegInfo("��Ʒ����")
    SaveSetting "ZLSOFT", "ע����Ϣ", "����֧����", zlRegInfo("����֧����", , -1)
    SaveSetting "ZLSOFT", "ע����Ϣ", "������", zlRegInfo("��Ʒ������", , -1)
    SaveSetting "ZLSOFT", "ע����Ϣ", "WEB֧���̼���", zlRegInfo("֧���̼���")
    SaveSetting "ZLSOFT", "ע����Ϣ", "WEB֧��EMAIL", zlRegInfo("֧����MAIL")
    SaveSetting "ZLSOFT", "ע����Ϣ", "WEB֧��URL", zlRegInfo("֧����URL")

    gstrSystems = " (ϵͳ =100 Or ϵͳ Is NULL)"
    glngSys = 100
    
    '-------------------------------------------------------------
    '����ͬ���
    '-------------------------------------------------------------
    zlDatabase.ExecuteProcedure "Zl_Createsynonyms(" & glngSys & ")", "����ͬ���"
    

    Unload SplashObj
    
    CodeMan 1290
End Sub


Public Sub CodeMan(ByVal lngModul As Long)
    '------------------------------------------------
    '���ܣ� ����������ָ�����ܣ�����ִ����س���
    '������
    '   lngModul:��Ҫִ�еĹ������
    '���أ�
    '------------------------------------------------
    Dim rsUser As ADODB.Recordset
    
    gstrAviPath = GetSetting(appName:="ZLSOFT", Section:="ע����Ϣ", Key:=UCase("gstrAviPath"), Default:="")
    gstrSysName = GetSetting(appName:="ZLSOFT", Section:="ע����Ϣ", Key:=UCase("gstrSysName"), Default:="")
    gstrVersion = GetSetting(appName:="ZLSOFT", Section:="ע����Ϣ", Key:=UCase("gstrVersion"), Default:="")
    gstr��λ���� = GetUnitName()
    
    '��ȡ�û�����Ϣ
    Set rsUser = zlDatabase.GetUserInfo
    If rsUser.RecordCount <> 0 Then
        glngUserId = Nvl(rsUser!ID)
        gstrUserCode = Nvl(rsUser!���)
        gstrUserName = Nvl(rsUser!����)
        gstrUserAbbr = Nvl(rsUser!����)
        glngDeptId = Nvl(rsUser!����ID)
        gstrDeptCode = Nvl(rsUser!������)
        gstrDeptName = Nvl(rsUser!������)
    Else
        glngUserId = 0
        gstrUserCode = ""
        gstrUserName = ""
        gstrUserAbbr = ""
        glngDeptId = 0
        gstrDeptCode = ""
        gstrDeptName = ""
    End If
    
    gstrPrivs = GetPrivFunc(glngSys, lngModul)
    '-------------------------------------------------
    
    Select Case lngModul
        Case 1290
            frmBrowserStation.Show
    End Select
End Sub


Public Function TranPasswd(strOld As String) As String
    '------------------------------------------------
    '���ܣ� ����ת������
    '������
    '   strOld��ԭ����
    '���أ� �������ɵ�����
    '------------------------------------------------
    Dim iBit As Integer, strBit As String
    Dim strNew As String
    If Len(Trim(strOld)) = 0 Then TranPasswd = "": Exit Function
    strNew = ""
    For iBit = 1 To Len(Trim(strOld))
        strBit = UCase(Mid(Trim(strOld), iBit, 1))
        Select Case (iBit Mod 3)
        Case 1
            strNew = strNew & _
                Switch(strBit = "0", "W", strBit = "1", "I", strBit = "2", "N", strBit = "3", "T", strBit = "4", "E", strBit = "5", "R", strBit = "6", "P", strBit = "7", "L", strBit = "8", "U", strBit = "9", "M", _
                   strBit = "A", "H", strBit = "B", "T", strBit = "C", "I", strBit = "D", "O", strBit = "E", "K", strBit = "F", "V", strBit = "G", "A", strBit = "H", "N", strBit = "I", "F", strBit = "J", "J", _
                   strBit = "K", "B", strBit = "L", "U", strBit = "M", "Y", strBit = "N", "G", strBit = "O", "P", strBit = "P", "W", strBit = "Q", "R", strBit = "R", "M", strBit = "S", "E", strBit = "T", "S", _
                   strBit = "U", "T", strBit = "V", "Q", strBit = "W", "L", strBit = "X", "Z", strBit = "Y", "C", strBit = "Z", "X", True, strBit)
        Case 2
            strNew = strNew & _
                Switch(strBit = "0", "7", strBit = "1", "M", strBit = "2", "3", strBit = "3", "A", strBit = "4", "N", strBit = "5", "F", strBit = "6", "O", strBit = "7", "4", strBit = "8", "K", strBit = "9", "Y", _
                   strBit = "A", "6", strBit = "B", "J", strBit = "C", "H", strBit = "D", "9", strBit = "E", "G", strBit = "F", "E", strBit = "G", "Q", strBit = "H", "1", strBit = "I", "T", strBit = "J", "C", _
                   strBit = "K", "U", strBit = "L", "P", strBit = "M", "B", strBit = "N", "Z", strBit = "O", "0", strBit = "P", "V", strBit = "Q", "I", strBit = "R", "W", strBit = "S", "X", strBit = "T", "L", _
                   strBit = "U", "5", strBit = "V", "R", strBit = "W", "D", strBit = "X", "2", strBit = "Y", "S", strBit = "Z", "8", True, strBit)
        Case 0
            strNew = strNew & _
                Switch(strBit = "0", "6", strBit = "1", "J", strBit = "2", "H", strBit = "3", "9", strBit = "4", "G", strBit = "5", "E", strBit = "6", "Q", strBit = "7", "1", strBit = "8", "X", strBit = "9", "L", _
                   strBit = "A", "S", strBit = "B", "8", strBit = "C", "5", strBit = "D", "R", strBit = "E", "7", strBit = "F", "M", strBit = "G", "3", strBit = "H", "A", strBit = "I", "N", strBit = "J", "F", _
                   strBit = "K", "O", strBit = "L", "4", strBit = "M", "K", strBit = "N", "Y", strBit = "O", "D", strBit = "P", "2", strBit = "Q", "T", strBit = "R", "C", strBit = "S", "U", strBit = "T", "P", _
                   strBit = "U", "B", strBit = "V", "Z", strBit = "W", "0", strBit = "X", "V", strBit = "Y", "I", strBit = "Z", "W", True, strBit)
        End Select
    Next
    TranPasswd = strNew

End Function


Public Function OraDataOpen(ByVal strServerName As String, ByVal strUserName As String, ByVal strUserPwd As String) As Boolean
    '------------------------------------------------
    '���ܣ� ��ָ�������ݿ�
    '������
    '   strServerName�������ַ���
    '   strUserName���û���
    '   strUserPwd������
    '���أ� ���ݿ�򿪳ɹ�������true��ʧ�ܣ�����false
    '------------------------------------------------
    Dim strSQL As String
    Dim strError As String
    
    On Error Resume Next
    err = 0
    DoEvents
    With gcnOracle
        If .State = adStateOpen Then .Close
        .Provider = "MSDataShape"
        .Open "Driver={Microsoft ODBC for Oracle};Server=" & strServerName, strUserName, strUserPwd
        If err <> 0 Then
            '���������Ϣ
            strError = err.Description
            If InStr(strError, "�Զ�������") > 0 Then
                MsgBox "���Ӵ��޷��������������ݷ��ʲ����Ƿ�������װ��", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-12154") > 0 Then
                MsgBox "�޷���������������" & vbCrLf & "������Oracle�������Ƿ���ڸñ�������������������ַ�������", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-12541") > 0 Then
                MsgBox "�޷����ӣ�����������ϵ�Oracle�����������Ƿ�������", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-01033") > 0 Then
                MsgBox "ORACLE���ڳ�ʼ�����ڹرգ����Ժ����ԡ�", vbInformation, gstrSysName
            Else
                MsgBox "�����û�������������ָ�������޷�ע�ᡣ", vbInformation, gstrSysName
            End If
            
            OraDataOpen = False
            Exit Function
        End If
    End With
    
    err = 0
    On Error GoTo errHand
    
    gstrDbUser = UCase(strUserName)
    SetDbUser gstrDbUser
    OraDataOpen = True
    Exit Function
    
errHand:
    If ErrCenter() = 1 Then Resume
    OraDataOpen = False
    err = 0
End Function

Public Function OraDataClose() As Boolean
    '------------------------------------------------
    '���ܣ� �ر����ݿ�
    '������
    '���أ� �ر����ݿ⣬����True��ʧ�ܣ�����False
    '------------------------------------------------
    err = 0
    On Error Resume Next
    gcnOracle.Close
    OraDataClose = True
    err = 0

End Function


'-------------------------------------------------------------------------------------------------
'sample usage - Debug.Print GetKeyValue(HKEY_CLASSES_ROOT, "COMCTL.ListviewCtrl.1\CLSID", "")
'-------------------------------------------------------------------------------------------------
Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String) As String
'���ܣ���ע���
    Dim i As Long                                           ' ѭ��������
    Dim rc As Long                                          ' ���ش���
    Dim hKey As Long                                        ' ����򿪵�ע���ؼ���
    Dim hDepth As Long                                      '
    Dim sKeyVal As String
    Dim lKeyValType As Long                                 ' ע���ؼ�����������
    Dim tmpVal As String                                    ' ע���ؼ��ֵ���ʱ�洢��
    Dim KeyValSize As Long                                  ' ע���ؼ��ֱ����ߴ�
    
    ' �� KeyRoot {HKEY_LOCAL_MACHINE...} �´�ע���ؼ���
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' ��ע���ؼ���
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' �������...
    
    tmpVal = String$(1024, 0)                             ' ��������ռ�
    KeyValSize = 1024                                       ' ��Ǳ����ߴ�
    
    '------------------------------------------------------------
    ' ����ע���ؼ��ֵ�ֵ...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         lKeyValType, tmpVal, KeyValSize)    ' ���/�����ؼ��ֵ�ֵ
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' ������
      
    tmpVal = Left$(tmpVal, InStr(tmpVal, Chr(0)) - 1)

    '------------------------------------------------------------
    ' �����ؼ���ֵ��ת������...
    '------------------------------------------------------------
    Select Case lKeyValType                                  ' ������������...
    Case REG_SZ, REG_EXPAND_SZ                              ' �ַ���ע���ؼ�����������
        sKeyVal = tmpVal                                     ' �����ַ�����ֵ
    Case REG_DWORD                                          ' ���ֽ�ע���ؼ�����������
        For i = Len(tmpVal) To 1 Step -1                    ' ת��ÿһλ
            sKeyVal = sKeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' һ���ַ�һ���ַ�������ֵ��
        Next
        sKeyVal = Format$("&h" + sKeyVal)                     ' ת�����ֽ�Ϊ�ַ���
    End Select
    
    GetKeyValue = sKeyVal                                   ' ����ֵ
    rc = RegCloseKey(hKey)                                  ' �ر�ע���ؼ���
    Exit Function                                           ' �˳�
    
GetKeyError:    ' ����������������...
    GetKeyValue = vbNullString                              ' ���÷���ֵΪ���ַ���
    rc = RegCloseKey(hKey)                                  ' �ر�ע���ؼ���
End Function


Public Function UpdatePassword(ByVal strUserName As String, ByVal strPasswd As String) As Boolean
    '-------------------------------------------------------------
    '���ܣ�����ԱID���޸�������
    '������CurrUser
    '      ��ǰ�û���
    '���أ�����ɹ����˻�True�����򷵻�False
    '-------------------------------------------------------------
    err = 0
    On Error GoTo ErrorHand
    
    DoEvents
    gcnOracle.Execute "alter user " & strUserName & " identified by " & strPasswd
    UpdatePassword = True
    Exit Function
    
ErrorHand:
    If ErrCenter() = 1 Then Resume
    UpdatePassword = False

End Function

Public Function UpdateKey(KeyRoot As Long, KeyName As String, SubKeyName As String, SubKeyValue As String) As Boolean
'���ܣ�дע���
    Dim rc As Long                                      ' ���ش���
    Dim hKey As Long                                    ' ����һ��ע���ؼ���
    Dim hDepth As Long                                  '
    Dim lpAttr As SECURITY_ATTRIBUTES                   ' ע���ȫ����
    
    lpAttr.nLength = 50                                 ' ���ð�ȫ����Ϊȱʡֵ...
    lpAttr.lpSecurityDescriptor = 0                     ' ...
    lpAttr.bInheritHandle = True                        ' ...

    '------------------------------------------------------------
    '- ����/��ע���ؼ���...
    '------------------------------------------------------------
    rc = RegCreateKeyEx(KeyRoot, KeyName, _
                        0, REG_SZ, _
                        REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, lpAttr, _
                        hKey, hDepth)                   ' ����/��//KeyRoot//KeyName
    
    If (rc <> ERROR_SUCCESS) Then GoTo CreateKeyError   ' ������...
    
    '------------------------------------------------------------
    '- ����/�޸Ĺؼ���ֵ...
    '------------------------------------------------------------
    If (SubKeyValue = "") Then SubKeyValue = " "        ' Ҫ��RegSetValueEx() ������Ҫ����һ���ո�...
    
    ' ����/�޸Ĺؼ���ֵ
    rc = RegSetValueEx(hKey, SubKeyName, _
                       0, REG_SZ, _
                       SubKeyValue, LenB(StrConv(SubKeyValue, vbFromUnicode)))
                       
    If (rc <> ERROR_SUCCESS) Then GoTo CreateKeyError   ' ������
    '------------------------------------------------------------
    '- �ر�ע���ؼ���...
    '------------------------------------------------------------
    rc = RegCloseKey(hKey)                              ' �رչؼ���
    
    UpdateKey = True                                    ' ���سɹ�
    Exit Function                                       ' �˳�
CreateKeyError:
    UpdateKey = False                                   ' ���ô��󷵻ش���
    rc = RegCloseKey(hKey)                              ' ��ͼ�رչؼ���
End Function


Public Sub ReadStudyListColor(ByVal lngDeptID As Long)

  gdblColor������ = GetStudyListColor(lngDeptID, "������")
  If gdblColor������ < 0 Then
    gdblColor������ = ColorConstants.vbWhite
  End If
  
  gdblColor������ = GetStudyListColor(lngDeptID, "������")
  If gdblColor������ < 0 Then
    gdblColor������ = ColorConstants.vbWhite
  End If
  
  gdblColor����� = GetStudyListColor(lngDeptID, "�����")
  If gdblColor����� < 0 Then
    gdblColor����� = ColorConstants.vbWhite
  End If
  
  gdblColor�ѱ��� = GetStudyListColor(lngDeptID, "�ѱ���")
  If gdblColor�ѱ��� < 0 Then
    gdblColor�ѱ��� = ColorConstants.vbWhite
  End If
  
  gdblColor�ѵǼ� = GetStudyListColor(lngDeptID, "�ѵǼ�")
  If gdblColor�ѵǼ� < 0 Then
    gdblColor�ѵǼ� = ColorConstants.vbWhite
  End If
  
  gdblColor�Ѽ�� = GetStudyListColor(lngDeptID, "�Ѽ��")
  If gdblColor�Ѽ�� < 0 Then
    gdblColor�Ѽ�� = ColorConstants.vbWhite
  End If
  
  gdblColor����� = GetStudyListColor(lngDeptID, "�����")
  If gdblColor����� < 0 Then
    gdblColor����� = ColorConstants.vbWhite
  End If
  
  gdblColor����� = GetStudyListColor(lngDeptID, "�����")
  If gdblColor����� < 0 Then
    gdblColor����� = ColorConstants.vbGreen
  End If
  
  gdblColor�ѱ��� = GetStudyListColor(lngDeptID, "�ѱ���")
  If gdblColor�ѱ��� < 0 Then
    gdblColor�ѱ��� = ColorConstants.vbWhite
  End If
  
  gdblColor�Ѿܾ� = GetStudyListColor(lngDeptID, "�Ѿܾ�")
  If gdblColor�Ѿܾ� < 0 Then
    gdblColor�Ѿܾ� = ColorConstants.vbYellow
  End If
End Sub


Public Function GetUserInfo() As Boolean
'���ܣ���ȡ��½�û���Ϣ
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    Set rsTmp = zlDatabase.GetUserInfo
    
    UserInfo.�û��� = gstrDbUser
    UserInfo.���� = gstrDbUser
    If Not rsTmp.EOF Then
        UserInfo.ID = rsTmp!ID
        UserInfo.��� = rsTmp!���
        UserInfo.����ID = IIf(IsNull(rsTmp!����ID), 0, rsTmp!����ID)
        UserInfo.���� = IIf(IsNull(rsTmp!����), "", rsTmp!����)
        UserInfo.���� = IIf(IsNull(rsTmp!����), "", rsTmp!����)
        UserInfo.�û��� = IIf(IsNull(rsTmp!�û���), "", rsTmp!�û���)
        GetUserInfo = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetUser����IDs(Optional ByVal bln���� As Boolean) As String
'���ܣ���ȡ����Ա�����Ŀ���(�������ڿ���+�������������Ŀ���),�����ж��
'�������Ƿ�ȡ���������µĿ���
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    
    strSQL = "Select ����ID From ������Ա Where ��ԱID=[1]"
    If bln���� Then
        strSQL = strSQL & " Union" & _
            " Select Distinct B.����ID From ������Ա A,��λ״����¼ B" & _
            " Where A.����ID=B.����ID And A.��ԱID=[1]"
    End If
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", UserInfo.ID)
    For i = 1 To rsTmp.RecordCount
        GetUser����IDs = GetUser����IDs & "," & rsTmp!����ID
        rsTmp.MoveNext
    Next
    GetUser����IDs = Mid(GetUser����IDs, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

'ȡ�ü���б�ָ����������ɫ
Public Function GetStudyListColor(ByVal lngDeptID As Long, ByVal strParameterName As String) As Double
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim lngTemp As Long
             
    On Error GoTo err
        
    strSQL = "select ID ,����ID,������,����ֵ from Ӱ�����̲��� where ����ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "ȡ�ü���б���ɫ", lngDeptID)
        
    GetStudyListColor = -1
    
    While Not rsTemp.EOF
        If rsTemp!������ = strParameterName Then
          GetStudyListColor = Val(rsTemp!����ֵ)
          Exit Function
        End If
        rsTemp.MoveNext
    Wend
    
    Exit Function
err:
    If ErrCenter() = 1 Then Resume Next
    Call SaveErrLog
End Function

Public Function getID_TO_����(ByVal lngID As Long, ByVal strDict As String) As String
Dim rsTemp As ADODB.Recordset
    On Error GoTo errHandle
    gstrSQL = "select ���� FROM " & strDict & " WHERE ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ͨ��������ȡID", lngID)
    If Not rsTemp.EOF Then
        getID_TO_���� = rsTemp!����
    End If
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Sub RemoveCheckImages(ByVal lngҽ��ID As Long, ByVal lng���ͺ� As Long)
    'ɾ��ָ��ҽ���ļ��Ӱ��
    Dim strSQL As String, rsTmp As New ADODB.Recordset
    
    Dim Inte As New clsFtp
    Dim strDeviceNO As String
    On Error GoTo ProcError
    '��ɾ��ͼ��
    strSQL = "select a.IP��ַ, a.FTPĿ¼, a.FTP�û���, a.FTP����, a.ҽ��ID, a.���ͺ�, a.���UID, a.λ��, a.�������� ,a.�豸�� ,c.ͼ��UID" & _
             " from (select IP��ַ, FTPĿ¼, FTP�û���, FTP����, ҽ��ID, ���ͺ�, ���UID, λ��һ as λ��, ��������, a.�豸�� " & _
             "       from Ӱ���豸Ŀ¼ a, Ӱ�����¼ b " & _
             "       Where a.�豸�� = B.λ��һ " & _
             "       Union All " & _
             "       select IP��ַ, FTPĿ¼, FTP�û���, FTP����, ҽ��ID, ���ͺ�, ���UID, λ�ö� as λ��, ��������, a.�豸��" & _
             "       from Ӱ���豸Ŀ¼ a, Ӱ�����¼ b " & _
             "       Where a.�豸�� = B.λ�ö� " & _
             "       Union All " & _
             "       select IP��ַ, FTPĿ¼, FTP�û���, FTP����, ҽ��ID, ���ͺ�, ���UID, λ���� as λ��, ��������, a.�豸�� " & _
             "       from Ӱ���豸Ŀ¼ a, Ӱ�����¼ b " & _
             "       Where a.�豸�� = B.λ���� " & _
             "       ) a , Ӱ�������� b , Ӱ����ͼ�� c " & _
             " Where a.���uid = B.���uid " & _
             " and b.����uid = c.����uid " & _
             " and a.ҽ��ID = [1] And ���ͺ� = [2] "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ͼ", lngҽ��ID, lng���ͺ�)
    Do Until rsTmp.EOF
        If strDeviceNO <> Nvl(rsTmp("�豸��")) Then
            strDeviceNO = Nvl(rsTmp("�豸��"))
            Inte.FuncFtpConnect Nvl(rsTmp("IP��ַ")), Nvl(rsTmp("FTP�û���")), Nvl(rsTmp("FTP����"))
        End If
        Inte.FuncDelFile IIf(IsNull(rsTmp("FTPĿ¼")), "", rsTmp("FTPĿ¼") & "/") & Format(rsTmp("��������"), "YYYYMMDD") & "/" & rsTmp("���UID"), rsTmp("ͼ��UID")
        rsTmp.MoveNext
    Loop
    strDeviceNO = ""
    Inte.FuncFtpDisConnect
    'ɾ��Ŀ¼
    strSQL = "select IP��ַ,FTPĿ¼,FTP�û���,FTP����,ҽ��ID,���ͺ�,���UID,�豸��,λ��,�������� from " & _
             "      (select IP��ַ,FTPĿ¼,FTP�û���,FTP����,ҽ��ID,���ͺ�,���UID,a.�豸��,λ��һ as λ��,�������� from Ӱ���豸Ŀ¼ a , Ӱ�����¼ b " & _
             "      Where a.�豸�� = B.λ��һ " & _
             "      Union All " & _
             "      select IP��ַ,FTPĿ¼,FTP�û���,FTP����,ҽ��ID,���ͺ�,���UID,a.�豸��,λ�ö� as λ��,�������� from Ӱ���豸Ŀ¼ a , Ӱ�����¼ b " & _
             "      Where a.�豸�� = B.λ�ö� " & _
             "      Union All " & _
             "      select IP��ַ,FTPĿ¼,FTP�û���,FTP����,ҽ��ID,���ͺ�,���UID,a.�豸��,λ���� as λ��,�������� from Ӱ���豸Ŀ¼ a , Ӱ�����¼ b " & _
             "      where a.�豸�� = b.λ���� ) a " & _
             " Where a.ҽ��ID = [1] And ���ͺ� = [2] "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��Ŀ¼", lngҽ��ID, lng���ͺ�)
    Do Until rsTmp.EOF
        If strDeviceNO <> Nvl(rsTmp("�豸��")) Then
            strDeviceNO = Nvl(rsTmp("�豸��"))
            Inte.FuncFtpConnect Nvl(rsTmp("IP��ַ")), Nvl(rsTmp("FTP�û���")), Nvl(rsTmp("FTP����"))
        End If
        Inte.FuncFtpDelDir IIf(IsNull(rsTmp("FTPĿ¼")), "", rsTmp("FTPĿ¼")), Format(rsTmp("��������"), "YYYYMMDD") & "/" & rsTmp("���UID")
        rsTmp.MoveNext
    Loop
    Inte.FuncFtpDisConnect
    Exit Sub
ProcError:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Public Function MovedByDate(ByVal vDate As Date) As Boolean
'���ܣ��ж�ָ������֮ǰ���Ƿ�����Ѿ�ִ��������ת��
'������vDate=ʱ����ʱ��εĿ�ʼʱ��
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "Select �ϴ����� From zlDataMove Where ϵͳ=[1] And ���=1 And �ϴ����� is Not Null"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", glngSys)
    If Not rsTmp.EOF Then
        '�ϴ�����û��ʱ��,"<"�ж���ת��������һ��
        If vDate < rsTmp!�ϴ����� Then
            MovedByDate = True
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function GetFullNO(ByVal strNO As String, ByVal intNUM As Integer) As String
'���ܣ����û�����Ĳ��ݵ��ţ�����ȫ���ĵ��š�
'������intNum=��Ŀ���,Ϊ0ʱ�̶��������
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, intType As Integer
    Dim curDate As Date
    
    If Len(strNO) >= 8 Then
        GetFullNO = Right(strNO, 8)
        Exit Function
    ElseIf Len(strNO) = 7 Then
        GetFullNO = PreFixNO & strNO
        Exit Function
    ElseIf intNUM = 0 Then
        GetFullNO = PreFixNO & Format(Right(strNO, 7), "0000000")
        Exit Function
    End If
    GetFullNO = strNO
    
    strSQL = "Select ��Ź���,Sysdate as ���� From ������Ʊ� Where ��Ŀ���=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlExpense", intNUM)
    If Not rsTmp.EOF Then
        intType = Nvl(rsTmp!��Ź���, 0)
        curDate = rsTmp!����
    End If

    If intType = 1 Then
        '���ձ��
        strSQL = Format(CDate("1992-" & Format(rsTmp!����, "MM-dd")) - CDate("1992-01-01") + 1, "000")
        GetFullNO = PreFixNO & strSQL & Format(Right(strNO, 4), "0000")
    Else
        '������
        GetFullNO = PreFixNO & Format(Right(strNO, 7), "0000000")
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Public Function InitSysPar() As Boolean
'��ʼ��ȫ�ֲ���
    Dim strValue As String
    On Error Resume Next
    
    gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "gstrSysName", "")
    strValue = zlDatabase.GetPara("���뷨")
    gstrIme = IIf(strValue = "", "���Զ�����", strValue)
    
    strValue = zlDatabase.GetPara(20, glngSys, , "7|7|7|7|7")
    gbytCardNOLen = Val(Split(strValue, "|")(4)) '���￨�ų���
    
        '���ý��С����λ��
    gbytDec = Val(zlDatabase.GetPara(9, glngSys, , 2))
    gstrDec = "0." & String(gbytDec, "0")
    
    If Not GetUserInfo Then
        MsgBox "��ǰ�û�δ���ö�Ӧ����Ա��Ϣ,����ϵͳ����Ա��ϵ,�ȵ��û���Ȩ���������ã�", vbInformation, gstrSysName
        Exit Function
    End If

    gstrUnitName = GetUnitName
    
    InitSysPar = True
End Function
Public Function MergeImageFiles(ByVal strCurrUID As String, ByVal strNewUID As String, _
    Optional ByVal strReceiveDate As String = "", Optional ByVal strMoveFiles As String = "") As Boolean
'------------------------------------------------
'���ܣ���һ������Ӱ���ļ�ת�Ƶ���������ȥ��֧��Ӱ�������ȡ������
'������ strCurrUID ����Դ���UID
'       strNewUID ����ת�ƺ��µ�Ŀ�ļ��UID
'       strReceiveDate ���� �������ڣ���������ͼ��洢·������strNewUID�������ݿ���ʱ������Ҫʹ�ñ�����
'       strMoveFiles ���� ��Ҫ�ƶ����ļ����б�ʹ��"|"�ָ��ļ��������û�У���ת��Դ���UIDָ���Ŀ¼�µ�����ͼ��
'���أ�True--�ɹ���False��ʧ��
'------------------------------------------------
    Dim objSrcFtp As New clsFtp, objDestFtp As New clsFtp
    Dim strSrcPath As String, strDestPath As String
    Dim rsTmp As New ADODB.Recordset, strSQL As String, strTmpFile As String
    Dim aFiles() As String, i As Integer, objFile As New Scripting.FileSystemObject
    Dim strFTPUser As String, strFTPPassw As String, strFTPHost As String, strFTPRoot As String
    Dim lngResult As Long       '��¼FTP�����Ľ��
        
    '����¼��UID���ɼ��UID������Ϊ�ϲ���ɣ���ֱ���˳�
    If strCurrUID = strNewUID Then
        MergeImageFiles = True
        Exit Function
    End If
    
    On Error GoTo errH

    '�����ƶ��ķ���ͬ��Դͼ�п����ڡ�Ӱ����ʱ��¼�����ߡ�Ӱ�����¼����
    '����ʱ����ʱ��¼���Ƶ�������¼��ȡ������ʱ��������¼���Ƶ���ʱ��¼
    
    strSQL = "Select D.FTP�û��� As FtpUser,D.FTP���� As FtpPwd," & _
        "D.IP��ַ As Host," & _
        "'/'||D.FtpĿ¼||'/' As Root,Decode(C.��������,Null,'',to_Char(C.��������,'YYYYMMDD')||'/')" & _
        "||C.���UID As URL " & _
        "From Ӱ�����¼ C,Ӱ���豸Ŀ¼ D " & _
        "Where Decode(C.λ��һ,Null,C.λ�ö�,C.λ��һ)=D.�豸��(+)" & _
        "And C.���UID= [1] Union All " & _
        "Select D.FTP�û��� As FtpUser,D.FTP���� As FtpPwd," & _
        "D.IP��ַ As Host," & _
        "'/'||D.FtpĿ¼||'/' As Root,Decode(C.��������,Null,'',to_Char(C.��������,'YYYYMMDD')||'/')" & _
        "||C.���UID As URL " & _
        "From Ӱ����ʱ��¼ C,Ӱ���豸Ŀ¼ D " & _
        "Where Decode(C.λ��һ,Null,C.λ�ö�,C.λ��һ)=D.�豸��(+)" & _
        "And C.���UID= [1]"
    '�����ݿ��в�ѯ�ɼ��UID
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "ZLPACSWork", strCurrUID)
    '��ǰ���UID�����ݿ��в����ڣ����˳�������
    If rsTmp.EOF Then
        Exit Function
    End If
    
    '�洢������FTP��������
    strFTPHost = Nvl(rsTmp("Host"))
    strFTPPassw = Nvl(rsTmp("FtpPwd"))
    strFTPRoot = Nvl(rsTmp("Root"))
    strFTPUser = Nvl(rsTmp("FtpUser"))
    strSrcPath = Nvl(rsTmp("Root")) & Nvl(rsTmp("URL"))
    lngResult = objSrcFtp.FuncFtpConnect(strFTPHost, strFTPUser, strFTPPassw)
    If lngResult = 0 Then Exit Function     'FTP����ʧ�ܣ��˳�����
    
    '�����ݿ��в�ѯ�¼��UID����ʼ��Ŀ��Ftp,���Ŀ��UID�����ڣ�����һ����·��
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "ZLPACSWork", strNewUID)
    If rsTmp.EOF Then
    '������ͼ��ת����ʱͼ���ʱ��Ŀ�ļ��UID��ʱ������������ݿ��У���ʱֱ��ʹ��ԭ�е�FTP����
    '�������ݿ���ת�Ƽ�¼��ʱ�򣬻���ʹ��ԭ����FTP����
        If strReceiveDate <> "" Then
                objDestFtp.FuncFtpConnect strFTPHost, strFTPUser, strFTPPassw
                strDestPath = strFTPRoot & Format(strReceiveDate, "YYYYMMDD") & "/" & strNewUID
                '����FTPĿ¼
                objDestFtp.FuncFtpMkDir strFTPRoot, Format(strReceiveDate, "YYYYMMDD") & "/" & strNewUID
        Else
            Exit Function
        End If
    Else
        objDestFtp.FuncFtpConnect Nvl(rsTmp("Host")), Nvl(rsTmp("FtpUser")), Nvl(rsTmp("FtpPwd"))
        strDestPath = Nvl(rsTmp("Root")) & Nvl(rsTmp("URL"))
    End If
    
    '��ȡ��Ҫ�ƶ����ļ���
    If strMoveFiles <> "" Then
        aFiles = Split(strMoveFiles, "|")
    Else
        aFiles = Split(objSrcFtp.FuncDirFiles(strSrcPath), "|")
    End If
    
    '��ת��ͼ��
    For i = 0 To UBound(aFiles)
        strTmpFile = App.Path & "\TmpImage\" & aFiles(i)
        lngResult = objSrcFtp.FuncDownloadFile(strSrcPath, strTmpFile, aFiles(i))
        If lngResult <> 0 Then Exit Function
        lngResult = objDestFtp.FuncUploadFile(strDestPath, strTmpFile, aFiles(i))
        If lngResult <> 0 Then Exit Function
    Next i
    
    'ת��ͼ��ɹ�����ɾ����ʱͼ���ԭ��FTP��ͼ���Ŀ¼���峡�������ִ�����Բ�����
    On Error Resume Next
    For i = 0 To UBound(aFiles)
        strTmpFile = App.Path & "\TmpImage\" & aFiles(i)
        Kill strTmpFile
        Call objSrcFtp.FuncDelFile(strSrcPath, aFiles(i))
    Next i
    Call objSrcFtp.FuncFtpDelDir(Replace(strSrcPath, strCurrUID, ""), strCurrUID)
    
    objSrcFtp.FuncFtpDisConnect
    objDestFtp.FuncFtpDisConnect
    MergeImageFiles = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub MkLocalDir(ByVal strDir As String)
'------------------------------------------------
'���ܣ���������Ŀ¼
'������ strDir��������Ŀ¼
'���أ���
'------------------------------------------------
    Dim objFile As New Scripting.FileSystemObject
    Dim aNestDirs() As String, i As Integer
    Dim strPath As String
    On Error Resume Next
    
    '��ȡȫ����Ҫ������Ŀ¼��Ϣ
    ReDim Preserve aNestDirs(0)
    aNestDirs(0) = strDir
    
    strPath = objFile.GetParentFolderName(strDir)
    Do While Len(strPath) > 0
        ReDim Preserve aNestDirs(UBound(aNestDirs) + 1)
        aNestDirs(UBound(aNestDirs)) = strPath
        strPath = objFile.GetParentFolderName(strPath)
    Loop
    '����ȫ��Ŀ¼
    For i = UBound(aNestDirs) To 0 Step -1
        MkDir aNestDirs(i)
    Next
End Sub


Public Sub ClearCacheFolder(ByVal strCacheFolder As String)
'------------------------------------------------
'���ܣ���ָ��Ŀ¼�Ĵ�С�ﵽһ���ٷֱ�ʱ����ո�Ŀ¼
'������ strCacheFolder--��Ҫ����Ƿ���յ�Ŀ¼
'���أ���
'------------------------------------------------
    Dim objFile As New Scripting.FileSystemObject
    Dim objCurFolder As Scripting.Folder, objCurFile As Scripting.File, objFiles As Scripting.Files
    Dim strDriver As String
    
    On Error Resume Next
    strDriver = objFile.GetDriveName(strCacheFolder)
    Set objCurFolder = objFile.GetFolder(strCacheFolder)
    If objCurFolder.Size / objFile.GetDrive(strDriver).FreeSpace > 0.2 Then
        objCurFolder.Delete True
    End If
End Sub

Public Function GetTrayHeight() As Long
    '------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�������ĸ߶�
    '------------------------------------------------------------------------------------------------------------------
    Dim lngHwd As Long
    Dim objRect As RECT
    
    On Error Resume Next
    
    lngHwd = FindWindow("shell_traywnd", "")
    Call GetWindowRect(lngHwd, objRect)

    GetTrayHeight = Screen.TwipsPerPixelX * (objRect.Bottom - objRect.Top)
    
    If GetTrayHeight < 0 Then GetTrayHeight = 0
    
End Function

Public Sub ResizeRegion(ByVal ImageCount As Integer, ByVal RegionWidth As Long, _
    ByVal RegionHeight As Long, Rows As Integer, Cols As Integer)
'-----------------------------------------------------------------------------
'���ܣ����������ͼ��������ͼ������Ŀ�Ⱥ͸߶ȣ�������ѵ�ͼ����������������
'������ ImageCount����ͼ������
'       RegionWidth--ͼ����ʾ����Ŀ��
'       RegionHeight--ͼ����ʾ����ĸ߶�
'       Rows����[����]�������
'       Cols����[����]�������
'���أ������������Rows���������Cols
'-----------------------------------------------------------------------------
    Dim iCols As Integer, iRows As Integer
    Dim iBase As Integer, blnDoLoop As Integer
    
    On Error GoTo err
    iCols = CInt(Sqr(ImageCount * RegionWidth / RegionHeight))
    iRows = CInt(Sqr(ImageCount * RegionHeight / RegionWidth))

    If iRows < 1 Then iRows = 1
    If iCols < 1 Then iCols = 1
    Do While iRows * iCols > ImageCount
        If RegionWidth / RegionHeight > 1 Then
            iCols = iCols - 1
        Else
            iRows = iRows - 1
        End If
    Loop
    
    If iRows < 1 Then iRows = 1
    If iCols < 1 Then iCols = 1
    Do While iRows * iCols < ImageCount
        If RegionWidth / RegionHeight > 1 Then
            iCols = iCols + 1
        Else
            iRows = iRows + 1
        End If
    Loop
    Rows = iRows: Cols = iCols
    
    If ImageCount <> 0 Then
        If Rows * Cols > ImageCount Then
            iBase = 6
            blnDoLoop = True
            
            While blnDoLoop
                iBase = iBase - 1
                
                If ImageCount Mod iBase = 0 Then
                    blnDoLoop = False
                End If
            Wend
        

            If RegionWidth > RegionHeight Then
                If ImageCount / iBase > iBase Then
                    Cols = ImageCount / iBase
                    Rows = iBase
                Else
                    Rows = ImageCount / iBase
                    Cols = iBase
                End If
            Else
                If ImageCount / iBase > iBase Then
                    Cols = iBase
                    Rows = ImageCount / iBase
                Else
                    Rows = iBase
                    Cols = ImageCount / iBase
                End If
            End If
        End If
    End If
err:
End Sub


Public Function funGetStudyUID(ByVal strOldStudyUID As String) As String
'-----------------------------------------------------------------------------
'����:��ѯ���ݿ⣬�жϵ�ǰͼ��ļ��UID�Ƿ��Ѿ����������������ʱ���У�
'     ������ڣ����ڼ��UID�������Ӻ�׺����������ֱ�ӷ�������ļ��UID
'�޸���:�ƽ�
'�޸�����:2007-1-27
'-----------------------------------------------------------------------------
    '
    Dim rsMatch As New ADODB.Recordset
    
    funGetStudyUID = strOldStudyUID
    gstrSQL = "select ���UID from Ӱ�����¼ where ���UID = [1]" & _
              " Union All Select ���UID from Ӱ����ʱ��¼ where ���UID = [1]"
    Set rsMatch = zlDatabase.OpenSQLRecord(gstrSQL, "PACSͼ�񱣴�", strOldStudyUID)
    If Not rsMatch.EOF Then
        '����һ���µļ��UID
        gstrSQL = "Select Ӱ����UID���_ID.Nextval From Dual"
        Set rsMatch = zlDatabase.OpenSQLRecord(gstrSQL, "PACSͼ�񱣴�")
        If Len(strOldStudyUID) <= 55 Then
            funGetStudyUID = strOldStudyUID & ".A" & rsMatch(0)
        Else
            funGetStudyUID = Left(strOldStudyUID, 55) & ".A" & rsMatch(0)
        End If
    End If
End Function


Public Function GetImageAttribute(objAttr As DicomAttributes, ByVal AttrName As String) As Variant
'-----------------------------------------------------------------------------
'����:��ȡDICOM���Լ��е�ָ������ֵ
'�޸���:�ƽ�
'�޸�����:2007-2-6
'-----------------------------------------------------------------------------
    Dim AttrTag() As String
    
    GetImageAttribute = ""
    AttrTag = Split(AttrName, ":")
    If objAttr("&h" & AttrTag(0), "&h" & AttrTag(1)).Exists Then
        GetImageAttribute = Nvl(objAttr("&h" & AttrTag(0), "&h" & AttrTag(1)).Value)
    End If
End Function

'Public Function funRelateSeries(lngҽ��ID As Long, lng���ͺ� As Long)
''-----------------------------------------------------------------------------
''����:����ͼ���ƶ�FTPͼ���µ�λ�ã��޸����ݿ��¼������ʱ��ת����ʽ����
''������ lngҽ��ID ����ҽ��ID
''       lng���ͺ� ���� ���ͺ�
''���أ���
''-----------------------------------------------------------------------------
'    Dim blnCancel As Boolean, rsTmp As ADODB.Recordset
'    Dim rsStudyUID As ADODB.Recordset
'    Dim strFilter As String
'    Dim strModality As String
'
'    On Error GoTo errHandle
'
'    gstrSQL = "Select Ӱ����� From Ӱ�����¼ Where ҽ��ID= [1] And ���ͺ� = [2]"
'    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "����ͼ����ȡӰ���౻", lngҽ��ID, lng���ͺ�)
'    strModality = Nvl(rsTmp!Ӱ�����)
'
'    gstrSQL = "Select 0 as ѡ��, A.���UID As ID,Nvl(A.����,' ') As ����,Nvl(A.Ӣ����,' ') As Ӣ����," & _
'            "Nvl(A.����,0) As ����,Nvl(A.�Ա�,' ') As �Ա�,Nvl(A.����,' ') As ����," & _
'            "Nvl(A.����豸,' ') As ����豸,to_char(Nvl(A.��������,Sysdate),'YYYY-MM-DD hh24:mi:ss') As ���ʱ��," & _
'            "to_char(A.��������,'YYYY-MM-DD') As ��������," & _
'            "Nvl(A.���,0) As ���,Nvl(A.����,0) As ����, c.��������,a.Ӱ����� " & _
'            " From Ӱ����ʱ��¼ a," & _
'            "(Select x.��������,x.���uid, row_number() over(partition by ���UID order by ���UID) As  rank from Ӱ����ʱ���� x) c " & _
'            " Where a.���UID = c.���UID And c.rank = 1"
'
'    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "����ͼ��", lngҽ��ID, lng���ͺ�)
'
'    frmSelectMuli.ShowSelect rsTmp, "ID,900,0,1;Ӱ�����,900,0,1;����,800,0,1;Ӣ����,800,0,1;����,900,0,1;" _
'            & "�Ա�,600,0,1;����,600,0,1;��������,1200,0,1;����豸,900,0,1;���ʱ��,1200,0,1;��������,1200,0,1;" _
'            & "���,500,0,1;����,500,0,1", 0, 0, 14000, 10000, "����ͼ��", , , strModality
'
'    If frmSelectMuli.mblnOK = True And frmSelectMuli.strFilter <> "ID=-1" Then
'
'        If MsgBox("�Ƿ�ȷ��ѡ���Ӱ���ǵ�ǰ���ģ�", vbDefaultButton2 + vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Function
'        strFilter = frmSelectMuli.strFilter
'        rsTmp.Filter = strFilter
'        '�����ѡ�е���ʱ��¼������ÿһ��ʱ��¼�Ĺ���
'        While Not rsTmp.EOF
'            '�ƶ�Ftp�ϵ�Ӱ���ļ�,�ƶ��ɹ����Ÿ������ݿ�
'            gstrSQL = "Select ���UID From Ӱ�����¼ Where ҽ��ID=[1] And ���ͺ�=[2]"
'            Set rsStudyUID = zlDatabase.OpenSQLRecord(gstrSQL, "����Ӱ��", lngҽ��ID, lng���ͺ�)
'            If Not rsStudyUID.EOF Then
'                If Len(Trim(Nvl(rsStudyUID(0)))) > 0 Then
'                    If MergeImageFiles(rsTmp("ID"), rsStudyUID(0)) = False Then
'                        MsgBox "�ļ�ת��ʧ�ܣ����ܹ���Ӱ��" & vbCrLf & vbCrLf & "�������������������⣬���顣"
'                        Exit Function
'                    End If
'                End If
'            End If
'
'            gstrSQL = "ZL_Ӱ����_SET(" & lngҽ��ID & "," & lng���ͺ� & ",'" & _
'                rsTmp("ID") & "')"
'            zlDatabase.ExecuteProcedure gstrSQL, "����Ӱ��"
'
'            rsTmp.MoveNext
'        Wend
'    End If
'
'    Exit Function
'errHandle:
'    If ErrCenter() = 1 Then Resume
'    Call SaveErrLog
'End Function
Public Function SetDeptPara(ByVal lngDeptID As Long, ByVal varPara As String, ByVal strValue As String) As Boolean
'���ܣ�����ָ���Ĳ���ֵ
'������lngDept=����ID
'      varPara=������
'      strValue=������ֵ
'���أ������Ƿ�ɹ�
    Dim strSQL As String
    
    On Error GoTo errH
        
    strSQL = "ZL_Ӱ�����̲���_UPDATE(" & lngDeptID & ",'" & varPara & "','" & strValue & "')"
    Call zlDatabase.ExecuteProcedure(strSQL, "SetPara")
    
    '���óɹ����������
    Set mrsDeptParas = Nothing
    
    SetDeptPara = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
End Function
Public Function GetDeptPara(ByVal lngDeptID As Long, ByVal varPara As String, Optional ByVal strDefault As String, Optional ByVal blnNotCache As Boolean) As String
'���ܣ���ȡָ���Ĳ���ֵ
'������lngDept=����ID
'      varPara=������
'      strDefault=�����ݿ���û�иò���ʱʹ�õ�ȱʡֵ(ע�ⲻ��Ϊ��ʱ)
'      blnNotCache=�Ƿ񲻴ӻ����ж�ȡ
'���أ�����ֵ���ַ�����ʽ
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnNew As Boolean
    
    On Error GoTo errH
    
    If blnNotCache Then
        Set rsTmp = New ADODB.Recordset
        strSQL = "Select ����ֵ from Ӱ�����̲��� where ����ID = [1] and ������=[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ����", lngDeptID, varPara)
        
        If Not rsTmp.EOF Then
            GetDeptPara = Nvl(rsTmp!����ֵ)
        Else
            GetDeptPara = strDefault
        End If
    Else
        '��һ�μ��ز�������
        If mrsDeptParas Is Nothing Then
            blnNew = True
        ElseIf mrsDeptParas.State = 0 Then
            blnNew = True
        End If
        If blnNew Then
            strSQL = "Select ����ֵ,������,����ID from Ӱ�����̲���"
            Set mrsDeptParas = New ADODB.Recordset
            Set mrsDeptParas = zlDatabase.OpenSQLRecord(strSQL, "��ȡ����")
        End If
        
        '���ݻ����ȡ����ֵ
        mrsDeptParas.Filter = "������='" & CStr(varPara) & "' AND ����ID=" & lngDeptID
        If Not mrsDeptParas.EOF Then
            GetDeptPara = Nvl(mrsDeptParas!����ֵ)
        Else
            GetDeptPara = strDefault
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
End Function

Public Function GetIsValidOfStorageDevice(ByVal lngDeptID As Long) As Boolean
'��ʼ�����Ҽ�����
    Dim rsTmp As New ADODB.Recordset
    Dim strSaveDeviceId As String
    
    On Error GoTo DBError
    
    '��ȡ�����洢�豸��
    strSaveDeviceId = GetDeptPara(lngDeptID, "�洢�豸��")
    
    gstrSQL = "Select �豸��,�豸�� From Ӱ���豸Ŀ¼ Where ����=1 and �豸��=[1] and NVL(״̬,0)=1"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�洢�豸��Ϣ", strSaveDeviceId)
    
    
    GetIsValidOfStorageDevice = Not rsTmp.EOF
    
    Exit Function
DBError:
    GetIsValidOfStorageDevice = False
    
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub subCancelSeriesRelate(lngAdviceNo As Long, lngSendNO As Long, strSeriesNo As String)
'-----------------------------------------------------------------------------
'����:ȡ������ͼ��Ĺ������ƶ�FTPͼ���µ�λ�ã��޸����ݿ��¼������ʽ���ƶ�����ʱ����
'������ lngAdviceNo ����ҽ��ID
'       lngSendNO ���� ���ͺ�
'       strSeriesNo ��������UID
'���أ���
'-----------------------------------------------------------------------------
    
    Dim mcnFTP As New clsFtp
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strCachePath As String
    Dim strCacheFileName As String
    Dim objFile As New Scripting.FileSystemObject
    Dim imgs As New DicomImages
    Dim img As New DicomImage
    Dim strNewStudyUID As String    '�����ɵļ��UID
    Dim strOldStudyUID As String    'ͼ������ԭ���ļ��UID
    Dim strDBStudyUID As String     '���ݿ��б���ļ��UID����ͼ��洢·�����
    Dim strMoveFiles As String  '�洢��Ҫ�ƶ���ͼ���ļ�����ʹ�á�|���ָ�
    Dim blnNoImage As Boolean   '1û��ͼ��ֱ�Ӷ�ȡ���ݿ���Ϣ��0��ͼ��ʹ��ͼ����Ϣ
    Dim lngResult As Long    '��¼FTP���ؽ��
    
    'ͼ���еĲ��˻�����Ϣ
    Dim strModality As String
    Dim strPatientID As String
    Dim strPatientName As String
    Dim strSex As String
    Dim strAge As String
    Dim strDateOfBirth As String
    Dim strManufacturer As String
    Dim strReceiveDateTime As String
    
    
    On Error GoTo DBError
    
    '���������е�һ��ͼ��� ����ID��Ӣ�������Ա����䣬�������ڣ����UID������豸������ʱ��
    strCachePath = App.Path & "\TmpImage\"
    strSQL = "Select A.ͼ���,D.FTP�û��� As User1,D.FTP���� As Pwd1,a.ͼ��UID, " & _
        "D.IP��ַ As Host1,c.���uid," & _
        "'/'||D.FtpĿ¼||'/' As Root1,Decode(C.��������,Null,'',to_Char(C.��������,'YYYYMMDD')||'/')" & _
        "||C.���UID||'/'||A.ͼ��UID As URL,d.�豸�� as �豸��1, " & _
        "E.FTP�û��� As User2,E.FTP���� As Pwd2," & "E.IP��ַ As Host2," & _
        "'/'||E.FtpĿ¼||'/' As Root2,e.�豸�� as �豸��2 " & _
        "From Ӱ����ͼ�� A,Ӱ�������� B,Ӱ�����¼ C,Ӱ���豸Ŀ¼ D,Ӱ���豸Ŀ¼ E " & _
        "Where A.����UID=B.����UID And B.���UID=C.���UID And C.λ��һ=D.�豸��(+) And C.λ�ö�=E.�豸��(+) " & _
        "And A.����UID= [1] Order By A.ͼ���"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "������", strSeriesNo)
    
    If Not rsTmp.EOF Then   '�����д���ͼ��
        strDBStudyUID = Nvl(rsTmp("���uid"))
        '�½�����Ŀ¼
        strCacheFileName = strCachePath & rsTmp("URL")
        MkLocalDir objFile.GetParentFolderName(strCacheFileName)
        
        '����ͼ��
        If rsTmp("�豸��1") <> "" And mcnFTP.FuncFtpConnect(Nvl(rsTmp("Host1")), Nvl(rsTmp("User1")), Nvl(rsTmp("Pwd1"))) <> 0 Then
            mcnFTP.FuncDownloadFile objFile.GetParentFolderName(Nvl(rsTmp("Root1")) & rsTmp("URL")), strCacheFileName, objFile.GetFileName(rsTmp("URL"))
            mcnFTP.FuncFtpDisConnect
        ElseIf rsTmp("�豸��2") <> "" And mcnFTP.FuncFtpConnect(Nvl(rsTmp("Host2")), Nvl(rsTmp("User2")), Nvl(rsTmp("Pwd2"))) <> 0 Then
            mcnFTP.FuncDownloadFile objFile.GetParentFolderName(Nvl(rsTmp("Root2")) & rsTmp("URL")), strCacheFileName, objFile.GetFileName(rsTmp("URL"))
            mcnFTP.FuncFtpDisConnect
        Else
            'FTP���Ӵ�����ʾ���˳�����ȡ����������
            MsgBox "FTP���Ӵ��󣬲���ȡ��������" & vbCrLf & vbCrLf & "�������������ӳ������⡣"
            Exit Sub
        End If
                    
        '��ȡͼ����Ϣ
        If Dir(strCacheFileName) <> vbNullString Then
            Set img = imgs.ReadFile(strCacheFileName)
            'ʹ�ñ�����ͼ�������Ϣ��ȡ����
            strOldStudyUID = img.StudyUID
            strModality = GetImageAttribute(img.Attributes, ATTR_Ӱ�����)
            strPatientID = img.PatientID
            strPatientName = img.Name
            strSex = img.Sex
            If IsDate(img.DateOfBirthAsDate) Then
                If img.Attributes(&H10, &H1010).Exists And Not IsNull(img.Attributes(&H10, &H1010)) Then
                    strAge = img.Attributes(&H10, &H1010).Value
                Else
                    strAge = CStr(Year(Date) - Year(img.DateOfBirthAsDate))
                End If
                        
                If img.DateOfBirthAsDate <> "0:00:00" Then
                    strDateOfBirth = Format(img.DateOfBirthAsDate, "YYYY-MM-DD")
                Else
                    strDateOfBirth = ""
                End If
            Else
                strAge = "": strDateOfBirth = ""
            End If
            strManufacturer = GetImageAttribute(img.Attributes, ATTR_����豸)
            strReceiveDateTime = GetImageAttribute(img.Attributes, ATTR_�������) & " " & _
                        Format(GetImageAttribute(img.Attributes, ATTR_���ʱ��), "HH:MM")
            'ɾ����ʱͼ��
            Set img = Nothing
            imgs.Remove (1)
            On Error Resume Next
            objFile.DeleteFile strCacheFileName
            On Error GoTo 0
        Else
            '�����һ��ͼ�����ز���ȷ����ȡ���ݿ���Ϣ���������������
            blnNoImage = True
        End If
    Else
        '������û��ͼ�󣬲���Ҫȡ��������Ӧ�ò�������������
        Exit Sub
    End If
    
    '����û��ͼ����Ϣ�ɶ�ȡ������ͼ����Ҫ��Ϣ��ȡ�������ģ�ֱ�Ӷ�ȡ���ݿ��е���Ϣ
    If blnNoImage = True Or Trim(strReceiveDateTime) = "" Then
        strSQL = "select a.Ӱ�����,a.����,a.����,a.Ӣ����,a.�Ա�,a.����,a.��������,a.���uid," & _
                " a.����豸,a.�������� from Ӱ�����¼ a where a.ҽ��id =[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ����", lngAdviceNo)
        If Not rsTmp.EOF Then
            If blnNoImage = True Then
                strOldStudyUID = Nvl(rsTmp("���uid"))
                strDBStudyUID = Nvl(rsTmp("���uid"))
                strPatientID = Nvl(rsTmp("����"))
                strPatientName = Nvl(rsTmp("Ӣ����"))
                strSex = Nvl(rsTmp("�Ա�"))
                strAge = Nvl(rsTmp("����"))
                strDateOfBirth = Nvl(rsTmp("��������"), "")
                strManufacturer = Nvl(rsTmp("����豸"))
            End If
            strModality = Nvl(rsTmp("Ӱ�����"))
            strReceiveDateTime = Nvl(rsTmp("��������"))
        End If
    End If
    '��֯ͼ���ļ����ƴ�
    strSQL = "select ͼ��UID from Ӱ�������� a,Ӱ����ͼ�� b where a.����UID =[1] and a.����UID = b.����UID"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ����", strSeriesNo)
    If Not rsTmp.EOF Then
        strMoveFiles = rsTmp(0)
        rsTmp.MoveNext
        While Not rsTmp.EOF
            strMoveFiles = strMoveFiles & "|" & rsTmp(0)
            rsTmp.MoveNext
        Wend
    End If
    
    '������UID�����ݿ����ִ�ļ��UID��ͬ���򴴽��µļ��UID�����޸�ͼ��FTP·��
    strNewStudyUID = funGetStudyUID(strOldStudyUID)
    If strNewStudyUID <> strDBStudyUID Then
        If MergeImageFiles(strDBStudyUID, strNewStudyUID, Format(strReceiveDateTime, "YYYY-MM-DD"), strMoveFiles) = False Then
            MsgBox "ͼ��ת�Ʋ��ɹ�������ȡ��������"
            Exit Sub
        End If
    End If
    
    '�޸����ݿ⣬������¼ת����ʱ��¼
    strSQL = "ZL_Ӱ����_PhotoCancel(" & lngAdviceNo & "," & lngSendNO & ",'" & strNewStudyUID & "','" & _
              strSeriesNo & "','" & strModality & "'," & Val(strPatientID) & ",'" & _
              strPatientName & "','" & strSex & "','" & strAge & "'," & _
              IIf(Len(strDateOfBirth) = 0, "null", "to_date('" & strDateOfBirth & "','YYYY-MM-DD')") & _
              ",'" & strManufacturer & "',to_date('" & strReceiveDateTime & "','YYYY-MM-DD HH24:MI:SS'))"
              
    zlDatabase.ExecuteProcedure strSQL, "ȡ������"
    
    Exit Sub
DBError:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Public Sub GetAllImages(dcmViewer As DicomViewer, blnMoved As Boolean, intSearchType As Integer, _
    Optional lngAdviceID As Long, Optional strSeriesUID As String, Optional intGetImgNum As Integer = 0, _
    Optional intShowImgNum As Integer = 0, Optional blnTempDB As Boolean = False, _
    Optional strStudyUID As String = "", Optional strImageUID As String = "")
'------------------------------------------------
'���ܣ�ɾ��dcmViewer�е�ͼ��󣬽���ȡ��ͼ���ļ�����dcmViewer��
'������ dcmViewer������ͼ���DicomViewer
'       blnMoved ���� �Ƿ�ת����
'       intSearchType ������������,ֻ����ʽ���ѯ��Ч  1������ҽ��ID�ͷ��ͺŲ飬2����������UID�飬3 - ����ͼ��UID��
'       lngAdviceID ���� ҽ��ID
'       strSeriesUID ���� ����UID
'       intGetImgNum �������ζ�ȡ��ͼ������
'       intShowImgNum ����������ʾ��ͼ������
'       blnTempDB - - �Ƿ����ʱ������ȡͼ��
'       strStudyUID - - ���UID,ֻ�д���ʱ����ҵ�ʱ�򣬲�ʹ���������
'       strImageUID - - ͼ��UID��ֻ�д���ʽ����ҵ�ʱ�򣬲�ʹ���������
'���أ��ޣ�ֱ���޸�dcmViewer����ʾ��ͼ��
'------------------------------------------------
    
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim curImage As DicomImage, i As Integer
    Dim iCols As Integer, iRows As Integer
    Dim objFile As New Scripting.FileSystemObject, strTmpFile As String
    Dim Inet1 As New clsFtp
    Dim Inet2 As New clsFtp
    Dim strCachePath As String
    Dim iCurrentIndex As Integer
    
    On Error GoTo DBError
    If blnTempDB = False Then       '����ʽͼ����в���ͼ��
        strSQL = "Select /*+RULE*/ A.ͼ���,D.FTP�û��� As User1,D.FTP���� As Pwd1," & _
            "D.IP��ַ As Host1,'/'||D.FtpĿ¼||'/' As Root1," & _
            "Decode(C.��������,Null,'',to_Char(C.��������,'YYYYMMDD')||'/')" & _
            "||C.���UID||'/'||A.ͼ��UID As URL,d.�豸�� as �豸��1, " & _
            "E.FTP�û��� As User2,E.FTP���� As Pwd2," & _
            "E.IP��ַ As Host2,'/'||E.FtpĿ¼||'/' As Root2," & _
            "e.�豸�� as �豸��2,C.���UID,B.����UID " & _
            "From Ӱ����ͼ�� A,Ӱ�������� B,Ӱ�����¼ C,Ӱ���豸Ŀ¼ D,Ӱ���豸Ŀ¼ E " & _
            "Where A.����UID=B.����UID And B.���UID=C.���UID And C.λ��һ=D.�豸��(+) And C.λ�ö�=E.�豸��(+) "
        If blnMoved Then
            strSQL = Replace(strSQL, "Ӱ����ͼ��", "HӰ����ͼ��")
            strSQL = Replace(strSQL, "Ӱ��������", "HӰ��������")
            strSQL = Replace(strSQL, "Ӱ�����¼", "HӰ�����¼")
        End If
        If intShowImgNum <> 0 Then
            strSQL = strSQL & " And Rownum<=[2] "
        End If
        
        If intSearchType = 1 Then       '1������ҽ��ID�ͷ��ͺŲ�
            strSQL = strSQL & "And C.ҽ��ID=[1] Order By A.ͼ���"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡͼ��", lngAdviceID, intGetImgNum)
        ElseIf intSearchType = 2 Then   '2����������UID��
            strSQL = strSQL & "And A.����UID= [1] Order By A.ͼ���"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡͼ��", strSeriesUID, intGetImgNum)
        ElseIf intSearchType = 3 Then   '3 - ����ͼ��UID��
            strSQL = strSQL & "And A.ͼ��UID = [1] "
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡͼ��", strImageUID, intGetImgNum)
        End If
        
    Else                '����ʱ���в���ͼ��
        
        strSQL = "Select /*+RULE*/ c.ͼ���,d.FTP�û��� As User1, d.FTP���� As Pwd1, d.Ip��ַ As Host1," _
                & "'/' || d.FtpĿ¼ || '/' As Root1," _
                & " Decode(a.��������, Null, '', To_Char(a.��������, 'YYYYMMDD') || '/') || a.���uid || '/' || c.ͼ��uid As URL," _
                & " d.�豸�� As �豸��1,a.���UID,b.����UID,d.FTP�û��� As User2, d.FTP���� As Pwd2, " _
                & " d.Ip��ַ As Host2, '/' || d.FtpĿ¼ || '/' As Root2, " _
                & " d.�豸�� As �豸��2 " _
                & " From Ӱ����ʱ��¼ a,Ӱ����ʱ���� b,Ӱ����ʱͼ�� c ,Ӱ���豸Ŀ¼ d " _
                & " Where a.���UID=b.���UID And b.����UID = c.����UID And a.λ��һ = d.�豸�� "
                
        If strStudyUID <> "" Then   '���ռ��uid����
            strSQL = strSQL & "And a.���UID=[1] and Rownum<=[2] Order By c.ͼ���  "
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡͼ��", strStudyUID, CLng(6))
        Else        '��������UID����
            strSQL = strSQL & "And b.����UID=[1] and Rownum<=[2] Order By c.ͼ���  "
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡͼ��", strSeriesUID, CLng(6))
        End If
    End If
    
        dcmViewer.Images.Clear
        If rsTmp.RecordCount > 0 Then
            If intShowImgNum = 0 Or intShowImgNum >= rsTmp.RecordCount Then
                ResizeRegion rsTmp.RecordCount, dcmViewer.Width, dcmViewer.Height, iRows, iCols
            Else
                ResizeRegion intShowImgNum, dcmViewer.Width, dcmViewer.Height, iRows, iCols
            End If
            dcmViewer.MultiColumns = iCols
            dcmViewer.MultiRows = iRows
            
            '��������Ŀ¼
            strCachePath = App.Path & "\TmpImage\"
            MkLocalDir strCachePath & objFile.GetParentFolderName(Nvl(rsTmp("URL")))
            
            Do While Not rsTmp.EOF
                If Dir(strCachePath & Nvl(rsTmp("URL"))) = vbNullString Then
                    '���ػ���ͼ�񲻴��ڣ����ȡFTPͼ��
                    strTmpFile = strCachePath & Nvl(rsTmp("URL"))
                    '����FTP����
                    If Nvl(rsTmp("�豸��1")) <> vbNullString And Inet1.hConnection = 0 Then
                        If Inet1.FuncFtpConnect(Nvl(rsTmp("Host1")), Nvl(rsTmp("User1")), Nvl(rsTmp("Pwd1"))) = 0 Then
                            If Nvl(rsTmp("�豸��2")) <> vbNullString Then
                                If Inet2.FuncFtpConnect(Nvl(rsTmp("Host2")), Nvl(rsTmp("User2")), Nvl(rsTmp("Pwd2"))) = 0 Then
                                    MsgBox "FTP�����������ӣ������������á�"
                                    Exit Sub
                                End If
                            Else
                                MsgBox "FTP�����������ӣ������������á�"
                                Exit Sub
                            End If
                        End If
                    End If
                    If Inet1.FuncDownloadFile(objFile.GetParentFolderName(Nvl(rsTmp("Root1")) & rsTmp("URL")), strTmpFile, objFile.GetFileName(rsTmp("URL"))) <> 0 Then
                        '���豸��1��ȡͼ��ʧ�ܣ�����豸��2��ȡͼ��
                        If Nvl(rsTmp("�豸��2")) <> vbNullString Then
                            If Inet2.hConnection = 0 Then Inet2.FuncFtpConnect Nvl(rsTmp("Host2")), Nvl(rsTmp("User2")), Nvl(rsTmp("Pwd2"))
                            Call Inet2.FuncDownloadFile(objFile.GetParentFolderName(Nvl(rsTmp("Root2")) & rsTmp("URL")), strTmpFile, objFile.GetFileName(rsTmp("URL")))
                        End If
                    End If
                End If
                If Dir(strCachePath & Nvl(rsTmp("URL"))) <> vbNullString Then
                    Set curImage = dcmViewer.Images.ReadFile(strCachePath & Nvl(rsTmp("URL")))
                    With curImage
                        .BorderStyle = 6
                        .BorderWidth = 1
                        .BorderColour = vbWhite
                    End With
                    
                    'ȡ���Զ���Ӱ,��ΪDicomObjects�ؼ�����Դ����Ӱ��BUG�����ڣ�0028��6100��ʱ�����Զ���ͼ����м�Ӱ��
                    '���½�ú��DSAͼ����������ʾ
                    '��Ȼ����ͼ���mask=0 ,����ȡ����Ӱ������ÿ��ͼ����ӵ��µ�Dicomimages֮���Զ��ֽ�mask���ó�1�ˣ�
                    '�����ڳ������޷��ܺõĿ��ƣ����ֱ��ȥ����0028��6100��������ԡ�
                    If Not IsNull(curImage.Attributes(&H28, &H6100).Value) Then
                        curImage.Attributes.Remove &H28, &H6100
                    End If
                End If
                
                rsTmp.MoveNext
            Loop
            If dcmViewer.Images.Count > 0 Then
                dcmViewer.CurrentIndex = 1
                dcmViewer.Images(1).BorderColour = vbRed
            End If
        Else
            dcmViewer.MultiColumns = 1
            dcmViewer.MultiRows = 1
        End If
    Inet1.FuncFtpDisConnect
    Inet2.FuncFtpDisConnect
    Exit Sub
DBError:
    If ErrCenter() = 1 Then
        Resume
    End If
    
    Call SaveErrLog
End Sub

Public Function funGetStorageDevice(strSaveDeviceId As String, ByRef strDirURL As String, ByRef strIp As String, _
        ByRef strUser As String, ByRef strPwd As String) As Boolean
'------------------------------------------------
'���ܣ������ݿ��ж�ȡ�ƶ��洢�豸ID��FTP���ʲ���
'������ strSaveDeviceID �����洢�豸ID
'       strDirURL����[OUT] FTPĿ¼
'       strIp ����[OUT] IP��ַ
'       strUser ���� [OUT]�û���
'       strPwd ����[OUT]�û���
'���أ�True������ȡ�ɹ���False������ȡʧ��
'-----------------------------------------------
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    '���洢�豸�Ƿ����
    strSQL = "Select '/'||Decode(FtpĿ¼,Null,'',FtpĿ¼||'/') As URL,FTP�û���,FTP����,IP��ַ " & _
        "From Ӱ���豸Ŀ¼ " & "Where �豸��=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, strSaveDeviceId)
     'û�д洢�豸ʱ�˳�
    If rsTemp.EOF = True Then
        MsgBox "û���ҵ��洢�豸,������ѡ��洢�豸!", vbInformation, gstrSysName
        funGetStorageDevice = False
        Exit Function
    End If
    strDirURL = Nvl(rsTemp("URL"))
    strIp = Nvl(rsTemp("IP��ַ"))
    strUser = Nvl(rsTemp("FTP�û���"))
    strPwd = Nvl(rsTemp("FTP����"))
    funGetStorageDevice = True
End Function

Public Function OpenViewer(ByRef objPacsCore As Object, lngAdviceID As Long, _
        blnAddImage As Boolean, objParent As Object, Optional ByVal strSerials As String = "", _
        Optional ByVal blnMoved As Boolean = False, Optional ByVal blnLocalizerBackward As Boolean = False, _
        Optional ByVal intImageInterval As Integer = 0, Optional ByVal strImageString As String = "") As Boolean
'------------------------------------------------
'���ܣ����ݴ����ҽ��ID�ͷ��ͺţ���objPacsCoreָ��Ĺ�Ƭվ
'������ objPacsCore ������Ƭվ����
'       lngAdviceID ����ҽ��ID
'       blnAddImage--True ��ԭ��ͼ����������ӵ�ǰͼ��Falseɾ��ԭ��ͼ�񣬴򿪵�ǰͼ��
'       objParent -- ������
'       strSerials������ѡ������UID���ƴ����ö��ŷָ�����������룬��ѡ��ȫ������
'       blnMoved������ѡ���Ƿ�ת��
'       blnLocalizerBackward--��ѡ����λ�����,��strImageString����
'       intImageInterval ---��ѡ����ͼ��ļ��������5����ʾÿ5��ͼ��һ��ͼ,��strImageString����
'       strImageString --- ��ѡ��ÿ����������Ҫ�򿪵�ͼ�����ϣ���intImageInterval��blnLocalizerBackward���⣬
'                           ��strImageStringΪ��
'                           �����ǡ�����UID1|1-3;5-27;33-100+����UID2|ȫ����,ȫ����ʾ��ȫ��ͼ��
'���أ�ͼ���ļ���������
'------------------------------------------------
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim strFTPHost As String
    Dim strSDPath As String, strSDUser As String, strSDPwd As String
    Dim strDeviceNO As String
    Dim i As Integer
    Dim blnConnectDS As Boolean         '�Ƿ����ӵ�ǰ�Ĺ���Ŀ¼
    
    On Error GoTo DBError
    strFTPHost = ""
           
    '������Ҫ�򿪵�����ͼ����Ϣ
    strSQL = "Select /*+RULE*/ D.IP��ַ As Host1,d.�豸�� as �豸��1," & _
        "Decode(C.��������,Null,'',to_Char(C.��������,'YYYYMMDD')||'/')" & _
        "||C.���UID||'/' As Path,E.IP��ַ As Host2,e.�豸�� as �豸��2, " & _
        "D.����Ŀ¼ AS ����Ŀ¼1, E.����Ŀ¼ AS ����Ŀ¼2,D.����Ŀ¼�û��� as ����Ŀ¼�û���1, " & _
        "E.����Ŀ¼�û��� AS ����Ŀ¼�û���2,D.����Ŀ¼���� AS ����Ŀ¼����1,E.����Ŀ¼���� AS ����Ŀ¼����2 " & _
        "From Ӱ�����¼ C,Ӱ���豸Ŀ¼ D,Ӱ���豸Ŀ¼ E " & _
        "Where C.λ��һ=D.�豸��(+) And C.λ�ö�=E.�豸��(+) And C.ҽ��ID=[1] "
    
    '�����ת����־�����ȡת������ʷ��
    If blnMoved Then
        strSQL = Replace(strSQL, "Ӱ����ͼ��", "HӰ����ͼ��")
        strSQL = Replace(strSQL, "Ӱ��������", "HӰ��������")
        strSQL = Replace(strSQL, "Ӱ�����¼", "HӰ�����¼")
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ����Ŀ¼��Ϣ", lngAdviceID)
    
    If rsTmp.RecordCount > 0 Then
        '�������صĻ���Ŀ¼����Ҫ�ڵ��ù�Ƭվ֮ǰ�ȴ������Ŀ¼����Ƭվ��ֻ�����أ����������ػ���Ŀ¼
        MkLocalDir App.Path & "\TmpImage\" & rsTmp("Path")
        
        '��ȡFTP�����������û��������룬IP��ַ��
        If rsTmp("�豸��1") <> "" Then
            strDeviceNO = rsTmp("�豸��1")
            strFTPHost = rsTmp("Host1")
            strSDPath = Nvl(rsTmp("����Ŀ¼1"))
            strSDUser = Nvl(rsTmp("����Ŀ¼�û���1"))
            strSDPwd = Nvl(rsTmp("����Ŀ¼����1"))
        ElseIf Nvl(rsTmp("�豸��2")) <> "" Then
            strDeviceNO = rsTmp("�豸��2")
            strFTPHost = rsTmp("Host2")
            strSDPath = Nvl(rsTmp("����Ŀ¼2"))
            strSDUser = Nvl(rsTmp("����Ŀ¼�û���2"))
            strSDPwd = Nvl(rsTmp("����Ŀ¼����2"))
        End If
        
        '�жϹ���Ŀ¼�Ƿ��Ѿ����ӣ����û�����ӣ����������
        blnConnectDS = True
        For i = 1 To UBound(gConnectedShardDir)
            If gConnectedShardDir(i) = strDeviceNO Then
                blnConnectDS = False
                Exit For
            End If
        Next i
        If blnConnectDS = True And strSDPath <> "" Then
            If funcConnectShardDir("\\" & strFTPHost & "\" & strSDPath, strSDUser, strSDPwd) = 0 Then
                ReDim Preserve gConnectedShardDir(UBound(gConnectedShardDir) + 1) As String
                gConnectedShardDir(UBound(gConnectedShardDir)) = strDeviceNO
            End If
        End If
        
        If objPacsCore Is Nothing Then
            Exit Function
        Else
            objPacsCore.CallOpenViewer strImageString, lngAdviceID, objParent, gcnOracle, blnMoved, blnAddImage, intImageInterval
        End If
    Else    'û�в��ҵ�ͼ���¼����ر�ԭ���Ѿ��򿪵Ĺ�Ƭ����
        If Not objPacsCore Is Nothing Then objPacsCore.Closefrom
    End If
    
    OpenViewer = True
    Exit Function
DBError:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckChargeState(ByVal lngҽ��ID As Long, ByVal lng��Դ As Long) As Integer
'�жϵ�ǰ��ҽ���Ƿ��շ�
'һ��ҽ�����жಿλ����ҽ��

    Dim rsTemp As New ADODB.Recordset
    Dim strTable As String
    
    CheckChargeState = 0
    
    'סԺ���˲�סԺ���ü�¼���������Ȳ��˲�������ü�¼
    If lng��Դ = 2 Then
        strTable = "סԺ���ü�¼"
    Else
        strTable = "������ü�¼"
    End If
    
    gstrSQL = "Select A.ҽ��id, B.��¼״̬" & vbNewLine & _
                "From ����ҽ������ A, " & strTable & " B" & vbNewLine & _
                "Where A.ҽ��id = [1] And A.NO = B.NO And A.��¼���� = B.��¼����"
                
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�Ƿ��շ�", lngҽ��ID)
    
    '�ɷѵļ��������1 û��¼ ��2 �м�¼ȫ��Ϊ��¼״̬=1 ��3 �м�¼���Ҳ��ݼ�¼״̬<>1����ʾ���˷ѻ��в���δ��
    'ֻ��2����������ѽɷѣ�ȫ�ɣ�
    If rsTemp.BOF Then Exit Function
    Do Until rsTemp.EOF
        If Nvl(rsTemp!��¼״̬, 0) <> 1 Then Exit Function
        rsTemp.MoveNext
    Loop
    CheckChargeState = 1
End Function
Public Function CheckConcurrentReport(ByVal lngOrderID As Long, Optional blnSilence As Boolean = False) As Boolean
'���ܣ���鵱ǰ�����Ƿ���ҽ�����ڲ�������
Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    CheckConcurrentReport = True
    gstrSQL = "Select ������� From Ӱ�����¼ Where ҽ��ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��¼", lngOrderID)
    
    If Not rsTemp Is Nothing Then
        If Not rsTemp.EOF Then
            If Nvl(rsTemp!�������) <> "" And Nvl(rsTemp!�������) <> UserInfo.���� Then
                If blnSilence = False Then
                    MsgBox "��ǰ���˵ı������ڱ� " & Nvl(rsTemp!�������) & " ���������Ժ����ԡ�", vbInformation, gstrSysName
                End If
                CheckConcurrentReport = False
            End If
        End If
    End If
    Exit Function
    
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Public Sub UpdateReporter(ByVal lngOrderID As Long, ByVal Reporter As String)
    On Error GoTo errHandle
    
    gstrSQL = "ZL_Ӱ�񱨸����_Update(" & lngOrderID & ",'" & Reporter & "')"
    zlDatabase.ExecuteProcedure gstrSQL, "���²�����"
    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Public Function bln����δ�󻮼۵�(ByVal lngҽ��ID As Long, ByVal lng��Դ As Long) As Boolean
    Dim rsTemp As ADODB.Recordset
    Dim strFeeTable As String
    
    'סԺ���˲�סԺ���ü�¼���������Ȳ��˲�������ü�¼
    If lng��Դ = 2 Then
        strFeeTable = "סԺ���ü�¼"
    Else
        strFeeTable = "������ü�¼"
    End If

    On Error GoTo errHandle
    gstrSQL = "Select /*+ RULE */ A.ID" & vbNewLine & _
            "From " & strFeeTable & " A" & vbNewLine & _
            "Where A.ҽ����� + 0 In (Select ID From ����ҽ����¼ Where ID = [1] Or ���id = [1]) And (A.��¼����, A.NO) In" & vbNewLine & _
            "      (Select ��¼����, NO" & vbNewLine & _
            "       From ����ҽ������" & vbNewLine & _
            "       Where ҽ��id In (Select ID From ����ҽ����¼ Where ID = [1] Or ���id = [1])" & vbNewLine & _
            "       Union All" & vbNewLine & _
            "       Select ��¼����, NO" & vbNewLine & _
            "       From ����ҽ������" & vbNewLine & _
            "       Where ҽ��id In (Select ID From ����ҽ����¼ Where ID = [1] Or ���id = [1])" & vbNewLine & _
            "       ) And A.���ʷ��� = 1 And A.��¼״̬ = 0"

    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡδ�󻮼۵�", lngҽ��ID)
    If rsTemp.EOF Then
        Exit Function
    Else
        bln����δ�󻮼۵� = True
    End If
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Public Function bln������Ժ(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
Dim rsTemp As ADODB.Recordset
    On Error GoTo errHandle
    gstrSQL = "SELECT to_char(��Ժ����,'YYYY-MM-DD HH24:MI:SS') as ��Ժ���� from ������ҳ where ����ID=[1] AND ��ҳID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��Ժʱ��", lng����ID, lng��ҳID)
    If rsTemp.EOF Then
        Exit Function
    Else
        If Nvl(rsTemp!��Ժ����) = "" Then
            bln������Ժ = True
        End If
    End If
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function GetFullPY(strIn As String) As String
'------------------------------------------------
'���ܣ��Ѵ�����ַ����а���������ת����ƴ����Ӣ����ĸ�����ֲ�������
'������ strIn ����������ַ���
'���أ��Ѻ���ת����ƴ������ַ���
'------------------------------------------------
    Dim i As Integer
    Dim strChar As String
    
    strIn = Trim(strIn)
    For i = 1 To Len(strIn)
        strChar = Mid(strIn, i, 1)
        If Asc(strChar) < 0 Then
            GetFullPY = GetFullPY & UCase(Replace(zlCommFun.mGetFullPY(strChar), vbCrLf, "")) & " "
        Else
            GetFullPY = GetFullPY & strChar
        End If
    Next i
    GetFullPY = Trim(GetFullPY)
End Function

Public Function GetRptImages(ByRef RptViewer As DicomViewer, ByVal lngOrderID As Long, ByVal blnMoved As Boolean)
'------------------------------------------------
'���ܣ���ȡ����ͼ�񵽱��أ���ˢ����ʾ
'������ RptViewer ������ʾͼ��Ŀؼ�
'       lngOrderID -- ҽ��ID
'       blnMoved -- �Ƿ�ת��
'���أ��ޣ�ֱ����RptViewer �м���ͼ��
'------------------------------------------------
    Dim objFileSystem As New Scripting.FileSystemObject
    Dim strSQL As String, rsTemp As New ADODB.Recordset
    Dim aryFiles() As String    '����ͼ������
    Dim strFiles As String      '���ֺŷָ��ĳɹ����ص��ļ�
    Dim aryRptFileName() As String    '�����ļ�������
    
    Dim cFtpNet As New cFTP
    Dim strVirtualPath As String
    Dim strLocalPath As String
    Dim IntCount As Integer
    Dim curImage As DicomImage
    
    '�����RptViewer �е�ͼ��
    RptViewer.Images.Clear
    
    '��鱾�ػ���ͼ��ĸ�Ŀ¼�Ƿ���ڣ�����������򴴽����ظ�Ŀ¼���������ʧ�ܣ���ֱ���˳�����
    If objFileSystem.FolderExists(App.Path & "\TmpImage\") = False Then objFileSystem.CreateFolder App.Path & "\TmpImage\"
    If objFileSystem.FolderExists(App.Path & "\TmpImage\") = False Then GetRptImages = False: Exit Function
    
    '�����ݿ��ȡͼ����Դ��Ϣ
    err = 0: On Error Resume Next
    strSQL = "Select To_Char(L.��������, 'yyyymmdd') As ��Ŀ¼, L.���uid, L.����ͼ��, A1.FtpĿ¼ As Root1, A1.Ip��ַ As Ip1," & vbNewLine & _
            "       A1.FTP�û��� As Usr1, A1.FTP���� As Pwd1, A2.FtpĿ¼ As Root2, A2.Ip��ַ As Ip2, A2.FTP�û��� As Usr2, A2.FTP���� As Pwd2" & vbNewLine & _
            "From Ӱ�����¼ L, Ӱ���豸Ŀ¼ A1, Ӱ���豸Ŀ¼ A2" & vbNewLine & _
            "Where L.λ��һ = A1.�豸��(+) And L.λ�ö� = A2.�豸��(+) And L.ҽ��id = [1]"
    If blnMoved = True Then
        strSQL = Replace(strSQL, "Ӱ�����¼", "HӰ�����¼")
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ����ͼ��", lngOrderID)
    If rsTemp.RecordCount <= 0 Then GetRptImages = False: Exit Function
    aryFiles = Split("" & rsTemp!����ͼ��, ";")
    aryRptFileName = Split("" & rsTemp!����ͼ��, ";")
    If UBound(aryFiles) < 0 Then GetRptImages = False: Exit Function
        
    '��鱾�������б��μ���Ŀ¼�Ƿ���ڣ�����������򴴽����ش洢Ŀ¼���������ʧ�ܣ����˳�����
    err = 0: On Error Resume Next
    strLocalPath = App.Path & "\TmpImage\" & rsTemp!��Ŀ¼
    If objFileSystem.FolderExists(strLocalPath) = False Then objFileSystem.CreateFolder strLocalPath
    If objFileSystem.FolderExists(strLocalPath) = False Then GetRptImages = False: Exit Function
    strLocalPath = strLocalPath & "\" & rsTemp!���uid
    If objFileSystem.FolderExists(strLocalPath) = False Then objFileSystem.CreateFolder strLocalPath
    If objFileSystem.FolderExists(strLocalPath) = False Then GetRptImages = False: Exit Function
        
    strFiles = ""
    '��鱾�ػ���ͼ���Ƿ���ڣ�������ڣ��򲻴�FTP���أ���ֱ�Ӷ�ȡ��������ͼ��
    For IntCount = 0 To UBound(aryFiles)
        '����ļ����ڣ�����Ҫ���أ����ñ��
        If Dir(strLocalPath & "\" & Trim(aryFiles(IntCount))) <> "" Then
            strFiles = strFiles & ";" & strLocalPath & "\" & Trim(aryFiles(IntCount))
            aryFiles(IntCount) = ""
        End If
    Next IntCount
    If strFiles <> "" Then strFiles = Mid(strFiles, 2)
    
    
    '������δ��ڵ��ļ���������Ҫ�򿪵��ļ�������һ�£����FTP���ر��������ڵ�ͼ��
    If UBound(Split(strFiles, ";")) <> UBound(aryFiles) Then
        '���ȴ��豸1����ͼ��
        If "" & rsTemp!Ip1 <> "" Then
            If cFtpNet.FuncFtpConnect("" & rsTemp!Ip1, "" & rsTemp!Usr1, "" & rsTemp!pwd1) <> 0 Then
                strVirtualPath = rsTemp!Root1 & "/" & rsTemp!��Ŀ¼ & "/" & rsTemp!���uid
                For IntCount = 0 To UBound(aryFiles)
                    If aryFiles(IntCount) <> "" Then
                        If cFtpNet.FuncDownloadFile(strVirtualPath, strLocalPath & "\" & Trim(aryFiles(IntCount)), Trim(aryFiles(IntCount))) = 0 Then
                            If Dir(strLocalPath & "\" & Trim(aryFiles(IntCount))) <> "" Then
                                strFiles = strFiles & ";" & strLocalPath & "\" & Trim(aryFiles(IntCount))
                                aryFiles(IntCount) = ""
                            End If
                        End If
                    End If
                Next IntCount
            End If
            cFtpNet.FuncFtpDisConnect
        End If
        
        '����豸1����ͼ���������ٴ��豸2����ͼ��
        If strFiles <> "" Then strFiles = Mid(strFiles, 2)
        If UBound(Split(strFiles, ";")) <> UBound(aryFiles) And "" & rsTemp!Ip2 <> "" Then
            If cFtpNet.FuncFtpConnect("" & rsTemp!Ip2, "" & rsTemp!Usr2, "" & rsTemp!pwd2) <> 0 Then
                strVirtualPath = rsTemp!Root2 & "/" & rsTemp!��Ŀ¼ & "/" & rsTemp!���uid
                For IntCount = 0 To UBound(aryFiles)
                    If aryFiles(IntCount) <> "" Then
                        If cFtpNet.FuncDownloadFile(strVirtualPath, strLocalPath & "\" & Trim(aryFiles(IntCount)), Trim(aryFiles(IntCount))) = 0 Then
                            If Dir(strLocalPath & "\" & Trim(aryFiles(IntCount))) <> "" Then
                                strFiles = strFiles & ";" & strLocalPath & "\" & Trim(aryFiles(IntCount))
                            End If
                        End If
                    End If
                Next IntCount
            End If
            cFtpNet.FuncFtpDisConnect
        End If
        If strFiles <> "" Then
            If Left(strFiles, 1) = ";" Then strFiles = Mid(strFiles, 2)
        End If
    End If
    
    '����õ��ļ�װ��
    Dim iRows As Integer, iCols As Integer
    aryFiles = Split(strFiles, ";")
    With RptViewer
        For IntCount = 0 To UBound(aryFiles)
            Set curImage = New DicomImage
            curImage.FileImport aryFiles(IntCount), "JPG"
            curImage.BorderWidth = 3: curImage.BorderColour = vbWhite
            curImage.Tag = aryRptFileName(IntCount)
            .Images.Add curImage
        Next
        If UBound(aryFiles) >= 0 Then
            .CurrentIndex = 1
            .Images(.CurrentIndex).BorderColour = vbBlue
        End If
        If .Images.Count > 0 Then
            ResizeRegion .Images.Count, .Width, .Height, iRows, iCols
            .MultiColumns = iCols: .MultiRows = iRows
        Else
            '��������
        End If
    End With
    
    GetRptImages = True: Exit Function

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

'Public Sub PromptResult(lngOrderID As Long, lngModul As Long, frmParent As Form)
'    Dim strResult As String
'
'    strResult = frmResult.zlGetResult(frmParent, lngModul, lngOrderID)    '��ʾ���������Ժ�Ӱ������
'    If strResult = "" Then Exit Sub
'
'    If Split(strResult, "-")(0) = 1 Then    '������
'        gstrSQL = "ZL_Ӱ����_���(" & lngOrderID & ",1)"
'    Else
'        gstrSQL = "ZL_Ӱ����_���(" & lngOrderID & ",0)"
'    End If
'    zlDatabase.ExecuteProcedure gstrSQL, "���������"
'
'    If lngModul = 1290 Then         'Ӱ��ҽ��վ�ż�¼Ӱ������
'        If Split(strResult, "-")(1) = 1 Then    'Ӱ������
'            gstrSQL = "Zl_Ӱ������_Update(" & lngOrderID & ",'��')"
'        Else
'            gstrSQL = "Zl_Ӱ������_Update(" & lngOrderID & ",'��')"
'        End If
'        zlDatabase.ExecuteProcedure gstrSQL, "Ӱ������"
'    End If
'End Sub
'Public Function BillingWarn(frmParent As Object, ByVal strPrivs As String, _
'    rsWarn As ADODB.Recordset, ByVal str���� As String, ByVal curʣ���� As Currency, _
'    ByVal cur���ս�� As Currency, ByVal cur���ʽ�� As Currency, ByVal cur������� As Currency, _
'    ByVal str�շ���� As String, ByVal str������� As String, str�ѱ���� As String, _
'    intWarn As Integer, Optional ByVal bln���� As Boolean) As Integer
''����:�Բ��˼��ʽ��б�����ʾ
''����:rsWarn=���������������õļ�¼��(�ò��˲���,�����ֺ���ҽ��)
''     str�շ����=��ǰҪ�������,���ڷ��౨��
''     str�������=�������,������ʾ
''     bln����=���ɻ��۷���ʱ�ı��������ƾ���ǿ�Ƽ���Ȩ��ʱ�Ĵ���
''     intWarn=�Ƿ���ʾѯ���Ե���ʾ,-1=Ҫ��ʾ,0=ȱʡΪ��,1-ȱʡΪ��
''����:str�ѱ����="CDE":�����ڱ��α�����һ�����,"-"Ϊ������𡣸÷������ڴ����ظ�����
''     intWarn=����ѯ������ʾ�е�ѡ����,0=Ϊ��,1-Ϊ��
''     0;û�б���,����
''     1:������ʾ���û�ѡ�����
''     2:������ʾ���û�ѡ���ж�
''     3:������ʾ�����ж�
''     4:ǿ�Ƽ��ʱ���,����
'    Dim bln�ѱ��� As Boolean, byt��־ As Byte
'    Dim byt��ʽ As Byte, byt�ѱ���ʽ As Byte
'    Dim arrtmp As Variant, vMsg As VbMsgBoxResult
'    Dim str���� As String, i As Long
'
'    BillingWarn = 0
'
'    '�����������:NULL��û������,0�������˵�
'    If rsWarn.State = 0 Then Exit Function
'    If rsWarn.EOF Then Exit Function
'    If IsNull(rsWarn!����ֵ) Then Exit Function
'
'    '��Ӧ���λ��Ч��������
'    If Not IsNull(rsWarn!������־1) Then
'        If rsWarn!������־1 = "-" Or InStr(rsWarn!������־1, str�շ����) > 0 Then byt��־ = 1
'        If rsWarn!������־1 = "-" Then str������� = "" '�������ʱ,������ʾ��������
'    End If
'    If byt��־ = 0 And Not IsNull(rsWarn!������־2) Then
'        If rsWarn!������־2 = "-" Or InStr(rsWarn!������־2, str�շ����) > 0 Then byt��־ = 2
'        If rsWarn!������־2 = "-" Then str������� = "" '�������ʱ,������ʾ��������
'    End If
'    If byt��־ = 0 And Not IsNull(rsWarn!������־3) Then
'        If rsWarn!������־3 = "-" Or InStr(rsWarn!������־3, str�շ����) > 0 Then byt��־ = 3
'        If rsWarn!������־3 = "-" Then str������� = "" '�������ʱ,������ʾ��������
'    End If
'    If byt��־ = 0 Then Exit Function '����Ч����
'
'    '������־2ʵ�����������жϢ٢�,����ֻ��һ���жϢ�
'    '���ִ����ǰ����һ�����ֻ������һ�ֱ�����ʽ(������������ʱ)
'    'ʾ����"-" �� ",ABC,567,DEF"
'    '������־2ʾ����"-��" �� ",ABC��,567��,DEF��"
'    bln�ѱ��� = InStr(str�ѱ����, str�շ����) > 0 Or str�ѱ���� Like "-*"
'
'    If bln�ѱ��� Then '��intWarn = -1ʱ,Ҳ��ǿ���ٱ���
'        If byt��־ = 2 Then
'            If str�ѱ���� Like "-*" Then
'                byt�ѱ���ʽ = IIf(Right(str�ѱ����, 1) = "��", 2, 1)
'            Else
'                arrtmp = Split(str�ѱ����, ",")
'                For i = 0 To UBound(arrtmp)
'                    If InStr(arrtmp(i), str�շ����) > 0 Then
'                        byt�ѱ���ʽ = IIf(Right(arrtmp(i), 1) = "��", 2, 1)
'                        'Exit For 'ȡ��˵����סԺ����ģ��
'                    End If
'                Next
'            End If
'        Else
'            Exit Function
'        End If
'    End If
'
'    If str������� <> "" Then str������� = """" & str������� & """����"
'    str���� = IIf(cur������� = 0, "", "(��������:" & Format(cur�������, "0.00") & ")")
'    curʣ���� = curʣ���� + cur������� - cur���ʽ��
'    cur���ս�� = cur���ս�� + cur���ʽ��
'
'    '---------------------------------------------------------------------
'    If rsWarn!�������� = 1 Then  '�ۼƷ��ñ���(����)
'        Select Case byt��־
'            Case 1 '���ڱ���ֵ(����Ԥ����ľ�)��ʾѯ�ʼ���
'                If curʣ���� < rsWarn!����ֵ Then
'                    If Not (InStr(";" & strPrivs & ";", ";ǿ�Ƽ���;") > 0 Or bln����) Then
'                        If intWarn = -1 Then
'                            vMsg = frmMsgBox.ShowMsgBox(str���� & " ��ǰʣ���" & str���� & ":" & Format(curʣ����, "0.00") & ",����" & str������� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & ",����ò��˼�����", frmParent)
'                            If vMsg = vbNo Or vMsg = vbCancel Then
'                                If vMsg = vbCancel Then intWarn = 0
'                                BillingWarn = 2
'                            ElseIf vMsg = vbYes Or vMsg = vbIgnore Then
'                                If vMsg = vbIgnore Then intWarn = 1
'                                BillingWarn = 1
'                            End If
'                        Else
'                            If intWarn = 0 Then
'                                BillingWarn = 2
'                            ElseIf intWarn = 1 Then
'                                BillingWarn = 1
'                            End If
'                        End If
'                    Else
'                        If intWarn = -1 Then
'                            vMsg = frmMsgBox.ShowMsgBox(IIf(bln����, "", "ǿ�Ƽ���") & "����:" & vbCrLf & vbCrLf & str���� & " ��ǰʣ���" & str���� & ":" & Format(curʣ����, "0.00") & " ����" & str������� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & "��", frmParent, True)
'                            If vMsg = vbIgnore Then intWarn = 1
'                        End If
'                        BillingWarn = 4
'                    End If
'                End If
'            Case 2 '���ڱ���ֵ��ʾѯ�ʼ���,Ԥ����ľ�ʱ��ֹ����
'                If Not bln�ѱ��� Then
'                    If curʣ���� < 0 Then
'                        byt��ʽ = 2
'                        If Not (InStr(";" & strPrivs & ";", ";ǿ�Ƽ���;") > 0 Or bln����) Then
'                            If intWarn = -1 Then
'                                vMsg = frmMsgBox.ShowMsgBox(str���� & " ��ǰʣ���" & str���� & "�Ѿ��ľ�," & str������� & "��ֹ���ʡ�", frmParent, True)
'                                If vMsg = vbIgnore Then intWarn = 1
'                            End If
'                            BillingWarn = 3
'                        Else
'                            If intWarn = -1 Then
'                                vMsg = frmMsgBox.ShowMsgBox(str������� & IIf(bln����, "", "ǿ�Ƽ���") & "����:" & vbCrLf & vbCrLf & str���� & " ��ǰʣ���" & str���� & "�Ѿ��ľ���", frmParent, True)
'                                If vMsg = vbIgnore Then intWarn = 1
'                            End If
'                            BillingWarn = 4
'                        End If
'                    ElseIf curʣ���� < rsWarn!����ֵ Then
'                        byt��ʽ = 1
'                        If Not (InStr(";" & strPrivs & ";", ";ǿ�Ƽ���;") > 0 Or bln����) Then
'                            If intWarn = -1 Then
'                                vMsg = frmMsgBox.ShowMsgBox(str���� & " ��ǰʣ���" & str���� & ":" & Format(curʣ����, "0.00") & ",����" & str������� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & ",����ò��˼�����", frmParent)
'                                If vMsg = vbNo Or vMsg = vbCancel Then
'                                    If vMsg = vbCancel Then intWarn = 0
'                                    BillingWarn = 2
'                                ElseIf vMsg = vbYes Or vMsg = vbIgnore Then
'                                    If vMsg = vbIgnore Then intWarn = 1
'                                    BillingWarn = 1
'                                End If
'                            Else
'                                If intWarn = 0 Then
'                                    BillingWarn = 2
'                                ElseIf intWarn = 1 Then
'                                    BillingWarn = 1
'                                End If
'                            End If
'                        Else
'                            If intWarn = -1 Then
'                                vMsg = frmMsgBox.ShowMsgBox(IIf(bln����, "", "ǿ�Ƽ���") & "����:" & vbCrLf & vbCrLf & str���� & " ��ǰʣ���" & str���� & ":" & Format(curʣ����, "0.00") & ",����" & str������� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & "��", frmParent, True)
'                                If vMsg = vbIgnore Then intWarn = 1
'                            End If
'                            BillingWarn = 4
'                        End If
'                    End If
'                Else
'                    '�ϴ��ѱ�����ѡ�������ǿ�Ƽ���
'                    If byt�ѱ���ʽ = 1 Then
'                        '�ϴε��ڱ���ֵ��ѡ�������ǿ�Ƽ���,���ٴ�����ڵ����,������Ҫ�ж�Ԥ�����Ƿ�ľ�
'                        If curʣ���� < 0 Then
'                            byt��ʽ = 2
'                            If Not (InStr(";" & strPrivs & ";", ";ǿ�Ƽ���;") > 0 Or bln����) Then
'                                If intWarn = -1 Then
'                                    vMsg = frmMsgBox.ShowMsgBox(str���� & " ��ǰʣ���" & str���� & "�Ѿ��ľ�," & str������� & "��ֹ���ʡ�", frmParent, True)
'                                    If vMsg = vbIgnore Then intWarn = 1
'                                End If
'                                BillingWarn = 3
'                            Else
'                                If intWarn = -1 Then
'                                    vMsg = frmMsgBox.ShowMsgBox(str������� & IIf(bln����, "", "ǿ�Ƽ���") & "����:" & vbCrLf & vbCrLf & str���� & " ��ǰʣ���" & str���� & "�Ѿ��ľ���", frmParent, True)
'                                    If vMsg = vbIgnore Then intWarn = 1
'                                End If
'                                BillingWarn = 4
'                            End If
'                        End If
'                    ElseIf byt�ѱ���ʽ = 2 Then
'                        '�ϴ�Ԥ�����Ѿ��ľ���ǿ�Ƽ���,���ٴ���
'                        Exit Function
'                    End If
'                End If
'            Case 3 '���ڱ���ֵ��ֹ����
'                If curʣ���� < rsWarn!����ֵ Then
'                    If Not (InStr(";" & strPrivs & ";", ";ǿ�Ƽ���;") > 0 Or bln����) Then
'                        If intWarn = -1 Then
'                            vMsg = frmMsgBox.ShowMsgBox(str���� & " ��ǰʣ���" & str���� & ":" & Format(curʣ����, "0.00") & ",����" & str������� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & ",��ֹ���ʡ�", frmParent, True)
'                            If vMsg = vbIgnore Then intWarn = 1
'                        End If
'                        BillingWarn = 3
'                    Else
'                        If intWarn = -1 Then
'                            vMsg = frmMsgBox.ShowMsgBox(IIf(bln����, "", "ǿ�Ƽ���") & "����:" & vbCrLf & vbCrLf & str���� & " ��ǰʣ���" & str���� & ":" & Format(curʣ����, "0.00") & ",����" & str������� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & "��", frmParent, True)
'                            If vMsg = vbIgnore Then intWarn = 1
'                        End If
'                        BillingWarn = 4
'                    End If
'                End If
'        End Select
'    ElseIf rsWarn!�������� = 2 Then  'ÿ�շ��ñ���(����)
'        Select Case byt��־
'            Case 1 '���ڱ���ֵ��ʾѯ�ʼ���
'                If cur���ս�� > rsWarn!����ֵ Then
'                    If Not (InStr(";" & strPrivs & ";", ";ǿ�Ƽ���;") > 0 Or bln����) Then
'                        If intWarn = -1 Then
'                            vMsg = frmMsgBox.ShowMsgBox(str���� & " ���շ���:" & Format(cur���ս��, gstrDec) & ",����" & str������� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & ",����ò��˼�����", frmParent)
'                            If vMsg = vbNo Or vMsg = vbCancel Then
'                                If vMsg = vbCancel Then intWarn = 0
'                                BillingWarn = 2
'                            ElseIf vMsg = vbYes Or vMsg = vbIgnore Then
'                                If vMsg = vbIgnore Then intWarn = 1
'                                BillingWarn = 1
'                            End If
'                        Else
'                            If intWarn = 0 Then
'                                BillingWarn = 2
'                            ElseIf intWarn = 1 Then
'                                BillingWarn = 1
'                            End If
'                        End If
'                    Else
'                        If intWarn = -1 Then
'                            vMsg = frmMsgBox.ShowMsgBox(IIf(bln����, "", "ǿ�Ƽ���") & "����:" & vbCrLf & vbCrLf & str���� & " ���շ���:" & Format(cur���ս��, gstrDec) & ",����" & str������� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & "��", frmParent, True)
'                            If vMsg = vbIgnore Then intWarn = 1
'                        End If
'                        BillingWarn = 4
'                    End If
'                End If
'            Case 3 '���ڱ���ֵ��ֹ����
'                If cur���ս�� > rsWarn!����ֵ Then
'                    If Not (InStr(";" & strPrivs & ";", ";ǿ�Ƽ���;") > 0 Or bln����) Then
'                        If intWarn = -1 Then
'                            vMsg = frmMsgBox.ShowMsgBox(str���� & " ���շ���:" & Format(cur���ս��, gstrDec) & ",����" & str������� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & ",��ֹ���ʡ�", frmParent, True)
'                            If vMsg = vbIgnore Then intWarn = 1
'                        End If
'                        BillingWarn = 3
'                    Else
'                        If intWarn = -1 Then
'                            vMsg = frmMsgBox.ShowMsgBox(IIf(bln����, "", "ǿ�Ƽ���") & "����:" & vbCrLf & vbCrLf & str���� & " ���շ���:" & Format(cur���ս��, gstrDec) & ",����" & str������� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & "��", frmParent, True)
'                            If vMsg = vbIgnore Then intWarn = 1
'                        End If
'                        BillingWarn = 4
'                    End If
'                End If
'        End Select
'    End If
'
'    '���ڼ�����Ĳ���,�����ѱ������
'    If BillingWarn = 1 Or BillingWarn = 4 Then
'        If byt��־ = 1 Then
'            If rsWarn!������־1 = "-" Then
'                str�ѱ���� = "-"
'            Else
'                str�ѱ���� = str�ѱ���� & "," & rsWarn!������־1
'            End If
'        ElseIf byt��־ = 2 Then
'            If rsWarn!������־2 = "-" Then
'                str�ѱ���� = "-"
'            Else
'                str�ѱ���� = str�ѱ���� & "," & rsWarn!������־2
'            End If
'            '���ӱ�ע���ж��ѱ����ľ��巽ʽ
'            str�ѱ���� = str�ѱ���� & IIf(byt��ʽ = 2, "��", "��")
'        ElseIf byt��־ = 3 Then
'            If rsWarn!������־3 = "-" Then
'                str�ѱ���� = "-"
'            Else
'                str�ѱ���� = str�ѱ���� & "," & rsWarn!������־3
'            End If
'        End If
'    End If
'End Function

'Public Function FinishBillingWarn(ByVal frmParent As Object, ByVal strPrivs As String, ByVal lng����ID As Long, _
'    ByVal lng��ҳID As Long, ByVal lng����ID As Long, ByVal cur��� As Currency, ByVal str��� As String, ByVal str����� As String) As Boolean
''���ܣ���ִ��������Զ���˵ķ���ʱ���Բ��˷��ý��м��ʱ�����
''������str���="CDE..."����������漰�����շ����
''      str�����="���,����,..."����Ӧ�������������ʾ
'    Dim rsPati As ADODB.Recordset
'    Dim rsWarn As ADODB.Recordset
'    Dim strWarn As String, intWarn As Integer
'    Dim strSQL As String, intR As Integer, i As Long
'    Dim cur���� As Currency
'
'    On Error GoTo errH
'
'    If lng��ҳID <> 0 Then
'        'סԺ���˱���
'        strSQL = _
'            " Select ����ID,Ԥ�����,�������,0 as Ԥ����� From ������� Where ����=1 And ����ID=[1]" & _
'            " Union ALL" & _
'            " Select A.����ID,0,0,Sum(���) From ����ģ����� A,������ҳ B" & _
'            " Where A.����ID=B.����ID And A.��ҳID=B.��ҳID And B.���� Is Not Null And A.����ID=[1] And A.��ҳID=[2] Group by A.����ID"
'        strSQL = "Select ����ID,Nvl(Sum(Ԥ�����),0)-Nvl(Sum(�������),0)+Nvl(Sum(Ԥ�����),0) as ʣ��� From (" & strSQL & ") Group by ����ID"
'
'        strSQL = "Select A.����,zl_PatiWarnScheme(A.����ID,B.��ҳID) as ���ò���,C.ʣ���," & _
'            " Decode(A.������,Null,Null,zl_PatientSurety(A.����ID,B.��ҳID)) as ������" & _
'            " From ������Ϣ A,������ҳ B,(" & strSQL & ") C" & _
'            " Where A.����ID=B.����ID And A.����ID=C.����ID(+)" & _
'            " And A.����ID=[1] And B.��ҳID=[2]"
'        Set rsPati = zlDatabase.OpenSQLRecord(strSQL, "FinishBillingWarn", lng����ID, lng��ҳID)
'    Else
'        '���������ﱨ��
'        strSQL = "Select ����ID,Ԥ�����,������� From ������� Where ����=1 And ����ID=[1]"
'        strSQL = "Select A.����,zl_PatiWarnScheme(A.����ID) as ���ò���,A.������," & _
'            " Nvl(B.Ԥ�����,0)-Nvl(B.�������,0)+Nvl(E.�ʻ����,0) as ʣ���" & _
'            " From ������Ϣ A,(" & strSQL & ") B,ҽ�����˹����� D,ҽ�����˵��� E" & _
'            " Where A.����ID=B.����ID(+) And A.����id = D.����id(+) And A.����=D.����(+)" & _
'            " And D.����=E.����(+) And D.����=E.����(+) And D.ҽ����=E.ҽ����(+) And D.��־(+)=1" & _
'            " And A.����ID=[1]"
'        Set rsPati = zlDatabase.OpenSQLRecord(strSQL, "FinishBillingWarn", lng����ID)
'    End If
'
'    intWarn = -1 '���ʱ���ʱȱʡҪ��ʾ
'    'ִ�б���:���ﲡ�˲���ID=0
'    strSQL = "Select Nvl(��������,1) as ��������,����ֵ,������־1,������־2,������־3 From ���ʱ����� Where Nvl(����ID,0)=[1] And ���ò���=[2]"
'    Set rsWarn = zlDatabase.OpenSQLRecord(strSQL, "FinishBillingWarn", lng����ID, CStr(Nvl(rsPati!���ò���)))
'    If Not rsWarn.EOF Then
'        If rsWarn!�������� = 2 Then cur���� = GetPatiDayMoney(lng����ID)
'        str����� = Mid(str�����, 2)
'        For i = 1 To Len(str���)
'            intR = BillingWarn(frmParent, strPrivs, rsWarn, Nvl(rsPati!����), Nvl(rsPati!ʣ���, 0), cur����, cur���, Nvl(rsPati!������, 0), Mid(str���, i, 1), Split(str�����, ",")(i - 1), strWarn, intWarn)
'            If InStr(",2,3,", intR) > 0 Then Exit Function
'        Next
'    End If
'
'    FinishBillingWarn = True
'    Exit Function
'errH:
'    If ErrCenter() = 1 Then Resume
'    Call SaveErrLog
'End Function

Public Function GetAdviceMoney(ByVal lngAdviceID As Long, ByVal lng��Դ As Long, str��� As String, str����� As String) As Currency
'���ܣ�����ָ����ҽ��ID������ȡҽ����Ӧδ��˵ļ��ʷ��úϼ�
'������lngAdviceID,strSendNo
'���أ�str���,str�����=���ڱ�����ʾ
'˵������ϵͳ����Ϊִ�к���˷���ʱ�ŷ��ء�
    Dim rsTmp As New ADODB.Recordset
    Dim curMoney As Currency
    Dim strFeeTable As String
    
    str��� = "": str����� = ""
    
    On Error GoTo errH
    
    '��Ҫ����ϵͳ�����жϣ�81�Ų�����"ִ�к��Զ���˻��۵�"
    If zlDatabase.GetPara(81, glngSys) <> "1" Then Exit Function
    
    'סԺ���˲�סԺ���ü�¼���������Ȳ��˲�������ü�¼
    If lng��Դ = 2 Then
        strFeeTable = "סԺ���ü�¼"
    Else
        strFeeTable = "������ü�¼"
    End If
    
    gstrSQL = "Select /*+ RULE */" & vbNewLine & _
                " B.����, B.����, Sum(A.ʵ�ս��) As ���" & vbNewLine & _
                "From " & strFeeTable & " A, �շ���Ŀ��� B" & vbNewLine & _
                "Where A.ҽ����� + 0 In (Select ID From ����ҽ����¼ Where ID = [1] Or ���id = [1]) And" & vbNewLine & _
                "      (A.��¼����, A.NO) In" & vbNewLine & _
                "      ( Select ��¼����, NO" & vbNewLine & _
                "        From ����ҽ������" & vbNewLine & _
                "        Where ҽ��id In (Select ID From ����ҽ����¼ Where ID = [1] Or ���id = [1])" & vbNewLine & _
                "        Union All" & vbNewLine & _
                "        Select ��¼����, NO" & vbNewLine & _
                "        From ����ҽ������" & vbNewLine & _
                "        Where ҽ��id In (Select ID From ����ҽ����¼ Where ID = [1] Or ���id = [1] )" & vbNewLine & _
                "       ) And A.���ʷ��� = 1 And A.��¼״̬ = 0 And A.�շ���� = B.���� " & vbNewLine & _
                "Group By B.����, B.����"
                
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "GetAdviceMoney", lngAdviceID)
    
    curMoney = 0
    Do While Not rsTmp.EOF
        curMoney = curMoney + Nvl(rsTmp!���, 0)
        str��� = str��� & rsTmp!����
        str����� = str����� & "," & rsTmp!����
        rsTmp.MoveNext
    Loop
    
    str����� = Mid(str�����, 2)
    GetAdviceMoney = curMoney
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Public Function GetPatiDayMoney(lng����ID As Long) As Currency
'���ܣ���ȡָ�����˵��췢���ķ����ܶ�
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select zl_PatiDayCharge([1]) as ��� From Dual"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng����ID)
    If Not rsTmp.EOF Then GetPatiDayMoney = Nvl(rsTmp!���, 0)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function funcConnectShardDir(strShareRemoteDir As String, strUserName As String, strPassWord As String) As Long
    '����������Դ
    Dim NetR As NETRESOURCE
    Dim lngResult As Long
    
    NetR.dwType = RESOURCETYPE_ANY
    NetR.lpLocalName = vbNullString
    NetR.lpRemoteName = strShareRemoteDir
    NetR.lpProvider = vbNullString
    lngResult = WNetAddConnection2(NetR, strPassWord, strUserName, 0)
    
    If lngResult <> 0 Then
        MsgBox "��������ʧ�ܣ��������������Ƿ���ȷ��"
    End If
    funcConnectShardDir = lngResult
End Function

Public Function bln����δ���(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal lngҽ��ID As Long, ByVal lng��Դ As Long) As Boolean
'�жϲ����Ƿ��ѳ�Ժ��Ϊ���ﲡ�ˣ����м��˷���δ���
'��Ҫ����ϵͳ�����жϣ�81�Ų�����"ִ�к��Զ���˻��۵�"
    
    bln����δ��� = False
    
    If zlDatabase.GetPara(81, glngSys) = 1 Then
        If Not bln������Ժ(lng����ID, lng��ҳID) And bln����δ�󻮼۵�(lngҽ��ID, lng��Դ) Then
            bln����δ��� = True
        End If
    End If
End Function

Public Function AssembleImage(AssembleViewer As DicomImages, ByVal intRows As Integer, ByVal intCols As Integer, _
    ByVal lngHeight As Long, ByVal lngWidth As Long) As DicomImage

'���viewer�е���ʾ������ͼ���һ��ͼ��

    Dim Image As New DicomImage '��ͼ��
    Dim imgs As New DicomImages '��ʱ�洢��Ļ�ɼ���ͼ��
    Dim intWidth As Integer     '��ͼ��Ŀ��
    Dim intHeight As Integer    '��ͼ��ĸ߶�
    Dim Simg As New DicomImage
    Dim sZoom As Single
    Dim intImgRectWidth As Integer  '����ͼ���ռ�õ�������
    Dim intImgRectHeight As Integer '����ͼ���ռ�õ�����߶�
    Dim i As Integer
    Dim intMaxWidth As Integer      'ƴ�Ӻ�ͼ��������
    Dim intMaxHeight As Integer     'ƴ�Ӻ�ͼ������߶�
    Dim intBorder As Integer        'ͼ��֮��ı߾�
    Dim intOffsetX As Integer       'ƴ��ʱX�����λ��
    Dim intOffsetY As Integer       'ƴ��ʱY�����λ��
    Dim lngWhiteX As Long           '��ͼ���ɫ�ĳɰ�ɫ��X���
    Dim lngWhiteY As Long           '��ͼ���ɫ�ĳɰ�ɫ��Y�߶�
    
    If AssembleViewer.Count <= 0 Then
        '����һ����ͼ**************
        Exit Function
    End If

    On Error GoTo err
    '������ͼ��Ŀ�Ⱥ͸߶�

    '��ͼ��Ŀ�Ⱥ͸߶Ȳ��ܹ�����intMaxWidth��intMaxHeight����ȡ��߶ȣ�
    intMaxWidth = 3073
    intMaxHeight = 3073
    intBorder = 10

    intImgRectWidth = 0
    intImgRectHeight = 0

    '������ͼ��Ŀ�Ⱥ͸߶�

    'ʹ��ԭͼ��Ŀ�Ⱥ͸߶Ⱥͣ�����Viewer�ı�����������

    '����ͼ����¿��
    For i = 1 To AssembleViewer.Count
        If intImgRectWidth < AssembleViewer(i).SizeX Then intImgRectWidth = AssembleViewer(i).SizeX
        If intImgRectHeight < AssembleViewer(i).SizeY Then intImgRectHeight = AssembleViewer(i).SizeY
    Next i
    
    '������������ͼ������
    intWidth = intImgRectWidth * intCols
    intHeight = intImgRectHeight * intRows
    
    '����ͼ��Ŀ�ߣ����ܴ������ֵ
    '�������intMaxWidth��intMaxHeight�򣬰���ͼ���ܳ���ȣ�ʹ��С�ڵ���intMaxWidth��intMaxHeight��Ϊ�¿��,
    If intWidth > intMaxWidth Or intHeight > intMaxHeight Then
        If intHeight / intWidth > intMaxHeight / intMaxWidth Then
            intWidth = intWidth / intHeight * intMaxHeight
            intHeight = intMaxHeight
        Else
            intHeight = intHeight / intWidth * intMaxWidth
            intWidth = intMaxWidth
        End If
    End If
    
    '�ɼ�ͼ��
    '��ͼ��ɼ�����ʱͼ��
    For i = 1 To AssembleViewer.Count
        '�������ű��� hj�޸�,�����ͼ�ϲ�ʱ���Ŵ��ͼ���޷������Ŵ������
        sZoom = intImgRectHeight / AssembleViewer(i).SizeY
        If sZoom > intImgRectWidth / AssembleViewer(i).SizeX Then
            sZoom = intImgRectWidth / AssembleViewer(i).SizeX
        End If
        
        AssembleViewer(i).StretchToFit = False
        AssembleViewer(i).Zoom = sZoom
        '�ɼ�ͼ��
        Set Simg = AssembleViewer(i).PrinterImage(8, 3, True, sZoom, 0, AssembleViewer(i).SizeX, 0, AssembleViewer(i).SizeY)
        imgs.Add Simg
    Next i

    '��ȷ������ͼ��Ŀ�Ⱥ͸߶�
    intImgRectWidth = 0
    intImgRectHeight = 0

    For i = 1 To imgs.Count
        If intImgRectWidth < imgs(i).SizeX Then intImgRectWidth = imgs(i).SizeX
        If intImgRectHeight < imgs(i).SizeY Then intImgRectHeight = imgs(i).SizeY
        imgs(i).Attributes.Add &H8, &H16, "doSOP_SecondaryCapture"
    Next i
    intImgRectWidth = intImgRectWidth + intBorder
    intImgRectHeight = intImgRectHeight + intBorder
    intWidth = intImgRectWidth * intCols
    intHeight = intImgRectHeight * intRows

    '������ͼ��
    Image.Name = "print"
    Image.PatientID = "print001"
    
    Image.Attributes.Add &H8, &H16, doSOP_SecondaryCapture
    Image.Attributes.Add &H28, &H2, 3 ' samples/pixel
    Image.Attributes.Add &H28, &H4, "RGB" ' photometric interpreation  'CT����MONOCHROME2,CR����MONOCHROME1��
    Image.Attributes.Add &H28, &H10, intHeight  'x,Rows
    Image.Attributes.Add &H28, &H11, intWidth 'Y,Columns
    Image.Attributes.Add &H28, &H100, 8 'bits allocated
    Image.Attributes.Add &H28, &H101, 8 ' bits stored
    Image.Attributes.Add &H28, &H102, 7 ' high bit
    ReDim pix(intWidth * 3, intHeight * 3) As Byte
    For lngWhiteX = 0 To intWidth * 3
        For lngWhiteY = 0 To intHeight * 3
            pix(lngWhiteX, lngWhiteY) = 255
        Next lngWhiteY
    Next lngWhiteX
    Image.Attributes.Add &H7FE0, &H10, pix

    'ƴ����ͼ��
    For i = 1 To imgs.Count
        '����ͼ����λ��
        intOffsetX = (intImgRectWidth - imgs(i).SizeX - intBorder) / 2
        intOffsetY = (intImgRectHeight - imgs(i).SizeY - intBorder) / 2
        Image.Blt imgs(i), 0, 0, ((i - 1) Mod intCols) * intImgRectWidth + intOffsetX, ((i - 1) \ intCols) * intImgRectHeight + intOffsetY, imgs(i).SizeX, imgs(i).SizeY, 1, 1, 1, False
    Next i

    Set AssembleImage = Image
    Exit Function
err:
End Function

Public Function FunLogIn(frmParent As Form, str���� As String) As String
'���ܣ��Գ������ע�ᣬ���ע��ɹ����򷵻�ע��ʱ��
'������ frmParent ---������
'       str���� ---'��ע������ʹ�õ���������
'����ֵ��ע��ɹ�ע�����ڣ�ע��ʧ�ܷ��ؿ�

    Dim intNUM As Integer
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim strIP��ַ As String         '��Ҫע���IP��ַ
    
    On Error GoTo err
    
    strIP��ַ = funGetOneIP
    
    '��ע��������ȡ��Ȩ��������-1--�����ƣ�0--��ֹ��X��X>0��--������������
    intNUM = gint��Ƶ�豸����
    
    'intNUM >0 ,����ù���ע�����
    If intNUM > 0 Then  '����������
        strSQL = "Zl_Ӱ�������¼_Update('" & strIP��ַ & "','" & str���� & "'," & intNUM & ")"
        zlDatabase.ExecuteProcedure strSQL, "ע��" & str����
        '���ע���Ƿ�ɹ�
        strSQL = "Select ����ʱ��,IP��ַ from Ӱ�������¼ where  ����=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ����ʱ��", str����)
        
        If rsTemp.RecordCount <= intNUM Then
            rsTemp.Filter = "IP��ַ='" & strIP��ַ & "'"
            If rsTemp.RecordCount = 1 Then  'ע��ɹ�
                FunLogIn = rsTemp!����ʱ��
                Exit Function
            End If
        End If
    ElseIf intNUM = -1 Then     '������
        FunLogIn = Now
        Exit Function
    Else    '=0����������ֵ����ֹ�������κδ�����������ʾ
    
    End If
    
    'ע��ʧ�ܣ�����������ԭ��
    '1��ע���������������ɵ��������޷�ע��IP��ַ
    '2��ֱ��ͨ��SQL����������IP��ַ�����±��еļ�¼��������������ɵ�����
    Call MsgBoxD(frmParent, "�򿪵�" & str���� & "�������������������" & intNUM & "�������������Ӧ����ϵ��", vbOKOnly, gstrSysName)
    FunLogIn = ""
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
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
    Dim intNUM As Integer
    
    On Error GoTo err
    strIP��ַ = funGetOneIP
    
    '����ʱ��Ϊ�գ���ʾע��ʧ�ܣ�û����������������˳���ʱ���ټ�����ݿ�
    If str����ʱ�� = "" Then
        FunLogOut = True
        Exit Function
    End If
    
    '��ע��������ȡ��Ȩ��������-1--�����ƣ�0--��ֹ��X��X>0��--������������
    intNUM = gint��Ƶ�豸����
    
    If intNUM > 0 Then '������������
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
    ElseIf intNUM = -1 Then     '������
        FunLogOut = True
    Else    '=0����������ֵ����ֹ
    
    End If
    If FunLogOut = False Then
        Call MsgBoxD(frmParent, "�򿪵�" & str���� & "�������������������" & intNUM & "�������������Ӧ����ϵ��", vbOKOnly, gstrSysName)
    End If
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function getLicenseCount(strLicenseName As String) As Integer
'��ȡ��Ȩ������
'������ strLicenseName --- ��Ȩ����
    Dim strLiceseCount As String
    
    On Error GoTo err
    
    strLiceseCount = zl9comlib.zlRegInfo(strLicenseName)
    If strLiceseCount = "" Then '������
        getLicenseCount = -1
    ElseIf Val(strLiceseCount) > 0 Then '������������
        getLicenseCount = Val(strLiceseCount)
    Else '��ֹ
        getLicenseCount = 0
    End If
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function getStudyState(ByVal lngOrderID As Long, Optional ByRef lngSendNO As Long, _
        Optional ByRef str������ As String, Optional ByRef strǩ�� As String, Optional ByRef str������ As String, _
        Optional ByRef bln���������� As Boolean) As Integer
'��鱨���ǩ�������ȷ�����μ����еĳ̶ȡ�
'������ lngOrderID [IN] --- ҽ��id
'       lngSendNo [OUT] --- ���أ����ͺ�
'       str������ [OUT] --- ���أ�����Ĵ�����
'       strǩ��   [OUT] --- ���أ���������ǩ��
'       str������ [OUT] --- ���أ��������󱣴���
'       bln����������[OUT]--- ���أ���������Ƿ��Ѿ�����,True-�����룬False-δ����
'����ֵ��1--�ѵǼǣ�2--�ѱ�����3--�Ѽ�飻4--�ѱ��棻5--����ˣ�6--����ɣ������̲������������ֵ��
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim rsLevel As ADODB.Recordset
    
    On Error GoTo err
    
    strSQL = "Select d.ҽ��id As Ӱ��ҽ��ID,e.ҽ��id As ����ҽ��ID,c.���ͺ�,d.���uid, " _
             & " e.����id,e.������, e.������, e.ǩ������, e.���ʱ��, e.���汾,c.������� " _
             & " From ����ҽ������ c, Ӱ�����¼ d, " _
             & " (Select a.ҽ��id,a.����id,b.������, b.������, b.ǩ������, b.���ʱ��, b.���汾 " _
             & "  From ����ҽ������ a, ���Ӳ�����¼ b Where a.ҽ��id = [1] And a.����id = b.Id) e " _
             & " Where c.ҽ��id = [1] And d.ҽ��id(+) = c.ҽ��id And e.ҽ��id(+) = c.ҽ��id "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�Ƿ�ǩ��", CLng(lngOrderID))
    
    '�����ѯû�н�������˳�
    If rsTemp.EOF = True Then Exit Function
    
    lngSendNO = rsTemp!���ͺ�
    str������ = Nvl(rsTemp!������)
    str������ = Nvl(rsTemp!������)
    bln���������� = Not IsNull(rsTemp!�������)
    
    '���Ӱ��ҽ��IDΪ�գ������=1,�ѵǼ�
    '�������ҽ��IDΪ�գ��� ���UIDΪ�գ������=2���ѱ���
    '�������ҽ��IDΪ�գ����UID��Ϊ�գ������=3���Ѽ��
    '�������ǩ���ͱ�����������ȷ������Ϊ2,3,4��5���ѱ���,�Ѽ��,�ѱ��棬�����
    
    If Nvl(rsTemp!Ӱ��ҽ��ID) = "" Then     '����=1,�ѵǼ�
        getStudyState = 1
    ElseIf Nvl(rsTemp!����ҽ��ID) = "" And Nvl(rsTemp!���uid) = "" Then    '����=2���ѱ���
        getStudyState = 2
    ElseIf Nvl(rsTemp!����ҽ��ID) = "" And Nvl(rsTemp!���uid) <> "" Then    '����=3���Ѽ��
        getStudyState = 3
    Else    '���ǩ���ͱ���������,ȷ������Ϊ2,3,4��5���ѱ���,�Ѽ��,�ѱ��棬�����
        If Nvl(rsTemp!���ʱ��) = "" And rsTemp!���汾 = 1 Then
            'δǩ������ �����һ��ҽʦ��ǩ��ִ�й�����ͼ��Ϊ�Ѽ�飬��ͼ��Ϊ�ѱ���
            getStudyState = IIf(Nvl(rsTemp!���uid) = "", 2, 3)
        Else
            '�жϵ�ǰ�����ǩ���������������Ӳ������ݡ����д���1��ǩ��������������ˡ�
            If rsTemp!ǩ������ > 1 Then '�����
                getStudyState = 5
            ElseIf rsTemp!ǩ������ = 0 And rsTemp!���汾 > 1 Then
                '���˳��ֵ�״̬����������ˣ���Ҫ��顰���Ӳ������ݡ�������ǩ������
                strSQL = "Select Ҫ�ر�ʾ As ǩ������,�����ı� as ǩ��  From ���Ӳ������� Where �ļ�ID=[1] " _
                        & " And ��������= 8 And ��ʼ�� = [2] order by ǩ������ desc "
                Set rsLevel = zlDatabase.OpenSQLRecord(strSQL, "��ȡǩ������", CLng(rsTemp!����Id), CLng(rsTemp!���汾 - 1))
                
                If rsLevel.EOF = False Then
                    If rsLevel!ǩ������ > 1 Then
                        getStudyState = 5
                        strǩ�� = Split(Nvl(rsLevel!ǩ��), ";")(0)
                    Else
                        getStudyState = 4
                    End If
                Else
                    getStudyState = 4
                End If
            Else
                getStudyState = 4
            End If
        End If
    End If
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function funGetOneIP() As String
'��ȡ��ǰ��������׸�IP��ַ
    Dim strIP��ַ As String
    
    On Error Resume Next
    
    strIP��ַ = funcGetLocalIP
    If strIP��ַ = "" Then
        funGetOneIP = "127.0.0.1"
    ElseIf InStr(strIP��ַ, ",") <> 0 Then
        funGetOneIP = Split(strIP��ַ, ",")(0)
    Else
        funGetOneIP = strIP��ַ
    End If
End Function

Private Function funcGetLocalIP() As String
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
        Exit Sub
    End If

    If lobyte(WSAD.wversion) < WS_VERSION_MAJOR Or (lobyte(WSAD.wversion) = _
        WS_VERSION_MAJOR And hibyte(WSAD.wversion) < WS_VERSION_MINOR) Then

        sHighByte = Trim$(Str$(hibyte(WSAD.wversion)))
        sLowByte = Trim$(Str$(lobyte(WSAD.wversion)))
        sMsg = "Windows Sockets version " & sLowByte & "." & sHighByte
        sMsg = sMsg & " is not supported by winsock.dll "
        MsgBox sMsg
        Exit Sub
    End If

    ''''''''''''''''iMaxSockets is not used in winsock 2. So the following check is only
    ''''''''''''''''necessary for winsock 1. If winsock 2 is requested,
    ''''''''''''''''the following check can be skipped.

    If WSAD.iMaxSockets < MIN_SOCKETS_REQD Then
        sMsg = "This application requires a minimum of "
        sMsg = sMsg & Trim$(Str$(MIN_SOCKETS_REQD)) & " supported sockets."
        MsgBox sMsg
        Exit Sub
    End If
End Sub

Private Sub SocketsCleanup()
Dim lReturn As Long

    lReturn = WSACleanup()

    If lReturn <> 0 Then
        MsgBox "Socket error " & Trim$(Str$(lReturn)) & " occurred in Cleanup "
        Exit Sub
    End If
End Sub

Private Function hibyte(ByVal wParam As Integer)
    hibyte = wParam \ &H100 And &HFF&
End Function

Private Function lobyte(ByVal wParam As Integer)
    lobyte = wParam And &HFF&
End Function
