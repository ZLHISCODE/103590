VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit
'##############################################################################
'#���ܣ�����ϵͳ��ػ�ײ���ض�����API��װ�Ĺ�������
'�������򣺺�����������Ϊͨ���׶���Ӣ�ļ���
'�ϼ��ࣺclsComLib
'##############################################################################
'ע����ؼ��ָ�����
Private Enum REGRoot
    HKEY_CLASSES_ROOT = &H80000000 '��¼Windows����ϵͳ�����������ļ��ĸ�ʽ�͹�����Ϣ����Ҫ��¼��ͬ�ļ����ļ�����׺����֮��Ӧ��Ӧ�ó��������Ӽ��ɷ�Ϊ���࣬һ�����Ѿ�ע��ĸ����ļ�����չ���������Ӽ�ǰ�涼��һ������������һ���Ǹ����ļ������й���Ϣ��
    HKEY_CURRENT_USER = &H80000001 '�˸��������˵�ǰ��¼�û����û������ļ���Ϣ����Щ��Ϣ��֤��ͬ���û���¼�����ʱ��ʹ���Լ��ĸ��Ի����ã������Լ������ǽֽ���Լ����ռ��䡢�Լ��İ�ȫ����Ȩ�޵ȡ�
    HKEY_LOCaL_MaCHINE = &H80000002 '�˸��������˵�ǰ��������������ݣ���������װ��Ӳ���Լ����������á���Щ��Ϣ��Ϊ���е��û���¼ϵͳ����ġ���������ע��������Ӵ�Ҳ������Ҫ�ĸ�����
    HKEY_USERS = &H80000003 '�˸�������Ĭ���û�����Ϣ��Default�Ӽ�����������ǰ��¼�û�����Ϣ��
    HKEY_PERFORMANCE_DATA = &H80000004 '��Windows NT/2000/XPע�������Ȼû��HKEY_DYN_DATA����������ȴ������һ����Ϊ��HKEY_ PERFOR MANCE_DATA����������ϵͳ�еĶ�̬��Ϣ���Ǵ���ڴ��Ӽ��С�ϵͳ�Դ���ע����༭���޷������˼�
    HKEY_CURRENT_CONFIG = &H80000005  '�˸���ʵ������HKEY_LOCAL_MACHINE�е�һ���֣����д�ŵ��Ǽ������ǰ���ã�����ʾ������ӡ���������������Ϣ�ȡ������Ӽ���HKEY_LOCAL_ MACHINE\ Config\0001��֧�µ�������ȫһ����
    HKEY_DYN_DATA = &H80000006 '�˸����б���ÿ��ϵͳ����ʱ��������ϵͳ���ú͵�ǰ������Ϣ���������ֻ������Windows 98�С�
End Enum

'ע�����������
Private Enum REGValueType
    REG_NONE = 0                       ' No value type
    REG_SZ = 1 'Unicode���ս��ַ���
    REG_EXPAND_SZ = 2 'Unicode���ս��ַ���
    REG_BINARY = 3 '��������ֵ
    REG_DWORD = 4 '32-bit ����
    REG_DWORD_BIG_ENDIAN = 5
    REG_LINK = 6
    REG_MULTI_SZ = 7 ' ��������ֵ��
End Enum
'�򿪴���
Private Enum REGErr
    ERROR_SUCCESS = &H0
    ERROR_FILE_NOT_FOUND = &H2 'The system cannot find the file specified
    ERROR_BADDB = 1009&
    ERROR_BADKEY = 1010&
    ERROR_CANTOPEN = 1011&
    ERROR_CANTREAD = 1012&
    ERROR_CANTWRITE = 1013&
    ERROR_OUTOFMEMORY = 14&
    ERROR_INVALID_PARAMETER = 87&
    ERROR_ACCESS_DENIED = 5&
    ERROR_NO_MORE_ITEMS = 259&
    ERROR_MORE_DATA = 234&
End Enum
'ע�������Ȩ
Private Enum REGRights
    KEY_QUERY_VaLUE = &H1
    KEY_SET_VaLUE = &H2
    KEY_CREaTE_Sub_KEY = &H4
    KEY_ENUMERaTE_Sub_KEYS = &H8
    KEY_NOTIFY = &H10
    KEY_CREaTE_LINK = &H20
    KEY_aLL_aCCESS = &H3F
    KEY_READ = &H20019
End Enum
' ���价���ַ�����������������������д�������Ϊ��ࡣҲ����˵�����ɰٷֺŷ�������Ļ���������ת�����Ǹ����������ݡ����磬��%path%�������������·����
Private Declare Function ExpandEnvironmentStrings Lib "kernel32" Alias "ExpandEnvironmentStringsA" (ByVal lpSrc As String, ByVal lpDst As String, ByVal nSize As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal uloptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Private Declare Function RegQueryValueEx_ValueType Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegQueryValueEx_Long Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Long, lpcbData As Long) As Long
Private Declare Function RegQueryValueEx_String Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function RegQueryValueEx_BINARY Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
Private Declare Function RegSetValueEx_String Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal lpcbData As Long) As Long
Private Declare Function RegSetValueEx_Long Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long
Private Declare Function RegSetValueEx_BINARY Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Byte, ByVal cbData As Long) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long

'��ͣ(Wait)
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'��������(ComputerName)
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
'�Ƿ���64λ���̣�Is64bit��
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function IsWow64Process Lib "kernel32" (ByVal hProc As Long, bWow64Process As Long) As Long
'��ʱIP��ȡ
Private Const MAX_IP = 5                                                    'To make a buffer... i dont think you have more than 5 ip on your pc..
Private Type IPINFO
    dwAddr As Long                                                          ' IP address
    dwIndex As Long                                                         ' interface index
    dwMask As Long                                                          ' subnet mask
    dwBCastAddr As Long                                                     ' broadcast address
    dwReasmSize  As Long                                                    ' assembly size
    unused1 As Integer                                                      ' not currently used
    unused2 As Integer                                                      '; not currently used
End Type
Private Type MIB_IPADDRTABLE
    dEntrys As Long                                                         'number of entries in the table
    mIPInfo(MAX_IP) As IPINFO                                               'array of IP address entries
End Type
Private Declare Function GetIpAddrTable Lib "IPHlpApi" (pIPAdrTable As Byte, pdwSize As Long, ByVal Sort As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
'��дini �ļ���API
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal LpApplicationName As String, ByVal LpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal LpApplicationName As String, ByVal LpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
'ϵͳ��ɫת����SysColor2RGB��
Private Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
'����ϵͳ�п��õ����뷨�����������뷨����Layout,����Ӣ�����뷨��
Private Declare Function GetKeyboardLayoutList Lib "user32" (ByVal nBuff As Long, lpList As Long) As Long
'��ȡĳ�����뷨������
Private Declare Function ImmGetDescription Lib "imm32.dll" Alias "ImmGetDescriptionA" (ByVal hkl As Long, ByVal lpsz As String, ByVal uBufLen As Long) As Long
'�ж�ĳ�����뷨�Ƿ��������뷨
Private Declare Function ImmIsIME Lib "imm32.dll" (ByVal hkl As Long) As Long
'�л���ָ�������뷨��
Private Declare Function ActivateKeyboardLayout Lib "user32" (ByVal hkl As Long, ByVal flags As Long) As Long

Public Sub Wait(ByVal lngMilliseconds As Long)
'���ܣ�ִ�й���һ��ʱ��
'lngMilliseconds=��������1000����=1��
    Call Sleep(lngMilliseconds)
End Sub

Public Function TempPath() As String
'����:��ȡ��ʱ�ļ���
    Dim strFileTemp As String
    Dim lngTemp As Long
    Dim objFile As New FileSystemObject
    '����ʹ��VB�Դ���ȡ������ʧ����ͨ��API
    strFileTemp = objFile.GetSpecialFolder(TemporaryFolder)
    If strFileTemp = "" Then
        strFileTemp = Space(256)
        lngTemp = GetTempPath(256, strFileTemp)
        TempPath = gobjComLib.zlStr.TruncZero(strFileTemp)
    Else 'C:\Users\ADMINI~1\AppData\Local\Temp\
        TempPath = strFileTemp & "\"
    End If
End Function

Public Function IP(Optional ByRef strErr As String) As String
    '���ܣ�ͨ��API��ȡ��ʱIP
    
    Dim ret As Long, Tel As Long
    Dim bBytes() As Byte
    Dim TempList() As String
    Dim TempIP As String
    Dim Tempi As Long
    Dim Listing As MIB_IPADDRTABLE
    Dim L3 As String
    Dim strTmpErr As String, strALLErr As String
    
    strErr = ""
    On Error GoTo Errhand
    GetIpAddrTable ByVal 0&, ret, True
    If ret <= 0 Then Exit Function
    ReDim bBytes(0 To ret - 1) As Byte
    ReDim TempList(0 To ret - 1) As String
    'retrieve the data
    GetIpAddrTable bBytes(0), ret, False
    'Get the first 4 bytes to get the entry's.. ip installed
    CopyMemory Listing.dEntrys, bBytes(0), 4
    For Tel = 0 To Listing.dEntrys - 1
        'Copy whole structure to Listing..
        CopyMemory Listing.mIPInfo(Tel), bBytes(4 + (Tel * Len(Listing.mIPInfo(0)))), Len(Listing.mIPInfo(Tel))
        TempList(Tel) = ConvertAddressToString(Listing.mIPInfo(Tel).dwAddr, strTmpErr)
        If strTmpErr <> "" Then strALLErr = strALLErr & IIf(strALLErr = "", "", "|") & strTmpErr
    Next Tel
    'Sort Out The IP For WAN
        TempIP = TempList(0)
        For Tempi = 0 To Listing.dEntrys - 1
            L3 = Left(TempList(Tempi), 3)
            If L3 <> "169" And L3 <> "127" And L3 <> "192" Then
                TempIP = TempList(Tempi)
            End If
        Next Tempi
        IP = TempIP 'Return The TempIP
    Exit Function
    strErr = strALLErr
    '------------------------------------------------------------------------------------------------------------------
Errhand:
    strErr = strALLErr & IIf(strALLErr = "", "", "|") & Err.Description
    Err.Clear
End Function

Private Function ConvertAddressToString(longAddr As Long, Optional ByRef strErr As String) As String
    Dim myByte(3) As Byte
    Dim Cnt As Long
    
    strErr = ""
    On Error GoTo errH
    CopyMemory myByte(0), longAddr, 4
    For Cnt = 0 To 3
        ConvertAddressToString = ConvertAddressToString + CStr(myByte(Cnt)) + "."
    Next Cnt
    ConvertAddressToString = Left$(ConvertAddressToString, Len(ConvertAddressToString) - 1)
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errH:
    strErr = Err.Description
    Err.Clear
End Function

Public Function SystemImes() As Variant
'���ܣ���ϵͳ�������뷨���Ʒ��ص�һ���ַ���������
'���أ�����������������뷨,�򷵻ؿմ�
    Dim arrIme(99) As Long, arrName() As String
    Dim lngLen As Long, strName As String * 255
    Dim lngCount As Long, i As Long, j As Long
    
    lngCount = GetKeyboardLayoutList(UBound(arrIme) + 1, arrIme(0))
    For i = 0 To lngCount - 1
        If ImmIsIME(arrIme(i)) = 1 Then
            ReDim Preserve arrName(j)
            lngLen = ImmGetDescription(arrIme(i), strName, Len(strName))
            arrName(j) = Mid(strName, 1, InStr(strName, Chr(0)) - 1)
            j = j + 1
        End If
    Next
    SystemImes = IIf(j > 0, arrName, vbNullString)
End Function

Public Function SysColor2RGB(ByVal lngColor As Long) As Long
'���ܣ���VB��ϵͳ��ɫת��ΪRGBɫ
    If lngColor < 0 Then
        Call OleTranslateColor(lngColor, 0, lngColor)
    End If
    SysColor2RGB = lngColor
End Function

Public Sub ShowChildWindow(ByVal lnghwnd As Long, Optional ByVal lngMainHwnd As Long, Optional ByVal blnMaximized As Boolean)
'���ܣ���ʾָ���Ĵ��壬���Ӵ��巽ʽ
'������lngHwnd=Ҫ��Ϊ�Ӵ�����ʾ�Ĵ���ľ��
'      lngMainHwnd=��������������ʱ�������������Ӵ�����ʾ
'      blnMaximized=�Ƿ�̶���������󻯴�������ΪAPI����Ϊ�Ӵ���󣬴���ԭ�е�������Խ�ʧЧ
'˵�����������Ҫ������ZLBH���ںϵ���ZLHIS������ʾ
    Call gobjComLib.zlCommFun.ShowChildWindow(lnghwnd, lngMainHwnd, blnMaximized)
End Sub

Public Sub SetWindowsInTaskBar(ByVal lnghwnd As Long, ByVal blnShow As Boolean)
'���ܣ����ô����Ƿ�������������ʾ
    Call gobjComLib.zlCommFun.SetWindowsInTaskBar(lnghwnd, blnShow)
End Sub

Public Function ExpandEnvStr(ByVal strInput As String) As String
'���ܣ����ַ����еĻ��������滻Ϊ����ֵ
'         strInput=���������������ַ���
'���أ���ʵ�ʵ�ֵ�滻�ַ����еĻ�����������ַ���
    '// �磺 %PATH% �򷵻� "c:\;c:\windows;"
    Dim lngLen As Long, strBuf As String, strOld As String
    strOld = strInput & "  " ' ��֪ΪʲôҪ�������ַ������򷵻�ֵ������������ַ���
    strBuf = "" '// ��֧��Windows 95
    '// get the length
    lngLen = ExpandEnvironmentStrings(strOld, strBuf, lngLen)
    '// չ���ַ���
    strBuf = String$(lngLen - 1, Chr$(0))
    lngLen = ExpandEnvironmentStrings(strOld, strBuf, LenB(strBuf))
    '// ���ػ�������
    ExpandEnvStr = gobjComLib.zlStr.TruncZero(strBuf)
End Function

Public Function GetAllSubKey(ByVal strKey As String) As Variant
'����:��ȡĳ�����������
'���أ�=��������
    Dim lnghKey As Long, lngRet As Long, strName As String, lngIdx As Long
    Dim hRootKey As Long, strKeyName As String
    Dim strSubKey As Variant
    strSubKey = Array()
    lngIdx = 0: strName = String(256, Chr(0))
     If Not GetKeyValueInfo(strKey, "", hRootKey, strKeyName) Then Exit Function
    lngRet = RegOpenKey(hRootKey, strKeyName, lnghKey)
    If lngRet = 0 Then
        Do
            lngRet = RegEnumKey(lnghKey, lngIdx, strName, Len(strName))
            If lngRet = 0 Then
                ReDim Preserve strSubKey(UBound(strSubKey) + 1)
                strSubKey(UBound(strSubKey)) = Left(strName, InStr(strName, Chr(0)) - 1)
                lngIdx = lngIdx + 1
            End If
        Loop Until lngRet <> 0
    End If
    RegCloseKey lnghKey
    GetAllSubKey = strSubKey
End Function

Private Function GetKeyValueInfo(ByVal strKey As String, Optional ByVal strValueName As String, Optional ByRef hRootKey As REGRoot, Optional ByRef strSubKey As String, Optional ByRef lngType As Long) As Boolean
'���ܣ����ݼ�λ��ȡ����ֵ���ӽ�,�Լ�ֵ����
'������strKey=ע�����λ���硰HKEY_CURRENT_USER\Printers\DevModePerUser"
'          strValueName=������
'���Σ�
'          hRootKey=����
'          strSubKey=�ӽ�
'          lngType=������
'���أ��Ƿ��ȡ�ɹ�
    Dim strRoot As String, lngPos As String, hKey As Long
    Dim lngReturn As Long, strName As String * 255
    
    On Error GoTo errH
    hRootKey = 0: strSubKey = "": lngType = 0
    lngPos = InStr(strKey, "\")
    If lngPos = 0 Then Exit Function
    strRoot = Mid(strKey, 1, lngPos - 1)
    strSubKey = Mid(strKey, lngPos + 1)
    
    hRootKey = gobjComLib.Decode(UCase(strRoot), "HKEY_CLASSES_ROOT", HKEY_CLASSES_ROOT, _
                                                                         "HKEY_CURRENT_USER", HKEY_CURRENT_USER, _
                                                                         "HKEY_LOCAL_MACHINE", HKEY_LOCaL_MaCHINE, _
                                                                         "HKEY_USERS", HKEY_USERS, _
                                                                         "HKEY_PERFORMANCE_DATA", HKEY_PERFORMANCE_DATA, _
                                                                         "HKEY_CURRENT_CONFIG", HKEY_CURRENT_CONFIG, _
                                                                         "HKEY_DYN_DATA", HKEY_DYN_DATA, 0)
    If hRootKey = 0 Then Exit Function
    If lngType <> -1 Then
        'ʹ�ò�ѯ��ʽ�򿪣����м������Ͳ�ѯ
        lngReturn = RegOpenKeyEx(hRootKey, strSubKey, 0, KEY_QUERY_VaLUE, hKey)
        If lngReturn <> ERROR_SUCCESS Then
            Exit Function
        End If
        If strValueName <> "" Then
            lngReturn = RegQueryValueEx_ValueType(hKey, strValueName, ByVal 0&, lngType, ByVal strName, Len(strName))
            'SetRegKey����������ص�����Ϊ�ܴ��������ֵ���̶�,�������Ϊ0�����ݴ������������ж�
            If lngReturn = ERROR_BADKEY Then
                If lngType < REG_NONE Or lngType > REG_MULTI_SZ Then lngType = REG_NONE
            End If
            '�����ֶγ��������Ȳ��������Գ������˳�
            'If lngReturn <> ERROR_SUCCESS Then: RegCloseKey (hKey): Exit Function
        End If
        RegCloseKey (hKey)
    End If
    GetKeyValueInfo = True
    Exit Function
errH:
    If 0 = 1 Then
        Resume
    End If
    Err.Clear
End Function

Public Function GetRegValue(ByVal strKey As String, ByVal strValueName As String, ByRef varValue As Variant, Optional blnOneString As Boolean = False) As Boolean
'���ܣ���ȡע�����ָ��λ�õ�ֵ
'������strKey=ע�����λ���硰HKEY_CURRENT_USER\Printers\DevModePerUser"
'          strValueName=������
'          strValue=����ֵ
'          strValueType=�������ͣ�Ĭ��Ϊ�ַ���
'           blnOneString = ��REG_EXPAND_SZ��REG_MULTI_SZ,REG_BINARY��Ч��-  True �������ص�һ�ַ������Ҳ����κδ�����ֻȥ���ַ���β��
'���أ��Ƿ��ȡ�ɹ�
'˵������ǰֻ��REG_SZ, REG_EXPAND_SZ, REG_MULTI_SZ��REG_DWORD��REG_BINARYʵ���˶�ȡ��û�в�ѯ�������Զ����Ҽ���
    Dim hRootKey As REGRoot, strSubKey As String
    Dim lngReturn As Long
    Dim lngKey As Long, ruType As REGValueType
    Dim lngLength As Long, varBufData As Variant, strBufVar() As String, lngBuf As Long, bytBuf() As Byte, strBuf As String
    Dim i As Long, strReturn As String, strTmp As String
    '������Ч��ע�����λ,��ȡ��������
    If Not GetKeyValueInfo(strKey, strValueName, hRootKey, strSubKey, ruType) Then Exit Function
    '�򿪱���
    lngReturn = RegOpenKeyEx(hRootKey, strSubKey, 0, KEY_QUERY_VaLUE, lngKey)
    If lngReturn <> ERROR_SUCCESS Then
        Exit Function
    End If
    On Error GoTo errH
    Select Case ruType
        Case REG_SZ, REG_EXPAND_SZ, REG_MULTI_SZ '�ַ������Ͷ�ȡ
'            lngReturn = RegQueryValueEx(lngKey, strValueName, 0, ruType, 0, lngLength)
'            If lngReturn <> ERROR_SUCCESS Then Err.Clear '���ܳ����������������
            lngLength = 1024: strBuf = Space(lngLength)
            lngReturn = RegQueryValueEx_String(lngKey, strValueName, 0, ruType, strBuf, lngLength)
            If lngReturn <> ERROR_SUCCESS Then: RegCloseKey (lngKey): Exit Function
            Select Case ruType
                Case REG_SZ
                    varValue = gobjComLib.zlStr.TruncZero(strBuf)
                Case REG_EXPAND_SZ ' ���价���ַ�������ѯ���������ͷ��ض���ֵ
                    If Not blnOneString Then
                        varValue = gobjComLib.zlStr.TruncZero(ExpandEnvStr(gobjComLib.zlStr.TruncZero(strBuf)))
                    Else
                        varValue = gobjComLib.zlStr.TruncZero(strBuf)
                    End If
                Case REG_MULTI_SZ ' �����ַ���
                    If Not blnOneString Then
                        If Len(strBuf) <> 0 Then ' �������Ƿǿ��ַ��������Էָ
                            strBufVar = Split(Left$(strBuf, Len(strBuf) - 1), Chr$(0))
                        Else ' ���ǿ��ַ�����Ҫ����S(0) �����������
                            ReDim strBufVar(0) As String
                        End If
                        ' ��������ֵ������һ���ַ������飿��
                        varValue = strBufVar()
                    Else
                        varValue = gobjComLib.zlStr.TruncZero(strBuf)
                    End If
            End Select
        Case REG_DWORD
            lngReturn = RegQueryValueEx_Long(lngKey, strValueName, ByVal 0&, ruType, lngBuf, Len(lngBuf))
            If lngReturn <> ERROR_SUCCESS Then: RegCloseKey (lngKey): varValue = 0: Exit Function
            varValue = lngBuf
        Case REG_BINARY
            lngReturn = RegQueryValueEx_BINARY(lngKey, strValueName, 0, ruType, ByVal 0, lngLength)
            If lngReturn <> ERROR_SUCCESS And lngReturn <> ERROR_MORE_DATA Then
                RegCloseKey lngKey: Exit Function
                If blnOneString Then
                    varValue = "00"
                Else
                    ReDim bytBuf(0)
                    varValue = bytBuf()
                End If
            End If
            ReDim bytBuf(lngLength - 1)
            lngReturn = RegQueryValueEx_BINARY(lngKey, strValueName, 0, ruType, bytBuf(0), lngLength)
            If lngReturn <> ERROR_SUCCESS And lngReturn <> ERROR_MORE_DATA Then
                RegCloseKey lngKey: Exit Function
                If blnOneString Then
                    varValue = "00"
                Else
                    ReDim bytBuf(0)
                    varValue = bytBuf()
                End If
            End If
            If lngLength <> UBound(bytBuf) + 1 Then
               ReDim Preserve bytBuf(0 To lngLength - 1) As Byte
            End If
            ' �����ַ�����ע�⣺Ҫ���ֽ��������ת����
            If blnOneString Then
                'ѭ�����ݣ����ֽ�ת��Ϊ16�����ַ���
                For i = LBound(bytBuf) To UBound(bytBuf)
                   strTmp = CStr(Hex(bytBuf(i)))
                   If (Len(strTmp) = 1) Then strTmp = "0" & strTmp
                   strReturn = strReturn & " " & strTmp
                Next i
                varValue = Trim$(strReturn)
            Else
                varValue = bytBuf()
            End If
    End Select
    RegCloseKey lngKey
    GetRegValue = True
    Exit Function
errH:
    If 0 = 1 Then
        Resume
    End If
End Function

Public Function SetRegValue(ByVal strKey As String, ByVal strValueName As String, varValue As Variant) As Boolean
'���ܣ�����ע�����ָ��λ�õ�ֵ
'������strKey=ע�����λ���硰HKEY_CURRENT_USER\Printers\DevModePerUser"
'          strValueName=������
'          strValue=����ֵ
'          strValueType=�������ͣ�Ĭ��Ϊ�ַ���
'���أ��Ƿ����óɹ�
    Dim hRootKey As REGRoot, strSubKey As String
    Dim lngReturn As Long
    Dim lngKey As Long, ruType As REGValueType
    Dim lngLength As Long, varBufData As Variant, strBufVar() As String, lngBuf As Long, bytBuf() As Byte, strBuf As String
    Dim i As Long, lb As Long, ub As Long, strReturn As String, strTmp As String
    '������Ч��ע�����λ,��ȡ��������
    If Not GetKeyValueInfo(strKey, strValueName, hRootKey, strSubKey, ruType) Then Exit Function
    '�򿪱���
    lngReturn = RegOpenKeyEx(hRootKey, strSubKey, 0, KEY_SET_VaLUE, lngKey)
    If lngReturn <> ERROR_SUCCESS Then
        Exit Function
    End If
    '���޷���ȡ��������ʱ�����ݴ������������ȷ����ֵ���͡�
    If ruType = REG_NONE Then
        Select Case TypeName(varValue)
            Case "Byte()"
                ruType = REG_BINARY
            Case "String"
                ruType = REG_SZ
            Case "String()"
                ruType = REG_MULTI_SZ
            Case Else
                If IsNumeric(varValue) Then
                    ruType = REG_DWORD
                End If
        End Select
    End If
    On Error GoTo errH
    Select Case ruType
        Case REG_SZ, REG_EXPAND_SZ, REG_MULTI_SZ
            If ruType = REG_MULTI_SZ And varType(varValue) = vbArray + vbString Then 'string���飬������ϳ��ַ���
                lngLength = UBound(varValue) - LBound(varValue) + 1
                For i = LBound(varValue) To UBound(varValue)
                    strBuf = strBuf & varValue(i) & Chr$(0)
                Next
                strBuf = gobjComLib.zlStr.TruncZero(strBuf)
                lngLength = gobjComLib.zlStr.ActualLen(strBuf)
            Else
                strBuf = gobjComLib.zlStr.TruncZero(varValue)
                lngLength = gobjComLib.zlStr.ActualLen(strBuf)
            End If
            lngReturn = RegSetValueEx_String(lngKey, strValueName, ByVal 0&, ruType, ByVal strBuf, lngLength)
            If lngReturn <> 0 Then RegCloseKey lngKey: Exit Function
        Case REG_DWORD
            lngBuf = Val(varValue): lngLength = Len(lngBuf)
            lngReturn = RegSetValueEx_Long(lngKey, strValueName, ByVal 0&, ruType, lngBuf, lngLength)
            If lngReturn <> 0 Then RegCloseKey lngKey: Exit Function
        Case REG_BINARY
            ' 1��varValue �� �ֽ����飬�� B()
            If varType(varValue) = vbArray + vbByte Then
                Dim binValue() As Byte, Length As Long
                bytBuf = varValue
                lngLength = UBound(bytBuf) - LBound(bytBuf) + 1
                lngReturn = RegSetValueEx_BINARY(lngKey, strValueName, 0, ruType, bytBuf(0), lngLength)
                If lngReturn <> 0 Then RegCloseKey lngKey: Exit Function
            ' 2��varValue �� ���ͻ����ͣ��� 520
            ElseIf varType(varValue) = vbLong Or varType(varValue) = vbInteger Then
                lngBuf = Val(varValue): lngLength = Len(lngBuf)
                lngReturn = RegSetValueEx_Long(lngKey, strValueName, 0, ruType, lngBuf, lngLength)
                If lngReturn <> 0 Then RegCloseKey lngKey: Exit Function
            ' 3��varValue ���ַ������� "BE 3E FF AB"
            ElseIf varType(varValue) = vbString Then
                ' ת������
                Dim ByteArray() As Byte
                Dim tmpArray() As String '//ת��ASCII�ַ���16�����ֽ�
                strTmp = varValue
                ' �Կո�ָ��ַ���
                strBufVar = Split(strTmp, " ")
                lb = LBound(strBufVar): ub = UBound(strBufVar)
                ' Ϊ��̬�������ռ�
                ReDim bytBuf(lb To ub)
                ' ѭ��ת��
                For i = lb To ub - 1
                    bytBuf(i) = CByte(Val("&H" & Right$(strBufVar(i), 2)))
                Next i
                ' ע�⣺���һ����֪���ַ����������2��ʲô��Ҫ�� Left$(tmpArray(ub), 2)
                bytBuf(ub) = CByte(Val("&H" & Left$(strBufVar(ub), 2)))
                ' ������д�뵽ע�����ע�⣺����� ub - lb + 1
                lngReturn = RegSetValueEx_BINARY(lngKey, strValueName, 0, ruType, bytBuf(0), ub - lb + 1)
                If lngReturn <> 0 Then RegCloseKey lngKey: Exit Function
            End If
    End Select
    RegCloseKey lngKey
    SetRegValue = True
    Exit Function
errH:
    If 0 = 1 Then
        Resume
    End If
End Function

Public Function DeleteRegValue(ByVal strKey As String, ByVal strValueName As String) As Boolean
'���ܣ�ɾ��ע�����ָ��λ�õ�ֵ
'������strKey=ע�����λ���硰HKEY_CURRENT_USER\Printers\DevModePerUser"
'          strValueName=������
'���أ��Ƿ��ȡ�ɹ�
    Dim lngLength As Long, lngReturn As Long
    Dim lngKey As Long, lngType As Long
    Dim hRootKey As REGRoot, strSubKey As String
    
    '������Ч��ע�����λ
    If Not GetKeyValueInfo(strKey, strValueName, hRootKey, strSubKey, -1) Then Exit Function
    '�򿪼�
    lngReturn = RegOpenKeyEx(hRootKey, strSubKey, 0, KEY_SET_VaLUE, lngKey)
    If lngReturn <> 0 Then
        Exit Function
    End If
    'ɾ����
    lngReturn = RegDeleteValue(lngKey, strValueName)
    If lngReturn = 0 Then
        DeleteRegValue = True
    End If
    '�رռ�
    RegCloseKey lngKey
End Function

Public Sub PressKeyEx(ByVal KeyCode As Integer, Optional ByVal Shift As Integer)
'���ܣ�����̷���һ����,����SendKey
'������KeyCode=VirtualKey Codes��1-254��������vbKeyTab,vbKeyReturn,vbKeyF4,vbKeyA��
'      Shift=vbKeyControl,vbKeyShift,vbKeyMenu(ALT)���������ϣ���Or����
    Call gobjComLib.zlCommFun.PressKeyEx(KeyCode, Shift)
End Sub

Public Sub PressKey(bytKey As Byte)
'���ܣ�����̷���һ����,����SendKey
'������bytKey=VirtualKey Codes��1-254��������vbKeyTab,vbKeyReturn,vbKeyF4
    Call gobjComLib.zlCommFun.PressKey(bytKey)
End Sub

Public Function OpenImeByName(Optional strIme As String) As Boolean
'����:�����ƴ��������뷨,��ָ������ʱ�ر��������뷨��֧�ֲ������ơ�
    OpenImeByName = gobjComLib.zlCommFun.OpenIme(Trim(strIme) = "", Trim(strIme))
End Function

Public Function OpenIme(Optional blnOpen As Boolean = False) As Boolean
'����:���������뷨����ر����뷨
    OpenIme = gobjComLib.zlCommFun.OpenIme(blnOpen)
End Function

Public Function OpenDir(hwndOwner As Long, Optional strTitle As String, Optional ByVal strInitDir As String) As String
'���ܣ�ѡ��Windows���ļ���
    OpenDir = gobjComLib.zlCommFun.OpenDir(hwndOwner, strTitle, strInitDir)
End Function

Public Function IsInstallExcel() As Boolean
'���ܣ��жϱ�����װ��EXCELû��
'������
'���أ����򷵻�True
    Dim objTemp  As Object
    
    On Error GoTo errH
    Set objTemp = CreateObject("Excel.Application") '��һ��EXCEL����
    Set objTemp = Nothing
    IsInstallExcel = True
    Exit Function
errH:
    Set objTemp = Nothing
    IsInstallExcel = False
    Err.Clear
End Function

Public Function IsDesinMode() As Boolean
'���ܣ� ȷ����ǰģʽΪ���ģʽ
     Err = 0: On Error Resume Next
     Debug.Print 1 / 0
     If Err <> 0 Then
        IsDesinMode = True
     Else
        IsDesinMode = False
     End If
     Err.Clear: Err = 0
 End Function
 
Public Function IsWindowsNT() As Boolean
'���ܣ��ж��Ƿ�WindowNT����ϵͳ
    IsWindowsNT = gobjComLib.zlCommFun.IsWindowsNT
End Function

Public Function IsWindows9X() As Boolean
'���ܣ��ж��Ƿ�Window95����ϵͳ
    IsWindows9X = gobjComLib.zlCommFun.IsWindows9X
End Function

Public Function Is64bit() As Boolean
    '******************************************************************************************************************
    '���ܣ��Ƿ���64λϵͳ
    '���أ�
    '******************************************************************************************************************
    Dim handle As Long
    Dim lngFunc As Long
        
    lngFunc = 0
    handle = GetProcAddress(GetModuleHandle("kernel32"), "IsWow64Process")
    If handle > 0 Then
        IsWow64Process GetCurrentProcess(), lngFunc
    End If
    Is64bit = lngFunc <> 0
End Function

Public Function IniRead(strItem As String, strKey As String, strPath As String) As String
'���ܣ�INi�ļ���
    Dim GetStr As String
    On Error GoTo errH

    GetStr = String(128, 0)
    GetPrivateProfileString strItem, strKey, "", GetStr, 256, strPath
    GetStr = Replace(GetStr, Chr(0), "")
    IniRead = GetStr
    Exit Function
errH:
    Err.Clear
    IniRead = ""
End Function

Public Function IniWrite(strItem As String, strKey As String, strVal As String, strPath As String) As Boolean
'���ܣ�INi�ļ�д
    On Error GoTo errH
    IniWrite = True
    WritePrivateProfileString strItem, strKey, strVal, strPath
    Exit Function
errH:
    Err.Clear
    IniWrite = False
End Function

Public Function Dec2Bin(bDec As Byte) As String
'���ܣ�ʮ����תΪ�����ƺ���
'�÷���String  Dec2Bin(Bdec as Byte)
'���أ�  ʮ���ƵĶ����� �ַ���(String)
'����  ����"0"

    Dim strBin As String

    On Error GoTo Err
    If bDec > 255 Then
        Dec2Bin = "-1"
        Exit Function
    End If
    strBin = ""
    'תΪ�ַ���
    While bDec > 0
        strBin = bDec Mod 2 & strBin
        bDec = Fix(bDec / 2)
    Wend
    '������8λ
    If Len(strBin) < 9 Then
        While Len(strBin) < 8
            strBin = "0" & strBin
        Wend
    End If
    Dec2Bin = strBin
    Exit Function
Err:
    Err.Clear
   Dec2Bin = "0"
End Function

Public Function Bin2Dec(strBin As String) As Long
    '���ܣ�������תΪʮ���ƺ���
    '�÷���Long  bin2dec(strBin as String)
    '���أ�  �����Ƶ�ʮ���� ��������Long��
    '����  ����-1
    Dim lDec As Long
    Dim lCount As Long
    Dim i As Long
    
    On Error GoTo Errhand
    lDec = 0
    If strBin = "" Then strBin = "0"
    lCount = Len(strBin)
    For i = 1 To lCount
        lDec = lDec + CInt(Left(strBin, 1)) * 2 ^ (Len(strBin) - 1)
        strBin = Right(strBin, Len(strBin) - 1)
        DoEvents
    Next
    Bin2Dec = lDec
    Exit Function
Errhand:
    Err.Clear
    Bin2Dec = -1
End Function

Public Function HIWORD(LongIn As Long) As Integer
'�� �� ����HIWORD
'��    �룺LongIn(Long) - 32λֵ
'��    ����(Integer) - 32λֵ�ĸ�16λ
'����������ȡ��32λֵ�ĸ�16λ

   ' ȡ��32λֵ�ĸ�16λ
     HIWORD = (LongIn And &HFFFF0000) \ &H10000
End Function

Public Function LOWORD(LongIn As Long) As Integer
'�� �� ����LOWORD
'��    �룺LongIn(Long) - 32λֵ
'��    ����(Integer) - 32λֵ�ĵ�16λ
'����������ȡ��32λֵ�ĵ�16λ
   ' ȡ��32λֵ�ĵ�16λ
     LOWORD = LongIn And &HFFFF&
End Function

Public Function ComputerName() As String
    '******************************************************************************************************************
    '���ܣ���ȡ��������
    '������
    '˵����
    '******************************************************************************************************************
    Dim strComputer As String * 256
    Call GetComputerName(strComputer, 255)
    ComputerName = strComputer
    ComputerName = Trim(Replace(ComputerName, Chr(0), ""))
End Function

Public Function ObjectHasProperty(obj As Object, ByVal strProName As String) As Boolean
'�����ж�����Ƿ�߱�ָ��������
    On Error GoTo errH
    Dim objinfo As InterfaceInfo
    Dim meminfo As MemberInfo
    Dim strTmp As String
    On Error GoTo errH
    Set objinfo = InterfaceInfoFromObject(obj)
    For Each meminfo In objinfo.Members
        If meminfo.InvokeKind = INVOKE_PROPERTYGET Then
            If UCase(strProName) = UCase(meminfo.Name) Then
                ObjectHasProperty = True: Exit For
            End If
        End If
    Next
    ObjectHasProperty = True
    Exit Function
errH:
    Err.Clear
    ObjectHasProperty = False
End Function

Public Function ObjectHasMethod(obj As Object, ByVal strFunName As String) As Boolean
'���棺���ж϶�̬�������˷�������������ޱ�ѡ�������򷽷�����óɹ�
'���ܣ��ж�����Ƿ�߱�ָ���ķ���
'������obj=����
'          strFunName=������������
'���أ�True-���ڷ�����False-�����ڷ���
    Dim objinfo As InterfaceInfo
    Dim meminfo As MemberInfo
    Dim strTmp As String
    On Error GoTo errH
    Set objinfo = InterfaceInfoFromObject(obj)
    For Each meminfo In objinfo.Members
        If meminfo.InvokeKind = INVOKE_FUNC Then
            If UCase(strFunName) = UCase(meminfo.Name) Then
                ObjectHasMethod = True: Exit For
            End If
'            Debug.Print meminfo.Name
        End If
    Next
    Exit Function
errH:
    ObjectHasMethod = False
    Err.Clear
End Function

Public Function TaskbarHeight() As Integer
    '-----------------------------------------------------------------------------------------------------------
    '����:��ȡ�������߶�
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '-----------------------------------------------------------------------------------------------------------
    Dim lRes As Long
    Dim vRect As RECT
    Err = 0: On Error GoTo Errhand:
    lRes = SystemParametersInfo(SPI_GETWORKAREA, 0, vRect, 0)
    TaskbarHeight = ((Screen.Height / Screen.TwipsPerPixelX) - vRect.Bottom) * Screen.TwipsPerPixelX
Errhand:
    Err.Clear
End Function

