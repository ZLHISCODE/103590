Attribute VB_Name = "mdlMain"
Option Explicit

Public gcnHIS As New ADODB.Connection
Public gcnPACS As New ADODB.Connection

Public gstrSysName As String                'ϵͳ����
Public gstrDbUser As String                 '��ǰ���ݿ��û�
Public gstrHISUser As String                'HIS���ݿ��û���
Public gstrHISPassw As String               'HIS���ݿ�����
Public gstrHISsid As String                 'HIS���ݿ�SID��
Public glngInterval As Long                 '����ʱ��������λ��
Public gstrPACSUser As String               'PACS���ݿ��û���
Public gstrPACSPassw As String              'PACS���ݿ�����
Public gstrPACSsid As String                'PACS���ݿ�SID��
Public gstrPACSport As String                'PACS�˿�
Public gstrPACSIP As String                 'PACS���ݿ��������IP��ַ
Public gstrRegPath As String                'ע���·��




Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long
Public Declare Function ShowWindow Lib "user32" (ByVal Hwnd As Long, ByVal nCmdShow As Long) As Long
Public Const SRCCOPY = &HCC0020
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long


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
Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Boolean
End Type


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


Public Function OraDataOpen() As Boolean
    '------------------------------------------------
    '���ܣ� ��ָ�������ݿ�
    '������
    '���أ� ���ݿ�򿪳ɹ�������true��ʧ�ܣ�����false
    '------------------------------------------------
    Dim strSQL As String
    Dim strError As String
        
    On Error Resume Next
    Err = 0
    DoEvents
    With gcnHIS
        If .State = adStateOpen Then .Close
        .Provider = "MSDataShape"
        .Open "Driver={Microsoft ODBC for Oracle};Server=" & gstrHISsid, gstrHISUser, gstrHISPassw
        If Err <> 0 Then
            '���������Ϣ
            strError = Err.Description
            If InStr(strError, "�Զ�������") > 0 Then
                MsgBox "���Ӵ��޷��������������ݷ��ʲ����Ƿ�������װ��", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-12154") > 0 Then
                MsgBox "�޷���������������" & vbCrLf & "������Oracle�������Ƿ���ڸñ�������������������ַ�������", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-12541") > 0 Then
                MsgBox "�޷����ӣ�����������ϵ�Oracle�����������Ƿ�������", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-01033") > 0 Then
                MsgBox "ORACLE���ڳ�ʼ�����ڹرգ����Ժ����ԡ�", vbInformation, gstrSysName
            Else
                MsgBox "�����û�������������ָ�������޷�ע�ᡣ " + Err.Description, vbInformation, gstrSysName
            End If
            
            OraDataOpen = False
            Exit Function
        End If
    End With
    
    Err = 0
    On Error GoTo errHand
    
    gstrDbUser = UCase(gstrHISUser)
    SetDbUser gstrDbUser
    OraDataOpen = True
    
    
    '�򿪿´� ���ݿ�����
    '��ʱʹ��HISģ��
    With gcnPACS
        If .State = adStateOpen Then .Close
       ' .Provider = "MSDataShape"
       ' .Open "Driver={Microsoft ODBC for Oracle};Server=" & gstrPACSsid, gstrPACSUser, gstrPACSPassw
       'odbc
        .Open "Driver={SYBASE ASE ODBC Driver};NA=" & gstrPACSIP & "," & gstrPACSport & ";Uid=" & gstrPACSUser & ";Pwd=" & gstrPACSPassw & ";"
        'ole
       '.Open "Provider=Sybase.ASEOLEDBProvider.2;" & _
              "Server Name=" & gstrPACSIP & ";" & _
              "Server Port Address=" & gstrPACSport & ";" & _
              "Initial Catalog=" & gstrPACSsid & ";" & _
              "User ID=" & gstrPACSUser & ";" & _
              "Password=" & gstrPACSPassw & ";"
        If Err <> 0 Then
            '���������Ϣ
            strError = Err.Description
            If InStr(strError, "�Զ�������") > 0 Then
                MsgBox "���Ӵ��޷��������������ݷ��ʲ����Ƿ�������װ��", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-12154") > 0 Then
                MsgBox "�޷���������������" & vbCrLf & "������Oracle�������Ƿ���ڸñ�������������������ַ�������", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-12541") > 0 Then
                MsgBox "�޷����ӣ�����������ϵ�Oracle�����������Ƿ�������", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-01033") > 0 Then
                MsgBox "ORACLE���ڳ�ʼ�����ڹرգ����Ժ����ԡ�", vbInformation, gstrSysName
            Else
                MsgBox "�����û�������������ָ�������޷�ע�ᡣ " + Err.Description, vbInformation, gstrSysName
            End If
            
            OraDataOpen = False
            Exit Function
        End If
    End With
    
    '�´�ʹ�õ���SYBASE�����ݿ⣬��Ҫͨ��ADODC�������������֣�Ȼ���滻�����³���ġ�Open����
'    With gcnPACS
'        If .State = adStateOpen Then .Close
'        .Open "Driver={Microsoft ODBC for Oracle};Server=" & gstrPACSsid, gstrPACSUser, gstrPACSPassw
'        If Err <> 0 Then
'            '���������Ϣ
'            strError = Err.Description
'            MsgBox "����SQL Server����"
'            OraDataOpen = False
'            Exit Function
'        End If
'    End With


       
'�����SQLServer2000���ݿ����ӵ����ӣ����Բ��ù�
'    With gcnSQL2K
'        If .State = adStateOpen Then .Close
'
'        .Open "Provider=SQLOLEDB.1;Persist Security Info=False;Initial Catalog=yygl;Data Source=" & gstrHISIP, gstrUser, gstrPassw
'        If Err <> 0 Then
'            MsgBox "����SQL Server����"
'        End If
'    End With
    
    
    
    Exit Function
    
errHand:
    If ErrCenter() = 1 Then Resume
    OraDataOpen = False
    Err = 0
End Function


Public Sub Main()
    '�û�ע��
    gstrRegPath = "HIStoKodakPacs"
    
    '��ע����ȡ�������ݿ����Ӳ���
    gstrHISUser = GetSetting("ZLSOFT", gstrRegPath, "HIS�û���", "zlhis")
    gstrHISPassw = GetSetting("ZLSOFT", gstrRegPath, "HIS����", "his")
    gstrHISsid = GetSetting("ZLSOFT", gstrRegPath, "HISsid", "")
    
    gstrPACSIP = GetSetting("ZLSOFT", gstrRegPath, "PACSIP��ַ", "172.16.9.13")
    gstrPACSUser = GetSetting("ZLSOFT", gstrRegPath, "PACS�û���", "zlhis")
    gstrPACSPassw = GetSetting("ZLSOFT", gstrRegPath, "PACS����", "his123")
    gstrPACSsid = GetSetting("ZLSOFT", gstrRegPath, "PACSsid", "ris")
    gstrPACSport = GetSetting("ZLSOFT", gstrRegPath, "gstrPACSport", "4100")
    
    glngInterval = GetSetting("ZLSOFT", gstrRegPath, "�������", "5")
    
    '�������ݿ����Ӳ���д��ע���
    SaveSetting "ZLSOFT", gstrRegPath, "HIS�û���", gstrHISUser
    SaveSetting "ZLSOFT", gstrRegPath, "HIS����", gstrHISPassw
    SaveSetting "ZLSOFT", gstrRegPath, "HISsid", gstrHISsid
   
    
    SaveSetting "ZLSOFT", gstrRegPath, "�������", glngInterval
    
    SaveSetting "ZLSOFT", gstrRegPath, "PACSIP��ַ", gstrPACSIP
    SaveSetting "ZLSOFT", gstrRegPath, "PACS�û���", gstrPACSUser
    SaveSetting "ZLSOFT", gstrRegPath, "PACS����", gstrPACSPassw
    SaveSetting "ZLSOFT", gstrRegPath, "PACSsid", gstrPACSsid
    SaveSetting "ZLSOFT", gstrRegPath, "gstrPACSport", gstrPACSport
    
    OraDataOpen
    
    
    If gcnHIS.State <> adStateOpen Or gcnPACS.State <> adStateOpen Then
        Exit Sub
    End If
    
    frmSendOrder.Show 1
End Sub
