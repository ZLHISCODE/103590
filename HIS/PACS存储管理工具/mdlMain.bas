Attribute VB_Name = "mdlMain"
Option Explicit

Public gobjPacsSrv As Object                '����̨
Public SplashObj As New frmSplash
Public gcnOracle As New ADODB.Connection    '�������ݿ�����

Public gstrSysName As String                'ϵͳ����
Public gstrVersion As String                'ϵͳ�汾
Public gstrAviPath As String                'AVI�ļ��Ĵ��Ŀ¼

Public gstrUserFlag As String               '��ǰ�û���־(��λ��ʾ)����1λ���Ƿ�DBA����2λ��ϵͳ������

Public gstrDbUser As String                 '��ǰ���ݿ��û�
Public glngUserId As Long                   '��ǰ�û�id
Public gstrUserCode As String               '��ǰ�û�����
Public gstrUserName As String               '��ǰ�û�����
Public gstrUserAbbr As String               '��ǰ�û�����

Public glngDeptId As Long                   '��ǰ�û�����id
Public gstrDeptCode As String               '��ǰ�û����ű���
Public gstrDeptName As String               '��ǰ�û���������

Public gstrStation As String                '������վ����
Public gstrMenuSys As String                'ϵͳ�˵�

Public strTimePolicy As String              'ʱ�����  "time,month,29,23:12:00,0,1"
'Public strCapacityPolicy As String          '��������  "capacity,30,0,1"
'Public strPercentPolicy As String           '�����ٷֱȲ���  "percent,10,1,1"
Public beginDay As Date                     '����ϵͳ��ʱ��
Public bAutoArchive As String
Public Const cAllStorageDevice = 1111       'ѡ��ȫ���洢�豸�ı�ʶ

Public ZlCommFun As New zl9comlib.clsCommFun
Public gobjFile As New FileSystemObject

'-----------------------------------------
'�����롢ע���롢�������������ע���������
Public gstrRegCode As String
Public gstrPublish As String
Public gstrParseRegCode As String
Public gstrParsePublish As String
'-----------------------------------------
Public gstrSystems As String

Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Const SRCCOPY = &HCC0020
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

'---------------------------------------------------------------
'-ע��� API ����...
'---------------------------------------------------------------
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByRef lpSecurityAttributes As SECURITY_ATTRIBUTES, ByRef phkResult As Long, ByRef lpdwDisposition As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Public Declare Function SetActiveWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
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


'---------------------------------------------------------------
'   ��Ȩ���˵������ð汾
'---------------------------------------------------------------
Public Sub Main()
    Dim lngReturn As Long
    Dim StrUnitName As String
    Dim BlnShowFlash As Boolean
    Dim strCode As String, IntCount As Integer, StrStyle As String
    Dim rsMenu As ADODB.Recordset, StrHaveSys As String
    
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
        Exit Sub
    End If
    
    
    '��ʼ����������
    InitCommon gcnOracle
    If RegCheck = False Then
        Unload SplashObj
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

    On Error Resume Next
    Err = 0
    Unload SplashObj
    frmMain.Show
End Sub

Public Function MenuGranted() As ADODB.Recordset
    '-------------------------------------------------------------
    '���ܣ�������Ȩʹ�ò���װ�Ĳ���������������Ȩʹ�õĲ˵�����
    '������ע����
    '-------------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    Dim strCodes As String
    Dim strObjs As String
    Dim IntCount As Integer
    Dim gstrSystems1 As String
    Dim gstrMenuSys As String
    Dim BlnOnlySys As Boolean 'ֻ�б���ϵͳ
    
    strCodes = zlRegFunctions(gstrRegCode) & " Or ϵͳ Is NULL "
    BlnOnlySys = (gstrSystems = "")
    If BlnOnlySys Then
        gstrSystems = " ϵͳ Is NULL"
        gstrSystems1 = " G.ϵͳ Is NULL"
        strCodes = "(" & strCodes & ") Or " & gstrSystems
    Else
        gstrSystems1 = Replace(gstrSystems, "ϵͳ", "G.ϵͳ")
        strCodes = "(" & strCodes & ") And " & gstrSystems
        
    End If
    
    '--����Ȩ�޲˵�--
    With rsTemp
        gstrMenuSys = IIf(Command() = "", "ȱʡ", Command())
        strObjs = "'ZL9DEMAND'"
        
        strSQL = "select Distinct M.���,M.���,nvl(M.�ϼ�,0) as �ϼ�,M.����,M.�̱���,M.���,M.˵��,nvl(M.ģ��,0) as ģ��,M.ϵͳ,nvl(M.ͼ��,0) ͼ��,upper(rtrim(P.����)) ����,Decode(upper(rtrim(P.����)),'ZL9REPORT',1,0) " & _
                " from (select level as ���,ID ���,�ϼ�ID �ϼ�,����,�̱���,���,˵��,ģ��,Nvl(ϵͳ,0) ϵͳ,ͼ��" & _
                "         from zlMenus " & IIf(BlnOnlySys, "", " where " & gstrSystems) & _
                "         start with ���='" & gstrMenuSys & "'" & " and �ϼ�ID is null" & _
                "         connect by prior ID=�ϼ�ID and ���='" & gstrMenuSys & "'" & ") M," & _
                "     (select distinct ID ���" & _
                "         from zlMenus" & _
                "         start with ģ�� in (" & _
                "             select ���" & _
                "             from zlPrograms" & _
                "             where ��� in (select distinct ��� from zlProgFuncs where (" & strCodes & "))" & _
                "             and " & gstrSystems & _
                "                  and (exists(select 1 from session_roles where role='DBA')" & _
                "                       or ϵͳ in (select ��� from zlSystems where upper(������)=user " & _
                "                                   Union" & _
                "                                   Select 0 From dual" & _
                "                                   Where Exists(select 1 from zlSystems where upper(������)=User))" & _
                "                       or ��� in (select ��� from zlRoleGrant G,session_roles S where G.��ɫ=S.Role" & " And " & gstrSystems1 & ")) " & _
                "                  and upper(rtrim(����)) in (" & strObjs & ")" & _
                "             and " & gstrSystems & _
                "       ) and ���='" & gstrMenuSys & "' and " & gstrSystems & _
                "         connect by prior �ϼ�ID=ID and ���='" & gstrMenuSys & "' and " & gstrSystems & ") G,"
        strSQL = strSQL & "" & _
                "     (Select Nvl(ϵͳ,0) ϵͳ,���,���� From zlPrograms " & _
                "             Where ��� in (select distinct ��� from zlProgFuncs where (" & strCodes & "))" & _
                "             and " & gstrSystems & _
                "                  and (exists(select 1 from session_roles where role='DBA')" & _
                "                       or ϵͳ in (select ��� from zlSystems where upper(������)=user " & _
                "                                   Union" & _
                "                                   Select 0 From dual" & _
                "                                   Where Exists(select 1 from zlSystems where upper(������)=User))" & _
                "                       or ��� in (select ��� from zlRoleGrant G,session_roles S where G.��ɫ=S.Role" & " And " & gstrSystems1 & ")) " & _
                "                  and upper(rtrim(����)) in (" & strObjs & ")" & _
                "             and " & gstrSystems & _
                ")P" & _
                " where M.���=G.��� and M.ģ��=P.���(+) And M.ϵͳ=P.ϵͳ(+) and M.ģ��=110 " & _
                " order by M.���,Decode(upper(rtrim(P.����)),'ZL9REPORT',1,0),M.���"

        If .State = adStateOpen Then .Close
        .Open strSQL, gcnOracle, adOpenKeyset
    End With
    
    Set MenuGranted = rsTemp
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
    Err = 0
    DoEvents
    With gcnOracle
        If .State = adStateOpen Then .Close
        .Provider = "MSDataShape"
        .Open "Driver={Microsoft ODBC for Oracle};Server=" & strServerName, strUserName, strUserPwd
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
                MsgBox "�����û�������������ָ�������޷�ע�ᡣ", vbInformation, gstrSysName
            End If
            
            OraDataOpen = False
            Exit Function
        End If
    End With
    
    Err = 0
    On Error GoTo ErrHand
    
    gstrDbUser = UCase(strUserName)
    SetDbUser gstrDbUser
    
    OraDataOpen = True
    Exit Function
    
ErrHand:
    If ErrCenter() = 1 Then Resume
    OraDataOpen = False
    Err = 0
End Function

Public Function OraDataClose() As Boolean
    '------------------------------------------------
    '���ܣ� �ر����ݿ�
    '������
    '���أ� �ر����ݿ⣬����True��ʧ�ܣ�����False
    '------------------------------------------------
    Err = 0
    On Error Resume Next
    gcnOracle.Close
    OraDataClose = True
    Err = 0

End Function

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

Public Function UpdatePassword(ByVal strUserName As String, ByVal strPasswd As String) As Boolean
    '-------------------------------------------------------------
    '���ܣ�����ԱID���޸�������
    '������CurrUser
    '      ��ǰ�û���
    '���أ�����ɹ����˻�True�����򷵻�False
    '-------------------------------------------------------------
    Err = 0
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

Public Function GetUserInfo(ByVal strSystems As String)
    Dim rsTmp As New ADODB.Recordset, rsUser As New ADODB.Recordset
    Dim strSQL As String, i As Integer
    '���û���Ϣ���蹫����������������ʹ��
    
    With rsTmp
        If .State = adStateOpen Then .Close
        strSQL = "Select S.*" & _
                " From zlSystems S,(Select Distinct owner From All_Tables Where Table_Name='���ű�') D" & _
                " Where Upper(S.������)=D.Owner And S.��� In (" & strSystems & ") Order by S.���"
        .Open strSQL, gcnOracle, adOpenKeyset
        If Not .EOF Then
            '��Ϊ���ܸ��û����ж��ϵͳ����ݣ�����ѭ��ȡ���
            glngUserId = 0 '��ǰ�û�id
            gstrUserCode = "" '��ǰ�û�����
            gstrUserName = "" '��ǰ�û�����
            gstrUserAbbr = "" '��ǰ�û�����
            glngDeptId = 0 '��ǰ�û�����id
            gstrDeptCode = "" '��ǰ�û�
            gstrDeptName = "" '��ǰ�û�
            
            For i = 1 To .RecordCount
                strSQL = "Select R.*,D.���� as ���ű���,D.���� as ��������,P.���,P.����,P.����" & _
                        " From " & !������ & ".�ϻ���Ա�� U," & !������ & ".��Ա�� P," & !������ & ".���ű� D," & !������ & ".������Ա R" & _
                        " Where U.��ԱID = P.ID And R.����ID = D.ID And P.ID=R.��ԱID and U.�û���=USER and R.ȱʡ=1"
                Set rsUser = New ADODB.Recordset
                rsUser.CursorLocation = adUseClient
                rsUser.Open strSQL, gcnOracle, adOpenKeyset
                If Not rsUser.EOF Then
                    glngUserId = rsUser!��ԱID '��ǰ�û�id
                    gstrUserCode = rsUser!��� '��ǰ�û�����
                    gstrUserName = IIf(IsNull(rsUser!����), "", rsUser!����) '��ǰ�û�����
                    gstrUserAbbr = IIf(IsNull(rsUser!����), "", rsUser!����) '��ǰ�û�����
                    glngDeptId = rsUser!����id '��ǰ�û�����id
                    gstrDeptCode = rsUser!���ű��� '��ǰ�û�
                    gstrDeptName = rsUser!�������� '��ǰ�û�
                    Exit For
                End If
                DoEvents
                .MoveNext
            Next
        End If
        .Close
    End With
End Function

Public Function OpenSQLRecord(ByVal strSQL As String, ByVal strTitle As String, ParamArray arrInput() As Variant) As ADODB.Recordset
'���ܣ�ͨ��Command����򿪴�����SQL�ļ�¼��
'������strSQL=�����а���������SQL���,������ʽΪ"[x]"
'             x>=1Ϊ�Զ��������,"[]"֮�䲻���пո�
'             ͬһ�������ɶദʹ��,�����Զ���ΪADO֧�ֵ�"?"����ʽ
'             ʵ��ʹ�õĲ����ſɲ�����,������Ĳ���ֵ��������(��SQL���ʱ��һ��Ҫ�õ��Ĳ���)
'      arrInput=���������Ĳ���ֵ,��������˳�����δ���,��������ȷ����
'      strTitle=����SQLTestʶ��ĵ��ô���/ģ�����
'���أ���¼����CursorLocation=adUseClient,LockType=adLockReadOnly,CursorType=adOpenStatic
'������
'SQL���Ϊ="Select ���� From ������Ϣ Where (����ID=[3] Or �����=[3] Or ���� Like [4]) And �Ա�=[5] And �Ǽ�ʱ�� Between [1] And [2] And ���� IN([6],[7])"
'���÷�ʽΪ��Set rsPati=OpenSQLRecord(strSQL, Me.Caption, CDate(Format(rsMove!ת������,"yyyy-MM-dd")),dtpʱ��.Value, lng����ID, "��%", "��", 20, 21)
    Static cmdData As New ADODB.Command
    Dim strPar As String, arrPar As Variant
    Dim lngLeft As Long, lngRight As Long
    Dim strSeq As String, intMax As Integer, i As Integer
    Dim strLog As String, varValue As Variant
    
    '�����Զ���[x]����
    lngLeft = InStr(1, strSQL, "[")
    Do While lngLeft > 0
        lngRight = InStr(lngLeft + 1, strSQL, "]")
        
        '������������"[����]����"
        strSeq = Mid(strSQL, lngLeft + 1, lngRight - lngLeft - 1)
        If IsNumeric(strSeq) Then
            i = CInt(strSeq)
            strPar = strPar & "," & i
            If i > intMax Then intMax = i
        End If
        
        lngLeft = InStr(lngRight + 1, strSQL, "[")
    Loop

    '�滻Ϊ"?"����
    strLog = strSQL
    For i = 1 To intMax
        strSQL = Replace(strSQL, "[" & i & "]", "?")
        
        '��������SQL���ٵ����
        varValue = arrInput(i - 1)
        Select Case TypeName(varValue)
        Case "Byte", "Integer", "Long", "Single", "Double", "Currency" '����
            strLog = Replace(strLog, "[" & i & "]", varValue)
        Case "String" '�ַ�
            strLog = Replace(strLog, "[" & i & "]", "'" & Replace(varValue, "'", "''") & "'")
        Case "Date" '����
            strLog = Replace(strLog, "[" & i & "]", "To_Date('" & Format(varValue, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')")
        Case "Variant" '����ȷ����
            strLog = Replace(strLog, "[" & i & "]", "?")
        End Select
    Next

    '���ԭ�в���:��Ȼ�����ظ�ִ��
    cmdData.CommandText = "" '��Ϊ����ʱ�����������
    Do While cmdData.Parameters.Count > 0
        cmdData.Parameters.Delete 0
    Loop
    
    '�����µĲ���
    arrPar = Split(Mid(strPar, 2), ",")
    For i = 0 To UBound(arrPar)
        varValue = arrInput((arrPar(i) - 1))
        Select Case TypeName(varValue)
        Case "Byte", "Integer", "Long", "Single", "Double", "Currency" '����
            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adVarNumeric, adParamInput, 30, varValue)
        Case "Date" '����
            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adDBTimeStamp, adParamInput, , varValue)
        Case "String" '�ַ�
            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adVarChar, adParamInput, 500, varValue)
        Case "Variant" '����ȷ����
        End Select
    Next

    'ִ�з��ؼ�¼��
    If cmdData.ActiveConnection Is Nothing Then
        Set cmdData.ActiveConnection = gcnOracle '���Ƚ���
    End If
    cmdData.CommandText = strSQL
    
    Call SQLTest(App.ProductName, strTitle, strLog)
    Set OpenSQLRecord = cmdData.Execute
    Call SQLTest
End Function




