Attribute VB_Name = "mdlMain"
Option Explicit

Public SplashObj As New frmSplash
Public gcnOracle As New adodb.Connection    '�������ݿ�����

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

Public gstrPrivs As String                   '��ǰ�û����еĵ�ǰģ��Ĺ���
Public gstr��λ���� As String
Public glngSys As Long
Public gdtStart As Long
Public gblnEmerge As Boolean                '�Ƿ����ּ��� 2008-12-24
Public gblnClearData As Boolean             '�Ƿ������־
Public gstr�������� As String               '���汾������������

Public gobjRegister As Object               'ע����Ȩ����zlRegister

Public Type T��������
    ID      As Long
    ����    As Integer  '0-COM�ڷ�ʽ 1-IP��ʽ
    COM��   As Integer
    ������  As Long
    ����λ  As String
    У��λ  As String
    ֹͣλ  As String
    ����    As String
    IP�˿�  As Long
    IP      As String
    ����     As Long
    �ַ�ģʽ As String
    SaveAsID As Long
    �������� As String
    �Զ�Ӧ�� As String  '�Զ�Ӧ��������λ�룬Ϊ<=0ʱ�����á�
    �ɷ��Ѻ˱걾 As Long '>0���Է� ,0<=�����Է�
    ͨѶĿ¼ As String  '���ճ���Ĵ��Ŀ¼
    ͨѶ���� As String  'ͨѶ������
    �Զ������ As String
    �Զ������ʿ� As Integer '0-�����㣬1-Ҫ����
    ���Ϊͨ���� As Integer '0-�����Ϊ����ȡ��Ĭ�ϣ���1-��������ȡ
End Type
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
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

'---------------------------------------------------------------
'-ע��� API ����...
'---------------------------------------------------------------
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByRef lpSecurityAttributes As SECURITY_ATTRIBUTES, ByRef phkResult As Long, ByRef lpdwDisposition As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
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

'--- ��������
Public gobjComLib As Object
Public gobjCommFun As Object
Public gobjControl As Object
Public gobjDatabase As Object
Public gobjPrintMode As Object
Public g����() As T��������

'---- ��ֹ�����õ�API
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Const PROCESS_QUERY_INFORMATION = &H400
Const STILL_ALIVE = &H103
'----------------------------------

Public mstrConn As String '���Ӵ��������Զ���������
Public gstr�������� As String '""-������,0-��ֹ,>0 ��������

'---------------------------------------------------------------
'   ��Ȩ���˵������ð汾
'---------------------------------------------------------------
Public Sub Main()
    Dim lngReturn As Long
    Dim StrUnitName As String
    Dim BlnShowFlash As Boolean
    Dim strCode As String, IntCount As Integer, StrStyle As String
    Dim rsMenu As adodb.Recordset, StrHaveSys As String
    Dim strTitle As String, strTag As String
    Dim objLogin As Object

    
    
    'Ϊʵ��XP�������ʾ����ǰ����ִ�иú���
    
    Call InitCommonControls
    '������������
    If gobjComLib Is Nothing Then Set gobjComLib = GetObject("", "zl9Comlib.clsComlib")
    If gobjCommFun Is Nothing Then Set gobjCommFun = GetObject("", "zl9Comlib.clsCommfun")
    If gobjControl Is Nothing Then Set gobjControl = GetObject("", "zl9Comlib.clsControl")
    If gobjDatabase Is Nothing Then Set gobjDatabase = GetObject("", "zl9Comlib.clsDatabase")
    If gobjPrintMode Is Nothing Then Set gobjPrintMode = GetObject("", "zl9PrintMode.zlPrintMethod")
    
    BlnShowFlash = False
    Load SplashObj
    '��ע����л�ȡ�û�ע�������Ϣ,����û���λ���Ʋ�Ϊ��,����ʾ���ִ���
    StrUnitName = GetSetting("ZLSOFT", "ע����Ϣ", "��λ����", "")
    gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "��ʾ", "")

    If StrUnitName <> "" Then
        With SplashObj
            '��������Ҫ����
            Call gobjComLib.ApplyOEM_Picture(.ImgIndicate, "Picture")
            Call gobjComLib.ApplyOEM_Picture(.imgPic, "PictureB")
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
            .lbltag = GetSetting("ZLSOFT", "ע����Ϣ", "��Ʒϵ��", "")
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
    '����ע�Ჿ��(���ڵ�¼ʱ��ȡ���Ӷ���)
    On Error Resume Next
    Set gobjRegister = CreateObject("zlRegister.clsRegister")
    If gobjRegister Is Nothing Then
        Err.Clear
        MsgBox "����zlRegister��������ʧ��,�����ļ��Ƿ���ڲ�����ȷע�ᡣ", vbExclamation, gstrSysName
        Unload SplashObj
        Exit Sub
    End If
    On Error GoTo 0
    '�û�ע��
'    frmUserLogin.Show 1
    '���õ�½����
    
    If objLogin Is Nothing Then
        Set objLogin = CreateObject("ZLLogin.clsLogin")
    End If
    If objLogin Is Nothing Then
        MsgBox "����ZLLogin��������ʧ��,�����ļ��Ƿ���ڲ�����ȷע�ᡣ"
        Exit Sub
    Else
        Set gcnOracle = objLogin.Login(2, CStr(Command()))
        If gcnOracle Is Nothing Then
            Exit Sub
        ElseIf gcnOracle.State <> adStateOpen Then '��ֹgcnOracle��New�ķ�ʽ�����ġ�
            Exit Sub
        End If
    End If

    
    If gcnOracle.State <> adStateOpen Then
'        Unload frmUserLogin
'        Unload SplashObj
        Exit Sub
    End If
    
    '��ʼ����������

    
    gobjComLib.InitCommon gcnOracle

    
    '�����������Ч��Ϊ�ջ�Ϊ"-"�������˳�
    gstrParsePublish = gobjComLib.zlRegInfo("��Ʒ����")
    gstrParseRegCode = gobjComLib.zlRegInfo("��λ����", , -1)
    
    gstrSysName = gstrParsePublish & "���"
    SaveSetting "ZLSOFT", "ע����Ϣ", "��ʾ", gstrSysName
    SaveSetting "ZLSOFT", "ע����Ϣ", UCase("gstrSysName"), gstrSysName
    gstrVersion = App.major & "." & App.minor & "." & App.Revision
    SaveSetting "ZLSOFT", "ע����Ϣ", UCase("gstrVersion"), gstrVersion
    gstrAviPath = App.Path & "\�����ļ�"
    SaveSetting "ZLSOFT", "ע����Ϣ", UCase("gstrAviPath"), gstrAviPath
    
    gstr�������� = gobjComLib.zlRegInfo("������������")
    
    strTag = ""
    strTitle = gobjComLib.zlRegInfo("��Ʒ����")
    If strTitle <> "" Then
        If InStr(strTitle, "-") > 0 Then
            If Split(strTitle, "-")(1) = "Ultimate" Then
                strTag = "�콢��"
            ElseIf Split(strTitle, "-")(1) = "Professional" Then
                strTag = "רҵ��"
            End If
        End If
    End If
    strTitle = Split(strTitle, "-")(0)
    
    With SplashObj
        If BlnShowFlash = False Then
            .lblGrant = gstrParseRegCode
            .lbl����֧����.Caption = gobjComLib.zlRegInfo("����֧����", , -1)
            .LblProductName = strTitle
            .lbltag = strTag
            
            strCode = gobjComLib.zlRegInfo("��Ʒ������", , -1)
            .lbl������.Caption = ""
            For IntCount = 0 To UBound(Split(strCode, ";"))
                .lbl������.Caption = .lbl������.Caption & Split(strCode, ";")(IntCount) & vbCrLf
            Next
            Call gobjComLib.ApplyOEM_Picture(.ImgIndicate, "Picture")
            .Show
            BlnShowFlash = True
        End If
        DoEvents
    End With
    
    '���û�ע�������Ϣд��ע���,���´�����ʱ��ʾ

    SaveSetting "ZLSOFT", "ע����Ϣ", "��λ����", gstrParseRegCode
    SaveSetting "ZLSOFT", "ע����Ϣ", "��Ʒȫ��", strTitle
    SaveSetting "ZLSOFT", "ע����Ϣ", "��Ʒ����", gobjComLib.zlRegInfo("��Ʒ����")
    SaveSetting "ZLSOFT", "ע����Ϣ", "����֧����", gobjComLib.zlRegInfo("����֧����", , -1)
    SaveSetting "ZLSOFT", "ע����Ϣ", "������", gobjComLib.zlRegInfo("��Ʒ������", , -1)
    SaveSetting "ZLSOFT", "ע����Ϣ", "WEB֧���̼���", gobjComLib.zlRegInfo("֧���̼���")
    SaveSetting "ZLSOFT", "ע����Ϣ", "WEB֧��EMAIL", gobjComLib.zlRegInfo("֧����MAIL")
    SaveSetting "ZLSOFT", "ע����Ϣ", "WEB֧��URL", gobjComLib.zlRegInfo("֧����URL")
    SaveSetting "ZLSOFT", "ע����Ϣ", "��Ʒϵ��", strTag
    '�����ס�ZYB��2001-09-19�޸�
    '-------------------------------------------------------------
    '��鱾����װ����
    '-------------------------------------------------------------
    '-------------------------------------------------------------
    '��������ѡ����
    '-------------------------------------------------------------
    gstrSystems = " (ϵͳ =100 Or ϵͳ Is NULL)"
    
    '-------------------------------------------------------------
    '�����˵�������
    '-------------------------------------------------------------
'    Set rsMenu = MenuGranted
'    If rsMenu.EOF Then
'        MsgBox "��û�в����κ�ϵͳ��Ȩ��,�������˳���", vbInformation, gstrSysName
'        Unload SplashObj
'        Exit Sub
'    End If
    '-------------------------------------------------------------
    '����ͬ���
    '-------------------------------------------------------------
    
    glngSys = 100
    Call CreateSynonyms(glngSys, 1208)
    
    gblnFromDB = IsFromDb
    
    If gblnFromDB Then
        gblnEmerge = gobjDatabase.GetPara("����걾", glngSys, 1208, 0)
    Else
        gblnEmerge = Val(GetSetting("ZLSOFT", "����ģ��\zl9LisWork\frmLabMain", "����걾", 0))
    End If
    '-------------------------------------------------------------
    'ѡ����ò�ͬ��񵼺�̨
    '-------------------------------------------------------------
    On Error Resume Next
    Err = 0
    
    Unload SplashObj
    
    CodeMan 1208
End Sub

Public Sub CodeMan(ByVal lngModul As Long)
    '------------------------------------------------
    '���ܣ� ����������ָ�����ܣ�����ִ����س���
    '������
    '   lngModul:��Ҫִ�еĹ������
    '���أ�
    '------------------------------------------------
    Dim clsPublic As New clsPublic
    gstrAviPath = GetSetting(appName:="ZLSOFT", Section:="ע����Ϣ", Key:=UCase("gstrAviPath"), Default:="")
    gstrSysName = GetSetting(appName:="ZLSOFT", Section:="ע����Ϣ", Key:=UCase("gstrSysName"), Default:="")
    gstrVersion = GetSetting(appName:="ZLSOFT", Section:="ע����Ϣ", Key:=UCase("gstrVersion"), Default:="")
    gstr��λ���� = gobjComLib.GetUnitName()
    Call GetUserInfo
    
    gstrPrivs = gobjComLib.GetPrivFunc(glngSys, lngModul)
    '-------------------------------------------------
    
    Select Case lngModul
        Case 1208
            clsPublic.InitClsPublic
    End Select
End Sub

Private Function CreateSynonyms(ByVal lngSys As Long, ByVal lngModul As Long)
    Dim strSQL As String
    '����ģ����������ͬ���(����Ѵ����򲻻��ٴ���)
    On Error Resume Next
    strSQL = "Zl_Createsynonyms(" & lngSys & ")"
    gobjDatabase.ExecuteProcedure strSQL, "����ͬ���"
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
    On Error GoTo errHand
    mstrConn = gcnOracle.ConnectionString
    gstrDbUser = UCase(strUserName)
    gobjComLib.SetDbUser gstrDbUser
    OraDataOpen = True
    Exit Function
    
errHand:
    If gobjComLib.ErrCenter() = 1 Then Resume
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
    If gobjComLib.ErrCenter() = 1 Then Resume
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

Public Sub CheckDBConnect()
    On Error GoTo ConnErr
    If gcnOracle.State <> 1 Then gcnOracle.Open
    gcnOracle.Execute "select '����'  from dual"
    Exit Sub
ConnErr:
    On Error Resume Next
    If gcnOracle.State = 1 Then
        gcnOracle.Close
    End If
End Sub
Public Sub GetUserInfo()
'����:�õ��û�����Ϣ

    Dim rsTemp As New adodb.Recordset
    On Error GoTo errHand
    glngUserId = 0
    gstrUserCode = ""
    gstrUserName = ""
    gstrUserAbbr = ""
    glngDeptId = 0
    gstrDeptCode = ""
    gstrDeptName = ""
    
    Set rsTemp = gobjDatabase.GetUserInfo
    
    Do Until rsTemp.EOF
        glngUserId = Val("" & rsTemp.Fields("ID").Value)               '��ǰ�û�id
        gstrUserCode = "" & rsTemp.Fields("���").Value            '��ǰ�û�����
        gstrUserName = "" & rsTemp.Fields("����").Value            '��ǰ�û�����
        gstrUserAbbr = "" & rsTemp.Fields("����").Value          '��ǰ�û�����
        glngDeptId = Val("" & rsTemp.Fields("����id").Value)            '��ǰ�û�����id
        gstrDeptCode = "" & rsTemp.Fields("������").Value        '��ǰ�û�
        gstrDeptName = "" & rsTemp.Fields("������").Value        '��ǰ�û�
    
        rsTemp.MoveNext
    Loop
    Exit Sub
errHand:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
    Err = 0
End Sub


Private Function IsFromDb() As Boolean
    '�Ƿ�����ݿ��ȡ����
    Dim strSQL As String, rsTmp As New adodb.Recordset
    Dim strSet As String
    
    Dim aPorts As Variant, i As Integer, lngID As Long
    On Error GoTo errH
    
    strSQL = "Select ��� as ϵͳ From zlSystems Where Trunc(���/100)=1 And �汾�� >= '10.24.0'"
    Set rsTmp = gcnOracle.Execute(strSQL)
    Do Until rsTmp.EOF
        IsFromDb = True
        rsTmp.MoveNext
    Loop
    
    If IsFromDb Then
        '���ϵͳ�Ƿ��в��������û�У��ӱ�����ע����ж���
        
        strSet = Trim(gobjDatabase.GetPara("������������", glngSys, 1208, ""))
        If strSet = "" Then
            Err = 0: On Error Resume Next
            aPorts = GetAllSettings("ZLSOFT", "����ģ��\ZlLISSrv")
            On Error GoTo errH
            
            If Not IsEmpty(aPorts) Then
                ReDim g����(UBound(aPorts))
                
                For i = LBound(aPorts) To UBound(aPorts)
                    lngID = Val(GetSetting("ZLSOFT", "����ģ��\ZlLISSrv\" & aPorts(i, 0), "Device", 0))
                    If lngID > 0 Then
                        If aPorts(i, 0) Like "COM*" Then
                            g����(i).���� = 0
                            g����(i).COM�� = Val(Replace(aPorts(i, 0), "COM", ""))

                            g����(i).�ַ�ģʽ = Val(GetSetting("ZLSOFT", "����ģ��\ZlLISSrv\" & aPorts(i, 0), "InputMode", "0"))
                        Else
                            g����(i).���� = 1
                            g����(i).COM�� = 0
                            g����(i).�ַ�ģʽ = Val(GetSetting("ZLSOFT", "����ģ��\ZlLISSrv\" & aPorts(i, 0), "InMode", "0"))
                        End If
                        
                        With g����(i)
                            .ID = lngID
                            g����(i).������ = Val(GetSetting("ZLSOFT", "����ģ��\ZlLISSrv\" & aPorts(i, 0), "Speed", "9600"))
                            g����(i).����λ = Val(GetSetting("ZLSOFT", "����ģ��\ZlLISSrv\" & aPorts(i, 0), "DataBit", "8"))
                            g����(i).У��λ = GetSetting("ZLSOFT", "����ģ��\ZlLISSrv\" & aPorts(i, 0), "Parity", "N")
                            g����(i).ֹͣλ = Val(GetSetting("ZLSOFT", "����ģ��\ZlLISSrv\" & aPorts(i, 0), "StopBit", "1"))
                            g����(i).���� = Val(GetSetting("ZLSOFT", "����ģ��\ZlLISSrv\" & aPorts(i, 0), "HandShaking", "0"))
                            .IP�˿� = Val(GetSetting("ZLSOFT", "����ģ��\ZlLISSrv\" & aPorts(i, 0), "Port", "6666"))
                            .IP = GetSetting("ZLSOFT", "����ģ��\ZlLISSrv\" & aPorts(i, 0), "IP", "127.0.0.1")
                            .SaveAsID = Val(GetSetting("ZLSOFT", "����ģ��\ZlLISSrv\" & aPorts(i, 0), "SaveAs", "0"))
                            .���� = Val(GetSetting("ZLSOFT", "����ģ��\ZlLISSrv\" & aPorts(i, 0), "Host", "0"))
                            
                            .�Զ�Ӧ�� = GetSetting("ZLSOFT", "����ģ��\ZlLISSrv\" & aPorts(i, 0), "Auto", "0")
                            .�ɷ��Ѻ˱걾 = GetSetting("ZLSOFT", "����ģ��\ZlLISSrv\" & aPorts(i, 0), "blnSend", "1")
                        End With
                    End If
                Next
                
                gblnFromDB = True
                Call SavePortsSetting
            End If
        End If
    End If
    
    Exit Function
errH:
End Function

Public Function KillProc(ByVal strFileName As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:��ָֹ���ĳ���
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:�¶�
    '����:2010-06-02
    '-----------------------------------------------------------------------------------------------------------
    Dim pid As Long, hProcess As Long, ExitCode As Long
    
    pid = Shell("taskkill.exe /im " & strFileName & " /f", vbHide)
    hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, 0, pid)
    Do
        Call GetExitCodeProcess(hProcess, ExitCode)
        DoEvents
    Loop While ExitCode = STILL_ALIVE
    Call CloseHandle(hProcess)
    KillProc = True
End Function



