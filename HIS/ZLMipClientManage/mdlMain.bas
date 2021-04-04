Attribute VB_Name = "mdlMain"
Option Explicit

Public ZlBrowerDll As Object                '����̨
Public SplashObj As New frmSplash
'Public gcnOracle As New ADODB.Connection    '�������ݿ�����

Public gstrSysName As String                'ϵͳ����
Public gstrVersion As String                'ϵͳ�汾
Public gstrAviPath As String                'AVI�ļ��Ĵ��Ŀ¼
Public gstrUserFlag As String               '��ǰ�û���־(��λ��ʾ)����1λ���Ƿ�DBA����2λ��ϵͳ������

Public gstrDbUser As String                 '��ǰ���ݿ��û�
Public glngUserId As Long                   '��ǰ�û�id
Public gstrUserCode As String               '��ǰ�û�����
Public gstrUserName As String               '��ǰ�û�����
Public gstrUserAbbr As String               '��ǰ�û�����
Public gstrServerName As String
Public glngDeptId As Long                   '��ǰ�û�����id
Public gstrDeptCode As String               '��ǰ�û����ű���
Public gstrDeptName As String               '��ǰ�û���������

Public gstrStation As String                '������վ����
Public gstrMenuSys As String                'ϵͳ�˵�
Public gstrCommand As String
Public gstrSystems As String

Public gobjFile As New FileSystemObject

Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Const SRCCOPY = &HCC0020
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, _
        ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
'---------------------------------------------------------------
'-ע��� API ����...
'---------------------------------------------------------------
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByRef lpSecurityAttributes As SECURITY_ATTRIBUTES, ByRef phkResult As Long, ByRef lpdwDisposition As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Public Declare Function SetActiveWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

'�л���ָ�������뷨��
Public Declare Function ActivateKeyboardLayout Lib "user32" (ByVal hkl As Long, ByVal flags As Long) As Long
'����ϵͳ�п��õ����뷨�����������뷨����Layout,����Ӣ�����뷨��
Public Declare Function GetKeyboardLayoutList Lib "user32" (ByVal nBuff As Long, lpList As Long) As Long
'��ȡĳ�����뷨������
Public Declare Function ImmGetDescription Lib "imm32.dll" Alias "ImmGetDescriptionA" (ByVal hkl As Long, ByVal lpsz As String, ByVal uBufLen As Long) As Long
'�ж�ĳ�����뷨�Ƿ��������뷨
Public Declare Function ImmIsIME Lib "imm32.dll" (ByVal hkl As Long) As Long



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
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
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

Public Enum REGISTER
    ע����Ϣ
    ˽��ģ��
    ˽��ȫ��
    ����ģ��
    ����ȫ��
End Enum

'---------------------------------------------------------------
'����ʱ�䣬�����ж�������Ļ�ĵȴ�ʱ��
'---------------------------------------------------------------
Public gdtStart As Long

'---------------------------------------------------------------
'   ��Ȩ���˵������ð汾
'---------------------------------------------------------------

'----------------------------------------------------------------------------------------------------------------------------------------------------------------
'����:������ؽ��̴����API����:2008-10-30 11:34:11:���˺�
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessID As Long) As Long
Private Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, ByVal lProcessId As Long) As Long
Private Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Public Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Type PROCESSENTRY32
      lSize             As Long
      lUsage            As Long
      lProcessId        As Long
      lDefaultHeapId    As Long
      lModuleId         As Long
      lThreads          As Long
      lParentProcessId  As Long
      lPriClassBase     As Long
      lFlags            As Long
      sExeFile          As String * 1024
End Type
Private Const PROCESS_TERMINATE = &H1
Public gcll_His_PId As Collection        '�洢��صĽ�����Ϣ:array(��������,PID,���ڸ���),"K"+������


Public zlCommFun As New zlMipClientComLib.clsCommFun
Public zlDataBase As New zlMipClientComLib.clsDatabase
Public zlComLib As New zlMipClientComLib.clsComLib
Public zlControl As New zlMipClientComLib.clsControl
Private mobjComLib As Object

Public gclsMsgSystem As New clsBusiness
Public gclsMsgOracle As New zlDataOracle.clsDataOracle

#Const SYS_TRYUSE = "��ʽ" '��ʽ/����

Public Sub Main()
    Dim lngReturn As Long
    Dim StrUnitName As String
    Dim BlnShowFlash As Boolean
    Dim strCode As String, intCount As Integer, strStyle As String
    Dim strTitle As String                  '��Ʒ����
    Dim strTag As String                    '�콢���־
    Dim rsMenu As ADODB.Recordset
    
'''    Call SetTNSNameFile 20110112-ZQ
    
     'Ϊʵ��XP�������ʾ����ǰ����ִ�иú���
    Call InitCommonControls

    BlnShowFlash = False
    If InStr(Command(), "=") <= 0 Then Load SplashObj
    
    '��ע����л�ȡ�û�ע�������Ϣ,����û���λ���Ʋ�Ϊ��,����ʾ���ִ���
    StrUnitName = GetSetting("ZLSOFT", "ע����Ϣ", "��λ����", "")
    gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "��ʾ", "")
    If StrUnitName <> "" And StrUnitName <> "-" Then
        gdtStart = Timer
        With SplashObj
            '��������Ҫ����
            Call zlComLib.ApplyOEM_Picture(.ImgIndicate, "Picture")
            Call zlComLib.ApplyOEM_Picture(.imgPic, "PictureB")
            If InStr(Command(), "=") <= 0 Then .Show
            .lblGrant = StrUnitName
            StrUnitName = GetSetting("ZLSOFT", "ע����Ϣ", "������", "")
            If Trim(StrUnitName) = "" Then
                .Label3.Visible = False
                .lbl������.Visible = False
            Else
                .lbl������.Caption = ""
                For intCount = 0 To UBound(Split(StrUnitName, ";"))
                    .lbl������.Caption = .lbl������.Caption & Split(StrUnitName, ";")(intCount) & vbCrLf
                Next
            End If
            .LblProductName = GetSetting("ZLSOFT", "ע����Ϣ", "��Ʒȫ��", "")
            If Len(.LblProductName) > 10 Then
                .LblProductName.FontSize = 15.75 '����
            Else
                .LblProductName.FontSize = 21.75 '����
            End If
            .lbl����֧���� = GetSetting("ZLSOFT", "ע����Ϣ", "����֧����", "")
            .lbltag = GetSetting("ZLSOFT", "ע����Ϣ", "��Ʒϵ��", "")
            
        End With
        Do
            If (Timer - gdtStart) > 3 Then Exit Do
            DoEvents
        Loop
        
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
    
    Call zlKillHISPID
    
    '�û�ע��
    If InStr(Command(), "=") > 0 Then
        Call frmUserLogin.Docmd(Command())
    Else
        frmUserLogin.Show 1
    End If
        
    If gclsMsgOracle.DatabaseState <> adStateOpen Then
        Unload frmUserLogin
        Unload SplashObj
        Exit Sub
    End If
    
    'д�뱾�����������EXE�ļ���
    Call SaveSetting("ZLSOFT", "����ȫ��", "ִ���ļ�", App.EXEName & ".exe")
    
    '��ʼ����������
    Call zlComLib.InitCommon(gclsMsgOracle.DatabaseConnection)
    
    Set mobjComLib = CreateObject("zl9Comlib.clsComlib")
    If Not (mobjComLib Is Nothing) Then
        Call mobjComLib.InitCommon(gclsMsgOracle.DatabaseConnection)
        If mobjComLib.RegCheck = False Then
            Unload SplashObj
            Exit Sub
        End If
    End If
    gstrSysName = mobjComLib.zlRegInfo("��Ʒ����") & "���"
    
'    SaveSetting "ZLSOFT", "ע����Ϣ", "��ʾ", gstrSysName
'    SaveSetting "ZLSOFT", "ע����Ϣ", UCase("gstrSysName"), gstrSysName
'    gstrVersion = App.Major & "." & App.Minor & "." & App.Revision
'    SaveSetting "ZLSOFT", "ע����Ϣ", UCase("gstrVersion"), gstrVersion
'    gstrAviPath = App.Path & "\�����ļ�"
'    SaveSetting "ZLSOFT", "ע����Ϣ", UCase("gstrAviPath"), gstrAviPath
'
'    strTag = ""
'    strTitle = zlComLib.zlRegInfo("��Ʒ����")
'    If strTitle <> "" Then
'        If InStr(strTitle, "-") > 0 Then
'            If Split(strTitle, "-")(1) = "Ultimate" Then
'                strTag = "�콢��"
'            ElseIf Split(strTitle, "-")(1) = "Professional" Then
'                strTag = "רҵ��"
'            End If
'        End If
'    End If
'    strTitle = Split(strTitle, "-")(0)
'    With SplashObj
'        If BlnShowFlash = False Then
'            .lblGrant = zlComLib.zlRegInfo("��λ����", , -1)
'            .lbl����֧����.Caption = zlComLib.zlRegInfo("����֧����", , -1)
'
'            .LblProductName = strTitle
'            .lbltag = strTag
'            strCode = zlComLib.zlRegInfo("��Ʒ������", , -1)
'            .lbl������.Caption = ""
'            For IntCount = 0 To UBound(Split(strCode, ";"))
'                .lbl������.Caption = .lbl������.Caption & Split(strCode, ";")(IntCount) & vbCrLf
'            Next
'            Call zlComLib.ApplyOEM_Picture(.ImgIndicate, "Picture")
'            If InStr(Command(), "=") <= 0 Then .Show
'            BlnShowFlash = True
'        End If
'        DoEvents
'    End With
    
'    '���û�ע�������Ϣд��ע���,���´�����ʱ��ʾ
'    SaveSetting "ZLSOFT", "ע����Ϣ", "��λ����", zlComLib.zlRegInfo("��λ����", , -1)
'    SaveSetting "ZLSOFT", "ע����Ϣ", "��Ʒȫ��", strTitle
'    SaveSetting "ZLSOFT", "ע����Ϣ", "��Ʒ����", zlComLib.zlRegInfo("��Ʒ����")
'    SaveSetting "ZLSOFT", "ע����Ϣ", "����֧����", zlComLib.zlRegInfo("����֧����", , -1)
'    SaveSetting "ZLSOFT", "ע����Ϣ", "������", zlComLib.zlRegInfo("��Ʒ������", , -1)
'    SaveSetting "ZLSOFT", "ע����Ϣ", "WEB֧���̼���", zlComLib.zlRegInfo("֧���̼���")
'    SaveSetting "ZLSOFT", "ע����Ϣ", "WEB֧��EMAIL", zlComLib.zlRegInfo("֧����MAIL")
'    SaveSetting "ZLSOFT", "ע����Ϣ", "WEB֧��URL", zlComLib.zlRegInfo("֧����URL")
'    SaveSetting "ZLSOFT", "ע����Ϣ", "��Ʒϵ��", strTag
        
    Unload SplashObj
    
    SplashObj.Tag = gstrSystems
            
            
    '����Ƿ�װ����Ϣ����ƽ̨ZLHIS�ͻ��ˣ����û�а�װ������뵯����װ����
    '------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String

    strSQL = "Select �к�,���� From zlRegInfo Where ��Ŀ='��Ϣ����ƽ̨�ͻ���'"
    Set rsTemp = gclsMsgOracle.OpenSQLRecord(strSQL, gstrSysName)
    If rsTemp.BOF = True Then
        '�޼�¼��ʾδ��װ
        If frmAppCreate.ShowDialog Then
            MsgBox "��Ϣ����ƽ̨�ͻ����Ѿ���װ�ɹ�������ʹ�ã�", vbInformation, gstrSysName
        Else
            MsgBox "��Ϣ����ƽ̨�ͻ����Ѿ�ȡ����װ�����ȷ�����Զ��˳���", vbInformation, gstrSysName
            Unload frmUserLogin
            Unload SplashObj
            Exit Sub
        End If
    End If
    
    '------------------------------------------------------------------------------------------------------------------
    Call frmNativateStart.SetEnvironment(gstrSysName, gstrVersion, gstrAviPath, _
                          gstrUserFlag, gstrDbUser, glngUserId, _
                          gstrUserCode, gstrUserName, gstrUserAbbr, _
                          glngDeptId, gstrDeptCode, gstrDeptName, _
                          gstrStation, gstrMenuSys, CStr(Command()))
    Call frmNativateStart.InitBrower(SplashObj, gclsMsgOracle.DatabaseConnection, rsMenu)
        
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
    Call gclsMsgOracle.UpdateUserPassword(strUserName, strPasswd)
'    gcnOracle.Execute "alter user " & strUserName & " identified by " & strPasswd
    UpdatePassword = True
    Exit Function
    
ErrorHand:
    If zlComLib.ErrCenter() = 1 Then Resume
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

Public Function ReadStartKey() As String
'���ܣ���ȡע�����������ʼʱ���־(֮һ��Ч����)
    Dim strKey As String
    strKey = GetKeyValue(HKEY_CURRENT_USER, "SOFTWARE\VTCELUS6CS", "IXPHWP")  'FirstStart,1Start
    If strKey = "" Then strKey = GetKeyValue(HKEY_LOCAL_MACHINE, "SOFTWARE\EG5PZRELSML", "NXPHWP") 'SecondStart,2Start
    If strKey = "" Then strKey = GetKeyValue(HKEY_USERS, ".DEFAULT\SOFTWARE\S1NM9US6CS", "TXPHWP") 'ThirdStart,3Start
    If strKey <> "" Then ReadStartKey = CStr(CDate(strKey))
End Function

Public Function WriteStartKey() As Boolean
'����:��ע�����д������ʼʱ���־
    Dim curDate As Date
    curDate = Format(Date, "yyyy-MM-dd")
    WriteStartKey = UpdateKey(HKEY_CURRENT_USER, "SOFTWARE\VTCELUS6CS", "IXPHWP", CCur(curDate)) 'FirstStart,1Start
    WriteStartKey = WriteStartKey And UpdateKey(HKEY_LOCAL_MACHINE, "SOFTWARE\EG5PZRELSML", "NXPHWP", CCur(curDate)) 'SecondStart,2Start
    WriteStartKey = WriteStartKey And UpdateKey(HKEY_USERS, ".DEFAULT\SOFTWARE\S1NM9US6CS", "TXPHWP", CCur(curDate)) 'ThirdStart,3Start
End Function

Public Function ReadValidKey() As String
'���ܣ���ȡע������������ڱ�־(֮һ��Ч����)
    Dim strKey As String
    strKey = GetKeyValue(HKEY_CURRENT_USER, "SOFTWARE\PZ7Q64F9", "IRSUTR") 'OneValid,1Valid
    If strKey = "" Then strKey = GetKeyValue(HKEY_LOCAL_MACHINE, "SOFTWARE\SDDQ64F9", "NRSUTR") 'TwoValid,2Valid
    If strKey = "" Then strKey = GetKeyValue(HKEY_USERS, ".DEFAULT\SOFTWARE\S1CKGZHPNO", "TRSUTR") 'ThreeValid,3Valid
    If strKey <> "" Then ReadValidKey = strKey
End Function

Public Function WriteValidKey() As Boolean
    '����:��ע�����д�������ڱ�־
    WriteValidKey = UpdateKey(HKEY_CURRENT_USER, "SOFTWARE\PZ7Q64F9", "IRSUTR", "Q64F9") 'OneValid,1Valid
    WriteValidKey = WriteStartKey And UpdateKey(HKEY_LOCAL_MACHINE, "SOFTWARE\SDDQ64F9", "NRSUTR", "Q64F9") 'TwoValid,2Valid
    WriteValidKey = WriteStartKey And UpdateKey(HKEY_USERS, ".DEFAULT\SOFTWARE\S1CKGZHPNO", "TRSUTR", "Q64F9") 'ThreeValid,3Valid
End Function

Public Function GetUserInfo(ByVal strSystems As String)
    Dim rsTmp As New ADODB.Recordset, rsUser As New ADODB.Recordset
    Dim strSQL As String, i As Integer
    '���û���Ϣ���蹫����������������ʹ��
'
'    With rsTmp
'        If .State = adStateOpen Then .Close
'        strSQL = "Select S.*" & _
'                " From zlSystems S,(Select Distinct owner From All_Tables Where Table_Name='���ű�') D" & _
'                " Where Upper(S.������)=D.Owner And S.��� In (" & strSystems & ") Order by S.���"
'        .Open strSQL, gcnOracle, adOpenKeyset
        
    Set rsTmp = gclsMsgSystem.GetSystemInfo(strSystems)
    With rsTmp
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
                
                
                Set rsUser = gclsMsgSystem.GetUserInfo(!������)
                
'                strSQL = "Select R.*,D.���� as ���ű���,D.���� as ��������,P.���,P.����,P.����" & _
'                        " From " & !������ & ".�ϻ���Ա�� U," & !������ & ".��Ա�� P," & !������ & ".���ű� D," & !������ & ".������Ա R" & _
'                        " Where U.��ԱID = P.ID And R.����ID = D.ID And P.ID=R.��ԱID and U.�û���=USER And (P.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or P.����ʱ�� Is Null) and R.ȱʡ=1"
'                Set rsUser = New ADODB.Recordset
'                rsUser.CursorLocation = adUseClient
'                rsUser.Open strSQL, gcnOracle, adOpenKeyset
                Set rsUser.ActiveConnection = Nothing
                If Not rsUser.EOF Then
                    glngUserId = rsUser!��ԱID '��ǰ�û�id
                    gstrUserCode = rsUser!��� '��ǰ�û�����
                    gstrUserName = IIf(IsNull(rsUser!����), "", rsUser!����) '��ǰ�û�����
                    gstrUserAbbr = IIf(IsNull(rsUser!����), "", rsUser!����) '��ǰ�û�����
                    glngDeptId = rsUser!����ID '��ǰ�û�����id
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

Private Function RunningInIDE() As Boolean
    '--����Ƿ�Դ���뻷��
    RunningInIDE = (App.EXEName = "zl9WizardMain")
End Function

'**********************************************************************************************************************
'����:���´�����ؽ��̵ĺ���
'����:���˺�
'����:2008-10-30 11:38:58
Public Function zlKillHISPID() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:ɱ������HIS����������Ƴ�����(ɱ��������:����ZLHIS+.exe�Ľ��������κδ���)
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-10-30 11:06:16
    '-----------------------------------------------------------------------------------------------------------
    Dim lngProcess As Long, i As Long
    
    zlKillHISPID = False
    Err = 0: On Error GoTo errHand:
    '��һ��:��Ҫ������ص�ZLHIS����ؽ���
    Set gcll_His_PId = New Collection
    If zlHISPidToCollect(gcll_His_PId) = False Then zlKillHISPID = True: Exit Function  '���������صĴ��󣬾�ֱ�ӷ�����
    If gcll_His_PId Is Nothing Then zlKillHISPID = True: Exit Function
    If gcll_His_PId.Count = 0 Then zlKillHISPID = True: Exit Function
    
    '�ڶ���:��Ҫ�������ZLHIS����ؽ��̵���ش��ڸ���,�����ź��жϳ���صĽ����Ƿ�����쳣,�����쳣�ģ��͵�ɱ��
    Call EnumWindows(AddressOf EnumWindowsProc, 0&)
    For i = 1 To gcll_His_PId.Count
        If Val(gcll_His_PId(i)(2)) <= 1 Then
            '�϶�������С��1����,��ô�϶����쳣����Ҫɱ����
            If Val(gcll_His_PId(i)(1)) <> 0 Then
                '����δ�ɹ������޴���������
                Call TerminatePID(Val(gcll_His_PId(i)(1)))
            End If
        End If
    Next
    zlKillHISPID = True
errHand:
End Function

Private Function EnumWindowsProc(ByVal hWnd As Long, ByVal lParam As Long) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:��ȡ���д��ڷ���HIS�Ľ��̵Ĵ���
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-10-30 10:26:02
    '-----------------------------------------------------------------------------------------------------------
    Dim strTittle As String, lngPID As Long, strName As String
    Dim lngCount As Long
    
    If GetParent(hWnd) = 0 Then
        '��ȡ hWnd ���Ӵ�����
        strTittle = String(80, 0)
        Call GetWindowText(hWnd, strTittle, 80)
        strTittle = Left(strTittle, InStr(strTittle, Chr(0)) - 1)
        If Trim(strTittle) <> "" Then
            Call GetWindowThreadProcessId(hWnd, lngPID)
            If IsWindowVisible(hWnd) Then
                Err = 0: On Error Resume Next
                strName = gcll_His_PId("K" & lngPID)(0)
                If Err = 0 Then
                    lngCount = Val(gcll_His_PId("K" & lngPID)(2)) + 1
                    gcll_His_PId.Remove "K" & lngPID
                    gcll_His_PId.Add Array(strName, lngPID, lngCount), "K" & lngPID
                End If
                Err.Clear: On Error GoTo 0
            End If
        End If
    End If
    EnumWindowsProc = True ' ��ʾ�����о� hWnd
    Exit Function
End Function

Private Function TerminatePID(ByVal lngPID As Long) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:����ָ���Ľ���
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-10-30 11:06:16
    '-----------------------------------------------------------------------------------------------------------
    Dim lngProcess As Long
    TerminatePID = False
    
    Err = 0: On Error GoTo errHand:
    lngProcess = OpenProcess(PROCESS_TERMINATE, 0&, lngPID)
    Call TerminateProcess(lngProcess, 1&)
    
    TerminatePID = True
errHand:
End Function

Private Function zlHISPidToCollect(ByRef cll_His_Pid As Collection) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:��ȡZLHIS�Ľ��̸���صļ���(gcll_HIS_Pid)
    '���:
    '����:cll_His_Pid-������HIS.exe�ĳ���װ�ظü�����
    '����:
    '����:���˺�
    '����:2008-10-30 10:07:38
    '-----------------------------------------------------------------------------------------------------------
    Dim strExeName  As String, lngSnapShot As Long, lngProcess As Long, lngCount  As Long
    Dim strCurExeName As String, lngCurPid As Long
    Dim uProcess   As PROCESSENTRY32
    Const TH32CS_SNAPPROCESS = &H2
    
    Err = 0: On Error GoTo errHand:
    strCurExeName = "*" & UCase(App.EXEName) & "*"
    
    lngCurPid = GetCurrentProcessId '��ȡ��ǰӦ�ó������
    lngSnapShot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
    If lngSnapShot <> 0 Then
        uProcess.lSize = Len(uProcess)
        lngProcess = ProcessFirst(lngSnapShot, uProcess)
        lngCount = 0
        Do While lngProcess
            '�����ڵ�ǰ���̵ĲŴ���
            If lngCurPid <> uProcess.lProcessId Then
                strExeName = UCase(Left(uProcess.sExeFile, InStr(1, uProcess.sExeFile, vbNullChar) - 1))
                If strExeName Like strCurExeName Then
                    cll_His_Pid.Add Array(strExeName, uProcess.lProcessId, 0), "K" & uProcess.lProcessId
                End If
            End If
            lngProcess = ProcessNext(lngSnapShot, uProcess)
        Loop
        CloseHandle (lngSnapShot)
    End If
    zlHISPidToCollect = True
    Exit Function
errHand:
End Function

''''**********************************************************************************************************************
'''
'''Private Function SetTNSNameFile() As Boolean
'''    Dim strFile As String, intFile As Integer
'''    Dim arrData() As Byte
'''
'''    On Error GoTo errHandle
'''
'''    strFile = GetKeyValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\VisualStudio\6.0\Setup\Microsoft Visual Basic", "ProductDir")
'''    If strFile <> "" Then
'''        If gobjFile.FolderExists(strFile) Then
'''            SetTNSNameFile = True
'''            Exit Function
'''        End If
'''    End If
'''
'''    strFile = GetKeyValue(HKEY_LOCAL_MACHINE, "SOFTWARE\ORACLE", "ORACLE_HOME")
'''    If Not gobjFile.FolderExists(strFile) Then '10G
'''        strFile = GetKeyValue(HKEY_LOCAL_MACHINE, "SOFTWARE\ORACLE", "ORA_CRS_HOME")
'''    End If
'''    If Not gobjFile.FolderExists(strFile) Then '10Gr2
'''        strFile = GetKeyValue(HKEY_LOCAL_MACHINE, "SOFTWARE\ORACLE\KEY_OraDb10g_home1", "ORACLE_HOME")
'''    End If
'''    If Not gobjFile.FolderExists(strFile) Then '10Gr2
'''        strFile = GetKeyValue(HKEY_LOCAL_MACHINE, "SOFTWARE\ORACLE\KEY_OraDb10g_home2", "ORACLE_HOME")
'''    End If
'''    If Not gobjFile.FolderExists(strFile) Then '10G ��ҵ��
'''        strFile = GetKeyValue(HKEY_LOCAL_MACHINE, "SOFTWARE\ORACLE\KEY_OraClient10g_home1", "ORACLE_HOME")
'''    End If
'''    strFile = strFile & "\network\admin\tnsnames.ora"
'''    If Not gobjFile.FileExists(strFile) Then Exit Function
'''    gobjFile.DeleteFile strFile
'''
'''    arrData = LoadResData(101, "CUSTOM")
'''    intFile = FreeFile
'''
'''    Open strFile For Binary As intFile
'''    Put intFile, , arrData()
'''    Close intFile
'''
'''    SetTNSNameFile = True
'''
'''    Exit Function
'''errHandle:
'''    'MsgBox "����" & Err.Number & vbCrLf & vbTab & Err.Description, vbExclamation, App.Title
'''End Function

Public Function SetRegister(ByVal enmRegister As REGISTER, ByVal strSection As String, ByVal strKey As String, ByVal strKeyValue As String) As Boolean
    '******************************************************************************************************************
    '���ܣ� ��ָ������Ϣ������ע�����
    '������ enmRegister-ע������
    '       strSection-ע���Ŀ¼
    '       strKey-����
    '       strKeyValue-��ֵ
    '���أ�
    '******************************************************************************************************************
    On Error GoTo errHand
    
    Select Case enmRegister
    Case ע����Ϣ
        
        Call SaveSetting("ZLSOFT", "ע����Ϣ\" & strSection, strKey, strKeyValue)
        
    Case ˽��ģ��

        Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & strSection, strKey, strKeyValue)
        
    Case ˽��ȫ��

        Call SaveSetting("ZLSOFT", "˽��ȫ��\" & gstrDbUser & "\" & strSection, strKey, strKeyValue)
        
    Case ����ģ��

        Call SaveSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & strSection, strKey, strKeyValue)
        
    Case ����ȫ��
        
        Call SaveSetting("ZLSOFT", "����ȫ��\" & strSection, strKey, strKeyValue)
        
    End Select
    
    SetRegister = True
    
errHand:
    
End Function

Public Function GetRegister(ByVal enmRegister As REGISTER, ByVal strSection As String, ByVal strKey As String, ByVal strDefKeyValue As String) As String
    '******************************************************************************************************************
    '���ܣ� ��ָ����ע����Ϣ��ȡ����
    '������ enmRegister-ע������
    '       strSection-ע���Ŀ¼
    '       strKey-����
    '       strDefKeyValue-ȱʡ��ֵ
    '���أ� strKeyValue-��ֵ
    '******************************************************************************************************************

    Dim strValue As String
    
    On Error GoTo errHand
    
    Select Case enmRegister
    Case ע����Ϣ
        
        strValue = GetSetting("ZLSOFT", "ע����Ϣ\" & strSection, strKey, strDefKeyValue)
        
    Case ˽��ģ��

        strValue = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & strSection, strKey, strDefKeyValue)
        
    Case ˽��ȫ��

        strValue = GetSetting("ZLSOFT", "˽��ȫ��\" & gstrDbUser & "\" & strSection, strKey, strDefKeyValue)
        
    Case ����ģ��

        strValue = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & strSection, strKey, strDefKeyValue)
        
    Case ����ȫ��
        
        strValue = GetSetting("ZLSOFT", "����ȫ��\" & strSection, strKey, strDefKeyValue)
        
    End Select
    
    GetRegister = strValue
    
errHand:
End Function
