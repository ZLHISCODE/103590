Attribute VB_Name = "MdlBrower"
Option Explicit
'MDI����
Public Type Menu_Type
    ���ܲ˵� As Long
    ���ڲ˵� As Long
    �������ܲ˵� As Long
    �ָ��˵� As Long
End Type
Public �˵���׼ As Menu_Type
Public Enum �����嵥
    ���������嵥 = 10
    �ֵ������ = 11
    ��Ϣ�շ����� = 12
    ϵͳѡ������ = 13
    EXCEL������ = 14
    ���ز������� = 15
End Enum
'��ҹ���
Public gobjPlugIn As Object

Public gobjRelogin As Object                   '���������
Public FrmMainface As Form
Public gcnOracle As ADODB.Connection

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

Public gstrObj() As String
Public gobjCls() As Object
Public grsMenus As New ADODB.Recordset       '�˵���¼��
Public gstrMenuSys As String                '�˵�����
Public gstrCommand As String                '�����в��� �¶� 2010-12-06
Private mlngSysPre As Long                  '�ϴε���˽��ͬ��ʼ�鴴��ʱ��ϵͳ��
Private mlngWin32 As Long
Private mblnע�� As Boolean

Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessID As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Const Process_Query_Information = &H400
Private Const Still_Active = &H103
'---------------------------------------------------------------------------------------------------
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'�ر�ϵͳ��صı�����API����
'----------------------------------------------------------------------------------------------------
Public Type LUID
    UsedPart As Long
    IgnoredForNowHigh32BitPart As Long
End Type
Public Type TOKEN_PRIVILEGES
    PrivilegeCount As Long
    TheLuid As LUID
    Attributes As Long
End Type
Public Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Public Declare Function GetVersion Lib "kernel32" () As Long
'The GetCurrentProcess function returns a pseudohandle for the current process.
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long
'The OpenProcessToken function opens the access token associated with a process.
Public Declare Function OpenProcessToken Lib "advapi32" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
'The LookupPrivilegeValue function retrieves the locally unique identifier (LUID) used on a specified system to locally represent the specified privilege name.
Public Declare Function LookupPrivilegeValue Lib "advapi32" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LUID) As Long
'The AdjustTokenPrivileges function enables or disables privileges in the specified access token. Enabling or disabling privileges in an access token requires TOKEN_ADJUST_PRIVILEGES access.
Public Declare Function AdjustTokenPrivileges Lib "advapi32" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long
Public Declare Sub SetLastError Lib "kernel32" (ByVal dwErrCode As Long)
Public Declare Function GetLastError Lib "kernel32" () As Long
'����ExitWindowsEx
Private Const M_lng�رռ��������Դ As Long = 8
Public Const EWX_FORCE = 4 'ǿ�йرճ���ע��
'�Զ���
Public Const WINDOWS95 = 0
Public Const WINDOWSNT = 1

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
'����������ڼ���Ƿ�Ϸ�����
Public Declare Function GlobalAddAtom Lib "kernel32" Alias "GlobalAddAtomA" (ByVal lpString As String) As Integer
Public Declare Function GlobalDeleteAtom Lib "kernel32" (ByVal nAtom As Integer) As Integer

'Command�������
Public Const INFINITE As Long = &HFFFF&
Private Const SW_HIDE                           As Integer = 0 '���ش��ڣ�������һ������
Private Const NORMAL_PRIORITY_CLASS             As Long = &H20&
Public Const STARTF_USESTDHANDLES = &H100&
Public Const STARTF_USESHOWWINDOW = &H1
Private Const ERROR_EXTENDED_ERROR          As Long = 1208
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200

Public Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessID As Long
    dwThreadID As Long
End Type
Public Type STARTUPINFO
    Cb As Long
    lpReserved As Long
    lpDesktop As Long
    lpTitle As Long
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type
'ע���ȫ��������
Public Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Boolean
End Type

Public Declare Function CreatePipe Lib "kernel32" (phReadPipe As Long, phWritePipe As Long, lpPipeAttributes As SECURITY_ATTRIBUTES, ByVal nSize As Long) As Long
Public Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, lpProcessAttributes As SECURITY_ATTRIBUTES, lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDriectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Public Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Public Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As String, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function FormatMessage Lib "kernel32.dll" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Private Declare Function WNetGetLastError Lib "mpr.dll" Alias "WNetGetLastErrorA" (lpError As Long, ByVal lpErrorBuf As String, ByVal nErrorBufSize As Long, ByVal lpNameBuf As String, ByVal nNameBufSize As Long) As Long


Public Sub ExecuteFunc(lngSys As Long, Components As String, Modul As Long, Optional ByVal strPara As String) ', Identity As Byte
    '-------------------------------------------------------------
    '���ܣ�����ִ��ָ�������Ĺ��ܳ���
    '������
    '   frmbrower��������
    '   Components������
    '   Modul��ģ����
    '   Identity����ִ�������Ҫ��
    '-------------------------------------------------------------
    Dim rsCheck As New ADODB.Recordset                  '���汾�Ƿ����ϵͳ����
    Dim IntCount As Integer, intClients As Integer
    Dim objNow As Object                                '�����Ĳ�������
    Dim BlnExecute As Boolean                           '�Ƿ���ڸò���
    Dim StrVersion As String, StrCompareVersion As String
    Dim ArrayVersion
    '�Ϸ��Լ��
    Dim intAtom As Integer, strCommon As String
    Dim strSQL  As String
    
    Err = 0: On Error Resume Next
    FrmMainface.MousePointer = 11
    
    IntCount = UBound(gstrObj)
    If Err <> 0 Then IntCount = -1
    Err = 0
    
    BlnExecute = False
    If IntCount >= 0 Then
        For IntCount = 0 To UBound(gstrObj)
            If gstrObj(IntCount) = Components Then
                BlnExecute = True
                Exit For
            End If
        Next
    End If
    
    'ʹ���²�������
    If UCase(Components) = UCase("zl9EmrInterface") And BlnExecute = False Then
        IntCount = UBound(gstrObj)
        IntCount = IntCount + 1
        ReDim Preserve gstrObj(IntCount)
        gstrObj(IntCount) = Components
        If FrmMainface.mobjEmr Is Nothing Then
            MsgBox "�����������ʧ�ܣ����鲢���µ�¼��", vbInformation, gstrSysName
            Exit Sub
        ElseIf FrmMainface.mobjEmr.IsInited = False Then
            MsgBox "�������δ�ܳ�ʼ��," & FrmMainface.mobjEmr.GetError(), vbInformation, gstrSysName
            Exit Sub
        End If
        If Not gobjRelogin.IsEMRProxy Then 'ʹ�ô����û���¼���򲻼��Ȩ��
            Dim strSpecify As String 'Ƭ�Σ�����Ȩ�޹̶��ڵ���ǰ����
            If Not FrmMainface.mobjEmr.HasInjectAuthorization(2201) Then
                strSpecify = GetPrivFunc(lngSys, 2201)
                Call FrmMainface.mobjEmr.InjectAuthorization(2201, strSpecify)
            End If
            If Not FrmMainface.mobjEmr.HasInjectAuthorization(2203) Then
                strSpecify = GetPrivFunc(lngSys, 2203)
                Call FrmMainface.mobjEmr.InjectAuthorization(2203, strSpecify)
            End If
        End If
        BlnExecute = True
    End If
    '--���û�иò���,�򴴽�--
    If BlnExecute = False Then
        Set objNow = CreateObject(Components & ".Cls" & Mid(Components, 4))
    
        If Err = 0 Then
            On Error GoTo errH
            '--���ò����İ汾�Ƿ�����ϵͳ����(���汾-3;�ΰ汾-3;���汾-3)[�Զ��屨��������]--
            If Not (UCase(Components) = "ZL9REPORT") And Not (UCase(Components) = "ZL9DOC") And Not OS.IsDesinMode Then
                strSQL = " Select nvl(���汾,1) ���汾,nvl(�ΰ汾,0) �ΰ汾,nvl(���汾,0) ���汾,���� " & _
                          " From ZlComponent Where Upper(Rtrim(����))=[1] And ϵͳ=[2]"
                Set rsCheck = zlDatabase.OpenSQLRecord(strSQL, "�����汾���", UCase(Components), lngSys)
                With rsCheck
                    If .EOF Then
                        MsgBox "ϵͳ������ZlComponent���ݲ����������������������ϵ��", vbInformation, gstrSysName
                        FrmMainface.MousePointer = 0
                        Exit Sub
                    End If
                    
                    '��װ�汾��Ϊ��λ���汾����λ�ΰ汾����λ���汾
                    StrCompareVersion = String(3 - Len(!���汾), "0") & !���汾 & "." & _
                                        String(3 - Len(!�ΰ汾), "0") & !�ΰ汾 & "." & _
                                        String(3 - Len(!���汾), "0") & !���汾
                    ArrayVersion = Split(objNow.Version, ".")
                    StrVersion = String(3 - Len(ArrayVersion(0)), "0") & ArrayVersion(0) & "." & _
                                 String(3 - Len(ArrayVersion(1)), "0") & ArrayVersion(1) & "." & _
                                 String(3 - Len(ArrayVersion(2)), "0") & ArrayVersion(2)
                    
                    If StrVersion < StrCompareVersion Then
                        MsgBox "�ò����İ汾�Ѳ�������ϵͳ���������������������ϵ����" & !���� & "��", vbInformation, gstrSysName
                        FrmMainface.MousePointer = 0
                        Exit Sub
                    End If
                End With
            End If
        
            IntCount = 0
            On Error Resume Next
            IntCount = UBound(gstrObj)
            IntCount = IntCount + 1
            If Err.Number <> 0 Then Err.Clear
            On Error GoTo errH
            ReDim Preserve gobjCls(IntCount)
            Set gobjCls(IntCount) = objNow
            ReDim Preserve gstrObj(IntCount)
            gstrObj(IntCount) = Components
        '��������ʧ�ܣ�Ӧ����ʾ
        Else
            Screen.MousePointer = 0
            MsgBox "���� " & Components & ".Cls" & Mid(Components, 4) & " �����������������鰲װ�Ƿ���ȷ����Ϣ��" & vbNewLine & Err.Description, vbExclamation, gstrSysName
            Err.Clear
            Exit Sub
        End If
    End If
    
    Err = 0: On Error GoTo errH
    '--ִ�иù���--
    If gstrObj(IntCount) = Components Then
        If UCase(Components) = "ZL9REPORT" Then
            If Modul = �˵���׼.�������ܲ˵� Then
                gobjCls(IntCount).ReportMan gcnOracle, FrmMainface
            Else
                
'                strPara = "��ʼ����=2013-01-01"
                If strPara <> "" Then
                    Dim varPara As Variant
                                        
                    varPara = Split(strPara, "|")
'                    varPara(0) = "��ʼ����=2013-01-01"
'                    varPara(1) = "��������=2014-05-01"
                    
                    '���֧��10������������10���Ĳ���
                    Select Case UBound(varPara)
                    Case 0
                        gobjCls(IntCount).ReportOpen gcnOracle, lngSys, Modul, FrmMainface, CStr(varPara(0))
                    Case 1
                        gobjCls(IntCount).ReportOpen gcnOracle, lngSys, Modul, FrmMainface, CStr(varPara(0)), CStr(varPara(1))
                    Case 2
                        gobjCls(IntCount).ReportOpen gcnOracle, lngSys, Modul, FrmMainface, CStr(varPara(0)), CStr(varPara(1)), CStr(varPara(2))
                    Case 3
                        gobjCls(IntCount).ReportOpen gcnOracle, lngSys, Modul, FrmMainface, CStr(varPara(0)), CStr(varPara(1)), CStr(varPara(2)), CStr(varPara(3))
                    Case 4
                        gobjCls(IntCount).ReportOpen gcnOracle, lngSys, Modul, FrmMainface, CStr(varPara(0)), CStr(varPara(1)), CStr(varPara(2)), CStr(varPara(3)), CStr(varPara(4))
                    Case 5
                        gobjCls(IntCount).ReportOpen gcnOracle, lngSys, Modul, FrmMainface, CStr(varPara(0)), CStr(varPara(1)), CStr(varPara(2)), CStr(varPara(3)), CStr(varPara(4)), CStr(varPara(5))
                    Case 6
                        gobjCls(IntCount).ReportOpen gcnOracle, lngSys, Modul, FrmMainface, CStr(varPara(0)), CStr(varPara(1)), CStr(varPara(2)), CStr(varPara(3)), CStr(varPara(4)), CStr(varPara(5)), CStr(varPara(6))
                    Case 7
                        gobjCls(IntCount).ReportOpen gcnOracle, lngSys, Modul, FrmMainface, CStr(varPara(0)), CStr(varPara(1)), CStr(varPara(2)), CStr(varPara(3)), CStr(varPara(4)), CStr(varPara(5)), CStr(varPara(6)), CStr(varPara(7))
                    Case 8
                        gobjCls(IntCount).ReportOpen gcnOracle, lngSys, Modul, FrmMainface, CStr(varPara(0)), CStr(varPara(1)), CStr(varPara(2)), CStr(varPara(3)), CStr(varPara(4)), CStr(varPara(5)), CStr(varPara(6)), CStr(varPara(7)), CStr(varPara(8))
                    Case 9
                        gobjCls(IntCount).ReportOpen gcnOracle, lngSys, Modul, FrmMainface, CStr(varPara(0)), CStr(varPara(1)), CStr(varPara(2)), CStr(varPara(3)), CStr(varPara(4)), CStr(varPara(5)), CStr(varPara(6)), CStr(varPara(7)), CStr(varPara(8)), CStr(varPara(9))
                    Case Else
                        gobjCls(IntCount).ReportOpen gcnOracle, lngSys, Modul, FrmMainface, CStr(varPara(0)), CStr(varPara(1)), CStr(varPara(2)), CStr(varPara(3)), CStr(varPara(4)), CStr(varPara(5)), CStr(varPara(6)), CStr(varPara(7)), CStr(varPara(8)), CStr(varPara(9))
                    End Select
                    
                Else
                    gobjCls(IntCount).ReportOpen gcnOracle, lngSys, Modul, FrmMainface
                End If
                
            End If
        ElseIf UCase(Components) = UCase("zl9EmrInterface") Then
            Dim strFuncs As String, strModul As String
            
            strSQL = " Select ���⡡From zlPrograms Where ���=[1] "
            Set rsCheck = zlDatabase.OpenSQLRecord(strSQL, "ϵͳģ����", Modul)
            With rsCheck
                    If .EOF Then
                        MsgBox "ϵͳ�����ݲ����������������������ϵ��", vbInformation, gstrSysName
                        FrmMainface.MousePointer = 0
                        Exit Sub
                    Else
                        strModul = !����
                    End If
            End With
            strFuncs = GetPrivFunc(lngSys, Modul)
            Call FrmMainface.mobjEmr.CodeMain(Modul, strModul, FrmMainface.hwnd, gobjRelogin.EMRUser, gobjRelogin.EMRPwd, strFuncs)
        Else
            Call CreateSynonyms(lngSys, Modul)
            
            '�û�վ������� (��ʽ�漰���ð�)
            intClients = Val(zlRegInfo("��Ȩվ��"))
            If intClients > 0 Then
                If GetCurStates > intClients Then
                    MsgBox "��ǰ�û���¼�������������Ȩ��" & intClients & ",ϵͳ���Զ��������У�", vbInformation, gstrSysName
                    Exit Sub
                End If
            End If

            
            'ΪͨѶԭ�Ӹ�ֵ
            strCommon = Format(Now, "yyyyMMddHHmm")
            strCommon = TranPasswd(strCommon) & "||" & OS.ComputerName
            '����ͨѶԭ��
            intAtom = GlobalAddAtom(strCommon)
            Call SaveSetting("ZLSOFT", "����ȫ��", "����", intAtom)
            gobjCls(IntCount).CodeMan lngSys, Modul, gcnOracle, FrmMainface, gstrDbUser
            Call GlobalDeleteAtom(intAtom)
            
            '��ҽ������ֻ��CodeMan()���ܻ�ȡϵͳ�ţ��ڶ�ȡ����ʱ����֪��ϵͳ�ţ���д��ע������ҽ��������Ĭ��Ϊ 100
            Call SaveSetting("ZLSOFT", "����ȫ��", "ϵͳ��", lngSys)
        End If
    End If
    FrmMainface.MousePointer = 0
    Exit Sub
errH:
    FrmMainface.MousePointer = 0
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Sub ReLogin()
    '����:���������¼
    mblnע�� = True
    
    Call gobjRelogin.ReLogin(FrmMainface)
End Sub

Public Function OwnerUser(ByVal strUserName As String) As Boolean
    Dim RecUser As New ADODB.Recordset
    Dim strSQL As String
    OwnerUser = True
'    With RecUser
    On Error GoTo errH
        strSQL = "Select Count(*) ������ From ZlSystems Where ������='" & strUserName & "'"
         Set RecUser = zlDatabase.OpenSQLRecord(strSQL, "������")
'        .Open "Select Count(*) ������ From ZlSystems Where ������='" & strUserName & "'", gcnOracle By zq
        
        If RecUser.EOF Then
            If Not IsNull(RecUser!������) Then
                If RecUser!������ = 0 Then OwnerUser = False
            End If
        End If
    Exit Function
errH:
    OwnerUser = False
    If ErrCenter() = 1 Then
        Resume
    End If
'    End With
End Function

Public Function CreateSynonyms(ByVal lngSys As Long, ByVal lngModul As Long)
    Dim strSQL As String
    '����ģ����������ͬ���(����Ѵ����򲻻��ٴ���)
    On Error Resume Next
    If mlngSysPre <> lngSys Then
        strSQL = "Zl_Createsynonyms(" & lngSys & ")"
        zlDatabase.ExecuteProcedure strSQL, "����ͬ���"
        mlngSysPre = lngSys
    End If
End Function

Public Sub AddHistory(ByVal strModul As String)
    Dim strϵͳ As String, str��� As String, intMax As Integer
    Dim arrϵͳ As Variant, arr��� As Variant, strValue As String
    Dim intϵͳ_Cur As Integer, int���_Cur As Integer
    Dim intϵͳ_Max As Integer, int���_Max As Integer
    '������еĳ���ʼ���ڵ�һ��λ�ã�����Ѵ�������ʷ��¼�У��������ڵ�һ��λ��
    'strModul:ϵͳ & "," & ģ��
    
    intMax = 6
    
    strValue = zlDatabase.GetPara("���ʹ��ģ��")
    If UBound(Split(strValue, "|")) >= 1 Then
        strϵͳ = Trim(Split(strValue, "|")(0))
        str��� = Trim(Split(strValue, "|")(1))
    End If
    If strϵͳ = "" Or str��� = "" Then
        strϵͳ = Split(strModul, ",")(0)
        str��� = Split(strModul, ",")(1)
        Call zlDatabase.SetPara("���ʹ��ģ��", strϵͳ & "|" & str���)
        Exit Sub
    End If
    
    arrϵͳ = Split(strϵͳ, ",")
    arr��� = Split(str���, ",")
    intϵͳ_Max = UBound(arrϵͳ)
    int���_Max = UBound(arr���)
    strϵͳ = Split(strModul, ",")(0): str��� = Split(strModul, ",")(1)
    If intϵͳ_Max > intMax Then intϵͳ_Max = intMax
    
    For intϵͳ_Cur = 0 To intϵͳ_Max
        int���_Cur = intϵͳ_Cur
        If int���_Cur > int���_Max Then Exit For
        If Not (arrϵͳ(intϵͳ_Cur) = Split(strModul, ",")(0) And arr���(int���_Cur) = Split(strModul, ",")(1)) Then
            strϵͳ = strϵͳ & "," & arrϵͳ(intϵͳ_Cur)
            str��� = str��� & "," & arr���(int���_Cur)
        End If
    Next
    Call zlDatabase.SetPara("���ʹ��ģ��", strϵͳ & "|" & str���)
End Sub

Public Sub CheckWinVersion()
    Dim lngVersion As Long
    
    mblnע�� = False
    lngVersion = GetVersion()
    If ((lngVersion And &H80000000) = 0) Then
        mlngWin32 = WINDOWSNT
    Else
        mlngWin32 = WINDOWS95
    End If
End Sub

Public Sub ShutDown(ByVal blnCloseWin As Boolean)
    If mblnע�� Then Exit Sub
    If Not blnCloseWin Then Exit Sub
    If mlngWin32 = WINDOWSNT Then
        'ExitWindowsEx lng�رռ��������Դ Or EWX_FORCEIFHUNG, 0
        Call AdjustToken
        Call ExitWindowsEx(M_lng�رռ��������Դ Or EWX_FORCE, 0)
    Else
        Call ExitWindowsEx(M_lng�رռ��������Դ Or EWX_FORCE, 0)
    End If
End Sub

Public Function AdjustToken() As Boolean
    Const TOKEN_ADJUST_PRIVILEGES = &H20
    Const TOKEN_QUERY = &H8
    Const SE_PRIVILEGE_ENABLED = &H2
    Dim hdlProcessHandle As Long
    Dim hdlTokenHandle As Long
    Dim tmpLuid As LUID
    Dim tkp As TOKEN_PRIVILEGES
    Dim tkpNewButIgnored As TOKEN_PRIVILEGES
    Dim lBufferNeeded As Long
    
    'Set the error code of the last thread to zero using the'SetLast Error function
    SetLastError 0
    
    '�õ���ǰ���̵ľ��
    hdlProcessHandle = GetCurrentProcess()
    If GetLastError <> 0 Then Exit Function
    
    '�õ���ǰ���̵�Ȩ�޾��
    OpenProcessToken hdlProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY), hdlTokenHandle
    If GetLastError <> 0 Then Exit Function
     
    '�ҵ��ر�Ȩ�޲�����LUID
    'SE_REMOTE_SHUTDOWN_NAME = "SeRemoteShutdownPrivilege
    'SE_SHUTDOWN_NAME = "SeShutdownPrivilege"
    LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid
    
    tkp.PrivilegeCount = 1    ' One privilege to set
    tkp.TheLuid = tmpLuid
    tkp.Attributes = SE_PRIVILEGE_ENABLED
    
    'Enable the shutdown privilege in the access token of this process
    AdjustTokenPrivileges hdlTokenHandle, False, tkp, Len(tkpNewButIgnored), tkpNewButIgnored, lBufferNeeded
    If GetLastError <> 0 Then Exit Function
    
    AdjustToken = True
End Function

Public Function TranPasswd(strOld As String) As String
    '------------------------------------------------
    '���ܣ� ����ת������
    '������
    '   strOld��ԭ����
    '���أ� �������ɵ�����
    '------------------------------------------------
    Dim intDo As Integer
    Dim StrPass As String, strReturn As String, strSource As String, strTarget As String
    
    StrPass = "WriteByZybZL"
    strReturn = ""
    
    For intDo = 1 To 12
        strSource = Mid(strOld, intDo, 1)
        strTarget = Mid(StrPass, intDo, 1)
        strReturn = strReturn & Chr(Asc(strSource) Xor Asc(strTarget))
    Next
    TranPasswd = strReturn
End Function

Public Function RunCommand(ByVal strCommand As String, Optional ByRef strErr As String, Optional ByVal blnCiper As Boolean, Optional ByVal lngWait As Long = INFINITE) As String
'���ܣ�ִ�������У�����ȡ���������
    Dim piProc          As PROCESS_INFORMATION '������Ϣ
    Dim stStart         As STARTUPINFO '������Ϣ
    Dim saSecAttr       As SECURITY_ATTRIBUTES '��ȫ����
    Dim lnghReadPipe    As Long '��ȡ�ܵ����
    Dim lnghWritePipe   As Long 'д��ܵ����
    Dim lngBytesRead    As Long '�������ݵ��ֽ���
    Dim strBuffer       As String * 256 '��ȡ�ܵ����ַ���buffer
    Dim lngRet          As Long 'API��������ֵ
    Dim lngRetPro       As Long
    Dim strlpOutputs    As String '���������ս��
    
    DoEvents
    On Error Resume Next
    '���ð�ȫ����
    With saSecAttr
        .nLength = LenB(saSecAttr)
        .bInheritHandle = True
        .lpSecurityDescriptor = 0
    End With
    
    '�����ܵ�
    lngRet = CreatePipe(lnghReadPipe, lnghWritePipe, saSecAttr, 0)
    If lngRet = 0 Then
        strErr = "�޷������ܵ���" & GetLastDllErr()
        Exit Function
    End If
    '���ý�������ǰ����Ϣ
    With stStart
        .Cb = LenB(stStart)
        .dwFlags = STARTF_USESHOWWINDOW Or STARTF_USESTDHANDLES
        .wShowWindow = SW_HIDE
        .hStdOutput = lnghWritePipe '��������ܵ�
        .hStdError = lnghWritePipe '���ô���ܵ�
    End With
    '��������
    'Command = "c:\windows\system32\ipconfig.exe /all" 'DOS������ipconfig.exeΪ��
    lngRetPro = CreateProcess(vbNullString, strCommand & vbNullChar, saSecAttr, saSecAttr, 1&, NORMAL_PRIORITY_CLASS, ByVal 0&, vbNullString, stStart, piProc)
    If lngRetPro = 0 Then
        strErr = "�޷��������̡�" & GetLastDllErr()
        lngRet = CloseHandle(lnghWritePipe)
        lngRet = CloseHandle(lnghReadPipe)
        Exit Function
    Else
        '��Ϊ����д�����ݣ������ȹر�д��ܵ��������������رմ˹ܵ��������޷���ȡ����
        lngRet = CloseHandle(lnghWritePipe)
        WaitForSingleObject piProc.hProcess, lngWait
        Do
            lngRet = ReadFile(lnghReadPipe, strBuffer, 256, lngBytesRead, ByVal 0)
            If lngRet <> 0 Then
                strlpOutputs = strlpOutputs & Left(strBuffer, lngBytesRead)
            Else
                strlpOutputs = strlpOutputs & Left(strBuffer, lngBytesRead)
            End If
            DoEvents
        Loop While (lngRet <> 0) '��ret=0ʱ˵��ReadFileִ��ʧ�ܣ��Ѿ�û�����ݿɶ���
        '��ȡ������ɣ��رո����
        lngRet = CloseHandle(lngRetPro)
        lngRet = CloseHandle(piProc.hProcess)
        lngRet = CloseHandle(piProc.hThread)
        lngRet = CloseHandle(lnghReadPipe)
    End If
    RunCommand = Replace(strlpOutputs, vbNullChar, "")
End Function

Public Function GetLastDllErr(Optional ByVal lngErr As Long) As String
    Dim strReturn As String
    If lngErr = 0 Then
        lngErr = GetLastError
    End If
    If lngErr = ERROR_EXTENDED_ERROR Then
        GetLastDllErr = GetWNetErr(lngErr)
    Else
        strReturn = String$(256, 32)
        FormatMessage FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, 0&, lngErr, 0&, strReturn, Len(strReturn), ByVal 0
        strReturn = Trim(strReturn)
        GetLastDllErr = Replace(Replace(strReturn, Chr(10), ""), Chr(13), "")
    End If
End Function

Private Function GetWNetErr(ByVal lngErr As Long) As String
    Dim strErr As String * 256
    Dim strName As String * 256
    Dim lngRet As Long
    lngRet = WNetGetLastError(lngErr, strErr, Len(strErr), strName, Len(strName))
    GetWNetErr = Replace(Replace("[" & TruncZero(strName) & "]" & TruncZero(strErr), Chr(10), ""), Chr(13), "")
End Function

Public Function TruncZero(ByVal strInput As String) As String
'���ܣ�ȥ���ַ�����\0�Ժ���ַ�
    Dim lngPos As Long
    
    lngPos = InStr(strInput, Chr(0))
    If lngPos > 0 Then
        TruncZero = Mid(strInput, 1, lngPos - 1)
    Else
        TruncZero = strInput
    End If
End Function
