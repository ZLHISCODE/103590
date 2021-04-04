Attribute VB_Name = "mdlHisCrust"
Option Explicit

'*******************************************************************************************************
'˵������ģ���ZLLogin��ģ��Ӧ�ñ���һ��
'*******************************************************************************************************
'���������������API
'----------------------------------------------------------------------------------------------------
'Window�汾����
'win2000 ���°汾
Private Type OSVERSIONINFO 'for GetVersionEx API call
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Private Type OSVERSIONINFOEX
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128      '  Maintenance string for PSS usage
    wServicePackMajor As Integer 'win2000 only
    wServicePackMinor As Integer 'win2000 only
    wSuiteMask As Integer 'win2000 only
    wProductType As Byte 'win2000 only
    wReserved As Byte
End Type
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFOEX) As Long
'wSuiteMask
Private Const VER_SUITE_BACKOFFICE = &H4                'Microsoft BackOffice components are installed.
Private Const VER_SUITE_BLADE = &H400                   'Windows Server 2003, Web Edition is installed.
Private Const VER_SUITE_COMPUTE_SERVER = &H4000         'Windows Server 2003, Compute Cluster Edition is installed.
Private Const VER_SUITE_DATACENTER = &H80               'Windows Server 2008 Datacenter, Windows Server 2003, Datacenter Edition, or Windows 2000 Datacenter Server is installed.
Private Const VER_SUITE_ENTERPRISE = &H2                'Windows Server 2008 Enterprise, Windows Server 2003, Enterprise Edition, or Windows 2000 Advanced Server is installed. Refer to the Remarks section for more information about this bit flag.
Private Const VER_SUITE_EMBEDDEDNT = &H40               'Windows XP Embedded is installed.
Private Const VER_SUITE_PERSONAL = &H200                'Windows Vista Home Premium, Windows Vista Home Basic, or Windows XP Home Edition is installed.
Private Const VER_SUITE_SINGLEUSERTS = &H100            'Remote Desktop is supported, but only one interactive session is supported. This value is set unless the system is running in application server mode.
Private Const VER_SUITE_SMALLBUSINESS = &H1             'Microsoft Small Business Server was once installed on the system, but may have been upgraded to another version of Windows. Refer to the Remarks section for more information about this bit flag.
Private Const VER_SUITE_SMALLBUSINESS_RESTRICTED = &H20 'Microsoft Small Business Server is installed with the restrictive client license in force. Refer to the Remarks section for more information about this bit flag.
Private Const VER_SUITE_STORAGE_SERVER = &H2000         'Windows Storage Server 2003 R2 or Windows Storage Server 2003is installed.
Private Const VER_SUITE_TERMINAL = &H10                 'Terminal Services is installed. This value is always set.
                                                        'If VER_SUITE_TERMINAL is set but VER_SUITE_SINGLEUSERTS is not set, the system is running in application server mode.
Private Const VER_SUITE_WH_SERVER = &H8000              'Windows Home Server is installed.
'wProductType
Private Const VER_NT_DOMAIN_CONTROLLER = &H2            'The system is a domain controller and the operating system is Windows Server 2012 , Windows Server 2008 R2, Windows Server 2008, Windows Server 2003, or Windows 2000 Server.
Private Const VER_NT_SERVER = &H3                       'The operating system is Windows Server 2012, Windows Server 2008 R2, Windows Server 2008, Windows Server 2003, or Windows 2000 Server.
                                                        'Note that a server that is also a domain controller is reported as VER_NT_DOMAIN_CONTROLLER, not VER_NT_SERVER.
Private Const VER_NT_WORKSTATION = &H1                  'The operating system is Windows 8, Windows 7, Windows Vista, Windows XP Professional, Windows XP Home Edition,
'dwPlatformId
Private Const VER_PLATFORM_WIN32s = 0
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32_NT = 2
'GetSystemMetrics
Private Const SM_TABLETPC = 86                          'Windows XP Tablet PC Edition
Private Const SM_MEDIACENTER = 87                       'Windows XP Media Center Edition
Private Const SM_STARTER = 88                           'Windows XP Starter Edition
Private Const SM_SERVERR2 = 89                          'Windows Server 2003 R2
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Type SYSTEM_INFO
'    dwOemID As Long
    wProcessorArchitecture As Integer
    wReserved As Integer
    dwPageSize As Long
    lpMinimumApplicationAddress As Long
    lpMaximumApplicationAddress As Long
    dwActiveProcessorMask As Long
    dwNumberOrfProcessors As Long
    dwProcessorType As Long
    dwAllocationGranularity As Long
    dwReserved As Long
End Type
Private Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As SYSTEM_INFO)
'wProcessorArchitecture
Private Const PROCESSOR_ARCHITECTURE_AMD64 = 9          'x64 (AMD Or Intel)
Private Const PROCESSOR_ARCHITECTURE_ARM = 5            'ARM
Private Const PROCESSOR_ARCHITECTURE_IA64 = 6           'Intel Itanium - based
Private Const PROCESSOR_ARCHITECTURE_INTEL = 0          'x86
Private Const PROCESSOR_ARCHITECTURE_UNKNOWN = &HFFFF   'Unknown architecture.
Private Const PROCESSOR_INTEL_386 = 386
Private Const PROCESSOR_INTEL_486 = 486
Private Const PROCESSOR_INTEL_PENTIUM = 586
Private Const PROCESSOR_INTEL_IA64 = 2200
Private Const PROCESSOR_AMD_X8664 = 8664
Private Const PROCESSOR_MIPS_R4000 = 4000      ' incl R4101 & R3910 for Windows CE
Private Const PROCESSOR_ALPHA_21064 = 21064
Private Const PROCESSOR_PPC_601 = 601
Private Const PROCESSOR_PPC_603 = 603
Private Const PROCESSOR_PPC_604 = 604
Private Const PROCESSOR_PPC_620 = 620
Private Const PROCESSOR_HITACHI_SH3 = 10003    ' Windows CE
'��ȡ�ڴ�
Private Type MEMORYSTATUS  'win2000�����°汾
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long

End Type
Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)

Private Type MEMORYSTATUSEX
        dwLength       As Long
        dwMemoryLoad   As Long
        ullTotalPhys   As Currency
        ullAvailPhys   As Currency
        ullTotalPageFile   As Currency
        ullAvailPageFile   As Currency
        ullTotalVirtual    As Currency
        ullAvailVirtual    As Currency
        ullAvailExtendedVirtual   As Currency
End Type
Private Declare Function GlobalMemoryStatusEx Lib "kernel32.dll" (ByRef lpBuffer As MEMORYSTATUSEX) As Long
'ȡӲ�̴�С
Private Const DRIVE_FIXED = 3
Private Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTotalNumberOfClusters As Long) As Long
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Const STRSPLIT As String = "���"

'API������Ϣ��ȡ
Private Declare Function GetLastError Lib "kernel32" () As Long
Private Declare Function FormatMessage Lib "kernel32.dll" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200
Private Declare Function WNetGetLastError Lib "mpr.dll" Alias "WNetGetLastErrorA" (lpError As Long, ByVal lpErrorBuf As String, ByVal nErrorBufSize As Long, ByVal lpNameBuf As String, ByVal nNameBufSize As Long) As Long
Private Const ERROR_EXTENDED_ERROR          As Long = 1208
'�ļ�������Ϣ�ж�
Private Declare Function GetFileVersionInfoSize Lib "version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Function GetFileVersionInfo Lib "version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwHandle As Long, ByVal dwLen As Long, lpData As Any) As Long
Private Declare Function VerQueryValue Lib "version.dll" Alias "VerQueryValueA" (ByVal pBlock As Long, ByVal lpSubBlock As String, lplpBuffer As Long, puLen As Long) As Long
'Public Const FVN_Comments           As String = "Comments"          'ע��
'Public Const FVN_InternalName       As String = "InternalName"      '�ڲ�����
'Public Const FVN_ProductName        As String = "ProductName"       '��Ʒ��
'Public Const FVN_CompanyName        As String = "CompanyName"       '��˾��
'Public Const FVN_ProductVersion     As String = "ProductVersion"    '��Ʒ�汾
'Public Const FVN_FileDescription    As String = "FileDescription"   '�ļ�����
'Public Const FVN_OriginalFilename   As String = "OriginalFilename"  'ԭʼ�ļ���
'Public Const FVN_FileVersion        As String = "FileVersion"       '�ļ��汾
'Public Const FVN_SpecialBuild       As String = "SpecialBuild"      '��������
'Public Const FVN_PrivateBuild       As String = "PrivateBuild"      '˽�б����
'Public Const FVN_LegalCopyright     As String = "LegalCopyright"    '�Ϸ���Ȩ
'Public Const FVN_LegalTrademarks    As String = "LegalTrademarks"   '�Ϸ��̱�
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
'hModule��һ��ģ��ľ����������һ��DLLģ�飬������һ��Ӧ�ó����ʵ�����������ò���ΪNULL���ú������ظ�Ӧ�ó���ȫ·��?
Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal LpApplicationName As String, ByVal LpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Public gstrExeFile      As String '���õ�¼������EXE·��
Public gstrSetupPath    As String 'APPSOFT·��
Public glnghInstance    As Long
Public gblnTimer        As Boolean  '�Ƿ�ʱ�������Ŀͻ��˸��¼��

Public Function CheckAllowByTerminal() As Boolean
'����:����Ƿ�����ʹ�ñ�����վ,�Լ����е�ǰ����վ��Ϣ�ĵǼ�
'     �ж��Ƿ�����ù���վʹ�ó���
'     �����Ҫ�滻���ز�������ִ���滻�����������Ҫ�������������ǳ��򣬲��ر��˳�
'����:�ɹ�,����true,���򷵻�False
'���棺���ڻ�û�г�ʼ���������������Ӷ��󣬸ú����в���ʹ�ù��������е����ݿ���ʷ���

    Dim rsTmp As ADODB.Recordset, strSQL As String, strRowID As String '�ͻ��˵�ROWID
    Dim strComuterInfo As String, arrComputer As Variant, strComputerName As String, strIpAddress As String
    Dim strTmp As String, arrTmp As Variant, i As Integer
    Dim bln���վ�� As Boolean, lng��վ�� As Long, bln��վ�� As Boolean, bln��վ�� As Boolean
    Dim strվ��       As String, strվ���� As String, str���� As String, strȱʡ����
    Dim blnAllow As Boolean, blnUpdate As Boolean
    Dim int��������� As Integer, int������ƵԴ As Integer, int������ As Integer, int������־ As Integer
    
'    Call SQLTest(App.EXEName, "mdlHisCrust", "�°���Ӳ����Զ��������")
    Call UpdateEmrInterface '�°���Ӳ����Զ�����
'    Call SQLTest

    strIpAddress = IP '��oracle���ӵ�IP��ַΪ��
    strComputerName = OS.ComputerName
    '����Ƿ�����������
    If CheckRepeatLogin(strIpAddress) = True Then
        CheckAllowByTerminal = False
        Exit Function
    End If
    '�ж��Ƿ�����ʹ��
    strComuterInfo = AnalyseConfigure
    arrComputer = Split(strComuterInfo, STRSPLIT)
    '1.��վ�������
    If Err.Number <> 0 Then Err.Clear
    On Error Resume Next
    strSQL = "Select Rowid as ID,վ��,����,Nvl(��ֹʹ��,0) as ����,Nvl(������־,0) as ����,Nvl(�ռ���־,0) as �ռ�,������,������ƵԴ From zlClients Where ����վ=[1]"
    Set rsTmp = OpenSQLRecord(strSQL, "��鹤��վ-��վ��Ϊ��", strComputerName)
    '��������δ��Ȩ��ԭ�򣬵��²�ѯ������ʱ������ʾ��ֹ��¼
    If rsTmp Is Nothing Then
        MsgBox Err.Description & vbNewLine & "������������ϵͳ��������ϵϵͳ����Ա���½��н�ɫ��Ȩ��", vbInformation, gstrSysName
        Exit Function
    End If
    '2.δ���ִ�վ��,����IP��ʽ���ң���ֻ��һ��ʱ�Ÿ��¼�����
    If rsTmp.EOF Then
        strSQL = "Select Rowid as ID,վ��,����, Nvl(��ֹʹ��,0) as ����,Nvl(������־,0) as ����,Nvl(�ռ���־,0) as �ռ�,������,������ƵԴ From zlClients Where IP=[1]"
        Set rsTmp = OpenSQLRecord(strSQL, "��鹤��վ-��վ��Ϊ��", strIpAddress)
        If rsTmp.RecordCount > 1 Then
            '������������,���CPU,�ڴ�,Ӳ��Ϊ��������.
            strSQL = "" & _
                "   Select Rowid as ID,վ��,����,Nvl(��ֹʹ��,0) as ����,Nvl(������־,0) as ����,Nvl(�ռ���־,0) as �ռ�,������,������ƵԴ " & _
                "   From zlClients Where IP=[1] and CPU=[2] and  �ڴ�=[3] and Ӳ��=[4]"
            Set rsTmp = OpenSQLRecord(strSQL, "��鹤��վ-��վ��Ϊ��", strIpAddress, CStr(arrComputer(2)), CStr(arrComputer(3)), CStr(arrComputer(4)))
        End If
    End If
    bln���վ�� = True
    '��������ڶ��,����ܴ���IP��ͻ�����,��˲����ж���Ҫ������ص�վ��.ֻ�ܵ����µ�վ���ϴ�
    If rsTmp.RecordCount > 1 Or rsTmp.EOF Then
        strRowID = ""
    Else '��ʾ������ص���Ϣ
        strRowID = NVL(rsTmp!id)
        int������ƵԴ = Val(NVL(rsTmp!������ƵԴ))
        '�������½,�������û�ѡ��,ֱ�Ӷ�ȡ
        If gstrCommand <> "" Then
            '�·���
            If InStr(gstrCommand, "ZLHISCRUSTCALL=1") > 0 And InStr(gstrCommand, "USER=") > 0 And InStr(gstrCommand, "PASS=") > 0 Then
                bln���վ�� = False
                strվ���� = NVL(rsTmp!վ��)
                gclsLogin.DeptName = NVL(rsTmp!����)
            '�ϵ��жϷ���
            ElseIf InStrRev(gstrCommand, "/", -1) > 0 And InStrRev(gstrCommand, ",", -1) = 0 Then
                bln���վ�� = False
                strվ���� = NVL(rsTmp!վ��)
                gclsLogin.DeptName = NVL(rsTmp!����)
            End If
        End If
        blnAllow = Val(rsTmp!���� & "") = 0
        int������ = Val(rsTmp!������ & "")  '0-��ʾ������
        blnUpdate = Val(rsTmp!���� & "") = 1
        If Not blnUpdate Then blnUpdate = Val(rsTmp!�ռ� & "") = 1
    End If

    If bln���վ�� Then
        strSQL = "Select b.����, a.վ��, a.ȱʡ" & vbNewLine & _
                "From (Select Distinct c.վ��, b.ȱʡ" & vbNewLine & _
                "       From �ϻ���Ա�� a, ������Ա b, ���ű� c" & vbNewLine & _
                "       Where a.��Աid = b.��Աid And b.����id = c.Id And a.�û��� = [1]) a, Zlnodelist b" & vbNewLine & _
                "Where a.վ�� = b.���(+)" & vbNewLine & _
                "Order By վ��"

        Set rsTmp = OpenSQLRecord(strSQL, "��鲢ȷ������Ժ��", UCase(gclsLogin.DBUser))
        If rsTmp Is Nothing Then
            MsgBox Err.Description & vbNewLine & "������������ϵͳ��������ϵϵͳ����Ա���½��н�ɫ��Ȩ��", vbInformation, gstrSysName
            Exit Function
        End If
        Do While Not rsTmp.EOF
            If NVL(rsTmp!վ��, "") <> "" Then
                strվ�� = strվ�� & "," & NVL(rsTmp!վ��, "")
                str���� = str���� & "," & NVL(rsTmp!����)
                lng��վ�� = lng��վ�� + 1
            Else
                bln��վ�� = True
            End If
            If NVL(rsTmp!ȱʡ, "0") = 1 Then
                strȱʡ���� = NVL(rsTmp!����)
            End If
            rsTmp.MoveNext
        Loop
        '�����ǰ��¼��Ա�������Ŷ�û������վ�㣬���������ڲ��Ҹ�Ժ�Ƿ�������վ�����!
        If strվ�� = "" Or (bln��վ�� And lng��վ�� <> 1) Then
            '������װ�°�LISʱҲ��Ҫ��������ȡվ��
            strTmp = GetLISStation()
            If strTmp <> "" Then
                arrTmp = Split(strTmp, ";")
                strվ�� = arrTmp(0)
                str���� = arrTmp(1)
            Else
                strվ�� = "": str���� = ""
                strSQL = "select distinct (A.վ��),B.���� from ���ű� A,zlNodeList B where A.վ��=B.��� And A.վ�� is not null order by A.վ��"
                Set rsTmp = OpenSQLRecord(strSQL, "����Ƿ�����վ�����")
                If Not rsTmp Is Nothing Then
                    Do While Not rsTmp.EOF
                        If NVL(rsTmp!վ��, "") <> "" Then
                            strվ�� = strվ�� & "," & NVL(rsTmp!վ��, "")
                            str���� = str���� & "," & NVL(rsTmp!����)
                        End If
                        rsTmp.MoveNext
                    Loop
                End If
            End If
        End If
        If strվ�� <> "" Then
            strվ�� = Mid(strվ��, 2)
            str���� = Mid(str����, 2)
            arrTmp = Split(strվ��, ",")
            For i = LBound(arrTmp) To UBound(arrTmp)
                If i = LBound(arrTmp) Then
                    strվ���� = arrTmp(i)
                Else
                    If strվ���� <> arrTmp(i) Then
                        bln��վ�� = True
                        Exit For
                    End If
                End If
            Next
            If bln��վ�� Then '��ʾ�û�ѡ��ǰ�����λ�����ڵĲ��š�
                strվ���� = GetSetting("ZLSOFT", "˽��ģ��\" & gclsLogin.DBUser & "\" & App.ProductName & "\" & App.EXEName, "��ǰվ��ѡ��", "")
                Call frmSelClient.ShowEdit(strվ��, str����, strվ����, strȱʡ����)
                strվ���� = IIf(frmSelClient.gstrվ�� = "��", "", frmSelClient.gstrվ��)
                gclsLogin.DeptName = frmSelClient.gstrCurվ��
                Call SaveSetting("ZLSOFT", "˽��ģ��\" & gclsLogin.DBUser & "\" & App.ProductName & "\" & App.EXEName, "��ǰվ��ѡ��", strվ����)
            End If
        End If
    End If
    gclsLogin.NodeNo = IIf(strվ���� <> "", strվ����, "-")
    If gclsLogin.DeptName = "" Then gclsLogin.DeptName = strȱʡ����
    If strRowID = "" Then '�����Ĺ���վ����û�иù���վ�����ݣ��ϴ���IP����������CPU���ڴ桢Ӳ�̡�����ϵͳ��
        int��������� = GetDefaultFileServer
        If int��������� = -1 Then '��ȡĬ�Ϸ�����ʧ�ܣ����������ָ���������ŵĳ�ʼֵ
            int��������� = 0
            int������־ = 0
        Else
            int������־ = 1
        End If
        strSQL = "Zl_Zlclients_Set(0,Null,'" & strComputerName & "','" & strIpAddress & "','" & arrComputer(2) & "','" & arrComputer(3) & _
                    "','" & arrComputer(4) & "','" & arrComputer(5) & "','" & gclsLogin.DeptName & "',Null,Null," & int��������� & "," & int������־ & _
                    ",0,'" & strվ���� & "',0,Null,Null," & int������ƵԴ & ")"
        ExecuteProcedure strSQL, "��������վ"
        '�����ͻ��˲���������ֱ���˳�
        If int������־ = 0 Then
            CheckAllowByTerminal = True
            Exit Function
        End If
        blnUpdate = True
    Else
        strSQL = "Zl_Zlclients_Set(1,'" & strRowID & "','" & strComputerName & "','" & strIpAddress & "','" & arrComputer(2) & "','" & arrComputer(3) & _
                    "','" & arrComputer(4) & "','" & arrComputer(5) & "','" & gclsLogin.DeptName & "',Null,Null,Null,Null," & int������ & ",'" & strվ���� & "',0,Null,Null," & int������ƵԴ & ")"
        '��Ҫ������ص�վ����Ϣ
        ExecuteProcedure strSQL, "���¹���վ"
        If Not blnAllow Then
            MsgBox "�ù���վ�ѱ�����Ա���ã�", vbInformation, gstrSysName
            Exit Function
        End If
        '�������������
        If int������ > 0 Then
            strSQL = "Select SID From gv$Session Where Upper(PROGRAM) Like 'ZL%.EXE' And Status<>'KILLED' And MACHINE=(Select Max(MACHINE) From v$Session Where AUDSID=UserENV('SessionID'))"
            Set rsTmp = OpenSQLRecord(strSQL, "�����������")
            If rsTmp.RecordCount > int������ Then
                MsgBox "��ǰ����վ���ֻ���� " & int������ & " ����¼���ӣ���ǰ�Ѿ��� " & rsTmp.RecordCount - 1 & " �����ӡ�", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End If
    On Error GoTo Errhand
AutoUpGrude:      'ִ����������
    If blnUpdate Then
        blnAllow = UpdateZLHIS(strComputerName)
    End If
    CheckAllowByTerminal = blnAllow
    Exit Function
Errhand:
    MsgBox "���������ִ���" & Err.Description & "��������ϵϵͳ����Ա���н����", vbInformation, gstrSysName
End Function

Public Function StartHisCrust(ByVal str�������� As String, ByVal strJobName As String, Optional ByVal lngWait As Long, Optional ByVal StrPass As String) As Boolean
'���ܣ������Զ��������
'������str��������=����ֱ�Ӵ�����ļ�·����Ҳ���Դ��ļ���
'      strJobName=�������ƣ����ߵ��ó�����
'      lngWait=��ʽ����ʱ���ȴ���N���Ӻ����ʽ����
'���أ��Ƿ�ɹ�
    Dim strUP As String
    Dim strUPFile  As String, strFileName As String
    Dim strConnString As String, lngErr As Long
    Dim objFile As New FileSystemObject
    Dim strCheck As String, strCommand As String
    Dim strExeName As String
    
    
    On Error Resume Next
    If gstrExeFile <> "" Then
        strExeName = objFile.GetFileName(gstrExeFile)
    Else
        strExeName = strJobName
    End If
    
    If objFile.GetDriveName(str��������) = "" Then
        strUPFile = gstrSetupPath & "\" & str��������
    Else
        strUPFile = str��������
        strFileName = objFile.GetFileName(str��������)
    End If
    If Not objFile.FileExists(strUPFile) Then
        MsgBox "û���ҵ��ͻ����Զ���������" & strFileName & "������ϵͳ����Ա��ϵ��", vbExclamation, gstrSysName
        Exit Function
    End If
    If OS.IsDesinMode Then
        '��װ�����У��Լ�����������У��λ
        strCommand = "Provider=MSDataShape.1;Extended Properties=""Driver={Microsoft ODBC for Oracle};Server=" & gclsLogin.ServerName & _
                                   """;Persist Security Info=True;User ID=" & gclsLogin.InputUser & ";Password=HIS;Data Provider=MSDASQL"
    Else
        '��װ�����У��Լ�����������У��λ
        strCommand = "Provider=MSDataShape.1;Extended Properties=""Driver={Microsoft ODBC for Oracle};Server=" & gclsLogin.ServerName & _
                                   """;Persist Security Info=True;User ID=" & gclsLogin.InputUser & ";Password=" & StrPass & ";Data Provider=MSDASQL"
    End If
    strCheck = "CMDCHECK:1" & "," & Len(strCommand)
    strCommand = strCommand & "||0"
    strCheck = strCheck & "," & Len(strCommand)
    strCommand = strCommand & "||" & strExeName
    strCheck = strCheck & "," & Len(strCommand)
    strCommand = strCommand & "||" & CStr(gstrCommand)
    strCheck = strCheck & "," & Len(strCommand)
    strCommand = strCommand & "||" & "USER=" & gclsLogin.InputUser & " PASS=" & gclsLogin.InputPwd
    strCheck = strCheck & "," & Len(strCommand)
    If lngWait <> 0 Then
        strCommand = strCommand & "||W:" & lngWait
        strCheck = strCheck & "," & Len(strCommand)
    End If
    strCommand = strCommand & "||" & strCheck
    lngErr = Shell(strUPFile & " " & strCommand, vbNormalFocus)
    StartHisCrust = True
    If lngErr = 0 Then
        MsgBox "�޷����������������̣���ʹ�ò���ϵͳ����Ա�����������", vbInformation, gstrSysName
    End If
End Function

Private Function AnalyseConfigure() As String
    '��д��:���� 2003-03-09
    '����:���������������ã�IP����������CPU���ڴ桢Ӳ�̡�����ϵͳ��
    Dim strCPU As String           'CPU
    Dim strMemory As String        '�ڴ�
    Dim strOS As String            '����ϵͳ
    Dim strComputerName As String  '�������
    Dim strHD As String            'Ӳ��
    Dim strIp As String            'IP��ַ
    Dim verinfo As OSVERSIONINFOEX
    Dim sysinfo As SYSTEM_INFO
    Dim memsts As MEMORYSTATUS
    Dim memstsex As MEMORYSTATUSEX
    Dim lngmemory As Long
    Dim curMemory As Currency
    
    strIp = OS.IP
    '��ȡ�������
    strComputerName = OS.ComputerName
    '��ȡӲ����Ϣ
    strHD = AnalyseHardDisk
    ' ��ò���ϵͳ��Ϣ
    strOS = GetVersionInfo
    ' ���CPU����
    GetSystemInfo sysinfo
    Select Case sysinfo.dwProcessorType
    Case PROCESSOR_INTEL_386
        strCPU = "Intel 386"
    Case PROCESSOR_INTEL_486
        strCPU = "Intel 486"
    Case PROCESSOR_INTEL_PENTIUM
        strCPU = "Intel Pentium"
    Case PROCESSOR_MIPS_R4000
        strCPU = "MIPS R4000"
    Case PROCESSOR_ALPHA_21064
        strCPU = "DEC Alpha 21064"
    Case Else
        strCPU = "(unknown)"
    End Select
    ' ���ʣ���ڴ�
    '���ж�ϵͳ�Ƿ�Ϊwin2000������
    '�����Windows2000�����°汾������GlobalMemoryStatusȡ
    verinfo.dwOSVersionInfoSize = Len(verinfo) 'should be 148/156
    If GetVersionEx(verinfo) = 0 Then 'try win2000 version
        GlobalMemoryStatus memsts
        lngmemory = memsts.dwTotalPhys
        strMemory = Format$(lngmemory& \ 1024 \ 1024, "###,###,###") + "M"
    Else
        memstsex.dwLength = Len(memstsex)
        GlobalMemoryStatusEx memstsex
        curMemory = memstsex.ullTotalPhys
        strMemory = CStr(Int(curMemory * 10000 / 1024 ^ 2)) & "M"
    End If
    AnalyseConfigure = strIp & STRSPLIT & strComputerName & STRSPLIT & strCPU & _
                       STRSPLIT & strMemory & STRSPLIT & strHD & STRSPLIT & strOS
End Function

Private Function AnalyseHardDisk() As String
    '��д��:���� 2003-03-09
    '����:��ȡӲ��������
    Dim lngSec As Long, lngByte As Long, lngFree As Long, lngClus As Long
    Dim strDrive As String, dblSum As Double
    
    strDrive = "C"
    Do Until strDrive > "Z"
        If GetDriveType(strDrive & ":\") = DRIVE_FIXED Then
            If GetDiskFreeSpace(strDrive & ":\", lngSec, lngByte, lngFree, lngClus) <> 0 Then
                dblSum = dblSum + lngSec * lngByte * CDbl(lngClus)
            End If
        End If
        
        strDrive = Chr(Asc(strDrive) + 1)
    Loop
    AnalyseHardDisk = Format(dblSum / 1024 / 1024 / 1024, "0.00") & "G"
End Function

Private Function GetVersionInfo() As String
    Dim myOS As OSVERSIONINFOEX
    Dim bExInfo As Boolean
    Dim strOS As String
    Dim sysinfo As SYSTEM_INFO
    'OSVERSIONINFO
    'Operating system    Version number  dwMajorVersion  dwMinorVersion  Other
    'Windows 10                 10.0*       10                  0   OSVERSIONINFOEX.wProductType == VER_NT_WORKSTATION
    'Windows Server 2016        10.0*       10                  0   OSVERSIONINFOEX.wProductType != VER_NT_WORKSTATION
    'Windows 8.1                6.3*        6                   3   OSVERSIONINFOEX.wProductType == VER_NT_WORKSTATION
    'Windows Server 2012 R2     6.3*        6                   3   OSVERSIONINFOEX.wProductType != VER_NT_WORKSTATION
    'Windows 8                  6.2         6                   2   OSVERSIONINFOEX.wProductType == VER_NT_WORKSTATION
    'Windows Server 2012        6.2         6                   2   OSVERSIONINFOEX.wProductType != VER_NT_WORKSTATION
    'Windows 7                  6.1         6                   1   OSVERSIONINFOEX.wProductType == VER_NT_WORKSTATION
    'Windows Server 2008 R2     6.1         6                   1   OSVERSIONINFOEX.wProductType != VER_NT_WORKSTATION
    'Windows Server 2008        6.0         6                   0   OSVERSIONINFOEX.wProductType != VER_NT_WORKSTATION
    'Windows Vista              6.0         6                   0   OSVERSIONINFOEX.wProductType == VER_NT_WORKSTATION
    'Windows Server 2003 R2     5.2         5                   2   GetSystemMetrics(SM_SERVERR2) != 0
    'Windows Server 2003        5.2         5                   2   GetSystemMetrics(SM_SERVERR2) == 0
    'Windows XP                 5.1         5                   1   Not applicable
    'Windows 2000               5.0         5                   0   Not applicable
    'OSVERSIONINFOEX
    'Operating system    Version number  dwMajorVersion  dwMinorVersion  Other
    'Windows 10                 10.0*       10                  0   OSVERSIONINFOEX.wProductType == VER_NT_WORKSTATION
    'Windows Server 2016        10.0*       10                  0   OSVERSIONINFOEX.wProductType != VER_NT_WORKSTATION
    'Windows 8.1                6.3*        6                   3   OSVERSIONINFOEX.wProductType == VER_NT_WORKSTATION
    'Windows Server 2012 R2     6.3*        6                   3   OSVERSIONINFOEX.wProductType != VER_NT_WORKSTATION
    'Windows 8                  6.2         6                   2   OSVERSIONINFOEX.wProductType == VER_NT_WORKSTATION
    'Windows Server 2012        6.2         6                   2   OSVERSIONINFOEX.wProductType != VER_NT_WORKSTATION
    'Windows 7                  6.1         6                   1   OSVERSIONINFOEX.wProductType == VER_NT_WORKSTATION
    'Windows Server 2008 R2     6.1         6                   1   OSVERSIONINFOEX.wProductType != VER_NT_WORKSTATION
    'Windows Server 2008        6.0         6                   0   OSVERSIONINFOEX.wProductType != VER_NT_WORKSTATION
    'Windows Vista              6.0         6                   0   OSVERSIONINFOEX.wProductType == VER_NT_WORKSTATION
    'Windows Server 2003 R2     5.2         5                   2   GetSystemMetrics(SM_SERVERR2) != 0
    'Windows Home Server        5.2         5                   2   OSVERSIONINFOEX.wSuiteMask & VER_SUITE_WH_SERVER
    'Windows Server 2003        5.2         5                   2   GetSystemMetrics(SM_SERVERR2) == 0
    'Windows XP Professional x64 Edition 5.2    5               2   (OSVERSIONINFOEX.wProductType == VER_NT_WORKSTATION) && (SYSTEM_INFO.wProcessorArchitecture==PROCESSOR_ARCHITECTURE_AMD64)
    'Windows XP                 5.1         5                   1   Not applicable
    'Windows 2000               5.0         5                   0   Not applicable
    '�����Windows2000�����°汾��������API��ȡһ��
    myOS.dwOSVersionInfoSize = Len(myOS) 'should be 148/156
    If GetVersionEx(myOS) = 0 Then 'try win2000 version
        myOS.dwOSVersionInfoSize = 148 'if fails,ignore reserved data
        If GetVersionEx(myOS) = 0 Then
            GetVersionInfo = "Windows (Unknown)"
            Exit Function
        End If
    Else
        bExInfo = True
    End If
    ' ���CPU����
    GetSystemInfo sysinfo
    With myOS
        Select Case .dwMajorVersion
            Case 3
                strOS = "Windows NT 3.1"
            Case 4
                Select Case .dwMinorVersion
                    Case 0
                        If .dwPlatformId = VER_PLATFORM_WIN32_NT Then
                            strOS = "Windows NT 4.0" '1996��7�·���
                        Else
                            strOS = "Windows 95"
                        End If
                    Case 10
                        strOS = "Windows 98"
                    Case 90
                        strOS = "Windows Me"
                End Select
            Case 5
                Select Case .dwMinorVersion
                    Case 0
                        strOS = "Windows 2000" '1999��12�·���
                        If .wProductType = VER_NT_WORKSTATION Then
                            strOS = strOS & " " & "Professional"
                        Else
                            If bExInfo Then
                                If .wSuiteMask = VER_SUITE_ENTERPRISE Then
                                    strOS = strOS & " " & "Advanced Server"
                                ElseIf .wSuiteMask = VER_SUITE_DATACENTER Then
                                    strOS = strOS & " " & "Datacenter Server"
                                Else
                                    strOS = strOS & " " & "Server"
                                End If
                            End If
                        End If
                    Case 1
                        strOS = "Windows XP" '2001��8�·���
                        If .wSuiteMask = VER_SUITE_EMBEDDEDNT Then
                            strOS = strOS & " " & "Embedded"
                        ElseIf .wSuiteMask = VER_SUITE_PERSONAL Then
                            strOS = strOS & " " & "Home Edition"
                        Else
                            strOS = strOS & " " & "Professional"
                        End If
                    Case 2
                        If .wProductType = VER_NT_WORKSTATION And sysinfo.wProcessorArchitecture = PROCESSOR_ARCHITECTURE_AMD64 Then
                            strOS = "Windows XP Professional x64 Edition"
                        ElseIf GetSystemMetrics(SM_SERVERR2) = 0 Then
                            strOS = "Windows Server 2003" '2003��3�·���
                        Else
                            strOS = "Windows Server 2003 R2"
                        End If
                        
                        If GetSystemMetrics(SM_SERVERR2) = 0 Then
                            If .wSuiteMask = VER_SUITE_BLADE Then
                                strOS = strOS & " " & "Web Edition"
                            ElseIf .wSuiteMask = VER_SUITE_COMPUTE_SERVER Then
                                strOS = strOS & " " & "Compute Cluster Edition"
                            ElseIf .wSuiteMask = VER_SUITE_STORAGE_SERVER Then
                                strOS = strOS & " " & "Storage Server"
                            ElseIf .wSuiteMask = VER_SUITE_DATACENTER Then
                                strOS = strOS & " " & "Datacenter Edition"
                            ElseIf .wSuiteMask = VER_SUITE_ENTERPRISE Then
                                strOS = strOS & " " & "Enterprise Edition"
                            End If
                        ElseIf .wSuiteMask = VER_SUITE_STORAGE_SERVER Then
                            strOS = strOS & " " & "Storage Server"
                        End If
                End Select
            Case 6
                Select Case .dwMinorVersion
                    Case 0
                        If .wProductType = VER_NT_WORKSTATION Then
                            strOS = "Microsoft Windows Vista"
                            If .wSuiteMask = VER_SUITE_PERSONAL Then
                                strOS = strOS & " " & "Home"
                            End If
                        Else
                            strOS = "Microsoft Windows Server 2008"
                            If .wSuiteMask = VER_SUITE_DATACENTER Then
                                strOS = strOS & " " & "Datacenter Server"
                            ElseIf .wSuiteMask = VER_SUITE_ENTERPRISE Then
                                strOS = strOS & " " & "Enterprise"
                            End If
                        End If
                    Case 1
                        If .wProductType = VER_NT_WORKSTATION Then
                            strOS = "Windows 7"
                        Else
                            strOS = "Windows Server 2008 R2"
                        End If
                    Case 2
                        If .wProductType = VER_NT_WORKSTATION Then
                            strOS = "Windows 8"
                        Else
                            strOS = "Windows Server 2012"
                        End If
                    Case 3
                        If .wProductType = VER_NT_WORKSTATION Then
                            strOS = "Windows 8.1"
                        Else
                            strOS = "Windows Server 2012 R2"
                        End If
                End Select
        End Select
    End With
    GetVersionInfo = strOS
End Function

Private Function CheckRepeatLogin(ByVal strIpAddress As String) As Boolean
    '����Ƿ����ظ���¼
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    Dim strProgram As String
    On Error GoTo Errhand
    
    strProgram = App.EXEName & ".exe"
    strSQL = "Select A.UserName, A.Program, B.IP" & vbNewLine & _
            "From gv$Session A, zlClients B" & vbNewLine & _
            "Where A.Terminal = B.����վ" & vbNewLine & _
            "      And A.Terminal = (Select Terminal From v$Session Where AudsID = Userenv('SessionID') and RowNum =1)" & vbNewLine & _
            "      And A.Program =[1] And A.AudsID <> Userenv('SessionID')" & vbNewLine & _
            "      And B.IP <> [2]"

    Set rsTemp = OpenSQLRecord(strSQL, "����ظ�����վ", strProgram, strIpAddress)
    If rsTemp.RecordCount = 0 Then '���Ե�¼
        CheckRepeatLogin = False
        Exit Function
    Else
        MsgBox "�������д�����ͬ���Ƶļ������¼," & vbCrLf & "�Է�IP��:[" & NVL(rsTemp!IP) & "]", vbInformation, gstrSysName
        CheckRepeatLogin = True
        Exit Function
    End If
    Exit Function
Errhand:
    MsgBox "���ͬ�����������" & Err.Description & ",����ϵ������Ա���н����", vbInformation, gstrSysName
End Function

Private Function GetCallEXE() As String
'���ܣ���ȡ���õ�ǰDLL��EXE����
    Dim strPName As String, strFileName As String

    strPName = String(256, Chr(0))
    Call GetModuleFileName(0, strPName, 256)
    strFileName = Left(strPName, InStr(strPName, Chr(0)) - 1)
    strFileName = UCase(Mid(strFileName, InStrRev(strFileName, "\") + 1))
    GetCallEXE = strFileName
End Function

Private Function GetLISStation() As String
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'����   �õ������°�LIS��վ��
'����   �õ�վ���վ������  ��Ϊû��վ��
'        �е���֯��ʽΪ ,1,2;,վ��1,վ��2
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim strվ��  As String, strվ������ As String
    
    On Error GoTo Errhand
    '�ж��Ƿ������װ
    strSQL = "select 1 ���� from zlsystems where ��� = 2500 and ����� is null"
    Set rsTmp = OpenSQLRecord(strSQL, "����Ƿ������װ�°�LIS")
    If rsTmp.EOF Then Exit Function
    '�����Ƿ���Ĭ�ϵ�վ��
    strSQL = "Select Distinct A.վ��, B.����" & vbNewLine & _
            "From (Select Distinct A.վ��" & vbNewLine & _
            "       From ����������¼ A, ����������Ա B, ��Ա�� C,�ϻ���Ա�� d" & vbNewLine & _
            "       Where A.Id = B.����id And A.վ�� Is Not Null And B.��Աid = C.Id and c.id = d.��ԱID And d.�û��� = [1]) A, Zlnodelist B" & vbNewLine & _
            "Where A.վ�� = B.���" & vbNewLine & _
            "Order By A.վ��"
    Set rsTmp = OpenSQLRecord(strSQL, "վ���ѯ", gclsLogin.DBUser)
    Do While Not rsTmp.EOF
        strվ�� = strվ�� & "," & rsTmp!վ��
        strվ������ = strվ������ & "," & rsTmp!����
        rsTmp.MoveNext
    Loop
    If strվ�� <> "" Then
        GetLISStation = strվ�� & ";" & strվ������
    End If
    Exit Function
Errhand:
    MsgBox "��ȡLIS����վ����" & Err.Description & ",����ϵ������Ա���н����", vbInformation, gstrSysName
End Function

Private Sub UpdateEmrInterface()
    Dim objEMR As Object
    
    If Not GetEMRLoginUser Then Exit Sub
    On Error Resume Next
    Err.Clear
    Set objEMR = CreateObject("zl9EmrInterface.ClsEmrInterface")
    If Err.Number = 0 Then
        Call objEMR.CheckUpdate1(gclsLogin.EMRUser, gclsLogin.EMRPwd, IIf(gstrCommand <> "", False, True))
        If Err.Number <> 0 Then
            Err.Clear
            Call objEMR.CheckUpdate(gclsLogin.EMRUser, gclsLogin.EMRPwd)
        End If
        Set gclsLogin.mobjEmr = objEMR
    Else
        Set gclsLogin.mobjEmr = Nothing
        Err.Clear
        Exit Sub
    End If
    On Error GoTo 0
End Sub

Private Function GetEMRLoginUser() As Boolean
'���ܣ���ȡEMP��ʼ�����û�������
'���أ��Ƿ��ȡ�ɹ���������ֻ����2500ϵͳ����������ļ���ȡ����������100��2500ϵͳ���򷵻�FALSE

    Dim strSQL      As String
    Dim rsTmp       As ADODB.Recordset
    Dim strConn     As String
    Dim objFSO      As New FileSystemObject
    Dim arrTmp      As Variant, arrInfo As Variant
    
    On Error GoTo errH
    strSQL = "Select Floor(a.��� / 100) ��� From zlSystems A Where Floor(a.��� / 100) In (1, 25)"
    Set rsTmp = OpenSQLRecord(strSQL, "GetEMRLoginUser")
    If rsTmp.RecordCount <> 0 Then
        rsTmp.Filter = "���=1"
        If rsTmp.RecordCount = 0 Then
            rsTmp.Filter = "���=25"
            If rsTmp.RecordCount <> 0 Then
                If objFSO.FileExists(App.Path & "\Apply\�ӿ�����.ini") Then
                    strConn = ReadIni("�ӿ�", "���ݷ���", App.Path & "\Apply\�ӿ�����.ini")
                    If strConn = "" Then Exit Function
                    strConn = Mid(strConn, 2)
                    strConn = DecipherV2("zlLis", strConn)
                    arrTmp = Split(strConn, ";")
                    If UBound(arrTmp) >= 1 Then
                        arrInfo = Split(arrTmp(1), ",")
                        If UBound(arrInfo) >= 1 Then
                            If arrInfo(0) <> "" And arrInfo(1) <> "" Then
                                gclsLogin.IsEMRProxy = True
                                gclsLogin.EMRUser = arrInfo(0)
                                gclsLogin.EMRPwd = IIf(UCase(arrInfo(0)) = "SYS" Or UCase(arrInfo(0)) = "SYSTEM", "[DBPASSWORD]", "") & arrInfo(1)
                                GetEMRLoginUser = True
                            End If
                        End If
                    End If
                End If
            End If
        Else
            gclsLogin.IsEMRProxy = False
            gclsLogin.EMRUser = gclsLogin.InputUser
            gclsLogin.EMRPwd = IIf(gclsLogin.IsTransPwd, "", "[DBPASSWORD]") & gclsLogin.InputPwd
            GetEMRLoginUser = True
        End If
    End If
    Exit Function
errH:
    Err.Clear
End Function

Public Function UpdateZLHIS(ByVal strComputerName As String, Optional ByVal blnBrwCall As Boolean, Optional ByVal blnForceUpdate As Boolean) As Boolean
'���ܣ�����ZLHIS��������
'      blnBrwCall=�Ƿ񵼺�̨����,����̨��������ʱ���Ԥ����ʱ��
    Dim strUpdateExe As String, strUpdateExePath As String
    Dim objFSO As New FileSystemObject
    Dim objConn As clsConnect, datCur           As Date
    Dim rsTemp As ADODB.Recordset, strSQL       As String
    Dim strJobName As String, blnDownload       As Boolean
    Dim strTmpPath As String, lngWait           As Long
    Dim strTmpGet  As String, blnMustNowUpdate  As Boolean
    
    glnghInstance = App.hInstance
    gstrExeFile = App.Path & "\" & App.EXEName & ".exe"
    gstrSetupPath = App.Path
    strUpdateExe = "zlHisCrust.exe"
    strTmpGet = gclsLogin.InputPwd
    'û�з��������û��ļ��嵥��������
    If Not IsHaveClientUpgradeSet(blnForceUpdate) Then '�ͻ����޸�ʱ��������Ϣ��ʾ��
        UpdateZLHIS = True
        Exit Function
    End If
    'û���������ռ����������Զ��˳�����
    If Not CheckJobs(strComputerName, strJobName, blnBrwCall, blnForceUpdate, blnMustNowUpdate) Then
        If blnForceUpdate Then
            MsgBox "��ǰֻ�ܽ���Ԥ�������޷����пͻ����޸���", vbInformation, gstrSysName
        Else
            UpdateZLHIS = True
        End If
        Exit Function
    End If
    
    If strJobName = "OfficialUpgrade" And blnBrwCall Then
        If blnMustNowUpdate Then
            MsgBox "��⵽ϵͳ��Ҫ������Ҫ�ĸ��£�1���Ӻ������������뼰ʱ����������д�����ݡ�", vbInformation, gstrSysName
            lngWait = 1 '���������ȴ�ʱ��
        Else
            If MsgBox("��⵽ϵͳ��Ҫ�������Ƿ���������?" & vbNewLine & "ѡ���������µ�¼����������", vbInformation + vbYesNo, gstrSysName) = vbNo Then
                UpdateZLHIS = True
                Exit Function
            End If
        End If
    End If
    If OS.IsDesinMode Then
        strUpdateExePath = "C:\APPSOFT\zlHisCrust.exe"
        strTmpPath = "C:\APPSOFT\ZLUPTMP"
    Else
        strUpdateExePath = gstrSetupPath & "\zlHisCrust.exe"
        strTmpPath = gstrSetupPath & "\ZLUPTMP"
    End If
    '�������򲻴��ڣ���׼������
    If Not objFSO.FileExists(strUpdateExePath) Then
        '��׼����ʱ����Ŀ¼
        If Not objFSO.FolderExists(strTmpPath) Then
            objFSO.CreateFolder (strTmpPath)
        End If
        strTmpPath = strTmpPath & "\" & Format(Now, "YYMMDDHHmmss")
        If Not objFSO.FolderExists(strTmpPath) Then
            Call objFSO.CreateFolder(strTmpPath)
        End If
        strTmpPath = strTmpPath & "\zlHisCrust.exe"
        Set objConn = New clsConnect
        If Not objConn.GetFileConnect(strComputerName) Then
            MsgBox "�޷����ӿͻ�������������""" & objConn.ServerPath & """,����ϵ����Ա��", vbExclamation, gstrSysName
            Exit Function
        End If
        blnDownload = objConn.DownloadFile("ZLHISCRUST.EXE", strTmpPath)
        If blnDownload Then
            objConn.CloseConnect
            On Error Resume Next
            '���������ļ�
            If objFSO.FileExists(strUpdateExePath) Then
                If FileSystem.GetAttr(strUpdateExePath) <> vbNormal Then
                     Call FileSystem.SetAttr(strUpdateExePath, vbNormal)
                End If
                Call objFSO.DeleteFile(strUpdateExePath)
            End If
            If Err.Number <> 0 Then Err.Clear
            '�ȸ��Ƶ�APPSOFT�£����ʧ�ܣ����Ƶ�APPLY��
            objFSO.CopyFile strTmpPath, strUpdateExePath, True
            If Err.Number <> 0 Then
                Err.Clear
                If OS.IsDesinMode Then
                    strUpdateExePath = "C:\APPSOFT\APPLY\zlHisCrust.exe"
                Else
                    strUpdateExePath = gstrSetupPath & "\APPLY\zlHisCrust.exe"
                End If
                '���������ļ�
                If objFSO.FileExists(strUpdateExePath) Then
                    If FileSystem.GetAttr(strUpdateExePath) <> vbNormal Then
                         Call FileSystem.SetAttr(strUpdateExePath, vbNormal)
                    End If
                    Call objFSO.DeleteFile(strUpdateExePath)
                End If
                If Err.Number <> 0 Then Err.Clear
                objFSO.CopyFile strTmpPath, strUpdateExePath, True
                If Err.Number <> 0 Then
                    Err.Clear
                    '�Ƿ����°��Զ�������ǣ��ǵĻ��������ֱ�Ӵ���ʱĿ¼������
                    If UCase(GetFileDesInfo(strTmpPath, "ProductName")) = "ZLHISINSTALLUPDATE" Then
                        strUpdateExePath = strTmpPath
                    End If
                End If
            End If
        End If
        If strTmpPath <> strUpdateExePath Then
            On Error Resume Next
            '��ʱ·��
            If objFSO.FileExists(strTmpPath) Then
                If FileSystem.GetAttr(strTmpPath) <> vbNormal Then
                     Call FileSystem.SetAttr(strTmpPath, vbNormal)
                End If
                Call objFSO.DeleteFile(strTmpPath)
            End If
            Call objFSO.DeleteFolder(objFSO.GetParentFolderName(strTmpPath))
        End If
        If Not objFSO.FileExists(strUpdateExePath) Then
            MsgBox "û���ҵ��ͻ����Զ���������" & strUpdateExe & "�����޷�ͨ���������������أ�����ϵͳ����Ա��ϵ��", vbExclamation, gstrSysName
            Exit Function
        End If
    End If
    'Ԥ���������ڵ���̨�����н���
    If StartHisCrust(strUpdateExePath, strJobName, lngWait, strTmpGet) Then
        If strJobName <> "PreUpgrade" Then
            Exit Function
        End If
    End If
    UpdateZLHIS = True
End Function

Private Function GetDefaultFileServer() As Integer
'���ܣ���ȡĬ�Ϸ�����
'���أ���û�з��������÷���-1�����ڣ������ⷵ��һ�����������
    Dim intDefaultSever As Integer, intServerType   As Integer
    Dim blnReadOld      As Boolean
    Dim strSQL          As String, rsTmp            As ADODB.Recordset
    
    On Error Resume Next
    intDefaultSever = -1
    strSQL = "Select ��� From Zltools.Zlupgradeserver Where �Ƿ����� = 1"
    Set rsTmp = OpenSQLRecord(strSQL, "��ȡ����������")
    If Err.Number <> 0 Then '���ܹ���Աʹ�õĹ�����������ͻ��˲�ƥ��
        Err.Clear
        blnReadOld = True
    ElseIf rsTmp.EOF Then
        blnReadOld = True
    End If
    On Error GoTo errH
    If Not blnReadOld Then
        intDefaultSever = Val(rsTmp!��� & "")
    Else
        strSQL = "select ���� from zlreginfo where ��Ŀ='��������'"
        Set rsTmp = OpenSQLRecord(strSQL, "���ʹ�õ���������")
        If Not rsTmp.EOF Then
            intServerType = Val(rsTmp!���� & "")
        End If
        If intServerType = 0 Then
            strSQL = "select replace(��Ŀ,'������Ŀ¼','') as ������ from zlreginfo where ��Ŀ like '������Ŀ¼%' and ���� is not null"
            Set rsTmp = OpenSQLRecord(strSQL, "����Ƿ�����õ��ļ����������")
        Else
            strSQL = "select replace(��Ŀ,'FTP������','') as ������ from zlreginfo where ��Ŀ like 'FTP������%' and ���� is not null"
            Set rsTmp = OpenSQLRecord(strSQL, "����Ƿ�����õ�FTP������")
        End If
        If Not rsTmp.EOF Then
            intDefaultSever = Val(rsTmp!������ & "")
        End If
    End If
    GetDefaultFileServer = intDefaultSever
    Exit Function
errH:
    GetDefaultFileServer = intDefaultSever
    If gblnTimer Then
        If ErrCenter() = 1 Then
            Resume
        End If
    Else
        MsgBox "��ȡȱʡ����������" & Err.Description, vbInformation, gstrSysName
        Err.Clear
    End If
End Function

Private Function IsHaveClientUpgradeSet(Optional ByVal blnMsg As Boolean) As Boolean
'���ܣ��Ƿ����������ص����á�
'������blnMsg=���ΪFalse��ʱ���Ƿ���ʾ
'���أ�IsHaveClientUpgradeSet=True:���ڿ������ļ�����������ã�False-����������һ��ȱʧ
    Dim intServerID As Integer
    Dim strSQL          As String, rsTmp            As ADODB.Recordset
    
    On Error GoTo errH
    IsHaveClientUpgradeSet = True
    '���ж��Ƿ���ڿ������ļ�
    strSQL = "Select 1 �������ļ� From Zltools.Zlfilesupgrade Where Md5 Is Not Null And Rownum < 2"
    Set rsTmp = OpenSQLRecord(strSQL, "����Ƿ���ڿ������ļ�")
    If Not rsTmp.EOF Then '�������ļ���������Ҫ��һ���ж��Ƿ�����������������
        intServerID = GetDefaultFileServer
        If intServerID = -1 Then
            If blnMsg Then
                MsgBox "û�����ÿͻ��������ļ����������޷����пͻ����޸���", vbInformation, gstrSysName
            End If
            IsHaveClientUpgradeSet = False
        End If
    Else
        If blnMsg Then
            MsgBox "��δ���������ļ��嵥���޷����пͻ����޸�������ϵ����Ա��", vbInformation, gstrSysName
        End If
        IsHaveClientUpgradeSet = False
    End If
    Exit Function
errH:
    IsHaveClientUpgradeSet = False
    If gblnTimer Then
        If ErrCenter() = 1 Then
            Resume
        End If
    Else
        MsgBox "��������ļ��嵥����" & Err.Description, vbInformation, gstrSysName
        Err.Clear
    End If
End Function

Private Function CheckJobs(ByVal strComputerName As String, ByRef strJobName As String, Optional ByVal blnBrwCall As Boolean, Optional ByVal blnForceUpdate As Boolean, Optional ByRef blnMustNowUpdate As Boolean) As Boolean
'����:��鲢��ȡ�������������
'      blnBrwCall=�Ƿ񵼺�̨����,����̨��������ʱ���Ԥ����ʱ��
'      blnForceUpdate=����̨����ͻ����޸�ʱ�ò���ΪTrue
'      blnMustNowUpdate=�Ƿ����ڱ�������
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim datCur As Date, blnOnlyOfficialUp As Boolean, blnOnlyPreUp As Boolean
    Dim blnPreUp As Boolean, blnOfficialUp As Boolean, blnPreComplete As Boolean, blnCollect As Boolean
    Dim strStartTime As String, strEndTime As String
    
    On Error GoTo errH
    strJobName = "": blnMustNowUpdate = False
    '���´���һ�㲻���ܳ���
    datCur = Currentdate
    '�ж������Ƿ������ȡ�Ƿ������˶�ʱ����
    strSQL = "Select Max(����) ���� From zlRegInfo Where ��Ŀ='�ͻ�����������'"
    Set rsTmp = OpenSQLRecord(strSQL, "��鶨ʱ����")
    If rsTmp!���� & "" <> "" Then
        If CDate(Format(datCur, "yyyy-MM-dd HH:mm:ss")) >= CDate(Format(NVL(rsTmp!����), "yyyy-MM-dd HH:mm:ss")) Then
            blnOnlyOfficialUp = True 'ֻ����ʽ����
        Else
            blnOnlyPreUp = True 'ֻ��Ԥ����
        End If
    Else
        blnOnlyOfficialUp = True
    End If
    On Error Resume Next
    Set rsTmp = Nothing
    '����û���Ƿ�Ԥ�����ֶ�(��ΪԤ����ʱ�����ݿ⻹û�������������Ҫ�������
    strSQL = "Select Ԥ��ʱ��,Nvl(�Ƿ�Ԥ����,0) �Ƿ�Ԥ����, Nvl(Ԥ�����, 0) Ԥ�����, Nvl(������־, 0) ������־, Nvl(�ռ���־, 0) �ռ���־,Nvl(�Ƿ���������,0) �Ƿ��������� From Zlclients Where ����վ = [1]"
    Set rsTmp = OpenSQLRecord(strSQL, "��鵱ǰ����", strComputerName)
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo errH
    If Not rsTmp Is Nothing Then
        If Not rsTmp.EOF Then
            blnPreUp = rsTmp!�Ƿ�Ԥ���� = 1
            blnOfficialUp = rsTmp!������־ = 1
            blnPreComplete = rsTmp!Ԥ����� = 1
            blnCollect = rsTmp!�ռ���־ = 1
            strStartTime = Format(datCur, "yyyy-mm-dd") & " " & Format(rsTmp!Ԥ��ʱ��, "HH:00:00")
            strEndTime = Format(datCur, "yyyy-mm-dd") & " " & Format(rsTmp!Ԥ��ʱ��, "HH:59:59")
            blnMustNowUpdate = rsTmp!�Ƿ��������� = 1
        End If
    Else
        '�����·�ʽ��ȡ��ʧ����ʹ���Ϸ�ʽ�����Ӽ�����
        strSQL = "Select Ԥ��ʱ��,Nvl(Ԥ�����, 0) Ԥ�����, Nvl(������־, 0) ������־, Nvl(�ռ���־, 0) �ռ���־ From Zlclients Where ����վ = [1]"
        Set rsTmp = OpenSQLRecord(strSQL, "��鵱ǰ����", strComputerName)
        If Not rsTmp.EOF Then
            blnPreUp = rsTmp!������־ = 1
            blnOfficialUp = rsTmp!������־ = 1
            blnPreComplete = rsTmp!Ԥ����� = 1
            blnCollect = rsTmp!�ռ���־ = 1
            strStartTime = Format(datCur, "yyyy-mm-dd") & " " & Format(rsTmp!Ԥ��ʱ��, "HH:00:00")
            strEndTime = Format(datCur, "yyyy-mm-dd") & " " & Format(rsTmp!Ԥ��ʱ��, "HH:59:59")
        End If
    End If
    '��ǰֻ�ܽ���Ԥ����
    If blnOnlyPreUp Then
        '��Ԥ��������
        If blnPreUp Or blnOfficialUp Then
            If Not blnPreComplete Then
                If datCur >= CDate(strStartTime) And datCur <= CDate(strEndTime) Then
                    strJobName = "PreUpgrade"
                Else
                    Exit Function
                End If
            Else
                Exit Function
            End If
        'û��Ԥ�������񣬵������ռ�����
        ElseIf blnCollect Then
            strJobName = "CollectClientFiles"
        Else
            Exit Function
        End If
    '��ǰֻ�ܽ�����ʽ����
    ElseIf blnOnlyOfficialUp Then
        If blnForceUpdate Then
            strJobName = "Repair"
        Else
            '����ʽ��������
            If blnOfficialUp Then
                strJobName = "OfficialUpgrade"
            'û����ʽ�������񣬵������ռ�����
            ElseIf blnCollect Then
                strJobName = "CollectClientFiles"
            Else
                Exit Function
            End If
        End If
    End If
    CheckJobs = True
    Exit Function
errH:
    If gblnTimer Then
        If ErrCenter() = 1 Then
            Resume
        End If
    Else
        MsgBox "���ͻ����������" & Err.Description, vbInformation, gstrSysName
        Err.Clear
    End If
End Function

Public Function DeCipher(ByVal strText As String) As String
'������ܳ���
    Const MIN_ASC = 32    '��СASCII��
    Const MAX_ASC = 126 '���ASCII�� �ַ�
    Const NUM_ASC = MAX_ASC - MIN_ASC + 1
    Dim lngOffset As Long, intLen As Integer, intSeedLen As Integer
    Dim intStart As Integer
    Dim i As Integer, intChr As Integer
    Dim strDeText As String
    
    If strText = "" Then Exit Function
    '������ӳ���
    intSeedLen = Asc(Mid(strText, 1, 1)) - MIN_ASC
    intLen = Len(strText)
    '���þɵ�����㷨
    If intSeedLen > 0 And intSeedLen < intLen - 3 And intSeedLen < 5 Then
        '��ȡ�������
        '������ӵ������Ϊ999
        Rnd (-1)
        Randomize (999)
        For i = 2 To 1 + intSeedLen
            intChr = Asc(Mid(strText, i, 1)) 'ȡ��ĸת���ASCII��
            If intChr >= MIN_ASC And intChr <= MAX_ASC Then
                intChr = intChr - MIN_ASC
                lngOffset = Int((NUM_ASC + 1) * Rnd())
                intChr = ((intChr - lngOffset) Mod NUM_ASC)
                If intChr < 0 Then
                    intChr = intChr + NUM_ASC
                End If
                intChr = intChr + MIN_ASC
                strDeText = strDeText & Chr(intChr)
            End If
        Next
        If Not IsNumeric(strDeText) Then
            strDeText = "123"
            intStart = 1
        Else
            intStart = 2 + intSeedLen
        End If
    Else
        strDeText = "123"
        intStart = 1
    End If
        
    '���ݽ��ܵ�����
    Rnd (-1)
    Randomize (Val(strDeText))
    strDeText = ""
    For i = intStart To intLen
        intChr = Asc(Mid(strText, i, 1)) 'ȡ��ĸת���ASCII��
        If intChr >= MIN_ASC And intChr <= MAX_ASC Then
            intChr = intChr - MIN_ASC
            lngOffset = Int((NUM_ASC + 1) * Rnd())
            intChr = ((intChr - lngOffset) Mod NUM_ASC)
            If intChr < 0 Then
                intChr = intChr + NUM_ASC
            End If
            intChr = intChr + MIN_ASC
            strDeText = strDeText & Chr(intChr)
        End If
    Next
    DeCipher = strDeText
End Function
Public Function DecipherV2(ByVal strPWD As String, ByVal strText As String) As String
    '����
    Const MIN_ASC = 32
    Const MAX_ASC = 126
    Const NUM_ASC = MAX_ASC - MIN_ASC + 1
    Dim lngOffset   As Long
    Dim intLen      As Integer
    Dim i           As Integer
    Dim intChr      As Integer
    Dim strReturn   As String
    
    lngOffset = NumericPassword(strPWD)
    Rnd -1
    Randomize lngOffset

    intLen = Len(strText)
    For i = 1 To intLen
        intChr = Asc(Mid$(strText, i, 1))
        If intChr >= MIN_ASC And intChr <= MAX_ASC Then
            intChr = intChr - MIN_ASC
            lngOffset = Int((NUM_ASC + 1) * Rnd)
            intChr = ((intChr - lngOffset) Mod NUM_ASC)
            If intChr < 0 Then intChr = intChr + NUM_ASC
            intChr = intChr + MIN_ASC
            strReturn = strReturn & Chr$(intChr)
        End If
    Next
    DecipherV2 = strReturn
End Function

Private Function NumericPassword(ByVal strPWD As String) As Long
    Dim lngValue    As Long
    Dim lngChr      As Long
    Dim lngShift1   As Long
    Dim lngShift2   As Long
    Dim i           As Integer
    Dim intLen     As Integer

    intLen = Len(strPWD)
    For i = 1 To intLen
        lngChr = Asc(Mid$(strPWD, i, 1))
        lngValue = lngValue Xor (lngChr * 2 ^ lngShift1)
        lngValue = lngValue Xor (lngChr * 2 ^ lngShift2)
        lngShift1 = (lngShift1 + 7) Mod 19
        lngShift2 = (lngShift2 + 13) Mod 23
    Next i
    NumericPassword = lngValue
End Function

Private Function ReadIni(strItem As String, strKey As String, strPath As String) As String
    Dim strReturn As String
    On Error GoTo errH

    strReturn = String(128, 0)
    GetPrivateProfileString strItem, strKey, "", strReturn, 256, strPath
    strReturn = Replace(strReturn, Chr(0), "")
    ReadIni = strReturn
    Exit Function
errH:
    Err.Clear
    ReadIni = ""
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
    GetWNetErr = Replace(Replace("[" & zlStr.TruncZero(strName) & "]" & zlStr.TruncZero(strErr), Chr(10), ""), Chr(13), "")
End Function

Private Function GetFileDesInfo(ByVal strFileName As String, ByVal strEntryName As String) As String
    Dim i               As Long
    Dim lngVerSize      As Long
    Dim bytVerBlock()   As Byte
    Dim strSubBlock  As String
    Dim bytTranslate()  As Byte, lngAdrTranslate    As Long, lngTranslateSize       As Long
    Dim bytBuffer()     As Byte, lngBuffer          As Long, lngAdrBuffer           As Long

    On Error GoTo errH
    lngVerSize = GetFileVersionInfoSize(strFileName, 0&)
    If lngVerSize <= 0 Then Exit Function
    
    ReDim bytVerBlock(lngVerSize - 1)
    Call GetFileVersionInfo(strFileName, 0&, lngVerSize, bytVerBlock(0))
    
    VerQueryValue VarPtr(bytVerBlock(0)), "\\VarFileInfo\\Translation", lngAdrTranslate, lngTranslateSize
    ReDim bytTranslate(lngTranslateSize - 1)
    CopyMemory bytTranslate(0), ByVal lngAdrTranslate, lngTranslateSize
    For i = 1 To lngTranslateSize / (UBound(bytTranslate) + 1)
        strSubBlock = "\\StringFileInfo\\"
        strSubBlock = strSubBlock & Byte2Hex(bytTranslate(), 0, 1, True)
        strSubBlock = strSubBlock & Byte2Hex(bytTranslate(), 2, 3, True)
        strSubBlock = strSubBlock & "\\" & strEntryName
        
        VerQueryValue VarPtr(bytVerBlock(0)), strSubBlock, lngAdrBuffer, lngBuffer
        If lngAdrBuffer <> 0 And lngBuffer <> 0 Then
            ReDim bytBuffer(lngBuffer - 1)
            CopyMemory bytBuffer(0), ByVal lngAdrBuffer, lngBuffer
            ReDim Preserve bytBuffer(InStrB(bytBuffer, ChrB(0)) - 2)
            GetFileDesInfo = StrConv(bytBuffer, vbUnicode)
        End If
    Next
    Exit Function
errH:
    Err.Clear
End Function

Private Function Byte2Hex(bytArray() As Byte, Optional ByVal lngStart As Long = 0, Optional ByVal lngEnd As Long = -1, Optional fReversed As Boolean = False) As String
    Dim i     As Long
    lngStart = IIf(lngStart < 0, 0, lngStart)
    lngEnd = IIf(lngEnd < 0, UBound(bytArray), lngEnd)
    
    If fReversed Then
        For i = lngEnd To lngStart Step -1
            Byte2Hex = Byte2Hex & Right$("00" & Hex(bytArray(i)), 2)
        Next
    Else
        For i = lngStart To lngEnd
            Byte2Hex = Byte2Hex & Right$("00" & Hex(bytArray(i)), 2)
        Next
    End If
End Function
