Attribute VB_Name = "mdlHisCrust"
Option Explicit

'���������������API
'----------------------------------------------------------------------------------------------------
Private Const PROCESSOR_INTEL_386 = 386
Private Const PROCESSOR_INTEL_486 = 486
Private Const PROCESSOR_INTEL_PENTIUM = 586
Private Const PROCESSOR_MIPS_R4000 = 4000
Private Const PROCESSOR_ALPHA_21064 = 21064
Private Type SYSTEM_INFO
    dwOemID As Long
    dwPageSize As Long
    lpMinimumApplicationAddress As Long
    lpMaximumApplicationAddress As Long
    dwActiveProcessorMask As Long
    dwNumberOrfProcessors As Long
    dwProcessorType As Long
    dwAllocationGranularity As Long
    dwReserved As Long
End Type
Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type
Private Type MEMORYSTATUS
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type

Private Const VER_PLATFORM_WIN32s = 0
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32_NT = 2
Private Const VER_NT_WORKSTATION = 1
Private Const VER_NT_DOMAIN_CONTROLLER = 2
Private Const VER_NT_SERVER = 3

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
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (ByRef lpVersionInformation As OSVERSIONINFOEX) As Long
Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)
Private Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As SYSTEM_INFO)
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


Public Function CheckAllowByTerminal() As Boolean
'����:����Ƿ�����ʹ�ñ�����վ,�Լ����е�ǰ����վ��Ϣ�ĵǼ�
'     �ж��Ƿ�����ù���վʹ�ó���
'     �����Ҫ�滻���ز�������ִ���滻�����������Ҫ�������������ǳ��򣬲��ر��˳�
'����:�ɹ�,����true,���򷵻�False
    
    Dim rsTmp As ADODB.Recordset, strSQL As String, strRowID As String '�ͻ��˵�ROWID
    Dim strComuterInfo As String, arrComputer As Variant, strComputerName As String, strIpAddress As String
    Dim strTmp As String, arrTmp As Variant, i As Integer
    Dim bln���վ�� As Boolean, lng��վ�� As Long, bln��վ�� As Boolean, bln��վ�� As Boolean
    Dim strվ��       As String, strվ���� As String, str���� As String, strȱʡ����
    Dim blnAllow As Boolean, blnUpdate As Boolean
    Dim int��������� As Integer, int������ƵԴ As Integer, int������ As Integer, int������־ As Integer
    
    Call SQLTest(App.EXEName, "mdlHisCrust", "�°���Ӳ����Զ��������")
    Call UpdateEmrInterface '�°���Ӳ����Զ�����
    Call SQLTest

    strIpAddress = Sys.IP '��oracle���ӵ�IP��ַΪ��
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
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��鹤��վ-��վ��Ϊ��", strComputerName)
    '��������δ��Ȩ��ԭ�򣬵��²�ѯ������ʱ������ʾ��ֹ��¼
    If rsTmp Is Nothing Then
        MsgBox Err.Description & vbNewLine & "������������ϵͳ��������ϵϵͳ����Ա���½��н�ɫ��Ȩ��", vbInformation, gstrSysName
        Exit Function
    End If
    '2.δ���ִ�վ��,����IP��ʽ���ң���ֻ��һ��ʱ�Ÿ��¼�����
    If rsTmp.EOF Then
        strSQL = "Select Rowid as ID,վ��,����, Nvl(��ֹʹ��,0) as ����,Nvl(������־,0) as ����,Nvl(�ռ���־,0) as �ռ�,������,������ƵԴ From zlClients Where IP=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��鹤��վ-��վ��Ϊ��", strIpAddress)
        If rsTmp.RecordCount > 1 Then
            '������������,���CPU,�ڴ�,Ӳ��Ϊ��������.
            strSQL = "" & _
                "   Select Rowid as ID,վ��,����,Nvl(��ֹʹ��,0) as ����,Nvl(������־,0) as ����,Nvl(�ռ���־,0) as �ռ�,������,������ƵԴ " & _
                "   From zlClients Where IP=[1] and CPU=[2] and  �ڴ�=[3] and Ӳ��=[4]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��鹤��վ-��վ��Ϊ��", strIpAddress, CStr(arrComputer(2)), CStr(arrComputer(3)), CStr(arrComputer(4)))
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
        If Command() <> "" Then
            '�·���
            If InStr(Command(), "ZLHISCRUSTCALL=1") > 0 And InStr(Command(), "USER=") > 0 And InStr(Command(), "PASS=") > 0 Then
                bln���վ�� = False
                strվ���� = NVL(rsTmp!վ��)
                gstrDeptName = NVL(rsTmp!����)
            '�ϵ��жϷ���
            ElseIf InStrRev(Command(), "/", -1) > 0 And InStrRev(Command(), ",", -1) = 0 Then
                bln���վ�� = False
                strվ���� = NVL(rsTmp!վ��)
                gstrDeptName = NVL(rsTmp!����)
            End If
        End If
        blnAllow = Val(rsTmp!���� & "") = 0
        int������ = Val(rsTmp!������ & "")  '0-��ʾ������
        blnUpdate = Val(rsTmp!���� & "") = 1
        If Not blnUpdate Then blnUpdate = Val(rsTmp!�ռ� & "") = 1
    End If

    If bln���վ�� Then
        strSQL = "select C.����,C.վ��,B.ȱʡ from �ϻ���Ա�� A,������Ա B, ���ű� C where A.��ԱID = B.��ԱID And B.����ID = C.ID And A.�û���=[1] order by C.վ��"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��鲢ȷ������Ժ��", gstrDbUser)
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
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "����Ƿ�����վ�����")
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
                strվ���� = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & App.EXEName, "��ǰվ��ѡ��", "")
                Call frmSelClient.ShowEdit(strվ��, str����, strվ����)
                strվ���� = IIf(frmSelClient.gstrվ�� = "��", "", frmSelClient.gstrվ��)
                gstrDeptName = frmSelClient.gstrCurվ��
                Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & App.EXEName, "��ǰվ��ѡ��", strվ����)
            End If
        End If
    End If
    If strվ���� <> "" Then
        Call zl9ComLib.SetNodeNo(strվ����)
    Else
        Call zl9ComLib.SetNodeNo("-")
    End If
    If gstrDeptName = "" Then gstrDeptName = strȱʡ����
    If strRowID = "" Then '�����Ĺ���վ����û�иù���վ�����ݣ��ϴ���IP����������CPU���ڴ桢Ӳ�̡�����ϵͳ��
        int������־ = 1
        strSQL = "select ���� from zlreginfo where ��Ŀ='��������'"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "���ʹ�õ���������")
        If Not rsTmp.EOF Then
            If NVL(rsTmp!����, 0) = 0 Then
                strSQL = "select replace(��Ŀ,'������Ŀ¼','') as ������ from zlreginfo where ��Ŀ like '������Ŀ¼%' and ���� is not null"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "����Ƿ�����õ��ļ����������")
                If rsTmp.EOF Then
                    int������־ = 0
                Else
                    int��������� = Val(rsTmp!������ & "")
                End If
            Else
                strSQL = "select replace(��Ŀ,'FTP������','') as FTP������ from zlreginfo where ��Ŀ like 'FTP������%' and ���� is not null"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "����Ƿ�����õ�FTP������")
                If rsTmp.EOF Then
                    int������־ = 0
                Else
                    int��������� = Val(rsTmp!FTP������ & "")
                End If
            End If
        End If
        strSQL = "Zl_Zlclients_Set(0,Null,'" & strComputerName & "','" & strIpAddress & "','" & arrComputer(2) & "','" & arrComputer(3) & _
                    "','" & arrComputer(4) & "','" & arrComputer(5) & "','" & gstrDeptName & "',Null,Null," & int��������� & "," & int������־ & _
                    ",0,'" & strվ���� & "',0,Null,Null," & int������ƵԴ & ")"
        zlDatabase.ExecuteProcedure strSQL, "��������վ"
        '�����ͻ��˲���������ֱ���˳�
        If int������־ = 0 Then
            CheckAllowByTerminal = True
            Exit Function
        End If
        blnUpdate = True
    Else
        strSQL = "Zl_Zlclients_Set(1,'" & strRowID & "','" & strComputerName & "','" & strIpAddress & "','" & arrComputer(2) & "','" & arrComputer(3) & _
                    "','" & arrComputer(4) & "','" & arrComputer(5) & "','" & gstrDeptName & "',Null,Null,Null,Null," & int������ & ",'" & strվ���� & "',0,Null,Null," & int������ƵԴ & ")"
        '��Ҫ������ص�վ����Ϣ
        zlDatabase.ExecuteProcedure strSQL, "���¹���վ"
        If Not blnAllow Then
            MsgBox "�ù���վ�ѱ�����Ա���ã�", vbInformation, gstrSysName
            Exit Function
        End If
        '�������������
        If int������ > 0 Then
            strSQL = "Select SID From gv$Session Where Upper(PROGRAM) Like 'ZL%.EXE' And Status<>'KILLED' And MACHINE=(Select Max(MACHINE) From v$Session Where AUDSID=UserENV('SessionID'))"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "�����������")
            If rsTmp.RecordCount > int������ Then
                MsgBox "��ǰ����վ���ֻ���� " & int������ & " ����¼���ӣ���ǰ�Ѿ��� " & rsTmp.RecordCount - 1 & " �����ӡ�", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End If
    On Error GoTo errHand
AutoUpGrude:      'ִ����������
    If blnUpdate Then
        blnAllow = UpdateZLHIS(strComputerName)
    End If
    CheckAllowByTerminal = blnAllow
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function StartHisCrust(ByVal str�������� As String, ByVal strJobName As String, Optional ByVal lngWait As Long) As Boolean
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
    
    On Error Resume Next
    If objFile.GetDriveName(str��������) = "" Then
        strUPFile = App.Path & "\" & str��������
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
        strCommand = "Provider=MSDataShape.1;Extended Properties=""Driver={Microsoft ODBC for Oracle};Server=" & gobjRelogin.ServerName & _
                                   """;Persist Security Info=True;User ID=" & gobjRelogin.InputUser & ";Password=HIS;Data Provider=MSDASQL"
    Else
        '��װ�����У��Լ�����������У��λ
        strCommand = "Provider=MSDataShape.1;Extended Properties=""Driver={Microsoft ODBC for Oracle};Server=" & gobjRelogin.ServerName & _
                                   """;Persist Security Info=True;User ID=" & gobjRelogin.InputUser & ";Password=" & gobjRegister.GetPassword(App.hInstance) & ";Data Provider=MSDASQL"
    End If
    strCheck = "CMDCHECK:1" & "," & Len(strCommand)
    strCommand = strCommand & "||0"
    strCheck = strCheck & "," & Len(strCommand)
    strCommand = strCommand & "||" & strJobName
    strCheck = strCheck & "," & Len(strCommand)
    strCommand = strCommand & "||" & CStr(Command())
    strCheck = strCheck & "," & Len(strCommand)
    strCommand = strCommand & "||" & "USER=" & gobjRelogin.InputUser & " PASS=" & gobjRelogin.InputPwd
    strCheck = strCheck & "," & Len(strCommand)
    If lngWait <> 0 Then
        strCommand = strCommand & "||" & lngWait
        strCheck = strCheck & "," & Len(strCommand)
    End If
    strCommand = strCommand & "||" & strCheck
    lngErr = Shell(strUPFile & " " & strCommand, vbNormalFocus)
    StartHisCrust = lngErr <> 0
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
    Dim verinfo As OSVERSIONINFO
    Dim sysinfo As SYSTEM_INFO
    Dim memsts As MEMORYSTATUS
    Dim memory As Long
    
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
    GlobalMemoryStatus memsts
    memory = memsts.dwTotalPhys
    strMemory = Format$(memory& \ 1024 \ 1024, "###,###,###") + "M"
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
    Dim sOS As String
    
    '�����Windows2000�����°汾��������API��ȡһ��
    myOS.dwOSVersionInfoSize = Len(myOS) 'should be 148/156
    'try win2000 version
    If GetVersionEx(myOS) = 0 Then
        'if fails
        myOS.dwOSVersionInfoSize = 148 'ignore reserved data
        If GetVersionEx(myOS) = 0 Then
            GetVersionInfo = "Windows (Unknown)"
            Exit Function
        End If
    Else
        bExInfo = True
    End If
    
    With myOS
        'is version 4
        If .dwPlatformId = VER_PLATFORM_WIN32_NT Then
            'nt platform
            Select Case .dwMajorVersion
            Case 3, 4
                sOS = "Windows NT"
            Case 5
                sOS = "Windows 2000"
            End Select
            If bExInfo Then
                'workstation/server?
                If .wProductType = VER_NT_SERVER Then
                    sOS = sOS & " Server"
                ElseIf .wProductType = VER_NT_DOMAIN_CONTROLLER Then
                    sOS = sOS & " Domain Controller"
                ElseIf .wProductType = VER_NT_WORKSTATION Then
                    sOS = sOS & IIf(.dwMajorVersion >= 5, " Professional", " WorkStation")
                End If
            End If
        ElseIf .dwPlatformId = VER_PLATFORM_WIN32_WINDOWS Then
            'get minor version info
            If .dwMinorVersion = 0 Then
                sOS = "Windows 95"
            ElseIf .dwMinorVersion = 10 Then
                sOS = "Windows 98"
            ElseIf .dwMinorVersion = 90 Then
                sOS = "Windows Millenium"
            Else
                sOS = "Windows 9?"
            End If
        End If
    End With
    GetVersionInfo = sOS
End Function

Private Function CheckRepeatLogin(ByVal strIpAddress As String) As Boolean
    '����Ƿ����ظ���¼
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    Dim strProgram As String
    On Error GoTo errHand
    
    strProgram = App.EXEName & ".exe"
    strSQL = "Select A.UserName, A.Program, B.IP" & vbNewLine & _
            "From gv$Session A, zlClients B" & vbNewLine & _
            "Where A.Terminal = B.����վ" & vbNewLine & _
            "      And A.Terminal = (Select Terminal From v$Session Where AudsID = Userenv('SessionID') and RowNum =1)" & vbNewLine & _
            "      And A.Program =[1] And A.AudsID <> Userenv('SessionID')" & vbNewLine & _
            "      And B.IP <> [2]"

    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "����ظ�����վ", strProgram, strIpAddress)
    If rsTemp.RecordCount = 0 Then '���Ե�¼
        CheckRepeatLogin = False
        Exit Function
    Else
        MsgBox "�������д�����ͬ���Ƶļ������¼," & vbCrLf & "�Է�IP��:[" & NVL(rsTemp!IP) & "]", vbInformation, gstrSysName
        CheckRepeatLogin = True
        Exit Function
    End If
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Function GetLISStation() As String
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'����   �õ������°�LIS��վ��
'����   �õ�վ���վ������  ��Ϊû��վ��
'        �е���֯��ʽΪ ,1,2;,վ��1,վ��2
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim strվ��  As String, strվ������ As String
    
    On Error GoTo errHand
    '�ж��Ƿ������װ
    strSQL = "select 1 ���� from zlsystems where ��� = 2500 and ����� is null"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "����Ƿ������װ�°�LIS")
    If rsTmp.EOF Then Exit Function
    '�����Ƿ���Ĭ�ϵ�վ��
    strSQL = "Select Distinct A.վ��, B.����" & vbNewLine & _
            "From (Select Distinct A.վ��" & vbNewLine & _
            "       From ����������¼ A, ����������Ա B, ��Ա�� C,�ϻ���Ա�� d" & vbNewLine & _
            "       Where A.Id = B.����id And A.վ�� Is Not Null And B.��Աid = C.Id and c.id = d.��ԱID And d.�û��� = [1]) A, Zlnodelist B" & vbNewLine & _
            "Where A.վ�� = B.���" & vbNewLine & _
            "Order By A.վ��"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "վ���ѯ", gstrDbUser)
    Do While Not rsTmp.EOF
        strվ�� = strվ�� & "," & rsTmp!վ��
        strվ������ = strվ������ & "," & rsTmp!����
        rsTmp.MoveNext
    Loop
    If strվ�� <> "" Then
        GetLISStation = strվ�� & ";" & strվ������
    End If
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub UpdateEmrInterface()
    Dim objEMR As Object
    
    On Error Resume Next
    Err.Clear
    Set objEMR = CreateObject("zl9EmrInterface.ClsEmrInterface")
    If Err.Number = 0 Then
        Call objEMR.CheckUpdate1(gobjRelogin.InputUser, IIf(gobjRelogin.IsTransPwd, "", "[DBPASSWORD]") & gobjRelogin.InputPwd, IIf(CStr(Command()) <> "", False, True))
        If Err.Number <> 0 Then
            Err.Clear
            Call objEMR.CheckUpdate(gobjRelogin.InputUser, IIf(gobjRelogin.IsTransPwd, "", "[DBPASSWORD]") & gobjRelogin.InputPwd)
        End If
        Set gobjRelogin.EMR = objEMR
    Else
        Set gobjRelogin.EMR = Nothing
        Err.Clear
        Exit Sub
    End If
    On Error GoTo 0
End Sub

Public Function UpdateZLHIS(ByVal strComputerName As String, Optional ByVal blnBrwCall As Boolean, Optional ByVal blnForceUpdate As Boolean) As Boolean
'���ܣ�����ZLHIS��������
'      blnBrwCall=�Ƿ񵼺�̨����,����̨��������ʱ���Ԥ����ʱ��

    Dim strUpdateExe As String, strUpdateExePath As String
    Dim objFSO As New FileSystemObject
    Dim objConn As clsConnect, datCur As Date
    Dim rsTemp As ADODB.Recordset, strSQL As String
    Dim strJobName As String, blnDownload As Boolean
    Dim strTmpPath As String, lngWait As Long
    
    'û���������ռ����������Զ��˳�����
    If Not CheckJobs(strComputerName, strJobName, blnBrwCall, blnForceUpdate) Then
        If blnForceUpdate Then
            MsgBox "��ǰֻ�ܽ���Ԥ�������޷����пͻ����޸���", vbInformation, gstrSysName
        Else
            UpdateZLHIS = True
        End If
        Exit Function
    End If
    
    If strJobName = "OfficialUpgrade" And blnBrwCall Then
        MsgBox "��⵽ϵͳ��Ҫ���������ϵͳ������������", vbInformation, gstrSysName
'        lngWait = 1
    End If
    strUpdateExe = "zlHisCrust.exe"
    If OS.IsDesinMode Then
        strUpdateExePath = "C:\APPSOFT\zlHisCrust.exe"
        strTmpPath = "C:\APPSOFT\ZLUPTMP"
    Else
        strUpdateExePath = App.Path & "\zlHisCrust.exe"
        strTmpPath = App.Path & "\ZLUPTMP"
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
            MsgBox "û���ҵ��ͻ����Զ���������" & strUpdateExe & "�����޷�ͨ���������������أ�����ϵͳ����Ա��ϵ��", vbExclamation, gstrSysName
            Exit Function
        End If
        blnDownload = objConn.DownloadFile("zlHisCrust.exe", strTmpPath)
        If blnDownload Then
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
                    strUpdateExePath = App.Path & "\APPLY\zlHisCrust.exe"
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
    Call SaveSetting("ZLSOFT", "����ȫ��", "��������", UCase(strUpdateExe)) '����ZLRegister�������ж�
    If StartHisCrust(strUpdateExePath, strJobName, lngWait) And strJobName <> "PreUpgrade" Then
        End
    End If
    UpdateZLHIS = True
End Function

Private Function CheckJobs(ByVal strComputerName As String, ByRef strJobName As String, Optional ByVal blnBrwCall As Boolean, Optional ByVal blnForceUpdate As Boolean) As Boolean
'����:��鲢��ȡ�������������
'      blnBrwCall=�Ƿ񵼺�̨����,����̨��������ʱ���Ԥ����ʱ��
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim datCur As Date, blnOnlyOfficialUp As Boolean, blnOnlyPreUp As Boolean
    Dim blnPreUp As Boolean, blnOfficialUp As Boolean, blnPreComplete As Boolean, blnCollect As Boolean
    Dim strStartTime As String, strEndTime As String
    
    On Error GoTo errH
    strJobName = ""
    '���´���һ�㲻���ܳ���
    datCur = zlDatabase.Currentdate
    '�ж������Ƿ������ȡ�Ƿ������˶�ʱ����
    strSQL = "Select Max(����) ���� From zlRegInfo Where ��Ŀ='�ͻ�����������'"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��鶨ʱ����")
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
    strSQL = "Select Ԥ��ʱ��,Nvl(�Ƿ�Ԥ����,0) �Ƿ�Ԥ����, Nvl(Ԥ�����, 0) Ԥ�����, Nvl(������־, 0) ������־, Nvl(�ռ���־, 0) �ռ���־ From Zlclients Where ����վ = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��鵱ǰ����", strComputerName)
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
        End If
    Else
        '�����·�ʽ��ȡ��ʧ����ʹ���Ϸ�ʽ�����Ӽ�����
        strSQL = "Select Ԥ��ʱ��,Nvl(Ԥ�����, 0) Ԥ�����, Nvl(������־, 0) ������־, Nvl(�ռ���־, 0) �ռ���־ From Zlclients Where ����վ = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��鵱ǰ����", strComputerName)
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
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function DeCipher(ByVal strText As String) As String
'������ܳ���
    Const MIN_ASC = 32    '��СASCII��
    Const MAX_ASC = 126 '���ASCII�� �ַ�
    Const NUM_ASC = MAX_ASC - MIN_ASC + 1
    Dim lngOffset As Long, intlen As Integer, intSeedLen As Integer
    Dim intStart As Integer
    Dim i As Integer, intChr As Integer
    Dim strDeText As String
    
    If strText = "" Then Exit Function
    '������ӳ���
    intSeedLen = Asc(Mid(strText, 1, 1)) - MIN_ASC
    intlen = Len(strText)
    '���þɵ�����㷨
    If intSeedLen > 0 And intSeedLen < intlen - 3 And intSeedLen < 5 Then
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
    For i = intStart To intlen
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


