Attribute VB_Name = "mdlHisCrust"
'Option Explicit
'
''���������������API
''----------------------------------------------------------------------------------------------------
'Private Const PROCESSOR_INTEL_386 = 386
'Private Const PROCESSOR_INTEL_486 = 486
'Private Const PROCESSOR_INTEL_PENTIUM = 586
'Private Const PROCESSOR_MIPS_R4000 = 4000
'Private Const PROCESSOR_ALPHA_21064 = 21064
'Private Type SYSTEM_INFO
'    dwOemID As Long
'    dwPageSize As Long
'    lpMinimumApplicationAddress As Long
'    lpMaximumApplicationAddress As Long
'    dwActiveProcessorMask As Long
'    dwNumberOrfProcessors As Long
'    dwProcessorType As Long
'    dwAllocationGranularity As Long
'    dwReserved As Long
'End Type
'Private Type OSVERSIONINFO
'    dwOSVersionInfoSize As Long
'    dwMajorVersion As Long
'    dwMinorVersion As Long
'    dwBuildNumber As Long
'    dwPlatformId As Long
'    szCSDVersion As String * 128
'End Type
'Private Type MEMORYSTATUS
'    dwLength As Long
'    dwMemoryLoad As Long
'    dwTotalPhys As Long
'    dwAvailPhys As Long
'    dwTotalPageFile As Long
'    dwAvailPageFile As Long
'    dwTotalVirtual As Long
'    dwAvailVirtual As Long
'End Type
'
'Private Const VER_PLATFORM_WIN32s = 0
'Private Const VER_PLATFORM_WIN32_WINDOWS = 1
'Private Const VER_PLATFORM_WIN32_NT = 2
'Private Const VER_NT_WORKSTATION = 1
'Private Const VER_NT_DOMAIN_CONTROLLER = 2
'Private Const VER_NT_SERVER = 3
'Private Type OSVERSIONINFOEX
'    dwOSVersionInfoSize As Long
'    dwMajorVersion As Long
'    dwMinorVersion As Long
'    dwBuildNumber As Long
'    dwPlatformId As Long
'    szCSDVersion As String * 128      '  Maintenance string for PSS usage
'    wServicePackMajor As Integer 'win2000 only
'    wServicePackMinor As Integer 'win2000 only
'    wSuiteMask As Integer 'win2000 only
'    wProductType As Byte 'win2000 only
'    wReserved As Byte
'End Type
'
'Private Declare Function GetIpAddrTable Lib "IPHlpApi" (pIPAdrTable As Byte, pdwSize As Long, ByVal Sort As Long) As Long
'Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (ByRef lpVersionInformation As OSVERSIONINFOEX) As Long
'Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)
'Private Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As SYSTEM_INFO)
'Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" _
'    (ByVal lpBuffer As String, nSize As Long) As Long
'
''ȡIP��API
'Private Const MAX_ADAPTER_NAME_LENGTH         As Long = 256
'Private Const MAX_ADAPTER_DESCRIPTION_LENGTH  As Long = 128
'Private Const MAX_ADAPTER_ADDRESS_LENGTH      As Long = 8
'Private Const ERROR_SUCCESS  As Long = 0
'Private Const MAX_IP = 5   'To make a buffer... i dont think you have more than 5 ip on your pc..
'Private Type IPINFO
'     dwAddr As Long   ' IP address
'    dwIndex As Long ' interface index
'    dwMask As Long ' subnet mask
'    dwBCastAddr As Long ' broadcast address
'    dwReasmSize  As Long ' assembly size
'    unused1 As Integer ' not currently used
'    unused2 As Integer '; not currently used
'End Type
'Private Type MIB_IPADDRTABLE
'    dEntrys As Long   'number of entries in the table
'    mIPInfo(MAX_IP) As IPINFO  'array of IP address entries
'End Type
'Private Type IP_Array
'    mBuffer As MIB_IPADDRTABLE
'    BufferLen As Long
'End Type
'Private Type IP_ADDRESS_STRING
'    IpAddr(0 To 15)  As Byte
'End Type
'Private Type IP_MASK_STRING
'    IpMask(0 To 15)  As Byte
'End Type
'Private Type IP_ADDR_STRING
'    dwNext     As Long
'    IpAddress  As IP_ADDRESS_STRING
'    IpMask     As IP_MASK_STRING
'    dwContext  As Long
'End Type
'Private Type IP_ADAPTER_INFO
'  dwNext                As Long
'  ComboIndex            As Long  '����
'  sAdapterName(0 To (MAX_ADAPTER_NAME_LENGTH + 3))        As Byte
'  sDescription(0 To (MAX_ADAPTER_DESCRIPTION_LENGTH + 3)) As Byte
'  dwAddressLength       As Long
'  sIPAddress(0 To (MAX_ADAPTER_ADDRESS_LENGTH - 1))       As Byte
'  dwIndex               As Long
'  uType                 As Long
'  uDhcpEnabled          As Long
'  CurrentIpAddress      As Long
'  IpAddressList         As IP_ADDR_STRING
'  GatewayList           As IP_ADDR_STRING
'  DhcpServer            As IP_ADDR_STRING
'  bHaveWins             As Long
'  PrimaryWinsServer     As IP_ADDR_STRING
'  SecondaryWinsServer   As IP_ADDR_STRING
'  LeaseObtained         As Long
'  LeaseExpires          As Long
'End Type
'Private Declare Function GetAdaptersInfo Lib "iphlpapi.dll" _
'    (pTcpTable As Any, pdwSize As Long) As Long
'Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
'(Destination As Any, Source As Any, ByVal Length As Long)
'
''ȡӲ�̴�С
'Private Const DRIVE_UNKNOWN = 0
'Private Const DRIVE_ABSENT = 1
'Private Const DRIVE_REMOVABLE = 2
'Private Const DRIVE_FIXED = 3
'Private Const DRIVE_REMOTE = 4
'Private Const DRIVE_CDROM = 5
'Private Const DRIVE_RAMDISK = 6
'Private Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTotalNumberOfClusters As Long) As Long
'Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
'
''����������ڵ���ע�����ָ����֧��һ���ļ���
''regedit /e d:\Win2000.reg "HKEY_CURRENT_USER\SOFTWARE\VB AND VBA PROGRAM SETTIGS\ZLSOFT"
''----------------------------------------------------------------------------------------------------
'Private Const STRSPLIT As String = "���"
'Private Const REGCMD As String = "REGEDIT /E"
'Private Const RegFile As String = "C:\REGFILE.REG"
'Private Const REGDATA As String = "C:\REGDATA.REG"
'Private Const REGDIRECTORY As String = """HKEY_CURRENT_USER\SOFTWARE\VB AND VBA PROGRAM SETTINGS\ZLSOFT"""
'
''����������ڼ���Ƿ�Ϸ�����
'Private Declare Function GlobalAddAtom Lib "kernel32" Alias "GlobalAddAtomA" (ByVal lpString As String) As Integer
'Private Declare Function GlobalDeleteAtom Lib "kernel32" (ByVal nAtom As Integer) As Integer
'Private Declare Function GlobalGetAtomName Lib "kernel32" Alias "GlobalGetAtomNameA" (ByVal nAtom As Integer, ByVal lpBuffer As String, ByVal nSize As Long) As Long
'
'
''����Ϊ���˺����:��Ҫ��ѡ���ļ���
''20060606
'Private Const OFS_MAXPATHNAME = 128
'Private Const OF_EXIST = &H4000
'
'Private Type OFSTRUCT
'        cBytes As Byte
'        fFixedDisk As Byte
'        nErrCode As Integer
'        Reserved1 As Integer
'        Reserved2 As Integer
'        szPathName(OFS_MAXPATHNAME) As Byte
'End Type
'Private Declare Function apiOpenFile Lib "kernel32" Alias "OpenFile" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long
'
'
'''''''���20110111''''''''������������
'Private Type NETRESOURCE
'    dwScope As Long
'    dwType As Long
'    dwDisplayType As Long
'    dwUsage As Long
'    lpLocalName As String
'    lpRemoteName As String
'    lpComment As String
'    lpProvider As String
'End Type
'Private Const INFINITE = -1&
'Private Const SYNCHRONIZE = &H100000
'
'Const NO_ERROR = 0
'Const CONNECT_UPDATE_PROFILE = &H1
'Const RESOURCETYPE_DISK = &H1
'Const RESOURCETYPE_PRINT = &H2
'Const RESOURCETYPE_ANY = &H0
'Const RESOURCE_CONNECTED = &H1
'Const RESOURCE_REMEMBERED = &H3
'Const RESOURCE_GLOBALNET = &H2
'Const RESOURCEDISPLAYTYPE_DOMAIN = &H1
'Const RESOURCEDISPLAYTYPE_GENERIC = &H0
'Const RESOURCEDISPLAYTYPE_SERVER = &H2
'Const RESOURCEDISPLAYTYPE_SHARE = &H3
'Const RESOURCEUSAGE_CONNECTABLE = &H1
'Const RESOURCEUSAGE_CONTAINER = &H2
'
'Private Declare Function WNetAddConnection2 Lib "mpr.dll" Alias _
'        "WNetAddConnection2A" _
'        (lpNetResource As NETRESOURCE, _
'        ByVal lpPassword As String, _
'        ByVal lpUserName As String, _
'        ByVal dwFlags As Long) As Long
'
'Private Declare Function WNetCancelConnection2 Lib "mpr.dll" Alias _
'        "WNetCancelConnection2A" _
'        (ByVal lpName As String, _
'        ByVal dwFlags As Long, _
'        ByVal fForce As Long) As Long
'
'
'''''''���20110111''''''''������������
'
'
'Public Function �Ƿ�����ʹ�ñ�����վ(Optional ByRef strIpAddress As String) As Boolean
'    '-----------------------------------------------------------------------------------------------------------
'    '����:����Ƿ�����ʹ�ñ�����վ��վ����Ϣ���ϴ�
'    '     �ж��Ƿ�����ù���վʹ�ó��������Ҫ�滻���ز�������ִ���滻�����������Ҫ�������������ǳ��򣬲��ر��˳�
'    '���:
'    '����:
'    '����:�ɹ�,����true,���򷵻�False
'    '����:���˺�
'    '����:2009-01-21 11:59:49
'    '-----------------------------------------------------------------------------------------------------------
'    Dim objFileSys As New FileSystemObject, rsClients As New ADODB.Recordset, rsTemp As ADODB.Recordset
'    Dim strSQL As String, strInfo As String, strCurrDate As String, strExeName As String
'    Dim str�������� As String, Error As Long, strComputerName As String, strRowID As String 'IP��ַ��վ����
'    Dim blnAllow As Boolean, blnUpdate As Boolean, int������ As Integer, int������־ As Integer
'    Dim int��������� As Integer, i As Integer
'    Dim strվ��       As String, strSouceվ�� As String, strվ���� As String, str���� As String, strȱʡ As String, strȱʡ����
'    Dim strSplitվ��()  As String, blnվ�� As Boolean, bln������ʽ As Boolean, bln���վ�� As Boolean, strCurIndex As String
'    Dim lng��վ�� As Long, bln��վ�� As Boolean
'    Err = 0: On Error Resume Next
'
'    strIpAddress = ""
'    blnAllow = False: blnUpdate = False: str�������� = "zlHisCrust.exe": �Ƿ�����ʹ�ñ�����վ = False
'    strExeName = GetSetting("ZLSOFT", "����ȫ��", "ִ���ļ�", "")
'
'    '�ж��Ƿ�����ʹ��
'    strComputerName = AnalyseComputer           '�����������
'    strInfo = AnalyseConfigure: strIpAddress = Split(strInfo, STRSPLIT)(0)
'    strIpAddress = zl_Ip_Address_FromOrc(strIpAddress)  '��oracle���ӵ�IP��ַΪ��
'
'    ''''''ZQ20101109''''''''''
'    ''''''����Ƿ�����������''''''
'    If CheckRepeatLogin() = True Then
'        �Ƿ�����ʹ�ñ�����վ = False
'        Exit Function
'    End If
'
'
''''    '���ͻ���վ���Ƿ�ΪNULL
''''    'ף��:2010-12-24 10:00:00
''''    strSQL = "select վ�� from zlclients where IP= Sys_Context('USERENV', 'IP_ADDRESS') and ����վ= SYS_CONTEXT('USERENV','TERMINAL')"
''''    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "���ͻ�վ���Ƿ�ȷ��")
''''    If rsTemp.RecordCount = 1 Then
''''        bln���վ�� = IIf(zlCommFun.NVL(rsTemp!վ��) = "", True, False)
''''        strվ���� = zlCommFun.NVL(rsTemp!վ��)
''''    Else
''''        bln���վ�� = True
''''        strվ���� = ""
''''    End If
'    '�������½,�������û�ѡ��,ֱ�Ӷ�ȡ ZQ20110114
'    '�������/����ʾ���û�������ĸ�ʽ���磺zlhis/his
'    If InStrRev(Command(), "/", -1) > 0 Then
'         bln���վ�� = False
'         strSQL = "select վ��,���� from zlclients where IP= Sys_Context('USERENV', 'IP_ADDRESS') and ����վ= SYS_CONTEXT('USERENV','TERMINAL')"
'         Set rsTemp = zlDataBase.OpenSQLRecord(strSQL, "���ͻ�վ���Ƿ�ȷ��")
'         If rsTemp.RecordCount = 1 Then
'            strվ���� = zlCommFun.NVL(rsTemp!վ��)
'            gstrDeptName = zlCommFun.NVL(rsTemp!����)
'         Else
'            strվ���� = ""
'            gstrDeptName = ""
'         End If
'    Else
'         bln���վ�� = True
'    End If
'
'
'    If bln���վ�� Then
'        strSQL = "select C.����,C.վ��,B.ȱʡ from �ϻ���Ա�� A,������Ա B, ���ű� C where A.��ԱID = B.��ԱID And B.����ID = C.ID And upper(A.�û���)=upper([1]) order by C.����"
'        Set rsTemp = zlDataBase.OpenSQLRecord(strSQL, "��鲢ȷ������Ժ��", gstrDbUser)
'        Do While Not rsTemp.EOF
'            If strվ�� = "" Then
'                If zlCommFun.NVL(rsTemp!վ��, "") <> "" Then
'                    strվ�� = zlCommFun.NVL(rsTemp!վ��, "") & ","
'                    str���� = zlCommFun.NVL(rsTemp!����) & ","
'                    lng��վ�� = lng��վ�� + 1
'                Else
'                    bln��վ�� = True
'                End If
'            Else
'                If zlCommFun.NVL(rsTemp!վ��, "") <> "" Then
'                    strվ�� = strվ�� & zlCommFun.NVL(rsTemp!վ��, "") & ","
'                    str���� = str���� & zlCommFun.NVL(rsTemp!����) & ","
'                    lng��վ�� = lng��վ�� + 1
'                Else
'                    bln��վ�� = True
'                End If
'            End If
'            If zlCommFun.NVL(rsTemp!ȱʡ, "0") = 1 Then
'                strȱʡ���� = zlCommFun.NVL(rsTemp!����)
'            End If
'            rsTemp.MoveNext
'        Loop
'
'        '  strվ�� = ""�����ǰ��¼��Ա�������Ŷ�û������վ�㣬���������ڲ��Ҹ�Ժ�Ƿ�������վ�����!
'        If strվ�� = "" Or (bln��վ�� And lng��վ�� <> 1) Then
'            strվ�� = ""
'            strSQL = "select distinct (A.վ��),B.���� from ���ű� A,zlNodeList B where A.վ��=B.��� And A.վ�� is not null order by A.վ��"
'            Set rsTemp = zlDataBase.OpenSQLRecord(strSQL, "����Ƿ�����վ�����")
'            Do While Not rsTemp.EOF
'                If strվ�� = "" Then
'                    If zlCommFun.NVL(rsTemp!վ��, "") <> "" Then
'                        strվ�� = zlCommFun.NVL(rsTemp!վ��, "") & ","
'                        str���� = zlCommFun.NVL(rsTemp!����) & ","
'                    End If
'                Else
'                    If zlCommFun.NVL(rsTemp!վ��, "") <> "" Then
'                        strվ�� = strվ�� & zlCommFun.NVL(rsTemp!վ��, "") & ","
'                        str���� = str���� & zlCommFun.NVL(rsTemp!����) & ","
'                    End If
'                End If
'                rsTemp.MoveNext
'            Loop
'        End If
'
'        If strվ�� <> "" Then
'            strSplitվ�� = Split(strվ��, ",")
'            For i = 0 To UBound(strSplitվ��) - 1
'                If i = 0 Then
'                    strSouceվ�� = strSplitվ��(i)
'                Else
'                    If strSouceվ�� <> strSplitվ��(i) Then
'                        blnվ�� = True
'                        Exit For
'                    End If
'                End If
'            Next
'
'            If blnվ�� Then
'            'blnվ�� = True ��ǰ��¼��Ա�������Ű������վ��,��ʾ�û�ѡ��ǰ�����λ�����ڵĲ��š�
'                strCurIndex = GetRegister(˽��ģ��, App.EXEName, "��ǰվ��ѡ��", "")
'                Call frmSelClient.ShowEdit(strվ��, str����, strCurIndex)
'                strվ���� = IIf(frmSelClient.gstrվ�� = "��", "", frmSelClient.gstrվ��)
'                gstrDeptName = IIf(blnվ��, strȱʡ����, frmSelClient.gstrCurվ��)
'                Call SetRegister(˽��ģ��, App.EXEName, "��ǰվ��ѡ��", strվ����)
'            Else
'            'blnվ��= False ��ǰ��¼��Ա�������Ŷ�����ͬ��վ�㣬���Ը�վ���ű�����"zlClients.վ��"�С�
'                strվ���� = strSouceվ��
'                gstrDeptName = strȱʡ����
'            End If
'
'        End If
'    End If
'    If strվ���� <> "" Then
'        zlComLib.gstrNodeNo = strվ����
'    Else
'        zlComLib.gstrNodeNo = "-"
'        gstrDeptName = strȱʡ����
'    End If
'
'    '���������м��.����:15640
'    '1.��վ�������
'    strSQL = "Select Rowid as ID,Nvl(��ֹʹ��,0) as ����,Nvl(������־,0) as ����,Nvl(�ռ���־,0) as �ռ�,������ From zlClients Where ����վ=[1]"
'    Set rsClients = zlDataBase.OpenSQLRecord(strSQL, "��鹤��վ-��վ��Ϊ��", strComputerName)
'    If rsClients.EOF Then
'        '2.δ���ִ�վ��,����IP��ʽ���ң���ֻ��һ��ʱ�Ÿ��¼�����
'        strSQL = "Select Rowid as ID, Nvl(��ֹʹ��,0) as ����,Nvl(������־,0) as ����,Nvl(�ռ���־,0) as �ռ�,������ From zlClients Where IP=[1]"
'        Set rsClients = zlDataBase.OpenSQLRecord(strSQL, "��鹤��վ-��վ��Ϊ��", strIpAddress)
'        If rsClients.RecordCount > 1 Then
'            '������������,���CPU,�ڴ�,Ӳ��Ϊ��������.
'            strSQL = "" & _
'                "   Select Rowid as ID, Nvl(��ֹʹ��,0) as ����,Nvl(������־,0) as ����,Nvl(�ռ���־,0) as �ռ�,������ " & _
'                "   From zlClients Where IP=[1] and CPU=[2] and  �ڴ�=[3] and Ӳ��=[4]"
'            Set rsClients = zlDataBase.OpenSQLRecord(strSQL, "��鹤��վ-��վ��Ϊ��", strIpAddress, CStr(Split(strInfo, STRSPLIT)(2)), CStr(Split(strInfo, STRSPLIT)(3)), CStr(Split(strInfo, STRSPLIT)(4)))
'            If rsClients.RecordCount > 1 Or rsClients.EOF Then
'                '��������ڶ��,����ܴ���IP��ͻ�����,��˲����ж���Ҫ������ص�վ��.ֻ�ܵ����µ�վ���ϴ�
'                strRowID = ""
'            Else '��ʾ������ص���Ϣ
'                strRowID = zlCommFun.NVL(rsClients!Id)
'            End If
'        ElseIf rsClients.RecordCount = 1 Then   '��ʾ������ص���Ϣ
'               strRowID = zlCommFun.NVL(rsClients!Id)
'        Else '��ʾ��Ҫ������ص�վ����Ϣ
'            strRowID = ""
'        End If
'    Else  '��ʾ������ص���Ϣ
'        strRowID = zlCommFun.NVL(rsClients!Id)
'    End If
'    int��������� = 0
'
'    If strRowID = "" Then
'        '��Ҫ������ص���Ϣ
'        '��û�иù���վ�����ݣ��ϴ���IP����������CPU���ڴ桢Ӳ�̡�����ϵͳ��
'        '���˺�:2010-04-27 10:13:17:bug:29279
'        strSQL = "select 1 from zlfilesupgrade   where rownum <=1"
'        Set rsTemp = zlDataBase.OpenSQLRecord(strSQL, "����Ƿ���������ļ�����")
'        int������־ = IIf(rsTemp.EOF, 0, 1)
'
'        If int������־ = 1 Then
'            '30622:Ҫ����Ƿ���������������Ƿ�����
'            'ף��:2010-12-24 10:00:00���һ�ַ�ʽFTP
'            strSQL = "select ���� from zlreginfo where ��Ŀ='��������'"
'            Set rsTemp = zlDataBase.OpenSQLRecord(strSQL, "���ʹ�õ���������")
'            If rsTemp.EOF = False Then
'                If zlCommFun.NVL(rsTemp!����, 0) = 0 Then
'                    bln������ʽ = False '�ļ�����
'                Else
'                    bln������ʽ = True  'FTP��ʽ
'                End If
'            End If
'
'            If bln������ʽ = False Then
'                strSQL = "select replace(��Ŀ,'������Ŀ¼','') as ������ from zlreginfo where ��Ŀ like '������Ŀ¼%' and ���� is not null"
'                Set rsTemp = zlDataBase.OpenSQLRecord(strSQL, "����Ƿ�����õ��ļ����������")
'                If rsTemp.EOF Then
'                    int������־ = 0
'                End If
'                int��������� = Val("" & rsTemp!������)
'            Else
'                strSQL = "select replace(��Ŀ,'FTP������','') as FTP������ from zlreginfo where ��Ŀ like 'FTP������%' and ���� is not null"
'                Set rsTemp = zlDataBase.OpenSQLRecord(strSQL, "����Ƿ�����õ�FTP������")
'                If rsTemp.EOF Then
'                    int������־ = 0
'                End If
'                int��������� = Val("" & rsTemp!������)
'            End If
'        End If
'
'        strSQL = " Insert into zlClients" & _
'                 " (IP,����վ,CPU,�ڴ�,Ӳ��,����ϵͳ,����,����������,������־,վ��)" & _
'                 " Values " & _
'                 "('" & strIpAddress & "','" & strComputerName & _
'                 "','" & Split(strInfo, STRSPLIT)(2) & "','" & Split(strInfo, STRSPLIT)(3) & _
'                 "','" & Split(strInfo, STRSPLIT)(4) & "','" & Split(strInfo, STRSPLIT)(5) & _
'                 "','" & gstrDeptName & "'," & int��������� & "," & int������־ & _
'                 ",'" & strվ���� & "')"
'        gcnOracle.Execute strSQL
'
'        If int������־ = 1 Then
'            blnAllow = True: int������ = 0: blnUpdate = True
'            GoTo AutoUpGrude:      'ִ����������
'        End If
'        �Ƿ�����ʹ�ñ�����վ = True
'        Exit Function
'    End If
'
'    With rsClients
'        blnAllow = IIf(IIf(IsNull(!����), 0, !����) = 0, True, False)
'        int������ = IIf(IsNull(!������), 0, !������) '0-��ʾ������
'        blnUpdate = IIf(IIf(IsNull(!����), 0, !����) = 1, True, False)
'        If Not blnUpdate Then blnUpdate = (IIf(IsNull(!�ռ�), 0, !�ռ�) = 1)
'    End With
'    '��Ҫ������ص�վ����Ϣ
'    strSQL = "" & _
'    "   Update zlClients " & _
'    "   set IP='" & strIpAddress & "'," & _
'    "       ����վ='" & strComputerName & "'," & _
'    "       CPU=decode(CPU,NULL,'" & Split(strInfo, STRSPLIT)(2) & "'" & ",CPU)," & _
'    "       �ڴ�=decode(�ڴ�,NULL,'" & Split(strInfo, STRSPLIT)(3) & "'" & ",�ڴ�)," & _
'    "       Ӳ��=decode(Ӳ��,NULL,'" & Split(strInfo, STRSPLIT)(4) & "'" & ",Ӳ��)," & _
'    "       ����ϵͳ=decode(����ϵͳ,NULL,'" & Split(strInfo, STRSPLIT)(5) & "'" & ",����ϵͳ)," & _
'    "       ����='" & gstrDeptName & "'," & _
'    "       վ��='" & strվ���� & "' " & _
'    "   Where RowID='" & strRowID & "'"
'    gcnOracle.Execute strSQL
'
'    If Not blnAllow Then
'        MsgBox "�ù���վ�ѱ�����Ա���ã�", vbInformation, gstrSysName
'        Exit Function
'    End If
'
'    '�������������
'    If int������ > 0 Then
'        strSQL = "Select SID From v$Session Where Upper(PROGRAM) Like 'ZLHIS%.EXE' And Status<>'KILLED' And MACHINE=(Select MACHINE From v$Session Where AUDSID=UserENV('SessionID'))"
'        If rsClients.State = 1 Then rsClients.Close
'        rsClients.Open strSQL, gcnOracle, adOpenKeyset
'        If rsClients.RecordCount > int������ Then
'            MsgBox "��ǰ����վ���ֻ���� " & int������ & " ����¼���ӣ���ǰ�Ѿ��� " & rsClients.RecordCount - 1 & " �����ӡ�", vbInformation, gstrSysName
'            Exit Function
'        End If
'    End If
'
'    On Error GoTo errHand
'    '���������Ҫ���µı�������������±���ע���
'    'If Not RegRestoreByManager Then Exit Function
'
'    '�����Ҫ�������������ǳ���
'AutoUpGrude:      'ִ����������
'    If blnUpdate Then
'        '�ж��Ƿ������˶�ʱ����
'        strSQL = "Select ���� From zlRegInfo Where ��Ŀ='�ͻ�����������'"
'        Set rsTemp = zlDataBase.OpenSQLRecord(strSQL, "��鶨ʱ����")
'        If rsTemp.RecordCount = 1 Then
'            '�����˶�ʱ����
'            If zlCommFun.NVL(rsTemp!����) <> "" Then
'                '�����������ʱ��Ƚ�
'                strCurrDate = zlDataBase.Currentdate
'                If CDate(Format(strCurrDate, "yyyy-MM-dd")) >= CDate(Format(zlCommFun.NVL(rsTemp!����), "yyyy-MM-dd")) Then
'                    strExeName = "OfficialUpgrade"
'                    blnAllow = StartHisCrust(str��������, strExeName)
'                Else
'                    blnAllow = True
'                End If
'            Else
'                blnAllow = StartHisCrust(str��������, strExeName)
'            End If
'        Else
'            blnAllow = StartHisCrust(str��������, strExeName)
'        End If
'    End If
'
'
'    '��ʱ����zlhisCrust.exe ���2011-01-11�汾����
'    '-----------------------------------------------------------------------------------------
'    Dim strSourceFile As String
'    Dim strSourceDate As String
'    Dim strTargetFile As String
'    Dim str��������   As String
'    Dim objFile As New FileSystemObject
'    If IsSourceCode Then
'        strSourceFile = "C:\APPSOFT\zlHisCrust.exe"
'    Else
'        strSourceFile = App.Path & "\zlHisCrust.exe"
'    End If
'    If objFile.FileExists(strSourceFile) Then
'        strSourceDate = Format(FileDateTime(strSourceFile), "yyyy-MM-DD hh:mm:ss")
'    Else
'        strSourceDate = "2011-01-11"
'    End If
'    If Format(strSourceDate, "YYYY-MM-DD") = "2011-01-11" Then
'
'        Dim rsTmp As New ADODB.Recordset
'        Dim strServerPath As String
'        Dim strVisitUser As String
'        Dim strVisitPassWord As String
'        Dim str�ռ����� As String
'
'        strSQL = "select ���������� from zlclients where upper(����վ)=upper(SYS_CONTEXT('USERENV','TERMINAL'))"
'        Set rsTmp = zlDataBase.OpenSQLRecord(strSQL, "��ȡ������������")
'        If rsTmp.EOF = False Then
'            If IsNull(rsTmp!����������) Then
'                str�������� = "0"
'            Else
'                str�������� = rsTmp!����������
'            End If
'        End If
'
'        If str�������� <> "" Then
'            strSQL = "Select ��Ŀ,���� From zlregInfo where ��Ŀ in('������Ŀ¼" & str�������� & "','�����û�" & str�������� & "','��������" & str�������� & "')"
'            Set rsTmp = zlDataBase.OpenSQLRecord(strSQL, "��ȡ������������Ϣ")
'            With rsTmp
'                Do While Not .EOF
'                    If !��Ŀ = "������Ŀ¼" & str�������� Then
'                        strServerPath = IIf(IsNull(!����), "", !����)
'                    End If
'                    If !��Ŀ = "�����û�" & str�������� Then
'                        strVisitUser = IIf(IsNull(!����), "", !����)
'                    End If
'                    If !��Ŀ = "��������" & str�������� Then
'                        strVisitPassWord = IIf(IsNull(!����), "", !����)
'                    End If
'                    If !��Ŀ = "�ռ�����" & str�������� Then
'                        str�ռ����� = IIf(IsNull(!����), "", !����)
'                    End If
'                    .MoveNext
'                Loop
'            End With
'
'            If IsNetServer(strServerPath, strVisitUser, strVisitPassWord) Then
'              '���ӳɹ�!
'               strTargetFile = strServerPath & "\zlHisCrust.exe"
'               On Error Resume Next
'               'ǿ�ƿ��������أ����ܳɹ����
'               objFile.CopyFile strTargetFile, strSourceFile, True
'            End If
'        End If
'    End If
'    '-----------------------------------------------------------------------------------------
'
'    �Ƿ�����ʹ�ñ�����վ = blnAllow
'    Exit Function
'errHand:
'    If zlComLib.ErrCenter = 1 Then
'        Resume
'    End If
'End Function
'
'Private Function StartHisCrust(ByVal str�������� As String, ByVal strExeName As String) As Boolean
'    Dim strPath As String
'    Err = 0: On Error Resume Next
'    strPath = App.Path ' objFileSys.GetParentFolderName(App.Path)
'    '2010-12-14 ��������в�������by�¶�
'    Error = Shell(strPath & "\" & str�������� & " " & gcnOracle.ConnectionString & "||0||" & strExeName & "||" & CStr(Command()), vbNormalFocus)
'    '������ǳ���
'    If Error = 0 Then
'        MsgBox "û���ҵ��ͻ����Զ��������ߣ�����ϵͳ����Ա��ϵ��", vbExclamation, gstrSysName
'        StartHisCrust = True
'    Else
'        StartHisCrust = False
'    End If
'End Function
'
'Private Function AnalyseComputer() As String
'    Dim strComputer As String * 256
'    Call GetComputerName(strComputer, 255)
'    AnalyseComputer = strComputer
'    AnalyseComputer = Trim(Replace(AnalyseComputer, Chr(0), ""))
'End Function
'
'Private Function AnalyseConfigure() As String
'    '��д��:���� 2003-03-09
'    '����:���������������ã�IP����������CPU���ڴ桢Ӳ�̡�����ϵͳ��
'    Dim strCPU As String           'CPU
'    Dim strMemory As String        '�ڴ�
'    Dim strOS As String            '����ϵͳ
'    Dim strComputerName As String  '�������
'    Dim strHD As String            'Ӳ��
'    Dim strIp As String            'IP��ַ
'    Dim verinfo As OSVERSIONINFO
'    Dim sysinfo As SYSTEM_INFO
'    Dim memsts As MEMORYSTATUS
'    Dim memory&
'
'    strIp = AnalyseIP
'
'    '��ȡ�������
'    strComputerName = AnalyseComputer
'
'    '��ȡӲ����Ϣ
'    strHD = AnalyseHardDisk
'
'    ' ��ò���ϵͳ��Ϣ
'    strOS = GetVersionInfo
'
'    ' ���CPU����
'    GetSystemInfo sysinfo
'    Select Case sysinfo.dwProcessorType
'    Case PROCESSOR_INTEL_386
'        strCPU = "Intel 386"
'    Case PROCESSOR_INTEL_486
'        strCPU = "Intel 486"
'    Case PROCESSOR_INTEL_PENTIUM
'        strCPU = "Intel Pentium"
'    Case PROCESSOR_MIPS_R4000
'        strCPU = "MIPS R4000"
'    Case PROCESSOR_ALPHA_21064
'        strCPU = "DEC Alpha 21064"
'    Case Else
'        strCPU = "(unknown)"
'    End Select
'
'    ' ���ʣ���ڴ�
'    GlobalMemoryStatus memsts
'    memory& = memsts.dwTotalPhys
'    strMemory = Format$(memory& \ 1024 \ 1024, "###,###,###") + "M"
'    'strMemory = "Total Physical Memory: "
'    'strMemory = strMemory + Format$(memory& \ 1024, "###,###,###") + "K"
''    memory& = memsts.dwAvailPhys
''    strMemory = strMemory + "Available Physical Memory: "
''    strMemory = strMemory + Format$(memory& \ 1024, "###,###,###") + "K"
''    memory& = memsts.dwTotalVirtual
''    strMemory = strMemory + "Total Virtual Memory: "
''    strMemory = strMemory + Format$(memory& \ 1024, "###,###,###") + "K"
''    memory& = memsts.dwAvailVirtual
''    strMemory = strMemory + "Available Virtual Memory: "
''    strMemory = strMemory + Format$(memory& \ 1024, "###,###,###") + "K"
'
'    AnalyseConfigure = strIp & STRSPLIT & strComputerName & STRSPLIT & strCPU & _
'                       STRSPLIT & strMemory & STRSPLIT & strHD & STRSPLIT & strOS
'End Function
'
'Private Function AnalyseHardDisk() As String
'    '��д��:���� 2003-03-09
'    '����:��ȡӲ��������
'    Dim lngSec As Long, lngByte As Long, lngFree As Long, lngClus As Long
'    Dim strDrive As String, dblSum As Double
'
'    strDrive = "C"
'    Do Until strDrive > "Z"
'        If GetDriveType(strDrive & ":\") = DRIVE_FIXED Then
'            If GetDiskFreeSpace(strDrive & ":\", lngSec, lngByte, lngFree, lngClus) <> 0 Then
'                dblSum = dblSum + lngSec * lngByte * CDbl(lngClus)
'            End If
'        End If
'
'        strDrive = Chr(Asc(strDrive) + 1)
'    Loop
'    AnalyseHardDisk = Format(dblSum / 1024 / 1024 / 1024, "0.00") & "G"
'End Function
'
'Private Function zl_Ip_Address_FromOrc(Optional strDefaultIp_Address As String = "") As String
'    '-----------------------------------------------------------------------------------------------------------
'    '����:ͨ��oracle��ȡ�ļ������IP��ַ
'    '���:strDefaultIp_Address-ȱʡIP��ַ
'    '����:
'    '����:����IP��ַ
'    '����:���˺�
'    '����:2009-01-21 11:08:47
'    '-----------------------------------------------------------------------------------------------------------
'    Dim rsTemp As ADODB.Recordset, strIp_Address As String, strSQL As String
'    Err = 0: On Error GoTo errHand:
'     strSQL = "Select Sys_Context('USERENV', 'IP_ADDRESS') as Ip_Address From Dual"
'    Set rsTemp = zlDataBase.OpenSQLRecord(strSQL, "��ȡIP��ַ")
'    If rsTemp.EOF = False Then
'        strIp_Address = zlCommFun.NVL(rsTemp!Ip_Address)
'    End If
'    If strIp_Address = "" Then strIp_Address = strDefaultIp_Address
'    If Replace(strIp_Address, " ", "") = "0.0.0.0" Then strIp_Address = ""
'    zl_Ip_Address_FromOrc = strIp_Address
'    Exit Function
'errHand:
'    If zlComLib.ErrCenter = 1 Then Resume
'End Function
'
'Private Function AnalyseIP() As String
'    Dim Ret As Long, Tel As Long
'    Dim bBytes() As Byte
'    Dim TempList() As String
'    Dim TempIP As String
'    Dim Tempi As Long
'    Dim Listing As MIB_IPADDRTABLE
'    Dim L3 As String
'
'
'    On Error GoTo END1
'        GetIpAddrTable ByVal 0&, Ret, True
'
'
'        If Ret <= 0 Then Exit Function
'        ReDim bBytes(0 To Ret - 1) As Byte
'        ReDim TempList(0 To Ret - 1) As String
'
'        'retrieve the data
'        GetIpAddrTable bBytes(0), Ret, False
'
'        'Get the first 4 bytes to get the entry's.. ip installed
'        CopyMemory Listing.dEntrys, bBytes(0), 4
'
'        For Tel = 0 To Listing.dEntrys - 1
'            'Copy whole structure to Listing..
'            CopyMemory Listing.mIPInfo(Tel), bBytes(4 + (Tel * Len(Listing.mIPInfo(0)))), Len(Listing.mIPInfo(Tel))
'            TempList(Tel) = ConvertAddressToString(Listing.mIPInfo(Tel).dwAddr)
'        Next Tel
'        'Sort Out The IP For WAN
'            TempIP = TempList(0)
'            For Tempi = 0 To Listing.dEntrys - 1
'                L3 = Left(TempList(Tempi), 3)
'                If L3 <> "169" And L3 <> "127" And L3 <> "192" Then
'                    TempIP = TempList(Tempi)
'                End If
'            Next Tempi
'            AnalyseIP = TempIP 'Return The TempIP
'
'
'    Exit Function
'END1:
'    AnalyseIP = ""
'End Function
'
'Private Function GetVersionInfo() As String
'    Dim myOS As OSVERSIONINFOEX
'    Dim bExInfo As Boolean
'    Dim sOS As String
'
'    '�����Windows2000�����°汾��������API��ȡһ��
'    myOS.dwOSVersionInfoSize = Len(myOS) 'should be 148/156
'    'try win2000 version
'    If GetVersionEx(myOS) = 0 Then
'        'if fails
'        myOS.dwOSVersionInfoSize = 148 'ignore reserved data
'        If GetVersionEx(myOS) = 0 Then
'            GetVersionInfo = "Windows (Unknown)"
'            Exit Function
'        End If
'    Else
'        bExInfo = True
'    End If
'
'    With myOS
'        'is version 4
'        If .dwPlatformId = VER_PLATFORM_WIN32_NT Then
'            'nt platform
'            Select Case .dwMajorVersion
'            Case 3, 4
'                sOS = "Windows NT"
'            Case 5
'                sOS = "Windows 2000"
'            End Select
'            If bExInfo Then
'                'workstation/server?
'                If .wProductType = VER_NT_SERVER Then
'                    sOS = sOS & " Server"
'                ElseIf .wProductType = VER_NT_DOMAIN_CONTROLLER Then
'                    sOS = sOS & " Domain Controller"
'                ElseIf .wProductType = VER_NT_WORKSTATION Then
'                    sOS = sOS & IIf(.dwMajorVersion >= 5, " Professional", " WorkStation")
'                End If
'            End If
'
'            'get version/build no
'            'sOS = sOS & " Version " & .dwMajorVersion & "." & .dwMinorVersion & " " & TrimNull(.szCSDVersion) & " (Build " & .dwBuildNumber & ")"
'
'        ElseIf .dwPlatformId = VER_PLATFORM_WIN32_WINDOWS Then
'            'get minor version info
'            If .dwMinorVersion = 0 Then
'                sOS = "Windows 95"
'            ElseIf .dwMinorVersion = 10 Then
'                sOS = "Windows 98"
'            ElseIf .dwMinorVersion = 90 Then
'                sOS = "Windows Millenium"
'            Else
'                sOS = "Windows 9?"
'            End If
'            'get version/build no
'            'sOS = sOS & "Version " & .dwMajorVersion & "." & .dwMinorVersion & " " & TrimNull(.szCSDVersion) & " (Build " & .dwBuildNumber & ")"
'        End If
'    End With
'    GetVersionInfo = sOS
'End Function
'
'Private Function ConvertAddressToString(longAddr As Long) As String
'    Dim myByte(3) As Byte
'    Dim Cnt As Long
'    CopyMemory myByte(0), longAddr, 4
'    For Cnt = 0 To 3
'        ConvertAddressToString = ConvertAddressToString + CStr(myByte(Cnt)) + "."
'    Next Cnt
'    ConvertAddressToString = Left$(ConvertAddressToString, Len(ConvertAddressToString) - 1)
'End Function
'
'Private Function CheckRepeatLogin() As Boolean
'    '����Ƿ����ظ���¼
'    Dim rsTemp As ADODB.Recordset
'    Dim strSQL As String
'    Dim strProgram As String
'    On Error GoTo errHand
'
'    strProgram = App.EXEName & ".exe"
'    strSQL = "Select A.UserName, A.Program, B.IP" & vbNewLine & _
'            "From gv$Session A, zlClients B" & vbNewLine & _
'            "Where A.Terminal = B.����վ" & vbNewLine & _
'            "      And A.Terminal = (Select Terminal From v$Session Where AudsID = Userenv('SessionID') and RowNum =1)" & vbNewLine & _
'            "      And A.Program =[1] And A.AudsID <> Userenv('SessionID')" & vbNewLine & _
'            "      And B.IP <> Sys_Context('USERENV', 'IP_ADDRESS')"
'
''    strSQL = "select  distinct(a.PROCESS),b.����վ,b.IP from v$session a,zlClients b where (substr(a.MACHINE,instr(a.MACHINE,'\')+1)) = b.����վ and a.USERNAME=[1] and a.PROGRAM =[2]"
'    Set rsTemp = zlDataBase.OpenSQLRecord(strSQL, "����ظ�����վ", strProgram)
'    If rsTemp.RecordCount = 0 Then '���Ե�¼
'        CheckRepeatLogin = False
'        Exit Function
'    Else
'        MsgBox "�������д�����ͬ���Ƶļ������¼," & vbCrLf & "�Է�IP��:[" & zlCommFun.NVL(rsTemp!IP) & "]", vbInformation, gstrSysName
'        CheckRepeatLogin = True
'        Exit Function
'    End If
'    Exit Function
'errHand:
'    If zlComLib.ErrCenter = 1 Then Resume
'End Function
'
'
''���20110111�汾����������������
'
'Private Function IsSourceCode() As Boolean
'    '-----------------------------------------------------------------------------------------
'    '����:ȷ���Ƿ�Դ����
'    '����:��ԭ����-true,����Դ����-false
'    '-----------------------------------------------------------------------------------------
'    Err = 0: On Error Resume Next
'    Debug.Print 1 / 0
'    IsSourceCode = Err <> 0
'End Function
'
'Public Function IsNetServer(ByVal gstrServerPath As String, ByVal gstrVisitUser As String, ByVal gstrVisitPassWord As String) As Boolean
'    '----------------------------------------------------------------------------------------------------------
'    '--����:���������Ƿ�����������
'    '----------------------------------------------------------------------------------------------------------
'    Dim NetR As NETRESOURCE
'    Dim objFile As New FileSystemObject
'
'    '���˺�:���ܴ���windows��Դ�������Ѿ��з��ʵ���
'    '
'    If objFile.FolderExists(gstrServerPath) Then
'            IsNetServer = True: Exit Function
'    End If
'
'    If objFile.FolderExists(gstrServerPath) Then '���ڴ��ļ���,�϶�û��Ȩ�޷���,��Ҫɾ������
'            Call zlNetCancelConnected 'Ŀǰȫ��ɱ��,ԭ���ǲ�֪���ļ���������:��:IP�ͻ���������
'    End If
'
'
'    With NetR
'        .dwScope = RESOURCE_GLOBALNET
'        .dwType = RESOURCETYPE_DISK
'        .dwDisplayType = RESOURCEDISPLAYTYPE_SHARE
'        .dwUsage = RESOURCEUSAGE_CONNECTABLE
'        .lpLocalName = "" 'ӳ���������
'        .lpRemoteName = gstrServerPath  '������·��
'    End With
'
'    Err = 0
'    On Error GoTo errHand:
'    If WNetAddConnection2(NetR, gstrVisitPassWord, gstrVisitUser, CONNECT_UPDATE_PROFILE) = NO_ERROR Then
'       IsNetServer = True
'    Else
'       IsNetServer = False
'    End If
'    Exit Function
'errHand:
'       IsNetServer = False
'End Function
'
'Public Function CancelNetServer(Optional strName As String, Optional strServerPath As String) As Boolean
'    '�Ͽ�����������
'    Dim lngReturn As Long
'
'    Err = 0
'    On Error Resume Next
'    lngReturn = WNetCancelConnection2(IIf(strName = "", strServerPath, strName), CONNECT_UPDATE_PROFILE, True)
'    If lngReturn = 0 Then
'        CancelNetServer = True
'    Else
'        CancelNetServer = False
'    End If
'    Err = 0
'End Function
