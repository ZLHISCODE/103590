Attribute VB_Name = "mdlHisCrust"
'Option Explicit
'
''分析本机配置相关API
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
''取IP的API
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
'  ComboIndex            As Long  '保留
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
''取硬盘大小
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
''下列语句用于导出注册表中指定分支到一个文件中
''regedit /e d:\Win2000.reg "HKEY_CURRENT_USER\SOFTWARE\VB AND VBA PROGRAM SETTIGS\ZLSOFT"
''----------------------------------------------------------------------------------------------------
'Private Const STRSPLIT As String = "♂♂"
'Private Const REGCMD As String = "REGEDIT /E"
'Private Const RegFile As String = "C:\REGFILE.REG"
'Private Const REGDATA As String = "C:\REGDATA.REG"
'Private Const REGDIRECTORY As String = """HKEY_CURRENT_USER\SOFTWARE\VB AND VBA PROGRAM SETTINGS\ZLSOFT"""
'
''下列语句用于检测是否合法调用
'Private Declare Function GlobalAddAtom Lib "kernel32" Alias "GlobalAddAtomA" (ByVal lpString As String) As Integer
'Private Declare Function GlobalDeleteAtom Lib "kernel32" (ByVal nAtom As Integer) As Integer
'Private Declare Function GlobalGetAtomName Lib "kernel32" Alias "GlobalGetAtomNameA" (ByVal nAtom As Integer, ByVal lpBuffer As String, ByVal nSize As Long) As Long
'
'
''以下为刘兴宏加入:主要是选择文件等
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
'''''''解决20110111''''''''不能升级问题
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
'''''''解决20110111''''''''不能升级问题
'
'
'Public Function 是否允许使用本工作站(Optional ByRef strIpAddress As String) As Boolean
'    '-----------------------------------------------------------------------------------------------------------
'    '功能:检查是否允许使用本工作站及站点信息的上传
'    '     判断是否允许该工作站使用程序；如果需要替换本地参数，则执行替换操作；如果需要升级，则调用外壳程序，并关闭退出
'    '入参:
'    '出参:
'    '返回:成功,返回true,否则返回False
'    '编制:刘兴洪
'    '日期:2009-01-21 11:59:49
'    '-----------------------------------------------------------------------------------------------------------
'    Dim objFileSys As New FileSystemObject, rsClients As New ADODB.Recordset, rsTemp As ADODB.Recordset
'    Dim strSQL As String, strInfo As String, strCurrDate As String, strExeName As String
'    Dim str升级程序 As String, Error As Long, strComputerName As String, strRowID As String 'IP地址及站点名
'    Dim blnAllow As Boolean, blnUpdate As Boolean, int连接数 As Integer, int升级标志 As Integer
'    Dim int服务器编号 As Integer, i As Integer
'    Dim str站点       As String, strSouce站点 As String, str站点编号 As String, str名称 As String, str缺省 As String, str缺省部门
'    Dim strSplit站点()  As String, bln站点 As Boolean, bln升级方式 As Boolean, bln检查站点 As Boolean, strCurIndex As String
'    Dim lng有站点 As Long, bln空站点 As Boolean
'    Err = 0: On Error Resume Next
'
'    strIpAddress = ""
'    blnAllow = False: blnUpdate = False: str升级程序 = "zlHisCrust.exe": 是否允许使用本工作站 = False
'    strExeName = GetSetting("ZLSOFT", "公共全局", "执行文件", "")
'
'    '判断是否允许使用
'    strComputerName = AnalyseComputer           '分析机算机名
'    strInfo = AnalyseConfigure: strIpAddress = Split(strInfo, STRSPLIT)(0)
'    strIpAddress = zl_Ip_Address_FromOrc(strIpAddress)  '以oracle连接的IP地址为主
'
'    ''''''ZQ20101109''''''''''
'    ''''''检查是否有重名机器''''''
'    If CheckRepeatLogin() = True Then
'        是否允许使用本工作站 = False
'        Exit Function
'    End If
'
'
''''    '检查客户端站点是否为NULL
''''    '祝庆:2010-12-24 10:00:00
''''    strSQL = "select 站点 from zlclients where IP= Sys_Context('USERENV', 'IP_ADDRESS') and 工作站= SYS_CONTEXT('USERENV','TERMINAL')"
''''    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "检查客户站点是否确定")
''''    If rsTemp.RecordCount = 1 Then
''''        bln检查站点 = IIf(zlCommFun.NVL(rsTemp!站点) = "", True, False)
''''        str站点编号 = zlCommFun.NVL(rsTemp!站点)
''''    Else
''''        bln检查站点 = True
''''        str站点编号 = ""
''''    End If
'    '升级后登陆,不在让用户选择,直接读取 ZQ20110114
'    '如果含有/，表示是用户加密码的格式，如：zlhis/his
'    If InStrRev(Command(), "/", -1) > 0 Then
'         bln检查站点 = False
'         strSQL = "select 站点,部门 from zlclients where IP= Sys_Context('USERENV', 'IP_ADDRESS') and 工作站= SYS_CONTEXT('USERENV','TERMINAL')"
'         Set rsTemp = zlDataBase.OpenSQLRecord(strSQL, "检查客户站点是否确定")
'         If rsTemp.RecordCount = 1 Then
'            str站点编号 = zlCommFun.NVL(rsTemp!站点)
'            gstrDeptName = zlCommFun.NVL(rsTemp!部门)
'         Else
'            str站点编号 = ""
'            gstrDeptName = ""
'         End If
'    Else
'         bln检查站点 = True
'    End If
'
'
'    If bln检查站点 Then
'        strSQL = "select C.名称,C.站点,B.缺省 from 上机人员表 A,部门人员 B, 部门表 C where A.人员ID = B.人员ID And B.部门ID = C.ID And upper(A.用户名)=upper([1]) order by C.名称"
'        Set rsTemp = zlDataBase.OpenSQLRecord(strSQL, "检查并确定所属院区", gstrDbUser)
'        Do While Not rsTemp.EOF
'            If str站点 = "" Then
'                If zlCommFun.NVL(rsTemp!站点, "") <> "" Then
'                    str站点 = zlCommFun.NVL(rsTemp!站点, "") & ","
'                    str名称 = zlCommFun.NVL(rsTemp!名称) & ","
'                    lng有站点 = lng有站点 + 1
'                Else
'                    bln空站点 = True
'                End If
'            Else
'                If zlCommFun.NVL(rsTemp!站点, "") <> "" Then
'                    str站点 = str站点 & zlCommFun.NVL(rsTemp!站点, "") & ","
'                    str名称 = str名称 & zlCommFun.NVL(rsTemp!名称) & ","
'                    lng有站点 = lng有站点 + 1
'                Else
'                    bln空站点 = True
'                End If
'            End If
'            If zlCommFun.NVL(rsTemp!缺省, "0") = 1 Then
'                str缺省部门 = zlCommFun.NVL(rsTemp!名称)
'            End If
'            rsTemp.MoveNext
'        Loop
'
'        '  str站点 = ""如果当前登录人员所属部门都没有设置站点，则不作处理。在查找该院是否启动了站点控制!
'        If str站点 = "" Or (bln空站点 And lng有站点 <> 1) Then
'            str站点 = ""
'            strSQL = "select distinct (A.站点),B.名称 from 部门表 A,zlNodeList B where A.站点=B.编号 And A.站点 is not null order by A.站点"
'            Set rsTemp = zlDataBase.OpenSQLRecord(strSQL, "检查是否启动站点控制")
'            Do While Not rsTemp.EOF
'                If str站点 = "" Then
'                    If zlCommFun.NVL(rsTemp!站点, "") <> "" Then
'                        str站点 = zlCommFun.NVL(rsTemp!站点, "") & ","
'                        str名称 = zlCommFun.NVL(rsTemp!名称) & ","
'                    End If
'                Else
'                    If zlCommFun.NVL(rsTemp!站点, "") <> "" Then
'                        str站点 = str站点 & zlCommFun.NVL(rsTemp!站点, "") & ","
'                        str名称 = str名称 & zlCommFun.NVL(rsTemp!名称) & ","
'                    End If
'                End If
'                rsTemp.MoveNext
'            Loop
'        End If
'
'        If str站点 <> "" Then
'            strSplit站点 = Split(str站点, ",")
'            For i = 0 To UBound(strSplit站点) - 1
'                If i = 0 Then
'                    strSouce站点 = strSplit站点(i)
'                Else
'                    If strSouce站点 <> strSplit站点(i) Then
'                        bln站点 = True
'                        Exit For
'                    End If
'                End If
'            Next
'
'            If bln站点 Then
'            'bln站点 = True 当前登录人员所属部门包含多个站点,提示用户选择当前计算机位置所在的部门。
'                strCurIndex = GetRegister(私有模块, App.EXEName, "当前站点选择", "")
'                Call frmSelClient.ShowEdit(str站点, str名称, strCurIndex)
'                str站点编号 = IIf(frmSelClient.gstr站点 = "无", "", frmSelClient.gstr站点)
'                gstrDeptName = IIf(bln站点, str缺省部门, frmSelClient.gstrCur站点)
'                Call SetRegister(私有模块, App.EXEName, "当前站点选择", str站点编号)
'            Else
'            'bln站点= False 当前登录人员所属部门都是相同的站点，则以该站点编号保存在"zlClients.站点"中。
'                str站点编号 = strSouce站点
'                gstrDeptName = str缺省部门
'            End If
'
'        End If
'    End If
'    If str站点编号 <> "" Then
'        zlComLib.gstrNodeNo = str站点编号
'    Else
'        zlComLib.gstrNodeNo = "-"
'        gstrDeptName = str缺省部门
'    End If
'
'    '分两部进行检查.问题:15640
'    '1.以站点名检查
'    strSQL = "Select Rowid as ID,Nvl(禁止使用,0) as 允许,Nvl(升级标志,0) as 升级,Nvl(收集标志,0) as 收集,连接数 From zlClients Where 工作站=[1]"
'    Set rsClients = zlDataBase.OpenSQLRecord(strSQL, "检查工作站-以站点为主", strComputerName)
'    If rsClients.EOF Then
'        '2.未发现此站点,则以IP方式查找，但只有一个时才更新计算名
'        strSQL = "Select Rowid as ID, Nvl(禁止使用,0) as 允许,Nvl(升级标志,0) as 升级,Nvl(收集标志,0) as 收集,连接数 From zlClients Where IP=[1]"
'        Set rsClients = zlDataBase.OpenSQLRecord(strSQL, "检查工作站-以站点为主", strIpAddress)
'        If rsClients.RecordCount > 1 Then
'            '大于两个以上,则加CPU,内存,硬盘为限制条件.
'            strSQL = "" & _
'                "   Select Rowid as ID, Nvl(禁止使用,0) as 允许,Nvl(升级标志,0) as 升级,Nvl(收集标志,0) as 收集,连接数 " & _
'                "   From zlClients Where IP=[1] and CPU=[2] and  内存=[3] and 硬盘=[4]"
'            Set rsClients = zlDataBase.OpenSQLRecord(strSQL, "检查工作站-以站点为主", strIpAddress, CStr(Split(strInfo, STRSPLIT)(2)), CStr(Split(strInfo, STRSPLIT)(3)), CStr(Split(strInfo, STRSPLIT)(4)))
'            If rsClients.RecordCount > 1 Or rsClients.EOF Then
'                '如果还存在多个,则可能存在IP冲突的情况,因此不能判定需要更新相关的站点.只能当成新的站点上传
'                strRowID = ""
'            Else '表示更新相关的信息
'                strRowID = zlCommFun.NVL(rsClients!Id)
'            End If
'        ElseIf rsClients.RecordCount = 1 Then   '表示更新相关的信息
'               strRowID = zlCommFun.NVL(rsClients!Id)
'        Else '表示需要新增相关的站点信息
'            strRowID = ""
'        End If
'    Else  '表示更新相关的信息
'        strRowID = zlCommFun.NVL(rsClients!Id)
'    End If
'    int服务器编号 = 0
'
'    If strRowID = "" Then
'        '需要新增相关的信息
'        '还没有该工作站的数据，上传（IP、机器名、CPU、内存、硬盘、操作系统）
'        '刘兴洪:2010-04-27 10:13:17:bug:29279
'        strSQL = "select 1 from zlfilesupgrade   where rownum <=1"
'        Set rsTemp = zlDataBase.OpenSQLRecord(strSQL, "检查是否存在升级文件配置")
'        int升级标志 = IIf(rsTemp.EOF, 0, 1)
'
'        If int升级标志 = 1 Then
'            '30622:要检查是否存在升级服务器是否配置
'            '祝庆:2010-12-24 10:00:00添加一种方式FTP
'            strSQL = "select 内容 from zlreginfo where 项目='升级类型'"
'            Set rsTemp = zlDataBase.OpenSQLRecord(strSQL, "检查使用的升级类型")
'            If rsTemp.EOF = False Then
'                If zlCommFun.NVL(rsTemp!内容, 0) = 0 Then
'                    bln升级方式 = False '文件共享
'                Else
'                    bln升级方式 = True  'FTP方式
'                End If
'            End If
'
'            If bln升级方式 = False Then
'                strSQL = "select replace(项目,'服务器目录','') as 服务器 from zlreginfo where 项目 like '服务器目录%' and 内容 is not null"
'                Set rsTemp = zlDataBase.OpenSQLRecord(strSQL, "检查是否存配置的文件共享服务器")
'                If rsTemp.EOF Then
'                    int升级标志 = 0
'                End If
'                int服务器编号 = Val("" & rsTemp!服务器)
'            Else
'                strSQL = "select replace(项目,'FTP服务器','') as FTP服务器 from zlreginfo where 项目 like 'FTP服务器%' and 内容 is not null"
'                Set rsTemp = zlDataBase.OpenSQLRecord(strSQL, "检查是否存配置的FTP服务器")
'                If rsTemp.EOF Then
'                    int升级标志 = 0
'                End If
'                int服务器编号 = Val("" & rsTemp!服务器)
'            End If
'        End If
'
'        strSQL = " Insert into zlClients" & _
'                 " (IP,工作站,CPU,内存,硬盘,操作系统,部门,升级服务器,升级标志,站点)" & _
'                 " Values " & _
'                 "('" & strIpAddress & "','" & strComputerName & _
'                 "','" & Split(strInfo, STRSPLIT)(2) & "','" & Split(strInfo, STRSPLIT)(3) & _
'                 "','" & Split(strInfo, STRSPLIT)(4) & "','" & Split(strInfo, STRSPLIT)(5) & _
'                 "','" & gstrDeptName & "'," & int服务器编号 & "," & int升级标志 & _
'                 ",'" & str站点编号 & "')"
'        gcnOracle.Execute strSQL
'
'        If int升级标志 = 1 Then
'            blnAllow = True: int连接数 = 0: blnUpdate = True
'            GoTo AutoUpGrude:      '执升升级程序
'        End If
'        是否允许使用本工作站 = True
'        Exit Function
'    End If
'
'    With rsClients
'        blnAllow = IIf(IIf(IsNull(!允许), 0, !允许) = 0, True, False)
'        int连接数 = IIf(IsNull(!连接数), 0, !连接数) '0-表示无限制
'        blnUpdate = IIf(IIf(IsNull(!升级), 0, !升级) = 1, True, False)
'        If Not blnUpdate Then blnUpdate = (IIf(IsNull(!收集), 0, !收集) = 1)
'    End With
'    '需要更新相关的站点信息
'    strSQL = "" & _
'    "   Update zlClients " & _
'    "   set IP='" & strIpAddress & "'," & _
'    "       工作站='" & strComputerName & "'," & _
'    "       CPU=decode(CPU,NULL,'" & Split(strInfo, STRSPLIT)(2) & "'" & ",CPU)," & _
'    "       内存=decode(内存,NULL,'" & Split(strInfo, STRSPLIT)(3) & "'" & ",内存)," & _
'    "       硬盘=decode(硬盘,NULL,'" & Split(strInfo, STRSPLIT)(4) & "'" & ",硬盘)," & _
'    "       操作系统=decode(操作系统,NULL,'" & Split(strInfo, STRSPLIT)(5) & "'" & ",操作系统)," & _
'    "       部门='" & gstrDeptName & "'," & _
'    "       站点='" & str站点编号 & "' " & _
'    "   Where RowID='" & strRowID & "'"
'    gcnOracle.Execute strSQL
'
'    If Not blnAllow Then
'        MsgBox "该工作站已被管理员禁用！", vbInformation, gstrSysName
'        Exit Function
'    End If
'
'    '连接数检查限制
'    If int连接数 > 0 Then
'        strSQL = "Select SID From v$Session Where Upper(PROGRAM) Like 'ZLHIS%.EXE' And Status<>'KILLED' And MACHINE=(Select MACHINE From v$Session Where AUDSID=UserENV('SessionID'))"
'        If rsClients.State = 1 Then rsClients.Close
'        rsClients.Open strSQL, gcnOracle, adOpenKeyset
'        If rsClients.RecordCount > int连接数 Then
'            MsgBox "当前工作站最多只允许 " & int连接数 & " 个登录连接，当前已经有 " & rsClients.RecordCount - 1 & " 个连接。", vbInformation, gstrSysName
'            Exit Function
'        End If
'    End If
'
'    On Error GoTo errHand
'    '如果存在需要更新的本机参数，则更新本机注册表
'    'If Not RegRestoreByManager Then Exit Function
'
'    '如果需要升级，则调用外壳程序
'AutoUpGrude:      '执升升级程序
'    If blnUpdate Then
'        '判断是否启动了定时升级
'        strSQL = "Select 内容 From zlRegInfo Where 项目='客户端升级日期'"
'        Set rsTemp = zlDataBase.OpenSQLRecord(strSQL, "检查定时升级")
'        If rsTemp.RecordCount = 1 Then
'            '启动了定时升级
'            If zlCommFun.NVL(rsTemp!内容) <> "" Then
'                '进行与服务器时间比较
'                strCurrDate = zlDataBase.Currentdate
'                If CDate(Format(strCurrDate, "yyyy-MM-dd")) >= CDate(Format(zlCommFun.NVL(rsTemp!内容), "yyyy-MM-dd")) Then
'                    strExeName = "OfficialUpgrade"
'                    blnAllow = StartHisCrust(str升级程序, strExeName)
'                Else
'                    blnAllow = True
'                End If
'            Else
'                blnAllow = StartHisCrust(str升级程序, strExeName)
'            End If
'        Else
'            blnAllow = StartHisCrust(str升级程序, strExeName)
'        End If
'    End If
'
'
'    '临时升级zlhisCrust.exe 解决2011-01-11版本问题
'    '-----------------------------------------------------------------------------------------
'    Dim strSourceFile As String
'    Dim strSourceDate As String
'    Dim strTargetFile As String
'    Dim str服务器号   As String
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
'        Dim str收集类型 As String
'
'        strSQL = "select 升级服务器 from zlclients where upper(工作站)=upper(SYS_CONTEXT('USERENV','TERMINAL'))"
'        Set rsTmp = zlDataBase.OpenSQLRecord(strSQL, "获取升级服务器号")
'        If rsTmp.EOF = False Then
'            If IsNull(rsTmp!升级服务器) Then
'                str服务器号 = "0"
'            Else
'                str服务器号 = rsTmp!升级服务器
'            End If
'        End If
'
'        If str服务器号 <> "" Then
'            strSQL = "Select 项目,内容 From zlregInfo where 项目 in('服务器目录" & str服务器号 & "','访问用户" & str服务器号 & "','访问密码" & str服务器号 & "')"
'            Set rsTmp = zlDataBase.OpenSQLRecord(strSQL, "获取升级服务器信息")
'            With rsTmp
'                Do While Not .EOF
'                    If !项目 = "服务器目录" & str服务器号 Then
'                        strServerPath = IIf(IsNull(!内容), "", !内容)
'                    End If
'                    If !项目 = "访问用户" & str服务器号 Then
'                        strVisitUser = IIf(IsNull(!内容), "", !内容)
'                    End If
'                    If !项目 = "访问密码" & str服务器号 Then
'                        strVisitPassWord = IIf(IsNull(!内容), "", !内容)
'                    End If
'                    If !项目 = "收集类型" & str服务器号 Then
'                        str收集类型 = IIf(IsNull(!内容), "", !内容)
'                    End If
'                    .MoveNext
'                Loop
'            End With
'
'            If IsNetServer(strServerPath, strVisitUser, strVisitPassWord) Then
'              '连接成功!
'               strTargetFile = strServerPath & "\zlHisCrust.exe"
'               On Error Resume Next
'               '强制拷贝到本地，不管成功与否
'               objFile.CopyFile strTargetFile, strSourceFile, True
'            End If
'        End If
'    End If
'    '-----------------------------------------------------------------------------------------
'
'    是否允许使用本工作站 = blnAllow
'    Exit Function
'errHand:
'    If zlComLib.ErrCenter = 1 Then
'        Resume
'    End If
'End Function
'
'Private Function StartHisCrust(ByVal str升级程序 As String, ByVal strExeName As String) As Boolean
'    Dim strPath As String
'    Err = 0: On Error Resume Next
'    strPath = App.Path ' objFileSys.GetParentFolderName(App.Path)
'    '2010-12-14 添加命令行参数传入by陈东
'    Error = Shell(strPath & "\" & str升级程序 & " " & gcnOracle.ConnectionString & "||0||" & strExeName & "||" & CStr(Command()), vbNormalFocus)
'    '调用外壳程序
'    If Error = 0 Then
'        MsgBox "没有找到客户端自动升级工具，请与系统管理员联系。", vbExclamation, gstrSysName
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
'    '编写人:朱玉宝 2003-03-09
'    '功能:分析出本机的配置（IP、机器名、CPU、内存、硬盘、操作系统）
'    Dim strCPU As String           'CPU
'    Dim strMemory As String        '内存
'    Dim strOS As String            '操作系统
'    Dim strComputerName As String  '计算机名
'    Dim strHD As String            '硬盘
'    Dim strIp As String            'IP地址
'    Dim verinfo As OSVERSIONINFO
'    Dim sysinfo As SYSTEM_INFO
'    Dim memsts As MEMORYSTATUS
'    Dim memory&
'
'    strIp = AnalyseIP
'
'    '获取计算机名
'    strComputerName = AnalyseComputer
'
'    '获取硬盘信息
'    strHD = AnalyseHardDisk
'
'    ' 获得操作系统信息
'    strOS = GetVersionInfo
'
'    ' 获得CPU类型
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
'    ' 获得剩余内存
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
'    '编写人:朱玉宝 2003-03-09
'    '功能:获取硬盘总容量
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
'    '功能:通过oracle获取的计算机的IP地址
'    '入参:strDefaultIp_Address-缺省IP地址
'    '出参:
'    '返回:返回IP地址
'    '编制:刘兴洪
'    '日期:2009-01-21 11:08:47
'    '-----------------------------------------------------------------------------------------------------------
'    Dim rsTemp As ADODB.Recordset, strIp_Address As String, strSQL As String
'    Err = 0: On Error GoTo errHand:
'     strSQL = "Select Sys_Context('USERENV', 'IP_ADDRESS') as Ip_Address From Dual"
'    Set rsTemp = zlDataBase.OpenSQLRecord(strSQL, "获取IP地址")
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
'    '如果是Windows2000或以下版本，则用新API再取一次
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
'    '检查是否有重复登录
'    Dim rsTemp As ADODB.Recordset
'    Dim strSQL As String
'    Dim strProgram As String
'    On Error GoTo errHand
'
'    strProgram = App.EXEName & ".exe"
'    strSQL = "Select A.UserName, A.Program, B.IP" & vbNewLine & _
'            "From gv$Session A, zlClients B" & vbNewLine & _
'            "Where A.Terminal = B.工作站" & vbNewLine & _
'            "      And A.Terminal = (Select Terminal From v$Session Where AudsID = Userenv('SessionID') and RowNum =1)" & vbNewLine & _
'            "      And A.Program =[1] And A.AudsID <> Userenv('SessionID')" & vbNewLine & _
'            "      And B.IP <> Sys_Context('USERENV', 'IP_ADDRESS')"
'
''    strSQL = "select  distinct(a.PROCESS),b.工作站,b.IP from v$session a,zlClients b where (substr(a.MACHINE,instr(a.MACHINE,'\')+1)) = b.工作站 and a.USERNAME=[1] and a.PROGRAM =[2]"
'    Set rsTemp = zlDataBase.OpenSQLRecord(strSQL, "检查重复工作站", strProgram)
'    If rsTemp.RecordCount = 0 Then '可以登录
'        CheckRepeatLogin = False
'        Exit Function
'    Else
'        MsgBox "局域网中存在相同名称的计算机登录," & vbCrLf & "对方IP是:[" & zlCommFun.NVL(rsTemp!IP) & "]", vbInformation, gstrSysName
'        CheckRepeatLogin = True
'        Exit Function
'    End If
'    Exit Function
'errHand:
'    If zlComLib.ErrCenter = 1 Then Resume
'End Function
'
'
''解决20110111版本不能升级自身问题
'
'Private Function IsSourceCode() As Boolean
'    '-----------------------------------------------------------------------------------------
'    '功能:确定是否源代码
'    '返回:是原代码-true,不是源代码-false
'    '-----------------------------------------------------------------------------------------
'    Err = 0: On Error Resume Next
'    Debug.Print 1 / 0
'    IsSourceCode = Err <> 0
'End Function
'
'Public Function IsNetServer(ByVal gstrServerPath As String, ByVal gstrVisitUser As String, ByVal gstrVisitPassWord As String) As Boolean
'    '----------------------------------------------------------------------------------------------------------
'    '--功能:检查服务器是否正常并连接
'    '----------------------------------------------------------------------------------------------------------
'    Dim NetR As NETRESOURCE
'    Dim objFile As New FileSystemObject
'
'    '刘兴洪:可能存在windows资源管理器已经有访问的了
'    '
'    If objFile.FolderExists(gstrServerPath) Then
'            IsNetServer = True: Exit Function
'    End If
'
'    If objFile.FolderExists(gstrServerPath) Then '存在此文件夹,肯定没有权限访问,则要删除连接
'            Call zlNetCancelConnected '目前全部杀死,原因是不知道文件服务器名:如:IP和机器名访问
'    End If
'
'
'    With NetR
'        .dwScope = RESOURCE_GLOBALNET
'        .dwType = RESOURCETYPE_DISK
'        .dwDisplayType = RESOURCEDISPLAYTYPE_SHARE
'        .dwUsage = RESOURCEUSAGE_CONNECTABLE
'        .lpLocalName = "" '映射的驱动器
'        .lpRemoteName = gstrServerPath  '服务器路径
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
'    '断开服务器连接
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
