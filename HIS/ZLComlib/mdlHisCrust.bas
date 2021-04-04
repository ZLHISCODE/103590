Attribute VB_Name = "mdlHisCrust"
Option Explicit

'分析本机配置相关API
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
'取硬盘大小
Private Const DRIVE_FIXED = 3
Private Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTotalNumberOfClusters As Long) As Long
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Const STRSPLIT As String = "♂♂"

'API错误信息获取
Private Declare Function GetLastError Lib "kernel32" () As Long
Private Declare Function FormatMessage Lib "kernel32.dll" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200
Private Declare Function WNetGetLastError Lib "mpr.dll" Alias "WNetGetLastErrorA" (lpError As Long, ByVal lpErrorBuf As String, ByVal nErrorBufSize As Long, ByVal lpNameBuf As String, ByVal nNameBufSize As Long) As Long
Private Const ERROR_EXTENDED_ERROR          As Long = 1208
'文件描述信息判断
Private Declare Function GetFileVersionInfoSize Lib "version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Function GetFileVersionInfo Lib "version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwHandle As Long, ByVal dwLen As Long, lpData As Any) As Long
Private Declare Function VerQueryValue Lib "version.dll" Alias "VerQueryValueA" (ByVal pBlock As Long, ByVal lpSubBlock As String, lplpBuffer As Long, puLen As Long) As Long
'Public Const FVN_Comments           As String = "Comments"          '注释
'Public Const FVN_InternalName       As String = "InternalName"      '内部名称
'Public Const FVN_ProductName        As String = "ProductName"       '产品名
'Public Const FVN_CompanyName        As String = "CompanyName"       '公司名
'Public Const FVN_ProductVersion     As String = "ProductVersion"    '产品版本
'Public Const FVN_FileDescription    As String = "FileDescription"   '文件描述
'Public Const FVN_OriginalFilename   As String = "OriginalFilename"  '原始文件名
'Public Const FVN_FileVersion        As String = "FileVersion"       '文件版本
'Public Const FVN_SpecialBuild       As String = "SpecialBuild"      '特殊编译号
'Public Const FVN_PrivateBuild       As String = "PrivateBuild"      '私有编译号
'Public Const FVN_LegalCopyright     As String = "LegalCopyright"    '合法版权
'Public Const FVN_LegalTrademarks    As String = "LegalTrademarks"   '合法商标
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)


Public Function CheckAllowByTerminal() As Boolean
'功能:检查是否允许使用本工作站,以及进行当前工作站信息的登记
'     判断是否允许该工作站使用程序；
'     如果需要替换本地参数，则执行替换操作；如果需要升级，则调用外壳程序，并关闭退出
'返回:成功,返回true,否则返回False
    
    Dim rsTmp As ADODB.Recordset, strSQL As String, strRowID As String '客户端的ROWID
    Dim strComuterInfo As String, arrComputer As Variant, strComputerName As String, strIpAddress As String
    Dim strTmp As String, arrTmp As Variant, i As Integer
    Dim bln检查站点 As Boolean, lng有站点 As Long, bln空站点 As Boolean, bln多站点 As Boolean
    Dim str站点       As String, str站点编号 As String, str名称 As String, str缺省部门
    Dim blnAllow As Boolean, blnUpdate As Boolean
    Dim int服务器编号 As Integer, int启用视频源 As Integer, int连接数 As Integer, int升级标志 As Integer
    
    Call SQLTest(App.EXEName, "mdlHisCrust", "新版电子病历自动升级检查")
    Call UpdateEmrInterface '新版电子病历自动升级
    Call SQLTest

    strIpAddress = Sys.IP '以oracle连接的IP地址为主
    strComputerName = OS.ComputerName
    '检查是否有重名机器
    If CheckRepeatLogin(strIpAddress) = True Then
        CheckAllowByTerminal = False
        Exit Function
    End If
    '判断是否允许使用
    strComuterInfo = AnalyseConfigure
    arrComputer = Split(strComuterInfo, STRSPLIT)
    '1.以站点名检查
    If Err.Number <> 0 Then Err.Clear
    On Error Resume Next
    strSQL = "Select Rowid as ID,站点,部门,Nvl(禁止使用,0) as 允许,Nvl(升级标志,0) as 升级,Nvl(收集标志,0) as 收集,连接数,启用视频源 From zlClients Where 工作站=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "检查工作站-以站点为主", strComputerName)
    '可能由于未授权等原因，导致查询出错，此时弹出提示禁止登录
    If rsTmp Is Nothing Then
        MsgBox Err.Description & vbNewLine & "不能正常访问系统，请您联系系统管理员重新进行角色授权！", vbInformation, gstrSysName
        Exit Function
    End If
    '2.未发现此站点,则以IP方式查找，但只有一个时才更新计算名
    If rsTmp.EOF Then
        strSQL = "Select Rowid as ID,站点,部门, Nvl(禁止使用,0) as 允许,Nvl(升级标志,0) as 升级,Nvl(收集标志,0) as 收集,连接数,启用视频源 From zlClients Where IP=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "检查工作站-以站点为主", strIpAddress)
        If rsTmp.RecordCount > 1 Then
            '大于两个以上,则加CPU,内存,硬盘为限制条件.
            strSQL = "" & _
                "   Select Rowid as ID,站点,部门,Nvl(禁止使用,0) as 允许,Nvl(升级标志,0) as 升级,Nvl(收集标志,0) as 收集,连接数,启用视频源 " & _
                "   From zlClients Where IP=[1] and CPU=[2] and  内存=[3] and 硬盘=[4]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "检查工作站-以站点为主", strIpAddress, CStr(arrComputer(2)), CStr(arrComputer(3)), CStr(arrComputer(4)))
        End If
    End If
    bln检查站点 = True
    '如果还存在多个,则可能存在IP冲突的情况,因此不能判定需要更新相关的站点.只能当成新的站点上传
    If rsTmp.RecordCount > 1 Or rsTmp.EOF Then
        strRowID = ""
    Else '表示更新相关的信息
        strRowID = NVL(rsTmp!id)
        int启用视频源 = Val(NVL(rsTmp!启用视频源))
        '升级后登陆,不在让用户选择,直接读取
        If Command() <> "" Then
            '新方法
            If InStr(Command(), "ZLHISCRUSTCALL=1") > 0 And InStr(Command(), "USER=") > 0 And InStr(Command(), "PASS=") > 0 Then
                bln检查站点 = False
                str站点编号 = NVL(rsTmp!站点)
                gstrDeptName = NVL(rsTmp!部门)
            '老的判断方法
            ElseIf InStrRev(Command(), "/", -1) > 0 And InStrRev(Command(), ",", -1) = 0 Then
                bln检查站点 = False
                str站点编号 = NVL(rsTmp!站点)
                gstrDeptName = NVL(rsTmp!部门)
            End If
        End If
        blnAllow = Val(rsTmp!允许 & "") = 0
        int连接数 = Val(rsTmp!连接数 & "")  '0-表示无限制
        blnUpdate = Val(rsTmp!升级 & "") = 1
        If Not blnUpdate Then blnUpdate = Val(rsTmp!收集 & "") = 1
    End If

    If bln检查站点 Then
        strSQL = "select C.名称,C.站点,B.缺省 from 上机人员表 A,部门人员 B, 部门表 C where A.人员ID = B.人员ID And B.部门ID = C.ID And A.用户名=[1] order by C.站点"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "检查并确定所属院区", gstrDbUser)
        If rsTmp Is Nothing Then
            MsgBox Err.Description & vbNewLine & "不能正常访问系统，请您联系系统管理员重新进行角色授权！", vbInformation, gstrSysName
            Exit Function
        End If
        Do While Not rsTmp.EOF
            If NVL(rsTmp!站点, "") <> "" Then
                str站点 = str站点 & "," & NVL(rsTmp!站点, "")
                str名称 = str名称 & "," & NVL(rsTmp!名称)
                lng有站点 = lng有站点 + 1
            Else
                bln空站点 = True
            End If
            If NVL(rsTmp!缺省, "0") = 1 Then
                str缺省部门 = NVL(rsTmp!名称)
            End If
            rsTmp.MoveNext
        Loop
        '如果当前登录人员所属部门都没有设置站点，则不作处理。在查找该院是否启动了站点控制!
        If str站点 = "" Or (bln空站点 And lng有站点 <> 1) Then
            '独立安装新版LIS时也需要按仪器读取站点
            strTmp = GetLISStation()
            If strTmp <> "" Then
                arrTmp = Split(strTmp, ";")
                str站点 = arrTmp(0)
                str名称 = arrTmp(1)
            Else
                str站点 = "": str名称 = ""
                strSQL = "select distinct (A.站点),B.名称 from 部门表 A,zlNodeList B where A.站点=B.编号 And A.站点 is not null order by A.站点"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "检查是否启动站点控制")
                If Not rsTmp Is Nothing Then
                    Do While Not rsTmp.EOF
                        If NVL(rsTmp!站点, "") <> "" Then
                            str站点 = str站点 & "," & NVL(rsTmp!站点, "")
                            str名称 = str名称 & "," & NVL(rsTmp!名称)
                        End If
                        rsTmp.MoveNext
                    Loop
                End If
            End If
        End If
        If str站点 <> "" Then
            str站点 = Mid(str站点, 2)
            str名称 = Mid(str名称, 2)
            arrTmp = Split(str站点, ",")
            For i = LBound(arrTmp) To UBound(arrTmp)
                If i = LBound(arrTmp) Then
                    str站点编号 = arrTmp(i)
                Else
                    If str站点编号 <> arrTmp(i) Then
                        bln多站点 = True
                        Exit For
                    End If
                End If
            Next
            If bln多站点 Then '提示用户选择当前计算机位置所在的部门。
                str站点编号 = GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & App.EXEName, "当前站点选择", "")
                Call frmSelClient.ShowEdit(str站点, str名称, str站点编号)
                str站点编号 = IIf(frmSelClient.gstr站点 = "无", "", frmSelClient.gstr站点)
                gstrDeptName = frmSelClient.gstrCur站点
                Call SaveSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & App.EXEName, "当前站点选择", str站点编号)
            End If
        End If
    End If
    If str站点编号 <> "" Then
        Call zl9ComLib.SetNodeNo(str站点编号)
    Else
        Call zl9ComLib.SetNodeNo("-")
    End If
    If gstrDeptName = "" Then gstrDeptName = str缺省部门
    If strRowID = "" Then '新增的工作站，还没有该工作站的数据，上传（IP、机器名、CPU、内存、硬盘、操作系统）
        int升级标志 = 1
        strSQL = "select 内容 from zlreginfo where 项目='升级类型'"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "检查使用的升级类型")
        If Not rsTmp.EOF Then
            If NVL(rsTmp!内容, 0) = 0 Then
                strSQL = "select replace(项目,'服务器目录','') as 服务器 from zlreginfo where 项目 like '服务器目录%' and 内容 is not null"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "检查是否存配置的文件共享服务器")
                If rsTmp.EOF Then
                    int升级标志 = 0
                Else
                    int服务器编号 = Val(rsTmp!服务器 & "")
                End If
            Else
                strSQL = "select replace(项目,'FTP服务器','') as FTP服务器 from zlreginfo where 项目 like 'FTP服务器%' and 内容 is not null"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "检查是否存配置的FTP服务器")
                If rsTmp.EOF Then
                    int升级标志 = 0
                Else
                    int服务器编号 = Val(rsTmp!FTP服务器 & "")
                End If
            End If
        End If
        strSQL = "Zl_Zlclients_Set(0,Null,'" & strComputerName & "','" & strIpAddress & "','" & arrComputer(2) & "','" & arrComputer(3) & _
                    "','" & arrComputer(4) & "','" & arrComputer(5) & "','" & gstrDeptName & "',Null,Null," & int服务器编号 & "," & int升级标志 & _
                    ",0,'" & str站点编号 & "',0,Null,Null," & int启用视频源 & ")"
        zlDatabase.ExecuteProcedure strSQL, "新增工作站"
        '新增客户端不能升级则直接退出
        If int升级标志 = 0 Then
            CheckAllowByTerminal = True
            Exit Function
        End If
        blnUpdate = True
    Else
        strSQL = "Zl_Zlclients_Set(1,'" & strRowID & "','" & strComputerName & "','" & strIpAddress & "','" & arrComputer(2) & "','" & arrComputer(3) & _
                    "','" & arrComputer(4) & "','" & arrComputer(5) & "','" & gstrDeptName & "',Null,Null,Null,Null," & int连接数 & ",'" & str站点编号 & "',0,Null,Null," & int启用视频源 & ")"
        '需要更新相关的站点信息
        zlDatabase.ExecuteProcedure strSQL, "更新工作站"
        If Not blnAllow Then
            MsgBox "该工作站已被管理员禁用！", vbInformation, gstrSysName
            Exit Function
        End If
        '连接数检查限制
        If int连接数 > 0 Then
            strSQL = "Select SID From gv$Session Where Upper(PROGRAM) Like 'ZL%.EXE' And Status<>'KILLED' And MACHINE=(Select Max(MACHINE) From v$Session Where AUDSID=UserENV('SessionID'))"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "检查连接数量")
            If rsTmp.RecordCount > int连接数 Then
                MsgBox "当前工作站最多只允许 " & int连接数 & " 个登录连接，当前已经有 " & rsTmp.RecordCount - 1 & " 个连接。", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End If
    On Error GoTo errHand
AutoUpGrude:      '执升升级程序
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

Public Function StartHisCrust(ByVal str升级程序 As String, ByVal strJobName As String, Optional ByVal lngWait As Long) As Boolean
'功能：调用自动升级外壳
'参数：str升级程序=可以直接传完成文件路径，也可以传文件名
'      strJobName=任务名称，或者调用程序名
'      lngWait=正式升级时，等待的N分钟后才正式升级
'返回：是否成功
    Dim strUP As String
    Dim strUPFile  As String, strFileName As String
    Dim strConnString As String, lngErr As Long
    Dim objFile As New FileSystemObject
    Dim strCheck As String, strCommand As String
    
    On Error Resume Next
    If objFile.GetDriveName(str升级程序) = "" Then
        strUPFile = App.Path & "\" & str升级程序
    Else
        strUPFile = str升级程序
        strFileName = objFile.GetFileName(str升级程序)
    End If
    If Not objFile.FileExists(strUPFile) Then
        MsgBox "没有找到客户端自动升级工具" & strFileName & "，请与系统管理员联系。", vbExclamation, gstrSysName
        Exit Function
    End If
    If OS.IsDesinMode Then
        '组装命令行，以及生成命令行校验位
        strCommand = "Provider=MSDataShape.1;Extended Properties=""Driver={Microsoft ODBC for Oracle};Server=" & gobjRelogin.ServerName & _
                                   """;Persist Security Info=True;User ID=" & gobjRelogin.InputUser & ";Password=HIS;Data Provider=MSDASQL"
    Else
        '组装命令行，以及生成命令行校验位
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
    '编写人:朱玉宝 2003-03-09
    '功能:分析出本机的配置（IP、机器名、CPU、内存、硬盘、操作系统）
    Dim strCPU As String           'CPU
    Dim strMemory As String        '内存
    Dim strOS As String            '操作系统
    Dim strComputerName As String  '计算机名
    Dim strHD As String            '硬盘
    Dim strIp As String            'IP地址
    Dim verinfo As OSVERSIONINFO
    Dim sysinfo As SYSTEM_INFO
    Dim memsts As MEMORYSTATUS
    Dim memory As Long
    
    strIp = OS.IP
    '获取计算机名
    strComputerName = OS.ComputerName
    '获取硬盘信息
    strHD = AnalyseHardDisk
    ' 获得操作系统信息
    strOS = GetVersionInfo
    ' 获得CPU类型
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
    ' 获得剩余内存
    GlobalMemoryStatus memsts
    memory = memsts.dwTotalPhys
    strMemory = Format$(memory& \ 1024 \ 1024, "###,###,###") + "M"
    AnalyseConfigure = strIp & STRSPLIT & strComputerName & STRSPLIT & strCPU & _
                       STRSPLIT & strMemory & STRSPLIT & strHD & STRSPLIT & strOS
End Function

Private Function AnalyseHardDisk() As String
    '编写人:朱玉宝 2003-03-09
    '功能:获取硬盘总容量
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
    
    '如果是Windows2000或以下版本，则用新API再取一次
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
    '检查是否有重复登录
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    Dim strProgram As String
    On Error GoTo errHand
    
    strProgram = App.EXEName & ".exe"
    strSQL = "Select A.UserName, A.Program, B.IP" & vbNewLine & _
            "From gv$Session A, zlClients B" & vbNewLine & _
            "Where A.Terminal = B.工作站" & vbNewLine & _
            "      And A.Terminal = (Select Terminal From v$Session Where AudsID = Userenv('SessionID') and RowNum =1)" & vbNewLine & _
            "      And A.Program =[1] And A.AudsID <> Userenv('SessionID')" & vbNewLine & _
            "      And B.IP <> [2]"

    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "检查重复工作站", strProgram, strIpAddress)
    If rsTemp.RecordCount = 0 Then '可以登录
        CheckRepeatLogin = False
        Exit Function
    Else
        MsgBox "局域网中存在相同名称的计算机登录," & vbCrLf & "对方IP是:[" & NVL(rsTemp!IP) & "]", vbInformation, gstrSysName
        CheckRepeatLogin = True
        Exit Function
    End If
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Function GetLISStation() As String
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'功能   得到独立新版LIS的站点
'返回   得到站点和站点名称  空为没有站点
'        有的组织方式为 ,1,2;,站点1,站点2
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim str站点  As String, str站点名称 As String
    
    On Error GoTo errHand
    '判断是否独立安装
    strSQL = "select 1 计数 from zlsystems where 编号 = 2500 and 共享号 is null"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "检查是否独立安装新版LIS")
    If rsTmp.EOF Then Exit Function
    '查找是否有默认的站点
    strSQL = "Select Distinct A.站点, B.名称" & vbNewLine & _
            "From (Select Distinct A.站点" & vbNewLine & _
            "       From 检验仪器记录 A, 检验仪器人员 B, 人员表 C,上机人员表 d" & vbNewLine & _
            "       Where A.Id = B.仪器id And A.站点 Is Not Null And B.人员id = C.Id and c.id = d.人员ID And d.用户名 = [1]) A, Zlnodelist B" & vbNewLine & _
            "Where A.站点 = B.编号" & vbNewLine & _
            "Order By A.站点"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "站点查询", gstrDbUser)
    Do While Not rsTmp.EOF
        str站点 = str站点 & "," & rsTmp!站点
        str站点名称 = str站点名称 & "," & rsTmp!名称
        rsTmp.MoveNext
    Loop
    If str站点 <> "" Then
        GetLISStation = str站点 & ";" & str站点名称
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
'功能：调用ZLHIS进行升级
'      blnBrwCall=是否导航台调用,导航台调用升级时检查预升级时点

    Dim strUpdateExe As String, strUpdateExePath As String
    Dim objFSO As New FileSystemObject
    Dim objConn As clsConnect, datCur As Date
    Dim rsTemp As ADODB.Recordset, strSQL As String
    Dim strJobName As String, blnDownload As Boolean
    Dim strTmpPath As String, lngWait As Long
    
    '没有升级，收集等任务，则自动退出升级
    If Not CheckJobs(strComputerName, strJobName, blnBrwCall, blnForceUpdate) Then
        If blnForceUpdate Then
            MsgBox "当前只能进行预升级，无法进行客户端修复！", vbInformation, gstrSysName
        Else
            UpdateZLHIS = True
        End If
        Exit Function
    End If
    
    If strJobName = "OfficialUpgrade" And blnBrwCall Then
        MsgBox "检测到系统需要升级，随后系统将进行升级！", vbInformation, gstrSysName
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
    '升级程序不存在，则准备下载
    If Not objFSO.FileExists(strUpdateExePath) Then
        '先准备临时升级目录
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
            MsgBox "没有找到客户端自动升级工具" & strUpdateExe & "并且无法通过升级服务器下载，请与系统管理员联系。", vbExclamation, gstrSysName
            Exit Function
        End If
        blnDownload = objConn.DownloadFile("zlHisCrust.exe", strTmpPath)
        If blnDownload Then
            On Error Resume Next
            '先清理本地文件
            If objFSO.FileExists(strUpdateExePath) Then
                If FileSystem.GetAttr(strUpdateExePath) <> vbNormal Then
                     Call FileSystem.SetAttr(strUpdateExePath, vbNormal)
                End If
                Call objFSO.DeleteFile(strUpdateExePath)
            End If
            If Err.Number <> 0 Then Err.Clear
            '先复制到APPSOFT下，如果失败，则复制到APPLY下
            objFSO.CopyFile strTmpPath, strUpdateExePath, True
            If Err.Number <> 0 Then
                Err.Clear
                If OS.IsDesinMode Then
                    strUpdateExePath = "C:\APPSOFT\APPLY\zlHisCrust.exe"
                Else
                    strUpdateExePath = App.Path & "\APPLY\zlHisCrust.exe"
                End If
                '先清理本地文件
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
                    '是否是新版自动升级外壳，是的话，则可以直接从临时目录启动。
                    If UCase(GetFileDesInfo(strTmpPath, "ProductName")) = "ZLHISINSTALLUPDATE" Then
                        strUpdateExePath = strTmpPath
                    End If
                End If
            End If
        End If
        If strTmpPath <> strUpdateExePath Then
            On Error Resume Next
            '临时路径
            If objFSO.FileExists(strTmpPath) Then
                If FileSystem.GetAttr(strTmpPath) <> vbNormal Then
                     Call FileSystem.SetAttr(strTmpPath, vbNormal)
                End If
                Call objFSO.DeleteFile(strTmpPath)
            End If
            Call objFSO.DeleteFolder(objFSO.GetParentFolderName(strTmpPath))
        End If
        If Not objFSO.FileExists(strUpdateExePath) Then
            MsgBox "没有找到客户端自动升级工具" & strUpdateExe & "并且无法通过升级服务器下载，请与系统管理员联系。", vbExclamation, gstrSysName
            Exit Function
        End If
    End If
    Call SaveSetting("ZLSOFT", "公共全局", "升级程序", UCase(strUpdateExe)) '用于ZLRegister中特殊判断
    If StartHisCrust(strUpdateExePath, strJobName, lngWait) And strJobName <> "PreUpgrade" Then
        End
    End If
    UpdateZLHIS = True
End Function

Private Function CheckJobs(ByVal strComputerName As String, ByRef strJobName As String, Optional ByVal blnBrwCall As Boolean, Optional ByVal blnForceUpdate As Boolean) As Boolean
'功能:检查并获取升级程序的任务
'      blnBrwCall=是否导航台调用,导航台调用升级时检查预升级时点
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim datCur As Date, blnOnlyOfficialUp As Boolean, blnOnlyPreUp As Boolean
    Dim blnPreUp As Boolean, blnOfficialUp As Boolean, blnPreComplete As Boolean, blnCollect As Boolean
    Dim strStartTime As String, strEndTime As String
    
    On Error GoTo errH
    strJobName = ""
    '以下代码一般不可能出错
    datCur = zlDatabase.Currentdate
    '判断任务是否合理，获取是否启用了定时升级
    strSQL = "Select Max(内容) 内容 From zlRegInfo Where 项目='客户端升级日期'"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "检查定时升级")
    If rsTmp!内容 & "" <> "" Then
        If CDate(Format(datCur, "yyyy-MM-dd HH:mm:ss")) >= CDate(Format(NVL(rsTmp!内容), "yyyy-MM-dd HH:mm:ss")) Then
            blnOnlyOfficialUp = True '只能正式升级
        Else
            blnOnlyPreUp = True '只能预升级
        End If
    Else
        blnOnlyOfficialUp = True
    End If
    On Error Resume Next
    Set rsTmp = Nothing
    '可能没有是否预升级字段(因为预升级时候，数据库还没升级），因此需要错误忽略
    strSQL = "Select 预升时点,Nvl(是否预升级,0) 是否预升级, Nvl(预升完成, 0) 预升完成, Nvl(升级标志, 0) 升级标志, Nvl(收集标志, 0) 收集标志 From Zlclients Where 工作站 = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "检查当前任务", strComputerName)
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo errH
    If Not rsTmp Is Nothing Then
        If Not rsTmp.EOF Then
            blnPreUp = rsTmp!是否预升级 = 1
            blnOfficialUp = rsTmp!升级标志 = 1
            blnPreComplete = rsTmp!预升完成 = 1
            blnCollect = rsTmp!收集标志 = 1
            strStartTime = Format(datCur, "yyyy-mm-dd") & " " & Format(rsTmp!预升时点, "HH:00:00")
            strEndTime = Format(datCur, "yyyy-mm-dd") & " " & Format(rsTmp!预升时点, "HH:59:59")
        End If
    Else
        '优先新方式读取，失败再使用老方式，增加兼容性
        strSQL = "Select 预升时点,Nvl(预升完成, 0) 预升完成, Nvl(升级标志, 0) 升级标志, Nvl(收集标志, 0) 收集标志 From Zlclients Where 工作站 = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "检查当前任务", strComputerName)
        If Not rsTmp.EOF Then
            blnPreUp = rsTmp!升级标志 = 1
            blnOfficialUp = rsTmp!升级标志 = 1
            blnPreComplete = rsTmp!预升完成 = 1
            blnCollect = rsTmp!收集标志 = 1
            strStartTime = Format(datCur, "yyyy-mm-dd") & " " & Format(rsTmp!预升时点, "HH:00:00")
            strEndTime = Format(datCur, "yyyy-mm-dd") & " " & Format(rsTmp!预升时点, "HH:59:59")
        End If
    End If
    '当前只能进行预升级
    If blnOnlyPreUp Then
        '有预升级任务
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
        '没有预升级任务，但是有收集任务
        ElseIf blnCollect Then
            strJobName = "CollectClientFiles"
        Else
            Exit Function
        End If
    '当前只能进行正式升级
    ElseIf blnOnlyOfficialUp Then
        If blnForceUpdate Then
            strJobName = "Repair"
        Else
            '有正式升级任务
            If blnOfficialUp Then
                strJobName = "OfficialUpgrade"
            '没有正式升级任务，但是有收集任务
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
'密码解密程序
    Const MIN_ASC = 32    '最小ASCII码
    Const MAX_ASC = 126 '最大ASCII码 字符
    Const NUM_ASC = MAX_ASC - MIN_ASC + 1
    Dim lngOffset As Long, intlen As Integer, intSeedLen As Integer
    Dim intStart As Integer
    Dim i As Integer, intChr As Integer
    Dim strDeText As String
    
    If strText = "" Then Exit Function
    '随机种子长度
    intSeedLen = Asc(Mid(strText, 1, 1)) - MIN_ASC
    intlen = Len(strText)
    '采用旧的随机算法
    If intSeedLen > 0 And intSeedLen < intlen - 3 And intSeedLen < 5 Then
        '获取随机种子
        '随机种子的随机数为999
        Rnd (-1)
        Randomize (999)
        For i = 2 To 1 + intSeedLen
            intChr = Asc(Mid(strText, i, 1)) '取字母转变成ASCII码
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
        
    '内容解密的种子
    Rnd (-1)
    Randomize (Val(strDeText))
    strDeText = ""
    For i = intStart To intlen
        intChr = Asc(Mid(strText, i, 1)) '取字母转变成ASCII码
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
'功能：去掉字符串中\0以后的字符
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


