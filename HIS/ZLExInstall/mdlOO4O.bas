Attribute VB_Name = "mdlOO4O"
Option Explicit

Public Function InstallOO4O(Optional ByRef strInfo As String) As Boolean
'功能：进行OO4O的安装
    Dim objTemp         As Object
    Dim strTmp          As String, strCLSID     As String
    Dim strOracleHome   As String, strOracleReg As String
    Dim strOracleVer    As String

    On Error Resume Next
    Set objTemp = CreateObject("OracleInProcServer.XOraServer")
    If Err.Number = 0 Then
        strInfo = "已经安装OO4O（可以成功创建类对象OracleInProcServer.XOraServer）"
        InstallOO4O = True
    Else
        Err.Clear
        '安装包是否存在
        strTmp = gstrAPPPath & "\ZLExFile\OO4O"
        '曾经存在BUG,解压后文件结构存在问题，因此这样判断
        If (Not gobjFSO.FolderExists(strTmp) Or Not gobjFSO.FolderExists(strTmp & "\8\Bin")) And gobjFSO.FileExists(strTmp & ".7Z") Then
            Call gobj7z.UnZipFile(strTmp & ".7Z", strTmp, False, , True)
        End If
        If Not gobjFSO.FolderExists(strTmp) Then
            strInfo = "OO4O安装文件不存在（路径：" & strTmp & "）"
            Exit Function
        End If
        'oracle是否安装
        'OracleHOme获取
        strOracleHome = GetOracleHome()
        If strOracleHome = "" Then
            strInfo = "未找到32位ORACLE客户端安装信息"
            Exit Function
        End If
        'ORacle注册表路径获取
        strOracleReg = GetOracleReg(strOracleHome)
        If strOracleReg = "" Then
            strInfo = "未找到Oracle_Home的注册表路径"
            Exit Function
        End If
        'Oracle版本获取
        strOracleVer = GetOracleClientVersion(strOracleHome & "\Bin")
        If strOracleVer = "" Then
            strInfo = "无法获取Oracle客户端版本，可能不支持该版本客户端的OO4O安装（支持8，10，11版本）"
            Exit Function
        End If
        '安装OO4O
        InstallOO4O = InstallComponent(strOracleVer, strOracleHome, strOracleReg)
        Err.Clear
        Set objTemp = CreateObject("OracleInProcServer.XOraServer")
        If Err.Number <> 0 Then
            strInfo = "再次安装验证失败。"
            InstallOO4O = False
        End If
    End If
End Function

Private Function GetOracleHome() As String
'功能：获取OracleHome路径
    Dim arrTmp  As Variant, arrSubKey   As Variant
    Dim strHome As String, strDefault   As String, strPath As String
    Dim i       As Integer
    Dim objPE   As New clsPEReader
    Dim blnRead As Boolean
    
    strHome = Environ("PATH")
    '1、PATH变量都没有，操作系统的环境变量存在问题或者非WIn系统，可能为麦金塔系统（MAC）
    If strHome = "" Then Exit Function
    arrTmp = Split(strHome, ";")
    strHome = ""
    For i = LBound(arrTmp) To UBound(arrTmp)
    
        If UCase(arrTmp(i)) Like "*ORA*\BIN" Then
            '判断Oracle的OCI基础部件是否存在
            If gobjFSO.FileExists(arrTmp(i) & "\oci.dll") Then
                If Not objPE.Is64BitPE(arrTmp(i) & "\oci.dll") Then
                    strHome = gobjFSO.GetParentFolderName(arrTmp(i))
                    If gobjFSO.FileExists(strHome & "\network\ADMIN\tnsnames.ora") Then
                        GetOracleHome = strHome
                        Exit Function
                    End If
                End If
            End If
        End If
    Next
    '2、寻找TNS_ADMIN:ORACLE_HOME & "\network\ADMIN
    strHome = Environ("TNS_ADMIN")
    If strHome <> "" Then
        If InStr(UCase(strHome), "\NETWORK\ADMIN") > 0 Then
            '判断TNSNAME
            If Not gobjFSO.FileExists(strHome & "\tnsnames.ora") Then
                strHome = ""
            End If
            '获取ORACLE_HOME,判断OCI
            If strHome <> "" Then
                strHome = gobjFSO.GetParentFolderName(gobjFSO.GetParentFolderName(strHome))
                If gobjFSO.FileExists(strHome & "\Bin\oci.dll") Then
                    If Not objPE.Is64BitPE(strHome & "\Bin\oci.dll") Then
                        GetOracleHome = strHome
                        Exit Function
                    End If
                End If
            End If
        End If
    End If
    '3、ORACLE_HOME环境变量
    strHome = Environ("ORACLE_HOME")
    If strHome <> "" Then
        If gobjFSO.FileExists(strHome & "\Bin\oci.dll") Then
            If Not objPE.Is64BitPE(strHome & "\Bin\oci.dll") Then
                If gobjFSO.FileExists(strHome & "\network\ADMIN\tnsnames.ora") Then
                    GetOracleHome = strHome
                    Exit Function
                End If
            End If
        End If
    End If
    
    '4、注册表判断,读取64位下32目录会自动定位到SOFTWARE\Wow6432Node\Oracle 2：读取32位下32位目录
    '4.1 ALL_HOMES
    '         DEFAULT_HOME"="DEFAULT_HOME"
    '      ALL_HOMES\ID0
    '        "NAME"="DEFAULT_HOME"
    '        "PATH"="F:\\instantclient_11_2_3"
    blnRead = GetRegValue("HKEY_LOCAL_MACHINE\SOFTWARE\" & IIf(gblnIs64Bits, "WOW6432Node\", "") & "Oracle\ALL_HOMES", "DEFAULT_HOME", strDefault)
    If blnRead And strDefault <> "" Then
        arrSubKey = GetAllSubKey("HKEY_LOCAL_MACHINE\SOFTWARE\" & IIf(gblnIs64Bits, "WOW6432Node\", "") & "Oracle\ALL_HOMES")
        If TypeName(arrSubKey) <> "Empty" Then
            For i = LBound(arrSubKey) To UBound(arrSubKey)
                strHome = ""
                blnRead = GetRegValue("HKEY_LOCAL_MACHINE\SOFTWARE\" & IIf(gblnIs64Bits, "WOW6432Node\", "") & "Oracle\ALL_HOMES\" & arrSubKey(i), "NAME", strHome)
                If blnRead And strHome <> "" Then
                    If strHome = strDefault Then
                        blnRead = GetRegValue("HKEY_LOCAL_MACHINE\SOFTWARE\" & IIf(gblnIs64Bits, "WOW6432Node\", "") & "Oracle\ALL_HOMES\" & arrSubKey(i), "PATH", strPath)
                        If blnRead And strPath <> "" Then
                            If Not objPE.Is64BitPE(strPath & "\Bin\oci.dll") Then
                                If gobjFSO.FileExists(strPath & "\network\ADMIN\tnsnames.ora") Then
                                    GetOracleHome = strPath
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
                End If
            Next
        End If
    End If
    '4.2非ALL_Homes方式,只获取第一个符合条件的。
    arrSubKey = Empty
    arrSubKey = GetAllSubKey("HKEY_LOCAL_MACHINE\SOFTWARE\" & IIf(gblnIs64Bits, "WOW6432Node\", "") & "Oracle")
    If TypeName(arrSubKey) <> "Empty" Then
        For i = LBound(arrSubKey) To UBound(arrSubKey)
            strHome = ""
            blnRead = GetRegValue("HKEY_LOCAL_MACHINE\SOFTWARE\" & IIf(gblnIs64Bits, "WOW6432Node\", "") & "Oracle\" & arrSubKey(i), "ORACLE_HOME", strHome)
            If blnRead And strHome <> "" Then
                If Not objPE.Is64BitPE(strHome & "\Bin\oci.dll") Then
                    If gobjFSO.FileExists(strHome & "\network\ADMIN\tnsnames.ora") Then
                        GetOracleHome = strHome
                        Exit Function
                    End If
                End If
            End If
        Next
    End If
End Function

Private Function GetOracleReg(ByVal strOracleHome As String) As String
'功能：通过Oracle_Home路径获取注册表中位置
    Dim arrTmp      As Variant, arrSubKey   As Variant
    Dim strHomeName As String, strHome      As String
    Dim i           As Integer
    Dim blnRead     As Boolean

    arrSubKey = GetAllSubKey("HKEY_LOCAL_MACHINE\SOFTWARE\" & IIf(gblnIs64Bits, "WOW6432Node\", "") & "Oracle")
    If TypeName(arrSubKey) <> "Empty" Then
        For i = LBound(arrSubKey) To UBound(arrSubKey)
            strHome = ""
            blnRead = GetRegValue("HKEY_LOCAL_MACHINE\SOFTWARE\" & IIf(gblnIs64Bits, "WOW6432Node\", "") & "Oracle\" & arrSubKey(i), "ORACLE_HOME", strHome)
            If blnRead And strHome <> "" Then
                If UCase(strHome) = UCase(strOracleHome) Then
                    GetOracleReg = "HKEY_LOCAL_MACHINE\SOFTWARE\" & IIf(gblnIs64Bits, "WOW6432Node\", "") & "Oracle\" & arrSubKey(i)
                    Exit Function
                End If
            End If
        Next
    End If
End Function

Private Function GetOracleClientVersion(ByVal strBinPath As String) As String
'功能：根据OralceHome路径部件获取Oracle版本，只返回大版本,
    Dim i As Long
    Dim arrTmp As Variant
    
    arrTmp = Split("8,10,11", ",")
    For i = LBound(arrTmp) To UBound(arrTmp)
        If gobjFSO.FileExists(strBinPath & "\ORANNZSBB" & arrTmp(i) & ".dll") Then
            GetOracleClientVersion = arrTmp(i)
            Exit Function
        ElseIf gobjFSO.FileExists(strBinPath & "\ORAJDBC" & arrTmp(i) & ".dll") Then
            GetOracleClientVersion = arrTmp(i)
            Exit Function
        ElseIf gobjFSO.FileExists(strBinPath & "\oraocci" & arrTmp(i) & ".dll") Then
            GetOracleClientVersion = arrTmp(i)
            Exit Function
        End If
    Next
End Function

Private Function InstallComponent(ByVal strOracleVer As String, ByVal strOracleHome As String, ByVal strOracleReg As String) As Boolean
'功能：安装OO4O部件
'参数：strOracleHome=OracleHOme
'strOracleVer:当前Oracle版本
'返回：是否安装成功
    Dim strSourcePath   As String
    Dim objFile         As File
    Dim strErr          As String
    
    strSourcePath = gstrAPPPath & "\ZLExFile\OO4O\" & strOracleVer
    Call XCopy(strSourcePath, strOracleHome)
    '11g在OracleHOME下有OraOO4Oic11.dll文件，其他两个版本没有
    '注册文件
    'regsvr32 /s "%BAT_DIR%bin\oradc.ocx"
    'regsvr32 /s "%BAT_DIR%bin\OIP11.dll"
    'regsvr32 /s "%BAT_DIR%bin\oo4ocodewiz.dll"
    'regsvr32 /s "%BAT_DIR%bin\odbtreeview.ocx"
    'regsvr32 /s "%BAT_DIR%bin\oo4oaddin.dll"
    If Not gclsRegCom.RegCom(strOracleHome & "\Bin\oradc.ocx", RFT_NormalReg, strErr) Then
        strErr = strErr & ",oradc.ocx"
    End If
    
    If Not gclsRegCom.RegCom(strOracleHome & "\Bin\OIP" & strOracleVer & ".dll", RFT_NormalReg, strErr) Then
        strErr = strErr & ",OIP" & strOracleVer & ".dll"
    End If
    
    If Not gclsRegCom.RegCom(strOracleHome & "\Bin\oo4ocodewiz.dll", RFT_NormalReg, strErr) Then
        strErr = strErr & ",oo4ocodewiz.dll"
    End If
    
    If Not gclsRegCom.RegCom(strOracleHome & "\Bin\odbtreeview.ocx", RFT_NormalReg, strErr) Then
        strErr = strErr & ",odbtreeview.ocx"
    End If
    
    If Not gclsRegCom.RegCom(strOracleHome & "\Bin\oo4oaddin.dll", RFT_NormalReg, strErr) Then
        strErr = strErr & ",oo4oaddin.dll"
    End If
    '注册表处理
    'echo Windows Registry Editor Version 5.00                              >  "%BAT_DIR%"\oo4o.reg
    'echo [HKEY_LOCAL_MACHINE\SOFTWARE\%ODAC_CFG_PREFIX%Oracle\KEY_%2]      >> "%BAT_DIR%"\oo4o.reg
    'echo "OO4O"="%REG_DIR%oo4o\\mesg"                                      >> "%BAT_DIR%"\oo4o.reg
    'echo [HKEY_LOCAL_MACHINE\SOFTWARE\%ODAC_CFG_PREFIX%Oracle\KEY_%2\OO4O] >> "%BAT_DIR%"\oo4o.reg
    'echo "CacheBlocks"="20"                                                >> "%BAT_DIR%"\oo4o.reg
    'echo "FetchLimit"="100"                                                >> "%BAT_DIR%"\oo4o.reg
    'echo "FetchSize"="4096"                                                >> "%BAT_DIR%"\oo4o.reg
    'echo "PerBlock"="16"                                                   >> "%BAT_DIR%"\oo4o.reg
    'echo "SliceSize"="256"                                                 >> "%BAT_DIR%"\oo4o.reg
    'echo "TempFileDirectory"="c:\\temp"                                    >> "%BAT_DIR%"\oo4o.reg
    'echo "OO4O_HOME"="%REG_DIR%oo4o"                                       >> "%BAT_DIR%"\oo4o.reg
    Call CreateRegKey(strOracleReg, "OO4O", strOracleHome & "\OO4O\mesg")
    Call CreateRegKey(strOracleReg & "\OO4O", "CacheBlocks", "20")
    Call CreateRegKey(strOracleReg & "\OO4O", "FetchLimit", "100")
    Call CreateRegKey(strOracleReg & "\OO4O", "FetchSize", "4096")
    Call CreateRegKey(strOracleReg & "\OO4O", "PerBlock", "16")
    Call CreateRegKey(strOracleReg & "\OO4O", "SliceSize", "256")
    Call CreateRegKey(strOracleReg & "\OO4O", "OO4O_HOME", strOracleHome & "\OO4O")
    InstallComponent = strErr = ""
End Function
