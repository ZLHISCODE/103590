Attribute VB_Name = "mdlMain"
Option Explicit
'**************************
'       OEM代号
'
'爱生    B0AEC9FA
'医业    D2BDD2B5
'托普    CDD0C6D5
'中软    D6D0C8ED
'金康泰  BDF0BFB5CCA9
'医院    D2BDD4BA
'宝信    B1A6D0C5
'**************************
Public gobjFile         As New FileSystemObject
Public gcolOwnerConn    As New Collection
Public gblnInIDE        As Boolean  '是否源代码运行
Public gstrOracleVer    As String 'Oracle版本
Public gstrOracleBigVer    As String 'Oracle大版本
Public gblnTrace        As Boolean '是否启用跟踪日志
Public gblnTestUpgrade  As Boolean '测试脚本升级
Public gblnClose11g     As Boolean '关闭11g重复数据不插入的新特性
Public glngInterval     As Long '升级的间隔
Public glngAtuoErr      As Long '自动错误处理
Public gobjRegister     As Object '注册授权部件
'----------------------------------------------------------------------------------------
'--安装脚本执行相关变量定义
Private mobjLog As TextStream

Public Type SQL_DEFINE
    varName As String
    varValue As String
End Type

Public Type listFile
  Filename      As String
  FileVision    As String
  FileEditDate  As String
  FileMD5       As String
End Type

'升级更新列表-zq 20101213
Public Type UpdateList
  uFile() As listFile
End Type
'连接方式
Public Enum enuProvider
    MSODBC = 0
    OraOLEDB = 1
    OriginalConnection = 9
End Enum

Public Const G_STR_USERS As String = "'SYS','SYSTEM','SCOTT','OUTLN','DBSNMP','MTSSYS','MDSYS','ORDSYS','ORDPLUGINS','CTXSYS','ZLTOOLS','XDB','WMSYS','TSMSYS','SYSMAN','SI_INFORMTN_SCHEMA','OLAPSYS','MGMT_VIEW','MDDATA','EXFSYS','DMSYS','DIP','ANONYMOUS'"
'刘兴宏:加入'XDB','WMSYS','TSMSYS','SYSMAN','SI_INFORMTN_SCHEMA','OLAPSYS','MGMT_VIEW','MDDATA','EXFSYS','DMSYS','DIP','ANONYMOUS'

Public gcnOracle As New ADODB.Connection     '以OraOLEDB方式打开的公共数据库连接
Public gcnOldOra As New ADODB.Connection    '以ODBC方式打开的连接，用于执行脚本，用OraOLEDB方式创建存储过程会发生执行成功但是过程没有被更新的问题
Public gcnSystem As ADODB.Connection        'SYSTEM用户连接
Public gcnTools As ADODB.Connection        'ZLTools用户连接

Public gstrUserName As String               '用户名
Public gstrPassword As String               '用户的数据库密码
Public gstrLoginPwd As String               '用户登录时输入的密码
Public gstrLoginUserName As String          '授权用户登录的用户名
Public gstrLoginUserPwd As String           '授权用户登录的数据库密码

Public gstrToolsPwd As String                  '管理工具的密码
Public gstrSysUser As String                     'SYS用户名
Public gstrSysPwd As String                     'SYS密码
Public gstrServer As String                       '服务器名

Public gobjFunction As Object
Public gobjReport As Object
Public gobjUsrProc As Object

Public gstrAppsoft As String                'APPSOFT路径

Public gstrSysName As String                '系统名称
Public gstrProductTitle As String
Public gstrUltimatetag  As String          '旗舰版，专业版标识
Public gstrProductName As String
Public gstrDevelopers As String
Public gstrSustainer As String
Public gstrWebSustainer As String
Public gstrWebURL As String
Public gstrWebEmail As String
Public gstr注册码 As String                 '得到注册码

Public gstrSQL    As String                 '通用的SQL语句变量
Public gblnCreate As Boolean                '是否已经创建管理工具
Public gblnDBA As Boolean                   '是否DBA
Public gblnRac As Boolean                 '是否是Rac环境
Public gintInstID As Integer                  'Rac环境下当前登录实例号
Public gblnOwner As Boolean                 '是否所有者
Public gfrmActive As Form                   '当前活动的子窗口
Public gcbsMain As CommandBars
Public gdtStart As Long
Public gstrHaveProg As String               '如果不是DBA或系统所有者登录，则判断是否有管理工具的权限
Public gblnSystemUser As Boolean            '判断是否为系统所有者登录

Public gstrComputerName As String           '记录当前客户端名称
Public glngSysNo As Long                    '主要用于单系统登录时，记录当前登录的系统的编号

Public gclsBase As New clsBase
Public glngTXTProc As Long


Public Const FindUserWidth = 4845   '查找窗口大小
Public Const FindUserHeight = 5595

Private mstrHasZltables As String  '是否有zltables这张表
Private mstrBigTable As String   '大表
Private mstrMiddleTable As String '中表
Private mstrMiddleTableRows As String

Private Enum REGRoot
    HKEY_CLASSES_ROOT = &H80000000 '记录Windows操作系统中所有数据文件的格式和关联信息，主要记录不同文件的文件名后缀和与之对应的应用程序。其下子键可分为两类，一类是已经注册的各类文件的扩展名，这类子键前面都有一个“。”；另一类是各类文件类型有关信息。
    HKEY_CURRENT_USER = &H80000001 '此根键包含了当前登录用户的用户配置文件信息。这些信息保证不同的用户登录计算机时，使用自己的个性化设置，例如自己定义的墙纸、自己的收件箱、自己的安全访问权限等。
    HKEY_LOCAL_MACHINE = &H80000002 '此根键包含了当前计算机的配置数据，包括所安装的硬件以及软件的设置。这些信息是为所有的用户登录系统服务的。它是整个注册表中最庞大也是最重要的根键！
    HKEY_USERS = &H80000003 '此根键包括默认用户的信息（Default子键）和所有以前登录用户的信息。
    HKEY_PERFORMANCE_DATA = &H80000004 '在Windows NT/2000/XP注册表中虽然没有HKEY_DYN_DATA键，但是它却隐藏了一个名为“HKEY_ PERFOR MANCE_DATA”键。所有系统中的动态信息都是存放在此子键中。系统自带的注册表编辑器无法看到此键
    HKEY_CURRENT_CONFIG = &H80000005  '此根键实际上是HKEY_LOCAL_MACHINE中的一部分，其中存放的是计算机当前设置，如显示器、打印机等外设的设置信息等。它的子键与HKEY_LOCAL_ MACHINE\ Config\0001分支下的数据完全一样。
    HKEY_DYN_DATA = &H80000006 '此根键中保存每次系统启动时，创建的系统配置和当前性能信息。这个根键只存在于Windows 98中。
End Enum

'注册表数据类型
Private Enum REGValueType
    REG_NONE = 0                       ' No value type
    REG_SZ = 1 'Unicode空终结字符串
    REG_EXPAND_SZ = 2 'Unicode空终结字符串
    REG_BINARY = 3 '二进制数值
    REG_DWORD = 4 '32-bit 数字
    REG_DWORD_BIG_ENDIAN = 5
    REG_LINK = 6
    REG_MULTI_SZ = 7 ' 二进制数值串
End Enum
Private Declare Function RegQueryValueEx_ValueType Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegQueryValueEx_String Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function RegQueryValueEx_Long Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Long, lpcbData As Long) As Long
Private Declare Function RegQueryValueEx_BINARY Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
Private Declare Function ExpandEnvironmentStrings Lib "kernel32" Alias "ExpandEnvironmentStringsA" (ByVal lpSrc As String, ByVal lpDst As String, ByVal nSize As Long) As Long

Private Declare Function RegSetValueEx_String Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal lpcbData As Long) As Long
Private Declare Function RegSetValueEx_Long Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long
Private Declare Function RegSetValueEx_BINARY Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Byte, ByVal cbData As Long) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long


Public Function ShowHelp(SHwnd As Long, ByVal htmName As String) As Boolean
'显示帮助窗体
'SHwnd:传入窗口句柄(作为宿主窗口)
'htmName:射映在CHM中的htm文件名称

    Dim Path As String
    Dim strSave As String
    On Error GoTo ShowHelpErr
    
    ShowHelp = False
    strSave = String(200, Chr$(0))
    Path = Left$(strSave, GetWindowsDirectory(strSave, Len(strSave))) + "\help\"
    If CBool(PathIsDirectory(Path)) = False Then GoTo ShowHelpErr
    strSave = "zl9server.CHM"
    Path = Trim(Path & strSave)
    If Trim(Dir(Path)) = "" Then GoTo ShowHelpErr
    Call Htmlhelp(SHwnd, Path, &H0, htmName & ".htm")
    ShowHelp = True
    Exit Function

ShowHelpErr:
    err.Clear
End Function

Public Sub Main()
    '为实现XP风格，在显示窗体前必须执行该函数
    Dim strRegErr As String, strFile As String, strTmp As String
    Dim strLog As String
    Dim strCommand As String
    Dim blnAnalysis As Boolean
    Dim objclsCiph As clsCipher
    
    Call InitCommonControls

    '为了实现注销功能，对全局变量进行初始化
    gblnCreate = False
    gblnDBA = False
    gblnOwner = False
    gblnRac = False
    gintInstID = 0
    Set gfrmActive = Nothing
    Set gobjFunction = Nothing
    Set gobjReport = Nothing
    Set gobjUsrProc = Nothing
    
    gblnInIDE = gclsBase.InDesign
    gblnTrace = gblnInIDE
    '是否启用管理工具跟踪
    gblnTrace = gblnTrace Or Val(GetSetting("ZLSOFT", "公共模块\服务器管理工具", "启用跟踪", "0")) = 1
    glngInterval = Val(GetSetting("ZLSOFT", "公共模块\服务器管理工具", "自动升迁频度", "0"))
    glngAtuoErr = Val(GetSetting("ZLSOFT", "公共模块\服务器管理工具", "错误重试次数", "0"))
    gblnTestUpgrade = Val(GetSetting("ZLSOFT", "公共模块\服务器管理工具", "测试脚本升级", "0")) = 1
    gblnClose11g = Val(GetSetting("ZLSOFT", "公共模块\服务器管理工具", "关闭11G新特性", "0")) = 1
    gstrComputerName = GetMyCompterName
    If glngInterval < 100 Then glngInterval = 100
    If gblnTrace Then
        strLog = GetLogPath(LT_跟踪日志)
        If strLog <> "" Then
            Set gobjLog = gobjFile.CreateTextFile(strLog, True)
        End If
    End If
    gdtStart = Timer
    '获取Shell字符串
    strCommand = CStr(Command())

    '判断是否对字符串进行了加密，若加密了，则进行解密
    If InStr(strCommand, "EncryptedLoginToken:") = 1 Then
        strCommand = Mid(strCommand, Len("EncryptedLoginToken:") + 1)
        Set objclsCiph = New clsCipher
        strCommand = objclsCiph.Decipher(MSTR_DBLINK_KEY, strCommand)
    End If
    '获取系统编号，并将系统编号信息截取掉
    glngSysNo = GetSysNo(strCommand)
    If InStr(strCommand, "=") <= 0 Then
        frmSplash.ShowSplash
    End If
    Do
        If (Timer - gdtStart) > 1 Then Exit Do
        DoEvents
    Loop
    
    On Error Resume Next
    Set gobjRegister = CreateObject("zlRegister.clsRegister")
    err.Clear: On Error GoTo 0
    If gobjRegister Is Nothing Then
        MsgBox "创建zlRegister部件对象失败。请检查文件是否存在并且正确注册。", vbExclamation, gstrSysName
        End
    End If
    
    '检查部件的MD5值(调试模式App.Path是当前源码工程的位置，所以不检查)
    If Not gblnInIDE Then
        strFile = App.Path & "\PUBLIC\zlRegister.Dll"
        strTmp = Md5_File_Calc(strFile)
        If strTmp <> "7F8912644328C37023F6839CDB4E7425" Then
            '10.35.90:13653ED7AF4144CAADB4CD5BF790C731
            '10.35.90SP1:F1335F5042068CF291B8141418775FD8
            MsgBox "验证注册授权部件失败,请检查文件" & strFile & "的版本是否与管理工具的版本匹配。", vbExclamation, gstrSysName
            End
        End If
    End If
    

    '获取当前登录系统编号，并将原字符串中的关于系统编号的参数剔除掉
    If InStr(strCommand, "=") > 0 Or glngSysNo <> -1 Then
        If Not frmUserLogin.Docmd(strCommand, blnAnalysis) Then
            If blnAnalysis = True Then  '表示以第一种方式解析成功，但是登录失败
                '若为单系统登录且登录失败，则不再提供手工输入
                If glngSysNo = -1 Then
                    frmUserLogin.Visible = True
                Else
                    Unload frmUserLogin
                    Set gcnOracle = Nothing
                End If
            Else  '表示以第一种方式解析失败，现尝试使用第二种方式解析
                frmUserLogin.ShowMe strCommand
                If glngSysNo <> -1 Then
                    Unload frmUserLogin
                    Set gcnOracle = Nothing
                End If
            End If
        End If
    Else
        frmUserLogin.ShowMe strCommand
    End If
    
    If InStr(strCommand, "=") <= 0 Then
        Unload frmSplash
    End If
    
    If gcnOracle.State = adStateOpen Then
        SaveSetting "ZLSOFT", "公共全局", "程序路径", App.Path & "\" & App.EXEName & ".exe"
        If gblnCreate = False Then
            '尚创建管理工具，进行创建
            MsgBox "首次运行系统，需要首先创建管理工具。", vbExclamation, "提示"
            frmSvrCreate.Show 1
        Else
            '以ODBC方式打开的连接，用于执行脚本，用OraOLEDB方式创建存储过程会发生执行成功但是过程没有被更新的问题
            Set gcnOldOra = gobjRegister.ReGetConnection(MSODBC, strRegErr)
            If strRegErr <> "" Then
                MsgBox strRegErr, vbQuestion, "提醒"
                gcnOracle.Close
                End
            End If
            Call SetSQLTrace(gstrServer, gstrUserName, gcnOldOra)
            Select Case gobjRegister.zlRegInfo("授权性质")
                Case "1"
                    '正式
                    SaveSetting "ZLSOFT", "注册信息", "Kind", ""
                Case "2"
                    '试用
                    SaveSetting "ZLSOFT", "注册信息", "Kind", "试用"
                Case "3"
                    '测试
                    SaveSetting "ZLSOFT", "注册信息", "Kind", "测试"
            End Select
            frmMDIMain.Show
        End If
    End If
End Sub

Private Function GetSysNo(ByRef strCmd As String) As Long
    '获取当前登录系统编号，并将原字符串中的关于系统编号的参数剔除掉
    '若为常规登录，则系统编号为-1
    Dim ArrCommand() As String
    Dim strCommand As String
    Dim i As Long
    
    ArrCommand = Split(strCmd, " ")
    For i = LBound(ArrCommand) To UBound(ArrCommand)
        If UCase(ArrCommand(i)) Like "SYS=*" Then
            GetSysNo = Val(Split(ArrCommand(i), "=")(1))
        Else
            strCommand = strCommand & " " & ArrCommand(i)
        End If
    Next
    strCmd = Trim(strCommand)
    If GetSysNo = 0 Then GetSysNo = -1
End Function

Public Sub SelAll(objTxt As Control)
'功能：对文本框的的文本选中
    If TypeName(objTxt) = "TextBox" Or TypeName(objTxt) = "ComboBox" Then
        If Trim(objTxt.Text) = "" Then Exit Sub
        objTxt.SelStart = 0: objTxt.SelLength = Len(objTxt.Text)
    ElseIf TypeName(objTxt) = "MaskEdBox" Then
        If Not IsDate(objTxt.Text) Then
            objTxt.SelStart = 0: objTxt.SelLength = Len(objTxt.Text)
        Else
            objTxt.SelStart = 0: objTxt.SelLength = 10
        End If
    End If
End Sub

Public Function StrIsValid(ByVal strInput As String, Optional ByVal intMax As Integer = 0) As Boolean
'检查字符串是否含有非法字符；如果提供长度，对长度的合法性也作检测。
    If InStr(strInput, "'") > 0 Or InStr(strInput, """") > 0 Then
        MsgBox "所输入内容含有非法字符。", vbExclamation, gstrSysName
        Exit Function
    End If
    If intMax > 0 Then
        If LenB(StrConv(strInput, vbFromUnicode)) > intMax Then
            MsgBox "所输入内容不能超过" & Int(intMax / 2) & "个汉字" & "或" & intMax & "个字母。", vbExclamation, gstrSysName
            Exit Function
        End If
    End If
    StrIsValid = True
End Function

Public Sub Get注册码()
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errH
    rsTemp.CursorLocation = adUseClient
    Set rsTemp = OpenCursor(gcnOracle, "ZLTOOLS.B_Public.Get_reginfo", "授权证章")
    gstr注册码 = ""
    Do Until rsTemp.EOF
        gstr注册码 = gstr注册码 & IIf(IsNull(rsTemp!内容), "", rsTemp!内容)
        rsTemp.MoveNext
    Loop
    Exit Sub
errH:
    gstr注册码 = ""
End Sub

Public Function CurrentDate() As Date
    '-------------------------------------------------------------
    '功能：提取服务器上当前日期
    '参数：
    '返回：由于Oracle日期格式的问题，所以
    '-------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    
    err = 0
    On Error GoTo errH
    '不能调用OpenSQLRecord,因为OpenSQLRecord也使用了该方法
    With rsTemp
        .CursorLocation = adUseClient
        .Open "SELECT SYSDATE FROM DUAL", gcnOracle, adOpenKeyset
    End With
    CurrentDate = rsTemp.Fields(0).value
    rsTemp.Close
    Exit Function
errH:
    If MsgBox(err.Description, vbRetryCancel, gstrSysName) = vbRetry Then
        Resume
    End If
    CurrentDate = 0
    err = 0
End Function


'将PictureBox模拟成3D平面按钮
'intStyle=0=平面,-1=凹下,1=凸起,-2=深凹下,2=深凸起
Public Sub RaisEffect(picBox As PictureBox, Optional IntStyle As Integer, Optional strName As String = "")
    Dim PicRect As RECT
    Dim lngTmp As Long
    With picBox
        lngTmp = .ScaleMode
        .ScaleMode = 3
        .Cls
        .BorderStyle = 0
        
        If IntStyle <> 0 Then
            PicRect.Left = .ScaleLeft
            PicRect.Top = .ScaleTop
            PicRect.Right = .ScaleWidth
            PicRect.Bottom = .ScaleHeight
            
            Select Case IntStyle
                Case 1
                    DrawEdge .hDC, PicRect, CLng(BDR_RAISEDINNER), BF_RECT
                Case 2
                    DrawEdge .hDC, PicRect, CLng(EDGE_RAISED), BF_RECT
                Case -1
                    DrawEdge .hDC, PicRect, CLng(BDR_SUNKENOUTER), BF_RECT
                Case -2
                    DrawEdge .hDC, PicRect, CLng(EDGE_SUNKEN), BF_RECT
            End Select
        End If
        .ScaleMode = lngTmp
        If strName <> "" Then
            .CurrentX = (.ScaleWidth - .TextWidth(strName)) / 2
            .CurrentY = (.ScaleHeight - .TextHeight(strName)) / 2
            picBox.Print strName
        End If
    End With
End Sub

Public Sub DeleteAllLog(FrmObj As Form, BlnRunTimeLog As Boolean)
    Dim strRemarks As String
    Dim strNote As String
    
    '验证身份并输入操作说明
    If Not CheckAuditStatus(frmMDIMain.gstrLastModule, "删除", strRemarks) Then Exit Sub
    Call OpenCursor(gcnOracle, "ZLTOOLS.B_Public.Delete_All_Log", IIf(BlnRunTimeLog, 1, 0))
    '插入重要操作日志
    Call SaveAuditLog(3, "删除", "删除所有日志", strRemarks)
    Call FrmObj.RefreshData
    Exit Sub
errHandle:
    If MsgBox(err.Description, vbRetryCancel, gstrSysName) = vbRetry Then
        Resume
    End If

End Sub

Public Sub DeleteCurLog(FrmObj As Form, BlnRunTimeLog As Boolean)
    Dim LngDelete As Long, ItemThis As ListItem
    Dim lng会话号 As Long, str工作站 As String, str用户名 As String, str部件名 As String
    Dim str工作内容 As String
    Dim date时间 As Date, lng类型 As Long, lng错误序号 As Long
    Dim strRemarks As String
    Dim strNote As String
    
    If MsgBox("你确认要删除所选择的日志记录吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    '验证身份并输入操作说明
    If Not CheckAuditStatus(frmMDIMain.gstrLastModule, "删除", strRemarks) Then Exit Sub
    On Error Resume Next
    err = 0
    
    gcnOracle.BeginTrans
    For LngDelete = 1 To FrmObj.LvwList.ListItems.Count
        If FrmObj.LvwList.ListItems(LngDelete).Selected Then
            Set ItemThis = FrmObj.LvwList.ListItems(LngDelete)
            If BlnRunTimeLog Then
                lng会话号 = Val(ItemThis.Tag)
                str工作站 = ItemThis.SubItems(1)
                str用户名 = ItemThis.SubItems(2)
                str部件名 = ItemThis.SubItems(3)
                str工作内容 = ItemThis.SubItems(4)
                date时间 = CDate(ItemThis.SubItems(5))
                Call OpenCursor(gcnOracle, "ZLTOOLS.B_Public.Delete_Diarylog", _
                                         lng会话号, str用户名, str工作站, str部件名, str工作内容, date时间)
                
            Else
                lng会话号 = Val(ItemThis.Tag)
                str工作站 = ItemThis.SubItems(1)
                str用户名 = ItemThis.SubItems(2)
                lng类型 = Val(IIf(ItemThis = "存储过程错误", 1, IIf(ItemThis = "数据联结层错误", 2, 3)))
                lng错误序号 = Val(ItemThis.SubItems(4))
                date时间 = CDate(ItemThis.SubItems(3))
                Call OpenCursor(gcnOracle, "ZLTOOLS.B_Public.Delete_Errorlog", _
                                         lng会话号, str用户名, str工作站, lng类型, lng错误序号, date时间)
            End If
        End If
    Next
    
    If err <> 0 Then
        MsgBox "删除时发生不可预知的错误！", vbInformation, gstrSysName
        gcnOracle.RollbackTrans
        Exit Sub
    End If
    gcnOracle.CommitTrans
    With FrmObj
    If BlnRunTimeLog Then
        If .Cbo工作站.Text <> "" Then strNote = ",工作站:" & .Cbo工作站.Text
        If .Cbo用户名.Text <> "" Then strNote = strNote & ",用户名:" & .Cbo用户名.Text
        If .Txt部件名.Text <> "" Then strNote = strNote & ",部件名:" & .Txt部件名.Text
        If .txt工作内容.Text <> "" Then strNote = strNote & ",模块名" & .txt工作内容.Text
        strNote = strNote & ",起始时间:" & Format(.dtpDateStart, "yyyy-MM-dd") & ",终止时间:" & Format(.dtpDateEnd, "yyyy-MM-dd")
    Else
        If .Cbo工作站.Text <> "" Then strNote = ",工作站:" & .Cbo工作站.Text
        If .Cbo用户名.Text <> "" Then strNote = strNote & ",用户名:" & .Cbo用户名.Text
        strNote = strNote & ",错误类型:" & .Cbo错误类型.Text & ",起始时间:" & Format(.dtpDateStart, "yyyy-MM-dd") & ",终止时间:" & Format(.dtpDateEnd, "yyyy-MM-dd")
    End If
    End With
    '插入重要操作日志
    Call SaveAuditLog(3, "删除", "删除条件为“" & Mid(strNote, 2) & "”的所有日志", strRemarks)
    Call FrmObj.RefreshData
End Sub

Public Function GetFileLineCount(ByVal txtStream As TextStream) As Long
    Do Until txtStream.AtEndOfStream
        txtStream.ReadLine
    Loop
    
    GetFileLineCount = txtStream.Line
End Function

Public Function CopyMenu(cnLink As ADODB.Connection, ByVal lngOldSys As Long, ByVal lngNewSys As Long) As Boolean
    
    On Error GoTo errHandle
    Call OpenCursor(cnLink, "ZLTOOLS.B_Popedom.Copy_menu", lngOldSys, lngNewSys)
    '重新调整序列
    Call AdjustNameSequece("zltools.zlMenus", cnLink)
    CopyMenu = True
    
    Exit Function
errHandle:
    If MsgBox(err.Description, vbRetryCancel, gstrSysName) = vbRetry Then
        Resume
    End If

End Function

Public Function CopyReport(cnLink As ADODB.Connection, ByVal lngOldSys As Long, ByVal lngNewSys As Long) As Boolean
    Call OpenCursor(gcnOracle, "ZLTOOLS.B_Expert.Copy_report", lngOldSys, lngNewSys)
    
    '重新调整序列
    Call AdjustNameSequece("zltools.zlRPTGroups", cnLink)
    Call AdjustNameSequece("zltools.zlReports", cnLink)
    Call AdjustNameSequece("zltools.zlRPTDatas", cnLink)
    Call AdjustNameSequece("zltools.zlRPTItems", cnLink)
    
    CopyReport = True
End Function

Public Function GetOwnerName(lngSys As Long, cnLink As ADODB.Connection) As String
    Dim rsReturn As New ADODB.Recordset
    
    Set rsReturn = OpenCursor(cnLink, "ZLTOOLS.B_Public.Get_Owner_name", lngSys)
    If rsReturn.RecordCount > 0 Then
        GetOwnerName = IIf(IsNull(rsReturn.Fields(0)), "", rsReturn.Fields(0))
    Else
        GetOwnerName = ""
    End If
    
End Function

Public Sub AdjustSequence(ByVal str所有者 As String, cnOwner As ADODB.Connection, Optional ByVal lngSys As Long)
    Dim rsTable As ADODB.Recordset, strTable As String
    
    Set rsTable = GetSequence(str所有者, cnOwner)
    Do Until rsTable.EOF
        strTable = rsTable!Owner & "." & rsTable!Table_Name
        Call AdjustNameSequece(strTable, cnOwner, rsTable!Column_Name)
        DoEvents
        rsTable.MoveNext
    Loop
    If lngSys \ 100 = 1 Then
        Call Adjust结帐ID(cnOwner)
    End If
End Sub

Public Function GetSequence(ByVal str所有者 As String, cnOwner As ADODB.Connection, Optional blnCurrUser As Boolean) As ADODB.Recordset
    Dim rsSeq As New ADODB.Recordset
    Dim strSql As String
    
    '10.26后新安装的系统没有视图"病人费用记录"
    If blnCurrUser Then
        strSql = "Select User, Decode(s.Sequence_Name, '病人费用记录_ID', '病人费用记录', s.Table_Name) as Table_Name, c.Column_Name, s.Sequence_Name" & vbNewLine & _
                "From User_Tab_Columns C," & vbNewLine & _
                "     (Select Sequence_Name, Decode(Sequence_Name, '病人费用记录_ID', '住院费用记录', Table_Name) Table_Name,Column_Name" & vbNewLine & _
                "       From (Select Sequence_Name, Substr(Sequence_Name, 1, Instr(Sequence_Name, '_') - 1) Table_Name," & vbNewLine & _
                "                     Substr(Sequence_Name, Instr(Sequence_Name, '_') + 1) Column_Name" & vbNewLine & _
                "              From User_Sequences)) S" & vbNewLine & _
                "Where c.Table_Name = s.Table_Name And c.Column_Name = s.Column_Name" & vbNewLine & _
                "Order By s.Table_Name"

    Else
        If str所有者 = "" Then
            strSql = " Where Sequence_Owner In (Select 'ZLTOOLS' From Dual Union Select 所有者 From Zlsystems)"
        Else
            strSql = " Where Sequence_Owner = '" & str所有者 & "'"
        End If
        
        strSql = "Select c.Owner, Decode(s.Sequence_Name, '病人费用记录_ID', '病人费用记录', s.Table_Name) as Table_Name, c.Column_Name, s.Sequence_Name" & vbNewLine & _
                "From All_Tab_Columns C," & vbNewLine & _
                "     (Select Sequence_Name, Sequence_Owner, Decode(Sequence_Name, '病人费用记录_ID', '住院费用记录', Table_Name) Table_Name, Column_Name" & vbNewLine & _
                "       From (Select Sequence_Name, Sequence_Owner, Substr(Sequence_Name, 1, Instr(Sequence_Name, '_') - 1) Table_Name," & vbNewLine & _
                "                     Substr(Sequence_Name, Instr(Sequence_Name, '_') + 1) Column_Name" & vbNewLine & _
                "              From All_Sequences" & strSql & ")) S" & vbNewLine & _
                "Where c.Table_Name = s.Table_Name And c.Column_Name = s.Column_Name And c.Owner = s.Sequence_Owner" & vbNewLine & _
                "Order By c.Owner, s.Table_Name"
    End If
    rsSeq.CursorLocation = adUseClient
    rsSeq.Open strSql, cnOwner, adOpenStatic, adLockReadOnly
    Set GetSequence = rsSeq
End Function

Public Function AdjustNameSequece(ByVal strTable As String, cnOwner As ADODB.Connection, Optional ByVal strColumn As String = "ID", Optional ByVal blnJustGetSQL As Boolean) As String
'功能：整理序列的当前号码
'参数：strTable=要调整的表名,注意要使用"user.table"这样的完整名称
'      strColumn 表中对应的序列字段名,一般为ID,可以为指定的其他字段
'      blnJustGetSQL=仅只获取SQL
    Dim dblTableID As Double
    Dim dblSequenceID As Double
    Dim lngIncrement As Long   '保存以前的增量
    Dim rsVal As New ADODB.Recordset, strSql As String, strTab As String
    Dim strReturn As String
    
    strTable = UCase(strTable)
    strTab = Mid(strTable, InStr(strTable, ".") + 1)
    dblTableID = 0
    dblSequenceID = 0
    If strTab = "门诊费用记录" Or strTab = "住院费用记录" Or strTab = "病人费用记录" Then
        strTab = Replace(strTable, strTab, "")    '所有者.
        strSql = "Select Max(MID) as MaxID From (" & _
                "Select Max(" & strColumn & ") as MID From " & strTab & "门诊费用记录 " & _
                "Union All Select Max(" & strColumn & ") as MID From " & strTab & "住院费用记录)"
    Else
        strSql = "Select Max(" & strColumn & ") as MaxID From " & strTable
    End If
    
    rsVal.CursorLocation = adUseClient
    rsVal.Open strSql, cnOwner, adOpenKeyset
    If Not rsVal.EOF Then
        If Not IsNull(rsVal!MAXID) Then
            dblTableID = CDbl(rsVal!MAXID)
        Else
            dblTableID = 0
        End If
    End If
    rsVal.Close
    
    rsVal.Open "Select " & strTable & "_" & strColumn & ".Nextval AS NextID From Dual", cnOwner, adOpenKeyset
    If Not IsNull(rsVal!NEXTID) Then
        dblSequenceID = CDbl(rsVal!NEXTID)
    Else
        dblSequenceID = 0
    End If
    rsVal.Close
    
    If dblTableID - dblSequenceID > 0 Then
        '修改增量
        rsVal.Open "Select Increment_By From All_Sequences Where Sequence_Owner = '" & Split(strTable, ".")(0) & "' And Sequence_Name ='" & Split(strTable, ".")(1) & "_" & strColumn & "'"
        If Not rsVal.EOF Then
            lngIncrement = Nvl(rsVal!Increment_By, 1)
        Else
            lngIncrement = 1
        End If
        rsVal.Close
        strSql = "Alter Sequence " & strTable & "_" & strColumn & " Increment by " & dblTableID - dblSequenceID
        If blnJustGetSQL Then
            strReturn = "--修改增量" & vbNewLine & strSql & ";"
        Else
            cnOwner.Execute strSql
        End If
        strSql = "Select " & strTable & "_" & strColumn & ".Nextval as NextID From Dual"
        If blnJustGetSQL Then
            strReturn = strReturn & vbNewLine & "--移动一次序列" & vbNewLine & strSql & ";"
        Else
            rsVal.Open strSql, cnOwner, adOpenKeyset
        End If
        If Not IsNull(rsVal!NEXTID) Then
            dblSequenceID = CDbl(rsVal!NEXTID)
        Else
            dblSequenceID = 0
        End If
        rsVal.Close
        '还原增量
        cnOwner.Execute "Alter Sequence " & strTable & "_" & strColumn & " Increment by " & lngIncrement
    End If
End Function

Public Sub Adjust结帐ID(cnOwner As ADODB.Connection)
'----------------------------------------------
'功能：针对结帐ID对病人结帐记录_ID进行特殊处理
'----------------------------------------------
    Dim dblTableID As Double, dblTmp As Double
    Dim dblSequenceID As Double
    Dim lngIncrement As Long   '保存以前的增量
    Dim rsVal As New ADODB.Recordset
    dblTableID = 0
    dblSequenceID = 0
    On Error Resume Next
    rsVal.Open "select max(结帐ID) as MAXID from 病人预交记录", cnOwner, adOpenStatic, adLockReadOnly
    If err <> 0 Then
        '可能该系统根本没有这些表
        err.Clear
        Exit Sub
    End If
    
    If Not rsVal.EOF Then
        If Not IsNull(rsVal!MAXID) Then
            dblTableID = CDbl(rsVal!MAXID)
        Else
            dblTableID = 0
        End If
    End If
    rsVal.Close
    rsVal.Open "select max(结帐ID) as MAXID from 门诊费用记录", cnOwner, adOpenStatic, adLockReadOnly
    If Not rsVal.EOF Then
        If Not IsNull(rsVal!MAXID) Then
            dblTmp = CDbl(rsVal!MAXID)
        Else
            dblTmp = 0
        End If
        If dblTmp > dblTableID Then
            dblTableID = dblTmp
        End If
    End If
    rsVal.Close
    
    
    rsVal.Open "select 病人结帐记录_ID.nextval AS NEXTID from dual", cnOwner, adOpenStatic, adLockReadOnly
    If Not IsNull(rsVal!NEXTID) Then
        dblSequenceID = CDbl(rsVal!NEXTID)
    Else
        dblSequenceID = 0
    End If
    
    rsVal.Close
    
    If dblTableID - dblSequenceID > 0 Then
        '修改增量
        rsVal.Open "select INCREMENT_BY from user_sequences where SEQUENCE_NAME = '病人结帐记录_ID'"
        If Not rsVal.EOF Then
            lngIncrement = IIf(IsNull(rsVal("INCREMENT_BY")), 1, rsVal("INCREMENT_BY"))
        Else
            lngIncrement = 1
        End If
        rsVal.Close
        
        cnOwner.Execute "alter sequence 病人结帐记录_ID increment by " & (dblTableID - dblSequenceID)
        
        rsVal.Open "select 病人结帐记录_ID.nextval AS NEXTID from dual", cnOwner, adOpenStatic, adLockReadOnly
        If Not IsNull(rsVal!NEXTID) Then
            dblSequenceID = CDbl(rsVal!NEXTID)
        Else
            dblSequenceID = 0
        End If
        rsVal.Close
'        cnOwner.Execute "select " & strTable & "_ID.nextval from dual"
        '还原增量
        cnOwner.Execute "alter sequence 病人结帐记录_ID increment by " & lngIncrement
    End If
End Sub

Public Sub ApplyOEM(objStatus As Object)
'针对状态栏应用OEM策略
    Dim strOEM As String
    Dim strTmp As String
    On Error Resume Next
    
    If objStatus.Panels(1).Bevel = sbrRaised Then
         strTmp = gobjRegister.zlRegInfo("产品简名")
         If strTmp <> "-" Then
             objStatus.Panels(1).Text = strTmp & "软件"
             If gobjFile Is Nothing Then Set gobjFile = New FileSystemObject
             If gstrAppsoft = "" Then
                 gstrAppsoft = App.Path
                 If gblnInIDE Then
                     gstrAppsoft = "C:\APPSOFT"
                 End If
             End If
             
             If gobjFile.FileExists(gstrAppsoft & "\附加文件\logo_app.jpg") Then
                  Set objStatus.Panels(1).Picture = LoadPicture(gstrAppsoft & "\附加文件\logo_app.jpg")
             Else
                 '处理状态栏图标的OEM策略
                 If strTmp = "中联" Then
                     If gobjRegister.zlRegInfo("授权性质") <> "1" Then
                         Set objStatus.Panels(1).Picture = LoadCustomPicture("Try")
                     Else
                         Set objStatus.Panels(1).Picture = LoadCustomPicture("Logo")
                     End If
                 Else
                     strOEM = GetOEM(strTmp)
                     Set objStatus.Panels(1).Picture = LoadCustomPicture(strOEM)
                     If err <> 0 Then
                         err.Clear
                         Set objStatus.Panels(1).Picture = LoadCustomPicture("Logo")
                     End If
                 End If
             End If
    
             If gobjRegister.zlRegInfo("授权性质") <> "1" Then
                 If strTmp = "中联" Then
                     objStatus.Panels(1).Text = ""
                 Else
                     objStatus.Panels(1).Text = strTmp & "(试用)"
                 End If
             End If
             objStatus.Panels(1).ToolTipText = ""
             objStatus.Height = 360
         End If
     End If
End Sub

Public Sub ApplyOEM_Picture(objPicture As Object, ByVal str属性 As String, Optional ByVal strProductName As String)
'针对各种图标应用OEM策略
    Dim strOEM As String
    Dim blnCorp As Boolean
    On Error Resume Next
    
    If strProductName = "" Then
        strProductName = GetSetting("ZLSOFT", "注册信息", "产品名称", "")
    End If

    If strProductName <> "中联" And strProductName <> "-" Then
        '处理状态栏图标的OEM策略
        If Right(str属性, 1) = "B" Then
            '表示产品图片
            blnCorp = False
            str属性 = Mid(str属性, 1, Len(str属性) - 1)
        Else
            '表示公司徽标
            blnCorp = True
        End If
        
        strOEM = GetOEM(strProductName, blnCorp)
        If str属性 = "Picture" Then
            Set objPicture.Picture = LoadCustomPicture(strOEM)
        ElseIf str属性 = "Icon" Then
            Set objPicture.Icon = LoadCustomPicture(strOEM)
        End If
        
        If err <> 0 Then
            err.Clear
        End If
    
    End If
End Sub

Public Function LoadCustomPicture(strID As String) As StdPicture
'功能:将资源文件中的指定资源生成磁盘文件
'参数:ID=资源号,strExt=要生成文件的扩展名(如BMP)
'返回:生成文件名
    Dim arrData() As Byte
    Dim intFile As Integer
    Dim strFile As String * 255, strR As String
    
    arrData = LoadResData(strID, "CUSTOM")
    intFile = FreeFile
    
    GetTempPath 255, strFile
    strR = Trim(Left(strFile, InStr(strFile, Chr(0)) - 1)) & CLng(Timer * 100) & ".pic"

    Open strR For Binary As intFile
    Put intFile, , arrData()
    Close intFile
    Set LoadCustomPicture = VB.LoadPicture(strR)
    Kill strR
End Function

Public Function GetOEM(ByVal strAsk As String, Optional ByVal blnCorp As Boolean = True) As String
    '-------------------------------------------------------------
    '功能：返回每个字线的ASCII码
    '参数：
    '返回：
    '-------------------------------------------------------------
    Dim intBit As Integer, iCount As Integer, blnCan As Boolean
    Dim strCode As String
    
    'OEM图片有两种类型 ，一是指公司徽标，另一个是产品标识
    strCode = IIf(blnCorp = True, "OEM_", "PIC_")
    For intBit = 1 To Len(strAsk)
        '取每个字的ASCII码
        strCode = strCode & Hex(Asc(Mid(strAsk, intBit, 1)))
    Next
    GetOEM = strCode
End Function


Public Sub ReCompileProcedure(ByVal cnOwner As ADODB.Connection)
'对本用户下所有已经失效的过程进行重新编译
    Dim rsTemp As New ADODB.Recordset
    Dim lngTime As Long
    
    For lngTime = 1 To 3
        '最多调用三次，因为有些过程是相互调用，一次编译不能解决问题
        '为了快速得到列表，不利用对象之间的引用关系
        If rsTemp.State = adStateOpen Then rsTemp.Close
        
        gstrSQL = "select OBJECT_NAME from user_objects where object_type='PROCEDURE' and STATUS='INVALID'"
        rsTemp.Open gstrSQL, cnOwner, adOpenStatic, adLockReadOnly
        
        On Error Resume Next
        If rsTemp.RecordCount = 0 Then
            '没有过程失效，直接退出
            Exit Sub
        Else
            Do Until rsTemp.EOF
                '有可能出错
                gstrSQL = "alter procedure " & rsTemp("OBJECT_NAME") & " compile"
                cnOwner.Execute gstrSQL
                rsTemp.MoveNext
            Loop
        End If
    Next
End Sub

Public Function LoadServer(ByRef strFileInfo As String) As Collection
'功能：读出本地的服务器列表
    Dim strPath As String, strFile As String, lngFile As Integer
    Dim strLine As String, lngPos As Long
    Dim strServer As String, strComputer As String, strSID As String
    Dim arrTmp As Variant
    Dim rsOraHome As ADODB.Recordset
    Dim intVersion As Integer, intTimes As Integer, intServer As Integer
    Dim i As Long
    Dim colServer As New Collection

    Set rsOraHome = New ADODB.Recordset
    With rsOraHome
        .Fields.Append "Name", adVarChar, 256 'Name
        .Fields.Append "VerSion", adInteger  '版本
        .Fields.Append "Times", adInteger '第几次安装
        .Fields.Append "Server", adInteger '1-服务器,2-客户端
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
        '1:读取64位下32目录会自动定位到SOFTWARE\Wow6432Node\Oracle 2：读取32位下32位目录
        arrTmp = GetAllSubKey(HKEY_LOCAL_MACHINE, "SOFTWARE\Oracle")
        If TypeName(arrTmp) = "Empty" Then
            If Is64bit Then
                strFileInfo = "没有找到注册表项HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Oracle！"
            Else
                strFileInfo = "没有找到注册表项HKEY_LOCAL_MACHINE\SOFTWARE\Oracle！"
            End If
        Else
            For i = LBound(arrTmp) To UBound(arrTmp)
                If UCase(arrTmp(i)) Like "KEY_ORA*HOME*" Then
                    intVersion = 0: intTimes = 0:  intServer = 1
                    If GetOraInfoByRegKey(arrTmp(i), intVersion, intTimes, intServer) Then
                        .AddNew Array("Name", "VerSion", "Times", "Server"), Array("\" & arrTmp(i), intVersion, intTimes, intServer)
                        .Update
                    End If
                End If
            Next
            If UBound(arrTmp) <> -1 Then ''顶级目录可能有Oracle_Home信息，默认读取这个
                .AddNew Array("Name", "VerSion", "Times", "Server"), Array("", 0, 0, 1): .Update
            End If
            .Sort = "VerSion Desc,Times Desc,Server"
            Do While Not .EOF
                strPath = GetKeyValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Oracle" & !name, "ORACLE_HOME")
                If strPath = "" And !name & "" = "" Then
                    strPath = GetKeyValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Oracle", "ORA_CRS_HOME")
                End If
                If strPath <> "" Then
                    strFile = strPath & "\network\ADMIN\tnsnames.ora" 'Oracle 8i以上
                    If gobjFile.FileExists(strFile) Then Exit Do
                    strFile = strPath & "\NET80\ADMIN\tnsnames.ora" 'Oracle 8
                    If gobjFile.FileExists(strFile) Then Exit Do
                End If
                strFile = ""
                .MoveNext
            Loop
        End If
    End With
    If strFile = "" Then Exit Function
    strFileInfo = "服务器列表来源:" & strFile
    lngFile = FreeFile()
    Open strFile For Input Access Read As lngFile
    Do Until EOF(lngFile)
        Input #lngFile, strLine
        strLine = Trim(strLine)
        If strLine <> "" And Left(strLine, 1) <> "#" Then
            '非注释行或空行
            If InStr(strLine, "(") = 0 And InStr(strLine, ")") = 0 Then
                '该行的内容就是服务器名了，把所有内容都初始化
                strServer = Trim(Mid(strLine, 1, InStr(strLine, "=") - 1))
                strComputer = ""
                strSID = ""
            ElseIf InStr(strLine, "(ADDRESS") > 0 Then
                '该行的内容是主机名
                If InStr(strLine, "PROTOCOL = TCP") > 0 And InStr(strLine, "PORT = ") > 0 Then
                    '符合我们的程序要求
                    strComputer = Mid(strLine, InStr(strLine, "HOST =") + Len("HOST ="))
                    strComputer = Trim(Mid(strComputer, 1, InStr(strComputer, ")") - 1))
                End If
            Else
                lngPos = InStr(strLine, "(SID")
                If lngPos = 0 Then
                    lngPos = InStr(strLine, "(SERVICE_NAME")
                End If
                
                If lngPos > 0 Then
                    '该行的内容是实例名
                    strSID = Mid(strLine, InStr(lngPos, strLine, "=") + 1)
                    strSID = Trim(Mid(strSID, 1, InStr(strSID, ")") - 1))
                    
                    If strServer <> "" And strComputer <> "" And strSID <> "" Then
                        '已经得到所有需要的内容
                        colServer.Add Array(strServer, strComputer, strSID)
                    End If
                End If
            End If
        End If
    Loop
    Close #lngFile
    
    Set LoadServer = colServer
End Function

Private Function GetOraInfoByRegKey(ByVal strOraHome As String, ByRef intVer As Integer, ByRef intTimes As Integer, ByRef intServer As Integer) As Boolean
'功能:通过OracleHome键获取Oracle信息
    Dim arrTmp As Variant
    Dim i As Long, blnRetrun As Boolean
    'KEY_OraDb11g_home1_32bit
    'Key_Ora*版本Home_32Bit
    'Key_Ora*版本_Home*
    arrTmp = Split(UCase(strOraHome), "_")
    For i = 1 To UBound(arrTmp)
        If arrTmp(i) Like "HOME*" Then
            intTimes = ValEx(arrTmp(2))
            blnRetrun = True
        ElseIf arrTmp(i) Like "*HOME*" Then
            intTimes = Val(Mid(arrTmp(1), InStr(UCase(arrTmp(1)), "HOME") + 4))
            blnRetrun = True
        End If
        If arrTmp(i) Like "ORADB*" Then
            intVer = ValEx(Mid(arrTmp(1), 6))
            intServer = 1
            blnRetrun = True
        ElseIf arrTmp(i) Like "ORACLIENT*" Then
            intVer = ValEx(Mid(arrTmp(1), 10))
            intServer = 2
            blnRetrun = True
        ElseIf arrTmp(i) Like "*CLIENT*" Then
            intServer = 2
            intVer = ValEx(arrTmp(i))
            blnRetrun = True
        End If
    Next
    GetOraInfoByRegKey = blnRetrun
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


Public Function ExpandEnvStr(ByVal strInput As String) As String
'功能：将字符串中的环境变量替换为常规值
'         strInput=包含环境变量的字符串
'返回：用实际的值替换字符串中的环境变量后的字符串
    '// 如： %PATH% 则返回 "c:\;c:\windows;"
    Dim lngLen As Long, strBuf As String, strOld As String
    strOld = strInput & "  " ' 不知为什么要加两个字符，否则返回值会少最后两个字符！
    strBuf = "" '// 不支持Windows 95
    '// get the length
    lngLen = ExpandEnvironmentStrings(strOld, strBuf, lngLen)
    '// 展开字符串
    strBuf = String$(lngLen - 1, Chr$(0))
    lngLen = ExpandEnvironmentStrings(strOld, strBuf, LenB(strBuf))
    '// 返回环境变量
    ExpandEnvStr = TruncZero(strBuf)
End Function

Public Function ValEx(ByVal varInput As Variant) As Variant
'功能：由于Val只能以数字开头识别，ValEx以第一个数字进行识别
    Dim arrTmp As Variant, lngPos As Long
    If Val(varInput) = 0 Then
        varInput = varInput & ""
        If Trim(varInput) = "" Then ValEx = 0: Exit Function
        For lngPos = 1 To Len(varInput)
            If IsNumeric(Mid(varInput, lngPos, 1)) Then Exit For
        Next
        If lngPos = Len(varInput) + 1 Then
            ValEx = 0
        Else
            ValEx = Val(Mid(varInput, lngPos))
        End If
    Else
        ValEx = Val(varInput)
    End If
End Function

 Public Function Is64bit() As Boolean
    '******************************************************************************************************************
    '功能：是否是64位系统
    '返回：
    '******************************************************************************************************************
    Dim handle As Long
    Dim bolFunc As Boolean
        
    bolFunc = False
    handle = GetProcAddress(GetModuleHandle("kernel32"), "IsWow64Process")
    If handle > 0 Then
        IsWow64Process GetCurrentProcess(), bolFunc
    End If
    Is64bit = bolFunc
End Function

Public Function GetAllSubKey(ByVal KeyRoot As Long, KeyName As String) As Variant
'功能:获取某项的所有子项
'返回：=子项数组
    Dim lnghKey As Long, lngRet As Long, strName As String, lngIdx As Long
    Dim strSubKey As Variant
    strSubKey = Array()
    lngIdx = 0: strName = String(256, Chr(0))
    lngRet = RegOpenKey(KeyRoot, KeyName, lnghKey)
    If lngRet = 0 Then
        Do
            lngRet = RegEnumKey(lnghKey, lngIdx, strName, Len(strName))
            If lngRet = 0 Then
                ReDim Preserve strSubKey(UBound(strSubKey) + 1)
                strSubKey(UBound(strSubKey)) = Left(strName, InStr(strName, Chr(0)) - 1)
                lngIdx = lngIdx + 1
            End If
        Loop Until lngRet <> 0
    End If
    RegCloseKey lnghKey
    GetAllSubKey = strSubKey
End Function

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String) As String
'功能：读注册表
    Dim i As Long                                           ' 循环计数器
    Dim rc As Long                                          ' 返回代码
    Dim hKey As Long                                        ' 处理打开的注册表关键字
    Dim hDepth As Long                                      '
    Dim sKeyVal As String
    Dim lKeyValType As Long                                 ' 注册表关键字数据类型
    Dim tmpVal As String                                    ' 注册表关键字的临时存储器
    Dim KeyValSize As Long                                  ' 注册表关键字变量尺寸
    
    ' 在 KeyRoot {HKEY_LOCAL_MACHINE...} 下打开注册表关键字
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' 打开注册表关键字
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' 处理错误...
    
    tmpVal = String$(1024, 0)                             ' 分配变量空间
    KeyValSize = 1024                                       ' 标记变量尺寸
    
    '------------------------------------------------------------
    ' 检索注册表关键字的值...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         lKeyValType, tmpVal, KeyValSize)    ' 获得/创建关键字的值
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' 错误处理
      
    tmpVal = Left$(tmpVal, InStr(tmpVal, Chr(0)) - 1)

    '------------------------------------------------------------
    ' 决定关键字值的转换类型...
    '------------------------------------------------------------
    Select Case lKeyValType                                  ' 搜索数据类型...
    Case REG_SZ, REG_EXPAND_SZ                              ' 字符串注册表关键字数据类型
        sKeyVal = tmpVal                                     ' 复制字符串的值
    Case REG_DWORD                                          ' 四字节注册表关键字数据类型
        For i = Len(tmpVal) To 1 Step -1                    ' 转换每一位
            sKeyVal = sKeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' 一个字符一个字符地生成值。
        Next
        sKeyVal = Format$("&h" + sKeyVal)                     ' 转换四字节为字符串
    End Select
    
    GetKeyValue = sKeyVal                                   ' 返回值
    rc = RegCloseKey(hKey)                                  ' 关闭注册表关键字
    Exit Function                                           ' 退出
    
GetKeyError:    ' 错误发生过后进行清除...
    GetKeyValue = vbNullString                              ' 设置返回值为空字符串
    rc = RegCloseKey(hKey)                                  ' 关闭注册表关键字
End Function

Private Function GetKeyValueInfo(ByVal strKey As String, Optional ByVal strValueName As String, Optional ByRef hRootKey As REGRoot, Optional ByRef strSubKey As String, Optional ByRef lngType As Long) As Boolean
'功能：根据键位获取根键值与子健,以及值类型
'参数：strKey=注册表键位，如“HKEY_CURRENT_USER\Printers\DevModePerUser"
'          strValueName=变量名
'出参：
'          hRootKey=根键
'          strSubKey=子健
'          lngType=键类型
'返回：是否获取成功
    Dim strRoot As String, lngPos As String, hKey As Long
    Dim lngReturn As Long, strName As String * 255
    
    On Error GoTo errH
    hRootKey = 0: strSubKey = "": lngType = 0
    lngPos = InStr(strKey, "\")
    If lngPos = 0 Then Exit Function
    strRoot = Mid(strKey, 1, lngPos - 1)
    strSubKey = Mid(strKey, lngPos + 1)
    
    hRootKey = Decode(UCase(strRoot), "HKEY_CLASSES_ROOT", HKEY_CLASSES_ROOT, _
                                                                         "HKEY_CURRENT_USER", HKEY_CURRENT_USER, _
                                                                         "HKEY_LOCAL_MACHINE", HKEY_LOCAL_MACHINE, _
                                                                         "HKEY_USERS", HKEY_USERS, _
                                                                         "HKEY_PERFORMANCE_DATA", HKEY_PERFORMANCE_DATA, _
                                                                         "HKEY_CURRENT_CONFIG", HKEY_CURRENT_CONFIG, _
                                                                         "HKEY_DYN_DATA", HKEY_DYN_DATA, 0)
    If hRootKey = 0 Then Exit Function
    If lngType <> -1 Then
        '使用查询方式打开，进行键名类型查询
        lngReturn = RegOpenKeyEx(hRootKey, strSubKey, 0, KEY_QUERY_VALUE, hKey)
        If lngReturn <> ERROR_SUCCESS Then
            Exit Function
        End If
        If strValueName <> "" Then
            lngReturn = RegQueryValueEx_ValueType(hKey, strValueName, ByVal 0&, lngType, ByVal strName, Len(strName))
            '可能字段超长，长度不够，所以出错不退出
            'If lngReturn <> ERROR_SUCCESS Then: RegCloseKey (hKey): Exit Function
        End If
        RegCloseKey (hKey)
    End If
    GetKeyValueInfo = True
    Exit Function
errH:
    If 0 = 1 Then
        Resume
    End If
    err.Clear
End Function

Public Function GetRegValue(ByVal strKey As String, ByVal strValueName As String, ByRef varValue As Variant, Optional blnOneString As Boolean = False) As Boolean
'功能：获取注册表中指定位置的值
'参数：strKey=注册表键位，如“HKEY_CURRENT_USER\Printers\DevModePerUser"
'          strValueName=变量名
'          strValue=变量值
'          strValueType=变量类型，默认为字符串
'           blnOneString = 对REG_EXPAND_SZ、REG_MULTI_SZ,REG_BINARY有效。-  True 则函数返回单一字符串，且不经任何处理，只去掉字符串尾！
'返回：是否读取成功
'说明：当前只对REG_SZ, REG_EXPAND_SZ, REG_MULTI_SZ，REG_DWORD，REG_BINARY实现了读取。没有查询到可以自动查找键名
    Dim hRootKey As REGRoot, strSubKey As String
    Dim lngReturn As Long
    Dim lngKey As Long, ruType As REGValueType
    Dim lngLength As Long, varBufData As Variant, strBufVar() As String, lngBuf As Long, bytBuf() As Byte, strBuf As String
    Dim i As Long, strReturn As String, strTmp As String
    '不是有效的注册表键位,获取键名类型
    If Not GetKeyValueInfo(strKey, strValueName, hRootKey, strSubKey, ruType) Then Exit Function
    '打开变量
    lngReturn = RegOpenKeyEx(hRootKey, strSubKey, 0, KEY_QUERY_VALUE, lngKey)
    If lngReturn <> ERROR_SUCCESS Then
        Exit Function
    End If
    On Error GoTo errH
    Select Case ruType
        Case REG_SZ, REG_EXPAND_SZ, REG_MULTI_SZ '字符串类型读取
'            lngReturn = RegQueryValueEx(lngKey, strValueName, 0, ruType, 0, lngLength)
'            If lngReturn <> ERROR_SUCCESS Then Err.Clear '可能出错，因此这样处理
            lngLength = 1024: strBuf = Space(lngLength)
            lngReturn = RegQueryValueEx_String(lngKey, strValueName, 0, ruType, strBuf, lngLength)
            If lngReturn <> ERROR_SUCCESS Then: RegCloseKey (lngKey): Exit Function
            Select Case ruType
                Case REG_SZ
                    varValue = TruncZero(strBuf)
                Case REG_EXPAND_SZ ' 扩充环境字符串，查询环境变量和返回定义值
                    If Not blnOneString Then
                        varValue = TruncZero(ExpandEnvStr(TruncZero(strBuf)))
                    Else
                        varValue = TruncZero(strBuf)
                    End If
                Case REG_MULTI_SZ ' 多行字符串
                    If Not blnOneString Then
                        If Len(strBuf) <> 0 Then ' 读到的是非空字符串，可以分割。
                            strBufVar = Split(Left$(strBuf, Len(strBuf) - 1), Chr$(0))
                        Else ' 若是空字符串，要定义S(0) ，否则出错！
                            ReDim strBufVar(0) As String
                        End If
                        ' 函数返回值，返回一个字符串数组？！
                        varValue = strBufVar()
                    Else
                        varValue = TruncZero(strBuf)
                    End If
            End Select
        Case REG_DWORD
            lngReturn = RegQueryValueEx_Long(lngKey, strValueName, ByVal 0&, ruType, lngBuf, Len(lngBuf))
            If lngReturn <> ERROR_SUCCESS Then: RegCloseKey (lngKey): varValue = 0: Exit Function
            varValue = lngBuf
        Case REG_BINARY
            lngReturn = RegQueryValueEx_BINARY(lngKey, strValueName, 0, ruType, ByVal 0, lngLength)
            If lngReturn <> ERROR_SUCCESS Then
                RegCloseKey lngKey: Exit Function
                If blnOneString Then
                    varValue = "00"
                Else
                    ReDim bytBuf(0)
                    varValue = bytBuf()
                End If
            End If
            ReDim bytBuf(lngLength - 1)
            lngReturn = RegQueryValueEx_BINARY(lngKey, strValueName, 0, ruType, bytBuf(0), lngLength)
            If lngReturn <> ERROR_SUCCESS Then
                RegCloseKey lngKey: Exit Function
                If blnOneString Then
                    varValue = "00"
                Else
                    ReDim bytBuf(0)
                    varValue = bytBuf()
                End If
            End If
            If lngLength <> UBound(bytBuf) + 1 Then
               ReDim Preserve bytBuf(0 To lngLength - 1) As Byte
            End If
            ' 返回字符串，注意：要将字节数组进行转化！
            If blnOneString Then
                '循环数据，把字节转换为16进制字符串
                For i = LBound(bytBuf) To UBound(bytBuf)
                   strTmp = CStr(Hex(bytBuf(i)))
                   If (Len(strTmp) = 1) Then strTmp = "0" & strTmp
                   strReturn = strReturn & " " & strTmp
                Next i
                varValue = Trim$(strReturn)
            Else
                varValue = bytBuf()
            End If
    End Select
    RegCloseKey lngKey
    GetRegValue = True
    Exit Function
errH:
    If 0 = 1 Then
        Resume
    End If
End Function

Public Function SetRegValue(ByVal strKey As String, ByVal strValueName As String, varValue As Variant) As Boolean
'功能：设置注册表中指定位置的值
'参数：strKey=注册表键位，如“HKEY_CURRENT_USER\Printers\DevModePerUser"
'          strValueName=变量名
'          strValue=变量值
'          strValueType=变量类型，默认为字符串
'返回：是否设置成功
    Dim hRootKey As REGRoot, strSubKey As String
    Dim lngReturn As Long
    Dim lngKey As Long, ruType As REGValueType
    Dim lngLength As Long, varBufData As Variant, strBufVar() As String, lngBuf As Long, bytBuf() As Byte, strBuf As String
    Dim i As Long, lb As Long, ub As Long, strReturn As String, strTmp As String
    '不是有效的注册表键位,获取键名类型
    If Not GetKeyValueInfo(strKey, strValueName, hRootKey, strSubKey, ruType) Then Exit Function
    '打开变量
    lngReturn = RegOpenKeyEx(hRootKey, strSubKey, 0, KEY_SET_VALUE, lngKey)
    If lngReturn <> ERROR_SUCCESS Then
        Exit Function
    End If
    On Error GoTo errH
    Select Case ruType
        Case REG_SZ, REG_EXPAND_SZ, REG_MULTI_SZ
            If ruType = REG_MULTI_SZ And varType(varValue) = vbArray + vbString Then 'string数组，则将数组合成字符串
                lngLength = UBound(varValue) - LBound(varValue) + 1
                For i = LBound(varValue) To UBound(varValue)
                    strBuf = strBuf & varValue(i) & Chr$(0)
                Next
                strBuf = TruncZero(strBuf)
                lngLength = ActualLen(strBuf)
            Else
                strBuf = TruncZero(varValue)
                lngLength = ActualLen(strBuf)
            End If
            lngReturn = RegSetValueEx_String(lngKey, strValueName, ByVal 0&, ruType, ByVal strBuf, lngLength)
            If lngReturn <> 0 Then RegCloseKey lngKey: Exit Function
        Case REG_DWORD
            lngBuf = Val(varValue): lngLength = Len(lngBuf)
            lngReturn = RegSetValueEx_Long(lngKey, strValueName, ByVal 0&, ruType, lngBuf, lngLength)
            If lngReturn <> 0 Then RegCloseKey lngKey: Exit Function
        Case REG_BINARY
            ' 1、varValue ＝ 字节数组，如 B()
            If varType(varValue) = vbArray + vbByte Then
                Dim binValue() As Byte, Length As Long
                bytBuf = varValue
                lngLength = UBound(bytBuf) - LBound(bytBuf) + 1
                lngReturn = RegSetValueEx_BINARY(lngKey, strValueName, 0, ruType, bytBuf(0), lngLength)
                If lngReturn <> 0 Then RegCloseKey lngKey: Exit Function
            ' 2、varValue ＝ 整型或长整型，如 520
            ElseIf varType(varValue) = vbLong Or varType(varValue) = vbInteger Then
                lngBuf = Val(varValue): lngLength = Len(lngBuf)
                lngReturn = RegSetValueEx_Long(lngKey, strValueName, 0, ruType, lngBuf, lngLength)
                If lngReturn <> 0 Then RegCloseKey lngKey: Exit Function
            ' 3、varValue ＝字符串，如 "BE 3E FF AB"
            ElseIf varType(varValue) = vbString Then
                ' 转化数据
                Dim ByteArray() As Byte
                Dim tmpArray() As String '//转换ASCII字符到16进制字节
                strTmp = varValue
                ' 以空格分割字符串
                strBufVar = Split(strTmp, " ")
                lb = LBound(strBufVar): ub = UBound(strBufVar)
                ' 为动态数组分配空间
                ReDim bytBuf(lb To ub)
                ' 循环转换
                For i = lb To ub - 1
                    bytBuf(i) = CByte(Val("&H" & Right$(strBufVar(i), 2)))
                Next i
                ' 注意：最后一个不知道字符串后面多了2个什么，要用 Left$(tmpArray(ub), 2)
                bytBuf(ub) = CByte(Val("&H" & Left$(strBufVar(ub), 2)))
                ' 将数据写入到注册表，注意：最后是 ub - lb + 1
                lngReturn = RegSetValueEx_BINARY(lngKey, strValueName, 0, ruType, bytBuf(0), ub - lb + 1)
                If lngReturn <> 0 Then RegCloseKey lngKey: Exit Function
            End If
    End Select
    RegCloseKey lngKey
    SetRegValue = True
    Exit Function
errH:
    If 0 = 1 Then
        Resume
    End If
End Function

Public Function DeleteRegValue(ByVal strKey As String, ByVal strValueName As String) As Boolean
'功能：删除注册表中指定位置的值
'参数：strKey=注册表键位，如“HKEY_CURRENT_USER\Printers\DevModePerUser"
'          strValueName=变量名
'返回：是否读取成功
    Dim lngLength As Long, lngReturn As Long
    Dim lngKey As Long, lngType As Long
    Dim hRootKey As REGRoot, strSubKey As String
    
    '不是有效的注册表键位
    If Not GetKeyValueInfo(strKey, strValueName, hRootKey, strSubKey, -1) Then Exit Function
    '打开键
    lngReturn = RegOpenKeyEx(hRootKey, strSubKey, 0, KEY_SET_VALUE, lngKey)
    If lngReturn <> 0 Then
        Exit Function
    End If
    '删除键
    lngReturn = RegDeleteValue(lngKey, strValueName)
    If lngReturn = 0 Then
        DeleteRegValue = True
    End If
    '关闭键
    RegCloseKey lngKey
End Function

Public Function CheckSpaceIsUse(ByVal strType As String, ByVal strName As String, ByVal strOwner As String) As Boolean
'功能：检查表空间或数据文件是否由其它用户使用
'参数：strType    表空间 数据文件
'      strName          表空间或数据文件的名字
'      strOwner         以区别其它用户的所有者名
    Dim rsTemp As New ADODB.Recordset
    
    If strType = "表空间" Then
        gstrSQL = "select owner from all_tables where tablespace_name='" & UCase(strName) & "' and owner<>'" & UCase(strOwner) & "' AND ROWNUM<2" & vbNewLine & _
                  "union " & vbNewLine & _
                  "select owner from all_indexes where tablespace_name='" & UCase(strName) & "' and owner<>'" & UCase(strOwner) & "' AND ROWNUM<2"
        
    Else
        gstrSQL = "select O.owner  from V$TABLESPACE T,V$DATAFILE F,all_tables O " & _
                  "Where T.TS# = F.TS# And T.name = O.TABLESPACE_NAME " & _
                  "    and F.name='" & UCase(strName) & "' and O.owner<>'" & UCase(strOwner) & "' AND ROWNUM<2 "
    End If
    
    On Error Resume Next
    rsTemp.Open gstrSQL, gcnOracle, , adLockReadOnly
    
    If rsTemp.RecordCount <= 0 Then
        '没有其他用户使用，可以删除
        Exit Function
    End If
    '有用户使用
    CheckSpaceIsUse = True
End Function

Public Function LvwSelectColumns(objSet As Object, ByVal strColumn As String, Optional ByVal blnInit As Boolean = False) As Boolean
'功能:对列表控件的列进行设置
'参数:
'   objSet：要设置的对象,目前只支持ListView，以后再加上FlexGrid,DataGrid。
'   strColumn；列串。格式是"列名,列宽,对齐数值,列特性;列名,列宽,对齐数值,列特性"    注意列之间是用分号
'      比如 "名称,2000,0,1;编码,800,0,0;简码,1440,0,0"
'      对ListView而言：列特性为1表示该列不可删除，列特性为0表示该列可以删除
'      对FlexGridView而言：列特性还要表示是否属于固定列，以便不能和其它列进行顺序调整
'   blnInit：True,不显示选择窗口，直接初始化
    Dim varColumns As Variant, varColumn As Variant
    Dim lngCol As Long

    If blnInit Then
        varColumns = Split(strColumn, ";")
        Select Case TypeName(objSet)
            Case "ListView"
                With objSet.ColumnHeaders
                    .Clear
                    For lngCol = LBound(varColumns) To UBound(varColumns)
                        varColumn = Split(varColumns(lngCol), ",")
                        .Add , "_" & varColumn(0), varColumn(0), varColumn(1), varColumn(2)
                    Next
                End With
            Case "MSHFlexGrid"
            Case "DataGrid"
        End Select
    End If
End Function

Public Sub NextLvwPos(lvwObj As Object, ByVal vIndex As Long)
        
    If lvwObj.ListItems.Count > 0 Then
        vIndex = IIf(lvwObj.ListItems.Count > vIndex, vIndex, lvwObj.ListItems.Count)
        lvwObj.ListItems(vIndex).Selected = True
        lvwObj.ListItems(vIndex).EnsureVisible
    End If
End Sub

Public Function Nvl(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'功能：相当于Oracle的NVL，将Null值改成另外一个预设值
    Nvl = IIf(IsNull(varValue), DefaultValue, varValue)
End Function


Public Function OpenCursor(ByVal cnOwner As ADODB.Connection, _
                              ByVal strPackagesName As String, _
                              ParamArray varParValue() As Variant) As ADODB.Recordset
'-----------------------------------------
'功能：调用存储过程返回记录集
'入参：strPackagesName ，格式为 [所有者.]包.过程名
'-----------------------------------------
    Static cmdPackage As New ADODB.Command
    Dim parPackage As ADODB.Parameter
    Dim arrPar As Variant, i As Integer
    Dim varValue As Variant, intMax As Integer
    Dim intMaxArr As Integer  '记录参数个数
    Dim varOutPar As Variant
    On Error GoTo errHandle

    '清除原有参数:不然不能重复执行
   
    
    cmdPackage.CommandText = "" '不为空有时清除参数出错
    Do While cmdPackage.Parameters.Count > 0
        cmdPackage.Parameters.Delete 0
    Loop
    
    '------ IN 参数
    For i = 0 To UBound(varParValue)
        varValue = varParValue(i)
        Select Case TypeName(varValue)
            Case "Byte", "Integer", "Long", "Single", "Double", "Currency" '数字
                cmdPackage.Parameters.Append cmdPackage.CreateParameter("P" & i, adVarNumeric, adParamInput, 30, varValue)
            Case "String" '字符
                intMax = LenB(StrConv(varValue, vbFromUnicode))
                If intMax = 0 Or intMax < 10 Then intMax = 10
                cmdPackage.Parameters.Append cmdPackage.CreateParameter("P" & i, adVarChar, adParamInput, intMax, varValue)
            Case "Date" '日期
                cmdPackage.Parameters.Append cmdPackage.CreateParameter("P" & i, adDBTimeStamp, adParamInput, , varValue)
        End Select
    Next

    If cmdPackage.ActiveConnection Is Nothing Then
        If cnOwner Is Nothing Then
            Set cmdPackage.ActiveConnection = gcnOracle
        Else
            Set cmdPackage.ActiveConnection = cnOwner
        End If
    Else
        If Not cnOwner Is Nothing Then
            If cmdPackage.ActiveConnection.ConnectionString <> cnOwner.ConnectionString Then
                Set cmdPackage.ActiveConnection = cnOwner
            End If
        End If
    End If
    
    cmdPackage.CommandType = adCmdStoredProc
    cmdPackage.CommandText = strPackagesName
    cmdPackage.Properties("PLSQLRSet") = True
    Set OpenCursor = cmdPackage.Execute
    cmdPackage.Properties("PLSQLRSet") = False
    Exit Function
errHandle:
    If MsgBox(err.Description, vbRetryCancel, gstrSysName) = vbRetry Then
        Resume
    End If

End Function

Public Function GetAllSystems(Optional blnAll As Boolean) As ADODB.Recordset
    
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHandle
    If blnAll Then
        gstrSQL = "Select A.编号, A.名称, A.所有者, A.版本号 From Zlsystems A Order By A.编号"
    Else
        gstrSQL = "SELECT a.编号, a.名称, a.所有者, a.版本号 " & _
                 "       FROM Zlsystems a, " & _
                 "            (SELECT Owner " & _
                 "              FROM All_Tables " & _
                 "              WHERE Table_Name IN ('部门表', '人员表', '部门人员', '上机人员表') " & _
                 "              GROUP BY Owner " & _
                 "              HAVING COUNT(Owner) = 4) b " & _
                 "       WHERE a.所有者 = b.Owner " & _
                 "       ORDER BY a.编号"
    End If
    If rsTemp.State = adStateOpen Then rsTemp.Close
    rsTemp.Open gstrSQL, gcnOracle, adOpenForwardOnly, adLockReadOnly
    Set GetAllSystems = rsTemp
    Exit Function
errHandle:
    MsgBox err.Description, vbCritical, gstrSysName
    
End Function


Public Function CheckHistorySpaces(ByVal cnOracle As ADODB.Connection, ByVal pgbProcess As ProgressBar, ByVal strBakOwner As String, ByVal strDbLink As String, _
                 ByVal lngSys As Long, ByVal strSysOwner As String, _
                 Optional bytCheckSys As Byte, Optional ByRef cllErrMsg As Variant) As Boolean
    '--------------------------------------------------------------------------------------------------------------------------
    '功能:检查历史数据空间的相对象是否完整
    '参数:cnOracle-在线数据库连接
    '     strDbLink-连接名
    '     pgbProcess-进度条
    '     lngSys-系统号
    '     strSysOwner-系统所有者
    '     strBakOwner-备份空间的所有者
    '     cllErr:错误信息集(0-对象类型,1-对象名称,2-错语信息,2-问题严重级别说明)
    '     bytCheckSys-0-仅仅检查是否在zlbakInfor表中存在系统(不检查表）,1-仅仅检查表对象,>1表示全检查:主要是检查对象和表
    '出参:strErrMsg-返回相关的错误信息
    '返回:如果检查合法,返回true,否则返回False
    '--------------------------------------------------------------------------------------------------------------------------
    Dim rsBakObject As New ADODB.Recordset, rsObject As New ADODB.Recordset
    Dim cllErr  As Collection, blnBakInfor As Boolean
    Dim strTemp  As String
    Dim rsTemp As New ADODB.Recordset
    Dim lngRow As Long
    
    
    If strDbLink <> "" Then
        '检查连接是否正常
        err = 0: On Error Resume Next
        gstrSQL = "Select 1 from dual@" & strDbLink
        OpenRecordset rsBakObject, gstrSQL, "连接验证", , , cnOracle
        If err <> 0 Then
            cllErr.Add Array("远程连接", strDbLink, "连接不正常", "严重：历史数据空间将不能正常运行")
            Exit Function
        End If
    End If
    
    err = 0: On Error GoTo ErrHand:
    blnBakInfor = True
    '检查是否存在历史数据空间
    gstrSQL = "Select 表名 From zlBakTables where 系统=" & lngSys
    OpenRecordset rsObject, gstrSQL, "检查是否存在历史数据空间", , , cnOracle
    If rsObject.EOF Then
        '不存在历史数据空间，则检查正确。
        CheckHistorySpaces = True
        Exit Function
    End If
    
            
    If strDbLink <> "" Then
        gstrSQL = "select table_name as 表名  from " & strBakOwner & ".user_tables"         '& " where  owner = '" & strBakOwnerName & "' "
    Else
        gstrSQL = "select table_name as 表名  from user_tables@" & strDbLink         '& " where  owner = '" & strBakOwnerName & "' "
    End If
    
    OpenRecordset rsBakObject, gstrSQL, "获取历史空间表", , , cnOracle
    
    'cllErr代表(0-对象类型,1-对象名称,2-错语信息,2-问题严重级别说明)
    
     Set cllErr = New Collection
    '检查zlBakInfo表是否存在
    rsBakObject.Filter = "表名='" & UCase("zlBakInfo") & "'"
    If rsBakObject.EOF Then
        cllErr.Add Array("表", "zlBakInfo", "不存在", "严重：历史数据空间将不能正常运行")
        blnBakInfor = False
    End If
    If (bytCheckSys = 0 Or bytCheckSys > 1) And blnBakInfor Then
        If strDbLink <> "" Then
            gstrSQL = "Select 1 From " & strBakOwner & ".zlBakInfo where 系统=" & lngSys
        Else
            gstrSQL = "Select 1 From zlBakInfo@" & strDbLink & " where 系统=" & lngSys
        End If
        OpenRecordset rsTemp, gstrSQL, "获取系统", , , cnOracle
        If rsTemp.EOF Then
            cllErr.Add Array("系统数据", lngSys, "系统编号为:" & lngSys & "不存在", "严重：影响历史数据空间的正常运行")
        End If
        rsTemp.Close
    End If
    
    Dim lngCount As Long
    lngRow = 0
    If bytCheckSys >= 1 Then
        With rsObject
            If .RecordCount <> 0 Then .MoveFirst
            Do While Not .EOF
                Call CheckHistoryTable(cnOracle, Nvl(!表名), strSysOwner, strBakOwner, Nvl(!表名), strDbLink, lngSys, cllErr)
                lngRow = lngRow + 1
                pgbProcess.value = lngRow \ .RecordCount * 100
                .MoveNext
            Loop
        End With
    End If
    rsBakObject.Close
    Set rsBakObject = Nothing
    rsObject.Close
    Set rsObject = Nothing
    If cllErr Is Nothing Then
    Else
    If cllErr.Count <> 0 Then Set cllErrMsg = cllErr: Exit Function
    End If
    CheckHistorySpaces = True
    Exit Function
ErrHand:
   ' Resume
    If cllErr.Count <> 0 Then cllErrMsg = cllErr
End Function


Public Sub OpenRecordset(rsTemp As ADODB.Recordset, strSql As String, ByVal strFormCaption As String, _
        Optional CursorType As CursorTypeEnum = adOpenStatic, Optional LockType As LockTypeEnum = adLockReadOnly, _
        Optional cnOracle As ADODB.Connection = Nothing)
        '功能：打开记录。同时保存SQL语句
    On Error GoTo errHandle
    If rsTemp.State = adStateOpen Then rsTemp.Close
    rsTemp.CursorLocation = adUseClient
    If cnOracle Is Nothing Then
        rsTemp.Open strSql, gcnOracle, CursorType, LockType
    ElseIf cnOracle.State = 1 Then
        rsTemp.Open strSql, cnOracle, CursorType, LockType
    Else
        rsTemp.Open strSql, gcnOracle, CursorType, LockType
    End If
    Exit Sub
errHandle:
    MsgBox err.Description, vbExclamation, gstrSysName
End Sub

Public Function Lpad(ByVal strCode As String, lngLen As Long, Optional strChar As String = " ") As String
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:按指定长度填制空格
    '--入参数:
    '--出参数:
    '--返  回:返回字串
    '-----------------------------------------------------------------------------------------------------------
    Dim lngTmp As Long
    Dim strTmp As String
    strTmp = strCode
    lngTmp = LenB(StrConv(strCode, vbFromUnicode))
    If lngTmp < lngLen Then
        strTmp = String(lngLen - lngTmp, strChar) & strTmp
    ElseIf lngTmp > lngLen Then  '大于长度时,自动载断
        strTmp = strCode
    End If
    Lpad = Replace(strTmp, Chr(0), strChar)
End Function
Public Function RPAD(ByVal strCode As String, lngLen As Long, Optional strChar As String = " ") As String
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:按指定长度填制空格
    '--入参数:
    '--出参数:
    '--返  回:返回字串
    '-----------------------------------------------------------------------------------------------------------
    Dim lngTmp As Long
    Dim strTmp As String
    strTmp = strCode
    lngTmp = LenB(StrConv(strCode, vbFromUnicode))
    If lngTmp < lngLen Then
        strTmp = strTmp & String(lngLen - lngTmp, strChar)
    Else
        '主要有空格引起的
        strTmp = strCode
    End If
    '取掉最后半个字符
    RPAD = Replace(strTmp, Chr(0), strChar)
End Function


Public Function CheckHistoryTable(ByVal cnOracle As ADODB.Connection, _
    ByVal strTableName As String, ByVal strSysOwner As String, ByVal strBakOwnner As String, ByVal strHistoryTableName As String, strDbLinkName As String, _
    ByVal lngSys As Long, ByRef cllErr As Variant) As Boolean
    '功能:检查指定表的列与历史数据表空间的表名是否一致
    '参数:cnOracle-在线库连接
    '     strTable_name-在线表名
    '     strSysOwner-在线数据库的所有者
    '     strBakOwnner-备份空间的所有者
    '     strHistoryTableName-历史表名
    '     strDbLinkNameName-远程连接名
    '出参:cllErr:错误信息集(0-对象类型,1-对象名称,2-错语信息,2-问题严重级别说明)
    '返回:检查成功,返回ture,否则返回False
    '-----------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim rsColumn As New ADODB.Recordset
    Dim rsBakColumn As New ADODB.Recordset
    Dim strTemp As String
    '看该表是否存在zlbakSpaces中
    gstrSQL = "Select 1 from zltools.zlbaktables where 系统=" & lngSys & " and 表名='" & strHistoryTableName & "'"
    OpenRecordset rsTemp, gstrSQL, "检查历史表数据", , , cnOracle
    If rsTemp.RecordCount = 0 Then
        rsTemp.Close
        CheckHistoryTable = True
        Exit Function
    End If
    err = 0: On Error Resume Next
    If strDbLinkName = "" Then
        gstrSQL = "select table_name as 表名  from  all_tables where owner='" & UCase(strBakOwnner) & "' and table_name='" & strHistoryTableName & "'"
    Else
        gstrSQL = "select table_name as 表名  from user_tables@" & strDbLinkName & " where  table_name='" & strHistoryTableName & "'"
    End If
 
    OpenRecordset rsTemp, gstrSQL, "检查历史表是否存在", , , cnOracle
    
    If err = 0 Then
        If rsTemp.EOF Then
            cllErr.Add Array("表", strHistoryTableName, "不存在", "严重：影响历史数据空间的正常运行")
            GoTo LoopH:
        End If
    Else
            cllErr.Add Array("表", strHistoryTableName, "不能正常访问连接", "严重：影响历史数据空间的正常运行")
            GoTo LoopH:
    End If
    '设图的有效性检查
    err = 0: On Error Resume Next
    gstrSQL = "Select 1 from H" & strHistoryTableName & " where 1=2 "
    If err <> 0 Then
        '-------------------------------------
        '设图无效
        cllErr.Add Array("视图", "H" & strHistoryTableName, "不存在", "严重：影响历史数据空间的正常运行")
    End If
    
    '检查相关的列是否正确
    If rsColumn.State = 1 Then rsColumn.Close
    If rsBakColumn.State = 1 Then rsBakColumn.Close
    
    If strDbLinkName = "" Then
        gstrSQL = "SELECT COLUMN_NAME,DATA_TYPE,DATA_LENGTH,DATA_PRECISION,DATA_SCALE,DATA_DEFAULT" & _
            "       From  ALL_TAB_COLUMNS" & _
            "       WHERE owner='" & strBakOwnner & "' and TABLE_NAME='" & strHistoryTableName & "'"
    Else
        gstrSQL = "SELECT COLUMN_NAME,DATA_TYPE,DATA_LENGTH,DATA_PRECISION,DATA_SCALE,DATA_DEFAULT" & _
            "       From USER_TAB_COLUMNS@" & strDbLinkName & _
            "       WHERE TABLE_NAME='" & strHistoryTableName & "'"
    End If
    
    rsBakColumn.Open gstrSQL, cnOracle
    
    gstrSQL = "SELECT COLUMN_NAME,DATA_TYPE,DATA_LENGTH,DATA_PRECISION,DATA_SCALE,DATA_DEFAULT" & _
        "       From ALL_TAB_COLUMNS" & _
        "       WHERE TABLE_NAME='" & strTableName & "' and OWNER='" & strSysOwner & "'"
    
    rsColumn.Open gstrSQL, cnOracle
                
    With rsColumn
        Do While Not .EOF
            rsBakColumn.Filter = "COLUMN_NAME='" & Nvl(!Column_Name) & "'"
            If rsBakColumn.EOF Then
                '不存列
                Select Case Nvl(!DATA_TYPE)
                Case "NUMBER"
                    strTemp = Nvl(!Column_Name) & " NUMBER(" & Nvl(!Data_Precision) & "," & Nvl(!Data_Scale) & ")"
                    If Not IsNull(!DATA_DEFAULT) Then strTemp = strTemp & " DEFAULT " & !DATA_DEFAULT
                Case "VARCHAR2"
                    strTemp = Nvl(!Column_Name) & " VARCHAR2(" & Nvl(!Data_Length) & ")"
                    If Not IsNull(!DATA_DEFAULT) Then strTemp = strTemp & " DEFAULT " & !DATA_DEFAULT
                Case Else
                    strTemp = Nvl(!Column_Name) & Space(2) & Nvl(!DATA_TYPE)
                End Select
                cllErr.Add Array("数据表", strHistoryTableName, "缺少列 " & strTemp, "严重：影响历史数据空间的正常运行")
            Else
                '检查长度
                    Select Case !DATA_TYPE
                    Case "NUMBER"
                        If Val(Nvl(!Data_Precision)) <> Val(Nvl(rsBakColumn!Data_Precision)) Or Val(Nvl(!Data_Scale)) <> Val(Nvl(rsBakColumn!Data_Scale)) Then
                            strTemp = Nvl(!Column_Name) & "列长度小于规定值：应为“" & "NUMBER(" & Nvl(!Data_Precision) & "," & Val(Nvl(!Data_Scale)) & ")”" & _
                                     " 现为“" & "NUMBER(" & rsBakColumn!Data_Precision & "," & Val(Nvl(rsBakColumn!Data_Scale)) & ")”"
                            If Val(Nvl(!Data_Precision)) > Val(Nvl(rsBakColumn!Data_Precision)) Then
                                cllErr.Add Array("数据表", strHistoryTableName, strTemp, "严重：影响历史数据空间的正常运行")
                            ElseIf Val(Nvl(!Data_Scale)) > Val(Nvl(rsBakColumn!Data_Scale)) Then
                                cllErr.Add Array("数据表", strHistoryTableName, strTemp, "严重：可能导致历史数据空间数据精度不足")
                            Else
                                cllErr.Add Array("数据表", strHistoryTableName, strTemp, "较轻：基本不影响历史数据空间的运行")
                            End If
                        End If
                    Case "VARCHAR2"
                        If Val(Nvl(!Data_Length)) <> Val(Nvl(rsBakColumn!Data_Length)) Then
                            strTemp = Nvl(!Column_Name) & "列长度小于规定值：应为“" & "VARCHAR2(" & Val(Nvl(!Data_Length)) & ")”" & _
                                     " 现为“" & "VARCHAR2(" & Val(Nvl(rsBakColumn!Data_Length)) & ")”"
                            If Val(Nvl(!Data_Length)) > Val(Nvl(rsBakColumn!Data_Length)) Then
                                cllErr.Add Array("数据表", strHistoryTableName, strTemp, "较重：可能导致历史数据空间数据的较长文本无法存储")
                            Else
                                cllErr.Add Array("数据表", strHistoryTableName, strTemp, "较轻：基本不影响历史数据空间的运行")
                            End If
                        End If
                    Case Else
                    End Select
                    If Nvl(!DATA_TYPE) <> Nvl(rsBakColumn!DATA_TYPE) Then
                        strTemp = Nvl(!Column_Name) & "列的类型不对,应为“" & Nvl(!DATA_TYPE) & "”" & _
                                 " 现为“" & Nvl(rsBakColumn!DATA_TYPE) & "”"
                        cllErr.Add Array("数据表", strHistoryTableName, strTemp, "较重：可能导致历史数据空间的数据存储问题")
                        
                    End If
            End If
                 
            .MoveNext
        Loop
    End With
LoopH:
CheckHistoryTable = True
 
End Function

Public Sub CheckBakConstraint(ByVal cnOracle As ADODB.Connection, ByVal strBakOwner As String, ByVal strDbLinkName As String, ByVal strTableName As String, _
        ByVal strConstraintName As String, ByVal strSql As String, ByVal lngSys As Long, ByRef cllErr As Variant)
    '---------------------------------------------------------------------------------
    '功能:检查备份数据库的约束
    '参数:cnOracle-在线数据库连接
    '     strDbLinkName-远程连接
    '     strTableName-表名
    '     strBakOwner-历名数据空间所有者
    '     strConstraintName-约束名
    '     strSQL-约束的SQL语句
    '出参:cllErr-返回错误信息
    '---------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim rsConstraint As New ADODB.Recordset
    Dim rsColumns As New ADODB.Recordset
    Dim strColumns As String
    Dim arySql As Variant
    Dim strTemp As String
    '看该表是否存在zlbakSpaces中
 
    
    gstrSQL = "Select 1 from zltools.zlbaktables where 系统=" & lngSys & " and 表名='" & strTableName & "'"
    OpenRecordset rsTemp, gstrSQL, "检查历史表数据", , , cnOracle
    If rsTemp.RecordCount = 0 Then
        rsTemp.Close
        Exit Sub
    End If
    
    If strDbLinkName <> "" Then
        gstrSQL = "select CONSTRAINT_TYPE,CONSTRAINT_NAME,STATUS,VALIDATED,BAD from USER_CONSTRAINTS@" & strDbLinkName & " where CONSTRAINT_NAME='" & strConstraintName & "'"
    Else
        gstrSQL = "select CONSTRAINT_TYPE,CONSTRAINT_NAME,STATUS,VALIDATED,BAD from all_CONSTRAINTS where  OWNER='" & strBakOwner & "' and CONSTRAINT_NAME='" & strConstraintName & "'"
    End If
    OpenRecordset rsConstraint, gstrSQL, "读取约束", , , cnOracle
    
    With rsConstraint
            If rsConstraint.EOF Then
                '约束不存在
                If InStr(1, strSql, " CHECK") > 0 Then
                    cllErr.Add Array("约束", strConstraintName, "不存在", "较轻：基本不影响历史数据空间的运行")
                ElseIf InStr(1, strSql, " FOREIGN ") = 0 Then
                    cllErr.Add Array("约束", strConstraintName, "不存在", "较重：可能导致历史数据空间的数据不一致，影响运行速度")
                End If
                Exit Sub
            End If
            If .Fields("STATUS").value <> "ENABLED" Then
                cllErr.Add Array("约束", strConstraintName, "当前处于禁止状态", "较重：可能历史数据空间已经存在问题")
                Exit Sub
            End If
            If !VALIDATED <> "VALIDATED" Then
                cllErr.Add Array("约束", strConstraintName, "当前处于无效状态", "较重：可能历史数据空间的数据一致性已被破坏")
                Exit Sub
            End If
            If Not IsNull(!BAD) Then
                cllErr.Add Array("约束", strConstraintName, "约束被意外损坏", "严重：可能历史数据空间存在硬件错误")
                Exit Sub
            End If
            strColumns = ""
            
            If strDbLinkName = "" Then
                gstrSQL = "" & _
                    "   Select COLUMN_NAME" & _
                    "   From all_CONS_COLUMNS" & _
                    "   where owner='" & strBakOwner & "' and CONSTRAINT_NAME='" & strConstraintName & "'" & _
                    "   order by POSITION"
            Else
                gstrSQL = "" & _
                    "   Select COLUMN_NAME" & _
                    "   From USER_CONS_COLUMNS@" & strDbLinkName & _
                    "   where CONSTRAINT_NAME='" & strConstraintName & "'" & _
                    "   order by POSITION"
            End If
            OpenRecordset rsColumns, gstrSQL, "读取的相关列", , , cnOracle
                 
            With rsColumns
                Do While Not .EOF
                    strColumns = strColumns & "," & !Column_Name
                    .MoveNext
                Loop
            End With
            If InStr(1, strSql, " PRIMARY ") > 0 Then
                If !constraint_type <> "P" Then
                    cllErr.Add Array("约束", strConstraintName, "约束类型错误，应为主键约束", "严重：可能影响历史数据空间的运行")
                Else
                    arySql = Split(strSql, " PRIMARY ")
                    strTemp = Replace(Replace(Replace(Left(arySql(1), InStr(1, arySql(1), ")") - 1), "KEY", ""), "(", ""), " ", "")
                    If strColumns <> "," & strTemp Then
                        cllErr.Add Array("约束", strConstraintName, "约束列错误，应为(" & strTemp & ")，现为(" & Mid(strColumns, 2) & ")", "严重：可能影响历史数据空间的运行")
                    End If
                End If
                Exit Sub
            End If
            If InStr(1, strSql, " UNIQUE") > 0 Then
                If !constraint_type <> "U" Then
                    cllErr.Add Array("约束", strConstraintName, "约束类型错误，应为唯一约束", "较重：可能影响历史数据空间的运行")
                Else
                    arySql = Split(strSql, " UNIQUE ")
                    If UBound(arySql) = 0 Then arySql = Split(strSql, " UNIQUE(")
                    strTemp = Replace(Replace(Left(arySql(1), InStr(1, arySql(1), ")") - 1), "(", ""), " ", "")
                    If strColumns <> "," & strTemp Then
                        cllErr.Add Array("约束", strConstraintName, "约束列错误，应为(" & strTemp & ")，现为(" & Mid(strColumns, 2) & ")", "较重：可能影响历史数据空间的运行")
                    End If
                End If
                Exit Sub
            End If
            If InStr(1, strSql, " CHECK") > 0 Then
                If !constraint_type <> "C" Then
                    cllErr.Add Array("约束", strConstraintName, "约束类型错误，应为检查约束", "较重：可能影响历史数据空间的运行")
                End If
            End If
    End With
End Sub
Public Sub CheckBakIndex(ByVal cnOracle As ADODB.Connection, ByVal strBakOwner As String, ByVal strDbLinkName As String, ByVal strTableName As String, _
        ByVal strIndexName As String, ByVal strSql As String, ByVal lngSys As Long, ByRef cllErr As Variant)
    '---------------------------------------------------------------------------------
    '功能:检查备份数据库的约束
    '参数:cnOracle-在线数据库连接
    '     strBakOwner-备份空间所有者
    '     strDbLinkName-远程连接
    '     strTableName-表名
    '     strIndexName-索引名
    '     strSQL-约束的SQL语句
    '出参:cllErr-返回错误信息
    '---------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim rsIndex As New ADODB.Recordset
    Dim rsColumns As New ADODB.Recordset
    Dim strColumns As String
    Dim arySql As Variant
    Dim strTemp As String
    '看该表是否存在zlbakSpaces中
 
    On Error GoTo errHandle
    gstrSQL = "Select 1 from zltools.zlbaktables where 系统=" & lngSys & " and 表名='" & strTableName & "'"
    OpenRecordset rsTemp, gstrSQL, "检查历史表数据", , , cnOracle
    If rsTemp.RecordCount = 0 Then
        rsTemp.Close
        Exit Sub
    End If
    If strDbLinkName = "" Then
        gstrSQL = "select INDEX_NAME,STATUS from all_INDEXES where owner='" & strBakOwner & "' and   INDEX_NAME='" & strIndexName & "'"
    Else
        gstrSQL = "select INDEX_NAME,STATUS from USER_INDEXES@" & strDbLinkName & " where  INDEX_NAME='" & strIndexName & "'"
    End If
    OpenRecordset rsIndex, gstrSQL, "读取索引", , , cnOracle
    
    With rsIndex
            If rsIndex.EOF Then
                '约束不存在
                cllErr.Add Array("索引", strIndexName, "不存在", "较重：可能影响历史数据空间的运行速度")
                Exit Sub
            End If
            If .Fields("STATUS").value <> "VALID" Then
                cllErr.Add Array("索引", strIndexName, "当前处于无效状态", "较重:可能影响历史数据空间的运行速度")
                Exit Sub
            End If
            
            If strDbLinkName = "" Then
                strTemp = "select TABLE_NAME,COLUMN_NAME" & _
                        " from all_IND_COLUMNS" & _
                        " where INDEX_OWNER ='" & strBakOwner & "' and INDEX_NAME='" & strIndexName & "'" & _
                        " order by COLUMN_POSITION"
            Else
                strTemp = "select TABLE_NAME,COLUMN_NAME" & _
                        " from USER_IND_COLUMNS@" & strDbLinkName & _
                        " where INDEX_NAME='" & strIndexName & "'" & _
                        " order by COLUMN_POSITION"
            End If
            
            OpenRecordset rsColumns, strTemp, "读取的相关索引列", , , cnOracle
           
            With rsColumns
                Do While Not .EOF
                    If .AbsolutePosition = 1 Then
                        strColumns = !Table_Name & "(" & !Column_Name
                    Else
                        strColumns = strColumns & "," & !Column_Name
                    End If
                    .MoveNext
                Loop
                strColumns = strColumns & ")"
            End With
            arySql = Split(strSql, " ON ")
            strTemp = Replace(Left(arySql(1), InStr(1, arySql(1), ")")), " ", "")
            If strColumns <> strTemp Then
               cllErr.Add Array("索引", strIndexName, "索引列错误，应为“" & strTemp & "”，现为“" & strColumns & "”", "较重：可能影响系统运行速度")
            End If
    End With
    Exit Sub
errHandle:
    MsgBox err.Description, vbExclamation, gstrSysName
End Sub

Public Sub CheckBakView(ByVal cnOracle As ADODB.Connection, ByVal strOwner As String, ByVal lngSys As Long, ByRef cllErr As Variant)
    '---------------------------------------------------------------------------------
    '功能:检查在线的历史数据空间视图是否存在
    '参数:cnOracle-在线数据库连接
    '     strOwner-所有者
    '     lngSys-系统号
    '出参:cllErr-返回错误信息
    '---------------------------------------------------------------------------------
    Dim rsBakTable As New ADODB.Recordset
    Dim rsView As New ADODB.Recordset
    Dim rsObject As New ADODB.Recordset
    Dim strSql As String
    '看该表是否存在zlbakSpaces中
    On Error GoTo errHandle
    strSql = "Select 表名 from zltools.zlbaktables where 系统=" & lngSys
    OpenRecordset rsBakTable, strSql, "检查历史表数据", , , cnOracle
    If rsBakTable.RecordCount = 0 Then
        rsBakTable.Close
        Exit Sub
    End If
    
    If gblnDBA Then
        strSql = "select VIEW_NAME from DBA_VIEWS where OWNER='" & strOwner & "'"
    Else
        strSql = "select VIEW_NAME from USER_VIEWS"
    End If
    OpenRecordset rsView, strSql, "检查历史表数据", , , cnOracle
    
    With rsBakTable
        Do While Not .EOF
            rsView.Filter = "VIEW_NAME='" & "H" & UCase(!表名) & "'"
            If rsView.EOF Then
                '视图不存在
                cllErr.Add Array("视图", "H" & UCase(!表名), "不存在", "较重：可能影响历史数据空间的数据转储")
            End If
            .MoveNext
        Loop
    End With
    Exit Sub
errHandle:
    MsgBox err.Description, vbExclamation, gstrSysName
End Sub


Public Sub ExecuteProcedure(strSql As String, ByVal strFormCaption As String, Optional cnOracle As ADODB.Connection)
'功能：执行过程语句,并自动对过程参数进行绑定变量处理
'参数：strSQL=过程语句,可能带参数,形如"过程名(参数1,参数2,...)"。
'说明：以下几种情况过程参数不使用绑定变量,仍用老的调用方法：
'  1.参数部份是表达式,这时程序无法处理绑定变量类型和值,如"过程名(参数1,100.12*0.15,...)"
'  2.中间没有传入明确的可选参数,这时程序无法处理绑定变量类型和值,如"过程名(参数1, , ,参数3,...)"
'  3.因为该过程是自动处理,不是一定使用绑定变量,对带"'"的字符参数,仍要使用"''"形式。
    Dim cmdData As New ADODB.Command
    Dim strProc As String, strPar As String
    Dim blnStr As Boolean, intBra As Integer
    Dim strTemp As String, i As Long
    Dim intMax As Integer, datCur As Date
    
    If Right(Trim(strSql), 1) = ")" Then
        '清除原有参数:不然不能重复执行
'        cmdData.CommandText = "" '不为空有时清除参数出错
'        Do While cmdData.Parameters.Count > 0
'            cmdData.Parameters.Delete 0
'        Loop
        
        '执行的过程名
        strTemp = Trim(strSql)
        strProc = Trim(Left(strTemp, InStr(strTemp, "(") - 1))
        
        '执行过程参数
        datCur = CDate(0)
        strTemp = Mid(strTemp, InStr(strTemp, "(") + 1)
        strTemp = Trim(Left(strTemp, Len(strTemp) - 1)) & ","
        For i = 1 To Len(strTemp)
            '是否在字符串内，以及表达式的括号内
            If Mid(strTemp, i, 1) = "'" Then blnStr = Not blnStr
            If Not blnStr And Mid(strTemp, i, 1) = "(" Then intBra = intBra + 1
            If Not blnStr And Mid(strTemp, i, 1) = ")" Then intBra = intBra - 1
            
            If Mid(strTemp, i, 1) = "," And Not blnStr And intBra = 0 Then
                strPar = Trim(strPar)
                With cmdData
                    If IsNumeric(strPar) Then '数字
                        .Parameters.Append .CreateParameter("PAR" & .Parameters.Count, adVarNumeric, adParamInput, 30, Val(strPar))
                    ElseIf Left(strPar, 1) = "'" And Right(strPar, 1) = "'" Then '字符串
                        strPar = Mid(strPar, 2, Len(strPar) - 2)
                        
                        'Oracle连接符运算:'ABCD'||CHR(13)||'XXXX'||CHR(39)||'1234'
                        If InStr(Replace(strPar, " ", ""), "'||") > 0 Then GoTo NoneVarLine
                        
                        '双"''"的绑定变量处理
                        If InStr(strPar, "''") > 0 Then strPar = Replace(strPar, "''", "'")
                        
                        '电子病历处理LOB时，如果用绑定变量转换为RAW时第2000个字符不正确
                        intMax = LenB(StrConv(strPar, vbFromUnicode))
                        If intMax = 0 Or intMax < 200 Then intMax = 200
                        If intMax > 1999 Then GoTo NoneVarLine
                        
                        .Parameters.Append .CreateParameter("PAR" & .Parameters.Count, adVarChar, adParamInput, intMax, strPar)
                    ElseIf UCase(strPar) Like "TO_DATE('*','*')" Then '日期
                        strPar = Split(strPar, "(")(1)
                        strPar = Trim(Split(strPar, ",")(0))
                        strPar = Mid(strPar, 2, Len(strPar) - 2)
                        If strPar = "" Then
                            'NULL值当成数字处理可兼容其他类型
                            .Parameters.Append .CreateParameter("PAR" & .Parameters.Count, adVarNumeric, adParamInput, , Null)
                        Else
                            If Not IsDate(strPar) Then GoTo NoneVarLine
                            .Parameters.Append .CreateParameter("PAR" & .Parameters.Count, adDBTimeStamp, adParamInput, , CDate(strPar))
                        End If
                    ElseIf UCase(strPar) = "SYSDATE" Then '日期
                        If datCur = CDate(0) Then datCur = CurrentDate
                        .Parameters.Append .CreateParameter("PAR" & .Parameters.Count, adDBTimeStamp, adParamInput, , datCur)
                    ElseIf UCase(strPar) = "NULL" Then 'NULL值当成字符处理可兼容其他类型
                        .Parameters.Append .CreateParameter("PAR" & .Parameters.Count, adVarChar, adParamInput, 200, Null)
                    ElseIf strPar = "" Then '可选参数当成NULL处理可能改变了缺省值:因此可选参数不能写在中间
                        GoTo NoneVarLine
                    Else '可能是其他复杂的表达式，无法处理
                        GoTo NoneVarLine
                    End If
                End With
                
                strPar = ""
            Else
                strPar = strPar & Mid(strTemp, i, 1)
            End If
        Next
        
        '补充?号
        strTemp = ""
        For i = 1 To cmdData.Parameters.Count
            strTemp = strTemp & ",?"
        Next
        strProc = "Call " & strProc & "(" & Mid(strTemp, 2) & ")"
        
        '执行过程
        'If cmdData.ActiveConnection Is Nothing Then
        If cnOracle Is Nothing Then
            Set cmdData.ActiveConnection = gcnOracle '这句比较慢
        Else
            Set cmdData.ActiveConnection = cnOracle '这句比较慢
        End If
            cmdData.CommandType = adCmdText
        'End If
        cmdData.CommandText = strProc
        
        Call cmdData.Execute

    Else
        GoTo NoneVarLine
    End If
    Exit Sub
NoneVarLine:
    
    '说明：为了兼容新连接方式
    '1.新连接用adCmdStoredProc方式在8i下面有问题
    '2.新连接如果不使用{},则即使过程没有参数也要加()
    strSql = "Call " & strSql
    If InStr(strSql, "(") = 0 Then strSql = strSql & "()"
    gcnOracle.Execute strSql, , adCmdText

End Sub

Public Sub AlterUserTableSpaces(ByVal cnOracle As ADODB.Connection, ByVal strUserName As String)
    '----------------------------------------------------------------------------------------------------------------------------
    '功能:修改指定用户的默认表空间(问题:10793)
    '参数:cnOracle-指定的Oracle连接
    '     strUserName-指定用户名
    '编制:刘兴宏
    '日期:2007/06/01
    '说明:
    '   经与周韬和张永康落实，一般情况下,将用户的表空间放在USERS和Temp下,如果用户更改掉了Users表空间和TMP表空间,
    '   则放在ZLTOOLSTBS和ZLTOOLSTMP表空间中
    '----------------------------------------------------------------------------------------------------------------------------
    '刘兴宏:20070531:用户的表空间不能用system,改为USERS的情况
    '因为Create User时授予了用户Create Table权限,从安全考虑,缺省表空间最好为USERS
    
    err = 0: On Error Resume Next
    '可能没有相应的表空间,出错继续
    '9i引入了缺省全局临时表空间,10G引入了临时表空间组,暂不做特殊考虑
    gstrSQL = "Alter User " & strUserName & " Default Tablespace USERS"
    cnOracle.Execute gstrSQL
    gstrSQL = "Alter User " & strUserName & " Temporary Tablespace TEMP"
    cnOracle.Execute gstrSQL
    If err <> 0 Then
        '更改成ZLToolsTBS表空间和ZLTOOLSTMP表空间
        gstrSQL = "Alter User " & strUserName & " Default Tablespace ZLTOOLSTBS"
        cnOracle.Execute gstrSQL
        gstrSQL = "Alter User " & strUserName & " Temporary Tablespace ZLTOOLSTMP"
        cnOracle.Execute gstrSQL
    End If
    err.Clear
End Sub


Private Function LogTime() As String
    LogTime = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] "
End Function

Public Sub zlInitRec(ByRef rsData As ADODB.Recordset)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查构建数据集
    '入参:rsData-数据集
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2009-08-19 14:00:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set rsData = New ADODB.Recordset
    With rsData
        .Fields.Append "序号", adDouble, 18, adFldIsNullable
        .Fields.Append "表名称", adLongVarChar, 80, adFldIsNullable
        .Fields.Append "对象名称", adLongVarChar, 80, adFldIsNullable
        .Fields.Append "类型", adLongVarChar, 10, adFldIsNullable  '类型:主键,唯一,外键,约束,索引,过程/函数,视图,...
        .Fields.Append "标志", adLongVarChar, 2, adFldIsNullable '1-存在对象,2-不存在对象,3-失效,4-缺少列,5-精度,6-处于禁止状态,7-约束不一致,8-处理主键时，需要先处理外键
        .Fields.Append "历史空间", adLongVarChar, 2, adFldIsNullable '1-历史数据库间错误,0-在线数据库对象错误不存在
        .Fields.Append "错误类型", adLongVarChar, 30, adFldIsNullable '不存在,缺少列...
        .Fields.Append "错误信息", adLongVarChar, 500, adFldIsNullable
        .Fields.Append "修正语句", adLongVarChar, 4000, adFldIsNullable '存储过程，由于数据太大，因此填为NULL
        .Fields.Append "字段名", adLongVarChar, 50, adFldIsNullable  '表时,需要更改的字段名
        .Fields.Append "原字段类型", adLongVarChar, 50, adFldIsNullable  '表时,需要更改的原字段类型
        .Fields.Append "现字段类型", adLongVarChar, 50, adFldIsNullable  '表时,需要更改的字段名的类型
        .Fields.Append "原字段长度", adLongVarChar, 20, adFldIsNullable  '表时,需要更改的字段名的长度,如果存在小数,则以逗号分离,如:16,5
        .Fields.Append "现字段长度", adLongVarChar, 20, adFldIsNullable  '表时,需要更改的字段名的长度,如果存在小数,则以逗号分离,如:16,5
        .Fields.Append "修正标志", adLongVarChar, 2, adFldIsNullable  '0-未修正,1-已经修正,2-修正错败,4-不能执行修正，需要手工调整
        .Fields.Append "修正说明", adLongVarChar, 500, adFldIsNullable '
        .CursorLocation = adUseClient: .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub

Public Function zlInsertRecData(ByVal rsData As ADODB.Recordset, ByVal str表名称 As String, ByVal str对象名称 As String, ByVal str类型 As String, ByVal int标志 As Integer, _
       ByVal bln历史空间 As Boolean, ByVal str修正语句 As String, ByVal str错误类型 As String, ByVal str错误信息 As String, Optional str字段名 As String, Optional str原字段类型 As String, Optional str现字段类型 As String, _
       Optional str原字段长度 As String, Optional str现字段长度 As String, Optional cllProcedureExecSQLs As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:向本地记录集中插入数据
    '入参:
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2009-08-19 14:25:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng序号 As Long, str修正说明 As String, byt修正标志 As Byte, varData As Variant
    Dim lng长度 As Long, lng小数 As Long
    
    str对象名称 = UCase(str对象名称)
    byt修正标志 = 0: str修正说明 = ""
    If Trim(str字段名) <> "" Then
        '需要检查，字段修正情况:
        ' 1.类型不对,不能执行修正，需要手工调整
        ' 2.字段长度小于了原字段长度,不能修修正(包含小数)，需要手工调整
        
        If str原字段类型 <> str现字段类型 And (str原字段类型 <> "") Then
            '0-未修正,1-已经修正,2-修正错败,4-不能执行修正，需要手工调整
            byt修正标志 = 4
            str修正说明 = "字段类型不对，不能修正, 变动情况:" & str原字段类型 & "--->" & str现字段类型
        Else
            varData = Split(str原字段长度 & ",", ",")
            lng长度 = Val(varData(0)): lng小数 = Val(varData(1))
            varData = Split(str现字段长度 & ",", ",")
            If UCase(str现字段类型) = "NUMBER" Then
                '如果是NUMBER的话，需要检查其长度
                If lng长度 > Val(varData(0)) Or lng小数 > Val(varData(1)) Then
                    byt修正标志 = 4
                    str修正说明 = "原字段精度大于了现字段精度，不能修正, 变动情况:" & str字段名 & "  NUMBER(" & lng长度 & IIf(lng小数 = 0, "", "," & lng小数) & ")" & "--->" & str字段名 & "  NUMBER(" & Val(varData(0)) & IIf(Val(varData(1)) = 0, "", "," & Val(varData(1))) & ")"
                End If
            ElseIf Left(UCase(Trim(str现字段类型)), 7) = "VARCHAR" Then
                '字符型，需检查长度
                If lng长度 > Val(varData(0)) Then
                    byt修正标志 = 4
                    str修正说明 = "原字段精度大于了现字段精度，不能修正, 变动情况:" & str字段名 & "  NUMBER(" & lng长度 & ")" & "--->" & str字段名 & "  NUMBER(" & Val(varData(0)) & ")"
                End If
            End If
        End If
    End If
    With rsData
        lng序号 = .RecordCount + 1
        If str类型 = "外键" Then
            .Filter = "对象名称='" & str对象名称 & "'"
            If .RecordCount = 0 Then
                .Filter = 0
                '不存在,只能新增,存在就更改:原因是可能存在级联这种情况
                .AddNew
            End If
        Else
            .AddNew
        End If
        !序号 = lng序号
        !表名称 = str表名称
        !对象名称 = str对象名称
        !类型 = str类型
        !标志 = int标志
        !历史空间 = IIf(bln历史空间, 1, 0)
        !修正语句 = IIf(str类型 = "过程/函数", "", str修正语句)
        !错误类型 = str错误类型
        !错误信息 = str错误信息
        !字段名 = str字段名
        !原字段类型 = str原字段类型
        !现字段类型 = str现字段类型
        !原字段长度 = str原字段长度
        !现字段长度 = str现字段长度
        !修正标志 = byt修正标志
        !修正说明 = str修正说明
        .Update
        If str类型 = "过程/函数" Then
            cllProcedureExecSQLs.Add str修正语句, "K" & lng序号
        End If
        .Filter = 0
    End With
End Function

Public Function Substr(ByVal strInfor As String, ByVal lngStart As Long, ByVal lngLen As Long) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:读取指定字串的值,字串中可以包含汉字
    '入参:strInfor-原串
    '      lngStart-直始位置
    '      lngLen-长度
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2009-08-20 12:04:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTmp As String, i As Long
    err = 0
    On Error GoTo ErrHand:
    Substr = StrConv(MidB(StrConv(strInfor, vbFromUnicode), lngStart, lngLen), vbUnicode)
    Substr = Replace(Substr, Chr(0), " ")
    Exit Function
ErrHand:
    Substr = ""
End Function

'去掉TextBox的默认右键菜单
Public Function WndMessage(ByVal hwnd As OLE_HANDLE, ByVal MSG As OLE_HANDLE, ByVal wp As OLE_HANDLE, ByVal lp As Long) As Long
    ' 如果消息不是WM_CONTEXTMENU，就调用默认的窗口函数处理
    If MSG <> WM_CONTEXTMENU Then WndMessage = CallNewWindowProc(hwnd, MSG, wp, lp)
End Function

Public Function CallNewWindowProc(ByVal hwnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Call CallWindowProc(glngTXTProc, hwnd, MSG, wParam, lParam)
    
    CallNewWindowProc = True
End Function

Public Function IsCharAlpha(ByVal strAsk As String) As Boolean
    '-------------------------------------------------------------
    '功能：判断指定字符串是否全部由英文字母构成    '
    '参数：
    '       strAsk
    '返回：
    '-------------------------------------------------------------
    Dim i As Integer, j As Integer
    
    If Len(Trim(strAsk)) > 0 Then
        For i = 1 To Len(Trim(strAsk))
            j = Asc(Mid(Trim(strAsk), i, 1))
            If Not ((j > 64 And j < 91) Or (j > 96 And j < 123)) Then
                IsCharAlpha = False
                Exit Function
            End If
        Next
    End If
    IsCharAlpha = True
End Function

Public Function IsCharChinese(ByVal strAsk As String) As Boolean
    '-------------------------------------------------------------
    '功能：判断指定字符串是否含有汉字
    '参数：
    '       strAsk
    '返回：
    '-------------------------------------------------------------
    Dim i As Integer, j As Integer
    
    If Len(Trim(strAsk)) > 0 Then
        For i = 1 To Len(Trim(strAsk))
            j = Asc(Mid(Trim(strAsk), i, 1))
            If j < 0 Then
                IsCharChinese = True
                Exit Function
            End If
        Next
    End If
    IsCharChinese = False
End Function

Public Function SQLRecordExecute(ByVal rs As ADODB.Recordset, Optional ByVal strTitle As String, Optional ByVal blnHaveTrans As Boolean = True) As Boolean
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    Dim blnTran As Boolean
    Dim intLoop As Integer
    Dim strSql As String
    
    On Error GoTo ErrHand
    
    If rs.RecordCount > 0 Then
        If Len(strTitle) = 0 Then strTitle = gstrSysName
        blnTran = True
        
        If blnHaveTrans Then gcnOracle.BeginTrans
        
        rs.MoveFirst
    
        For intLoop = 1 To rs.RecordCount
        
            strSql = CStr(rs("SQL").value)
            
            Call ExecuteProcedure(strSql, strTitle)
            
            rs.MoveNext
        Next
    
        If blnHaveTrans Then gcnOracle.CommitTrans
        blnTran = False
    End If
    
    SQLRecordExecute = True
    
    Exit Function
ErrHand:
    
    If blnTran And blnHaveTrans Then gcnOracle.RollbackTrans
    
    MsgBox err.Description, vbCritical, gstrSysName
    
    
End Function


Public Function GetCommpentVersion(ByVal strFile As String) As String
    '-----------------------------------------------------------------------------------------------------------
    '功能:获取指定控件的版本号
    '入参:
    '出参:
    '返回:成功,返回版本号,否则返回空
    '编制:刘兴洪
    '日期:2009-01-16 16:59:34
    '-----------------------------------------------------------------------------------------------------------
    Dim objFile As New FileSystemObject
    Dim strVer As String, varVersion As Variant
    
    err = 0: On Error Resume Next
    '获取文件版本号
    strVer = objFile.GetFileVersion(strFile)
    If err <> 0 Then
        err.Clear: err = 0
        GetCommpentVersion = ""
        Exit Function
    End If
    If Trim(strVer) <> "" Then
        varVersion = Split(strVer, ".")
        If UBound(varVersion) > 2 Then
            strVer = varVersion(0) & "." & varVersion(1) & "." & varVersion(3)
        ElseIf UBound(varVersion) = 2 Then
            strVer = varVersion(0) & "." & varVersion(1) & "." & varVersion(2)
        End If
    End If
    GetCommpentVersion = strVer
End Function

Function GetFileName(ByVal strFileName As String, Optional Path As String, Optional WithExt As Boolean = False) As String
'获得文件名
'strFilename 文件绝对路径
'Path 返回位置
'WithExt 是否返回后缀名称 True:带后缀名称返回 false:不带后缀名称返回
    Dim C As String
    Dim tmpString As String
    Dim i As Integer
    Dim szlen As Integer
    Dim Cnt As Integer
    
    szlen = Len(strFileName)
    Cnt = 0
    If InStr(strFileName, "\") = 0 Then
      tmpString = strFileName
      Cnt = InStr(tmpString, ".")
      If Cnt > 0 And Not WithExt Then
          GetFileName = Left(tmpString, Cnt - 1)
      Else
          GetFileName = tmpString
      End If
    Else
      For i = szlen To 1 Step -1
        C = Mid(strFileName, i, 1)
        If C = "\" Then
          Path = Left(strFileName, szlen - Cnt)
          tmpString = Right(strFileName, Cnt)
          Cnt = InStr(tmpString, ".")
          If Cnt > 0 And Not WithExt Then
              GetFileName = Left(tmpString, Cnt - 1)
          Else
              GetFileName = tmpString
          End If
          Exit For
        Else
          Cnt = Cnt + 1
        End If
      Next i
    End If
End Function


Public Function GetWinPath() As String
    '--------------------------------------------------------------------------------------------------------------
    '--功能:获取系统目录
    '--------------------------------------------------------------------------------------------------------------
    Dim Buffer As String
    Dim gstrWinPath As String
    Dim rtn As Long
    
    Buffer = Space(MAX_PATH)
    rtn = GetWindowsDirectory(Buffer, Len(Buffer))
    gstrWinPath = Left(Buffer, rtn)
    GetWinPath = gstrWinPath
End Function

Public Function GetWinSystemPath() As String
    
    Dim Buffer As String
    Dim strSystem As String
    Dim rtn As Long
    
    Buffer = Space(MAX_PATH)
    rtn = GetSystemDirectory(Buffer, Len(Buffer))
    strSystem = Left(Buffer, rtn)
    
    GetWinSystemPath = strSystem
End Function

Public Sub LvwFlatColumnHeader(ByVal lvw As Object)
'功能：使用ListView的列标题成为平面
    Const strHeaderClass As String = "msvb_lib_header"
    Const HDS_BUTTONS   As Long = 2
    
    Dim lngChild As Long, lngLen As Long, LngStyle As Long
    Dim strName As String * 255

    
    lngChild = GetWindow(lvw.hwnd, GW_CHILD)
    Do While lngChild <> 0
        lngLen = GetClassName(lngChild, strName, 255)
        If lngLen > 0 Then
            If Mid(strName, 1, lngLen) = strHeaderClass Then
                LngStyle = GetWindowLong(lngChild, GWL_STYLE)
                LngStyle = LngStyle And (Not HDS_BUTTONS)
                SetWindowLong lngChild, GWL_STYLE, LngStyle
                Exit Sub
            End If
        End If
        lngChild = GetWindow(lngChild, GW_HWNDNEXT)
    Loop
End Sub

Public Function CheckRushHours(ByVal strModuleNo As String, ByVal strFuncName As String) As Boolean
'功能：检查当前时间是否是业务高峰期
'参数：
'      strModuleNo=模块号
'      strFuncName=功能名称
'返回：是否可以进行操作

    Dim strSql As String
    Dim strTime As String
    Dim rsTemp As ADODB.Recordset
    Dim blnLimit As Boolean
    Dim dateNow As Date
    Dim strNote As String
    
    On Error GoTo errH
    CheckRushHours = True
    dateNow = CDate(Format(CurrentDate(), "HH:MM:SS"))
    strSql = "Select a.操作选项, a.限时原因, To_Char(b.开始时间, 'HH24:MI:SS') 开始时间, To_Char(b.结束时间, 'HH24:MI:SS') 结束时间" & vbNewLine & _
            "From Zlrunlimitset A, Zlrunlimittime B" & vbNewLine & _
            "Where a.方案序号 = b.方案 And a.系统 Is Null And a.模块 = [1] And a.功能 = [2] And b.星期 = To_Char(Sysdate, 'd') - 1" & vbNewLine & _
            "Order By b.开始时间"
    Set rsTemp = gclsBase.OpenSQLRecord(gcnOracle, strSql, "获取指定功能当天的限时时间数据", strModuleNo, strFuncName)
    With rsTemp
        If .RecordCount = 0 Then Exit Function
        If Nvl(!限时原因) = "" Then
            strNote = "当前时段处于业务高峰期，使用此功能可能会对系统使用造成一定影响"
        Else
            strNote = !限时原因
        End If
        Do While Not .EOF
            strTime = strTime & "，" & !开始时间 & "-" & !结束时间
            If dateNow > !开始时间 And dateNow < !结束时间 Then
                '说明当前时间在限制时间的范围里
                blnLimit = True
            End If
            .MoveNext
        Loop
        If blnLimit = True Then
            .MoveFirst
            If !操作选项 = 0 Then  '弹出提示，并禁止用户进行后续操作
                MsgBox strNote & vbNewLine & "故在时间范围：" & vbNewLine & Mid(strTime, 2) & vbNewLine & "内禁止使用此功能！", vbInformation, gstrSysName
            Else   '弹出提示，但不禁止用户进行后续操作
                If MsgBox(strNote & vbNewLine & "确定要继续吗？", vbInformation + vbOKCancel, gstrSysName) = vbOK Then
                    blnLimit = False
                Else
                    blnLimit = True
                End If
            End If
        End If
    End With
    CheckRushHours = Not blnLimit
    Exit Function
errH:
    MsgBox err.Description, vbInformation, gstrSysName
End Function

Public Sub SaveAuditLog(ByVal lngType As Long, ByVal strFunction As String, ByVal strContent As String, Optional ByVal strDescription As String, Optional ByVal strModule As String)
'功能：插入重要工作日志
'参数：lngType = 操作类型，1-新增，2-修改，3-删除
'      strFunction = 功能名称
'      strContent = 操作内容
'      strDescription = 操作说明
'      strModule = 操作模块   ，一般来讲操作模块使用frmMDIMain.gstrLastModule即可，但有可能一个操作模块调用另一个操作模块的功能，此时则需要手动设定一个操作模块
    Dim strSql As String
    
    On Error GoTo errH:
    If LenB(StrConv(strContent, vbFromUnicode)) > 1024 Then
        strContent = Mid(strContent, 1, 500)
    End If
    If strModule = "" Then strModule = frmMDIMain.gstrLastModule
    strSql = "zltools.Zl_Zlauditlog_Insert('" & gstrLoginUserName & "','" & _
                                                gstrComputerName & "'," & _
                                                lngType & ",Null,'" & _
                                                strModule & "','" & _
                                                strFunction & "','" & _
                                                strContent & "','" & _
                                                strDescription & "')"
    Call ExecuteProcedure(strSql, "保存重要工作日志")
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
End Sub


Public Function InstrEx(ByVal strTxt As String, ByVal strCheck As String, Optional ByVal strDeli As String) As Boolean
    '功能:比较strCheck是否存在于strTxt中
    'strDeli-字符串之间的分隔符,默认为,
    
    If strDeli = "" Then strDeli = ","
    strTxt = strDeli & strTxt & strDeli
    strCheck = strDeli & strCheck & strDeli
    
    InstrEx = InStr(1, strTxt, strCheck) > 0
    
End Function

Public Function GetVersion() As String
'功能：获取数据库的大版本号
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    Dim arrTmp As Variant
    
    On Error GoTo errH
    'CORE    10.2.0.3.0  Production
    strSql = "Select Banner From V$version Where Banner Like  'CORE%'"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSql, App.Title)
    If rsTmp.RecordCount > 0 Then
        arrTmp = Split(TrimEx(rsTmp!Banner & ""), " ")
        If UBound(arrTmp) = 2 Then
            GetVersion = Mid(arrTmp(1), 1, InStr(1, arrTmp(1), ".") - 1)
        End If
    End If
    
    Exit Function
errH:
    MsgBox err.Description, vbExclamation, "错误"
End Function


Public Sub GetRowPos(objVsf As Object, strTxt As String, strCol As String)
'功能: 根据传入的字符串定位到表格
'参数:strTxt-需要匹配的字段 strCol; strCol 需要匹配的列,每个字段之间用逗号间隔 ;objFocus-搜索完成后获取焦点的对象
    Dim intRow As Integer, i As Integer, j As Integer
    Dim strFiels() As String, blnResult As Boolean
    
    strFiels = Split(strCol, ",")
    blnResult = False
    '输入数据就进行匹配
    With objVsf
        '第一次循环,从当前行进行匹配,匹配至最后一行
        intRow = 0
        For i = .Row + 1 To .Rows - .FixedRows
            For j = 0 To UBound(strFiels)   '循环每个列,有一个满足就记为当前行符合条件
                If (UCase(.TextMatrix(i, .ColIndex(strFiels(j)))) Like "*" & UCase(strTxt) & "*" Or UCase(.RowData(i)) = UCase(strTxt)) And .RowHidden(i) = False Then
                    blnResult = True
                    Exit For
                End If
            Next
            
            If blnResult Then '定位至当前行
                intRow = i
                .Select i, 1
                .TopRow = IIf(Val(i - 10) < 0, i, i - 10)   '如果行数过多,确保定位在表格中间.
                Exit Sub
            End If
        Next
        '第二次循环,从第一行匹配至当前行
        If .Row <> .FixedRows And intRow = 0 Then
            If MsgBox("未找到匹配信息,是否从头重新寻找?", vbYesNo + vbQuestion + vbDefaultButton1, "") = vbYes Then
                For i = .FixedRows To .Row - 1
                    For j = 0 To UBound(strFiels)   '循环每个列,有一个满足就记为当前行符合条件
                        If (UCase(.TextMatrix(i, .ColIndex(strFiels(j)))) Like "*" & UCase(strTxt) & "*" Or UCase(.RowData(i)) = UCase(strTxt)) And .RowHidden(i) = False Then
                            blnResult = True
                            Exit For
                        End If
                    Next
                    
                    If blnResult Then '定位至当前行
                        intRow = i
                        .Select i, 1
                        .TopRow = IIf(Val(i - 10) < 0, i, i - 10)   '如果行数过多,确保定位在表格中间.
                        Exit Sub
                    End If
                Next
            End If
        End If
        
        '两次都没有找到,给予提示
        If intRow = 0 Then
            For j = 0 To UBound(strFiels)   '检查当前行
                If (UCase(.TextMatrix(.Row, .ColIndex(strFiels(j)))) Like "*" & UCase(strTxt) & "*" Or UCase(.RowData(.Row)) = UCase(strTxt)) And .RowHidden(.Row) = False Then
                    blnResult = True
                    Exit For
                End If
            Next
            
            If Not blnResult Then
                MsgBox "未在表格中匹配到数据。", , "提示"
            End If
        End If
    End With
End Sub


Public Function TranStr2Var(ByVal strTxt As String, ByVal strDeli, ByVal intLength) As Variant
'功能: 将超过指定长度字符串,转换成数组
    Dim varTmp As Variant, strTmp As String
    varTmp = Array()
    
    ReDim varTmp(0): varTmp(0) = strTxt
    Do While Len(strTxt) > intLength
        '直接取指定长度前一个分隔符作为数组最后一个元素
        strTmp = Left(strTxt, intLength)
        strTmp = Left(strTmp, InStrRev(strTmp, strDeli) - 1)
        varTmp(UBound(varTmp)) = strTmp
        
        '原字符串去掉截取出的部分
        strTxt = Mid(strTxt, Len(varTmp(UBound(varTmp))) + 2)
        
        ReDim Preserve varTmp(UBound(varTmp) + 1)
    Loop
    
    If strTxt <> "" Then
        varTmp(UBound(varTmp)) = strTxt
    End If
    
    TranStr2Var = varTmp
End Function

Public Function LoadUsers(Optional blnIncludeDBA As Boolean) As ADODB.Recordset
'功能:获取用户名,返回数据集
    Dim strSql As String
    
    On Error GoTo errH
    strSql = "Select Distinct a.Username 用户名, b.姓名, b.名称 部门, b.简码" & vbNewLine & _
                    "From Dba_Users A," & vbNewLine & _
                    "     (Select b.用户名, c.名称, d.姓名, d.简码" & vbNewLine & _
                    "       From 上机人员表 B, 部门表 C, 人员表 D, 部门人员 E" & vbNewLine & _
                    "       Where b.人员id = d.Id And e.人员id = d.Id And e.部门id = c.Id And e.缺省 = 1) B" & vbNewLine & _
                    "Where a.Username = b.用户名(+) And a.Account_Status Not In ('LOCKED', 'EXPIRED & LOCKED') " & vbNewLine & _
                    IIf(blnIncludeDBA, "", " And Not Exists (Select 1 From Dba_Role_Privs Where Granted_Role = 'DBA' And Grantee = Username) ") & vbNewLine & _
                    IIf(blnIncludeDBA, "", " And Not Exists (Select 1 From Dba_Sys_Privs Where Grantee = Username And Privilege = 'ADMINISTER DATABASE TRIGGER')")
                    
    Set LoadUsers = gclsBase.OpenSQLRecord(gcnOracle, strSql, "LoadUsers")
    Exit Function
errH:
    MsgBox err.Description
End Function

Public Function CheckExist(ByVal strFields As String, ByVal strCheck As String, ByVal rsData As ADODB.Recordset) As String
'功能:检查数据集中是否含有相关记录,如果不存在 ,就返回不存在的值
'参数:strFields-需要检索的字段,strCheck-需要检索的字符串,用","作为分隔 ,rsData-数据集
    Dim strTmp() As String, i As Integer
    Dim strResult As String
    
    strTmp = Split(strCheck, ",")
    
    For i = 0 To UBound(strTmp)
        rsData.Filter = strFields & "= '" & strTmp(i) & "'"
        If rsData.RecordCount = 0 Then
            If strResult = "" Then
                strResult = strTmp(i)
            Else
                strResult = strResult & "," & strTmp(i)
            End If
            rsData.Filter = 0
        End If
    Next
    
    CheckExist = strResult
    rsData.Filter = 0
End Function

Public Function FindUser(ByVal strUser) As String
'功能:根据传入的值模糊查询用户名,如果有多条记录,返回第一条,若无记录,返回空.
    
    Dim strSql As String, rsTmp As ADODB.Recordset
    Dim strResult As String
    
    On Error GoTo errH
    strSql = "Select Username" & vbNewLine & _
                    "From Dba_Users" & vbNewLine & _
                    "Where Username Like  '" & strUser & "%'  And Not Exists" & vbNewLine & _
                    " (Select 1 From Dba_Role_Privs Where Granted_Role = 'DBA' And Grantee = Username) And Rownum = 1"
    
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSql, "FindUser")
    
    If rsTmp.RecordCount = 0 Then
        strResult = ""
    Else
        rsTmp.MoveFirst
        strResult = rsTmp!USERNAME
    End If
    
    FindUser = strResult
    Exit Function
errH:
    MsgBox err.Description
End Function

Public Function CheckRAC(ByRef intInstID As Integer) As Boolean
'功能：检查是否为RAC环境
    Dim rsTmp As ADODB.Recordset, strSql As String
 
    strSql = "Select Value From V$parameter Where Name = 'cluster_database'"
    On Error GoTo errH
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSql, "CheckRAC")
    
    If rsTmp.RecordCount > 0 And rsTmp!value = "TRUE" Then
        CheckRAC = True
        
        strSql = "Select UserENV('instance') Inst_ID From dual"
        Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSql, "CheckRAC")
        intInstID = Val("" & rsTmp!INST_ID)
    Else
        intInstID = 0
        CheckRAC = False
    End If
    
    Exit Function
errH:
    MsgBox err.Description
End Function

Public Function CheckSQLPlan(ByVal strSQLCheck As String, Optional ByRef vsPlan As VSFlexGrid, _
    Optional ByVal intConnect As Integer, Optional ByRef blnSuccess As Boolean) As Boolean
'性能问题检查:
'         1.大表全表扫描zlbigtable+zlbaktables，
'         2.中型表全表扫描(如果有统计信息，User_tab_statistics:num_rows>3000(药品目录一般是这个值以上) AND num_rows<100 0000百万以内)
'         3.大表上引用基础表(非大表)的外键上的索引
'         4.大表和中型表索引全扫描（inex full scan，INDEX FAST FULL SCAN）
'         5.大表和中型表跳跃式索引扫描（INDEX SKIP SCAN）
'返回：blnReturn=true 有性能问题

    Dim rsPlan As ADODB.Recordset
    Dim i As Long, strSql As String
    Dim j As Long, blnReturn As Boolean
    Dim rsIndex As New Recordset
    Dim rsData As ADODB.Recordset
    Dim strTable As String
    Dim rsCons_FK As New Recordset
    Dim strPar As String
    Dim strTmp As String
    Dim strAllTable As String
    
    If intConnect > 0 Then
        blnSuccess = True
        CheckSQLPlan = False
        Exit Function
    End If
    
    Set rsPlan = GetSQLPlan(strSQLCheck, intConnect)
    If Not vsPlan Is Nothing Then
        vsPlan.Redraw = flexRDNone
        vsPlan.Rows = vsPlan.FixedRows
        vsPlan.FixedAlignment(1) = flexAlignLeftCenter
    End If
    
    blnSuccess = Not rsPlan Is Nothing
    
    If Not rsPlan Is Nothing Then
        If mstrBigTable = "" Then
            '先取大表,首次进入判断是否有zltables这张表
            If mstrHasZltables = "" Then
                mstrHasZltables = CheckTblExist("ZLTABLES")
            End If
            
            '有ZLTABLES,就去B类和C类作为大表,否则取zlbigtabls和zlbaktables中的表
            If mstrHasZltables = "True" Then
               strSql = " Select Distinct 表名 From Zltables Where 分类 In ('B1', 'B2', 'B3', 'C1', 'C2', 'C3') "
            Else
                strSql = "Select Distinct 表名" & vbNewLine & _
                        "From Zlbigtables" & vbNewLine & _
                        "Union All" & vbNewLine & _
                        "Select Distinct 表名 From Zlbaktables"
            End If
           Set rsIndex = gclsBase.OpenSQLRecord(gcnOracle, strSql, App.ProductName)
            Do While Not rsIndex.EOF
                mstrBigTable = mstrBigTable & "," & rsIndex!表名
                rsIndex.MoveNext
            Loop
            mstrBigTable = mstrBigTable & ","
        End If
        
        '再取中表（统计信息，User_tab_statistics:num_rows>3000）
        strSql = "Select A.参数名,Nvl(A.参数值,A.缺省值) As 参数值 " & _
                 "From zlParameters A " & _
                 "Where A.参数名 = '检查中型表' And a.系统 is null And a.模块 is null"
        Set rsData = gclsBase.OpenSQLRecord(gcnOracle, strSql, App.ProductName)
        If rsData.BOF = False Then
            strPar = Nvl(rsData("参数值").value, "0,0")
            If strPar <> "0,0" Then
                If strPar <> mstrMiddleTableRows Then
                    strSql = "Select Table_Name as 表名 From User_Tab_Statistics Where Num_Rows > [1] And Num_Rows < [2] "
                    Set rsIndex = gclsBase.OpenSQLRecord(gcnOracle, strSql, App.ProductName, Val(Split(strPar, ",")(0)), Val(Split(strPar, ",")(1)))
                    mstrMiddleTable = ""
                    Do While Not rsIndex.EOF
                        If InStr("," & mstrBigTable & ",", "," & rsIndex!表名 & ",") = 0 Then
                            mstrMiddleTable = mstrMiddleTable & "," & rsIndex!表名
                        End If
                        rsIndex.MoveNext
                    Loop
                    mstrMiddleTable = mstrMiddleTable & ","
                    mstrMiddleTableRows = strPar
                End If
            Else
                mstrMiddleTable = ""
                mstrMiddleTableRows = ""
            End If
        Else
            mstrMiddleTable = ""
            mstrMiddleTableRows = ""
        End If
        
        strAllTable = mstrMiddleTable & mstrBigTable
        
        For i = 1 To rsPlan.RecordCount
            strTmp = UCase(rsPlan!Operation & "")
            
            If Not vsPlan Is Nothing Then
                With vsPlan
                    .addItem rsPlan!Cardinality & vbTab & Trim(rsPlan!Operation) & " " & rsPlan!name & " " & IIf(rsPlan!Bytes & "" = "" And rsPlan!cost & "" = "" And rsPlan!Time & "" = "", "", " (bytes=" & rsPlan!Bytes & " cost=" & rsPlan!cost & " time=" & Format(Time / 24 / 60 / 60, "HH:MM:SS") & ")")
                    .RowOutlineLevel(.Rows - 1) = Len(rsPlan!Operation & "") - Len(LTrim(rsPlan!Operation & ""))
                    .IsSubtotal(.Rows - 1) = True
                End With
            End If
            If InStr(strTmp, "TABLE ACCESS FULL") > 0 Then
                '判断是否是大表中表全扫描
                If InStr(strAllTable, "," & rsPlan!name & ",") > 0 Then
                    If Not vsPlan Is Nothing Then
                        vsPlan.Cell(flexcpForeColor, vsPlan.Rows - 1, 0, vsPlan.Rows - 1, vsPlan.Cols - 1) = &HFF& '红
                    End If
                    blnReturn = True
                End If
            ElseIf InStr(strTmp, "INDEX FAST FULL SCAN") > 0 Or InStr(strTmp, "INDEX FULL SCAN") > 0 Or InStr(strTmp, "INDEX SKIP SCAN") > 0 Then
                '判断是否是大表中表索引全扫描或跳跃式索引
                strTable = Split(rsPlan!name & "_", "_")(0)
                If InStr(strAllTable, "," & strTable & ",") > 0 Then
                    If Not vsPlan Is Nothing Then
                        vsPlan.Cell(flexcpForeColor, vsPlan.Rows - 1, 0, vsPlan.Rows - 1, vsPlan.Cols - 1) = &HFF& '红
                    End If
                    blnReturn = True
                End If
            ElseIf InStr(strTmp, "INDEX RANGE SCAN") > 0 Then
                '大表上使用了基础表(非大表)外键索引
                strTable = Split(rsPlan!name & "_", "_")(0)
                
                If InStr("," & mstrBigTable & ",", "," & strTable & ",") > 0 Then
                    strSql = "Select distinct d.Table_Name, d.Index_Name, d.Column_Name,d.Column_Position" & vbNewLine & _
                        "              From User_Ind_Columns D" & vbNewLine & _
                        "              Where d.Index_Name = [1] " & vbNewLine & _
                        "              Order By d.Column_Position"
                    Set rsIndex = gclsBase.OpenSQLRecord(gcnOracle, strSql, App.ProductName, rsPlan!name & "")
                    If rsIndex.RecordCount > 0 Then
                        '找外键约束
                        Set rsCons_FK = GetConsFK(strTable, rsPlan!object_owner & "")
                        strTmp = ""
                        Do While Not rsIndex.EOF
                            strTmp = strTmp & "," & rsIndex!Column_Name
                            rsIndex.MoveNext
                        Loop
                        rsCons_FK.Filter = "Column_Name='" & Mid(strTmp, 2) & "'"
                        If rsCons_FK.RecordCount > 0 Then
                            strTable = Split(rsCons_FK!r_Constraint_Name & "_", "_")(0)
                            
                            '外键父表不是大表，则视为有性能问题
                            If InStr(mstrBigTable, "," & strTable & ",") = 0 Then
                                If Not vsPlan Is Nothing Then
                                    vsPlan.Cell(flexcpForeColor, vsPlan.Rows - 1, 0, vsPlan.Rows - 1, vsPlan.Cols - 1) = &HFF& '红
                                End If
                                blnReturn = True
                            End If
                        End If
                    End If
                End If
            End If
            
            rsPlan.MoveNext
        Next
        
        If Not vsPlan Is Nothing Then
            vsPlan.CellBorderRange 0, 0, vsPlan.Rows - 1, 0, &H808080, 0, 0, 1, 0, 0, 0
            vsPlan.CellBorderRange vsPlan.FixedRows - 1, 0, vsPlan.FixedRows - 1, vsPlan.Cols - 1, &H808080, 0, 0, 0, 1, 1, 0
            vsPlan.CellBorderRange vsPlan.Rows - 1, 0, vsPlan.Rows - 1, vsPlan.Cols - 1, &H808080, 0, 0, 0, 1, 1, 0
            vsPlan.AutoSize 0, vsPlan.Cols - 1
            vsPlan.Redraw = flexRDDirect
        End If
    End If
    
    CheckSQLPlan = blnReturn
End Function

Private Function GetSQLPlan(ByVal strSQLCheck As String, Optional ByVal intConnect As Integer = 0) As ADODB.Recordset
'功能：收集SQL的执行计划

    Dim strSql As String, strSID As String
    Dim rsTmp As ADODB.Recordset
    
    If strSQLCheck <> "" Then
        
        On Error Resume Next
        strSID = Time()
          
        '执行计划
        strSql = "explain plan set statement_id = '" & strSID & "' for " & strSQLCheck
        gcnOracle.Execute strSql
        If err.Number = 0 Then
            strSql = _
                    "Select Time From Plan_Table " & vbNewLine & _
                    "Connect By Prior ID = Parent_Id And Prior Statement_Id = Statement_Id " & vbNewLine & _
                    "Start With ID = 0 And Statement_Id = [1] " & vbNewLine & _
                    "Order By ID "
            On Error Resume Next
            Set GetSQLPlan = gclsBase.OpenSQLRecord(gcnOracle, strSql, "执行计划", strSID)
            strSql = _
                    "Select ID, LPad(' ', Level - 1) || Operation || ' ' || Options As Operation, Object_Name As Name" & _
                    "    ,Object_Owner, Cardinality, Bytes" & vbNewLine & _
                    "    ,Cost" & IIf(err.Number = 0, ", Time ", ",0 as Time ") & vbNewLine & _
                    "From Plan_Table " & vbNewLine & _
                    "Connect By Prior ID = Parent_Id And Prior Statement_Id = Statement_Id " & vbNewLine & _
                    "Start With ID = 0 And Statement_Id = [1] " & vbNewLine & _
                    "Order By ID "
            err.Clear
            Set GetSQLPlan = gclsBase.OpenSQLRecord(gcnOracle, strSql, "执行计划", strSID)
            gcnOracle.Execute "Delete plan_table"
        Else
            Set GetSQLPlan = Nothing
           MsgBox err.Description: err.Clear
        End If
    End If
End Function



Public Function CheckTblExist(ByVal strTableName As String) As Boolean
    '功能：根据表名判断表是否存在
    '参数：strTableName - 要查询的表名
    Dim strSql As String, rsData As ADODB.Recordset
    
    On Error Resume Next
    strSql = "select 1 from " & strTableName & " where rownum<1 "
    Set rsData = gclsBase.OpenSQLRecord(gcnOracle, strSql, "CheckTblExist")
    
    CheckTblExist = err.Number = 0
    err.Clear
End Function

Public Function CheckAuditStatus(ByVal strModuleNo As String, ByVal strFuncName As String, ByRef strRemarks As String) As Boolean
    '功能：检查传入的功能是否需要进行审核
    'strModuleNo = 模块编号
    'strFuncName = 功能名称
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errH
    strSql = "Select 是否需审核 From Zlauditlogconfig Where 系统 Is Null And 模块 = [1] And 功能 = [2]"
    Set rsTemp = gclsBase.OpenSQLRecord(gcnOracle, strSql, "获取对应模块功能是否需要审核", strModuleNo, strFuncName)
    With rsTemp
        If .RecordCount = 0 Then
            MsgBox "管理员身份验证失败，请联系开发人员检查模块功能名称是否有误！", vbInformation, gstrSysName
            Exit Function
        End If
        If !是否需审核 = 1 Then
            If Not frmUserCheckLogin.ShowLogin(UCT_AuditLog, , gstrUserName, , , , strRemarks) Then Exit Function
        Else
            strRemarks = ""
        End If
    End With
    CheckAuditStatus = True
    Exit Function
errH:
    MsgBox err.Description, vbInformation, gstrSysName
End Function

Private Function GetConsFK(ByVal strFind As String, ByVal strOwner As String) As ADODB.Recordset
'功能：获取指定表的外键约束记录集
'参数：strFind=表名
    Dim strSql As String
    Dim rsCons As New Recordset
    Dim rsCons_FK As New Recordset

    strSql = "Select" & vbNewLine & _
        "        f.Constraint_Name, f.r_Constraint_Name,e.Column_Name,e.Position" & vbNewLine & _
        "       From User_Cons_Columns E, User_Constraints F" & vbNewLine & _
        "       Where e.Constraint_Name = f.Constraint_Name And e.owner = f.owner  And f.Constraint_Type = 'R' And f.Table_Name = [1] And f.owner = [2]" & vbNewLine & _
        "       order by e.constraint_name,e.position"
    Set rsCons = gclsBase.OpenSQLRecord(gcnOracle, strSql, App.ProductName, strFind, strOwner)
    Set rsCons_FK = New ADODB.Recordset
    rsCons_FK.Fields.Append "r_Constraint_Name", adVarChar, 50, adFldIsNullable
    rsCons_FK.Fields.Append "Constraint_Name", adVarChar, 50, adFldIsNullable
    rsCons_FK.Fields.Append "Column_Name", adVarChar, 100, adFldIsNullable
    rsCons_FK.CursorLocation = adUseClient
    rsCons_FK.LockType = adLockOptimistic
    rsCons_FK.CursorType = adOpenStatic
    rsCons_FK.Open
    Do While Not rsCons.EOF
        rsCons_FK.Filter = "Constraint_Name='" & rsCons!Constraint_Name & "'"
        If rsCons_FK.RecordCount = 0 Then
            rsCons_FK.AddNew
            rsCons_FK!Constraint_Name = rsCons!Constraint_Name & ""
            rsCons_FK!r_Constraint_Name = rsCons!r_Constraint_Name & ""
            rsCons_FK!Column_Name = rsCons!Column_Name & ""
        Else
            rsCons_FK!Column_Name = rsCons_FK!Column_Name & "," & rsCons!Column_Name
        End If
        rsCons_FK.Update
        rsCons.MoveNext
    Loop
    Set GetConsFK = rsCons_FK
End Function

Public Function CheckIpValidate(ByVal strBeginIp As String, Optional ByVal strEndIp As String, Optional ByRef strErr As String) As Boolean
    '检查IP的合法性
    'strBeginIp -开始IP strEndIp-结束IP strErr-错误信息
    Dim arrStart As Variant, arrEnd As Variant
    Dim i As Integer
    
    If Not IsNumeric(Replace(strBeginIp, ".", "")) Then
        strErr = "IP必须为数字"
        Exit Function
    End If
    arrStart = Split(strBeginIp, ".")
    If UBound(arrStart) <> 3 Then
        strErr = "IP必须由4个IP段组成"
        Exit Function
    End If

    If strEndIp <> "" Then
        If Not IsNumeric(Replace(strEndIp, ".", "")) Then
            strErr = "IP必须为数字"
            Exit Function
        End If
        arrEnd = Split(strEndIp, ".")
        If UBound(arrEnd) <> 3 Then
            strErr = "IP必须由4个IP段组成"
            Exit Function
        End If
    End If
    
'    A类IP：1.0.0.0-126.0.0.255
'    B类IP：128.1.0.0--191.254.0.255
'    C类IP：192.0.1.0--223.255.254.255
'    D类IP：224.0.0.0到239.255.255.255
    '第一段
    If arrStart(0) >= 1 And arrStart(0) <= 239 Then
        If arrEnd(0) <> "" And arrEnd(0) <> arrStart(0) Then
            strErr = "开始IP与结束IP的首段必须相同"
            Exit Function
        End If
        
        '第二段
        If arrStart(1) >= 0 And arrStart(1) <= 255 Then
            If arrEnd(1) <> "" And arrEnd(1) <> arrStart(1) Then
                strErr = "开始IP与结束IP的次段必须相同"
                Exit Function
            End If
        Else
            strErr = "IP的次段只能介于0-255之间"
            Exit Function
        End If
        
        '第三段
        If arrStart(2) >= 0 And arrStart(2) <= 255 Then
            If arrEnd(2) <> "" Then
                If arrEnd(2) >= 0 And arrEnd(2) <= 255 Then
                    If arrEnd(2) < arrStart(2) Then
                        strErr = "结束IP的第三段必须大于或等于开始IP的第三段"
                        Exit Function
                    End If
                Else
                    strErr = "IP的第三段只能介于0-255之间"
                    Exit Function
                End If
            End If
        Else
            strErr = "IP的第三段只能介于0-255之间"
            Exit Function
        End If
        
        '第四段
        If arrStart(3) > 0 And arrStart(3) <= 255 Then
            If arrEnd(3) <> "" Then
                If arrEnd(3) > 0 And arrEnd(3) <= 255 Then
                    If arrEnd(3) < arrStart(3) Then
                        strErr = "结束IP的第四段必须大于或等于开始IP的第四段"
                        Exit Function
                    End If
                Else
                    strErr = "IP的第四段只能介于1-255之间"
                    Exit Function
                End If
            End If
        Else
            strErr = "IP的第四段只能介于1-255之间"
            Exit Function
        End If
        
    Else
        strErr = "IP首段只能介于1-239之间"
        Exit Function
    End If
    
    CheckIpValidate = True
End Function


Public Function CheckProcExist(ByVal strProc As String) As Integer
    '功能:根据传入的进程名称,返回正在运行的进程数

    Dim intResult As Integer
    Dim uProcess As PROCESSENTRY32
    Dim lngMdlProcess As Long, strExeName As String, lngSnapShot As Long
    
    '创建进程快照
    lngSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
    If lngSnapShot > 0 Then
        uProcess.dwSize = Len(uProcess)
        If Process32First(lngSnapShot, uProcess) Then
            Do
                strExeName = UCase(Left(Trim(uProcess.szExeFile), InStr(1, Trim(uProcess.szExeFile), vbNullChar) - 1))
                If strExeName = UCase(strProc) Then
                    intResult = intResult + 1
                End If
            Loop Until (Process32Next(lngSnapShot, uProcess) < 1)
        End If
    End If
    
    CheckProcExist = intResult
End Function

Public Sub ShowTipInfo(ByVal lngHwnd As Long, ByVal strInfo As String, Optional blnMultiRow As Boolean, Optional blnOutline As Boolean, Optional lngMaxWidth As Long, Optional strTitle As String, Optional blnChild As Boolean)
'功能：显示或者隐藏提示
'参数：lngHwnd=提示所针对的控件句柄,当传入为0时隐藏提示
'      strInfo=提示信息,当传入为空时隐藏提示
'      blnMultiRow=以一定的间距分行显示多行信息，每行按vbcrlf分隔
'      blnOutline=是否将每行文本中字符|前的文字做为提纲单独一行显示
'      lngMaxWidth=窗口的最大窗度，缺省为0表示按设计状态的窗体最大宽度为准
'      strTitle = 提示标题
'      blnChild=是否使用ChildWindowFromPoint方法

    Call frmTipInfo.ShowTipInfo(lngHwnd, strInfo, blnMultiRow, blnOutline, lngMaxWidth, strTitle, blnChild)
End Sub

