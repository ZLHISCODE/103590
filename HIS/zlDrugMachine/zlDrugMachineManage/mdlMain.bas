Attribute VB_Name = "mdlMain"
Option Explicit

'---------------------------------------------------------------
'说明：启动过程、逻辑处理模块
'编制：余智勇
'---------------------------------------------------------------

Private Const cstBase64 = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"

Public Sub Main()
    Dim blnNew As Boolean
    
    On Error Resume Next
    
    Set gobjFile = New FileSystemObject
    If Err.Number <> 0 Then
        MsgBox "创建“FileSystemObject”部件失败，程序将立即终止，请联系管理员！", vbInformation, GSTR_MSG
        Exit Sub
    End If
    
    Set gobjRegister = Nothing
    Set gobjComLib = CreateObject("zl9ComLib.clsComLib")
    If Err.Number <> 0 Then
        MsgBox "创建“zl9ComLib”部件失败，程序将立即终止，请联系管理员！", vbInformation, GSTR_MSG
        Exit Sub
    End If
    
    Set gobjRegister = CreateObject("zlRegister.clsRegister")
    If Err.Number <> 0 Then
'        MsgBox "创建“zlRegister”部件失败，程序将立即终止，请联系管理员！", vbInformation, GSTR_MSG
'        Exit Sub
        Set gobjRegister = Nothing
        Err.Clear
    End If
    
    Set gcnOracle = New ADODB.Connection
    If Err.Number <> 0 Then
        MsgBox "“Microsoft ADO”组件未安装，请联系管理员！", vbInformation, GSTR_MSG
        Exit Sub
    End If
    
    Set gobjXML = New clsXML
    If Err.Number <> 0 Then
        MsgBox "创建“clsXML”类失败， 程序将立即终止，请联系管理员！", vbInformation, GSTR_MSG
        Exit Sub
    End If
    
    Set gobjZLPrint = CreateObject("zl9PrintMode.zlPrintMethod")
    If Err.Number <> 0 Then
        Set gobjZLPrint = Nothing
        MsgBox "创建“zl9PrintMode”部件失败，将影响打印及相关功能！", vbInformation, GSTR_MSG
    End If
    
    Set gobjEncrypt = CreateObject("zlEncryptPub.clsEncrypt")
    If Err.Number <> 0 Then
        Set gobjEncrypt = Nothing
        MsgBox "创建“zlEncryptPub”部件失败，将影响与密钥相关功能！", vbInformation, GSTR_MSG
    End If
    
    On Error GoTo 0
    
    frmLogin.Show vbModal
    
    If Not gcnOracle Is Nothing Then
        If gcnOracle.State = adStateOpen Then
            gobjComLib.InitCommon gcnOracle
            If mdlMain.GetUserInfo(gstrUser) Then
                frmMain.Show
            End If
        End If
    End If
    
End Sub

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

Public Function OraDataOpen(ByVal strServerName As String, ByVal strUserName As String, ByVal strUserPwd As String) As Boolean
    '------------------------------------------------
    '功能： 打开指定的数据库
    '参数：
    '   strServerName：主机字符串
    '   strUserName：用户名
    '   strUserPwd：密码
    '返回： 数据库打开成功，返回true；失败，返回false
    '------------------------------------------------
    Dim strSQL As String
    Dim strError As String
    Dim rs As New ADODB.Recordset
    
    On Error Resume Next
    Err = 0
    DoEvents
    With gcnOracle
        
        If .State = adStateOpen Then .Close
        .Provider = "MSDataShape"
        .Properties("Persist Security Info") = True
        .Open "Driver={Microsoft ODBC for Oracle};Server=" & strServerName, strUserName, strUserPwd
        If Err <> 0 Then
            '保存错误信息
            strError = Err.Description
            If InStr(strError, "自动化错误") > 0 Then
                MsgBox "连接串无法创建，请检查数据访问部件是否正常安装。", vbInformation, GSTR_MSG
            ElseIf InStr(strError, "ORA-12154") > 0 Then
                MsgBox "无法分析服务器名，" & vbCrLf & "请检查在Oracle配置中是否存在该本地网络服务名（主机字符串）。", vbInformation, GSTR_MSG
            ElseIf InStr(strError, "ORA-12541") > 0 Then
                MsgBox "无法连接，请检查服务器上的Oracle监听器服务是否启动。", vbInformation, GSTR_MSG
            ElseIf InStr(strError, "ORA-01033") > 0 Then
                MsgBox "ORACLE正在初始化或在关闭，请稍候再试。", vbInformation, GSTR_MSG
            ElseIf InStr(strError, "ORA-01034") > 0 Then
                MsgBox "ORACLE不可用，请检查服务或数据库实例是否启动。", vbInformation, GSTR_MSG
            ElseIf InStr(strError, "ORA-02391") > 0 Then
                MsgBox "用户" & UCase(strUserName) & "已经登录，不允许重复登录(已达到系统所允许的最大登录数)。", vbExclamation, GSTR_MSG
            ElseIf InStr(strError, "ORA-01017") > 0 Then
                MsgBox "由于用户、口令或服务器指定错误，无法登录。", vbInformation, GSTR_MSG
            ElseIf InStr(strError, "ORA-28000") > 0 Then
                MsgBox "由于用户已经被禁用，无法登录。", vbInformation, GSTR_MSG
            Else
                MsgBox strError, vbInformation, GSTR_MSG
            End If
            
            OraDataOpen = False
            Exit Function
        End If
    End With
    
    Err = 0
    On Error GoTo errHand
    
    'gstrDbUser = UCase(strUserName)
    'SetDbUser gstrDbUser
    
    OraDataOpen = True
    Exit Function
    
errHand:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Err.Clear
End Function

Public Function TranPasswd(strOld As String) As String
    '------------------------------------------------
    '功能： 密码转换函数
    '参数：
    '   strOld：原密码
    '返回： 加密生成的密码
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

'Public Function GetUserInfo() As Boolean
''功能：获取登陆用户信息
''返回：True成功；False失败
'    Dim rsTmp As ADODB.Recordset
'
'    On Error GoTo hErr
'
'    UserInfo.姓名 = UserInfo.用户名
'    Set rsTmp = mdlMain.GetUserInfo
'    If Not rsTmp Is Nothing Then
'        If Not rsTmp.EOF Then
'            UserInfo.ID = rsTmp!ID
'            UserInfo.编号 = rsTmp!编号
'            UserInfo.部门ID = gobjComLib.zlCommFun.NVL(rsTmp!部门ID, 0)
'            UserInfo.简码 = gobjComLib.zlCommFun.NVL(rsTmp!简码)
'            UserInfo.姓名 = gobjComLib.zlCommFun.NVL(rsTmp!姓名)
'            UserInfo.用户名 = rsTmp!用户名
'            GetUserInfo = True
'        End If
'        rsTmp.Close
'    End If
'
'    Exit Function
'
'hErr:
'    If gobjComLib.ErrCenter = 1 Then Resume
'End Function

Public Function GetUserInfo(ByVal strDBUser As String) As Boolean
'功能：获取当前用户的基本信息
'返回：返回Ado记录集
    Dim strSQL As String, strDefault As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo hErr
    
    strDefault = " And C.缺省 = 1"
    strSQL = "Select User,A.Id, A.编号, A.简码, A.姓名, A.专业技术职务,B.用户名, C.部门id, D.编码 As 部门码, D.名称 As 部门名 " & vbNewLine & _
             "From 人员表 A, 上机人员表 B, 部门人员 C, 部门表 D " & vbNewLine & _
             "Where A.Id = B.人员id And A.Id = C.人员id And C.部门id = D.Id And B.用户名 = [1] "
    Set rsTemp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL & strDefault, "GetUserInfo", UCase(strDBUser))
    If rsTemp.RecordCount = 0 Then
        strDefault = " And Rownum < 2"
        Set rsTemp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL & strDefault, "GetUserInfo", UCase(strDBUser))
    End If
    If rsTemp.RecordCount > 0 Then
        UserInfo.ID = rsTemp!ID
        UserInfo.编号 = rsTemp!编号
        UserInfo.部门ID = gobjComLib.zlCommFun.NVL(rsTemp!部门ID, 0)
        UserInfo.简码 = gobjComLib.zlCommFun.NVL(rsTemp!简码)
        UserInfo.姓名 = gobjComLib.zlCommFun.NVL(rsTemp!姓名)
        UserInfo.用户名 = rsTemp!用户名
        GetUserInfo = True
    End If
    rsTemp.Close
    
    Exit Function
    
hErr:
    If gobjComLib.ErrCenter() = 1 Then Resume
End Function

Public Function OpenIme(Optional blnOpen As Boolean = False) As Boolean
'功能:打开中文输入法，或关闭输入法
    Dim arrIme(99) As Long, lngCount As Long, strName As String * 255
    Dim strIme As String, blnNotCloseIme As Boolean
    
    lngCount = GetKeyboardLayoutList(UBound(arrIme) + 1, arrIme(0))
    blnNotCloseIme = True
    Do
        lngCount = lngCount - 1
        If ImmIsIME(arrIme(lngCount)) = 1 Then
            If blnOpen = True Then
                '需要打开输入法。接着判断是否批定输入法
                ImmGetDescription arrIme(lngCount), strName, Len(strName)
                If InStr(1, Mid(strName, 1, InStr(1, strName, Chr(0)) - 1), strIme) > 0 And strIme <> "" Then
                    If ActivateKeyboardLayout(arrIme(lngCount), 0) <> 0 Then OpenIme = True: Exit Function
                End If
            End If
        ElseIf blnOpen = False Then
            '不是输入法，正好是应了关闭输入法的请求
            If ActivateKeyboardLayout(arrIme(lngCount), 0) <> 0 Then OpenIme = True: Exit Function
        End If
    Loop Until lngCount = 0
    
    If blnNotCloseIme And blnOpen = False Then
        '由于windows Vista系统的英文输入法用ImmIsIME测试出是true的输入法,因此,需要单独处理.
        '刘兴宏:2008/09/03
        If ActivateKeyboardLayout(arrIme(0), 0) <> 0 Then OpenIme = True: Exit Function
    End If
End Function

Public Sub LoadServer(ByVal cbxVar As ComboBox, ByRef colVar As Collection)
'功能：读出本地的服务器列表
    Dim strPath As String, strFile As String, lngFile As Integer
    Dim strLine As String, lngPos As Long
    Dim strServer As String, strComputer As String, strSID As String
    Dim blnFinish As Boolean
    
    cbxVar.Clear
    
'    strPath = GetKeyValue(HKEY_LOCAL_MACHINE, "SOFTWARE\ORACLE", "ORACLE_HOME")
'    If Not gobjFile.FolderExists(strPath) Then '10G
'        strPath = GetKeyValue(HKEY_LOCAL_MACHINE, "SOFTWARE\ORACLE", "ORA_CRS_HOME")
'    End If
'    If Not gobjFile.FolderExists(strPath) Then '10Gr2
'        strPath = GetKeyValue(HKEY_LOCAL_MACHINE, "SOFTWARE\ORACLE\KEY_OraDb10g_home1", "ORACLE_HOME")
'    End If
'    If Not gobjFile.FolderExists(strPath) Then '10Gr2
'        strPath = GetKeyValue(HKEY_LOCAL_MACHINE, "SOFTWARE\ORACLE\KEY_OraDb10g_home2", "ORACLE_HOME")
'    End If
'    If Not gobjFile.FolderExists(strPath) Then    '10G 企业版
'        strPath = GetKeyValue(HKEY_LOCAL_MACHINE, "SOFTWARE\ORACLE\KEY_OraClient10g_home1", "ORACLE_HOME")
'    End If
'    If Not gobjFile.FolderExists(strPath) Then    '10G 企业版
'        strPath = GetKeyValue(HKEY_LOCAL_MACHINE, "SOFTWARE\ORACLE\KEY_OraClient10g_home2", "ORACLE_HOME")
'    End If
'    If Not gobjFile.FolderExists(strPath) Then '11Gr2
'        strPath = GetKeyValue(HKEY_LOCAL_MACHINE, "SOFTWARE\ORACLE\KEY_OraDb11g_home1", "ORACLE_HOME")
'    End If
'    If Not gobjFile.FolderExists(strPath) Then '11Gr2
'        strPath = GetKeyValue(HKEY_LOCAL_MACHINE, "SOFTWARE\ORACLE\KEY_OraDb11g_home2", "ORACLE_HOME")
'    End If
'    strFile = strPath & "\network\ADMIN\tnsnames.ora" 'Oracle 8i以上
'    If Not gobjFile.FileExists(strFile) Then
'        strFile = strPath & "\NET80\ADMIN\tnsnames.ora" 'Oracle 8
'        If Not gobjFile.FileExists(strFile) Then Exit Sub
'    End If

    '遍历注册表，获取Oracle安装路径
    strPath = GetRegItemValue(HKEY_LOCAL_MACHINE, "SOFTWARE\ORACLE", blnFinish)
    strFile = strPath & "\network\ADMIN\tnsnames.ora" 'Oracle 8i以上
    If Not gobjFile.FileExists(strFile) Then
        strFile = strPath & "\NET80\ADMIN\tnsnames.ora" 'Oracle 8
        If Not gobjFile.FileExists(strFile) Then Exit Sub
    End If

    Set colVar = New Collection
    
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
                If InStr(strLine, "PROTOCOL = TCP") > 0 And strLine Like "*PORT = 152[0-9]*" Then
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
                        colVar.Add Array(strServer, strComputer, strSID)
                        cbxVar.AddItem strServer
                    End If
                End If
            End If
        End If
    Loop
    Close #lngFile
End Sub

Public Function GetParameter(ByVal objXML As clsXML, ByVal strName As String, Optional ByVal strDefaultVal As String) As String
'功能：从zlDrugMachine.cfg文件中获取指定参数的值
'参数：
'  objXML：cfg文件的内容加载后的XML对象
'  strName：参数名称，即：XML结点名称
'返回：参数值

    Dim strValue As String

    If objXML Is Nothing Then Exit Function
    
    strName = LCase(strName)
    
    If objXML.GetSingleNodeValue(strName, strValue) Then
        GetParameter = strValue
    Else
        GetParameter = strDefaultVal
    End If

End Function

Public Function VerifyConfigFile(ByVal strFile As String) As Boolean
'功能：检查配置文档是否存在，不存在就自动创建
'参数：
'返回：True检查成功；False检查失败

    Dim fsoFile As New FileSystemObject
    Dim tsmFile As TextStream
    
    On Error GoTo hErr
    
    If fsoFile.FileExists(strFile) = False Then
        '创建配置文档
        Set tsmFile = fsoFile.CreateTextFile(strFile)
        
        '默认生成文档内容
        With tsmFile
            .WriteLine "<root>"
            .WriteLine "    <log>"
            .WriteLine "        <output>0</output>"
            .WriteLine "        <detailed>0</detailed>"
            .WriteLine "        <savedays>7</savedays>"
            .WriteLine "    </log>"
            .WriteLine "    <timer>"
            .WriteLine "        <enabled>0</enabled>"
            .WriteLine "        <businessdata></businessdata>"
            .WriteLine "        <cycle>5</cycle>"
            .WriteLine "        <validdays>2</validdays>"
            .WriteLine "        <viewlines>200</viewlines>"
            .WriteLine "    </timer>"
            .WriteLine "</root>"
        End With
        tsmFile.Close
    End If
    
    VerifyConfigFile = True
    Exit Function
    
hErr:
    Call gobjComLib.ErrCenter
End Function

Public Sub SetTextMaxLen(ByRef txtVal As TextBox, ByVal strTableField As String)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errHandle
'    gstrSQL = zlStr.FormatString("Select [2] as 字段 From [1] Where Rownum < 1 ", _
'                        CStr(Split(strTableField, ".")(0)), _
'                        CStr(Split(strTableField, ".")(1)))
    strSQL = "Select " & Split(strTableField, ".")(1) & " as 字段 From " & Split(strTableField, ".")(0) & " Where Rownum < 1 "
    
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "获取字段信息")
    txtVal.MaxLength = rsTmp.Fields(0).DefinedSize
    rsTmp.Close

    Exit Sub
    
errHandle:
    If gobjComLib.ErrCenter = 1 Then Resume
End Sub

Public Sub CreateSOAP(ByRef objSOAP As Object)
    On Error Resume Next
    Set objSOAP = Nothing
    Set objSOAP = CreateObject("MSSOAP.SoapClient30")
    If Err.Number <> 0 Then
        Err.Clear
        Set objSOAP = CreateObject("MSSOAP.SoapClient")
        If Err.Number <> 0 Then
            MsgBox "实例化“SoapClient”失败，请联系技术人员！" & vbCrLf & _
                   "注意：SoapClient在WinXP下安装2.0版本。", _
                   vbInformation, GSTR_MSG
        End If
    End If
    On Error GoTo 0
End Sub

Public Sub CreateHTTP(ByRef objHTTP As Object)
    On Error Resume Next
    Set objHTTP = Nothing
    Set objHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
    If Err.Number <> 0 Then
        Err.Clear
        MsgBox "实例化“WinHttp”失败，请联系技术人员！", vbInformation, GSTR_MSG
    End If
    On Error GoTo 0
End Sub

Public Function GetControlRect(ByVal lnghwnd As Long, Optional ByVal blnTwip As Boolean = True) As RECT
'功能：获取指定控件在屏幕中的位置(Twip/Pixel)
'返回：blnTwip=True-返回Twip单位，False-返回像素单位

    Dim vRect As RECT
    
    Call GetWindowRect(lnghwnd, vRect)
    If blnTwip Then
        vRect.Left = vRect.Left * Screen.TwipsPerPixelX
        vRect.Right = vRect.Right * Screen.TwipsPerPixelX
        vRect.Top = vRect.Top * Screen.TwipsPerPixelY
        vRect.Bottom = vRect.Bottom * Screen.TwipsPerPixelY
    End If
    GetControlRect = vRect
End Function

Public Function FormatString(ByVal strFormat As String, ParamArray arrParams() As Variant) As String
'功能：格式化字符串
'参数：
'  strFormat：表达式；[1-x]为参数号关键字；例子："测试值为：[1]"
'  arrParams：表达式的参数，对应strFormat中的参数号关键字
'返回：格式化后的字符串

    Dim i As Integer, intSN As Integer
    Dim strKey As String, strTmp As String
    Dim blnStart As Boolean

    FormatString = strFormat

    If Len(strFormat) > 60000 Then Exit Function
    If Not strFormat Like "*[[]*[]]*" Then Exit Function
    If UBound(arrParams) < 0 Then Exit Function

    On Error GoTo errHandle

    For i = 1 To Len(strFormat)
        If Mid(strFormat, i, 1) = "[" Then
            blnStart = True
        End If
        If blnStart Then
            If Mid(strFormat, i, 1) = "]" Then
                intSN = Val(Mid(strKey, 2))
                If intSN > 0 Then
                    If UBound(arrParams) >= intSN - 1 Then
                        strTmp = strTmp & arrParams(intSN - 1)
                    End If
                Else
                    strTmp = strTmp & Mid(strKey, 2)
                End If
                blnStart = False
                strKey = ""
            Else
                strKey = strKey & Mid(strFormat, i, 1)
            End If
        Else
            strTmp = strTmp & Mid(strFormat, i, 1)
        End If
    Next

    FormatString = strTmp
    Exit Function

errHandle:
End Function

Public Function VerifyString(ByVal strTarget As String, ByVal strStandard As String, _
    Optional ByVal blnStandard As Boolean = True) As Boolean
    
'功能：检查目标字符串有无非标准字符
'参数：
'  strStandard：标准的字符集
'  strTarget：要检查的目标字符串
'  blnStandard：True-strStandard为标准字符；False-strStandard为非标准字符
'返回：True通过；False未通过

    Dim i As Integer
    
    For i = 1 To Len(strTarget)
        If blnStandard Then
            If InStr(strStandard, Mid(strTarget, i, 1)) <= 0 Then
                VerifyString = False
                Exit Function
            End If
        Else
            If InStr(strStandard, Mid(strTarget, i, 1)) > 0 Then
                VerifyString = False
                Exit Function
            End If
        End If
    Next
    
    VerifyString = True

End Function

Public Function Encrypt(ByVal strSource As String) As String
'加密
    Dim BLowData As Byte
    Dim BHigData As Byte
    Dim i As Long
    Dim k As Integer
    Dim strEncrypt As String
    Dim StrChar As String
    Dim KeyTemp As String
    Dim Key1 As Byte
    
    For k = 1 To 30
        KeyTemp = KeyTemp & CStr(Int(Rnd * (9) + 1))
    Next
    
    Key1 = CByte(Mid(KeyTemp, 11, 1) & Mid(KeyTemp, 27, 1))
    
    For i = 1 To Len(strSource)
        StrChar = Mid(strSource, i, 1)                                      '从待加密字符串中取出一个字符
        BLowData = AscB(MidB(StrChar, 1, 1)) Xor Key1                       '取字符的低字节和Key1进行异或运算
        BHigData = AscB(MidB(StrChar, 2, 1))                                '取字符的高字节
        strEncrypt = strEncrypt & ChrB(BLowData) & ChrB(BHigData)       '将运算后的数据合成新的字符
    Next i
    
    Encrypt = KeyTemp & strEncrypt
End Function

Public Function Decrypt(ByVal strSource As String) As String
'解密
    Dim BLowData As Byte
    Dim BHigData As Byte
    Dim i As Long
    Dim k As Integer
    Dim StrDecrypt As String
    Dim StrChar As String
    Dim KeyTemp As String
    Dim Key1 As Byte
    
    KeyTemp = Mid(strSource, 1, 30)
    Key1 = CByte(Mid(KeyTemp, 11, 1) & Mid(KeyTemp, 27, 1))
    
    For i = 31 To Len(strSource)
        StrChar = Mid(strSource, i, 1)                                      '从待解密字符串中取出一个字符
        BLowData = AscB(MidB(StrChar, 1, 1)) Xor Key1                       '取字符的低字节和Key1进行异或运算
        BHigData = AscB(MidB(StrChar, 2, 1))                                '取字符的高字节
        StrDecrypt = StrDecrypt & ChrB(BLowData) & ChrB(BHigData)       '将运算后的数据合成新的字符
    Next i
    
    Decrypt = StrDecrypt
End Function

Public Function Base64Encode(strSource As String) As String
    Dim arrBase64() As String
    Dim arrB() As Byte, bTmp(2) As Byte, bT As Byte
    Dim i As Long, j As Long
    
    On Error Resume Next
    
    If UBound(arrBase64) = -1 Then
        arrBase64 = Split(StrConv(cstBase64, vbUnicode), vbNullChar)
    End If
    
    arrB = StrConv(strSource, vbFromUnicode)

    j = UBound(arrB)
    For i = 0 To j Step 3
        Erase bTmp
        bTmp(0) = arrB(i + 0)
        bTmp(1) = arrB(i + 1)
        bTmp(2) = arrB(i + 2)

        bT = (bTmp(0) And 252) / 4
        Base64Encode = Base64Encode & arrBase64(bT)

        bT = (bTmp(0) And 3) * 16
        bT = bT + bTmp(1) \ 16
        Base64Encode = Base64Encode & arrBase64(bT)

        bT = (bTmp(1) And 15) * 4
        bT = bT + bTmp(2) \ 64
        If i + 1 <= j Then
            Base64Encode = Base64Encode & arrBase64(bT)
        Else
            Base64Encode = Base64Encode & "="
        End If

        bT = bTmp(2) And 63
        If i + 2 <= j Then
            Base64Encode = Base64Encode & arrBase64(bT)
        Else
            Base64Encode = Base64Encode & "="
        End If
    Next
End Function

Public Function Base64Decode(strEncoded As String) As String '??
    Dim arrB() As Byte, bTmp(3) As Byte, bT As Long, bRet() As Byte
    Dim i As Long, j As Long
    
    On Error Resume Next
    
    arrB = StrConv(strEncoded, vbFromUnicode)
    j = InStr(strEncoded & "=", "=") - 2
    ReDim bRet(j - j \ 4 - 1)
    For i = 0 To j Step 4
        Erase bTmp
        bTmp(0) = (InStr(cstBase64, Chr(arrB(i))) - 1) And 63
        bTmp(1) = (InStr(cstBase64, Chr(arrB(i + 1))) - 1) And 63
        bTmp(2) = (InStr(cstBase64, Chr(arrB(i + 2))) - 1) And 63
        bTmp(3) = (InStr(cstBase64, Chr(arrB(i + 3))) - 1) And 63

        bT = bTmp(0) * 2 ^ 18 + bTmp(1) * 2 ^ 12 + bTmp(2) * 2 ^ 6 + bTmp(3)

        bRet((i \ 4) * 3) = bT \ 65536
        bRet((i \ 4) * 3 + 1) = (bT And 65280) \ 256
        bRet((i \ 4) * 3 + 2) = bT And 255
    Next
    Base64Decode = StrConv(bRet, vbUnicode)
End Function

Public Function GetRegItemValue(ByVal lngKey As Long, ByVal strSubKey As String, _
    ByRef blnFinish As Boolean) As String
    
'功能：遍历子目录下特定的项目名称的值
'参数：
'  lngKey：注册表主键
'  strSubKey：注册表目录名
'返回：指定项目的值
    
    Const STR_HOME_KEY_1 As String = "ORACLE_HOME"
    Const STR_HOME_KEY_2 As String = "ORA_CRS_HOME"

    Dim lngRet As Long, lngResult As Long, lngLen As Long, lngIndex As Long, lngReserved As Long, lngClass As Long
    Dim strName As String, strClass As String, strResult As String, strTmp As String
    Dim LWT As FILETIME
    Dim blnTemp As Boolean
    
    lngRet = RegOpenKey(lngKey, strSubKey, lngResult)
    
    Do While lngRet = ERROR_SUCCESS
        strName = String(255, Chr(0))
        lngLen = Len(strName)
        lngRet = RegEnumKeyEx(lngResult, lngIndex, strName, lngLen, lngReserved, strClass, lngClass, LWT)
        If lngRet = ERROR_SUCCESS Then
            strName = Left(strName, InStr(strName, Chr(0)) - 1)
            strTmp = strSubKey & "\" & strName
'Debug.Print strTmp
            strResult = GetRegItemValue(lngKey, strTmp, blnTemp)
            If strResult = "" Then
                '无子目录时，开始找项目和项目值
                strResult = GetKeyValue(lngKey, strTmp, STR_HOME_KEY_1)
                If strResult <> "" Then
                    GetRegItemValue = strResult
                    blnFinish = True
                    Exit Do
                Else
                    strResult = GetKeyValue(lngKey, strTmp, STR_HOME_KEY_2)
                    If strResult <> "" Then
                        GetRegItemValue = strResult
                        blnFinish = True
                        Exit Do
                    End If
                End If
            ElseIf blnFinish Then
                GetRegItemValue = strResult
                Exit Do
            End If
        End If
        lngIndex = lngIndex + 1
    Loop
    
    Call RegCloseKey(lngRet)

End Function
