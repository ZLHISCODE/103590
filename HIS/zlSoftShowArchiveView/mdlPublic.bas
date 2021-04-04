Attribute VB_Name = "mdlPublic"
Option Explicit

'调用病案查阅   1:ZLHIS:83:1:0:0
'调用门诊医嘱   2:ZLHIS:1:1:0:0
'调用住院医嘱   3:ZLHIS:83:1:0:0
'调用PACS报告   4:ZLHIS:83:1:1:1008

Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

'进程间传递内存空间，可以传字符串
Public Type COPYDATASTRUCT
  dwData As Long
  cbData As Long
  lpData As Long
End Type

Public Const SW_RESTORE = 9
Public Const GWL_STYLE = (-16)
Public Const WS_SYSMENU = &H80000
Public Const WS_MAXIMIZE = &H1000000
Public Const WS_MAXIMIZEBOX = &H10000
Public Const WS_MINIMIZEBOX = &H20000
Public Const WS_CAPTION = &HC00000
Public Const WS_THICKFRAME = &H40000
Public Const WS_CHILD = &H40000000
Public C_LOG As Long  '是否记录日志,0不生成日志，1要记录生成日志

'消息Hook变量
Public plngPreWndProc As Long       '原来的消息处理程序
Public Const MSG_SPLIT = ":"

Private mobjRegister As Object                  '10.35.10之后的注册对象

Public Enum LogType
    ltError = 0
    ltDebug = 1
End Enum

Public gstrZLHIS主机字符串 As String
Public gstr用户名 As String
Public gstr密码 As String
Public gbln是否转换密码 As Boolean

'公共参数
Public gblnXWRISInterfaceLog As Boolean         '向数据库中写入日志
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Function GetRandom(ByVal lngBase As Long) As String
'------------------------------------------------
'功能：获取一个1到（lngBase-1）之间的随机数，返回它的ASCII码
'参数： lngBase  --  随机数的最大值，最大随机数为（lngBase-1）
'返回：返回它的ASCII码
'------------------------------------------------
    Dim lngNum As Long
    
    Randomize
    
    lngNum = Fix(Rnd * lngBase)
    
    If lngNum <= 0 Then lngNum = 1
    
    GetRandom = Chr(lngNum)
End Function


Public Function getEncryptionWord(ByVal strPassW As String) As String
'------------------------------------------------
'功能：获取加密密文，使用1-29之间的ASCII码，作为随机数，给字符串加密，这个算法只能适用于1-29的ASCII码，超过范围会导致解密失败
'参数： strPassW  --  需要加密的源文
'返回：返回它的密文
'------------------------------------------------
    Dim i As Integer
    Dim lngAsc  As Long
    Dim strTemp() As String
    Dim lngPassWLength As Integer
    Dim strRandom As String
    Dim strBase As String
        
    i = 0
    
    lngPassWLength = Len(strPassW)
    
    strBase = GetRandom(30)
    strRandom = GetRandom(30)
    
    '如果strBase=strRandom，加密后的密文，会出现原文，所以要确保这两个值不相同
    If strRandom = strBase Then
        If Asc(strBase) >= 29 Then
            strRandom = Chr(1)
        Else
            strRandom = Chr(Asc(strRandom) + 1)
        End If
    End If
    
    
    ReDim intAsc(0 To lngPassWLength - 1), strTemp(0 To lngPassWLength - 1)
     
    Do While i < lngPassWLength
        lngAsc = Asc(Mid(strPassW, i + 1, 1))
        lngAsc = lngAsc Xor Asc(strBase) Xor Asc(strRandom)
        strTemp(i) = Chr(lngAsc)
        i = i + 1
    Loop
    
    getEncryptionWord = strBase & Join(strTemp, "") & strRandom '加密后的字串
End Function

Public Function getDecryptionWord(ByVal strPassW As String) As String
'------------------------------------------------
'功能：获取解密的源文
'参数： strPassW  --  需要解密的密文
'返回：返回它的源文
'------------------------------------------------
    Dim i As Integer
    Dim lngAsc  As Integer
    Dim strTemp() As String
    Dim lngPassWLength As Integer
    Dim lngBase As Long
    Dim strRandom As String
    Dim strPassSouce As String

    i = 0
    
    strPassSouce = Mid(strPassW, 2, Len(strPassW) - 2)
    lngPassWLength = Len(strPassSouce)
    lngBase = Asc(Mid(strPassW, 1, 1))
    
    strRandom = Right(strPassW, 1)
    
    ReDim intAsc(0 To lngPassWLength - 1), strTemp(0 To lngPassWLength - 1)
    
    Do While i < lngPassWLength
        lngAsc = Asc(Mid(strPassSouce, i + 1, 1))
        lngAsc = lngAsc Xor Asc(strRandom) Xor lngBase
        strTemp(i) = Chr(lngAsc)
        i = i + 1
    Loop

    getDecryptionWord = Join(strTemp, "") '解密后的字串
End Function

Public Function errHandle(errSubName As String, errTitle As String, Optional errDesc As String = "") As Long
'------------------------------------------------
'功能：错误处理
'参数： logSubName  --  产生错误的函数名
'       logTitle   -- 错误名称
'       logDesc   --  错误描述
'返回：1-程序继续Resume；0-程序退出
'------------------------------------------------
    
    errHandle = 0
    
    '提示错误
    MsgBox errTitle & errDesc, vbOKOnly, "接口zlSoftShowHisForms出现错误"
    
    '清除错误
    err.Clear
    
End Function

Public Function ConnectDB(ByVal strDBUser As String) As Boolean
'------------------------------------------------
'功能：连接数据库，从注册表中读取加密后的数据库连接信息：用户名，密码，服务名
'参数：
'返回：True-成功；False-失败
'------------------------------------------------
    Dim strDBPassword As String
    Dim strDBServer As String
    Dim blnTransPassword As Boolean
    
    ConnectDB = False
    
    On Error GoTo err
    
    If gcnOracle.State <> adStateOpen Then
        strDBServer = gstrZLHIS主机字符串
        strDBUser = gstr用户名
        strDBPassword = gstr密码
        blnTransPassword = gbln是否转换密码
                
        '连接数据库
        If OraDataOpen(strDBServer, strDBUser, strDBPassword, blnTransPassword) = False Then
           
            Exit Function
        End If
    End If
    
    ConnectDB = True
    Exit Function
err:
    If errHandle("zlSoftShowHisForms.ConnectDB", "连接数据库函数出现错误", err.Description) = 1 Then Resume
End Function

Private Function OraDataOpen(ByVal strServerName As String, ByVal strUserName As String, ByVal strUserPwd As String, ByVal blnTransPassword As Boolean) As Boolean
    '------------------------------------------------
    '功能： 打开指定的数据库
    '参数：
    '   strServerName：主机字符串
    '   strUserName：用户名
    '   strUserPwd：密码
    '   blnTransPassword ： 是否需要转换密码
    '返回： 数据库打开成功，返回true；失败，返回false
    '------------------------------------------------
    Dim strError As String
    
    On Error GoTo ErrHand
    

    If gblnBefore3510 = True Then
        '如果是10.35.10之前的版本，直接用用户名和密码登录数据库
        OraDataOpen = OpenOracle(gcnOracle, strServerName, strUserName, IIf(UCase(strUserName) = "SYS" Or UCase(strUserName) = "SYSTEM", strUserPwd, IIf(blnTransPassword = True, TranPasswd(strUserPwd), strUserPwd)))
    Else
        '如果是10.35.10之后的版本，使用zlRegister获取数据库连接
        Set gcnOracle = mobjRegister.GetConnection(strServerName, strUserName, strUserPwd, blnTransPassword, , strError, True)
        If gcnOracle.State = adStateOpen Then
            OraDataOpen = True
        Else
            OraDataOpen = False
        End If
    End If
    
    If OraDataOpen = True Then
        gstrDBUser = UCase(strUserName) '这里为什么要强制大写？是不是comlib的要求？
        If gblnBefore3510 = True Then
            '10.35.10之前的版本
            gzlComLib.SetDbUser gstrDBUser
        End If
    End If
    
    Exit Function
    
ErrHand:
    
    If errHandle("zlSoftShowHisForms.OraDataOpen", "连接数据库出错", err.Description) = 1 Then Resume
    OraDataOpen = False
End Function

Private Function OpenOracle(ByRef cnOrcle As ADODB.Connection, ByVal strServerName As String, ByVal strUserName As String, ByVal strUserPwd As String) As Boolean
    '------------------------------------------------
    '功能： 打开指定的Oracle数据库
    '参数：
    '   cnOrcle ：数据库连接
    '   strServerName：主机字符串
    '   strUserName：用户名
    '   strUserPwd：密码
    '返回： 数据库打开成功，返回true；失败，返回false
    '------------------------------------------------
    Dim strError As String
    
    On Error Resume Next
    err = 0
    DoEvents
    With cnOrcle
        If .State = adStateOpen Then .Close
        .Provider = "MSDataShape"
        .Open "Driver={Microsoft ODBC for Oracle};Server=" & strServerName, strUserName, strUserPwd
        If err <> 0 Then
            '保存错误信息
            strError = err.Description
            If InStr(strError, "自动化错误") > 0 Then
                MsgBox "连接串无法创建，请检查数据访问部件是否正常安装。", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-12154") > 0 Then
                MsgBox "无法分析服务器名，" & vbCrLf & "请检查在Oracle配置中是否存在该本地网络服务名（主机字符串）。", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-12541") > 0 Then
                MsgBox "无法连接，请检查服务器上的Oracle监听器服务是否启动。", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-01033") > 0 Then
                MsgBox "ORACLE正在初始化或在关闭，请稍候再试。", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-01034") > 0 Then
                MsgBox "ORACLE不可用，请检查服务或数据库实例是否启动。", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-02391") > 0 Then
                MsgBox "用户" & UCase(strUserName) & "已经登录，不允许重复登录(已达到系统所允许的最大登录数)。", vbExclamation, gstrSysName
            ElseIf InStr(strError, "ORA-01017") > 0 Then
                MsgBox "由于用户、口令或服务器指定错误，无法登录。", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-28000") > 0 Then
                MsgBox "由于用户已经被禁用，无法登录。", vbInformation, gstrSysName
            Else
                MsgBox strError, vbInformation, gstrSysName
            End If
            
            OpenOracle = False
            Exit Function
        End If
    End With
    
    OpenOracle = True
    err = 0
    
    Exit Function
    
End Function

Private Function TranPasswd(strOld As String) As String
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

Public Sub ShowSubWindow(ByVal lngHwnd As Long, Optional ByVal lngMainHwnd As Long)
'功能：显示指定的窗体，以子窗体方式
'参数：lngHwnd=要作为子窗体显示的窗体的句柄
'      lngMainHwnd=父窗体句柄，不传时表明不以它的子窗体显示
'说明：该项函数主要用于在ZLBH中融合调用ZLHIS窗体显示
    Dim vRect1 As RECT, vRect2 As RECT
    Dim X As Long, Y As Long
    
    If lngHwnd <= 0 Then Exit Sub
    
    If lngMainHwnd <> 0 Then
        SetParent lngHwnd, lngMainHwnd
    Else
        SetParent lngHwnd, 0
    End If

    '显示在父窗体的中央
    If IsWindowVisible(lngHwnd) = 0 Then
        GetWindowRect lngHwnd, vRect1
        GetWindowRect lngMainHwnd, vRect2
        
        X = ((vRect2.Right - vRect2.Left) - (vRect1.Right - vRect1.Left)) / 2
        Y = ((vRect2.Bottom - vRect2.Top) - (vRect1.Bottom - vRect1.Top)) / 2
        If X < 0 Then X = 0
        If Y < 0 Then Y = 0
        
        SetWindowPos lngHwnd, 0, X, Y, 0, 0, &H40 Or &H1 'HWND_TOP=0
    End If
    
    ShowWindow lngHwnd, SW_RESTORE
End Sub

Public Function UpdateEmrInterface() As Object
    Dim objEmr As Object
    Dim strDBPassword As String
    Dim strDBServer As String
    Dim blnOut As Boolean
 
    On Error Resume Next

    writeTestLog "UpdateEmrInterface ** strDBServer=" & gstrZLHIS主机字符串 & ",strDBPassword=" & gstr密码
    Set objEmr = CreateObject("zl9EmrInterface.ClsEmrInterface")
    If Not objEmr Is Nothing Then
        '从注册表读取数据库连接信息
        strDBServer = gstrZLHIS主机字符串
        strDBPassword = gstr密码 '3510版后密码转换发生变化统一用未转换的密码
        
        If objEmr.CheckUpdate1(gstr用户名, strDBPassword, True) = False Then
           blnOut = False
        Else
            blnOut = True
        End If
        If err.Number <> 0 Then
            err.Clear
            If objEmr.CheckUpdate(gstrDBUser, strDBPassword) = False Then
                blnOut = False
            Else
                blnOut = True
            End If
        End If
     End If
  If blnOut Then
    Set UpdateEmrInterface = objEmr
   End If
End Function

Public Function InitInterface(ByVal strDBUser As String) As Boolean
'------------------------------------------------
'功能：初始化接口，创建ComLib，连接数据库
'参数：无
'返回：True-成功；False-失败
'------------------------------------------------
    
    On Error GoTo err
    InitInterface = False
    
    '初始化系统号为100，模块号为1287
    glngSys = 100
    glngModule = 1287
 
On Error Resume Next
    If mobjRegister Is Nothing Then
        Set mobjRegister = GetObject("", "zlRegister.clsRegister")
        If mobjRegister Is Nothing Then gblnBefore3510 = True '35.10之前的版本
    End If
    
    err.Clear
On Error GoTo err
    If gzlComLib Is Nothing Then
        If gblnBefore3510 Then
            '10.35.10之前的版本
            Set gzlComLib = CreateObject("zl9ComLib.clsComLib")
        Else
            '10.35.10之后的版本
            Set gzlComLib = GetObject("", "zl9ComLib.clsComLib")
        End If
    End If
    
    '如果是从RIS启动的DLL，数据库连接gzlComLib.CurrentConn是空的，需要从注册表读取用户名密码，并且连接数据库
    If gzlComLib.CurrentConn Is Nothing Then
        '从注册表读取用户名密码，连接数据库
        
        '如果gcnOracle不存在，要新建一个
        If gcnOracle Is Nothing Then Set gcnOracle = New ADODB.Connection
        Call ConnectDB(strDBUser)

        '初始化公共部件
        gzlComLib.InitCommon gcnOracle
        

        If gblnBefore3510 = True Then
            '10.35.10之前的版本
            If gzlComLib.RegCheck = False Then
                
                Exit Function
            End If
        End If
    Else
        '如果是从HIS导航台启动的DLL，则创建zl9ComLib之后，会自动包含有gzlComLib.CurrentConn
        '现在暂时没有从 CodeMan中取得 gcnOracle，所以需要从zl9ComLib取得gcnOracle对象
        'gstrDBUser从注册表中读取，见方法clsHISInner.SaveDBConnectInfo
        
        If gcnOracle Is Nothing Then Set gcnOracle = gzlComLib.CurrentConn
    End If
    
    InitInterface = True
    
  
    Exit Function
err:
    If errHandle("zlSoftShowHisForms.InitInterface", "初始化接口出错", err.Description) = 1 Then Resume
End Function

Public Function InitSysParameter() As Boolean
'------------------------------------------------
'功能：初始化全局参数
'参数：无
'返回：True-成功；False-失败
'------------------------------------------------
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    
    InitSysParameter = False
    
    ''获取是否启用影像信息系统接口
    gblnUseInterface = Val(gzlComLib.zlDatabase.GetPara(255, glngSys)) = 1
    
    InitSysParameter = True
    
On Error GoTo Error
    'gblnXWRISInterfaceLog默认为false
    gblnXWRISInterfaceLog = False
    
    strSQL = "Select 内容 From zlRegInfo Where 项目 = '记录专业版RIS日志'"
    Set rsData = gzlComLib.zlDatabase.OpenSQLRecord(strSQL, "记录专业版RIS日志")
    
    If rsData.RecordCount > 0 Then
        gblnXWRISInterfaceLog = Nvl(rsData!内容, "0") = "1"
    End If
    
    Exit Function
Error:
    If errHandle("zlSoftShowHisForms.InitSysParameter", "判断是否记录日志到数据库时出现错误", strSQL) = 1 Then Resume
End Function

Public Function AnalyseComputer() As String
'获取计算机名
    Dim strComputer As String * 256
    Call GetComputerName(strComputer, 255)
    AnalyseComputer = strComputer
    AnalyseComputer = Trim(Replace(AnalyseComputer, Chr(0), ""))
End Function

Public Function Nvl(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'clsCommFun存在该函数
'功能：相当于Oracle的NVL，将Null值改成另外一个预设值
    Nvl = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

Public Function Decode(ParamArray arrPar() As Variant) As Variant
'功能：模拟Oracle的Decode函数
    Dim varValue As Variant, i As Integer
    
    i = 1
    varValue = arrPar(0)
    Do While i <= UBound(arrPar)
        If i = UBound(arrPar) Then
            Decode = arrPar(i): Exit Function
        ElseIf varValue = arrPar(i) Then
            Decode = arrPar(i + 1): Exit Function
        Else
            i = i + 2
        End If
    Loop
End Function

Public Function ProcessMessage(strmsg As String) As Long
    '处理接收到的消息
    '根据传入的参数判断处理哪个部件，消息格式“HIS主机字符串:用户名:密码:是否转换密码(0/1):病人ID:主页ID”
    Dim lng病人ID As Long
    Dim lng主页ID As Long
    Dim blnLis As Boolean
    Dim strErr As String
    
    On Error GoTo err
    ProcessMessage = 1
    
    If UBound(Split(strmsg, MSG_SPLIT)) = 5 Then
        gstrZLHIS主机字符串 = Split(strmsg, MSG_SPLIT)(0)
        gstr用户名 = Split(strmsg, MSG_SPLIT)(1)
        gstr密码 = Split(strmsg, MSG_SPLIT)(2)
        gbln是否转换密码 = Val(Split(strmsg, MSG_SPLIT)(3)) = 1
        lng病人ID = Val(Split(strmsg, MSG_SPLIT)(4))
        lng主页ID = Val(Split(strmsg, MSG_SPLIT)(5))
    ElseIf UBound(Split(strmsg, MSG_SPLIT)) = 6 Then
        blnLis = Val(Split(strmsg, MSG_SPLIT)(0)) = 25
        gstrZLHIS主机字符串 = Split(strmsg, MSG_SPLIT)(1)
        gstr用户名 = Split(strmsg, MSG_SPLIT)(2)
        gstr密码 = Split(strmsg, MSG_SPLIT)(3)
        gbln是否转换密码 = Val(Split(strmsg, MSG_SPLIT)(4)) = 1
        lng病人ID = Val(Split(strmsg, MSG_SPLIT)(5))
        lng主页ID = Val(Split(strmsg, MSG_SPLIT)(6))
    Else
        Exit Function
    End If
    
    '检验报告浏览
    If blnLis Then
        If mobjLisInsideComm Is Nothing Then
            Set mobjLisInsideComm = CreateObject("zl9LisInsideComm.clsLisInsideComm")
            If mobjLisInsideComm.InitComponentsLIS(glngSys, glngModule, gcnOracle, strErr) = False Then
                If strErr <> "" Then
                    If errHandle("zlSoftShowHisForms.ProcessMessage", strErr) = 1 Then Resume
                    Exit Function
                End If
            End If
            If mobjLisInsideComm Is Nothing Then
                 MsgBox "LIS接口初始化失败", vbInformation, "提示"
                 Exit Function
            End If
        End If
        Call mobjLisInsideComm.PatientSampleBrowse(frmShowHisForms, lng病人ID, "", 0, 0, IIf(lng主页ID = 0, 1, 2), lng主页ID)
    Else
    
        '电子病案查询
        If mclsArchive Is Nothing Then
            '第一次调用电子病案查阅，赋值
            Set mclsArchive = New clsArchive
        End If
        Call mclsArchive.zlOpenArchiveForm(lng病人ID, lng主页ID)
    End If

    ProcessMessage = 0
    Exit Function
err:
    
End Function

Public Function CloseAllForms() As Boolean

    On Error GoTo err
    
    '关闭电子病案查阅窗口
    If Not mclsArchive Is Nothing Then
        mclsArchive.zlCloseArchiveForm
    End If
'
    '关闭LIS报告浏览器
    If Not mobjLisInsideComm Is Nothing Then
        Set mobjLisInsideComm = Nothing
    End If
    
    '关闭消息循环主窗口
    If Not mfrmShowHisForms Is Nothing Then
        Unload mfrmShowHisForms
        Set mfrmShowHisForms = Nothing
    End If
    
    CloseAllForms = True
    
    Exit Function
err:
   
    Resume Next
End Function

Public Sub writeTestLog(ByVal strInfo As String)
'API申明时钟函数
'Private Declare Function GetTickCount Lib "kernel32" () As Long
'引用  microsoft script runtime  即(C:\Windows\System32\scrrun.dll)
    Dim objFile As FileSystemObject
    Dim objText As TextStream
    Dim strFile As String
    Dim strTmp As String
    
    If C_LOG = 0 Then Exit Sub
    
    On Error Resume Next
    
    Set objFile = New FileSystemObject
    
    strFile = App.Path & "\zlSoftShowArchiveView" & Format(Now, "YYYY_MM_DD") & ".Log"
    
    If Not Dir(strFile) <> "" Then objFile.CreateTextFile strFile
    strTmp = strInfo
    Set objText = objFile.OpenTextFile(strFile, ForAppending)
    objText.WriteLine strTmp
    objText.Close
End Sub
