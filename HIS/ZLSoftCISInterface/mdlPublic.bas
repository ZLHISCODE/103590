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


Public Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function EnumWindows& Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long)
Public Declare Function BringWindowToTop Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long

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
Public Const C_LOG = 0 '是否记录日志,0不生成日志，1要记录生成日志

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
Public glng病人ID As Long
Public glng主页ID As Long
Public glngFunID As Long
Public glng病区ID As Long
Public glng科室ID As Long
Public glng功能号 As Long '0-病案查阅,1-LIS调用,2-医嘱处理;3-执行端付费配置;4-执行端付费;5-单据打印;99-自定义报表


Public glngPid As Long

Public gstrHwndOLD As String
Public gstrHwndNew As String



Private mclsReport As Object


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
    MsgBox errTitle & errDesc, vbOKOnly, "接口zlSoftCISInterface出现错误"
    
    '清除错误
    err.Clear
    
End Function


Public Sub MkLocalDir(ByVal strDir As String)
'------------------------------------------------
'功能：创建本地目录
'参数： strDir－－本地目录
'返回：无
'------------------------------------------------
    Dim objFile As New Scripting.FileSystemObject
    Dim aNestDirs() As String, i As Integer
    Dim strPath As String
    On Error Resume Next
    
    '读取全部需要创建的目录信息
    ReDim Preserve aNestDirs(0)
    aNestDirs(0) = strDir
    
    strPath = objFile.GetParentFolderName(strDir)
    Do While Len(strPath) > 0
        ReDim Preserve aNestDirs(UBound(aNestDirs) + 1)
        aNestDirs(UBound(aNestDirs)) = strPath
        strPath = objFile.GetParentFolderName(strPath)
    Loop
    '创建全部目录
    For i = UBound(aNestDirs) To 0 Step -1
        MkDir aNestDirs(i)
    Next
End Sub

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
    If errHandle("zlSoftCISInterface.ConnectDB", "连接数据库函数出现错误", err.Description) = 1 Then Resume
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
    
    If errHandle("zlSoftCISInterface.OraDataOpen", "连接数据库出错", err.Description) = 1 Then Resume
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

    
    On Error Resume Next
    err.Clear
    If GetEMRLoginUser(strDBServer, strDBPassword) Then
        Set objEmr = CreateObject("zl9EmrInterface.ClsEmrInterface")
        If err.Number = 0 Then
            Call objEmr.CheckUpdate1(gstrDBUser, strDBPassword, True)
            If err.Number <> 0 Then
                err.Clear
                If objEmr.CheckUpdate(gstrDBUser, strDBPassword) = False Then
                    Exit Function
                End If
            End If
        Else
            err.Clear
        End If
    Else
        Set objEmr = CreateObject("zl9EmrInterface.ClsEmrInterface")
        If err.Number = 0 Then
            '从注册表读取数据库连接信息
            strDBServer = gstrZLHIS主机字符串
            strDBPassword = gstr密码
    
            Call objEmr.CheckUpdate1(gstrDBUser, strDBPassword, True)
            If err.Number <> 0 Then
                err.Clear
                If objEmr.CheckUpdate(gstrDBUser, strDBPassword) = False Then
                    Exit Function
                End If
            End If
        Else
            err.Clear
        End If
    End If

    
    Set UpdateEmrInterface = objEmr
    On Error GoTo 0
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
    
    '创建日志目录
    Call MkLocalDir(gstrLogPath + "\")
    Call MkLocalDir(gstrBackupPath + "\")
 
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
    If errHandle("zlSoftCISInterface.InitInterface", "初始化接口出错", err.Description) = 1 Then Resume
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
    Set rsData = gzlComLib.zlDatabase.OpenSQLRecord(strSQL, "zlSoftCISInterface")
    
    If rsData.RecordCount > 0 Then
        gblnXWRISInterfaceLog = Nvl(rsData!内容, "0") = "1"
    End If
    
    Exit Function
Error:
    If errHandle("zlSoftCISInterface.InitSysParameter", "判断是否记录日志到数据库时出现错误", strSQL) = 1 Then Resume
End Function

'======================================================================================================================
'方法           ByteToHexString         将16进制字符串转换为字节组
'返回值         Byte()                  16进制字符串转换的字节组
'入参列表:
'参数名         类型                    说明
'bstrInput      String                  16进制字符串
'lngRetBytLen   Long(Optional)          指定返回的字节组的长度,0-按原始长度返回，<>0返回指定的长度，不足补齐（补0），多了截取
'======================================================================================================================
Public Function HexStringToByte(ByVal strInput As String, Optional ByVal lngRetBytLen As Long) As Byte()
    Dim arrReturn() As Byte
    Dim i           As Long
    Dim lngLen      As Long
    
    lngLen = Len(strInput)
    If lngRetBytLen <> 0 Then
        lngLen = lngLen \ 2
        If lngLen > lngRetBytLen Then
            lngLen = lngRetBytLen
        End If
        ReDim arrReturn(lngRetBytLen - 1)
    Else
        lngLen = lngLen \ 2
        ReDim arrReturn(lngLen - 1)
    End If
    
    For i = 0 To lngLen - 1
        arrReturn(i) = Val("&H" & Mid(strInput, 2 * i + 1, 2))
    Next
    
    HexStringToByte = arrReturn()
End Function

Private Function TruncZeroInside(ByVal strInput As String) As String
'功能：去掉字符串中\0以后的字符,仅用作该工程,可以单独是用clsstring
    Dim lngPos As Long
    
    lngPos = InStr(strInput, Chr(0))
    If lngPos > 0 Then
        TruncZeroInside = Mid(strInput, 1, lngPos - 1)
    Else
        TruncZeroInside = strInput
    End If
End Function

'======================================================================================================================
'方法           sm_version              获取ZLSM4的版本号
'返回值         Long                    ZLSM4的版本号
'入参列表:
'======================================================================================================================
Public Function sm_version() As Long
    Dim lngVersion As Long
    On Error Resume Next
    lngVersion = get_sm_version
    If err.Number <> 0 Then
        err.Clear
        sm_version = 1
    Else
        sm_version = lngVersion
    End If
End Function


Private Function GetKey(ByVal strKey As String, ByVal intType As Integer) As Byte()
    Dim arrReturn() As Byte
    Dim i           As Long
    If strKey <> "" Then
        arrReturn = HexStringToByte(strKey, 16)
    Else
        ReDim arrReturn(15)
        If intType = 0 Then
            For i = 0 To 15
                arrReturn(i) = i * 15
            Next
        ElseIf intType = 1 Then
            Rnd (-1)
            Randomize (SM4_CRYPT_RANDOMIZE_IV)
            For i = 0 To 15
                arrReturn(i) = Int(Rnd() * 256)
            Next
        ElseIf intType = 2 Then
            Rnd (-1)
            Randomize (SM4_CRYPT_RANDOMIZE_KEY)
            For i = 0 To 15
                arrReturn(i) = Int(Rnd() * 256)
            Next
        End If
    End If
    GetKey = arrReturn
End Function

'======================================================================================================================
'方法           Sm4DecryptEcb           SM4解密
'返回值         String                  解密后的值
'入参列表:
'参数名         类型                    说明
'strInput       String                  要解密的字符串（该字符串是Sm4EncryptEcb生成的结果）
'strKey         String(Optional)        加密密钥也就是解密密钥（32位的16进制字符串，可以通过HexStringToByte返回）
'======================================================================================================================
Public Function Sm4DecryptEcb(ByVal strInput As String, Optional ByVal strKey As String) As String
    Dim arrKey()        As Byte
    Dim arrInput()      As Byte
    Dim arrOutPut()     As Byte
    Dim lngVersion      As Long

    If M_SM4_VERSION = 0 Then
        M_SM4_VERSION = sm_version
    End If
    If strInput Like "ZLSV*:*" Then
        lngVersion = Val(Mid(strInput, 5, InStr(strInput, ":") - 5))
        strInput = Mid(strInput, InStr(strInput, ":") + 1)
        '当前客户端的ZLSM4不支持该版本的加密字符串解密，仍旧解密，因为一般来说都能解密出相同的字符串
'        If lngVersion > M_SM4_VERSION Then
'            Exit Function
'        End If
    Else
        Exit Function
    End If
    
    arrKey = GetKey(strKey, 2)
    arrInput = HexStringToByte(strInput)
    ReDim arrOutPut(UBound(arrInput))
    
    Call sm4_crypt_ecb(CM_Decrypt, UBound(arrInput) + 1, arrKey(0), arrInput(0), arrOutPut(0))
    If lngVersion = 1 Then
        Sm4DecryptEcb = Trim(StrConv(arrOutPut(), vbUnicode))
    Else
        Sm4DecryptEcb = TruncZeroInside(StrConv(arrOutPut(), vbUnicode))
    End If
End Function

Private Function GetEMRLoginUser(strUser As String, strPwd As String) As Boolean
'功能：获取EMP初始化的用户与密码
'返回：是否获取成功（当存在只存在2500系统，则从配置文件获取，若不存在100与2500系统，则返回FALSE

    Dim strSQL      As String
    Dim rsTmp       As ADODB.Recordset
    Dim rsTest      As ADODB.Recordset
    Dim strConn     As String
    Dim objFSO      As New FileSystemObject
    Dim arrInfo As Variant
    Dim strCode     As String

    On Error GoTo errH
    strSQL = "Select Floor(a.编号 / 100) 编号 From zlSystems A Where Floor(a.编号 / 100) In (1, 25)"
    Set rsTmp = gzlComLib.zlDatabase.OpenSQLRecord(strSQL, "GetEMRLoginUser")
    If rsTmp.RecordCount <> 0 Then
        rsTmp.Filter = "编号=1"
        If rsTmp.RecordCount = 0 Then
            rsTmp.Filter = "编号=25"
            If rsTmp.RecordCount <> 0 Then
                strSQL = "Select 参数值 From zlOptions Where  参数名 =[1]"
                Set rsTest = gzlComLib.zlDatabase.OpenSQLRecord(strSQL, "获取LIS连接配置", "LIS系统连接配置")
                If rsTest.RecordCount > 0 Then
                    strConn = rsTest("参数值") & ""
                End If
                If strConn <> "" Then
                    strCode = Sm4DecryptEcb(strConn)
                    arrInfo = Split(strCode, "<SP 1>")
                    If UBound(arrInfo) >= 1 Then
                        If arrInfo(0) <> "" And arrInfo(1) <> "" Then
                            strUser = arrInfo(0)
                            strPwd = IIf(UCase(arrInfo(0)) = "SYS" Or UCase(arrInfo(0)) = "SYSTEM", "[DBPASSWORD]", "") & arrInfo(1)
                            GetEMRLoginUser = True
                        End If
                    End If
                End If
            End If
        Else
            strUser = gstrDBUser
            strPwd = IIf(gbln是否转换密码, "", "[DBPASSWORD]") & gstr密码
            GetEMRLoginUser = True
        End If
    End If
    Exit Function
errH:
    err.Clear
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

Public Function InStrW(wStr As String, wIn As String, wTimes As Long) As Long
    Dim sPos As Long
    Dim s As String
    
    s = Replace(wStr, wIn, "", 1, wTimes - 1, vbBinaryCompare)
    sPos = InStr(s, wIn)
    InStrW = sPos + wTimes - 1
End Function


Public Function ProcessMessage(strmsg As String) As Long
    '-----------------------------------------------
    '处理接收到的消息
    '消息格式：Oracle连接字符串:用户名:密码:是否密码转换(0或1):调用功能号(0-病案查阅,1-浏览检查报告,2-医嘱处理;3-执行端付费配置;4-执行端付费;5-单据打印;99-报表处理;999-功能初始化):...
    '              功能号不同，后续参数的格式与含义也不同
    '              功能=0,1,2时:功能号后参数：病人ID,主页ID
    '              功能=3，999时，功能后无参数
    '              功能=4时,功能后为:病人ID:医嘱信息:NOs
    '                      其中医嘱信息或NOs，任传一个即可,医嘱信息：执行科室|医嘱IDs(多个用逗号分隔);NOs: 多个用逗号分隔
'                  功能=5时,功能后为：打印类别(0=含打印及预览,1=直接到预览,2=直接打印,3-输出到Excel,4-输出到PDF,99-打印设置):(格式：报表编号,单据号(par)报表编号,单据号)功能后为:  报表编号,单据号(par)报表编号,单据号(par)报表编号,单据号
    '              功能=99时,功能后为：系统号:报表编号:打印类别(0=含打印及预览,1=直接到预览,2=直接打印,3-输出到Excel,4-输出到PDF,99-打印设置):报表参数(可为空 示例格式："病人id=1<par>PDF=C:\1.PDF<par>ExcelFile=C:\1.xls")
    
    
    '-----------------------------------------------
    Dim blnLis As Boolean
    Dim strErr As String
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    Dim str医嘱信息 As String
    Dim strNos     As String
    
    Dim lng系统号 As Long
    Dim str报表编号 As String
    Dim lng打印类别 As Long '打印类别(0=含打印及预览,1=直接到预览,2=直接打印,3-输出到Excel,4-输出到PDF,99-打印设置)
    Dim str报表参数 As String '
    Dim arrPar As Variant
    Dim arrPar固定(50) As String
    Dim arrNoPar As Variant



    Dim i As Long
                     
    
    Dim varData As Variant
    
    
    On Error GoTo err
    ProcessMessage = 1
    
    varData = Split(strmsg & String(7, MSG_SPLIT), MSG_SPLIT)
    gstrZLHIS主机字符串 = varData(0)
    gstr用户名 = varData(1)
    gstr密码 = varData(2)
    gbln是否转换密码 = Val(varData(3)) = 1
    glng功能号 = Val(varData(4)) '0-病案查阅,1-LIS调用,2-医嘱处理;3-执行端付费配置;4-执行端付费;5-单据打印
    If glng功能号 <> 3 And glng功能号 <> 99 And glng功能号 <> 5 Then
        glng病人ID = Val(varData(5))
        If glng功能号 = 4 Then
            str医嘱信息 = varData(6)
            strNos = varData(7)
        Else
            glng主页ID = Val(varData(6))
        End If
        glngFunID = IIf(glng功能号 = 2, 3001, 0)
    End If
    
    If glng功能号 = 99 Then
        lng系统号 = Val(varData(5))
        str报表编号 = varData(6)
        lng打印类别 = Val(varData(7))
        
        str报表参数 = Mid(strmsg, InStrW(strmsg, MSG_SPLIT, 8) + 1) '不用split，预防报表参数里有冒号
        If str报表参数 <> "" Then
            str报表参数 = Mid(str报表参数, 2)
            str报表参数 = Mid(str报表参数, 1, Len(str报表参数) - 1)
        End If
    End If
    If glng功能号 = 5 Then
        lng打印类别 = Val(varData(5))
        str报表参数 = varData(6)
    End If
    
    
    
    If InStr("0,1,2", glng功能号) > 0 And (glng病人ID = 0) Then
        Exit Function
    ElseIf glng功能号 = 4 And (str医嘱信息 = "" And strNos = "") Then
        Exit Function
    ElseIf glng功能号 = 5 And str报表参数 = "" Then
        Exit Function
    ElseIf glng功能号 = 99 And (str报表编号 = "") Then
        Exit Function
    End If
    
    blnLis = glng功能号 = 1
    
    If glng功能号 = 2 Then
        strSQL = "Select a.当前病区ID,a.出院科室ID From 病案主页 a Where a.病人id=[1] and a.主页id=[2]"
        Set rsData = gzlComLib.zlDatabase.OpenSQLRecord(strSQL, "zlSoftCISInterface", glng病人ID, glng主页ID)
        If Not rsData.EOF Then
            glng病区ID = Val(rsData!当前病区ID & "")
            glng科室ID = Val(rsData!出院科室ID & "")
        End If
    End If
    
    If Not mfrmShowHisForms Is Nothing Then
        Call GetWindowThreadProcessId(mfrmShowHisForms.hWnd, glngPid)
    End If
    
    gstrHwndOLD = "": EnumChildWindows GetDesktopWindow, AddressOf EnumChildProcOld, ByVal 0
    Select Case glng功能号
        Case 0, 1 '病案查阅
            '检验报告浏览
            If Not mclsReport Is Nothing Then
                mclsReport.CloseWindows
            End If
            If blnLis Then
                If mobjLisInsideComm Is Nothing Then
                    Set mobjLisInsideComm = CreateObject("zl9LisInsideComm.clsLisInsideComm")
                    If mobjLisInsideComm.InitComponentsLIS(glngSys, glngModule, gcnOracle, strErr) = False Then
                        If strErr <> "" Then
                            If errHandle("zlSoftCISInterface.ProcessMessage", strErr) = 1 Then Resume
                            Exit Function
                        End If
                    End If
                    If mobjLisInsideComm Is Nothing Then
                         MsgBox "LIS接口初始化失败", vbInformation, "提示"
                         Exit Function
                    End If
                End If
                mfrmShowHisForms.TimerShow.Enabled = True
                Call mobjLisInsideComm.PatientSampleBrowse(frmShowHisForms, glng病人ID, "", 0, 0, IIf(glng主页ID = 0, 1, 2), glng主页ID)
            Else
                '电子病案查询
                If mclsArchive Is Nothing Then
                    '第一次调用电子病案查阅，赋值
                    Set mclsArchive = New clsArchive
                End If
                Call mclsArchive.zlOpenArchiveForm(glng病人ID, glng主页ID)
            End If
        Case 2
            If mclsOrder Is Nothing Then
                Set mclsOrder = New clsOrder
            End If
            
            mclsOrder.zlCloseOrderForm
            mclsOrder.zlOpenOrderForm
        Case 3 '执行端付费配置
            If mclsFee Is Nothing Then
                Set mclsFee = New clsFee
            End If
            Call mclsFee.zlDeviceSetup
        Case 4 '执行端付费
            If mclsFee Is Nothing Then
                Set mclsFee = New clsFee
            End If
            If mclsFee.zlSquareAffirm(glng病人ID, str医嘱信息, strNos) = False Then Exit Function
        Case 99 '自定义报表
            If mclsReport Is Nothing Then
                Set mclsReport = CreateObject("zl9Report.clsReport")
            End If
            
            If (Not mclsReport Is Nothing) And (Not gcnOracle Is Nothing) Then
                mclsReport.CloseWindows
                If lng打印类别 = 99 Then
                    Call mclsReport.ReportPrintSet(gcnOracle, lng系统号, str报表编号, mfrmShowHisForms)
                Else
                    If str报表参数 <> "" Then
                        arrPar = Split(str报表参数, "<par>")
                        For i = LBound(arrPar) To UBound(arrPar)
                            arrPar固定(i) = arrPar(i)
                        Next
                    End If
                    
                    mfrmShowHisForms.TimerShow.Enabled = True
                    
                    
                    Call mclsReport.ReportOpen(gcnOracle, lng系统号, str报表编号, mfrmShowHisForms, arrPar固定(0), arrPar固定(1), arrPar固定(2), arrPar固定(3), arrPar固定(4), arrPar固定(5), _
                             arrPar固定(6), arrPar固定(7), arrPar固定(8), arrPar固定(9), arrPar固定(10), arrPar固定(11), _
                             arrPar固定(12), arrPar固定(13), arrPar固定(14), arrPar固定(15), arrPar固定(16), arrPar固定(17), _
                             arrPar固定(18), arrPar固定(19), lng打印类别)
                End If
            End If
        Case 5 '单据打印
            If mclsReport Is Nothing Then
                Set mclsReport = CreateObject("zl9Report.clsReport")
            End If
            
            If (Not mclsReport Is Nothing) And (Not gcnOracle Is Nothing) Then
                mclsReport.CloseWindows
            
            
            
                mfrmShowHisForms.TimerShow.Enabled = True
                If lng打印类别 = 99 Then
                    Call mclsReport.ReportPrintSet(gcnOracle, 100, str报表参数, mfrmShowHisForms)
                Else
                    If str报表参数 <> "" Then
                        arrPar = Split(str报表参数, "(par)")
                        For i = LBound(arrPar) To UBound(arrPar)
                            arrNoPar = Split(arrPar(i), ",")
                            
                             Call mclsReport.ReportOpen(gcnOracle, 100, arrNoPar(0), mfrmShowHisForms, "NO=" & arrNoPar(1), "性质=1", "医嘱ID=0", "PrintEmpty=0", lng打印类别)
                        Next
                    End If
                End If
            End If
            
            If lng打印类别 = 2 Then Call CloseAllForms
        Case 999 '初始化
            
            On Error Resume Next
            If mobjLisInsideComm Is Nothing Then
                Set mobjLisInsideComm = CreateObject("zl9LisInsideComm.clsLisInsideComm")
                If Not mobjLisInsideComm Is Nothing Then
                    Call mobjLisInsideComm.InitComponentsLIS(glngSys, glngModule, gcnOracle, strErr)
                End If
            End If
            
            Set mclsReport = CreateObject("zl9Report.clsReport")
            
            err.Clear
            '电子病案查询
            If mclsArchive Is Nothing Then
                '第一次调用电子病案查阅，赋值
                Set mclsArchive = New clsArchive
                
                If Not mclsArchive Is Nothing Then
                   Call mclsArchive.zlOpenArchiveForm(0, 0, True)
                End If
            End If
            
            If mclsOrder Is Nothing Then
                Set mclsOrder = New clsOrder
            End If

            If mclsFee Is Nothing Then
                Set mclsFee = New clsFee
            End If
            
            err.Clear
            
    End Select
    ProcessMessage = 0
    Exit Function
err:
    MsgBox err.Description
End Function

Public Function CloseAllForms() As Boolean
    On Error GoTo err
    
    If Not mfrmShowHisForms Is Nothing Then
        Call GetWindowThreadProcessId(mfrmShowHisForms.hWnd, glngPid)
        KillPID glngPid
    End If
    
    
    '关闭报表对象
    If Not mclsReport Is Nothing Then
        Set mclsReport = Nothing
    End If
    
    
    '关闭收费对象
    If Not mclsFee Is Nothing Then
        Set mclsOrder = Nothing
    End If
    
    '关闭医嘱处理窗口
    If Not mclsOrder Is Nothing Then
        mclsOrder.zlCloseOrderForm
        Set mclsOrder = Nothing
    End If
    
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
    
    strFile = App.Path & "\zlSoftShowArchiveView.Log"
    
    If Not Dir(strFile) <> "" Then objFile.CreateTextFile strFile
    
'    If FileLen(strFile) > 52428800 Then     '判断文件是否大于50M
'        Name strFile As App.Path & "\CISJOBTest" & Format(Now, "yyyymmddhhmm") & ".bak"  '修改文件名
'        '当文件被重命名之后，需要重新检查并创建文件
'        Call writeTestLog(strInfo)
'        Exit Sub
'    End If

'    strTmp = "'" & strInfo & "' from dual Union All Select"
    strTmp = strInfo
    Set objText = objFile.OpenTextFile(strFile, ForAppending)
    objText.WriteLine strTmp
    objText.Close
'4072 个 Union All Select 左右
'insert into 测试数据 (ID,日期,文本)
'select 测试数据_ID.Nextval,sysdate,a.* from (
'select 文本 from  测试数据 Where 1 = 0 Union All Select

'是文件目录' from dual Union All Select
'是文件目录' from dual Union All Select
'是文件目录' from dual Union All Select

'文本 from  测试数据 Where 1 = 0) a;


'select id,文本,substr(文本,1,instr(文本,'|')-1) as 方法,
'substr(replace(文本,substr(文本,1,instr(文本,'|')) ,''),1,instr(replace(文本,substr(文本,1,instr(文本,'|')) ,''),'|')-1) as 模块,
'replace(文本,substr(文本,1,instr(文本,'|',-1)) ,'') as 工程
'from 测试数据 where id>=4848;
'解析日志
'    Set objText = objFile.OpenTextFile(strFile, ForReading)
'    Do While Not objText.AtEndOfStream
'        strTmp = objText.ReadLine
'    Loop
'    objText.Close
End Sub

Public Function EnumChildProc(ByVal hWnd As Long, ByVal lParam As Long) As Long
 Dim cklm As String * 50 '窗口类名
 Dim lngPid As Long
    On Error Resume Next
    GetClassName hWnd, cklm, 50
    lngPid = 0
    If InStr(LCase(Blank(cklm)), "form") > 0 And InStr(gstrHwndOLD, "," & hWnd & ",") = 0 Then
        Call GetWindowThreadProcessId(hWnd, lngPid)

        If lngPid = glngPid And lngPid <> 0 Then
            
            SetWindowPos hWnd, -1, 0, 0, 0, 0, &H1 Or &H2
            SetWindowPos hWnd, -2, 0, 0, 0, 0, &H1 Or &H2
            BringWindowToTop hWnd
            SetForegroundWindow hWnd
        End If
    End If
    EnumChildProc = 1
End Function


Public Function EnumChildProcOld(ByVal hWnd As Long, ByVal lParam As Long) As Long
 Dim cklm As String * 50 '窗口类名
 Dim lngPid As Long
    On Error Resume Next
    GetClassName hWnd, cklm, 50
    lngPid = 0
    If InStr(LCase(Blank(cklm)), "form") > 0 Then
        Call GetWindowThreadProcessId(hWnd, lngPid)

        If lngPid = glngPid And lngPid <> 0 Then
             gstrHwndOLD = "," & gstrHwndOLD & "," & hWnd & ","
        End If
    End If
    EnumChildProcOld = 1
End Function

Public Function Blank(ByVal szString As String) As String
    Dim l As Integer
    l = InStr(szString, Chr(0))
    If l > 0 Then
        Blank = Left(szString, l - 1)
    Else
        Blank = szString
    End If
End Function


Public Function KillPID(ByVal lngPid As Long) As Boolean
    
    '杀死进程
    On Error Resume Next
    Shell ("taskkill /pid " & lngPid & " -t -f")
End Function

