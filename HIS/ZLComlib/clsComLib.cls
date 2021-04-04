VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsComLib"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'---------------------------------------------------------------------------------
'模块名  ：clsComLib
'模块说明：本模块作为全局类，可直接被其它部件声明使用。
'模块内容：包含其它功能类的声明、基本函数的定义
'---------------------------------------------------------------------------------
Public zlCommFun As clsCommFun       '实现算法和特殊功能
Public zlControl As clsControl       '操作控件
Public zlDatabase As clsDatabase     '操作数据库

Public zlXML As clsXML        '

Public gstrNodeNo As String          '当前站点编号(登录时选择的分院)；如果未设置启用站点，则为"-"
Public gstrNodeName As String        '当前站点名称(登录时选择的分院)；如果未设置启用站点，则为"-"
Public gblnAutoConnect As Boolean   '是否断网后自动连接数据库
Public mlngConnectCount As Long     '相同SQL次数

Private mblnExeSQLTest As Boolean
Private mstrSQLPlan As String    '进行执行计划跟踪的SQL

Private msngTime As Single
Private mstrBigTable As String   '大表
Private mstrMiddleTable As String '中表
Private mstrMiddlePar As String

'错误日志处理相关变量
Private mlngErrNum As Long, mstrErrInfo As String, mbytErrType As Byte

'ADO事务错误状态判断
Private WithEvents mobjOracle As ADODB.Connection
Attribute mobjOracle.VB_VarHelpID = -1
Private mblnADORoll As Boolean
Private mlngADOErrNum As Long
Private mstrADOErrDesc As String

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private mrsIndex As Recordset
Private mrsCons_FK As Recordset
Private mstrVersion As String
'分类函数
Private mvarclsComboBox As clsComboBox
Private mvarclsGrid As clsGrid
Private mvarclsRecordset As clsRecordset
Private mvarclsSystem As clsSystem
Private mvarclsFormShow As clsFormShow
Private mvarclsString As clsString
Private mvarclsOS As clsOS

Public Property Get Version() As String
    Version = mstrVersion
End Property

Public Property Get OS() As clsOS
    If mvarclsOS Is Nothing Then
        Set mvarclsOS = New clsOS
    End If
    Set OS = mvarclsOS
End Property

Public Property Get zlStr() As clsString
    If mvarclsString Is Nothing Then
        Set mvarclsString = New clsString
    End If
    Set zlStr = mvarclsString
End Property

Public Property Get Grid() As clsGrid
    If mvarclsGrid Is Nothing Then
        Set mvarclsGrid = New clsGrid
    End If
    Set Grid = mvarclsGrid
End Property

Public Property Get Cbo() As clsComboBox
    If mvarclsComboBox Is Nothing Then
        Set mvarclsComboBox = New clsComboBox
    End If
    Set Cbo = mvarclsComboBox
End Property


Public Property Get Rec() As clsRecordset
    If mvarclsRecordset Is Nothing Then
        Set mvarclsRecordset = New clsRecordset
    End If
    Set Rec = mvarclsRecordset
End Property

Public Property Get Sys() As clsSystem
    If mvarclsSystem Is Nothing Then
        Set mvarclsSystem = New clsSystem
    End If
    Set Sys = mvarclsSystem
End Property

Public Property Get FS() As clsFormShow
    If mvarclsFormShow Is Nothing Then
        Set mvarclsFormShow = New clsFormShow
    End If
    Set FS = mvarclsFormShow
End Property

Private Sub Class_Initialize()
    msngTime = 0
    mstrVersion = App.Major & "." & App.Minor & "." & App.Revision
    If gcllComlibs Is Nothing Then
        Set gcllComlibs = New Collection
    End If
    gcllComlibs.Add Me, "K_" & ObjPtr(Me)
    Set gobjComLib = Me
    If zlCommFun Is Nothing Then Set zlCommFun = New clsCommFun
    If zlControl Is Nothing Then Set zlControl = New clsControl
    If zlDatabase Is Nothing Then Set zlDatabase = New clsDatabase
    If zlXML Is Nothing Then Set zlXML = New clsXML

    '当创建 clsComLib 类时，创建 mclsOS 对象
    Set mvarclsOS = New clsOS
    Set mvarclsString = New clsString
    Set mvarclsGrid = New clsGrid
    Set mvarclsComboBox = New clsComboBox
    Set mvarclsRecordset = New clsRecordset
    Set mvarclsSystem = New clsSystem
    Set mvarclsFormShow = New clsFormShow
    
    gstrNodeNo = g_NodeNo
    gstrNodeName = g_NodeName
    gblnAutoConnect = g_AutoConnect
    Set mobjOracle = gcnOracle
End Sub

Private Sub Class_Terminate()
    If Not gobjLogFile Is Nothing Then
        If Not gobjLogText Is Nothing Then
            gobjLogText.Close
            Set gobjLogText = Nothing
        End If
        Set gobjLogFile = Nothing
    End If
    If Not gobjPlanExFile Is Nothing Then
        If Not gobjPlanExText Is Nothing Then
            gobjPlanExText.Close
            Set gobjPlanExText = Nothing
        End If
        Set gobjPlanExFile = Nothing
    End If

    '释放小类
    Set mvarclsOS = Nothing
    Set mvarclsString = Nothing
    Set mvarclsGrid = Nothing
    Set mvarclsComboBox = Nothing
    Set mvarclsRecordset = Nothing
    Set mvarclsSystem = Nothing
    Set mvarclsFormShow = Nothing
    If Not gcllComlibs Is Nothing Then
        gcllComlibs.Remove "K_" & ObjPtr(Me)
        If gcllComlibs.count = 0 Then
            Set gcllComlibs = Nothing
            Call ClearLogVars  '清除日志变量及对象
        End If
    End If
    '释放公共对象(多个部件都引用zlComLib时，每个部件卸载时都会调用本过程，为了避免别的部件中使用下列对象的语句出错，所以不释放)
    'Set gobjComLib = Nothing
    'Set zlCommFun = Nothing
    'Set zlControl = Nothing
    'Set zlDatabase = Nothing
    'Set gobjRegister = Nothing
End Sub

Private Sub mobjOracle_BeginTransComplete(ByVal TransactionLevel As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
    mblnADORoll = False: mlngADOErrNum = 0: mstrADOErrDesc = ""
End Sub

Private Sub mobjOracle_CommitTransComplete(ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
    mblnADORoll = False: mlngADOErrNum = 0: mstrADOErrDesc = ""
End Sub

Private Sub mobjOracle_RollbackTransComplete(ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
    mblnADORoll = True
    If pConnection.Errors.count > 0 Then
        mlngADOErrNum = pConnection.Errors(0).NativeError
        mstrADOErrDesc = pConnection.Errors(0).Description
    Else
        mlngADOErrNum = 0: mstrADOErrDesc = ""
    End If
End Sub

Public Function CloseWindows() As Boolean
    '--------------------------------------
    '功能:关闭所有子窗口
    '--------------------------------------
    Dim frmThis As Form
    Set grsParas = Nothing
    Set grsUserParas = Nothing
    Set grsUserInfo = Nothing
    Set gcolMoveDate = Nothing
    
    For Each frmThis In Forms
        Unload frmThis
    Next
    CloseWindows = (Forms.count = 0)
    '注销时销毁对象，释放缓存
    Set gobjRegister = Nothing
    If Not (gclsPDF Is Nothing) Then
        Set gclsPDF = Nothing
    End If
End Function
    
'------------------------------------------------
'第一部分：与本系统数据结构无关的函数
'------------------------------------------------

Public Sub InitCommon(cnMain As ADODB.Connection)
'功能： 部件的初始化，将主程序的必要内容传递下来
'参数： cnMain:主程序的数据库连接
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strTmp As String, strFile As String
    Dim arrTmp As Variant
    
    '屏蔽错误输入，放置错误影响整个进程，提高程序健壮性
    If cnMain Is Nothing Then Exit Sub
    '当前模块使用同一个连接多次调用，则不再调用实际代码
    If gobjComLib Is Me And gcnOracle Is cnMain Then Exit Sub
    
    '注册部件初始化
    '启动程序在登录之前已创建，以及医保部件等可能再次调用，避免重复创建部件
    If gobjRegister Is Nothing Then
        '这些都是固定的，只有初次登录才需要读取
        gstrAviPath = GetSetting("ZLSOFT", "注册信息", "gstrAviPath", "")
        gstrSysName = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")
        gstrHelpPath = gstrAviPath & "\help"
        gstrComputerName = OS.ComputerName  '启动程序中写入到zlclients时，是以本地代码获取的机器名，不是通过SYS_CONTEXT('USERENV','TERMINAL')获取的
        
        On Error Resume Next
        Set gobjRegister = GetObject("", "zlRegister.clsRegister")
        Err.Clear
        
        If gobjRegister Is Nothing Then
            '用于支持未通过导航台（启动程序prjMain）调用本部件的情况。
            '这种情况下，由于没有调用GetConnection,无法通过该对象获取服务器名、用户名、切换连接
            Set gobjRegister = CreateObject("zlRegister.clsRegister")
            Err.Clear
            If gobjRegister Is Nothing Then
                MsgBox "创建zlRegister部件对象失败,请检查文件是否存在并且正确注册。", vbExclamation, gstrSysName
                cnMain.Close
                Exit Sub
            End If
        End If
        '检查部件的MD5值(调试模式App.Path是当前源码工程的位置，所以不检查)，只进行一次合法性检查
        If Not OS.IsDesinMode Then
            strFile = App.Path & "\zlRegister.Dll"
            strTmp = Md5_File_Calc(strFile)
            If strTmp <> "F1335F5042068CF291B8141418775FD8" Then
                MsgBox "验证注册授权部件失败,请检查文件" & strFile & "的版本是否与ZL9Comlib.dll的版本匹配。", vbExclamation, gstrSysName
                cnMain.Close
                Exit Sub
            End If
        End If
        '保证FOXIT7配置生成，防止各个业务部件访问到Printer的属性时导致打印配置生成
'        If gclsPDF Is Nothing Then
'            Call zlCommFun.PDFInitialize
'        End If
    End If
    
    '连接切换，一些变量需要重新初始化
    If Not gcnOracle Is cnMain Then
        '眉山医院断开问题，必须重新设置，尽管还是这个值，否则会出现462-远程服务器不存在或无效
        cnMain.CursorLocation = adUseClient '
        Set gcnOracleOLEDB = Nothing '标识连接改变
        Set gobjComLib = Me
        '公共变量及对象初始
        Call gobjRegister.zlRegInit(cnMain)
        gstrDBUser = gobjRegister.GetUserName
        Set grsUserInfo = Nothing
        Set gcnOracle = cnMain
        Set mobjOracle = cnMain
        '会话认证
        If RegCheck = False Then
            cnMain.Close
            Exit Sub
        End If
        
        
        'SQL跟踪参数
        gblnSQLLog = False: gblnSQLTest = False: gblnSQLPlan = False
        strTmp = gobjRegister.GetServerName
        If strTmp Like "SQLLOG*" Then gblnSQLLog = True
        If strTmp Like "SQLTEST*" Then gblnSQLTest = True
        If strTmp Like "SQLPLAN*" Then gblnSQLPlan = True
        If gblnSQLPlan = False Then
            '技术中心性能测试模式
            If GetUnitInfo("使用模式") = "性能测试" Then
                gblnSQLPlan = True
            End If
        End If
        
        
        '取SessionID和当前计算机名
        strSQL = "Select UserENV('SessionID') as SID From dual"
        
        If gblnSQLPlan Then mblnExeSQLTest = True   '避免进入SQLTest，因为其中CheckSQLPlan调用GetSQLPlan时需要用到glngSessionID
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "InitCommon")
        If gblnSQLPlan Then mblnExeSQLTest = False
        
        glngSessionID = Val(rsTmp!SID & "")
        If gblnSQLTest Then
            '重庆九院IIS服务器在该窗体加载时失败，报无效的图片。具体是读取该窗体资源，加载窗体图标时不识别图标导致。
            '该问题和用户注册项丢失有关，尚未查出问题。因此增加该逻辑，在需要调试的时候才加载该窗体。
            Load frmSQLTest
        End If
            
        '系统参数初始
        strSQL = "Select 参数号,参数值 From ZLOPTIONS Where 参数号 in(1,3)"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "InitCommon")
        If Not rsTmp.EOF Then
            rsTmp.Filter = "参数号=1"
            If Not rsTmp.EOF Then gblnRunLog = (IIf(IsNull(rsTmp!参数值), 0, rsTmp!参数值) = 1)
            rsTmp.Filter = "参数号=3"
            If Not rsTmp.EOF Then gblnErrLog = (IIf(IsNull(rsTmp!参数值), 0, rsTmp!参数值) = 1)
        End If
        
        '从工作站（主键,登录时写入）获取站点编号(如果是通过导航台进入，则启动程序中的“是否允许使用本工作站”会再次设置此变量)
        If gstrNodeNo = "" Then
            strSQL = "Select a.站点, b.名称 From Zlclients a, Zlnodelist b Where a.站点 = b.编号 And a.工作站 = [1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "InitCommon", gstrComputerName)
            If Not rsTmp.EOF Then
                gstrNodeNo = IIf(IsNull(rsTmp!站点), "-", rsTmp!站点)
                gstrNodeName = IIf(IsNull(rsTmp!名称), "-", rsTmp!名称)
            Else
                gstrNodeNo = "-"
                gstrNodeName = "-"
            End If
        End If
        g_NodeNo = gstrNodeNo
        g_NodeName = gstrNodeName
        
        g_AutoConnect = IIf(Val(zlDatabase.GetPara("网络断网自动重连")) = 1, True, 0)
        Call SaveSetting("ZLSOFT", "公共全局\网络断网自动重连", "AutoConnect", IIf(g_AutoConnect = True, 1, 0))
        Call LogInit     '日志初始化
        '切换所有公共部件实例的事务对象
        Call ChangeAllIntanceConn(cnMain)
    End If
End Sub

Public Sub LogInit(Optional ByVal blnForceInit As Boolean)
'------------------------------------------------
'功能： 日志初始化
'参数： blnForceInit：强制初始化，仅在系统选择设置中调用
'------------------------------------------------
    Dim strPath As String, strDate As String, strFormat As String, strFileName As String
    Dim strSQL As String, strExePath As String, strLogPath As String, lngClearLogDays As Long
    Dim rsLoglist As ADODB.Recordset

    On Error GoTo errh:
    
    '若是在系统设置中调用，则清空变量，并进行初始化
    If blnForceInit Then
        Call ClearLogVars
    '若不是在系统设置时调用，但是第一次调用，则允许进行初始化操作
    '若不是第一次调用，则不允许再进行初始化
    ElseIf Not gcolWriteLog Is Nothing Then
        Exit Sub
    End If
    
    '获取路径
    If OS.IsDesinMode Then
        strExePath = "C:\APPSOFT\"
    Else
        strExePath = Mid(App.Path, 1, InStr(5, App.Path, "\"))
    End If
    strLogPath = strExePath & "Log\日志跟踪"

    If gobjLogFile Is Nothing Then Set gobjLogFile = New FileSystemObject
    If gcolWriteLog Is Nothing Then Set gcolWriteLog = New Collection
    If gcolLastLogInfoHeader Is Nothing Then Set gcolLastLogInfoHeader = New Collection
    
    '判断日志路径是否存在，若不存在，则创建
    If gobjLogFile.FolderExists(strExePath & "\Log") = False Then
        gobjLogFile.CreateFolder strExePath & "\Log"
    End If
    If gobjLogFile.FolderExists(strLogPath) = False Then
        gobjLogFile.CreateFolder strLogPath
    End If
    
    '清除历史日志以及将日志路径存储到集合中
    strSQL = "Select Nvl(系统,0) 系统, 名称 From Zllogconfig"
    Set rsLoglist = zlDatabase.OpenSQLRecord(strSQL, "获取日志类型列表")
    With rsLoglist
        Do While Not .EOF
            If GetSetting("ZLSOFT", "公共全局\日志\" & rsLoglist!名称, "启用", 0) = 1 Then
                '清除日志
                lngClearLogDays = GetSetting("ZLSOFT", "公共全局\日志\" & rsLoglist!名称, "保存天数", "3")
                strDate = GetSetting("ZLSOFT", "公共全局\日志\" & rsLoglist!名称, "上次清理时间", "1901-01-01")
                '如果日志路径不存在，则创建该路径
                If gobjLogFile.FolderExists(strLogPath & "\" & rsLoglist!系统 & "_" & rsLoglist!名称) = False Then
                    gobjLogFile.CreateFolder strLogPath & "\" & rsLoglist!系统 & "_" & rsLoglist!名称
                Else
                    If Format(strDate, "YYYY-MM-DD") <> Format(Now, "YYYY-MM-DD") Then  '如果今天已清理过日志，则不再清理
                        strPath = Dir(strLogPath & "\" & rsLoglist!系统 & "_" & rsLoglist!名称 & "\*.log")   '查找第一个文件
                        On Error Resume Next
                        Do While Len(strPath) '循环到没有文件为止
                            strDate = Mid(strPath, Len(rsLoglist!名称) + 2, 14)
                            strFormat = Mid(strDate, 1, 4) & "-" & Mid(strDate, 5, 2) & "-" & Mid(strDate, 7, 2) & " " & Mid(strDate, 9, 2) & ":" & Mid(strDate, 11, 2) & ":" & Mid(strDate, 13, 2)
                            If DateDiff("d", Format(strFormat, "YYYY-MM-DD HH:mm:ss"), Format(Now, "YYYY-MM-DD HH:mm:ss")) > lngClearLogDays Then
                                gobjLogFile.DeleteFile strLogPath & "\" & rsLoglist!系统 & "_" & rsLoglist!名称 & "\" & strPath
                            End If
                            strPath = Dir '查找下一个文件
                        Loop
                        SaveSetting "ZLSOFT", "公共全局\日志\" & rsLoglist!名称, "上次清理时间", Now
                        If Err.Number <> 0 Then Err.Clear
                    End If
                End If
                
                '将日志文件路径存储到集合中，此时并没有生成文件，在写日志的时候才会创建日志文件
                strFileName = strLogPath & "\" & rsLoglist!系统 & "_" & rsLoglist!名称 & "\"
                gcolWriteLog.Add strFileName, "Key_" & rsLoglist!名称
                gcolLastLogInfoHeader.Add "", "Key_" & rsLoglist!名称
            Else
                '清除日志，因为该日志类型未启用，所以删除全部此类型日志
                If gobjLogFile.FolderExists(strLogPath & "\" & rsLoglist!系统 & "_" & rsLoglist!名称) = True Then
                    gobjLogFile.DeleteFolder strLogPath & "\" & rsLoglist!系统 & "_" & rsLoglist!名称
                End If
                
                '因为日志未启用，故将nothing存入集合
                gcolWriteLog.Add Nothing, "Key_" & rsLoglist!名称
            End If
            .MoveNext
        Loop
    End With
    Exit Sub
errh:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub LogWrite(ByVal strLogName As String, ByVal strModule As String, ByVal strFunction As String, ByVal strLog As String)
'------------------------------------------------
'功能：写入日志
'参数：
'      strLogName ：日志名称
'      strModule  ：模块名
'      strFunction：功能名
'      strLog     ：日志内容
'备注：在系统选项中开启日志后，在写入日志时创建日志文件，每个进程每个日志创建只创建一次
'------------------------------------------------
    Dim strDataType As String
    Dim strFileName As String
    Dim strLogPath As String, strExePath As String
    
    On Error GoTo errh

    If gcolWriteLog Is Nothing Then Exit Sub
    '若当前使用的日志对象未缓存了,重新取对象，并缓存起来
    If gstrLastLogName <> strLogName Then
        '先判断该日志名称是否存在
        '若多次调用同一个不存在的日志，会不停地弹出错误，以下几行代码是为了使错误只弹出一次
        On Error Resume Next
        strDataType = TypeName(gcolWriteLog("Key_" & strLogName))
        If Err.Number <> 0 Then
            '为了使工具类那边也能调用该方法，故先判断当前是否连接数据
            '若已连接，则说明日志名称确实填写错误，将错误弹出。若未连接，则说明为工具类那边调用，继续使用该日志
            If gcnOracle Is Nothing Then
                '生成日志存储路径
                strExePath = Mid(App.Path, 1, InStr(5, App.Path, "\"))
                strLogPath = strExePath & "Log\日志跟踪"
            
                If gobjLogFile Is Nothing Then Set gobjLogFile = New FileSystemObject
                If gcolWriteLog Is Nothing Then Set gcolWriteLog = New Collection
                If gcolLastLogInfoHeader Is Nothing Then Set gcolLastLogInfoHeader = New Collection
                
                '判断日志路径是否存在，若不存在，则创建
                If gobjLogFile.FolderExists(strExePath & "\Log") = False Then
                    gobjLogFile.CreateFolder strExePath & "\Log"
                End If
                If gobjLogFile.FolderExists(strLogPath) = False Then
                    gobjLogFile.CreateFolder strLogPath
                End If
                If gobjLogFile.FolderExists(strLogPath & "\" & "9999_" & strLogName) = False Then
                    gobjLogFile.CreateFolder strLogPath & "\" & "9999_" & strLogName
                End If
                
                strFileName = strLogPath & "\" & "9999_" & strLogName & "\"
                gcolWriteLog.Add strFileName, "Key_" & strLogName
                gcolLastLogInfoHeader.Add "", "Key_" & strLogName
            Else
                MsgBox "日志【" & strLogName & "】未在ZLLogConfig中注册，请联系开发人员进行日志注册！", vbInformation, gstrSysName
                Err.Clear
                gcolWriteLog.Add Nothing, "Key_" & strLogName
                Exit Sub
            End If
        End If
        On Error GoTo errh
    
        '当日志未启用时
        If TypeName(gcolWriteLog("Key_" & strLogName)) = "String" Then
            strFileName = gcolWriteLog("Key_" & strLogName) & strLogName & "_" & Format(Now, "YYYYMMDDHHmmss") & ".log"
            '根据文件路径，创建TextStream对象
            gcolWriteLog.Remove "Key_" & strLogName
            gcolWriteLog.Add gobjLogFile.CreateTextFile(strFileName), "Key_" & strLogName
        Else
            If gcolWriteLog("Key_" & strLogName) Is Nothing Then Exit Sub
        End If
        Set gobjLastLog = gcolWriteLog("Key_" & strLogName)
        gstrLastLogName = strLogName
    End If
    
    '判断本次使用的日志头和上一次使用的是否相同，若相同，则无需插入日志头，若不同，则从集合中取出对应日志头再进行判断。
    If gstrLastLogInfoHeader <> UCase(gstrDBUser & strModule & strFunction & strLogName) Then
        If gcolLastLogInfoHeader("Key_" & strLogName) <> UCase(gstrDBUser & strModule & strFunction & strLogName) Then
            '判断是否为第一次插入日志头，若不是，则插入一行空行
            If gcolLastLogInfoHeader("Key_" & strLogName) <> "" Then
                gobjLastLog.WriteLine
            End If
            gobjLastLog.WriteLine "[" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "]  " & zlStr.RPAD(gstrDBUser, 8) & "  " & zlStr.RPAD(strModule, 5) & "  " & strFunction
            '缓存日志头到集合
            gcolLastLogInfoHeader.Remove "Key_" & strLogName
            gcolLastLogInfoHeader.Add UCase(gstrDBUser & strModule & strFunction & strLogName), "Key_" & strLogName
        End If
        '缓存日志头到变量
        gstrLastLogInfoHeader = UCase(gstrDBUser & strModule & strFunction & strLogName)
    End If

    '记录日志主体内容
    gobjLastLog.WriteLine Format(Now, "hh:mm:ss") & "," & strLog
    Exit Sub
errh:
    Call ErrCenter
    Call SaveErrLog
End Sub

'清空日志相关变量及对象
Private Sub ClearLogVars()
    Dim i As Long

    gstrLastLogName = ""
    Set gobjLastLog = Nothing
    If Not gcolWriteLog Is Nothing Then
        For i = 1 To gcolWriteLog.count
            If TypeName(gcolWriteLog(i)) <> "String" Then
                If Not gcolWriteLog(i) Is Nothing Then
                    gcolWriteLog(i).Close
                End If
            End If
        Next
    End If
    Set gcolWriteLog = Nothing
    Set gcolLastLogInfoHeader = Nothing
    gstrLastLogInfoHeader = ""
    Set gcolLastLogInfoHeader = Nothing
End Sub

Public Function ErrCenter() As Byte
'------------------------------------------------
'功能： 数据事务错误处理中心
'参数：
'返回： cancel      返回 0
'       resume      返回 1
'------------------------------------------------
    Dim strNote As String, strTemp As String
    Dim bytReturnType As Byte
    Dim blnExeSQLTest As Boolean  '当前的SQLTest状态
    Static mstrErrRecentSQL As String
    
    bytReturnType = 1
    If gcnOracle.Errors.count <> 0 Then
        '处理事务中的错误不允许重试
        If mblnADORoll Then
            If mlngADOErrNum = gcnOracle.Errors(0).NativeError _
                And mstrADOErrDesc = gcnOracle.Errors(0).Description Then
                bytReturnType = 0
            End If
            mblnADORoll = False: mlngADOErrNum = 0: mstrADOErrDesc = ""
        End If
        
        'PL/SQL存储过程错误(包括嵌套的过程调用)
        strNote = gcnOracle.Errors(0).Description
        If InStr(UCase(strNote), "[ZLSOFT]") > 0 Then
            '日志变量
            mbytErrType = 1
            mlngErrNum = gcnOracle.Errors(0).NativeError
            mstrErrInfo = gcnOracle.Errors(0).Description
            MsgBox Split(strNote, "[ZLSOFT]")(1), vbExclamation, App.Title
            Exit Function
        End If

        
        'ORACLE其它错误
        '日志变量
        mbytErrType = 2
        mlngErrNum = gcnOracle.Errors(0).NativeError
        mstrErrInfo = gcnOracle.Errors(0).Description
        
        Select Case gcnOracle.Errors(0).NativeError
        Case 1
            strNote = "已经存在相同内容的数据（要求唯一的内容[如编号、名称等]有重复）。"
            bytReturnType = 0
        Case 903
            strNote = "表名称错误。"
            If gstrRecentSQL <> "" Then mstrErrInfo = mstrErrInfo & vbCrLf & vbCrLf & "错误SQL语句为：" & vbCrLf & vbCrLf & gstrRecentSQL
        Case 904, 920
            strNote = "列名称错误" & vbCrLf & vbCrLf & "SQL语句中使用了不存在的列或语句错误."
            If gstrRecentSQL <> "" Then mstrErrInfo = mstrErrInfo & vbCrLf & vbCrLf & "错误SQL语句为：" & vbCrLf & vbCrLf & gstrRecentSQL
        Case 942
            strNote = "表或视图不存在，很可能是你不具备使用该部分数据的权限或该部分对象同义词缺失。"
            bytReturnType = 0
            
            strTemp = mGetInvalidTable()
            If strTemp <> "" Then
                mstrErrInfo = "（建议请系统管理员重新授权、修复同义词）" & vbCrLf & "请对下列对象进行检查：" & vbCrLf & vbCrLf & vbTab & strTemp
            Else
                mstrErrInfo = "（建议请系统管理员重新授权、修复同义词）" & vbCrLf & "错误SQL语句为：" & vbCrLf & vbCrLf & gstrRecentSQL
            End If
        Case 1000
            strNote = "打开的数据表太多，必要时请系统管理员修改数据库的Open_Cursors配置。"
        Case 1005
            strNote = "错误的用户名或密码。"
        Case 1017
            strNote = "错误的用户名或密码。"
            bytReturnType = 0
        Case 1031
            strNote = "没有足够的权限。"
            bytReturnType = 0
        Case 1045
            strNote = "没有联结数据库的权限。"
            bytReturnType = 0
        Case 1400
            strNote = "由于给主键或要求非空列赋予了空值，导致增加失败。"
            bytReturnType = 0
        Case 1401
            strNote = "由于赋予的值超过了列宽限制，导致增加或更新失败。"
            bytReturnType = 0
        Case 1402
            strNote = "由于赋予的值不符合视图的条件限制，导致增加或更新失败。"
            bytReturnType = 0
        Case 1403
            strNote = "由于未检索到数据，导致后续处理失败。"
        Case 1404
            strNote = "修改列操作，导致相关的索引太大。"
        Case 1405
            strNote = "取得的列值为空。"
        Case 1406
            strNote = "取得的列值被切断而缩短了。"
        Case 1407
            strNote = "由于给主键或要求非空列赋予了空值，导致更新失败。"
            bytReturnType = 0
        Case 1408
            strNote = "指定的列已经建立了索引。"
        Case 1409
            strNote = "不能进行无顺序操作(NoSort)，因为本身就没排序。"
        Case 1410
            strNote = "错误的行ID(ROWID)，行ID必须是数字和字符组成的16进制格式。"
        Case 1411
            strNote = "当前列不能存储超过64K的数据。"
            bytReturnType = 0
        Case 1412
            strNote = "当前列数据类型不能存储零长度字符串。"
            bytReturnType = 0
        Case 1413
            strNote = "错误的小数位数，导致失败。"
            bytReturnType = 0
        Case 1415
            strNote = "不能对一个标签伪列指定外连接[Outer-Join(+)]"
        Case 1416
            strNote = "两张表不能同时指向一个外连接[Outer-Join(+)]"
        Case 1417
            strNote = "一张表只能指定指向不超过一张表的外连接[Outer-Join(+)]"
        Case 1418
            strNote = "指定的索引不存在。"
        Case 1424
            strNote = "错误或无效的换码字符(通配符中只能是'%'或'_')。"
        Case 1425
            strNote = "换码字符必须是长度为1的字符。"
        Case 1426
            strNote = "数值表达式的数据溢出(太大或太小)。"
        Case 1427
            strNote = "单行子查询返回了多行。"
        Case 1428
            strNote = "函数的参数错误或超界。"
        Case 1429
            strNote = "一个二进制日期格式超界。"
        Case 1430
            strNote = "希望增加的列已经存在。"
        Case 1431
            strNote = "授权命令(GRANT)导致内在的不一致。"
        Case 1432
            strNote = "希望删除的公共同义词已经不存在。"
        Case 1433
            strNote = "希望建立的同义词已经存在。"
        Case 1434
            strNote = "希望删除的同义词已经不存在。"
        Case 1435
            strNote = "指定的用户不存在。"
            bytReturnType = 0
        Case 1438
            strNote = "数值超过了列允许的精确程度。"
        Case 1439, 1440, 1441
            strNote = "只有空值列才能修改数据类型、将精度或尺寸减小"
        Case 1536
            strNote = "某个超出表空间的空间限量。"
        Case 2290
            strNote = "由于项目值超过允许的范围（违背了检查约束），导致增加或更新失败。"
            bytReturnType = 0
        Case 2291
            strNote = "由于未填写相关表中存在的项目值(违背了外键约束)，导致增加或更新失败。"
        Case 2091, 2292
            strNote = "因为该记录已经使用，导致删除或更新失败。"
            bytReturnType = 0
        Case 2391
            strNote = "用户已达到数据库所允许的最大登录数。"
        Case 12203
            strNote = "由于主机串书写、配置或服务器问题，不能正常连接。"
            bytReturnType = 0
        Case 20003
            strNote = "存储过程无效，请对失效的存储过程进行编译。"
            If gstrRecentSQL <> "" Then mstrErrInfo = mstrErrInfo & vbCrLf & vbCrLf & "错误SQL语句为：" & vbCrLf & vbCrLf & gstrRecentSQL
        Case Else
            strTemp = Err.Description
            If InStr(strTemp, "PLS-00201") > 0 And InStr(strTemp, "ZL_") > 0 Then
                Dim lngPos As Long
                
                lngPos = InStr(strTemp, "ZL_")
                strTemp = Mid(strTemp, lngPos)
                strTemp = Mid(strTemp, 1, InStr(strTemp, "'") - 1)
                
                strNote = "请在服务器管理工具的角色管理程序中增加对过程“" & strTemp & "”的授权。"
            Else
                strNote = "未知错误，发生在" & gcnOracle.Errors(0).Source
            End If
            If gstrRecentSQL <> "" Then mstrErrInfo = mstrErrInfo & vbCrLf & vbCrLf & "错误SQL语句为：" & vbCrLf & vbCrLf & gstrRecentSQL
        End Select
        
    Else
        'VB标准错误
        '日志变量
        mbytErrType = 3
        mlngErrNum = Err.Number
        mstrErrInfo = Err.Description
        
        Select Case Err.Number
            Case 3, 3 - 2146828288
                strNote = "未采用标准返回过程"
            Case 5, 5 - 2146828288
                strNote = "无效的过程或参数"
            Case 6, 6 - 2146828288
                strNote = "数据溢出"
            Case 7, 7 - 2146828288
                strNote = "内存溢出"
            Case 9, 9 - 2146828288
                strNote = "下标超界"
            Case 10, 10 - 2146828288
                strNote = "数组是固定数组或暂时锁定"
            Case 11, 11 - 2146828288
                strNote = "除数为零太小"
            Case 13, 13 - 2146828288
                strNote = "类型不匹配"
            Case 14, 14 - 2146828288
                strNote = "超过字符串允许长度"
            Case 16, 16 - 2146828288
                strNote = "表达式太复杂"
            Case 17, 17 - 2146828288
                strNote = "不支持要求的操作"
            Case 18, 18 - 2146828288
                strNote = "发生了用户中断"
            Case 20, 20 - 2146828288
                strNote = "无错误返回"
            Case 28, 28 - 2146828288
                strNote = "堆栈空间溢出"
            Case 35, 35 - 2146828288
                strNote = "过程或函数未定义"
            Case 47, 47 - 2146828288
                strNote = " 太多的动态联结库（DLL）应用客户"
            Case 48, 48 - 2146828288
                strNote = " 调用动态联结库（DLL）错误"
            Case 49, 49 - 2146828288
                strNote = " 动态联结库（DLL）约定错误"
            Case 51, 51 - 2146828288
                strNote = "内部错误"
            Case 52, 52 - 2146828288
                strNote = "错误的文件名或文件号"
            Case 53, 53 - 2146828288
                strNote = "文件未找到"
            Case 54, 54 - 2146828288
                strNote = "文件格式错误"
            Case 55, 55 - 2146828288
                strNote = "文件已经打开"
            Case 57, 57 - 2146828288
                strNote = "设备输入 / 输出错误"
            Case 58, 58 - 2146828288
                strNote = "文件已经存在"
            Case 59, 59 - 2146828288
                strNote = "错误的记录长度"
            Case 61, 61 - 2146828288
                strNote = "磁盘满"
            Case 62, 62 - 2146828288
                strNote = "输入超过文件尾"
            Case 63, 63 - 2146828288
                strNote = "错误的记录号"
            Case 67, 67 - 2146828288
                strNote = "文件太多"
            Case 68, 68 - 2146828288
                strNote = "设备无效或不支持"
            Case 70, 70 - 2146828288
                strNote = "拒绝访问"
            Case 71, 71 - 2146828288
                strNote = "磁盘未准备好"
            Case 74, 74 - 2146828288
                strNote = "不能命名为不同的驱动器"
            Case 75, 75 - 2146828288
                strNote = "路径 / 文件访问错误"
            Case 76, 76 - 2146828288
                strNote = "路径未找到"
            Case 91, 91 - 2146828288
                strNote = "对象变量或块变量为定义(未新建实例)"
            Case 92, 92 - 2146828288
                strNote = "循环未初始化"
            Case 93, 93 - 2146828288
                strNote = "错误的模式字符串"
            Case 94, 94 - 2146828288
                strNote = "错误地使用空(Null)"
            Case 96, 96 - 2146828288
                strNote = " 由于已经使用的对象时间超过了其设置的最大元素号，导致不可能进入事件"
            Case 97, 97 - 2146828288
                strNote = "不能调用一个未建立实例的类对象函数"
            Case 98, 98 - 2146828288
                strNote = " 不能使用一个私有对象的属性和方法?参数和返回值"
            Case 321, 321 - 2146828288
                strNote = "错误的文件格式"
            Case 322, 322 - 2146828288
                strNote = "不能创建需要的临时文件"
            Case 325, 325 - 2146828288
                strNote = "资源文件中错误的格式"
            Case 380, 380 - 2146828288
                strNote = "错误的属性值"
            Case 381, 381 - 2146828288
                strNote = "错误的属性数组索引"
            Case 382, 382 - 2146828288
                strNote = "不支持的运行时设置"
            Case 383, 383 - 2146828288
                strNote = "不支持的只读属性设置"
            Case 385, 384 - 2146828288
                strNote = "需要属性数组索引"
            Case 387, 387 - 2146828288
                strNote = "不允许的设置"
            Case 393, 393 - 2146828288
                strNote = "不支持的运行时读取"
            Case 394, 394 - 2146828288
                strNote = "不支持的只写属性读取"
            Case 422, 422 - 2146828288
                strNote = "不存在的属性"
            Case 423, 423 - 2146828288
                strNote = "不存在的属性或方法"
            Case 424, 424 - 2146828288
                strNote = "要求一个对象"
            Case 429, 429 - 2146828288
                strNote = "ActiveX不能创建部件"
            Case 430, 430 - 2146828288
                strNote = "类不支持的自动化操作或不支持的界面"
            Case 432, 432 - 2146828288
                strNote = "在自动操作期间未找到文件名或类名称"
            Case 438, 438 - 2146828288
                strNote = "对象不支持该属性或方法"
            Case 440, 440 - 2146828288
                strNote = "自动化对象错误"
            Case 442, 442 - 2146828288
                strNote = "到远程类库或对象库的联结丢失，按OK进入对话移去参照"
            Case 443, 443 - 2146828288
                strNote = "自动化对象没有缺省值"
            Case 445, 445 - 2146828288
                strNote = "对象不支持这种操作"
            Case 446, 446 - 2146828288
                strNote = "对象不支持命名参数"
            Case 447, 447 - 2146828288
                strNote = "对象不支持当前本地设置"
            Case 448, 448 - 2146828288
                strNote = "命名参数未找到"
            Case 449, 449 - 2146828288
                strNote = "参数不是可选的"
            Case 450, 450 - 2146828288
                strNote = "错误的参数个数和属性分配"
            Case 451, 451 - 2146828288
                strNote = "属性赋值(Let)过程和读取(Get)过程不返回对象"
            Case 452, 452 - 2146828288
                strNote = "无效的序号"
            Case 453, 453 - 2146828288
                strNote = "指定的DLL函数未找到"
            Case 454, 454 - 2146828288
                strNote = "代码资源未找到"
            Case 455, 455 - 2146828288
                strNote = "代码资源锁定错误"
            Case 457, 457 - 2146828288
                strNote = "该关键值已经与集合的另一元素结合"
            Case 458, 458 - 2146828288
                strNote = "VB不支持的可变自动化类型"
            Case 459, 459 - 2146828288
                strNote = "对象和类不支持的事件集"
            Case 460, 460 - 2146828288
                strNote = "错误的剪贴板格式"
            Case 461, 461 - 2146828288
                strNote = "方法或数据成员未找到"
            Case 462, 462 - 2146828288
                strNote = "远程服务器不存在或无效"
            Case 463, 463 - 2146828288
                strNote = "类没有在本地注册"
            Case 481, 481 - 2146828288
                strNote = "无效的图片格式"
            Case 482, 482 - 2146828288
                strNote = "打印机错误"
            Case 735, 735 - 2146828288
                strNote = "不能将存储为临时文件"
            Case 744, 744 - 2146828288
                strNote = "未找到搜索的主题"
            Case 746, 746 - 2146828288
                strNote = "太长的复制内容"
                
            'ADO错误
            Case -2147483647
                strNote = "未实现"
            Case -2147483646
                strNote = "内存不足"
            Case -2147483645
                strNote = "一个或多个参数无效"
            Case -2147483644
                strNote = "不支持这样的接口"
            Case -2147483643
                strNote = "无效指针"
            Case -2147483642
                strNote = "无效句柄"
            Case -2147483641
                strNote = "操作终止"
            Case -2147483640
                strNote = "不确定的错误"
            Case -2147483639
                strNote = "一般访问拒绝错误"
            Case -2147483638
                strNote = "完成操作所必需的数据不再可用"
            Case -2147467263
                strNote = "未实现"
            Case -2147467262
                strNote = "不支持这样的接口"
            Case -2147467261
                strNote = "无效指针"
            Case -2147467260
                strNote = "操作终止"
            Case -2147467259
                strNote = "不确定的错误"
            Case -2147467258
                strNote = "线程本地存储失败"
            Case -2147467257
                strNote = "获取共享的内存分配程序失败"
            Case -2147467256
                strNote = "获取内存分配程序失败"
            Case -2147467255
                strNote = "不能初始化类的高速缓存"
            Case -2147467254
                strNote = "不能初始化RPC服务"
            Case -2147467253
                strNote = "不能设置线程本地存储通道控制"
            Case -2147467252
                strNote = "不能分配线程本地存储通道控制"
            Case -2147467251
                strNote = "用户提供的内存分配程序不可接受"
            Case -2147467250
                strNote = "OLE服务互斥量已存在"
            Case -2147467249
                strNote = "OLE服务文件映射已存在"
            Case -2147467247
                strNote = "试图启动OLE服务失败"
            Case -2147467246
                strNote = "在单线程模型中试图再一次调用CoInitialize"
            Case -2147467245
                strNote = "需要一个远程激活，但是不允许"
            Case -2147467244
                strNote = "需要一个远程激活，但是提供的服务器名称无效"
            Case -2147467243
                strNote = "类运行配置的安全id与调用者不同"
            Case -2147467242
                strNote = "使用OLE1服务所需的DDE窗口被禁止"
            Case -2147467241
                strNote = "RunAs指定的必须是域名\用户名或只是用户名"
            Case -2147467240
                strNote = "服务进程不能启动，可能路径名不正确"
            Case -2147467239
                strNote = "当配置标识时服务进程不能启动，路径名可能不正确或无效"
            Case -2147467238
                strNote = "由于配置标识不正确，服务进程不能启动。检查用户名和口令"
            Case -2147467237
                strNote = "不允许客户启动这个服务器"
            Case -2147467236
                strNote = "提供这个服务的服务器不能启动"
            Case -2147467235
                strNote = "本计算机不能和服务器提供的其他计算机通信"
            Case -2147467234
                strNote = "服务器启动后不响应"
            Case -2147467233
                strNote = "服务器的注册信息不一致或不完整"
            Case -2147467232
                strNote = "这个接口的注册信息不一致或不完整"
            Case -2147467231
                strNote = "不支持试图执行的操作"
            Case -2147418113
                strNote = "灾难性失败"
            Case -2147024891
                strNote = "一般访问拒绝错误"
            Case -2147024890
                strNote = "无效句柄"
            Case -2147024882
                strNote = "内存不足"
            Case -2147024809
                strNote = "一个或多个参数无效"
            Case 3000
                strNote = "提供者执行请求的动作失败"
            Case 3001
                strNote = "参数类型错误，或数值超过范围，或与其他类型互相冲突。"
            Case 3002
                strNote = "当打开请求的文件时，发生错误"
            Case 3003
                strNote = "读指定的文件时出错"
            Case 3004
                strNote = "写文件时有错误"
            Case 3021
                strNote = "BOF和EOF中一个为True，或者当前记录已被删，而应用程序的请求操作需要当前记录"
            Case 3219
                strNote = "上下文环境不允许当前应用操作（可能是处于尚未结束的事务）。"
            Case 3220
                strNote = "不能改变提供者"
            Case 3246
                strNote = "在事务执行中，不能关闭一个联结对象。"
            Case 3251
                strNote = "提供者不支持该应用程序请求的操作。"
            Case 3265
                strNote = "ADO没找到应用程序要求的对应名称或序号（可能是列名称错误）。"
            Case 3367
                strNote = "对象已在集合中，不能追加"
            Case 3420
                strNote = "对象未引用或引用的对象不再有效。"
            Case 3421
                strNote = "当前操作使用了错误的数值类型。"
            Case 3704
                strNote = "如果对象已关闭，不允许应用程序请求的操作"
            Case 3705
                strNote = "如果对象已打开，不允许应用程序请求的操作"
            Case 3706
                strNote = "ADO不能找到指定的提供者"
            Case 3707
                strNote = "不能采用命令对象改变一个记录集的活动连接源等属性。"
            Case 3708
                strNote = "应用程序出现错误的参数定义。"
            Case 3709
                strNote = "应用程序请求对一个对象的操作时使用了一个引用，而该引用指向了一个关闭的或无效的Connection对象"
            Case 3710
                strNote = "操作不能重新执行"
            Case 3711
                strNote = "操作仍然在执行"
            Case 3712
                strNote = "操作被取消"
            Case 3713
                strNote = "操作仍然在连接中"
            Case 3714
                strNote = "事务无效"
            Case 3715
                strNote = "操作不在执行过程中"
            Case 3716
                strNote = "在这种情况下运行不安全"
            Case 3717
                strNote = "操作引出一个安全对话"
            Case 3718
                strNote = "操作引出一个安全对话头"
            Case 3719
                strNote = "违背数据的完整性，操作失败。"
            Case 3720
                strNote = "用户没有足够的权限完成操作，操作失败。"
            Case 3721
                strNote = "数据超出给定的数据类型的范围"
            Case 3722
                strNote = "动作违背了模式"
            Case 3723
                strNote = "表达式包含不匹配的符号"
            Case 3724
                strNote = "不能转换值不能创建资源"
            Case 3726
                strNote = "这一行中不存在指定的列"
            Case 3727
                strNote = "URL不存在"
            Case 3728
                strNote = "没有查看目录树的权限"
            Case 3729
                strNote = "提供的URL无效"
            Case 3730
                strNote = "资源被锁定"
            Case 3731
                strNote = "资源已经存在"
            Case 3732
                strNote = "不能完成动作"
            Case 3733
                strNote = "文件版本信息没找到"
            Case 3734
                strNote = "服务器得不到足够的空间完成操作，操作失败"
            Case 3735
                strNote = "资源超出范围"
            Case 3736
                strNote = "命令不可用"
            Case 3737
                strNote = "在命名的行中的URL不存在"
            Case 3738
                strNote = "不能删除资源，这超出了允许范围"
            Case 3739
                strNote = "对于选择的列，这个属性无效"
            Case 3740
                strNote = "给属性提供了一个无效的选择"
            Case 3741
                strNote = "给属性提供了一个无效的值"
            Case 3742
                strNote = "设置这个属性造成和其他属性冲突"
            Case 3743
                strNote = "不是所有的属性都能被设置"
            Case 3744
                strNote = "属性没有被设置"
            Case 3745
                strNote = "属性不能被设置"
            Case 3746
                strNote = "属性不被支持"
            Case 3747
                strNote = "类别没有设置所以动作不能执行"
            Case 3748
                strNote = "不能改变连接"
            Case 3749
                strNote = "Fields集合的Update方法失败"
            Case 3750
                strNote = "不能设置Deny权限，因为提供者不支持"
            Case 3751
                strNote = "提供者不支持请求的Deny类型"
                
            Case Else
                strNote = "发生未知的界面错误"
        End Select
        bytReturnType = 0
    End If
    
    
    g_AutoConnect = Val(GetSetting("ZLSOFT", "公共全局\网络断网自动重连", "AutoConnect", 0))
    gblnAutoConnect = g_AutoConnect
    If gblnAutoConnect Then '是否使用网络断开自动连接功能
        Dim blnConnect As Boolean
        Dim blnNumConnect As Boolean '检查次数是否重新连接
        Dim blnStatus As Boolean '是否其他错误引发的网络问题
        '通过过滤错误信息,检查是否为网络问题引发的错误。mbytErrType=2 Oracle提供的错误信息 mbytErrType=3 VB提供的错误信息
        If mbytErrType = 3 Then
            If mlngErrNum = -2147467259 Or mlngErrNum = -2147217900 Or mlngErrNum = 3709 Then
                '检查VB具体错误信息
                If CheckErrConnectInfo(mlngErrNum, strNote, mstrErrInfo, 1) Then
                      
                    '判断相同错误,如果2次以上正常错误提示。
                    If mstrErrRecentSQL = gstrRecentSQL And gstrRecentSQL <> "" Then
                        mlngConnectCount = mlngConnectCount + 1
                        If mlngConnectCount > 2 Then
                            blnNumConnect = False  '正常错误提示
                            mlngConnectCount = 0 '还原计数器
                        Else
                            blnNumConnect = True
                        End If
                    Else
                        mstrErrRecentSQL = gstrRecentSQL
                        mlngConnectCount = 1
                        blnNumConnect = True
                    End If
                Else
                    blnConnect = False '正常错误提示
                End If
            End If
        Else
            '错误号12543 TNS: 无法连接目标主机,1012-没有登录，0028-会话被终止
            If mlngErrNum = -2147467259 Or mlngErrNum = -2147217900 Or mlngErrNum = 0 Or mlngErrNum = 12543 Or mlngErrNum = 2399 Or mlngErrNum = 2396 Or mlngErrNum = 1012 Or mlngErrNum = 28 Then
                '检查ORACLE具体错误信息
                If CheckErrConnectInfo(mlngErrNum, strNote, mstrErrInfo, 2) Then
                    '判断相同错误,如果2次以上正常错误提示。
                    If mstrErrRecentSQL = gstrRecentSQL And gstrRecentSQL <> "" Then
                        mlngConnectCount = mlngConnectCount + 1
                        If mlngConnectCount > 2 Then
                            blnNumConnect = False  '正常错误提示
                            mlngConnectCount = 0 '还原计数器
                        Else
                            blnNumConnect = True
                        End If
                    Else
                        mstrErrRecentSQL = gstrRecentSQL
                        mlngConnectCount = 1
                        blnNumConnect = True
                    End If
                Else
                    blnConnect = False '正常错误提示
                End If
            End If
        End If
        
        '自动重新连接一次,检查是否能自动重新连接
        If blnNumConnect Then '与ORACLE连接已经断开
            blnExeSQLTest = mblnExeSQLTest
            mblnExeSQLTest = True
            If CheckAdoConnction(blnStatus) Then
                If blnStatus Then
                   blnConnect = False '正常错误提示
                Else
                   blnConnect = True '提示重连
                End If
            Else
                '与ORACLE重新连接成功,不需要提示。直接返回重新执行。
                blnConnect = False
                ErrCenter = 1
                mblnExeSQLTest = blnExeSQLTest
                Exit Function
            End If
            mblnExeSQLTest = blnExeSQLTest
        End If
    End If
    If bytReturnType = 1 Then
        ErrCenter = frmErrAsk.ShowEdit(mlngErrNum, strNote, mstrErrInfo, blnConnect)
    Else
        Call frmErrNote.ShowEdit(mlngErrNum, strNote, mstrErrInfo, blnConnect)
        ErrCenter = 0
    End If
    
    '清除错误
    Err.Clear
End Function

Public Sub SaveErrLog()
'功能：将刚才的错误信息写入数据库错误日志
    Dim strSQL As String
    
    If mlngErrNum <> 0 And mbytErrType <> 0 And gblnErrLog Then
        On Local Error Resume Next
        
        strSQL = "Zl_Zlerrorlog_Insert('" & gstrComputerName & "'," & mbytErrType & "," & mlngErrNum & "," & _
            gobjComLib.zlStr.SQLAdjust(mstrErrInfo) & "," & glngSessionID & ")"
        Call zlDatabase.ExecuteProcedure(strSQL, "保存错误日志")
        mlngErrNum = 0: mstrErrInfo = "": mbytErrType = 0
    End If
End Sub

Public Sub ShowAbout(frmParent As Object, strTitle As String, strName As String, strVersion As String)
'功能： 显示关于窗体
'参数：
'   frmParent：关于窗体的父窗体
'   strTitle=App.Title
'   strName=App.ProductName
'   strVersion=App.Major & "." & App.Minor & "." & App.Revision
    Dim rsTmp As New ADODB.Recordset
    Dim blnFind As Boolean, strAPP As String
    Dim i As Integer, objItem As ListItem
    Dim strCode As String
    Dim strWhere As String, strSQL As String
    Dim strBT As String, strTag As String   '关于窗体中的 旗舰版标识
    Dim objFile As New FileSystemObject
    Dim strDate As String
    Dim lngMaxWidth As Long
        
    On Error GoTo errh
    
    With frmAbout
        '取部件版本号
        strAPP = GetSetting("ZLSOFT", "公共全局", "程序路径", "")
        '测试代码
'        strAPP = "C:\APPsoft\ZLHIS+.EXE" '指定APP路径
'        strName = "zl9CustAcc" '测试PUBLIC部件
'        strName = "ZL9MouseWheel" '测试System32部件
'       strName = "zl9PacsWork"
        If strAPP <> "" Then
            If objFile.FileExists(strAPP) Then
                strWhere = objFile.GetParentFolderName(strAPP) & "\Apply\" & strName & ".dll"
                blnFind = objFile.FileExists(strWhere)
                If Not blnFind Then
                    strWhere = objFile.GetParentFolderName(strAPP) & "\PUBLIC\" & strName & ".dll"
                    blnFind = objFile.FileExists(strWhere)
                    If Not blnFind Then
                        strWhere = objFile.GetSpecialFolder(SystemFolder) & "\" & strName & ".dll"
                        blnFind = objFile.FileExists(strWhere)
                    End If
                End If
                If blnFind Then
                    strDate = Format(objFile.GetFile(strWhere).DateLastModified, "yyyy-MM-dd HH:mm")
                End If
            End If
        End If
        '系统与注册版本读取
        strSQL = "Select 系统, 注册产品名称, 注册产品简名, 注册产品版本 From zlComponent Where Upper(部件) = '" & UCase(strName) & "' Order By 系统"
        rsTmp.CursorLocation = adUseClient
        rsTmp.Open strSQL, gcnOracle, adOpenKeyset
        If zlRegInfo("产品简名", , -1) <> "中联" Then
            '旗舰版标识
            strTag = ""
            strBT = zlRegInfo("产品标题")
            If strBT <> "" Then
                If InStr(strBT, "-") > 0 Then
                    If Split(strBT, "-")(1) = "Ultimate" Then
                        strTag = "旗舰版"
                    ElseIf Split(strBT, "-")(1) = "Professional" Then
                        strTag = "专业版"
                    End If
                End If
            End If
            strBT = Split(strBT, "-")(0)
            .lblName.Caption = strBT & IIf(strTag = "", "", "(" & strTag & ")")
            If rsTmp.RecordCount <> 0 Then
                strSQL = "Select 版本号 from zlsystems where 编号=" & rsTmp!系统
                Set rsTmp = New ADODB.Recordset
                rsTmp.CursorLocation = adUseClient
                rsTmp.Open strSQL, gcnOracle, adOpenKeyset
                .lblRegVersion.Caption = "v" & rsTmp!版本号
            Else
                .lblRegVersion.Caption = ""
            End If
        Else
            If rsTmp.RecordCount <> 0 Then
                .lblName.Caption = rsTmp!注册产品名称 & ""
                .lblRegVersion.Caption = rsTmp!注册产品简名 & " v" & rsTmp!注册产品版本
            Else
                .lblName.Caption = "": .lblRegVersion.Caption = ""
            End If
        End If
        'Logo设置
        Call ApplyOEM_Picture(.imgLogo, "Picture")
        '下面版本与上面右对齐
        lngMaxWidth = IIf(.lblName.Width > .lblRegVersion.Width, .lblName.Width, .lblRegVersion.Width)
        .lblName.Left = .imgLogo.Left * 2 + .imgLogo.Width + lngMaxWidth - .lblName.Width
        .lblRegVersion.Left = .lblName.Left + .lblName.Width - .lblRegVersion.Width
        '编译信息
        If strTitle = strName Then
            .lblCompile.Caption = "编译信息：" & strTitle & " " & strVersion & IIf(strDate <> "", "，" & strDate, "")
        Else
            .lblCompile.Caption = "编译信息：" & strTitle & "，" & strName & " " & strVersion & IIf(strDate <> "", "，" & strDate, "")
        End If

        '其他信息设置
        .lblGrant.Caption = Replace(zlRegInfo("单位名称", , -1), ";", vbCrLf)
        .lbl技术支持商.Caption = zlRegInfo("技术支持商", , -1)
        If Trim$(.lbl技术支持商.Caption) = "" Then
            .Label1.Visible = False
            .lbl技术支持商.Visible = False
        Else
            .Label1.Visible = True
            .lbl技术支持商.Visible = True
        End If
        strCode = zlRegInfo("产品开发商", , -1)
        If Trim(strCode) = "" Then
            .Label3.Visible = False
            .lbl开发商.Visible = False
        Else
            .Label3.Visible = True
            .lbl开发商.Visible = True
            
            .lbl开发商.Caption = ""
            For i = 0 To UBound(Split(strCode, ";"))
                .lbl开发商.Caption = .lbl开发商.Caption & Split(strCode, ";")(i) & vbCrLf
            Next
        End If
        Set rsTmp = New ADODB.Recordset
        '可用功能
        If strTitle Like "*导航台" Then
            strSQL = "Select Distinct A.序号,A.标题,A.说明 From zlPrograms A,zlProgFuncs B,zlregFunc C" & _
                " Where A.系统=B.系统 And A.序号=B.序号 And trunc(a.系统/100)=C.系统 and A.序号=C.序号  Order by 序号"
        Else
            strSQL = "Select Distinct A.序号,A.标题,A.说明 From zlPrograms A,zlProgFuncs B,zlregFunc C" & _
                " Where A.系统=B.系统 And A.序号=B.序号 And Upper(A.部件)='" & UCase(strName) & "'" & _
                " And trunc(a.系统/100)=C.系统 and A.序号=C.序号 Order by 序号"
        End If
        rsTmp.CursorLocation = adUseClient
        rsTmp.Open strSQL, gcnOracle, adOpenKeyset
        .lvw.ListItems.Clear
        For i = 1 To rsTmp.RecordCount
            Set objItem = .lvw.ListItems.Add(, , rsTmp!序号, , 1)
            objItem.SubItems(1) = rsTmp!标题
            objItem.SubItems(2) = IIf(IsNull(rsTmp!说明), "", rsTmp!说明)
            rsTmp.MoveNext
        Next
        .Refresh
    End With
    If frmParent Is Nothing Then
        frmAbout.Show 1
    Else
        frmAbout.Show 1, frmParent
    End If
    Exit Sub
errh:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Sub SaveReportDate(lngProgId As Long, strBegin As String, strEnd As String)
'功能：将指定的报表时间范围条件存入注册表
'参数：lngProgID=报表模块号
'      strBegin=开始时间,格式yyyy-MM-dd [hh:mm:ss]
'      strEnd=结束时间,格式yyyy-MM-dd [hh:mm:ss]
    Dim intCount As Integer, i As Integer
    
    intCount = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\报表\" & lngProgId & "\历史查询", "Count", 0)
    intCount = IIf(intCount = 100, 100, intCount + 1)
    
    For i = intCount To 2 Step -1
        SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\报表\" & lngProgId & "\历史查询", i, GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\报表\" & lngProgId & "\历史查询", i - 1, "")
    Next
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\报表\" & lngProgId & "\历史查询", "Count", intCount
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\报表\" & lngProgId & "\历史查询", 1, strBegin & " 至 " & strEnd
End Sub

Public Function GetReportDate(lngProgId As Long) As String
'功能:读取指定报表的历史查询时间范围
'参数:lngProgID=报表模块号
'返回:如"1:2000-01-01 至 2000-01-31 23:59:59;2:..."
    Dim intCount As Integer, i As Integer
    
    intCount = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\报表\" & lngProgId & "\历史查询", "Count", 0)
    For i = 1 To intCount
        GetReportDate = GetReportDate & ";" & i & ":" & GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\报表\" & lngProgId & "\历史查询", i, "")
    Next
    GetReportDate = Mid(GetReportDate, 2)
End Function

Public Sub SQLTest(Optional strProject As String, Optional strForm As String, Optional ByVal strSQL As String, Optional strNote As String)
'功能：将部件中执行的SQL语句输出到窗体或文件中，并附加开始结束时间，执行时间
'参数：strProject=部件名称,具体可取App.Title
'      strForm=窗体名,具体可取Form.Caption
'      strSQL=将要执行的SQL语句,在Open时传入,如果不传，表示最近一次SQL执行完毕
'      strNote=SQL语句说明
'说明：
'     1.窗体中包含一个RTF控件,其内容可以保存到文件,可以编辑。
'     2. 该窗体在第一次调用时显示，主程序结束时关闭。
    Dim strTmp As String, strInsert As String
    Dim sngEnd As Single, i As Integer
    Dim strPlanEx As String
    Dim arrTmp As Variant
    
    If mblnExeSQLTest Then Exit Sub
    gstrRecentSQL = strSQL  '保存最近执行的SQL语句
    If Not (gblnSQLTest Or gblnSQLLog Or gblnSQLPlan) Then Exit Sub
    mblnExeSQLTest = True
    
    If strSQL <> "" Then
        If gblnSQLLog Then
            If gobjLogFile Is Nothing Then
                Set gobjLogFile = New FileSystemObject
                On Local Error Resume Next
                Set gobjLogText = gobjLogFile.OpenTextFile("SQL_" & gstrDBUser & "_" & Format(Date, "yyyyMMdd") & ".log", ForAppending, True, TristateFalse)
                On Local Error GoTo 0
            End If
            If Not gobjLogText Is Nothing Then
                strTmp = "[" & Format(Time, "HH:mm:ss") & "]"
                gobjLogText.WriteLine strTmp & "Application:" & strProject & "\" & strForm
                gobjLogText.WriteLine strTmp & "SQL:" & strSQL
                msngTime = timer
            End If
        ElseIf gblnSQLPlan Then
            mstrSQLPlan = ""
            If gcnSysConn Is Nothing And gstrSysUser = "" Then
                '必须使用SYS登录
                gstrSysUser = "SYS"
'                Set gcnSysOLEDB = New ADODB.Connection
                Call zlDatabase.UserIdentify(Nothing, "执行计划跟踪必须先使用SYS用户连接,否则不能生成日志文件。", 0, 0, "SYS", gcnSysConn, True, , gcnSysOLEDB)
            End If
            If Not gcnSysConn Is Nothing Then
                If gobjPlanExFile Is Nothing Then
                    Set gobjPlanExFile = New FileSystemObject
                    On Local Error Resume Next
                    Set gobjPlanExText = gobjPlanExFile.OpenTextFile("SQLPlanEx_" & gstrDBUser & "_" & Format(Date, "yyyyMMdd") & ".log", ForAppending, True, TristateFalse)
                    On Local Error GoTo 0
                End If
                If Not gobjPlanExText Is Nothing Then
                    strTmp = "[" & Format(Time, "HH:mm:ss") & "]"
                    mstrSQLPlan = strTmp & "Application:" & strProject & "\" & strForm & ",SQL:" & vbCrLf & strSQL
                    msngTime = timer
                End If
            End If
        ElseIf gblnSQLTest Then
            If Not gblnShow Then
                frmSQLTest.Show
                gblnShow = True
            End If
            If frmSQLTest.rtf.Visible Then
                With frmSQLTest.rtf
                    If Len(.Text) > 10000 Then .Text = ""
                    .SelStart = Len(.Text)
                    If .Text <> "" Then
                        If Right(.Text, 2) <> vbCrLf Then
                            .SelText = vbCrLf & vbCrLf
                        Else
                            .SelText = vbCrLf
                        End If
                    End If
                    .SelColor = &H800000
                    .SelText = "●部件：" & strProject & "，位置：" & strForm & vbCrLf
                    
                    .SelColor = &H800000
                    .SelText = "●SQL语句：" & strNote & vbCrLf
                    
                    .SelColor = &H808080
                    .SelText = strSQL
                    
                    If Right(.Text, 2) <> vbCrLf Then .SelText = vbCrLf
                    .SelColor = &H800000
                    msngTime = timer
                    .SelText = "●开始时间：" & Format(msngTime, "0.0000") & "，"
                End With
            ElseIf frmSQLTest.txt.Visible Then
                With frmSQLTest.txt
                    If Len(.Text) > 10000 Then .Text = ""
                    .SelStart = Len(.Text)
                    If .Text <> "" Then
                        If Right(.Text, 2) <> vbCrLf Then
                            .SelText = vbCrLf & vbCrLf
                        Else
                            .SelText = vbCrLf
                        End If
                    End If
                    .SelText = "●部件：" & strProject & "，位置：" & strForm & vbCrLf
                    
                    .SelText = "●SQL语句：" & strNote & vbCrLf
                    
                    .SelText = strSQL
                    
                    If Right(.Text, 2) <> vbCrLf Then .SelText = vbCrLf
                    msngTime = timer
                    .SelText = "●开始时间：" & Format(msngTime, "0.0000") & "，"
                End With
            ElseIf frmSQLTest.txtObject.Visible Then
                With frmSQLTest.txtObject
                    strInsert = ""
                    If UCase(strSQL) Like UCase("Select *_ID.Nextval From Dual") Then
                        strTmp = Replace(UCase(Replace(UCase(strSQL), UCase("Select"), "")), UCase(".Nextval From Dual"), "")
                        strInsert = "Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(XXXX,YYYY,'基本','" & strTmp & "',USER,'SELECT');"
                        If InStr(.Text, strInsert) = 0 Then .Text = .Text & strInsert & vbCrLf
                    ElseIf UCase(strSQL) Like "ZL*_*(*)" Then
                        strTmp = Left(strSQL, InStr(strSQL, "(") - 1)
                        strInsert = "Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(XXXX,YYYY,'基本','" & strTmp & "',USER,'EXECUTE');"
                        If InStr(.Text, strInsert) = 0 Then .Text = .Text & strInsert & vbCrLf
                    Else
                        strTmp = SQLObject(strSQL)
                        For i = 0 To UBound(Split(strTmp, ","))
                            If UCase(Split(strTmp, ",")(i)) <> "DUAL" Then
                                strInsert = "Insert Into zlProgPrivs(系统,序号,功能,对象,所有者,权限) Values(XXXX,YYYY,'基本','" & Split(strTmp, ",")(i) & "',USER,'SELECT');"
                                If InStr(.Text, strInsert) = 0 Then .Text = .Text & strInsert & vbCrLf
                            End If
                        Next
                    End If
                End With
            End If
        End If
    Else
        If gblnSQLLog Then
            If Not gobjLogText Is Nothing Then
                sngEnd = timer
                strTmp = "[" & Format(Time, "HH:mm:ss") & "]"
                gobjLogText.WriteLine strTmp & "Expend:" & Format(sngEnd - msngTime, "0.0000")
                gobjLogText.WriteBlankLines 1
            End If
        ElseIf gblnSQLPlan Then
            If Not gcnSysConn Is Nothing Then
                If Not gobjPlanExText Is Nothing Then
                    sngEnd = timer
                    strTmp = "[" & Format(Time, "HH:mm:ss") & "]"
                    
                    If CheckSQLPlan(strPlanEx) Then
                        For i = 0 To UBound(Split(mstrSQLPlan, vbCrLf))
                            gobjPlanExText.WriteLine Split(mstrSQLPlan, vbCrLf)(i)
                        Next
                        arrTmp = Split(strPlanEx, vbCrLf)
                        For i = 0 To UBound(arrTmp)
                            gobjPlanExText.WriteLine arrTmp(i)
                        Next
                        
                        gobjPlanExText.WriteLine strTmp & "Expend:" & Format(sngEnd - msngTime, "0.0000") & "秒"
                        gobjPlanExText.WriteBlankLines 1
                    End If
                End If
                mstrSQLPlan = ""
            End If
        ElseIf gblnSQLTest Then
            sngEnd = timer
            If Not frmSQLTest.txtObject.Visible Then
                If frmSQLTest.rtf.Visible Then
                    With frmSQLTest.rtf
                        .SelStart = Len(.Text)
                        
                        .SelColor = &H800000
                        .SelText = "结束时间：" & Format(sngEnd, "0.0000") & "，"
                        
                        .SelColor = vbRed
                        .SelText = "执行时间：" & Format(sngEnd - msngTime, "0.0000")
                    End With
                ElseIf frmSQLTest.txt.Visible Then
                    With frmSQLTest.txt
                        .SelStart = Len(.Text)
                        
                        .SelText = "结束时间：" & Format(sngEnd, "0.0000") & "，"
                        
                        .SelText = "执行时间：" & Format(sngEnd - msngTime, "0.0000")
                    End With
                End If
            End If
        End If
    End If
    gblnOK = False
    mblnExeSQLTest = False
End Sub

Private Function GetSQLPlan() As ADODB.Recordset
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim strSQLID As String, lngPrev_Num As Long
        
    On Error GoTo errh
    
    strSQL = "select prev_sql_id,PREV_CHILD_NUMBER from gv$session where AUDSID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "GetSQLPlan", glngSessionID)
    If rsTmp.RecordCount = 0 Then Exit Function
    strSQLID = rsTmp!prev_sql_id & ""
    lngPrev_Num = Val(rsTmp!PREV_CHILD_NUMBER & "")
    
    strSQL = "Select null as ID,null as Operation,null as Name,null as Cardinality,null as Bytes,null as Cost,null as Time ,null as Object_Owner, '|                      Operation                    |             Name             |  Rows | Bytes|  Cost  | Pstart| Pstop |' As ""PlanTable"" From Dual Union All" & vbNewLine & _
                "Select null,null,null,null,null,null,null,null,'--------------------------------------------------------------------------------' From Dual Union All" & vbNewLine & _
                "Select * From (Select /*+ no_merge */ ID,LPad(' ', Level - 1) || Operation || ' ' || Options As Operation, Object_Name As Name, Cardinality, Bytes,Cost, Time,Object_Owner," & vbNewLine & _
                "        Rpad('| ' || Substr(Lpad(' ', 1 * (Level - 1)) || Operation || Decode(Options, Null, '', ' ' || Options), 1, 52),52, ' ') || '|' || Rpad(Substr(Object_Name || ' ', 1, 30), 30, ' ') || '|' ||" & vbNewLine & _
                "         Lpad(Decode(Cardinality, Null, ' ',Decode(Sign(Cardinality - 1000), -1, Cardinality || ' '," & vbNewLine & _
                "                             Decode(Sign(Cardinality - 1000000), -1, Trunc(Cardinality / 1000) || 'K',Decode(Sign(Cardinality - 1000000000), -1, Trunc(Cardinality / 1000000) || 'M'," & vbNewLine & _
                "                                             Trunc(Cardinality / 1000000000) || 'G')))), 7, ' ') || '|' ||" & vbNewLine & _
                "         Lpad(Decode(Bytes, Null, ' ',Decode(Sign(Bytes - 1024), -1, Bytes || ' '," & vbNewLine & _
                "                             Decode(Sign(Bytes - 1048576), -1, Trunc(Bytes / 1024) || 'K',Decode(Sign(Bytes - 1073741824), -1, Trunc(Bytes / 1048576) || 'M'," & vbNewLine & _
                "                                             Trunc(Bytes / 1073741824) || 'G')))), 6, ' ') || '|' ||" & vbNewLine & _
                "         Lpad(Decode(Cost, Null, ' ',Decode(Sign(Cost - 10000000), -1, Cost || ' '," & vbNewLine & _
                "                             Decode(Sign(Cost - 1000000000), -1, Trunc(Cost / 1000000) || 'M', Trunc(Cost / 1000000000) || 'G'))),8, ' ') || '|' ||" & vbNewLine & _
                "         Lpad(Decode(Partition_Start, 'ROW LOCATION', 'ROWID',Decode(Partition_Start, 'KEY', 'KEY'," & vbNewLine & _
                "                             Decode(Partition_Start, 'KEY(INLIST)', 'KEY(I)',Decode(Substr(Partition_Start, 1, 6), 'NUMBER'," & vbNewLine & _
                "                                             Substr(Substr(Partition_Start, 8, 10), 1, Length(Substr(Partition_Start, 8, 10)) - 1)," & vbNewLine & _
                "                                             Decode(Partition_Start, Null, ' ', Partition_Start))))) || ' ', 7, ' ') || '|' ||" & vbNewLine & _
                "         Lpad(Decode(Partition_Stop, 'ROW LOCATION', 'ROW L',Decode(Partition_Stop, 'KEY', 'KEY'," & vbNewLine & _
                "                             Decode(Partition_Stop, 'KEY(INLIST)', 'KEY(I)',Decode(Substr(Partition_Stop, 1, 6), 'NUMBER'," & vbNewLine & _
                "                                             Substr(Substr(Partition_Stop, 8, 10), 1, Length(Substr(Partition_Stop, 8, 10)) - 1)," & vbNewLine & _
                "                                             Decode(Partition_Stop, Null, ' ', Partition_Stop))))) || ' ', 7, ' ') || '|' As ""Explain plan""" & vbNewLine & _
                "       From (Select * From v$sql_plan Where Sql_Id = [1] and child_number = [2]) a " & vbNewLine & _
                "       Start With a.Id = 0  Connect By Prior a.Id = a.Parent_Id  Order By Id, Position)" & vbNewLine & _
                "Union All Select null,null,null,null,null,null,null,null,'--------------------------------------------------------------------------------' From Dual"
    '这里v$sql_plan必须要用一个子查询，否则会很慢
    '外面要用Select *，否则会报ID字段不存在
    Set GetSQLPlan = zlDatabase.OpenSQLRecord(strSQL, "GetSQLPlan", strSQLID, lngPrev_Num)

    Exit Function
errh:
    If ErrCenter() = 1 Then Resume
End Function

Private Function InitrsIndex(ByRef rsIndex As Recordset, ByRef rsCons_FK As Recordset) As Boolean
'功能：性能问题检查所需变量
    Set rsIndex = New ADODB.Recordset
    rsIndex.Fields.Append "Table_Name", adVarChar, 50, adFldIsNullable
    rsIndex.Fields.Append "Index_Name", adVarChar, 50, adFldIsNullable
    rsIndex.Fields.Append "Column_Name", adVarChar, 100, adFldIsNullable
    rsIndex.Fields.Append "Index_Owner", adVarChar, 50, adFldIsNullable
    rsIndex.CursorLocation = adUseClient
    rsIndex.LockType = adLockOptimistic
    rsIndex.CursorType = adOpenStatic
    rsIndex.Open
    
    Set rsCons_FK = New ADODB.Recordset
    rsCons_FK.Fields.Append "r_Constraint_Name", adVarChar, 50, adFldIsNullable
    rsCons_FK.Fields.Append "Constraint_Name", adVarChar, 50, adFldIsNullable
    rsCons_FK.Fields.Append "Column_Name", adVarChar, 100, adFldIsNullable
    rsCons_FK.Fields.Append "Table_Name", adVarChar, 50, adFldIsNullable
    rsCons_FK.Fields.Append "Index_Owner", adVarChar, 50, adFldIsNullable
    rsCons_FK.CursorLocation = adUseClient
    rsCons_FK.LockType = adLockOptimistic
    rsCons_FK.CursorType = adOpenStatic
    rsCons_FK.Open
    
End Function

Private Function CheckSQLPlan(ByRef strPlanEx As String) As Boolean
'功能：检查性能问题，并返回存在性能问题的SQL执行计划:
'         1.大表全表扫描zlbigtable+zlbaktables，
'         2.中型表全表扫描(如果有统计信息，User_tab_statistics:num_rows>3000(药品目录一般是这个值以上) AND num_rows<100 0000百万以内)
'         3.大表上引用基础表(非大表)的外键上的索引
'         4.大表和中型表索引全扫描（inex full scan，INDEX FAST FULL SCAN）
'         5.大表和中型表跳跃式索引扫描（INDEX SKIP SCAN）
'返回：blnReturn=true 有性能问题
    Dim rsPlan As ADODB.Recordset
    Dim rsCons_FK As Recordset
    Dim rsTemp As Recordset
    
    Dim i As Long, strSQL As String
    Dim blnReturn As Boolean
    Dim strTable As String
    Dim strTmp As String
    Dim strAllTable As String

    gblnSys = True
    Set rsPlan = GetSQLPlan()
    If Not rsPlan Is Nothing Then
        If mstrBigTable = "" Then
            '先取大表（zlbigtable+zlbaktables）
            strSQL = "Select Distinct 表名" & vbNewLine & _
                    "From Zlbigtables" & vbNewLine & _
                    "Union All" & vbNewLine & _
                    "Select Distinct 表名 From Zlbaktables"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName)
            Do While Not rsTemp.EOF
                mstrBigTable = mstrBigTable & "," & rsTemp!表名
                rsTemp.MoveNext
            Loop
            mstrBigTable = mstrBigTable & ","
        End If
        '再取中表（统计信息，User_tab_statistics:num_rows>3000）
        If mstrMiddlePar = "" Then
            mstrMiddlePar = GetSetting("ZLSOFT", "公共模块\zl9Report\SQLPlanEx", "MiddleTable", "1")
            If Val(mstrMiddlePar) = 1 Then
                strTmp = GetSetting("ZLSOFT", "公共模块\zl9Report\SQLPlanEx", "MiddleTableRows", "3000|1000000")

                strSQL = "Select Table_Name as 表名 From DBA_Tab_Statistics Where Num_Rows > [1] And Num_Rows < [2]" & _
                        " and owner In(Select 'ZLTOOLS' From Dual Union All Select 所有者 from zlsystems)"
                Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, Val(Split(strTmp, "|")(0)), Val(Split(strTmp, "|")(1)))
                mstrMiddleTable = ""
                Do While Not rsTemp.EOF
                    If InStr("," & mstrBigTable & ",", "," & rsTemp!表名 & ",") = 0 Then
                        mstrMiddleTable = mstrMiddleTable & "," & rsTemp!表名
                    End If
                    rsTemp.MoveNext
                Loop
                mstrMiddleTable = mstrMiddleTable & ","
            Else
                mstrMiddleTable = ""
            End If
        End If
        
        strAllTable = mstrMiddleTable & mstrBigTable
        
        For i = 1 To rsPlan.RecordCount
            strPlanEx = strPlanEx & vbCrLf & rsPlan!PlanTable
            If blnReturn = False Then
                strTmp = UCase(rsPlan!Operation & "")
                If InStr(strTmp, "TABLE ACCESS FULL") > 0 Then
                    '判断是否是大表中表全扫描
                    blnReturn = InStr(strAllTable, "," & rsPlan!Name & ",") > 0
                    
                ElseIf InStr(strTmp, "INDEX FAST FULL SCAN") > 0 Or InStr(strTmp, "INDEX FULL SCAN") > 0 Or InStr(strTmp, "INDEX SKIP SCAN") > 0 Then
                    '判断是否是大表中表索引全扫描或跳跃式索引
                    strTable = Split(rsPlan!Name & "_", "_")(0)
                    blnReturn = InStr(strAllTable, "," & strTable & ",") > 0
                 
                ElseIf InStr(strTmp, "INDEX RANGE SCAN") > 0 Then
                    '大表上使用了基础表(非大表)外键索引
                    strTable = Split(rsPlan!Name & "_", "_")(0)
                    
                    If InStr(mstrBigTable, "," & strTable & ",") > 0 Then
                        If InStr("病案主页_IX_婴儿科室ID,病案主页_IX_婴儿病区ID", rsPlan!Name) = 0 Then '排除几个特殊索引
                             If mrsIndex Is Nothing Then Call InitrsIndex(mrsIndex, mrsCons_FK)
                        
                             mrsIndex.Filter = "Index_Name='" & rsPlan!Name & "' And Index_Owner='" & rsPlan!object_owner & "'"
                             If mrsIndex.RecordCount = 0 Then
                                 strSQL = "Select d.Table_Name, d.Index_Name, d.Column_Name,d.Column_Position" & vbNewLine & _
                                 "              From DBA_Ind_Columns D" & vbNewLine & _
                                 "              Where d.Index_Name = [1]  And d.Index_Owner = [2] " & vbNewLine & _
                                 "              Order By d.Column_Position"
                                 Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, rsPlan!Name, rsPlan!object_owner)
                                 If rsTemp.RecordCount > 0 Then
                                     mrsIndex.AddNew
                                     mrsIndex!Index_Owner = rsPlan!object_owner
                                     mrsIndex!Table_Name = rsTemp!Table_Name
                                     mrsIndex!Index_Name = rsPlan!Name
                                     strTmp = ""
                                     Do While Not rsTemp.EOF
                                         strTmp = strTmp & "," & rsTemp!Column_Name
                                         rsTemp.MoveNext
                                     Loop
                                     mrsIndex!Column_Name = Mid(strTmp, 2)
                                     mrsIndex.Update
                                 End If
                                 mrsIndex.Filter = "Index_Name='" & rsPlan!Name & "' And Index_Owner='" & rsPlan!object_owner & "'"
                             End If
                             If mrsIndex.RecordCount > 0 Then
                                 mrsCons_FK.Filter = "Table_Name='" & strTable & "' And Index_Owner='" & rsPlan!object_owner & "'"
                                 If mrsCons_FK.RecordCount = 0 Then
                                     '找外键约束
                                     Set rsCons_FK = GetConsFK(strTable, rsPlan!object_owner)
                                     Do While Not rsCons_FK.EOF
                                         mrsCons_FK.AddNew
                                         mrsCons_FK!r_Constraint_Name = rsCons_FK!r_Constraint_Name
                                         mrsCons_FK!Constraint_Name = rsCons_FK!Constraint_Name
                                         mrsCons_FK!Column_Name = rsCons_FK!Column_Name
                                         mrsCons_FK!Index_Owner = rsPlan!object_owner
                                         mrsCons_FK!Table_Name = strTable
                                         mrsCons_FK.Update
                                         rsCons_FK.MoveNext
                                     Loop
                                 End If
                                 mrsCons_FK.Filter = "Table_Name='" & strTable & "' And Index_Owner='" & rsPlan!object_owner & "' And Column_Name='" & mrsIndex!Column_Name & "'"
                                 If mrsCons_FK.RecordCount > 0 Then
                                     strTable = Split(mrsCons_FK!r_Constraint_Name & "_", "_")(0)
                                     '外键父表不是大表，则视为有性能问题
                                     blnReturn = InStr(mstrBigTable, "," & strTable & ",") = 0
                                 End If
                             End If
                        End If
                    End If
                End If
            End If
            
            rsPlan.MoveNext
        Next
    End If
    If blnReturn = False Then strPlanEx = ""
    gblnSys = False
    CheckSQLPlan = blnReturn
End Function

Private Function GetConsFK(ByVal strTable As String, ByVal strOwner As String) As ADODB.Recordset
'功能：获取指定表的外键约束记录集
'参数：strFind=表名
    Dim strSQL As String
    Dim rsCons As Recordset
    Dim rsCons_FK As Recordset

    strSQL = "Select f.Constraint_Name, f.r_Constraint_Name, e.Column_Name, e.Position" & vbNewLine & _
            "From Dba_Cons_Columns E, Dba_Constraints F" & vbNewLine & _
            "Where e.Constraint_Name = f.Constraint_Name And e.owner = f.owner  And f.Constraint_Type = 'R' And f.Table_Name = [1] And f.owner = [2]" & vbNewLine & _
            "Order By f.Constraint_Name, e.Position"
    Set rsCons = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, strTable, strOwner)
    
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
            rsCons_FK!Constraint_Name = rsCons!Constraint_Name
            rsCons_FK!r_Constraint_Name = rsCons!r_Constraint_Name
            rsCons_FK!Column_Name = rsCons!Column_Name
        Else
            rsCons_FK!Column_Name = rsCons_FK!Column_Name & "," & rsCons!Column_Name
        End If
        rsCons_FK.Update
        rsCons.MoveNext
    Loop
    Set GetConsFK = rsCons_FK
End Function

Public Function SaveWinState(objForm As Object, Optional ByVal strProjectName As String, Optional ByVal strUserDef As String) As Boolean
'功能：保存窗体及其中各种控件的状态
'参数：objForm:要保存的窗体
'      strProjectName：当前工程名，通常可用app.ProductName传递，用以区分不同工程中的同名窗体，保证恢复的正确性；
'      strUserDef：主要适用于工程中，一个窗体多个程序使用(程序使用 set frmxxx=new frm设计窗体形式)，为了按不同应用保存恢复各自的个性化状态，需要直接确定命名。
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Integer
    Dim strTmp As String, strIndex As String
    Dim objThis As Object, blnAutoHide As Boolean
    Dim lngIndex As Long
    
    On Error GoTo errh
    
    If Not gcnOracle Is Nothing And strProjectName <> "" And UCase(strProjectName) <> "ZL9COMMON" And gblnRunLog Then
        If gcnOracle.State = 1 Then
            '正常退出
            On Error Resume Next
            strSQL = "Zl_Zldiarylog_Update(" & "'" & gstrComputerName & "','" & UCase(strProjectName) & _
                    "','" & UCase(objForm.Name) & "',1," & glngSessionID & ")"
            Call zlDatabase.ExecuteProcedure(strSQL, "更新工作日志")
            If Err.Number <> 0 Then Err.Clear
        End If
    End If
    
    On Error Resume Next
    Call gfrmMain.Shut任务(objForm)
    On Error GoTo 0

    If Val(zlDatabase.GetPara("使用个性化风格")) = 0 Then
        Call DelWinState(objForm, strProjectName, strUserDef)
        SaveWinState = True: Exit Function
    End If
    blnAutoHide = Val(zlDatabase.GetPara("界面区域隐藏")) = 1
    
    If strProjectName <> "" Then strProjectName = strProjectName & "\"
    
    '保存窗体状态、位置、大小
    With objForm
        Select Case .WindowState
            Case 0
                SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strProjectName & objForm.Name & strUserDef & "\Form", "状态", objForm.WindowState & "," & .Left & "," & .Top & "," & .Width & "," & .Height
            Case 1
                SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strProjectName & objForm.Name & strUserDef & "\Form", "状态", 0
            Case 2
                '获取当前监视器索引
                lngIndex = GetMonitorIndex(objForm.hWnd)
                
                If UBound(gMonitors) > 1 Then
                    '说明有多个屏幕，注册表状态参数最后增加保存的屏幕索引
                    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strProjectName & objForm.Name & strUserDef & "\Form", "状态", objForm.WindowState & ",,,,," & CStr(lngIndex)
                Else
                    '单个屏幕，维持原来的处理方式
                    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strProjectName & objForm.Name & strUserDef & "\Form", "状态", objForm.WindowState
                End If
                
        End Select
    End With
    
    '保存各种控件的各种状态
    For Each objThis In objForm.Controls
        strTmp = "": strIndex = ""
        On Error Resume Next
        If UCase(TypeName(objThis)) = UCase("Menu") Then
            If objThis.Caption Like "标准按钮*" Or _
                objThis.Caption Like "文本标签*" Or _
                objThis.Caption Like "状态栏*" Or _
                UCase(objThis.Name) Like UCase("mnuViewTool*") Then
                '特殊菜单的复选
                strTmp = objThis.Checked & "," & objThis.Enabled
            Else
                strTmp = ""
            End If
            
            strIndex = objThis.Index
            If Err.Number <> 0 Then Err.Clear: strIndex = ""
        ElseIf (UCase(objThis.Name) Like "*_S" Or _
            UCase(TypeName(objThis)) = UCase("StatusBar") Or _
            UCase(TypeName(objThis)) = UCase("Toolbar") Or _
            UCase(TypeName(objThis)) = UCase("Coolbar")) And objForm.Visible Then

            'Left,Top,Width、Height,Visible
            strTmp = strTmp & "," & objThis.Left
            If Err.Number <> 0 Then Err.Clear: strTmp = strTmp & ",-32767"
            
            strTmp = strTmp & "," & objThis.Top
            If Err.Number <> 0 Then Err.Clear: strTmp = strTmp & ",-32767"
            
            strTmp = strTmp & "," & objThis.Width
            If Err.Number <> 0 Then Err.Clear: strTmp = strTmp & ",-32767"
            
            strTmp = strTmp & "," & objThis.Height
            If Err.Number <> 0 Then Err.Clear: strTmp = strTmp & ",-32767"
            
            strTmp = strTmp & "," & objThis.Visible
            If Err.Number <> 0 Then Err.Clear: strTmp = strTmp & ",-32767"
            
            strIndex = objThis.Index
            If Err.Number <> 0 Then Err.Clear: strIndex = ""
            
            strTmp = Mid(strTmp, 2)
        End If
        If strTmp <> "" Then
            SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strProjectName & objForm.Name & strUserDef & "\" & TypeName(objThis), objThis.Name & strIndex & "状态", strTmp
        End If
        
        Select Case UCase(TypeName(objThis))
            Case UCase("CommandBars") 'CommandBar
                If objThis.ActiveMenuBar.Visible Then '有菜单栏的才处理
                    '状态栏
                    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strProjectName & objForm.Name & strUserDef & "\" & TypeName(objThis), objThis.Name & "StatusBarVisible", IIf(objThis.StatusBar.Visible, 1, 0)
                    '标准按钮,文本标签
                    If objThis.count >= 2 Then
                        SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strProjectName & objForm.Name & strUserDef & "\" & TypeName(objThis), objThis.Name & "Visible", IIf(objThis(2).Visible, 1, 0)
                        SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strProjectName & objForm.Name & strUserDef & "\" & TypeName(objThis), objThis.Name & "ButtonText", objThis(2).Controls(1).Style
                    End If
                    '大图标
                    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strProjectName & objForm.Name & strUserDef & "\" & TypeName(objThis), objThis.Name & "LargeIcon", IIf(objThis.Options.LargeIcons, 1, 0)
                End If
            Case UCase("DockingPane") 'DockingPane
                'SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strProjectName & objForm.Name & strUserDef & "\" & TypeName(objThis), objThis.Name, objThis.SaveStateToString
                If Not blnAutoHide Then
                    DeleteSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strProjectName & objForm.Name & strUserDef & "\" & TypeName(objThis)
                    DeleteSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & Replace(strProjectName, "\", "") & objForm.Name & strUserDef & "\" & TypeName(objThis)
                End If
            Case UCase("ReportControl") 'ReportControl
                SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strProjectName & objForm.Name & strUserDef & "\" & TypeName(objThis), objThis.Name, objThis.SaveSettings
            Case UCase("Toolbar")
                If objThis.Buttons.count > 0 Then
                    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strProjectName & objForm.Name & strUserDef & "\" & TypeName(objThis), objThis.Name & "文本", IIf(objThis.Buttons(1).Caption <> "", 1, objThis.ButtonHeight)
                End If
            Case UCase("ListView")
                strIndex = objThis.Index
                If Err.Number <> 0 Then Err.Clear: strIndex = ""
                SaveListViewState objThis, strProjectName & objForm.Name & strUserDef, strIndex
            Case UCase("CoolBar")
                strTmp = ""
                For i = 1 To objThis.Bands.count
                    strTmp = strTmp & "," & objThis.Bands(i).NewRow
                Next
                SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strProjectName & objForm.Name & strUserDef & "\" & TypeName(objThis), objThis.Name & "行序", Mid(strTmp, 2)
                
                strTmp = ""
                For i = 1 To objThis.Bands.count
                    strTmp = strTmp & "," & objThis.Bands(i).Visible
                Next
                SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strProjectName & objForm.Name & strUserDef & "\" & TypeName(objThis), objThis.Name & "可见栏", Mid(strTmp, 2)
            Case UCase("MSHFlexGrid"), UCase("BillEdit"), UCase("VSFlexGrid")
                SaveFlexState objThis, strProjectName & objForm.Name & strUserDef
            Case UCase("DataGrid")
                SaveDBGridState objThis, strProjectName & objForm.Name & strUserDef
        End Select
    Next
    SaveWinState = True
    Exit Function
errh:
    If gobjComLib.ErrCenter() = 1 Then Resume
End Function

Public Function RestoreWinState(objForm As Object, Optional ByVal strProjectName As String, Optional ByVal strUserDef As String, Optional ByVal blnAutoNotHide As Boolean) As Boolean
'功能：恢复窗体的状态，当左顶边界超出时，则自动设置为0
'参数：objForm:要恢复的窗体
'      strProjectName：当前工程名，通常可用app.ProductName传递，用以区分不同工程中的同名窗体，保证恢复的正确性；
'      strUserDef：主要适用于工程中，一个窗体多个程序使用(程序使用 set frmxxx=new frm设计窗体形式)，为了按不同应用保存恢复各自的个性化状态，需要直接确定命名。
'      blnAutoNotHide:强制界面显示区域
    Dim aryInfo() As String
    Dim strTmp As String, i As Integer
    Dim objThis As Object, objSub As Object
    Dim strIndex As String, blnDo As Boolean
    Dim strSQL As String, strOEM As String
    Dim strUesrInfo As String, strCaption As String
    Dim rsUser As ADODB.Recordset
    Dim lngIndex As Long
    Dim lngtwipPerPixersX As Long
    Dim lngtwipPerPixersY As Long
    
    Dim blnAutoHide As Boolean
    
    '窗体标题添加操作员信息
    strTmp = objForm.Caption
    strCaption = strTmp
    If Not gcnOracle Is Nothing Then
        If gcnOracle.State = 1 And strProjectName <> "" Then
            Set rsUser = zlDatabase.GetUserInfo
            If rsUser.EOF = False Then
                strUesrInfo = "" & rsUser!姓名
                If strTmp Like "*-" & strUesrInfo Then
                    strTmp = Mid(strTmp, 1, Len(strTmp) - Len("-" & strUesrInfo))
                ElseIf strTmp Like "*- " & strUesrInfo Then
                    strTmp = Mid(strTmp, 1, Len(strTmp) - Len("- " & strUesrInfo))
                    If Right(strTmp, 1) = "" Then strTmp = Mid(strTmp, 1, Len(strTmp) - 1)
                ElseIf strTmp Like strUesrInfo & "-*" Then
                    strTmp = Mid(strTmp, Len(strUesrInfo & "-") + 1)
                ElseIf strTmp Like strUesrInfo & " -*" Then
                    strTmp = Mid(strTmp, Len(strUesrInfo & " -") + 1)
                    If Left(strTmp, 1) = "" Then strTmp = Mid(strTmp, 2)
                ElseIf strTmp Like "*" & "(操作员：" & strUesrInfo & ")" & "*" Then
                    strTmp = Replace(strTmp, "(操作员：" & strUesrInfo & ")", "")
                ElseIf strTmp Like "*" & "(当前用户：" & strUesrInfo & ")" & "*" Then
                    strTmp = Replace(strTmp, "(当前用户：" & strUesrInfo & ")", "")
                End If
                objForm.Caption = strTmp
                strCaption = strTmp & "(当前用户：" & strUesrInfo & ")" & IIf(g_NodeName = "-", "", " (院区：" & g_NodeName & ")")
            End If
        End If
    End If
    If Not gcnOracle Is Nothing And strProjectName <> "" And UCase(strProjectName) <> "ZL9COMMON" And gblnRunLog Then
        If gcnOracle.State = 1 Then
            On Error Resume Next
            strSQL = "Zl_Zldiarylog_Insert('" & gstrComputerName & "','" & UCase(strProjectName) & "'," & _
                " '" & UCase(objForm.Name) & "','" & UCase(objForm.Caption) & "'," & glngSessionID & ")"
            Call zlDatabase.ExecuteProcedure(strSQL, "保存工作日志")
            If Err.Number <> 0 Then Err.Clear
        End If
    End If
    
    On Error Resume Next
    
    blnDo = Val(zlDatabase.GetPara("使用个性化风格")) <> 0
    blnAutoHide = Val(zlDatabase.GetPara("界面区域隐藏")) = 1
    If blnAutoHide = False And blnAutoNotHide Then blnAutoHide = True
    
    If strProjectName <> "" Then strProjectName = strProjectName & "\"
    
    '恢复窗体的状态、位置、大小
    strTmp = "0," & (Screen.Width - objForm.Width) / 2 & "," & (Screen.Height - objForm.Height) / 2 & "," & objForm.Width & "," & objForm.Height
    If blnDo Then
        aryInfo = Split(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strProjectName & objForm.Name & strUserDef & "\Form", "状态", strTmp), ",")
    Else
        aryInfo = Split(strTmp, ",")
    End If
    With objForm

        .WindowState = aryInfo(0)
        .Caption = strCaption
        If UBound(aryInfo) = 4 Then
            .Left = IIf(aryInfo(1) < 0, 0, aryInfo(1))
            .Top = IIf(aryInfo(2) < 0, 0, aryInfo(2))
            .Width = IIf(aryInfo(3) > Screen.Width, Screen.Width, aryInfo(3))
            .Height = IIf(aryInfo(4) > Screen.Height, Screen.Height, aryInfo(4))
        ElseIf UBound(aryInfo) = 5 Then
            '这种情况说明多屏环境并且最大化状态
            
            '保证gMonitors有值
            If UBound(gMonitors()) < 2 Then
                Call GetMonitorIndex(objForm.hWnd)
            End If
            lngtwipPerPixersX = Screen.TwipsPerPixelX
            lngtwipPerPixersY = Screen.TwipsPerPixelY
            
            .Left = lngtwipPerPixersX * gMonitors(Val(aryInfo(5) + 1)).monitorInf.rcMonitor.Left + _
            ((lngtwipPerPixersX * Abs(gMonitors(Val(aryInfo(5) + 1)).monitorInf.rcMonitor.Right - gMonitors(Val(aryInfo(5) + 1)).monitorInf.rcMonitor.Left) - objForm.Width) / 2)
            
            .Top = (lngtwipPerPixersY * Abs(gMonitors(Val(aryInfo(5) + 1)).monitorInf.rcMonitor.Bottom - gMonitors(Val(aryInfo(5) + 1)).monitorInf.rcMonitor.Top) - objForm.Height) / 2
        Else
            .Left = (Screen.Width - objForm.Width) / 2
            .Top = (Screen.Height - objForm.Height) / 2
        End If
    End With

    '恢复窗体中各种控件的各种状态
    For Each objThis In objForm.Controls
        If blnDo Then
            strTmp = "": strIndex = ""
            If UCase(TypeName(objThis)) = UCase("Menu") Then
                '特殊菜单的复选
                If objThis.Caption Like "标准按钮*" Or _
                    objThis.Caption Like "文本标签*" Or _
                    objThis.Caption Like "状态栏*" Or _
                    UCase(objThis.Name) Like UCase("mnuViewTool*") Then
                    
                    strIndex = objThis.Index
                    If Err.Number <> 0 Then Err.Clear: strIndex = ""
                    
                    strTmp = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strProjectName & objForm.Name & strUserDef & "\" & TypeName(objThis), objThis.Name & strIndex & "状态", "")
                    If UBound(Split(strTmp, ",")) = 1 Then
                        objThis.Checked = Split(strTmp, ",")(0)
                        objThis.Enabled = Split(strTmp, ",")(1)
                    End If
                End If
            ElseIf UCase(objThis.Name) Like "*_S" Or _
                UCase(TypeName(objThis)) = UCase("StatusBar") Or _
                UCase(TypeName(objThis)) = UCase("Toolbar") Or _
                UCase(TypeName(objThis)) = UCase("Coolbar") Then
                
                strIndex = objThis.Index
                If Err.Number <> 0 Then Err.Clear: strIndex = ""
                
                strTmp = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strProjectName & objForm.Name & strUserDef & "\" & TypeName(objThis), objThis.Name & strIndex & "状态", "")
                If strTmp <> "" Then
                    'Left,Top,Width、Height,Visible
                    If UBound(Split(strTmp, ",")) = 4 Then
                        If Split(strTmp, ",")(0) <> "-32767" Then objThis.Left = Split(strTmp, ",")(0)
                        If Split(strTmp, ",")(1) <> "-32767" Then objThis.Top = Split(strTmp, ",")(1)
                        If Split(strTmp, ",")(2) <> "-32767" Then objThis.Width = Split(strTmp, ",")(2)
                        If Split(strTmp, ",")(3) <> "-32767" Then objThis.Height = Split(strTmp, ",")(3)
                        If Split(strTmp, ",")(4) <> "-32767" Then objThis.Visible = Split(strTmp, ",")(4)
                    End If
                End If
            End If
        End If
        
        Select Case UCase(TypeName(objThis))
            Case UCase("CommandBars") 'CommandBar
                If blnDo Then
                    If objThis.ActiveMenuBar.Visible Then '有菜单栏的才处理
                        '状态栏
                        strTmp = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strProjectName & objForm.Name & strUserDef & "\" & TypeName(objThis), objThis.Name & "StatusBarVisible", "")
                        If strTmp <> "" Then objThis.StatusBar.Visible = Val(strTmp) <> 0
                        '标准按钮,文本标签
                        If objThis.count >= 2 Then
                            strTmp = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strProjectName & objForm.Name & strUserDef & "\" & TypeName(objThis), objThis.Name & "Visible", "")
                            If strTmp <> "" Then
                                For i = 2 To objThis.count
                                    objThis(i).Visible = Val(strTmp) <> 0
                                Next
                            End If
                            strTmp = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strProjectName & objForm.Name & strUserDef & "\" & TypeName(objThis), objThis.Name & "ButtonText", "")
                            If strTmp <> "" Then
                                For i = 2 To objThis.count
                                    For Each objSub In objThis(i).Controls
                                        objSub.Style = Val(strTmp)
                                    Next
                                Next
                            End If
                        End If
                        '大图标
                        strTmp = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strProjectName & objForm.Name & strUserDef & "\" & TypeName(objThis), objThis.Name & "LargeIcon", "")
                        If strTmp <> "" Then objThis.Options.LargeIcons = Val(strTmp) <> 0
                    End If
                End If
            Case UCase("DockingPane") 'DockingPane
                If blnDo Then
                    'strTmp = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strProjectName & objForm.Name & strUserDef & "\" & TypeName(objThis), objThis.Name, "")
                    'If strTmp <> "" Then objThis.LoadStateFromString strTmp
                End If
                If Not blnAutoHide Then
                    'PaneNoHideable(2) Or PaneNoCloseable(1) or PaneNoFloatable(4)，没有Hide会自动出现Close,允许Float双击有问题
                    For i = 1 To objThis.PanesCount
                        'PaneNoCaption(8)
                        If objThis.Panes(i).Options And 8 = 8 Then
                            objThis.Panes(i).Options = 15
                        Else
                            objThis.Panes(i).Options = 7
                        End If
                    Next
                    DeleteSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strProjectName & objForm.Name & strUserDef & "\" & TypeName(objThis)
                    DeleteSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & Replace(strProjectName, "\", "") & objForm.Name & strUserDef & "\" & TypeName(objThis)
                End If
            Case UCase("ReportControl") 'ReportControl
                If blnDo Then
                    strTmp = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strProjectName & objForm.Name & strUserDef & "\" & TypeName(objThis), objThis.Name, "")
                    If strTmp <> "" Then objThis.LoadSettings strTmp
                End If
            Case UCase("StatusBar")
                '状态条试用标志
'                If zlRegInfo("授权性质") <> "1" Then
'                    If objThis.Panels(1).Bevel = sbrRaised Then
'                        objThis.Panels(1).Text = ""
'                        Set objThis.Panels(1).Picture = LoadCustomPicture("Try")
'                        objThis.Panels(1).ToolTipText = ""
'                        objThis.Height = 360
'                    End If
'                Else
                    If objThis.Panels(1).Bevel = sbrRaised Then
                        strTmp = zlRegInfo("产品简名")
                        If strTmp <> "-" Then
                            objThis.Panels(1).Text = strTmp & "软件"
                            If gobjLogFile Is Nothing Then Set gobjLogFile = New FileSystemObject
                            If gstrAppsoft = "" Then
                                gstrAppsoft = gobjLogFile.GetParentFolderName(App.Path)
                                If OS.IsDesinMode Then
                                    gstrAppsoft = "C:\APPSOFT"
                                End If
                            End If
                            
                            If gobjLogFile.FileExists(gstrAppsoft & "\附加文件\logo_app.jpg") Then
                                 Set objThis.Panels(1).Picture = LoadPicture(gstrAppsoft & "\附加文件\logo_app.jpg")
                            Else
                                '处理状态栏图标的OEM策略
                                If strTmp = "中联" Then
                                    If zlRegInfo("授权性质") <> "1" Then
                                        Set objThis.Panels(1).Picture = LoadCustomPicture("Try")
                                    Else
                                        Set objThis.Panels(1).Picture = LoadCustomPicture("Logo")
                                    End If
                                Else
                                    strOEM = mGetOEM(strTmp)
                                    Set objThis.Panels(1).Picture = LoadCustomPicture(strOEM)
                                    If Err <> 0 Then
                                        Err.Clear
                                        Set objThis.Panels(1).Picture = LoadCustomPicture("Logo")
                                    End If
                                End If
                            End If
 
                            If zlRegInfo("授权性质") <> "1" Then
                                If strTmp = "中联" Then
                                    objThis.Panels(1).Text = ""
                                Else
                                    objThis.Panels(1).Text = strTmp & "(试用)"
                                End If
                            End If
                            objThis.Panels(1).ToolTipText = ""
                            objThis.Height = 360
                        End If
                    End If
'                End If
            Case UCase("Menu")
                If UCase(objThis.Name) = UCase("mnuHelpWeb") Then
                    'WEB上的中联
                    strTmp = zlRegInfo("支持商简名")
                    If strTmp <> "-" Then
                        objThis.Caption = "&WEB上的" & strTmp
                    End If
                ElseIf UCase(objThis.Name) = UCase("mnuHelpWebHome") Then
                    '中联主页
                    strTmp = zlRegInfo("支持商简名")
                    If strTmp <> "-" Then
                        objThis.Caption = strTmp & "主页(&H)"
                    End If
                ElseIf UCase(objThis.Name) = UCase("mnuHelpWebForum") Then
                    '中联论坛
                    strTmp = zlRegInfo("支持商简名")
                    If strTmp <> "-" Then
                        objThis.Caption = strTmp & "论坛(&F)"
                    End If
                    If strTmp <> "中联" Then objThis.Visible = False
                End If
            Case UCase("Toolbar")
                If blnDo Then
                    If objThis.Buttons.count > 0 Then
                        strTmp = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strProjectName & objForm.Name & strUserDef & "\" & TypeName(objThis), objThis.Name & "文本", 1)
                        For i = 1 To objThis.Buttons.count
                            objThis.Buttons(i).Caption = IIf(strTmp = 1, objThis.Buttons(i).Tag, "")
                        Next
                    End If
                End If
            Case UCase("ListView")
                If blnDo Then
                    strIndex = objThis.Index
                    If Err.Number <> 0 Then Err.Clear: strIndex = ""
                    RestoreListViewState objThis, strProjectName & objForm.Name & strUserDef, strIndex
                End If
            Case UCase("CoolBar")
                If blnDo Then
                    strTmp = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strProjectName & objForm.Name & strUserDef & "\" & TypeName(objThis), objThis.Name & "行序", "")
                    If UBound(Split(strTmp, ",")) >= 0 Then
                        For i = 0 To UBound(Split(strTmp, ","))
                            objThis.Bands(i + 1).NewRow = Split(strTmp, ",")(i)
                        Next
                    End If
            
                    strTmp = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strProjectName & objForm.Name & strUserDef & "\" & TypeName(objThis), objThis.Name & "可见栏", "")
                    If UBound(Split(strTmp, ",")) >= 0 Then
                        For i = 0 To UBound(Split(strTmp, ","))
                            objThis.Bands(i + 1).Visible = Split(strTmp, ",")(i)
                        Next
                    End If
                End If
            Case UCase("MSHFlexGrid"), UCase("BillEdit"), UCase("VSFlexGrid")
                If blnDo Then
                    RestoreFlexState objThis, strProjectName & objForm.Name & strUserDef
                End If
            Case UCase("DataGrid")
                If blnDo Then
                    RestoreDBGridState objThis, strProjectName & objForm.Name & strUserDef
                End If
        End Select
    Next
    Call gfrmMain.Show任务(objForm)
    RestoreWinState = True
End Function

Public Sub RestoreDBGridState(objThis As Object, strForm As String)
    Dim strText As String, arrText As Variant
    Dim strWidth As String, i As Integer
    
    On Error Resume Next
    
    If Val(zlDatabase.GetPara("使用个性化风格")) = 0 Then Exit Sub
    
    '检查是否需要恢复
    strText = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strForm & "\" & TypeName(objThis), objThis.Name & objThis.Tag & "名称", "")
    If strText <> "" Then
        arrText = Split(strText, ",")
        
        '列数变了,不恢复而使用缺省
        If UBound(arrText) + 1 <> objThis.Columns.count Then Exit Sub
        
        '列标题变了,不恢复而使用缺省
        For i = 0 To objThis.Columns.count - 1
            If objThis.Columns(i).Caption <> arrText(i) Then Exit Sub
        Next
    End If
    
    strWidth = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strForm & "\" & TypeName(objThis), objThis.Name & objThis.Tag & "宽度", "")
    If UBound(Split(strWidth, ",")) >= objThis.Columns.count - 1 Then
        For i = 0 To objThis.Columns.count - 1
            objThis.Columns(i).Width = Split(strWidth, ",")(i)
        Next
    End If
End Sub

Public Function RestoreFlexState(objThis As Object, strForm As String) As Boolean
    Dim strWidth As String, strText As String
    Dim arrText As Variant, i As Integer
        
    On Error Resume Next
    
    If Val(zlDatabase.GetPara("使用个性化风格")) = 0 Then
        RestoreFlexState = True: Exit Function
    End If
    
    '检查是否需要恢复
    strText = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strForm & "\" & TypeName(objThis), objThis.Name & objThis.Tag & "名称", "")
    If strText <> "" Then
        '固定行变了,不恢复而使用缺省
        If UCase(TypeName(objThis)) = UCase("BillEdit") Then
            If objThis.msfObj.FixedRows <> 1 Then Exit Function
        Else
            If objThis.FixedRows <> 1 Then Exit Function
        End If
        
        '列数变了,不恢复而使用缺省
        arrText = Split(strText, ",")
        If objThis.Cols <> UBound(arrText) + 1 Then Exit Function
        
        '列标题变了,不恢复而使用缺省
        For i = 0 To objThis.Cols - 1
            If objThis.TextMatrix(0, i) <> arrText(i) Then Exit Function
        Next
    End If
        
    strWidth = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strForm & "\" & TypeName(objThis), objThis.Name & objThis.Tag & "宽度", "")
    If UBound(Split(strWidth, ",")) >= objThis.Cols - 1 Then
        For i = 0 To objThis.Cols - 1
            objThis.ColWidth(i) = Split(strWidth, ",")(i)
        Next
        RestoreFlexState = True
    End If
    
    If TypeName(objThis) = "VSFlexGrid" Then
        strText = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strForm & "\" & TypeName(objThis), objThis.Name & objThis.Tag & "冻结", "")
        If strText <> "" And IsNumeric(strText) Then objThis.FrozenCols = Val(strText)
    End If
End Function

Public Sub SaveFlexState(objThis As Object, strForm As String)
    Dim strWidth As String, strText As String, i As Integer
        
    On Error Resume Next
    
    strWidth = "": strText = ""
    For i = 0 To objThis.Cols - 1
        strWidth = strWidth & "," & objThis.ColWidth(i)
        If UCase(TypeName(objThis)) = UCase("BillEdit") Then
            If objThis.msfObj.FixedRows = 1 Then strText = strText & "," & objThis.TextMatrix(0, i)
        Else
            If objThis.FixedRows = 1 Then strText = strText & "," & objThis.TextMatrix(0, i)
        End If
    Next
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strForm & "\" & TypeName(objThis), objThis.Name & objThis.Tag & "宽度", Mid(strWidth, 2)
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strForm & "\" & TypeName(objThis), objThis.Name & objThis.Tag & "名称", Mid(strText, 2)
    
    If TypeName(objThis) = "VSFlexGrid" Then
        SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strForm & "\" & TypeName(objThis), objThis.Name & objThis.Tag & "冻结", objThis.FrozenCols
    End If
End Sub

Public Sub SaveDBGridState(objThis As Object, strForm As String)
    Dim strWidth As String, strText As String, i As Integer
    
    On Error Resume Next
    
    strWidth = "": strText = ""
    For i = 0 To objThis.Columns.count - 1
        strWidth = strWidth & "," & objThis.Columns(i).Width
        strText = strText & "," & objThis.Columns(i).Caption
    Next
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strForm & "\" & TypeName(objThis), objThis.Name & objThis.Tag & "宽度", Mid(strWidth, 2)
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strForm & "\" & TypeName(objThis), objThis.Name & objThis.Tag & "名称", Mid(strText, 2)
End Sub

Public Sub SaveListViewState(objLvw As Object, ByVal strForm As String, ByVal strIndex As String)
'功能：保存ListView的各种特性
'参数：objLvw=ListView对象,strForm=窗体关键字
'说明：视图方式、列宽、列位置、列标题、列对齐、排序
    Dim lngCol As Long
    Dim strWidth As String
    Dim strPosition As String
    Dim strText As String
    Dim strAlign As String
    
    For lngCol = 1 To objLvw.ColumnHeaders.count
        strWidth = strWidth & "," & objLvw.ColumnHeaders(lngCol).Width
        strPosition = strPosition & "," & objLvw.ColumnHeaders(lngCol).Position
        strText = strText & "," & objLvw.ColumnHeaders(lngCol).Text
        strAlign = strAlign & "," & objLvw.ColumnHeaders(lngCol).Alignment
    Next
    
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strForm & "\ListView", objLvw.Name & strIndex & "视图", objLvw.View
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strForm & "\ListView", objLvw.Name & strIndex & "宽度", Mid(strWidth, 2)
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strForm & "\ListView", objLvw.Name & strIndex & "位置", Mid(strPosition, 2)
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strForm & "\ListView", objLvw.Name & strIndex & "名称", Mid(strText, 2)
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strForm & "\ListView", objLvw.Name & strIndex & "对齐", Mid(strAlign, 2)
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strForm & "\ListView", objLvw.Name & strIndex & "排序", objLvw.SortKey & "," & objLvw.SortOrder & "," & objLvw.Sorted
End Sub

Public Sub RestoreListViewState(objLvw As Object, ByVal strForm As String, ByVal strIndex As String)
'功能：恢复ListView的各种特性
'参数：objLvw=ListView对象,strForm=窗体关键字
'说明：视图方式、列宽、列位置、列标题、列对齐、排序
    Dim strWidth As String, strPosition As String
    Dim strText As String, arrText As Variant
    Dim strAlign As String, strSort As String
    Dim lngCol As Long
    
    On Error Resume Next
    
    If Val(zlDatabase.GetPara("使用个性化风格")) = 0 Then Exit Sub
    
    strText = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strForm & "\ListView", objLvw.Name & strIndex & "名称")
    If strText <> "" Then
        arrText = Split(strText, ",")
        If objLvw.Tag = "可变化的" Then
            '只对需要的ListView进行列变化
            objLvw.ColumnHeaders.Clear
            For lngCol = LBound(arrText) To UBound(arrText)
                '列缺省关键字为"_" & 列标题
                objLvw.ColumnHeaders.Add , "_" & arrText(lngCol), arrText(lngCol)
            Next
        Else
            '检查是否需要恢复
            '列数变了,不恢复而使用缺省
            If UBound(arrText) + 1 <> objLvw.ColumnHeaders.count Then Exit Sub
            '列标题变了,不恢复而使用缺省
            For lngCol = 1 To objLvw.ColumnHeaders.count
                If objLvw.ColumnHeaders(lngCol).Text <> arrText(lngCol - 1) Then Exit Sub
            Next
        End If
    End If
    
    '视图缺省保持初始值
    lngCol = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strForm & "\ListView", objLvw.Name & strIndex & "视图", -1)
    If lngCol <> -1 Then objLvw.View = lngCol
    
    '列的宽度,顺序,对齐
    strWidth = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strForm & "\ListView", objLvw.Name & strIndex & "宽度")
    strPosition = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strForm & "\ListView", objLvw.Name & strIndex & "位置")
    strAlign = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strForm & "\ListView", objLvw.Name & strIndex & "对齐")
    For lngCol = 1 To objLvw.ColumnHeaders.count
        '列缺省关键字为"_" & 列标题
        objLvw.ColumnHeaders(lngCol).Key = "_" & objLvw.ColumnHeaders(lngCol).Text
        If strWidth <> "" Then objLvw.ColumnHeaders(lngCol).Width = Split(strWidth, ",")(lngCol - 1)
        If strAlign <> "" Then objLvw.ColumnHeaders(lngCol).Alignment = Split(strAlign, ",")(lngCol - 1)
    Next
    For lngCol = objLvw.ColumnHeaders.count To 1 Step -1
        If strPosition <> "" Then objLvw.ColumnHeaders(lngCol).Position = Split(strPosition, ",")(lngCol - 1)
    Next
    
    '排序特性
    strSort = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strForm & "\ListView", objLvw.Name & strIndex & "排序")
    If strSort <> "" Then
        objLvw.SortKey = Split(strSort, ",")(0)
        objLvw.SortOrder = Split(strSort, ",")(1)
        objLvw.Sorted = Split(strSort, ",")(2)
    End If
End Sub

Public Function DelWinState(objForm As Object, Optional ByVal strProjectName As String, Optional ByVal strUserDef As String) As Boolean
'功能：删除窗体个性化设置值
'参数：objForm:要恢复的窗体
'      strProjectName：当前工程名，通常可用app.ProductName传递，用以区分不同工程中的同名窗体，保证恢复的正确性；
'      strUserDef：主要适用于工程中，一个窗体多个程序使用(程序使用 set frmxxx=new frm设计窗体形式)，为了按不同应用保存恢复各自的个性化状态，需要直接确定命名。
    Dim strProject As String
    Dim lngR As Long
    Dim objThis As Object
    
    strProject = strProjectName
    If strProjectName <> "" Then strProjectName = strProjectName & "\"
    
    For Each objThis In objForm.Controls
        lngR = RegDeleteKey(HKEY_CURRENT_USER, "Software\VB and VBA Program Settings\ZLSOFT\私有模块\" & gstrDBUser & "\界面设置\" & strProjectName & objForm.Name & strUserDef & "\" & TypeName(objThis) & Chr(0))
        If lngR <> 0 And lngR <> 2 Then Exit Function
    Next
    
    lngR = RegDeleteKey(HKEY_CURRENT_USER, "Software\VB and VBA Program Settings\ZLSOFT\私有模块\" & gstrDBUser & "\界面设置\" & strProjectName & objForm.Name & strUserDef & "\Form" & Chr(0))
    If lngR <> 0 And lngR <> 2 Then Exit Function
    lngR = RegDeleteKey(HKEY_CURRENT_USER, "Software\VB and VBA Program Settings\ZLSOFT\私有模块\" & gstrDBUser & "\界面设置\" & strProjectName & objForm.Name & strUserDef & Chr(0))
    If lngR <> 0 And lngR <> 2 Then Exit Function
    
    DelWinState = True
End Function

Public Function GetUnitInfo(Optional strItem As String) As Variant
'功能：读取用户注册信息表内的自定义项目的内容(标准的注册项目需通过zlRegist部件读取)
'参数：strItem=要读取的项目
'返回：读取单项目返回字符串,读取整个信息返回记录集
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    On Error GoTo errh
    
    strSQL = "Select 项目,行号,内容 From zlRegInfo" & IIf(strItem = "", " Order by 项目,行号", " Where 项目='" & strItem & "' Order by 行号")
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "读取用户注册信息")
    
    If strItem <> "" Then
        If rsTmp.EOF Then
            GetUnitInfo = ""
        Else
            Do While Not rsTmp.EOF
                GetUnitInfo = GetUnitInfo & IIf(IsNull(rsTmp!内容), "", rsTmp!内容)
                rsTmp.MoveNext
            Loop
        End If
    Else
        If Not rsTmp.EOF Then
            Set GetUnitInfo = rsTmp
        Else
            Set GetUnitInfo = Nothing
        End If
    End If
    Exit Function
errh:
    If gobjComLib.ErrCenter() = 1 Then Resume
    If strItem <> "" Then
        GetUnitInfo = ""
    Else
        Set GetUnitInfo = Nothing
    End If
End Function

'========与用户信息相关函数=============================================================================================
Public Function GetUnitName() As String
'功能：返回当前注册用户单位名称(该函数已无用)
    GetUnitName = zlRegInfo("单位名称", , -1)
End Function

Public Function GetDBConnection(Optional ByVal Index As Integer = 0) As ADODB.Connection
'功能：通过数据连接编号获取对应的数据连接对象
'参数：
'  Index：数据连接编号
    Dim strPass As String, strServer As String, strSQL As String
    Dim cnOracle As ADODB.Connection
    Dim rsTemp As ADODB.Recordset

    If Index <= 0 Then
        Set GetDBConnection = gcnOracle
    Else
        On Error GoTo hErr
        '数据获取
        strSQL = _
                "Select 编号, 用户名, 密码, Ip, 端口, 实例名 " & vbNewLine & _
                "From ZlConnections " & vbNewLine & _
                "Order By 编号 "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取数据连接信息")
        '加载数据连接
        Set cnOracle = New ADODB.Connection
        With rsTemp
            Do While Not .EOF
                If NVL(!编号, 0) = Index Then
                    '初始化连接对象
                    strServer = NVL(!IP) & _
                                IIf(NVL(!端口) = "", ":1521", ":" & NVL(!端口)) & _
                                IIf(NVL(!实例名) = "", "", "/" & NVL(!实例名))
                    strPass = NVL(!密码)
                    '解密
                    strPass = Decipher(MSTR_DBLINK_KEY, strPass)
                    Set cnOracle = gobjRegister.GetConnection(strServer, NVL(!用户名), strPass, False, OraOLEDB, , False)
                    Exit Do
                End If
                .MoveNext
            Loop
        End With
        Set GetDBConnection = cnOracle
    End If
    Exit Function
hErr:
    Call ErrCenter
End Function


Public Function zlHomePage(hWnd As Long) As Boolean
'功能：根据产品发行码，联结主页
    Dim strCode As String
    
    strCode = zlRegInfo("支持商URL")
    If strCode <> "-" Then
        ShellExecute hWnd, "open", "http://" & strCode, "", "", 1
        zlHomePage = True
    End If
End Function

Public Function zlWebForum(hWnd As Long) As Boolean
'功能：根据产品发行码，联结论坛
    Dim strCode As String
    
    'strCode = zlRegInfo("支持商BBS")
    strCode = "www.zlsoft.com/techbbs/index.asp"
    If strCode <> "-" Then
        ShellExecute hWnd, "open", "http://" & strCode, "", "", 1
        zlWebForum = True
    End If
End Function

Public Function zlMailTo(hWnd As Long) As Boolean
'功能：根据产品发行码发送电子邮件
    Dim strCode As String
    strCode = zlRegInfo("支持商MAIL")
    If strCode <> "-" Then
        ShellExecute hWnd, "open", "mailto:" & strCode, "", "", 1
        zlMailTo = True
    End If
End Function

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
    strR = Trim(Left(strFile, InStr(strFile, Chr(0)) - 1)) & CLng(timer * 100) & ".pic"

    Open strR For Binary As intFile
    Put intFile, , arrData()
    Close intFile
    Set LoadCustomPicture = VB.LoadPicture(strR)
    Kill strR
End Function


Public Sub SetMainForm(frmMain As Object)
'功能：设置导航台主窗体
    Set gfrmMain = frmMain
End Sub

Public Sub SetDbUser(ByVal strDBUser As String)
'功能：设置当前连接用户名(为兼容旧程序的调用而存在，从10.35.10开始，已在InitCommon中设置gstrDBUser)
'说明：该方法是兼容以前的，只供Zl9Comlib调用，其他部件对象不能调用
    gstrDBUser = UCase(strDBUser)
    Set grsUserInfo = Nothing
End Sub

Public Sub SetNodeNo(ByVal strNodeNo As String)
'功能：设置站点编号
    g_NodeNo = strNodeNo
    gstrNodeNo = g_NodeNo
End Sub

Public Sub SetNodeName(ByVal strNodeName As String)
'功能：设置站点名称
    g_NodeName = strNodeName
    gstrNodeName = g_NodeName
End Sub

Public Function GetCurStates() As Integer
'功能：返回当前已连接站点数

'注销时需使用的函数
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    On Error GoTo errh
    
    strSQL = "Select Distinct Machine" & vbNewLine & _
            "From Gv$session" & vbNewLine & _
            "Where Upper(Program) Like 'ZL%.EXE' And Status <> 'KILLED' And Exists" & vbNewLine & _
            " (Select 1 From Gv$session Where Audsid = Userenv('SessionID'))"
            
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "已连接站点数")
    
    If Not rsTmp.EOF Then GetCurStates = rsTmp.RecordCount
    Exit Function
errh:
    If ErrCenter() = 1 Then Resume
End Function

Public Function ShowHelp(ByVal ChmName As String, SHwnd As Long, ByVal htmName As String, Optional Sys As Integer = 1) As Boolean
    '-----------------------------------------------------------------------------------------------------------------------------
    '功能:显示帮助窗体
    '参数:ChmName:CHM格式文件(目前传入的是:App.ProductName)
    '     SHwnd:传入窗口句柄(作为宿主窗口)
    '     htmName:射映在CHM中的htm文件名称
    '返回:
    '修改人:刘兴宏
    '日期:2007/09/05
    '-----------------------------------------------------------------------------------------------------------------------------

    Dim strPath As String
    Dim strFile As String
    
    Dim strSave As String
    On Error GoTo ShowHelpErr
    
    ShowHelp = False
    strSave = String(200, Chr$(0))
    strPath = Left$(strSave, GetWindowsDirectory(strSave, Len(strSave))) + "\help\"

    If Sys = 0 Then
        '公共导航台调用
        strFile = strPath & "zl9Start" & ".chm"
        If Trim(Dir(strFile)) = "" Then GoTo ShowHelpErr
        Call Htmlhelp(SHwnd, strFile, &H0, htmName & ".htm")
    Else
        Select Case UCase(ChmName)
        Case "ZL9LISWORK", "ZL9LISCOMM", "ZLLISDEV"   'LIS系统
            strFile = strPath & "ZLLIS.chm"
        Case "ZL9PACSWORK", "ZL9PACSCORE", "ZLPACSSRV"            'PACS系统     :ZLPACSSRV－网管接收程序
            strFile = strPath & "zlPacs.chm"
        Case "ZL9OPS"                                 '手麻系统
            strFile = strPath & "Zl9Ops.Chm"
        Case "ZL9PEIS"                                '体检系统
            strFile = strPath & "zl9Peis.Chm"
        Case "ZL9BLOOD"                               '血库
            strFile = strPath & "zl9Blood.chm"
        Case "ZL9TRANSFUSION"                         '门诊输液
            strFile = strPath & "ZL9Transfunction.chm"
        Case "ZL9OUTEXSE", "ZL9PATIENT", "ZL9REGEVENT", "ZL9MEDICAL", "ZL9MEDISTORE", _
             "ZL9ANALYSIS", "ZL9STUFF", "ZL9INSURE", "ZL9CISWORK", "ZL9CISJOB", _
             "ZL9CUSTACC", "ZL9BASEITEM", "ZL9CASHBILL", _
             "ZL9CISBASE", "ZL9CISCORE", "ZLCISKERNEL", "ZLPICTUREEDITOR", "ZLRICHEPR", _
             "ZL9INEXSE", "ZL9INPATIENT", "ZL9DUE", "ZL9DRUGSTORE", _
             "ZL9IMGCAPTURE", "ZL9NEWQUERY", "ZL9TECHCORE", _
             "ZL9LEDVOICE", "ZL9TAXBILL", "ZLIDCARD", "ZL9CISAUDIT"
            '门诊费用部件;病人管理部件;挂号事务部件;体检管理部件;药库管理部件;
            '查询分析部件;卫生材料部件;医保结算系统;临床工作部件;临床工作部件;
            '专项记帐单  ;基础编码部件;财务监控部件;
            '临床基础部件;临床核心部件;临床核心部件;临床核心部件;电子病历部件
            '住院费用部件;住院管理部件;应付款系统;药房管理部件;
            '影像采集部件;病人导诊咨询;医嘱管理部件;
            '语音报价部件;税控设置部件;电子病案审查部件
            strFile = strPath & "zl9App1.chm"
        Case "ZL9BASECODE", "ZL9CLINIQUE", "ZL9CUREBASE", "ZL9CUREREC", "ZL9DRADVICE", "ZL9NURSE", _
             "ZL9MEDICODE", "ZL9PUBQUERY", "ZL9ORDERS"
            '9系列
            '基础编码部件;门诊诊治部件;诊疗基础部件;住院病历部件;住院医嘱部件;住院护理部件
            '药品编码部件;公共查询部件;医嘱处理部件;
            '与HIS10版本共享的部件:zl9Analysis(查询分析部件);zl9CashBill,zl9DrugStore;zl9InExse;zl9Inpatient
            '   zl9CustAcc,Zl9DrugStore,zl9MediStore,zl9NewQuery,zl9OutExse,zl9Patient,zl9RegEvent,
            
            strFile = strPath & "zl9App1.chm"
        Case "ZL9MEDREC"                '病案事务部件    300
            strFile = strPath & "zl9App3.chm"
        Case "ZL9MATERIAL"              '物资事务部件    400
            strFile = strPath & "zl9App4.chm"
        Case "ZL9ACCOUNT", "ZL9FINANCEFUNC", "ZL9EXCEL"             '帐务管理部件    500
            strFile = strPath & "zl9App5.chm"
        Case "ZL9DEVICE"                '设备管理系统    600
            strFile = strPath & "zl9App6.chm"
        Case "ZL9COST"                  '成本核算系统    700
            strFile = strPath & "zl9App7.chm"
        Case "ZL9HUMAN"                 '人事系统        200
            strFile = strPath & "zl9App2.chm"
        Case Else '暂无帮助
            MsgBox "部件<<" & ChmName & ">>暂无帮助，请与开发商联系!", vbInformation + vbDefaultButton1, gstrSysName
            Exit Function
        End Select
        If Trim(Dir(strFile)) = "" Then
            'MsgBox "未找到帮助文件(" & strFile & ")，请检查!", vbInformation + vbDefaultButton1, gstrSysName
            GoTo ShowHelpErr
        End If
        
        Call Htmlhelp(SHwnd, strFile, &H0, ChmName & "\" & htmName & ".htm")
    End If
    ShowHelp = True
    Exit Function
ShowHelpErr:
    Err.Clear
End Function

Public Function RegCheck() As Boolean
'说明：该方法是兼容以前的，只供Zl9Comlib调用，其他部件对象不能调用
    Dim strRegErr As String
    Dim intCount As Integer
    
    '--进行注册相关检查--
    RegCheck = False
    
    '注册检查
    strRegErr = zlRegCheck
    If strRegErr <> "" Then
        MsgBox strRegErr & ",系统不能继续运行！", vbInformation, gstrSysName
        Exit Function
    End If

'    用户站点数检测 (0表示不限制)
    intCount = Val(zlRegInfo("授权站点"))
    If intCount > 0 Then
        If GetCurStates > intCount Then
            MsgBox "当前用户登录数超过了最大授权数,系统将自动结束运行！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
                          
    RegCheck = True
End Function

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
        
        strOEM = mGetOEM(strProductName, blnCorp)
        If str属性 = "Picture" Then
            Set objPicture.Picture = LoadCustomPicture(strOEM)
        ElseIf str属性 = "Icon" Then
            Set objPicture.Icon = LoadCustomPicture(strOEM)
        End If
        
        If Err <> 0 Then
            Err.Clear
        End If
    
    End If
End Sub

Private Function mGetOEM(ByVal strAsk As String, Optional ByVal blnCorp As Boolean = True) As String
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
    mGetOEM = strCode
End Function

Private Function mGetInvalidTable() As String
'功能：得到在最近使用的SQL语句中不能访问的表或视图
    Dim varTables As Variant
    Dim strTable As String, lngCount As Long
    Dim strInvalidTable As String
    
    varTables = Split(SQLObject(gstrRecentSQL), ",")
    
    On Error Resume Next
    
    For lngCount = 0 To UBound(varTables)
        strTable = varTables(lngCount)
        
        '测试该对象是否可用
        gcnOracle.Execute "select 1 from " & strTable & " where rownum<1"
        If Err <> 0 Then
            Err.Clear
            strInvalidTable = strInvalidTable & "," & strTable
        End If
    Next
    
    If strInvalidTable <> "" Then
        '去掉第一个逗号
        mGetInvalidTable = Mid(strInvalidTable, 2)
    End If
End Function

'--------------------------------------------------
'功能：验证系统注册授权的正确性
'参数：blnTemp-是否从未保存的临时注册信息验证
'返回：正确返回"";错误返回错误信息
'--------------------------------------------------
Public Function zlRegCheck(Optional blnTemp As Boolean) As String
'说明：该方法是兼容以前的，只供Zl9Comlib调用，其他部件对象不能调用
    zlRegCheck = gobjRegister.zlRegCheck(blnTemp)
End Function

'--------------------------------------------------
'功能：获得指定的产品发行或注册授权信息
'参数： strItem-指定的授权项目
'       blnTemp-是否从未保存的临时注册信息验证
'       intBits-对于同时有多项信息的单位名称、产品开发商等指定获得第几个信息,0-N,为-1时表示返回";"间隔的多个
'返回：正确时返回指定的信息；错误返回""
'--------------------------------------------------
Public Function zlRegInfo(strItem As String, Optional blnTemp As Boolean, Optional intBits As Integer) As String
    zlRegInfo = gobjRegister.zlRegInfo(strItem, blnTemp, intBits)
End Function

'--------------------------------------------------
'功能：获得授权工具信息
'返回：按2的工具末位次方返回工具许可
'--------------------------------------------------
Public Function zlRegTool(Optional blnTemp As Boolean) As Long
    zlRegTool = gobjRegister.zlRegTool(blnTemp)
End Function


Public Function GetPrivFunc(lngSys As Long, lngProgId As Long) As String
'功能：返回当前用户具有的指定程序的功能串
'参数：lngSys     如果是固定模块，则为0
'      lngProgId  程序序号
'返回：分号间隔的功能串,为空表示没有权限
    GetPrivFunc = gobjRegister.zlRegFunc(lngSys, lngProgId)
End Function

Public Function GetPrivFuncByUser(lngSys As Long, lngProgId As Long, Optional ByVal strDBUser As String) As String
'功能：返回其他用户具有的指定程序的功能串,传登录用户
'参数：lngSys     如果是固定模块，则为0
'      lngProgId  程序序号
'      strDBUser=数据库用户
'返回：分号间隔的功能串,为空表示没有权限
    GetPrivFuncByUser = Sys.GetPrivFuncByUser(lngSys, lngProgId, strDBUser)
End Function

Public Function ConnectMip(ByVal lngParentForm As Long) As Boolean
    '******************************************************************************************************************
    '功能：连接消息服务平台
    '参数：无
    '说明：成功返回True,否则返回False
    '******************************************************************************************************************
    Dim strSQL As String
    If gclsMipClient Is Nothing Then Set gclsMipClient = New clsMipClient
    
    ConnectMip = gclsMipClient.ConnectMip(lngParentForm)
    
End Function

Public Function AddMipModule(ByRef objMipModule As clsMipModule) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：成功返回True,否则返回False
    '******************************************************************************************************************
    If Not (gclsMipClient Is Nothing) Then AddMipModule = gclsMipClient.AddModule(objMipModule)
End Function

Public Function DelMipModule(ByRef objMipModule As clsMipModule) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：成功返回True,否则返回False
    '******************************************************************************************************************
    If Not (gclsMipClient Is Nothing) Then
        DelMipModule = gclsMipClient.DeleteModule(objMipModule)
    End If
End Function

Public Function DisConnectMip() As Boolean
    '******************************************************************************************************************
    '功能：断开消息服务平台
    '参数：无
    '说明：成功返回True,否则返回False
    '******************************************************************************************************************
    If Not (gclsMipClient Is Nothing) Then
        DisConnectMip = gclsMipClient.DisConnectMip
    End If
End Function

'常用函数，放在ComLib，可以不加前缀
Public Function NVL(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'clsCommFun存在该函数
'功能：相当于Oracle的NVL，将Null值改成另外一个预设值
    NVL = IIf(IsNull(varValue), DefaultValue, varValue)
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

Public Function Between(X, A, b) As Boolean
'功能：判断x是否在a和b之间
    If A < b Then
        Between = X >= A And X <= b
    Else
        Between = X >= b And X <= A
    End If
End Function

Public Function IntEx(vNumber As Variant) As Variant
'功能：取大于指定数值的最小整数
    IntEx = -1 * Int(-1 * Val(vNumber))
End Function

Public Function ZVal(ByVal varValue As Variant, Optional ByVal blnForceNum As Boolean, Optional ByVal varDefault As Variant = 0) As String
'功能：将0零转换为"NULL"串,在生成SQL语句时用
'参数：blnForceNum=当为Null时，是否强制表示为数字型
    ZVal = zlStr.ZVal(varValue, blnForceNum, varDefault)
End Function

Public Function RoundEx(ByVal dblNumber As Double, ByVal intBit As Integer) As Double
'功能：四舍五入方式格式化数字
'参数：intBit=最大小数位数
'问题号：94552
'说明：VB自带的Round是银行家舍入法,与实际不一致。如Round(57.575,2)=57.58,Round(57.565,2)=57.56
    RoundEx = zlStr.RoundEx(dblNumber, intBit)
End Function

Public Function FormatEx(ByVal vNumber As Variant, ByVal intBit As Integer, _
    Optional blnShowZero As Boolean = True, Optional ByVal blnAddZero As Boolean, _
    Optional intMinBit As Integer) As String
'功能：四舍五入方式格式化显示数字,保证小数点最后不出现0,小数点前要有0
'参数：vNumber=Single,Double,Currency类型的数字,
'          intBit=最大小数位数
'         blnShowZero=vNumber为0时是否显示0值
'         blnAddZero=小数位不足是否补零
'       intMinBit=最少保留小数位数,如:FormatEx(1,6,True,False,2) 结果为：1.00
'返回：格式化后的字符串
    FormatEx = zlStr.FormatEx(vNumber, intBit, blnShowZero, blnAddZero, intMinBit)
End Function

Public Function SystemVersion(Optional ByVal lngSys As Long = 100) As String
'功能：提取指定系统版本号，默认为医院系统标准版
    SystemVersion = Sys.SystemVersion(lngSys)
End Function

Public Function SystemOwner(Optional ByVal lngSys As Long = 100) As String
'功能：获取指定系统的所有者
    SystemOwner = Sys.SystemOwner(lngSys)
End Function

Public Function SystemShareWith(ByVal lngSysShareNO As Long, Optional ByVal lngSysMainNO As Long = 100) As Boolean
'功能：标准系统和其他系统是否是共享安装
'参数：lngSysShareNO= 共享安装的系统
'           lngSysMainNO=主系统
    SystemShareWith = Sys.SystemShareWith(lngSysShareNO, lngSysMainNO)
End Function

Public Function IsSysSetUp(ByVal lngSys As Long) As Boolean
'功能：是否安装某个系统
    IsSysSetUp = Sys.IsSysSetUp(lngSys)
End Function

Public Function CurrentConn() As ADODB.Connection
    '******************************************************************************************************************
    '功能：返回当前连接
    '参数：无
    '说明：返回当前连接
    '******************************************************************************************************************
    Set CurrentConn = gcnOracle
End Function

Friend Sub ChangeIntanceConn(cnMain As ADODB.Connection)
'功能：同步当前实例中的mcnoracle进行事务重试按钮的禁用。
'参数：同步当前实例的连接
'说明：该方法只供zl9comlib自身使用，其他部件不能使用。
    If Not mobjOracle Is cnMain Then
        Set mobjOracle = cnMain
    End If
End Sub
