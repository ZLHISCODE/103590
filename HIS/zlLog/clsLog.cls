VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum LogLevel
    lglUnDefined = -1                   '尚未设置，应用于模块部件级别
    lglLogOFF = 0                       '不记录日志
    lglError = 1                        '该日志是属于业务错误或者VB错误、程序错误等，会影响程序运行的
    lglWarn = 2                         '该日志不属于错误，不影响程序运行，但是可能造成流程变动或者程序功能不全。可能与数据控制或者当前环境相关，如缺失某部件，仍然可以继续使用，但是对应功能缺失。
    lglInfo = 3                         '该日志属于重要信息记录，用于重要信息的记录，如费用，交易等数据。
    lglTrace = 4                        '该信息是程序的运行跟踪信息。用于跟踪程序运行，以便方便错误查证。
    lglAll = 5                          '记录所有日志信息
End Enum

Public Enum LogMode
    lgmNone = 0                         '不记录日志
    lgmOnlyLocal = 1                    '仅记录本地日志
    lgmOnlyDB = 2                       '仅记录数据库日志
    lgmBoth = 3                         '本地日志和数据库日志都记录
End Enum

Public Enum LogCallState
    lcsCallBegin = 0
    lcsCallEnd = 1
End Enum

Private Enum InfoIndex
    ifiLogLevel = 0
    ifiLogMode = 1
    ifiComponent = 0
    ifiFunction = 1
    ifiModule = 2
    ifiCall = 3
End Enum

Private Const MLNG_MAXLINE  As Long = 2 ^ 28            '本地日志文档最大行数

Private mcnDbBusiness       As ADODB.Connection         '业务数据库（ZLHIS）
Private mcnDbLog            As ADODB.Connection         '数据库日志连接
Private mobjLogFile         As FileSystemObject
Private mobjLocalLog        As TextStream               '本地日志对象
Private mstrStation         As String                   '工作站即电脑名
Private mlngPID             As Long                     '进程ID
Private mstrProcessName     As String                   '进程名

Private mstrLogLevelName(6) As String                   '日志级别名称
Private mblnUseLog          As Boolean                  '是否需要日志输出。True需要；False不需要
'Private mstrParas           As String
Private mstrLogTiltle       As String
Private mstrCurLogName      As String
Private mstrCurFuncs        As String

'存储日志数据库的连接信息，防止自动断开，用于自动重连
Private marrCurLogSet       As Variant                  '当前日志配置信息； Index=0，日志等级；Index=1，日志保存方式；Index=2，部件名；Index=3，模块；Index=4，功能；Index=5，调用用户名； 后面动态扩展（部件、模块、功能、用户）
Private mdicLogControl      As Scripting.Dictionary     '日志配置信息（列表对象）
'  日志类别
'    日志名  Array
'      LogLevel     日志等级
'      LogMode      日志方式
'      部件1        Scripting.Dictionary
'      模块1        Scripting.Dictionary
'      功能1        Scripting.Dictionary
'      CallName1    Scripting.Dictionary
'      ...
'      部件n        Scripting.Dictionary
'      模块n        Scripting.Dictionary
'      功能n        Scripting.Dictionary
'      CallNamen    Scripting.Dictionary

Private mstrLogFile                     As String
Private mstrLogDB                       As String       '日志数据库的服务名
Private mstrLogUser                     As String
Private mstrOccurServer                 As String
Private mstrOccurUser                   As String
Private mlngOccurSID                    As Long
Private mstrOccurIP                     As String
Private mlngKeepDays                    As Long
Private mblnSilentMode                  As Boolean

'CurrentLogLevel当前日志级别(R/W)
Private Property Get CurrentLogLevel() As LogLevel
    On Error Resume Next
    CurrentLogLevel = -1
    CurrentLogLevel = marrCurLogSet(ifiLogLevel)
End Property
'CurrentLogMode当前的日志模式。0-不记录日志。1-仅记录本地日志。2-仅记录数据库日志，3-本地日志和数据库日志都记录(R)
Private Property Get CurrentLogMode() As LogMode
    On Error Resume Next
    CurrentLogMode = -1
    CurrentLogMode = marrCurLogSet(ifiLogMode)
End Property
''LogFile日志文件文件名称(R)
'Public Property Get LogFile() As String
'    LogFile = mstrLogFile
'End Property
''LogDB 日志存储的数据库(R)
'Public Property Get LogDB() As String
'    LogDB = mstrLogDB
'End Property
''LogUser 写日志的用户(R)
'Public Property Get LogUser() As String
'    LogUser = mstrLogUser
'End Property
''OccurServer 日志对应业务发生的数据库(R)
'Public Property Get OccurServer() As String
'    OccurServer = mstrOccurServer
'End Property
''OccurUser 日志记录内容的操作用户(R)
'Public Property Get OccurUser() As String
'    OccurUser = mstrOccurUser
'End Property
''OccurSID 日志记录内容的业务操作的会话SID(R)
'Public Property Get OccurSID() As Long
'    OccurSID = mstrOccurUser
'End Property
''OccurIP 日志记录内容的发生的IP，可能存在多网卡，因此通过数据库判定(R)
'Public Property Get OccurIP() As String
'    OccurIP = mstrOccurIP
'End Property
''CurrentLogCount当前已经记录的行数(R)
'Public Property Get CurrentLogCount() As Long
'    CurrentLogCount = mlngCurLogCount
'End Property

'静默提示对话框。注意：本部件的弹出对话框能不弹出尽量不弹出，必要弹出的需要通过该属性判断
Public Property Get SilentMode() As Boolean
    SilentMode = mblnSilentMode
End Property
Public Property Let SilentMode(Value As Boolean)
    mblnSilentMode = Value
End Property

Public Function SetBusinessDB(ByVal cnMain As ADODB.Connection) As Boolean
'功能：设置日志的位置(必要)。即：业务发生的数据库信息
'参数：
'  cnMain：业务的连接信息
'返回：是否连接成功
    
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    '检查参数
    If mcnDbBusiness Is cnMain Then
        SetBusinessDB = True
        Exit Function
    End If
    
    Set mcnDbBusiness = cnMain
    If mcnDbBusiness Is Nothing Then
        MsgBoxEx "SetBusinessDB", "业务数据库连接未设置！", vbInformation, App.Title
        Exit Function
    ElseIf mcnDbBusiness.State <> adStateOpen Then
        Set mcnDbBusiness = Nothing
        MsgBoxEx "SetBusinessDB", "业务数据库连接未开启！", vbInformation, App.Title
        Exit Function
    End If
    
    mlngOccurSID = 0
    mstrOccurIP = ""
    mstrOccurUser = ""
    mstrOccurServer = ""
        
    '当前会话无需访问Gv$
    On Error GoTo ErrH
    strSql = _
        "Select User, Audsid, Sys_Context('USERENV', 'IP_ADDRESS') As Ip" & vbNewLine & _
        "From v$session" & vbNewLine & _
        "Where Audsid = Userenv('SessionID')"
    Set rsTmp = mdlDataBase.OpenSQLRecord(mcnDbBusiness, strSql, "查询终端信息")
    If rsTmp.EOF = False Then
        mlngOccurSID = rsTmp!Audsid
        mstrOccurIP = rsTmp!IP
        mstrOccurUser = rsTmp!User
        mstrOccurServer = mdlPublic.GetServerInfo(mcnDbBusiness)
    End If
    rsTmp.Close
    
    SetBusinessDB = ConnectLogDB()         '设置日志存储的数据库
    Exit Function
    
ErrH:
    Err.Clear
End Function

Private Function ConnectLogDB() As Boolean
'功能：日志存储的数据库
'参数：无
'返回：是否连接成功
'备注：
'ZlRegInfo表.项目[日志服务器]格式：
'    a. 指定日志存储服务器以及日志记录用户。zlStr.Sm4EncryptEcb("SERVER=127.0.0.1:1521/TestBase USER=ZLHIS PASS=AQA TRANS=1")
'    b. 仅指定日志存储服务器。用户使用ZLUA。zlStr.Sm4EncryptEcb("SERVER=127.0.0.1:1521/TestBase")
'    c. 仅指定用户。日志存储当前数据库。zlStr.Sm4EncryptEcb("USER=ZLHIS PASS=AQA TRANS=1")
'    d. 都未指定。日志数据库存储当前数据库，且用户为ZLUA。zlStr.Sm4EncryptEcb("")

    Const G_UA_PWD As String = "FA74C8A530DE7E088B1ACA673DD6297D"
    Const G_UA_KEY As String = "0016FDE250354FA9A4BA45433DBCC35D"
    
    Dim cnTmp As ADODB.Connection
    Dim rsTmp As ADODB.Recordset
    Dim strCommand As String, strCurServer As String, strSql As String
    Dim blnOpen As Boolean
    
    mblnUseLog = False
    mstrCurLogName = ""
    mstrCurFuncs = ""
    mstrLogDB = ""
    mstrLogUser = ""
    mstrLogFile = ""
    Call ClearObject(mobjLocalLog)
    
    If mcnDbBusiness Is Nothing Then
        MsgBoxEx "ConnectLogDB", "请先设置业务数据库（SetBusinessDB方法）！", vbInformation, App.Title
        Exit Function
    ElseIf mcnDbBusiness.State <> adStateOpen Then
        MsgBoxEx "ConnectLogDB", "业务数据库的连接未开启！", vbInformation, App.Title
        Exit Function
    End If

    On Error GoTo hErr
    
    marrCurLogSet = Array(lglLogOFF, lgmNone, Nothing, Nothing, Nothing, Nothing)
    Set mdicLogControl = New Scripting.Dictionary
    
    '判断日志存储数据库是不是非业务数据库
    strSql = "Select Max(内容) 内容 From zlRegInfo Where 项目 = [1]"
    Set rsTmp = mdlDataBase.OpenSQLRecord(mcnDbBusiness, strSql, "日志服务器信息", "日志服务器")
    If rsTmp.EOF Then
        '业务发生与日志记录为同一数据库
        Set cnTmp = mcnDbBusiness
    Else
        If rsTmp!内容 & "" <> "" Then
            '其他数据库
            If rsTmp!内容 & "" Like "ZLSV*:*" Then
                '解密
                strCommand = mdlPublic.Sm4DecryptEcb(rsTmp!内容)
            Else
                '原文
                strCommand = rsTmp!内容
            End If
            '创建日志记录的数据连接
            Set cnTmp = CreateConnectionByCommand(strCommand _
                , "ZLUA" _
                , mdlPublic.Sm4DecryptEcb("ZLSV2:" & G_UA_PWD, GetGeneralAccountKey(G_UA_KEY)) _
                , False, strCurServer, strCurServer, mstrLogUser)
        Else
            '业务发生与日志记录为同一数据库
            Set cnTmp = mcnDbBusiness
        End If
    End If
    If cnTmp Is Nothing Then
        '缺省使用业务数据库
        Set cnTmp = mcnDbBusiness
    End If
    strCurServer = mdlPublic.GetServerInfo(cnTmp)

    If cnTmp = mcnDbBusiness Then
'        '创建日志记录的数据连接
'        mstrLogUser = "ZLUA"
'        strCommand = "SERVER=" & strCurServer & _
'                     " USER=" & mstrLogUser & _
'                     " PASS=" & mdlPublic.Sm4DecryptEcb("ZLSV2:" & G_UA_PWD, GetGeneralAccountKey(G_UA_KEY)) & _
'                     " TRANS=0"
'        Set mcnDbLog = CreateConnectionByCommand(strCommand)
        
        '尽量减少数据库连接会话，如果未设置额外的日志数据库就沿用mcnDBBusiness对象，并且不单独以ZLUA帐号提交日志
        mstrLogUser = GetConnectionUser(mcnDbBusiness)
        Set mcnDbLog = mcnDbBusiness
    Else
        Set mcnDbLog = cnTmp
    End If
    
    blnOpen = mcnDbLog.State = adStateOpen
    If blnOpen Then
        mstrLogDB = strCurServer
        ConnectLogDB = ReadLogSet(mcnDbLog, mstrOccurUser, mstrOccurIP)
        If ConnectLogDB Then
            mblnUseLog = mdicLogControl.Count > 0
        Else
            mblnUseLog = False
        End If
    Else
        Set mcnDbLog = Nothing
    End If

    ConnectLogDB = True
    Exit Function
    
hErr:
    Err.Clear
    Set mcnDbLog = Nothing
    ConnectLogDB = False
End Function

Public Function Log(ByVal strLogName As String, ByVal strComponentName As String, ByVal strModule As String _
    , ByVal strFuncName As String, ByVal llLogLevel As LogLevel, ByVal strLogTilte As String _
    , ParamArray arrPars() As Variant) As Boolean
'功能：记录指定日志级别的日志。
'返回：True记录成功；False记录失败
'参数：
'  strLogName：日志的业务分类名称。如一卡通日志等。传""时取最近一次不为空
'  strComponentName：日志发生的部件名称。使用App.EXEName。传""时取最近一次不为空
'  strModule：日志发生的模块。可以是ZLHIS体系的模块可以是VB的模块等。传""时取最近一次不为空
'  strFuncName：日志的发生的功能名。或者发生的VB函数。传""时取最近一次不为空
'  llLogLevel：当前记录日志性质，详见枚举说明
'  arrPars：产生格式：strLogTilte: arrPars(0)    arrPars(1)    ...arrPars(n)
    
    Dim strText As String, strParas As String
    Dim blnCanLog As Boolean
    Dim i As Integer, intIndex As Integer, intCurLogLevel As Integer, intCurLogMode As InfoIndex
    Dim varTemp As Variant, arrTmp As Variant
    
    On Error GoTo ErrH
    
    If Not mcnDbBusiness Is Nothing Then
'        If mstrCurLogName <> strLogName And LenB(strLogName) <> 0 Then
            Call ChangeCurrentLog(strLogName)           '变更当前日志（marrCurLogSet）
'        End If
        intCurLogLevel = CurrentLogLevel
        If intCurLogLevel >= lglInfo Then     '通过 marrCurLogSet(0-lglInfo) 获取当前日志等级
'            If LenB(strComponentName) = 0 Or LenB(strModule) = 0 Or LenB(strFuncName) = 0 Or LenB(strLogName) = 0 Then
'                If LenB(strLogName) = 0 Then strLogName = mstrCurLogName
'                arrTmp = Split(mstrCurFuncs & "...", ".")
'                If LenB(strComponentName) = 0 Then strComponentName = arrTmp(0)
'                If LenB(strModule) = 0 Then strModule = arrTmp(1)
'                If LenB(strFuncName) = 0 Then strFuncName = arrTmp(2)
'            End If
            
            '判断日志是否输出
            blnCanLog = IsCanWriteLog(UCase(strComponentName), UCase(strModule), UCase(strFuncName), "")
            
            If blnCanLog Then
                '日志输出
                
                '准备参数串
                strParas = ""
                For intIndex = LBound(arrPars) To UBound(arrPars)
                    If TypeName(arrPars(intIndex)) Like "*()" Then
                        '数组格式： Array(a1, a2, ... aN)
                        strText = ""
                        For i = 0 To UBound(arrPars(intIndex))
                            varTemp = arrPars(intIndex)(i)
                            If UCase$(TypeName(varTemp)) = "ERROR" Then
                                strText = strText & ", 空缺"
                            Else
                                strText = strText & ", " & mdlPublic.DisPlayOneValue(varTemp, intCurLogLevel > lglAll)
                            End If
                        Next
                        If Mid(strText, 3) <> "" Then
                            strParas = strParas & "    Array(" & Mid(strText, 3) & ")"
                        End If
                    Else
                        If UCase$(TypeName(arrPars(intIndex))) = "ERROR" Then
                            strParas = strParas & "    空缺"
                        Else
                            strParas = strParas & "    " & mdlPublic.DisPlayOneValue(arrPars(intIndex), intCurLogLevel > lglAll)
                        End If
                    End If
                Next
                If Len(strParas) <> 0 Then
                    strParas = strLogTilte & ": " & LTrim$(strParas)
                End If
    
                intCurLogMode = CurrentLogMode
                If intCurLogMode = lgmOnlyLocal Or intCurLogMode = lgmBoth Then
                    '本地日志输出
                    If OpenLogFile(mstrLogFile, mobjLogFile, mobjLocalLog) Then
                        Call GetLogTiltle(llLogLevel)
                        If mstrCurLogName <> strLogName Then
                            If LenB(mstrCurLogName) <> 0 Then
                                If LenB(mstrCurFuncs) <> 0 Then
                                    mobjLocalLog.WriteLine String(Len(mstrLogTiltle), " ") & "┣━┻" & mstrCurFuncs
                                End If
                                mobjLocalLog.WriteLine String(Len(mstrLogTiltle), " ") & "┗" & mstrCurLogName
                            End If
                            If LenB(strLogName) <> 0 Then
                                mobjLocalLog.WriteLine String(Len(mstrLogTiltle), " ") & "┏" & strLogName
                                mstrCurLogName = strLogName
                            End If
                            If LenB(strComponentName) <> 0 Or LenB(strModule) <> 0 Or LenB(strFuncName) <> 0 Then
                                mstrCurFuncs = strComponentName & "." & strModule & "." & strFuncName
                                mobjLocalLog.WriteLine String(Len(mstrLogTiltle), " ") & "┣━┳" & mstrCurFuncs
                            End If
                        Else
                            If mstrCurFuncs <> strComponentName & "." & strModule & "." & strFuncName Then
                                If LenB(mstrCurFuncs) <> 0 Then
                                    mobjLocalLog.WriteLine String(Len(mstrLogTiltle), " ") & "┣━┻" & mstrCurFuncs
                                End If
                                If LenB(strComponentName) <> 0 Or LenB(strModule) <> 0 Or LenB(strFuncName) <> 0 Then
                                    mstrCurFuncs = strComponentName & "." & strModule & "." & strFuncName
                                    mobjLocalLog.WriteLine String(Len(mstrLogTiltle), " ") & "┣━┳" & mstrCurFuncs
                                End If
                            End If
                        End If
                        mobjLocalLog.WriteLine mstrLogTiltle & "┣━╋━" & strParas
                        Call ClearObject(mobjLocalLog)
                    End If
                End If
    
                If intCurLogMode = lgmOnlyDB Or intCurLogMode = lgmBoth Then
                    '数据库日志输出
                    Call WriteDBLog(llLogLevel, strLogName, strComponentName, strModule, strFuncName, strParas)
                End If
            End If
        End If
    End If
    
    Log = True
    Exit Function
    
ErrH:
    If 0 = 1 Then
        Resume
    End If
    If mblnUseLog Then
        If intCurLogLevel >= lglError Then
            If blnCanLog Then
                intCurLogMode = CurrentLogMode
                If intCurLogMode = lgmOnlyLocal Or intCurLogMode = lgmBoth Then
                    Call GetLogTiltle(lglError)
                    If mobjLocalLog Is Nothing Then
                        If OpenLogFile(mstrLogFile, mobjLogFile, mobjLocalLog) Then
                            mobjLocalLog.WriteLine mstrLogTiltle & "┣━╋━╋Log发生错误：" & Err.Number & "-" & Err.Description
                        End If
                    Else
                        mobjLocalLog.WriteLine mstrLogTiltle & "┣━╋━╋Log发生错误：" & Err.Number & "-" & Err.Description
                    End If
                    Call ClearObject(mobjLocalLog)
                End If
                If intCurLogMode = lgmOnlyDB Or intCurLogMode = lgmBoth Then
                    Call WriteDBLog(lglError, strLogName, strComponentName, strModule, strFuncName _
                        , "Log发生错误：" & Err.Number & "-" & Err.Description)
                End If
            End If
        End If
    End If
    Err.Clear
End Function

Public Function LogCall(ByVal strLogName As String, ByVal strComponentName As String, ByVal strModule As String _
    , ByVal strFuncName As String, ByVal strCallName As String, ByVal lcsCurentLogCallState As LogCallState _
    , ParamArray arrPars() As Variant) As Boolean
    
'功能：记录重要接口的调用情况。如WebAPI或者存储过程。该日志的性质属于LogLevel_Info（重要信息级别）
'返回：True记录成功；False记录失败
'参数：
'  strLogName：日志的业务分类名称。如一卡通日志等。传""时取最近一次不为空
'  strComponentName：日志发生的部件名称。使用App.EXEName。传""时取最近一次不为空
'  strModule：日志发生的模块。可以是ZLHIS体系的模块可以是VB的模块等。传""时取最近一次不为空
'  strFuncName：日志的发生的功能名。或者发生的VB函数。传""时取最近一次不为空
'  strCallName：WebAPI名称或者存储过程名称
'  lcsCurentLogCallState：标识调用的时机。开始调用或者结束调用。
'  arrPars：产生格式： arrPars(0), arrPars(1), ... , arrPars(n)
    
    Dim i As Integer, intIndex As Integer, intCurLogLevel As Integer, intCurLogMode As Integer
    Dim blnWrite As Boolean, blnCanLog As Boolean
    Dim varTemp As Variant, arrTmp As Variant
    Dim strText As String, strParas As String, strTemp As String
    
    On Error GoTo ErrH
    
    If mblnUseLog Then
'        If mstrCurLogName <> strLogName And LenB(strLogName) <> 0 Then
            Call ChangeCurrentLog(strLogName)
'        End If
        intCurLogLevel = CurrentLogLevel()
        If intCurLogLevel >= lglInfo Then
'            If LenB(strComponentName) = 0 Or LenB(strModule) = 0 Or LenB(strFuncName) = 0 Or LenB(strLogName) = 0 Then
'                If LenB(strLogName) = 0 Then strLogName = mstrCurLogName
'                arrTmp = Split(mstrCurFuncs & "...", ".")
'                If LenB(strComponentName) = 0 Then strComponentName = arrTmp(0)
'                If LenB(strModule) = 0 Then strModule = arrTmp(1)
'                If LenB(strFuncName) = 0 Then strFuncName = arrTmp(2)
'            End If
            blnCanLog = IsCanWriteLog(UCase(strComponentName), UCase(strModule), UCase(strFuncName), UCase(strCallName))
            If blnCanLog Then
                strParas = ""
                For intIndex = 0 To UBound(arrPars)
                    If TypeName(arrPars(intIndex)) Like "*()" Then
                        '数组
                        strText = ""
                        For i = 0 To UBound(arrPars(intIndex))
                            varTemp = arrPars(intIndex)(i)
                            strTemp = ""
                            Call JoinParamString(strTemp, varTemp)
                            If strTemp <> "" Then
                                strText = strText & ", " & strTemp
                            End If
                        Next
                        If Mid(strText, 3) <> "" Then
                            strParas = strParas & "    " & Mid(strText, 3)
                        End If
                    Else
                        strTemp = ""
                        Call JoinParamString(strTemp, arrPars(intIndex))
                        strParas = strParas & "    " & strTemp
                    End If
                Next
                strParas = LTrim(strParas)
                intCurLogMode = CurrentLogMode
                
                If intCurLogMode = lgmOnlyLocal Or intCurLogMode = lgmBoth Then
                    '本地日志输出
                    If OpenLogFile(mstrLogFile, mobjLogFile, mobjLocalLog) Then
                        Call GetLogTiltle(lglInfo)
                        If mstrCurLogName <> strLogName Then
                            If LenB(mstrCurLogName) <> 0 Then
                                If LenB(mstrCurFuncs) <> 0 Then
                                    mobjLocalLog.WriteLine String(Len(mstrLogTiltle), " ") & "┣━┻" & mstrCurFuncs
                                End If
                                mobjLocalLog.WriteLine String(Len(mstrLogTiltle), " ") & "┗" & mstrCurLogName
                            End If
                            If LenB(strLogName) <> 0 Then
                                mobjLocalLog.WriteLine String(Len(mstrLogTiltle), " ") & "┏" & strLogName
                                mstrCurLogName = strLogName
                            End If
                            If LenB(strComponentName) <> 0 Or LenB(strModule) <> 0 Or LenB(strFuncName) <> 0 Then
                                mstrCurFuncs = strComponentName & "." & strModule & "." & strFuncName
                                mobjLocalLog.WriteLine String(Len(mstrLogTiltle), " ") & "┣━┳" & mstrCurFuncs
                            End If
                        Else
                            If mstrCurFuncs <> strComponentName & "." & strModule & "." & strFuncName Then
                                If LenB(mstrCurFuncs) <> 0 Then
                                    mobjLocalLog.WriteLine String(Len(mstrLogTiltle), " ") & "┣━┻" & mstrCurFuncs
                                End If
                                If LenB(strComponentName) <> 0 Or LenB(strModule) <> 0 Or LenB(strFuncName) <> 0 Then
                                    mstrCurFuncs = strComponentName & "." & strModule & "." & strFuncName
                                    mobjLocalLog.WriteLine String(Len(mstrLogTiltle), " ") & "┣━┳" & mstrCurFuncs
                                End If
                            End If
                        End If
                        If lcsCurentLogCallState = lcsCallBegin Then
                            mobjLocalLog.WriteLine mstrLogTiltle & "┣━╋━┳Begin " & strCallName & " Using " & strParas
                        Else
                            mobjLocalLog.WriteLine mstrLogTiltle & "┣━╋━┻End " & strCallName & " Return " & strParas
                        End If
                        blnWrite = True
                        Call ClearObject(mobjLocalLog)
                    End If
                End If
    
                If intCurLogMode = lgmOnlyDB Or intCurLogMode = lgmBoth Then
                    '数据库日志输出
                    Call WriteDBLog(lglInfo, strLogName, strComponentName, strModule, strFuncName _
                        , strParas, strCallName, lcsCurentLogCallState)
                End If
            End If
        End If
    End If
    
    LogCall = True
    Exit Function
    
ErrH:
    If 0 = 1 Then
        Resume
    End If
    If mblnUseLog Then
        If CurrentLogLevel >= lglError Then
            If blnCanLog Then
                intCurLogMode = CurrentLogMode
                If intCurLogMode = lgmOnlyLocal Or intCurLogMode = lgmBoth Then
                    '本地日志输出
                    Call GetLogTiltle(lglError)
                    If mobjLocalLog Is Nothing Then
                        If OpenLogFile(mstrLogFile, mobjLogFile, mobjLocalLog) Then
                            If Not blnWrite And lcsCurentLogCallState = lcsCallBegin Then
                                mobjLocalLog.WriteLine mstrLogTiltle & "┣━╋━┳Begin " & strCallName
                            End If
                            mobjLocalLog.WriteLine mstrLogTiltle & "┣━╋━╋LogCall发生错误：" & Err.Number & "-" & Err.Description
                            If Not blnWrite And lcsCurentLogCallState <> lcsCallBegin Then
                                mobjLocalLog.WriteLine mstrLogTiltle & "┣━╋━┻End " & strCallName
                            End If
                        End If
                    Else
                        If Not blnWrite And lcsCurentLogCallState = lcsCallBegin Then
                            mobjLocalLog.WriteLine mstrLogTiltle & "┣━╋━┳Begin " & strCallName
                        End If
                        mobjLocalLog.WriteLine mstrLogTiltle & "┣━╋━╋LogCall发生错误：" & Err.Number & "-" & Err.Description
                        If Not blnWrite And lcsCurentLogCallState <> lcsCallBegin Then
                            mobjLocalLog.WriteLine mstrLogTiltle & "┣━╋━┻End " & strCallName
                        End If
                    End If
                    Call ClearObject(mobjLocalLog)
                End If
                If intCurLogMode = lgmOnlyDB Or intCurLogMode = lgmBoth Then
                    '数据库日志输出
                    If Not blnWrite And lcsCurentLogCallState = lcsCallBegin Then
                        Call WriteDBLog(lglInfo, strLogName, strComponentName, strModule, strFuncName _
                            , "发生错误！" & strParas, strCallName, lcsCurentLogCallState)
                    End If
                    Call WriteDBLog(lglError, strLogName, strComponentName, strModule, strFuncName _
                        , "LogCall发生错误：" & Err.Number & "-" & Err.Description, strCallName)
                    If Not blnWrite And lcsCurentLogCallState <> lcsCallBegin Then
                        Call WriteDBLog(lglInfo, strLogName, strComponentName, strModule, strFuncName _
                            , "发生错误！" & strParas, strCallName, lcsCurentLogCallState)
                    End If
                End If
            End If
        End If
    End If
    Err.Clear
End Function

Private Function GetLogName() As String
    Dim strLogFile As String, strTmp As String
    Dim arrTmp As Variant
    Dim i As Long, lngCount As Long
    Dim objFSO As FileSystemObject
    Dim objTSM As TextStream
    
    On Error GoTo ErrH
    
    '生成日志文件名。进程名_S会话ID P进程ID _YYYYMMDD
    strLogFile = mdlPublic.AppsoftPath & "\Log\日志跟踪\" & mobjLogFile.GetBaseName(mdlPublic.StartExePath) & "_S" & _
                 mdlPublic.SessionID & "P" & mdlPublic.CurrentPID
    strTmp = strLogFile & "_" & Format(Now, "YYYYMMDD") & ".log"
    If mobjLogFile.FileExists(strTmp) Then
        '文件存在
        Set objFSO = New FileSystemObject
        Set objTSM = objFSO.OpenTextFile(strTmp, ForReading)
        lngCount = 0
        Do While objTSM.AtEndOfLine = False
            lngCount = lngCount + 1
            objTSM.SkipLine
        Loop
        objTSM.Close
        
        If lngCount >= MLNG_MAXLINE Then
            '行超最大限制，新建日志文档记录日志
            strLogFile = strLogFile & "_" & Format(Now, "YYYYMMDD-HH") & ".log"
        Else
            strLogFile = strTmp
        End If
    Else
        '文件不存在，并且依次判断文件路径是否存在
        On Error GoTo hErr
        
        strLogFile = strTmp
        If mobjLogFile.FolderExists(mobjLogFile.GetParentFolderName(strLogFile)) Then
            arrTmp = Split(strLogFile, "\")
            strTmp = arrTmp(0) & "\"
            For i = 1 To UBound(arrTmp) - 1
                strTmp = strTmp & arrTmp(i) & "\"
                If Not mobjLogFile.FolderExists(strTmp) Then
                    Call mobjLogFile.CreateFolder(strTmp)
                End If
            Next
        End If
    End If
    
    GetLogName = strLogFile
    Exit Function
    
ErrH:
    MsgBoxEx "GetLogName", "获取日志文档异常！" & vbCr & Err.Description, vbInformation, App.Title
    Exit Function
    
hErr:
    MsgBoxEx "GetLogName", "存储文档的目录可能受限，不能正常输出日志！" & vbCr & strLogFile, vbInformation, App.Title
End Function

Private Sub GetLogTiltle(ByVal intLogLevel As LogLevel)
'功能：获取日志头，并进行行控制
    'Call LogWriteHeader
    mstrLogTiltle = mstrLogLevelName(intLogLevel + 1) & Format(Now, "YYYY-MM-DD hh:mm:ss")
End Sub

Private Sub LogWriteHeader(Optional ByVal strLogFile As String)
'功能：书写日志头

    If strLogFile = "" Then Exit Sub
    If mobjLogFile Is Nothing Then Exit Sub
    
    Call OpenLogFile(strLogFile, mobjLogFile, mobjLocalLog)
    If mobjLocalLog Is Nothing Then
        MsgBoxEx "LogWriteHeader", "存储文档的目录可能受限，不能正常输出日志！" & vbCr & strLogFile, vbInformation, App.Title
        Exit Sub
    End If
    
    mobjLocalLog.WriteLine String(80, "-")
    mobjLocalLog.WriteLine "[Create Time]  " & Format(Now, "yyyy-mm-dd hh:mm:ss")
    mobjLocalLog.WriteLine "[Environment] " & _
        " ProcessID=" & mdlPublic.CurrentPID & _
        " ProcessUser=" & mdlPublic.ProcessUserFullName & _
        " OSSessionID=" & mdlPublic.SessionID & _
        " OSSessionUser=" & mdlPublic.SessionUserFullName
    Call LogWriteUser(True)
    mobjLocalLog.WriteLine String(80, "-")
    
    Call ClearObject(mobjLocalLog)
End Sub

Private Function OpenLogFile(ByVal strFile As String, ByRef objFSO As FileSystemObject _
    , ByRef objTextStream As TextStream) As Boolean
    
    On Error Resume Next
    Set objTextStream = Nothing
    Set objTextStream = objFSO.OpenTextFile(strFile, ForAppending, True)
    OpenLogFile = Err.Number = 0
End Function

Private Sub LogWriteUser(Optional ByVal blnFirst As Boolean = False)
    If mobjLocalLog Is Nothing Then Exit Sub
    
    If Not blnFirst Then
        Call GetLogTiltle(lglInfo)
        If LenB(mstrCurLogName) <> 0 Then
            If LenB(mstrCurFuncs) <> 0 Then
                mobjLocalLog.WriteLine String$(Len(mstrLogTiltle), " ") & "┣━┻" & mstrCurFuncs
            End If
            mobjLocalLog.WriteLine String$(Len(mstrLogTiltle), " ") & "┗" & mstrCurLogName
        End If
    End If
    mstrCurLogName = ""
    mstrCurFuncs = ""
    If Not mobjLocalLog Is Nothing Then
        If Not blnFirst Then
            mobjLocalLog.WriteLine String$(80, "-")
            mobjLocalLog.WriteLine "[BusDBChange]  " & Format$(Now, "yyyy-mm-dd hh:mm:ss")
        End If
        mobjLocalLog.WriteLine "[Business DB]  User=" & mstrOccurUser & " Server=" & mstrOccurServer & _
                               " SID=" & mlngOccurSID & " IP=" & mstrOccurIP
        mobjLocalLog.WriteLine "[Log      DB]  User=" & mstrLogUser & " Server=" & mstrLogDB
        If Not blnFirst Then
            mobjLocalLog.WriteLine String$(80, "-")
        End If
    End If
End Sub

Private Sub Class_Initialize()
    Const STR_INFO As String = "[UnDefine];[ LogOFF ];[  Error ];[  Warn  ];[  Info  ];[  Trace ];[   All  ]"
    
    Dim i As Integer
    Dim arrVal() As String
    
    On Error Resume Next
    
    Set mobjLogFile = New FileSystemObject
    If mobjLogFile Is Nothing Then
        MsgBoxEx "Class_Initialize", "操作系统环境异常，缺少“Microsoft Scripting Runtime”组件，无法正常输出业务日志！"
        Exit Sub
    End If
    
    arrVal = Split(STR_INFO, ";")
    For i = 0 To UBound(arrVal)
        mstrLogLevelName(i) = arrVal(i)
    Next

    mstrProcessName = mobjLogFile.GetBaseName(mdlPublic.StartExePath())
    mlngPID = mdlPublic.CurrentPID
    mstrStation = mdlPublic.ComputerName
    marrCurLogSet = Array(lglLogOFF, lgmNone, Nothing, Nothing, Nothing, Nothing)
    Set mdicLogControl = New Scripting.Dictionary
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    
    Erase mstrLogLevelName
    Set marrCurLogSet = Nothing
    Set mobjLocalLog = Nothing
    Set mobjLogFile = Nothing
    Set mcnDbBusiness = Nothing
    Set mcnDbLog = Nothing
End Sub

Private Sub ClearObject(objInput As Object)
    If Not objInput Is Nothing Then
        On Error Resume Next
        If TypeOf objInput Is ADODB.Connection Then
            'ADO.Connection对象不能使用Close方法，会引起传入clsLogger的连接对象丢失
            '...
        Else
            objInput.Close
        End If
        Set objInput = Nothing
        On Error GoTo 0
    End If
End Sub

Private Function GetGeneralAccountKey(ByRef strKey As String) As String
    Dim arrTmp()    As Byte
    Dim i           As Long
    arrTmp = mdlPublic.HexStringToByte(strKey, 16)
    For i = LBound(arrTmp) To UBound(arrTmp)
        If i Mod 2 = 0 Then
            arrTmp(i) = 255 - arrTmp(i)
        ElseIf i Mod 3 = 0 Then
            arrTmp(i) = (arrTmp(i) + i) Mod 256
        End If
    Next
    GetGeneralAccountKey = mdlPublic.ByteToHexString(arrTmp)
End Function

Private Function CreateConnectionByCommand(ByVal strCommand As String, Optional ByVal strDefaultUser As String _
    , Optional ByVal strDefaultPWD As String, Optional ByVal blnDefaultTrans As Boolean _
    , Optional ByVal strDefaultServer As String, Optional strOutSerrver As String _
    , Optional strOutUser As String) As ADODB.Connection
    
    Dim strServer As String, strUser As String, strPass As String
    Dim strSID As String, strIp As String, strPort As String
    Dim arrTmp As Variant
    Dim i As Long
'    Dim blnTrans As Boolean
    
    arrTmp = Split(strCommand, " ")
    strCommand = ""
    strOutSerrver = ""
    strOutUser = ""
    For i = LBound(arrTmp) To UBound(arrTmp)
        If Trim$(arrTmp(i)) <> "" Then
            If arrTmp(i) Like "USER=*" Then
                strUser = Mid$(arrTmp(i), Len("USER=*"))
            ElseIf arrTmp(i) Like "PASS=*" Then
                strPass = Mid$(arrTmp(i), Len("PASS=*"))
'            ElseIf arrTmp(i) Like "TRANS=*" Then
'                blnTrans = Val(Mid$(arrTmp(i), Len("TRANS=*"))) = 1
            ElseIf arrTmp(i) Like "SERVER=*" Then
                strServer = Mid$(arrTmp(i), Len("SERVER=*"))
            Else
                If LenB(strServer) = 0 Then
                    strServer = arrTmp(i)
                End If
            End If
        End If
    Next
    If strServer = "" Then
        strServer = strDefaultServer
    End If

    If strUser = "" Then
        strUser = strDefaultUser
        strPass = strDefaultPWD
'        blnTrans = blnDefaultTrans
    End If
    
    If InStr(strServer, "/") > 0 Then
        arrTmp = Split(strServer, "/")
        strSID = arrTmp(1)
        If InStr(arrTmp(0), ":") > 0 Then
            arrTmp = Split(arrTmp(0), ":")
            strIp = arrTmp(0)
            strPort = arrTmp(1)
        Else
            strIp = arrTmp(0)
            strPort = "1521"
        End If
        strServer = "(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=" & strIp & ")(PORT=" & strPort & "))(CONNECT_DATA=(SERVICE_NAME=" & strSID & ")))"
        
        '下面这种加了ADDRESS_LIST的写法，在ODBC下，只支持SID，不支持SERVICE_NAME;OLEDB则两种都支持
        'If bytProvider = enuProvider.MSODBC Then
        'strServer = "(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=" & strIP & ")(PORT=" & strPort & ")))(CONNECT_DATA=(SID=" & strSID & ")))"
    End If
    
    On Error GoTo ErrH
    Set CreateConnectionByCommand = New ADODB.Connection
    CreateConnectionByCommand.Provider = "OraOLEDB.Oracle"
    Call CreateConnectionByCommand.Open( _
         "Provider=OraOLEDB.Oracle.1;Persist Security Info=True;PLSQLRSet=1;Data Source=" & strServer _
        , strUser, strPass)
    strOutUser = strUser
    Exit Function
    
ErrH:
    Set CreateConnectionByCommand = Nothing
    Err.Clear
End Function

Private Function ReadLogSet(ByRef cnLogDB As ADODB.Connection, ByVal strUser As String, ByVal strIp As String) As Boolean
'功能：准备日志记录需要的数据
'返回：True成功；False失败
'参数：
'  cnLogDB：日志数据库的连接对象
'  strUser：业务发生的用户
'  strIP：业务发生的终端IP
    
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, strPreName As String
    Dim blnLocalLog As Boolean, blnDBLog As Boolean
    Dim arrSets() As Variant
    Dim lngStart As Long
    Dim dtCurrent As Date
    Dim intCurMode As Integer

    On Error GoTo ErrH
    
    '获取日志生产的配置信息（级别、输出形态（本地、数据）等，以及仔细的日志存储条件）
    strSql = _
        "Select c.Id, c.Name, c.Log_Keep_Days, c.Log_Mode, c.Log_Level, c.Component_Names, c.Module_Names, c.Function_Names," & vbNewLine & _
        "    c.Call_Names, c.Sortid " & vbNewLine & _
        "From (Select a.Id, a.Name, a.Log_Keep_Days, a.Log_Mode, a.Log_Level, b.Station, b.Ip, b.User_Name, b.Component_Names," & vbNewLine & _
        "          b.Module_Names, b.Function_Names, b.Call_Names," & vbNewLine & _
        "          Decode(b.Station, Null, 0, 1) + Decode(b.User_Name, Null, 0, 1) + Decode(b.Ip, Null, 0, 1) Sortid " & vbNewLine & _
        "      From Zllogcategory A, Zllogset B " & vbNewLine & _
        "      Where a.Id = b.Category_Id(+) And Sysdate Between a.Begin_Time And a.End_Time And a.Log_Mode > 0 " & vbNewLine & _
        "          And a.Log_Level > 0 And Nvl(Station, [1]) =[1] And Nvl(User_Name, [2]) =[2] And Nvl(Ip, [3]) =[3]" & vbNewLine & _
        ") C " & vbNewLine & _
        "Order By c.Id, c.Sortid Desc"
    Set rsTmp = mdlDataBase.OpenSQLRecord(cnLogDB, strSql, "ReadLogSet", mstrStation, strUser, strIp)
    Do While Not rsTmp.EOF
        If strPreName <> rsTmp!Name Then
            If strPreName <> "" Then
                mdicLogControl.Add strPreName, arrSets
            End If
            strPreName = rsTmp!Name & ""
            
            ReDim arrSets(5)
            arrSets(ifiLogLevel) = Val(rsTmp!Log_Level & "")                                            'Index=0 表 Level
            arrSets(ifiLogMode) = Val(rsTmp!Log_Mode & "")                                              'Index=1 表 Mode
            lngStart = ifiLogMode + 1
            Set arrSets(lngStart + ifiComponent) = AnalyseSet(UCase$(rsTmp!Component_Names & ""))       'Index=2 表 Component，并且是Scripting.Dictionary类型
            Set arrSets(lngStart + ifiFunction) = AnalyseSet(UCase$(rsTmp!Function_Names & ""))         'Index=3 表 Function
            Set arrSets(lngStart + ifiModule) = AnalyseSet(UCase$(rsTmp!Module_Names & ""))             'Index=4 表 Module
            Set arrSets(lngStart + ifiCall) = AnalyseSet(UCase$(rsTmp!Call_Names & ""))                 'Index=5 表 CallName
            
            intCurMode = arrSets(ifiLogMode)
            If intCurMode = lgmOnlyLocal Or intCurMode = lgmBoth Then
                blnLocalLog = True
            End If
            If intCurMode = lgmOnlyDB Or intCurMode = lgmBoth Then
                blnDBLog = True
            End If
        Else
            If rsTmp!Sortid <> 0 Then
                '其他配置
                lngStart = UBound(arrSets) + 1
                ReDim Preserve arrSets(UBound(arrSets) + 4)
                Set arrSets(lngStart + ifiComponent) = AnalyseSet(UCase(rsTmp!Component_Names & ""))
                Set arrSets(lngStart + ifiFunction) = AnalyseSet(UCase(rsTmp!Function_Names & ""))
                Set arrSets(lngStart + ifiModule) = AnalyseSet(UCase(rsTmp!Module_Names & ""))
                Set arrSets(lngStart + ifiCall) = AnalyseSet(UCase(rsTmp!Call_Names & ""))
            End If
        End If
        rsTmp.MoveNext
    Loop
    '收尾处理
    If strPreName <> "" Then
        If mdicLogControl.Exists(strPreName) Then
            mdicLogControl(strPreName) = arrSets    '更新
        Else
            mdicLogControl.Add strPreName, arrSets  '新增
        End If
    End If
    
    '获取当前时间，和最大的日志保留天数，根据这些信息清理本地日志。数据库日志通过自动作业处理。
    strSql = "Select Max(a.Log_Keep_Days) Log_Keep_Days, Sysdate Cur " & vbNewLine & _
             "From Zllogcategory A " & vbNewLine & _
             "Where Sysdate Between a.Begin_Time And a.End_Time And a.Log_Mode > 0 And a.Log_Level > 0"
    Set rsTmp = mdlDataBase.OpenSQLRecord(cnLogDB, strSql, "ReadLogSet")
    If rsTmp.EOF = False Then
        dtCurrent = rsTmp!Cur
        mlngKeepDays = Val(rsTmp!Log_Keep_Days & "")
        If mlngKeepDays <= 0 Then mlngKeepDays = 7
        Call ClearLocalLog(mlngKeepDays, dtCurrent)
    End If
    
    '准备日志对象（日志文件和写入日志头）
    If blnLocalLog Then
        mstrLogFile = GetLogName()
        If mstrLogFile <> "" Then
            Set mobjLocalLog = mobjLogFile.OpenTextFile(mstrLogFile, ForAppending, True)
            Call LogWriteHeader(mstrLogFile)
        Else
            Call ClearObject(mobjLocalLog)
        End If
    Else
        Call ClearObject(mobjLocalLog)  '本地日志未开启就清除
    End If
    
    '准备日志对象（日志数据库的连接）
    If blnDBLog Then
        Set mcnDbLog = cnLogDB          '有配置数据库存储日志
    Else
        On Error Resume Next
        If cnLogDB <> mcnDbBusiness Then
            cnLogDB.Close               '关闭非传入的连接
        End If
        On Error GoTo 0
        Set cnLogDB = Nothing           '直接nothing断开指向，不要去Close操作
    End If
    mblnUseLog = blnDBLog Or blnLocalLog
    ReadLogSet = True
    Exit Function
    
ErrH:
    Err.Clear
End Function

Private Function AnalyseSet(strSets As String) As Scripting.Dictionary
    Dim objSet As Scripting.Dictionary
    Dim arrTmp As Variant
    Dim i As Long
    
    If strSets <> "*" Then
        arrTmp = Split(strSets, ";")
        Set objSet = New Scripting.Dictionary
        For i = 0 To UBound(arrTmp)
            If Trim(arrTmp(i)) <> "" Then
                objSet.Add arrTmp(i), 1
            End If
        Next
        If objSet.Count = 0 Then
            Set objSet = Nothing
        End If
    End If
    Set AnalyseSet = objSet
End Function

Private Sub ClearLocalLog(ByVal lngKeepDays As Long, ByVal dtCurrent As Date)
    Dim objFile As File
    Dim arrFile As Variant
    Dim strMin As String, strCur As String
    Dim i As Long
    
    On Error GoTo ErrH
    
    If lngKeepDays <= 0 Then lngKeepDays = 7
    strMin = Format(DateAdd("d", -lngKeepDays, Date), "yyyymmdd")
    arrFile = Array()
    For Each objFile In mobjLogFile.GetFolder(mdlPublic.AppsoftPath & "\Log\日志跟踪").Files
        If UCase(objFile.Name) Like "*_S*P*_*.LOG" Then
            strCur = Mid(objFile.Name, InStrRev(objFile.Name, "_") + 1)
            If strMin > strCur Then
                '待删除的日志文件
                ReDim Preserve arrFile(UBound(arrFile) + 1)
                arrFile(UBound(arrFile)) = objFile.Path
            End If
        End If
    Next
    For i = 0 To UBound(arrFile)
        On Error Resume Next
        Call mobjLogFile.DeleteFile(arrFile(i), True)
        If Err.Number <> 0 Then Err.Clear
    Next
    Exit Sub
    
ErrH:
    Err.Clear
End Sub

Private Sub ChangeCurrentLog(strLogName As String)
    If mdicLogControl.Exists(strLogName) Then
        '指定日志配置
        marrCurLogSet = mdicLogControl(strLogName)
        If IsNull(marrCurLogSet) Or IsEmpty(marrCurLogSet) Then
            '缺省日志配置
            mdicLogControl(strLogName) = Array(lglLogOFF, lgmNone, Nothing, Nothing, Nothing, Nothing)
            marrCurLogSet = mdicLogControl(strLogName)
        End If
    Else
        '缺省日志配置
        mdicLogControl.Add strLogName, Array(lglLogOFF, lgmNone, Nothing, Nothing, Nothing, Nothing)
        marrCurLogSet = mdicLogControl(strLogName)
    End If
End Sub

Private Function MatchingLogSet(ByVal objLogSet As Variant, ByVal lngIndex As Long, ByVal strName As String) As Byte
'功能：是否匹配关键的信息
'返回：0-未设置；1-匹配设置；2-不匹配设置

    Dim dicInfo As Scripting.Dictionary
    Dim varRet As Variant
    
    On Error Resume Next
    
    MatchingLogSet = False
    Set dicInfo = objLogSet(lngIndex + ifiComponent)
    On Error GoTo 0
    If Not dicInfo Is Nothing Then
        If dicInfo.Exists(strName) Then
            varRet = dicInfo(strName)
            If IsNull(varRet) Or IsEmpty(varRet) Then
                '不匹配
                MatchingLogSet = 2
            Else
                '匹配
                MatchingLogSet = 1
            End If
        End If
    Else
        MatchingLogSet = 0
    End If
End Function

Private Function IsCanWriteLog(ByVal strComponentName As String, ByVal strModule As String, ByVal strFuncName As String _
    , ByVal strCallName As String) As Boolean
'功能：判断是否可以写日志
'返回：True写日志；False不写日志
'参数：
'  strComponentName：部件名
'  strModule：模块名
'  strFuncName：功能名
'  strCallName：调用的用户名

    Dim lngStart As Long
    Dim bytResult As Byte

    On Error GoTo ErrH

    For lngStart = ifiLogMode + 1 To UBound(marrCurLogSet) Step 4
        '部件名
        bytResult = MatchingLogSet(marrCurLogSet, lngStart + ifiComponent, strComponentName)
        '模块名
        If bytResult = Val("2-不匹配设置") Or bytResult = Val("0-未设置") Then
            bytResult = MatchingLogSet(marrCurLogSet, lngStart + ifiModule, strModule)
        Else
            Exit For    '匹配设置就不再检查后面的要素
        End If
        '功能名
        If bytResult = Val("2-不匹配设置") Or bytResult = Val("0-未设置") Then
            bytResult = MatchingLogSet(marrCurLogSet, lngStart + ifiFunction, strFuncName)
        Else
            Exit For
        End If
        '用户名
        If bytResult = Val("2-不匹配设置") Or bytResult = Val("0-未设置") Then
            bytResult = MatchingLogSet(marrCurLogSet, lngStart + ifiCall, strCallName)
        Else
            Exit For
        End If
    Next
    
    IsCanWriteLog = bytResult <> Val("2-不匹配设置")
    Exit Function

ErrH:
    Err.Clear
End Function

Private Sub WriteDBLog(llCurent As LogLevel, strLogName As String, strComponentName As String, strModule As String _
    , strFuncName As String, strParasInfo As String, Optional strCallName As String _
    , Optional ByVal lcsCurentLogCallState As LogCallState = -1)
    
    On Error GoTo ErrH
    If LenB(strParasInfo) < 4000 Then   '不能大于 zlLogInfo.LOG_INFO 字段长度
        Call mdlDataBase.CallProcedure(mcnDbLog, "Zlloginfo_Insert", "WriteDBLog", False, llCurent _
            , mstrOccurServer, mstrOccurUser, mlngOccurSID, mstrOccurIP, mstrStation, mlngPID _
            , mstrProcessName, strLogName, strComponentName, strModule, strFuncName, strCallName _
            , IIf(lcsCurentLogCallState < 0, Null, lcsCurentLogCallState), strParasInfo)
    Else
        Call mdlDataBase.CallProcedure(mcnDbLog, "Zlloginfo_Insert", "WriteDBLog", False, llCurent _
            , mstrOccurServer, mstrOccurUser, mlngOccurSID, mstrOccurIP, mstrStation, mlngPID _
            , mstrProcessName, strLogName, strComponentName, strModule, strFuncName, strCallName _
            , IIf(lcsCurentLogCallState < 0, Null, lcsCurentLogCallState), "", strParasInfo)
    End If
    Exit Sub
    
ErrH:
    mcnDbLog.Errors.Clear
    Err.Clear
End Sub

Private Sub JoinParamString(ByRef strParams As String, ByVal varPar As Variant)
    Dim objCommand As Object
    Dim i As Integer, intCurLogLevel As Integer
    Dim strText As String, strResult As String
    
    intCurLogLevel = CurrentLogLevel
    If IsObject(varPar) Then
        '对象参数
        If TypeOf varPar Is ADODB.Parameters Then
            Set objCommand = varPar
            strResult = ", " & "{ADODB.Parameters("
            strText = ""
            For i = 0 To objCommand.Parameters.Count - 1
                strText = strText & _
                    ", " & _
                    mdlPublic.DisPlayOneValue(objCommand.Parameters(i).Value, intCurLogLevel >= lglAll)
            Next
            strResult = strResult & Mid(strText, 3) & ")}"
        Else
            '其他对象暂不支持
        End If
    ElseIf UCase$(TypeName(varPar)) = "ERROR" Then
        strResult = ", 空缺"
    Else
        '常规参数
        strResult = ", " & mdlPublic.DisPlayOneValue(varPar, intCurLogLevel >= lglAll)
    End If
    
    If strResult Like ", *" Then
        strParams = strParams & Mid$(strResult, 3)
    Else
        strParams = strParams & strResult
    End If
End Sub

Private Function GetConnectionUser(ByVal objCN As ADODB.Connection) As String
    If objCN Is Nothing Then Exit Function
    If Not UCase$(objCN.ConnectionString) Like "*[;| ]USER[ |=]*" Then Exit Function
    
    On Error Resume Next
    GetConnectionUser = objCN.Properties("USER ID").Value
End Function

Private Function MsgBoxEx(ByVal strProcName As String, ByVal Prompt As String _
    , Optional ByVal Buttons As VbMsgBoxStyle _
    , Optional ByVal Title As String) As VbMsgBoxResult

    If mblnSilentMode Then
        Call Log("公共日志", App.ProductName, "clsLog", strProcName, lglTrace, Prompt)
    Else
        MsgBoxEx = MsgBox(Prompt, Buttons, Title)
    End If
End Function
