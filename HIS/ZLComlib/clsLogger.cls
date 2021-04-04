VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@模块 clsLogger-2019/9/16
'@编写 lshuo
'@功能
'
'@引用
'
'@备注
'
Option Explicit
'---------------------------------------------------------------------------
'                0、API和常量声明
'---------------------------------------------------------------------------
Public Enum LogLevel
    LogLevel_UnDefined = -1                 '尚未设置，应用于模块部件级别
    LogLevel_LogOFF = 0                     '不记录日志
    LogLevel_Error = 1                      '只记录错误
    LogLevel_Warn = 2                       '记录警告
    LogLevel_Info = 3                       '记录重要信息
    LogLevel_Trace = 4                      '记录跟踪信息
    LogLevel_All = 5                        '记录所有日志信息
End Enum

Public Enum LogMode
    LogMode_None = 0               '不记录日志
    LogMode_OnlyLocal = 1          '仅记录本地日志
    LogMode_OnlyDB = 2             '仅记录数据库日志
    LogMode_Both = 3               '本地日志和数据库日志都记录
End Enum

Public Enum LogCallState
    LogCallState_CallBegin = 0
    LogCallState_CallEnd = 1
End Enum

Private Enum InfoIndex
    Info_LogLevel = 0
    Info_LogMode = 1
    Info_Component = 0
    Info_Function = 1
    Info_Module = 2
    Info_Call = 3
End Enum
'---------------------------------------------------------------------------
'                1、常规变量
'---------------------------------------------------------------------------
Private mcnDbLog            As ADODB.Connection         '数据库日志连接。
Private mobjLocalLog        As TextStream               '本地日志对象。

Private mstrStation         As String                   '工作站即电脑名
Private mlngPID             As Long
Private mstrProcessName     As String

Private mstrLogLevelName(6)                             As String               '日志级别名称
Private mblnUseLog                                      As Boolean
Private mlngCurLogCount                                 As Long
Private mlngMaxLogCount                                 As Long

'临时变量
Private mstrTmp                                         As String
Private mstrParas                                       As String
Private mlngIndex                                       As Long
Private mstrLogTiltle                                   As String
Private mstrCurLogName                                  As String
Private mstrCurFuncs                                    As String
'存储日志数据库的连接信息，防止自动断开，用于自动重连
Private mstrLogDBCommand                                As String
Private marrCurLogSet                                   As Variant
Private mobjLogControl                                  As Scripting.Dictionary
'   日志类别
'       日志名  Array
'               LogLevel 日志等级
'               LogMode  日志方式
'               [n]
'                  部件     Scripting.Dictionary
'                  模块     Scripting.Dictionary
'                  功能     Scripting.Dictionary
'                  CallName Scripting.Dictionary


'---------------------------------------------------------------------------
'                2、属性变量与定义
'---------------------------------------------------------------------------
'CurrentLogLevel当前日志级别
'CurrentLogMode  当前日志模式

'LogFile日志文件文件路径(R)
Private mstrLogFile                     As String
'LogDB 日志存储的数据库,IP:Port/SID(R)
Private mstrLogDB                       As String
'LogUser 写日志的用户(R)
Private mstrLogUser                     As String
'OccurServer 日志记录内容发生的数据库,IP:Port/SID(R)
Private mstrOccurServer                 As String
'OccurUser 日志记录内容的操作用户(R)
Private mstrOccurUser                   As String
'OccurSID 日志记录内容的业务操作的会话SID(R)
Private mlngOccurSID                    As Long
'OccurIP 日志记录内容的发生的IP，可能存在多网卡，因此通过数据库判定(R)
Private mstrOccurIP                     As String
'KeepDays   日志保存时限(R)
Private mlngKeepDays                    As Long
'IsReady 日志是准备就绪。（R）日志打开是否符合数据库配置，且准备就绪。
Private mblnReady                       As Boolean


'CurrentLogLevel当前日志级别(R/W)
Public Property Get CurrentLogLevel() As LogLevel
    CurrentLogLevel = marrCurLogSet(Info_LogLevel)
End Property
'CurrentLogMode当前的日志模式。0-不记录日志。1-仅记录本地日志。2-仅记录数据库日志，3-本地日志和数据库日志都记录(R)
Public Property Get CurrentLogMode() As LogMode
    CurrentLogMode = marrCurLogSet(Info_LogMode)
End Property
'LogFile日志文件文件名称(R)
Public Property Get LogFile() As String
    LogFile = mstrLogFile
End Property
'LogDB 日志存储的数据库(R)
Public Property Get LogDB() As String
    LogDB = mstrLogDB
End Property
'LogUser 写日志的用户(R)
Public Property Get LogUser() As String
    LogUser = mstrLogUser
End Property
'OccurServer 日志对应业务发生的数据库(R)
Public Property Get OccurServer() As String
    OccurServer = mstrOccurServer
End Property
'OccurUser 日志记录内容的操作用户(R)
Public Property Get OccurUser() As String
    OccurUser = mstrOccurUser
End Property
'OccurSID 日志记录内容的业务操作的会话SID(R)
Public Property Get OccurSID() As Long
    OccurSID = mstrOccurUser
End Property
'OccurIP 日志记录内容的发生的IP，可能存在多网卡，因此通过数据库判定(R)
Public Property Get OccurIP() As String
    OccurIP = mstrOccurIP
End Property
'CurrentLogCount当前已经记录的行数(R)
Public Property Get CurrentLogCount() As Long
    CurrentLogCount = mlngCurLogCount
End Property
'IsLogReady 日志是准备就绪。（R）日志打开是否符合数据库配置，且准备就绪。
Public Property Get IsLogReady() As Boolean
    IsLogReady = mblnReady
End Property

'---------------------------------------------------------------------------
'                3、公共方法
'---------------------------------------------------------------------------
'@方法    SetBusinessDB
'   设置日志的位置。即业务发生的数据库信息。如果使用ZL9Comlib的公共连接，则不需要调用。在InitCommon会自动调用
'@返回值  Boolean
'   是否连接成功。第一次调用该方法时，必须在是ConnectLogDB之前调用，否则读取不到正确的日志配置。
'@参数:
'cnMain In Optional  Variant
'   业务的连接
'   1）ADODB.Connection，业务连接
'   2)Nothing 不传,此时使用gcnOracle作为业务库连接
'@备注
'    若使用单独的连接，该连接没有使用ZL9Comlib.InitCommon，则需要调用该方法，该连接业务处理完毕后，再次调用该方法切换到全局连接。
Public Function SetBusinessDB(Optional cnMain As ADODB.Connection) As Boolean
    Dim cnTmp                   As ADODB.Connection
    Dim strSQL                  As String, rsTmp            As ADODB.Recordset
    If cnMain Is Nothing Then
        Set cnTmp = gcnOracle
    Else
        Set cnTmp = cnMain
    End If
    mlngOccurSID = 0
    mstrOccurIP = ""
    mstrOccurUser = ""
    mstrOccurServer = ""
    On Error GoTo ErrH
    If Not cnTmp Is Nothing Then
        If cnTmp.State <> adStateClosed Then
            
            strSQL = "Select User, Audsid, Sid, Program, Terminal, Sys_Context('USERENV', 'IP_ADDRESS') As Ip" & vbNewLine & _
                    "From Gv$session" & vbNewLine & _
                    "Where Audsid = Userenv('SessionID')"
            Set rsTmp = mdlDataBase.OpenSQLRecord(cnTmp, strSQL, "zl9ComLib.clsLogger")
            mstrOccurUser = rsTmp!User
            mlngOccurSID = rsTmp!Audsid
            mstrOccurIP = rsTmp!IP
            mstrOccurServer = zlStr.GetServerInfo(cnTmp)
            SetBusinessDB = True
            Call LogWriteUser
        End If
    End If
    Exit Function
ErrH:
    Err.Clear
End Function

'@方法    ConnectLogDB
'   连接日志库。如果使用ZL9Comlib的公共连接，则不需要调用。在InitCommon会自动调用
'@返回值  Boolean
'   是否连接成功。该方法必须在SetBusinessLocation(只需要调用一次，但是ConnectLogDB可以调用多次)调用后调用，否则可能找不到对应的日志配置。
'@参数:
'varDataBaseInfo In Optional Variant
'   包含业务库或者日志库所在位置的信息。通过这些信息连接日志管理库。
'   支持类型：
'   1）ADODB.Connection，当前业务库，或者日志存储库的连接。业务库有ZlRegInfo.项目[日志服务器]信息，日志存储库无该信息。
'                   若业务库没有该信息时，业务库也是日志存储库。
'   2)SERVER=IP:Port/SID USER=* PASS=* TRANS=*   (TRANS为是否转换密码）
'       业务库或日志库所在位置以及连接信息。
'   3)SERVER=IP:Port/SID  或 IP:Port/SID   或 strServerName(如TestBase或ORCL)
'       业务库或日志库所在位置。使用ZLUA用户访问信息。
'   4)IsMissing(varDataBaseInfo) 不传,此时使用gcnOracle作为业务库或者日志库连接
'@备注
'ZlRegInfo.项目[日志服务器]格式：
'   1)指定日志存储服务器以及日志记录用户。zlStr.Sm4EncryptEcb("SERVER=127.0.0.1:1521/TestBase USER=ZLHIS PASS=AQA TRANS=1")
'   2)仅指定日志存储服务器。用户使用ZLUA。zlStr.Sm4EncryptEcb("SERVER=127.0.0.1:1521/TestBase")
'   3)仅指定用户。日志存储当前数据库。zlStr.Sm4EncryptEcb("USER=ZLHIS PASS=AQA TRANS=1")
'   4)都未指定。日志数据库存储当前数据库，且用户为ZLUA。zlStr.Sm4EncryptEcb("")
Public Function ConnectLogDB(Optional varDataBaseInfo As Variant) As Boolean
    Dim cnMain                  As ADODB.Connection, cnTmp  As ADODB.Connection
    Dim strCommand              As String
    Dim strCurServer            As String
    Dim strSQL                  As String, rsTmp            As ADODB.Recordset

    Const G_UA_PWD              As String = "FA74C8A530DE7E088B1ACA673DD6297D"
    Const G_UA_KEY              As String = "0016FDE250354FA9A4BA45433DBCC35D"
    
    mblnUseLog = False
    mblnUseLog = False
    marrCurLogSet = Array(LogLevel_LogOFF, LogMode_None, Nothing, Nothing, Nothing, Nothing)
    Set mobjLogControl = New Scripting.Dictionary
    mstrLogDBCommand = ""
    mstrCurLogName = ""
    mstrCurFuncs = ""
    mstrLogDB = ""
    mstrLogUser = ""
    
    mstrLogFile = ""
    mlngCurLogCount = 0
    Call ClearObject(mobjLocalLog)
    Call ClearObject(mcnDbLog)
    
    If IsObject(varDataBaseInfo) Then
        If varDataBaseInfo Is Nothing Then
            Set cnTmp = gcnOracle
            If Not cnTmp Is Nothing Then
                If cnTmp.State <> adStateClosed Then strCurServer = zlStr.GetServerInfo(cnTmp)
            End If
        Else
            If TypeOf varDataBaseInfo Is ADODB.Connection Then
                Set cnTmp = varDataBaseInfo
                strCurServer = zlStr.GetServerInfo(cnTmp)
            End If
        End If
    ElseIf varType(varDataBaseInfo) = vbString Then
        strCommand = varDataBaseInfo
        If InStr(strCommand, "=") = 0 Then
            strCommand = "SERVER=" & Trim(strCommand)
        End If
        If Not gcnOracle Is Nothing Then
            If gcnOracle.State <> adStateClosed Then
                strCurServer = zlStr.GetServerInfo(gcnOracle)
            End If
        End If
        Set cnTmp = CreateConectinByCommand(strCommand, "ZLUA", zlStr.Sm4DecryptEcb("ZLSV2:" & G_UA_PWD, GetGeneralAccountKey(G_UA_KEY)), False, strCurServer, strCurServer)
    ElseIf IsMissing(varDataBaseInfo) Then
        Set cnTmp = gcnOracle
        If Not cnTmp Is Nothing Then
            If cnTmp.State <> adStateClosed Then strCurServer = zlStr.GetServerInfo(cnTmp)
        End If
    End If
    If Not cnTmp Is Nothing Then
        If cnTmp.State <> adStateClosed Then
            On Error Resume Next
            strSQL = "Select User,Max(内容) 内容 From zlRegInfo Where 项目 = [1]"
            Set rsTmp = mdlDataBase.OpenSQLRecord(cnTmp, strSQL, "zl9ComLib.clsLogger", "日志服务器")
            '不支持多级指定日志数据库。如库A->库B->库C 这种不支持。
            '若连接的数据库没有配置日志库，则连接的数据库为日志数据库，否则指向的数据库为日志数据库。不再判断指向库的日志库配置
            If IsNull(rsTmp!内容) Then
                If strCommand = "" Then
                    strCommand = "SERVER=" & strCurServer & " USER=ZLUA PASS=" & zlStr.Sm4DecryptEcb("ZLSV2:" & G_UA_PWD, GetGeneralAccountKey(G_UA_KEY)) & " TRANS=0"
                    mstrLogUser = "ZLUA"
                    Set cnMain = CreateConectinByCommand(strCommand)
                Else
                    Set cnMain = CreateConectinByCommand(strCommand, "ZLUA", zlStr.Sm4DecryptEcb("ZLSV2:" & G_UA_PWD, GetGeneralAccountKey(G_UA_KEY)), False, strCurServer, strCurServer, mstrLogUser)
                End If
                If cnMain Is Nothing Then
                    Set cnMain = cnTmp
                    mstrLogUser = rsTmp!User & ""
                End If
            Else
                If rsTmp!内容 & "" Like "ZLSV*:*" Then
                    strCommand = zlStr.Sm4DecryptEcb(rsTmp!内容)
                Else
                    strCommand = rsTmp!内容
                End If
                Set cnMain = CreateConectinByCommand(strCommand, "ZLUA", zlStr.Sm4DecryptEcb("ZLSV2:" & G_UA_PWD, GetGeneralAccountKey(G_UA_KEY)), False, strCurServer, strCurServer, mstrLogUser)
            End If
            If Not cnMain Is Nothing Then
                ConnectLogDB = cnMain.State <> adStateClosed
                If ConnectLogDB Then
                    mstrLogDB = strCurServer
                    ConnectLogDB = ReadLogSet(cnMain, mstrOccurUser, mstrOccurIP)
                    If ConnectLogDB Then
                        mblnUseLog = mobjLogControl.Count > 0
                        mblnReady = True
                        mstrLogDBCommand = strCommand
                    Else
                        mstrLogDB = ""
                    End If
                End If
            End If
        End If
    End If
End Function

'@方法    Log
'   记录指定日志级别的日志。
'@返回值  Boolean
'
'@参数:
'strLogName In String
'   日志的业务分类名称。如一卡通日志等。传""时取最近一次不为空
'strComponentName In strComponentName
'   日志发生的部件名称。使用App.EXEName。传""时取最近一次不为空
'strModule In String
'   日志发生的模块。可以是ZLHIS体系的模块可以是VB的模块等。传""时取最近一次不为空
'strFuncName In String
'   日志的发生的功能名。或者发生的VB函数。传""时取最近一次不为空
'llLogLevel In LogLevel
'   当前记录日志性质。
'   LogLevel_Error：该日志是属于业务错误或者VB错误、程序错误等，会影响程序运行的。
'   LogLevel_Warn：该日志不属于错误，不影响程序运行，但是可能造成流程变动或者程序功能不全。可能与数据控制或者当前环境相关，如缺失某部件，仍然可以继续使用，但是对应功能缺失。
'   LogLevel_Info：该日志属于重要信息记录，用于重要信息的记录，如费用，交易等数据。
'   LogLevel_Trace：该信息是程序的运行跟踪信息。用于跟踪程序运行，以便方便错误查证。
'arrPars In ParamArray
'   产生格式：strLogTilte: arrPars(0)=arrPars(1)  arrPars(2)=arrPars(3) ...arrPars(2n)=arrPars(2n+1)
'@备注
'
Public Function Log(strLogName As String, strComponentName As String, strModule As String, strFuncName As String, ByVal llLogLevel As LogLevel, strLogTilte As String, ParamArray arrPars() As Variant) As Boolean
    Dim strText         As String
    Dim arrTmp          As Variant
    Dim blnCanLog       As Boolean
    
    On Error GoTo ErrH
    If mblnUseLog Then
        If mstrCurLogName <> strLogName And LenB(strLogName) <> 0 Then Call ChangeCurrentLog(strLogName)
        If Me.CurrentLogLevel >= LogLevel_Info Then
            If LenB(strComponentName) = 0 Or LenB(strModule) = 0 Or LenB(strFuncName) = 0 Or LenB(strLogName) = 0 Then
                If LenB(strLogName) = 0 Then strLogName = mstrCurLogName
                arrTmp = Split(mstrCurFuncs & "...", ".")
                If LenB(strComponentName) = 0 Then strComponentName = arrTmp(0)
                If LenB(strModule) = 0 Then strModule = arrTmp(1)
                If LenB(strFuncName) = 0 Then strFuncName = arrTmp(2)
            End If
            blnCanLog = IsCanWriteLog(UCase(strComponentName), UCase(strModule), UCase(strFuncName), "")
            
            If blnCanLog Then
                mstrParas = ""
                For mlngIndex = LBound(arrPars) To UBound(arrPars) Step 2
                    If mlngIndex = UBound(arrPars) Then
                        mstrParas = mstrParas & zlStr.DisPlayOneValue(arrPars(mlngIndex), Me.CurrentLogLevel > LogLevel_All)
                    Else
                        mstrParas = mstrParas & zlStr.DisPlayOneValue(arrPars(mlngIndex), Me.CurrentLogLevel > LogLevel_All) & "=" & zlStr.DisPlayOneValue(arrPars(mlngIndex + 1), Me.CurrentLogLevel > LogLevel_All) & "    "
                    End If
                Next
    
                If Len(mstrParas) <> 0 Then
                    mstrParas = strLogTilte & ":  " & mstrParas
                Else
                    mstrParas = mstrParas
                End If
    
                If Me.CurrentLogMode = LogMode_OnlyLocal Or Me.CurrentLogMode = LogMode_Both Then
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
                    mobjLocalLog.WriteLine mstrLogTiltle & "┣━╋━" & mstrParas
                End If
    
                If Me.CurrentLogMode = LogMode_OnlyDB Or Me.CurrentLogMode = LogMode_Both Then
                    Call WriteDBLog(llLogLevel, strLogName, strComponentName, strModule, strFuncName, mstrParas)
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
        If Me.CurrentLogLevel >= LogLevel_Error Then
            If blnCanLog Then
                If Me.CurrentLogMode = LogMode_OnlyLocal Or Me.CurrentLogMode = LogMode_Both Then
                    Call GetLogTiltle(LogLevel_Error)
                    mobjLocalLog.WriteLine mstrLogTiltle & "┣━╋━╋Log发生错误：" & Err.Number & "-" & Err.Description
                End If
                If Me.CurrentLogMode = LogMode_OnlyDB Or Me.CurrentLogMode = LogMode_Both Then
                    Call WriteDBLog(LogLevel_Error, strLogName, strComponentName, strModule, strFuncName, "Log发生错误：" & Err.Number & "-" & Err.Description)
                End If
            End If
        End If
    End If
    Err.Clear
End Function

'@方法    LogCall
'   记录重要接口的调用情况。如WebAPI或者存储过程。该日志的性质属于LogLevel_Info（重要信息级别）
'@返回值  Boolean
'
'@参数:
'@参数:
'strLogName In String
'   日志的业务分类名称。如一卡通日志等。传""时取最近一次不为空
'strComponentName In strComponentName
'   日志发生的部件名称。使用App.EXEName。传""时取最近一次不为空
'strModule In String
'   日志发生的模块。可以是ZLHIS体系的模块可以是VB的模块等。传""时取最近一次不为空
'strFuncName In String
'   日志的发生的功能名。或者发生的VB函数。传""时取最近一次不为空
'strCallName In String
'   WebAPI名称或者存储过程名称
'lcsCurentLogCallState In LogCallState
'   标识调用的时机。开始调用或者结束调用。
'arrPars In ParamArray
'   产生格式： arrPars(0),arrPars(1),...,arrPars(n)
'@备注
'
Public Function LogCall(strLogName As String, strComponentName As String, strModule As String, strFuncName As String, strCallName As String, lcsCurentLogCallState As LogCallState, ParamArray arrPars() As Variant) As Boolean
    Dim blnNormal   As Boolean
    Dim i           As Long
    Dim objCommand  As ADODB.Command
    Dim blnWrite    As Boolean
    Dim arrTmp      As Variant
    Dim blnCanLog   As Boolean

    On Error GoTo ErrH
    If mblnUseLog Then
        If mstrCurLogName <> strLogName And LenB(strLogName) <> 0 Then Call ChangeCurrentLog(strLogName)
        If Me.CurrentLogLevel >= LogLevel_Info Then
            If LenB(strComponentName) = 0 Or LenB(strModule) = 0 Or LenB(strFuncName) = 0 Or LenB(strLogName) = 0 Then
                If LenB(strLogName) = 0 Then strLogName = mstrCurLogName
                arrTmp = Split(mstrCurFuncs & "...", ".")
                If LenB(strComponentName) = 0 Then strComponentName = arrTmp(0)
                If LenB(strModule) = 0 Then strModule = arrTmp(1)
                If LenB(strFuncName) = 0 Then strFuncName = arrTmp(2)
            End If
            blnCanLog = IsCanWriteLog(UCase(strComponentName), UCase(strModule), UCase(strFuncName), UCase(strCallName))
            If blnCanLog Then
                mstrParas = ""
                For mlngIndex = 0 To UBound(arrPars)
                    blnNormal = True
                    If IsObject(arrPars(mlngIndex)) Then
                        If TypeOf arrPars(mlngIndex) Is ADODB.Parameters Then
                            blnNormal = False
                            Set objCommand = arrPars(mlngIndex)
                            If mlngIndex = 0 Then
                                mstrParas = "{ADODB.Parameters("
                            Else
                                mstrParas = mstrParas & "," & "{ADODB.Parameters("
                            End If
                            For i = 0 To objCommand.Parameters.Count - 1
                                If i = 0 Then
                                    mstrParas = mstrParas & zlStr.DisPlayOneValue(objCommand.Parameters(i).Value, Me.CurrentLogLevel >= LogLevel_All)
                                Else
                                    mstrParas = mstrParas & "," & zlStr.DisPlayOneValue(objCommand.Parameters(i).Value, Me.CurrentLogLevel >= LogLevel_All)
                                End If
                            Next
                            mstrParas = mstrParas & ")}"
                        End If
                    End If
                    If blnNormal Then
                        If mlngIndex = 0 Then
                            mstrParas = zlStr.DisPlayOneValue(arrPars(mlngIndex), Me.CurrentLogLevel >= LogLevel_All)
                        Else
                            mstrParas = mstrParas & "," & zlStr.DisPlayOneValue(arrPars(mlngIndex), Me.CurrentLogLevel >= LogLevel_All)
                        End If
                    End If
                Next
                If Me.CurrentLogMode = LogMode_OnlyLocal Or Me.CurrentLogMode = LogMode_Both Then
                    Call GetLogTiltle(LogLevel_Info)
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
                    If lcsCurentLogCallState = LogCallState_CallBegin Then
                        mobjLocalLog.WriteLine mstrLogTiltle & "┣━╋━┳Begin " & strCallName & " Using " & mstrParas
                    Else
                        mobjLocalLog.WriteLine mstrLogTiltle & "┣━╋━┻End " & strCallName & " Return " & mstrParas
                    End If
                    blnWrite = True
                End If
    
                If Me.CurrentLogMode = LogMode_OnlyDB Or Me.CurrentLogMode = LogMode_Both Then
                    Call WriteDBLog(LogLevel_Info, strLogName, strComponentName, strModule, strFuncName, mstrParas, strCallName, lcsCurentLogCallState)
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
        If Me.CurrentLogLevel >= LogLevel_Error Then
            If blnCanLog Then
                If Me.CurrentLogMode = LogMode_OnlyLocal Or Me.CurrentLogMode = LogMode_Both Then
                    Call GetLogTiltle(LogLevel_Error)
                    If Not blnWrite And lcsCurentLogCallState = LogCallState_CallBegin Then
                        mobjLocalLog.WriteLine mstrLogTiltle & "┣━╋━┳Begin " & strCallName
                    End If
                    mobjLocalLog.WriteLine mstrLogTiltle & "┣━╋━╋LogCall发生错误：" & Err.Number & "-" & Err.Description
                    If Not blnWrite And lcsCurentLogCallState <> LogCallState_CallBegin Then
                        mobjLocalLog.WriteLine mstrLogTiltle & "┣━╋━┻End " & strCallName
                    End If
                End If
                If Me.CurrentLogMode = LogMode_OnlyDB Or Me.CurrentLogMode = LogMode_Both Then
                    If Not blnWrite And lcsCurentLogCallState = LogCallState_CallBegin Then
                        Call WriteDBLog(LogLevel_Info, strLogName, strComponentName, strModule, strFuncName, "发生错误！" & mstrParas, strCallName, lcsCurentLogCallState)
                    End If
                    Call WriteDBLog(LogLevel_Error, strLogName, strComponentName, strModule, strFuncName, "LogCall发生错误：" & Err.Number & "-" & Err.Description, strCallName)
                    If Not blnWrite And lcsCurentLogCallState <> LogCallState_CallBegin Then
                        Call WriteDBLog(LogLevel_Info, strLogName, strComponentName, strModule, strFuncName, "发生错误！" & mstrParas, strCallName, lcsCurentLogCallState)
                    End If
                End If
            End If
        End If
    End If
    Err.Clear
End Function

'---------------------------------------------------------------------------
'                4、私有方法
'---------------------------------------------------------------------------

'@方法    GetLogTiltle
'   获取日志头，并进行行控制
'@返回值
Private Sub GetLogTiltle(ByVal llLogLevel As LogLevel)
    If mlngCurLogCount >= mlngMaxLogCount Or mlngCurLogCount = 0 Then
        Call LogWriteHeader
    End If
    mlngCurLogCount = mlngCurLogCount + 1
    mstrLogTiltle = mstrLogLevelName(llLogLevel + 1) & Format(Now, "YYYY-MM-DD hh:mm:ss")
End Sub
'@方法    LogWriteHeader
'   书写日志头
'@返回值  Boolean
Private Sub LogWriteHeader()
    Dim strLogFile      As String
    If mlngCurLogCount <> 0 Then
        On Error Resume Next
        mobjLocalLog.Close
        Set mobjLocalLog = Nothing
        strLogFile = GetLogName
        Set mobjLocalLog = gobjLogFile.OpenTextFile(strLogFile, ForAppending, True)
        mstrLogFile = strLogFile
        mlngCurLogCount = 0
        If Not mobjLocalLog Is Nothing Then
            mobjLocalLog.WriteLine String(80, "-")
            mobjLocalLog.WriteLine "[Create Time]  " & Format(Now, "yyyy-mm-dd  hh:mm:ss")
            mobjLocalLog.WriteLine "[Environment]  ProcessID=" & OS.CurrentPID & " ProcessUser=" & OS.ProcessUserFullName & " OSSessionID=" & OS.SessionID & " OSSessionUser=" & OS.SessionUserFullName
            Call LogWriteUser(True)
            mobjLocalLog.WriteLine String(80, "-")
        End If
    Else
        If Not mobjLocalLog Is Nothing Then
            mlngCurLogCount = mlngCurLogCount + 1
            mobjLocalLog.WriteLine String(80, "-")
            mobjLocalLog.WriteLine "[Create Time]  " & Format(Now, "yyyy-mm-dd  hh:mm:ss")
            mobjLocalLog.WriteLine "[Environment]  ProcessID=" & OS.CurrentPID & " ProcessUser=" & OS.ProcessUserFullName & " OSSessionID=" & OS.SessionID & " OSSessionUser=" & OS.SessionUserFullName
            Call LogWriteUser(True)
            mobjLocalLog.WriteLine String(80, "-")
            
        End If
    End If
End Sub

'@方法    LogWriteUser
'   写当前业务库的信息
'@返回值
'
'@参数:
'Param 方向
'
'@备注
'
Private Sub LogWriteUser(Optional ByVal blnFirst As Boolean)
    If Not blnFirst Then
        Call GetLogTiltle(LogLevel_Info)
        If LenB(mstrCurLogName) <> 0 Then
            If LenB(mstrCurFuncs) <> 0 Then
                mobjLocalLog.WriteLine String(Len(mstrLogTiltle), " ") & "┣━┻" & mstrCurFuncs
            End If
            mobjLocalLog.WriteLine String(Len(mstrLogTiltle), " ") & "┗" & mstrCurLogName
        End If
    End If
    mstrCurLogName = ""
    mstrCurFuncs = ""
    If Not mobjLocalLog Is Nothing Then
        If Not blnFirst Then
            mobjLocalLog.WriteLine String(80, "-")
            mobjLocalLog.WriteLine "[BusDBChange]  " & Format(Now, "yyyy-mm-dd  hh:mm:ss")
        End If
        mobjLocalLog.WriteLine "[Business DB]  User=" & mstrOccurUser & " Server=" & mstrOccurServer & " SID=" & mlngOccurSID & " IP=" & mstrOccurIP
        mobjLocalLog.WriteLine "[Log      DB]  User=" & mstrLogUser & " Server=" & mstrLogDB
        If Not blnFirst Then
            mobjLocalLog.WriteLine String(80, "-")
        End If
    End If
End Sub
'@方法    WriteDBLog
'   将日志写入数据库
'@返回值
'
'@参数:
'llCurent LogLevel(In)
'   当前日志级别
'strProcedureName        String(In)
'   记录的存储过程名
'strParas        String
'   存储过程参数
'lngExecutePhase Long(In)
'   执行阶段。0-开始阶段。1-结束阶段
'@备注
'
Private Sub WriteDBLog(llCurent As LogLevel, strLogName As String, strComponentName As String, strModule As String, strFuncName As String, strParasInfo As String, Optional strCallName As String, Optional ByVal lcsCurentLogCallState As LogCallState = -1)
    On Error GoTo ErrH
    If LenB(strParasInfo) < 8000 Then
        Call mdlDataBase.CallProcedure(mcnDbLog, "Zlloginfo_Insert", "WriteDBLog", False, llCurent, mstrOccurServer, mstrOccurUser, mlngOccurSID, mstrOccurIP, mstrStation, mlngPID, mstrProcessName, strLogName, strComponentName, strModule, strFuncName, strCallName, IIf(lcsCurentLogCallState < 0, Null, lcsCurentLogCallState), strParasInfo)
    Else
        Call mdlDataBase.CallProcedure(mcnDbLog, "Zlloginfo_Insert", "WriteDBLog", False, llCurent, mstrOccurServer, mstrOccurUser, mlngOccurSID, mstrOccurIP, mstrStation, mlngPID, mstrProcessName, strLogName, strComponentName, strModule, strFuncName, strCallName, IIf(lcsCurentLogCallState < 0, Null, lcsCurentLogCallState), "", strParasInfo)
    End If
    Exit Sub
ErrH:
    Err.Clear
End Sub

'@方法    ChangeCurrentLog
'   切换当前日志
'@返回值
'
'@参数:
'strLogName In String
'   当前日志名称
'@备注
'
Private Sub ChangeCurrentLog(strLogName As String)
    marrCurLogSet = mobjLogControl(strLogName)
    If IsNull(marrCurLogSet) Or IsEmpty(marrCurLogSet) Then
        marrCurLogSet = Array(LogLevel_LogOFF, LogMode_None, Nothing, Nothing, Nothing, Nothing)
    End If
End Sub


'@方法    IsCanWriteLog
'   判断是否可以写日志
'@返回值  Boolean
'
'@参数:
'Param 方向
'
'@备注
'
Private Function IsCanWriteLog(strComponentName As String, strModule As String, strFuncName As String, strCallName As String) As Boolean
    Dim lngStart    As Long, objDic             As Scripting.Dictionary
    Dim varRet      As Variant, blnCurCheck     As Boolean
    
    On Error GoTo ErrH
    For lngStart = Info_LogMode + 1 To UBound(marrCurLogSet) Step 4
        blnCurCheck = True
        Set objDic = marrCurLogSet(lngStart + Info_Component)
        If Not objDic Is Nothing Then
            varRet = objDic(strComponentName)
            If IsNull(varRet) Or IsEmpty(varRet) Then blnCurCheck = False
        End If
        If blnCurCheck Then
            Set objDic = marrCurLogSet(lngStart + Info_Module)
            If Not objDic Is Nothing Then
                varRet = objDic(strModule)
                If IsNull(varRet) Or IsEmpty(varRet) Then blnCurCheck = False
            End If
        End If
        
        If blnCurCheck Then
            Set objDic = marrCurLogSet(lngStart + Info_Function)
            If Not objDic Is Nothing Then
                varRet = objDic(strFuncName)
                If IsNull(varRet) Or IsEmpty(varRet) Then blnCurCheck = False
            End If
        End If
        
        If blnCurCheck Then
            Set objDic = marrCurLogSet(lngStart + Info_Call)
            If Not objDic Is Nothing Then
                varRet = objDic(strCallName)
                If IsNull(varRet) Or IsEmpty(varRet) Then blnCurCheck = False
            End If
        End If
        '整个检查过来都符合，则直接退出。不再检查下一组配置。
        '若当前组配置不符合，只有继续检查下一组
        If blnCurCheck Then
            IsCanWriteLog = True
            Exit For
        End If
    Next
    Exit Function
ErrH:
    Err.Clear
End Function

'@方法    CreateConectinByCommand
'   通过命令行获取连接
'@返回值  ADODB.Connection
'
'@参数:
'strCommand In/Out String
'   格式：SERVER=127.0.0.1:1521/TestBase USER=ZLHIS PASS=AQA TRANS=0
'   如果命令行不完整，会自动补充并返回
'strDefaultUser In String
'   没有用户时缺省的用户
'strDefaultPWD In String
'   没有密码时缺省用户的密码
'blnDefaultTrans In Boolean
'   缺省密码是否需要转换
'strDefaultServer In String
'   没有服务器时，缺省的服务器。IP:Port/SID 或这TNSName
'strOutCommand In String
'   返回命令行，方便断网重新连接
'@备注
'
Private Function CreateConectinByCommand(strCommand As String, Optional ByVal strDefaultUser As String, Optional ByVal strDefaultPWD As String, Optional ByVal blnDefaultTrans As Boolean, Optional ByVal strDefaultServer As String, Optional strOutSerrver As String, Optional strOutUser As String) As ADODB.Connection
    Dim strServer   As String, strUser      As String, strPass  As String, blnTrans As Boolean
    Dim arrTmp      As Variant, i           As Long
    
    arrTmp = Split(strCommand, " ")
    strCommand = ""
    strOutSerrver = ""
    For i = LBound(arrTmp) To UBound(arrTmp)
        If Trim(arrTmp(i)) <> "" Then
            If arrTmp(i) Like "USER=*" Then
                strUser = Mid(arrTmp(i), Len("USER=*"))
            ElseIf arrTmp(i) Like "PASS=*" Then
                strPass = Mid(arrTmp(i), Len("PASS=*"))
            ElseIf arrTmp(i) Like "TRANS=*" Then
                blnTrans = Val(Mid(arrTmp(i), Len("TRANS=*"))) = 1
            ElseIf arrTmp(i) Like "SERVER=*" Then
                strServer = Mid(arrTmp(i), Len("SERVER=*"))
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
        blnTrans = blnDefaultTrans
    End If
    
    If gobjRegister Is Nothing Then
        On Error Resume Next
        Set gobjRegister = GetObject("", "zlRegister.clsRegister")
        Err.Clear
        On Error GoTo ErrH
    End If
    
    If Not gobjRegister Is Nothing Then
        If strPass <> "" Then
            Set CreateConectinByCommand = gobjRegister.GetConnection(strServer, strUser, strPass, blnTrans, 1, "", False)
            If CreateConectinByCommand.State = adStateClosed Then
                Set CreateConectinByCommand = Nothing
            Else
                strOutSerrver = zlStr.GetServerInfo(strServer)
                strCommand = "SERVER=" & strServer & " USER=" & strUser & " PASS=" & strPass & " TRANS=" & IIf(blnTrans, 1, 0)
                strOutUser = strUser
            End If
        End If
    End If
    Exit Function
ErrH:
    Err.Clear
End Function

'@方法    GetGeneralAccountKey
'   密钥转换
'@返回值  String
'
'@参数:
'@备注
'
Private Function GetGeneralAccountKey(ByRef strKey As String) As String
    Dim arrTmp()    As Byte
    Dim i           As Long
    arrTmp = zlStr.HexStringToByte(strKey, 16)
    For i = LBound(arrTmp) To UBound(arrTmp)
        If i Mod 2 = 0 Then
            arrTmp(i) = 255 - arrTmp(i)
        ElseIf i Mod 3 = 0 Then
            arrTmp(i) = (arrTmp(i) + i) Mod 256
        End If
    Next
    GetGeneralAccountKey = zlStr.ByteToHexString(arrTmp)
End Function

'@方法    ReadLogSet
'   读取日志配置
'@返回值  Boolean
'
'@参数:
'Param 方向
'
'@备注
'
Private Function ReadLogSet(cnLogDB As ADODB.Connection, ByVal strUser As String, ByVal strIp As String) As Boolean
    Dim strSQL          As String, rsTmp        As ADODB.Recordset
    Dim strPreName      As String
    Dim blnLocalLog     As Boolean, blnDBLog    As Boolean
    Dim arrSets()       As Variant, lngStart    As Long
    Dim strLogFile      As String, lngKeepDays  As Long, dtCurrent      As Date
    
    On Error GoTo ErrH
    strSQL = "Select c.Id, c.Name, c.Log_Keep_Days, c.Log_Mode, c.Log_Level, c.Component_Names, c.Module_Names, c.Function_Names," & vbNewLine & _
            "       c.Call_Names, c.Sortid" & vbNewLine & _
            "From (Select a.Id, a.Name, a.Log_Keep_Days, a.Log_Mode, a.Log_Level, b.Station, b.Ip, b.User_Name, b.Component_Names," & vbNewLine & _
            "              b.Module_Names, b.Function_Names, b.Call_Names," & vbNewLine & _
            "              Decode(b.Station, Null, 0, 1) + Decode(b.User_Name, Null, 0, 1) + Decode(b.Ip, Null, 0, 1) Sortid" & vbNewLine & _
            "       From Zllogcategory A, Zllogset B" & vbNewLine & _
            "       Where a.Id = b.Category_Id(+) And Sysdate Between a.Begin_Time And a.End_Time And a.Log_Mode > 0 And" & vbNewLine & _
            "             a.Log_Level > 0 And Nvl(Station, [1]) =[1] And Nvl(User_Name, [2]) =[2] And Nvl(Ip, [3]) =[3]) C" & vbNewLine & _
            "Order By c.Id, c.Sortid Desc"
    Set rsTmp = mdlDataBase.OpenSQLRecord(cnLogDB, strSQL, "ReadLogSet", mstrStation, strUser, strIp)
    Do While Not rsTmp.EOF
        If strPreName <> rsTmp!Name Then
            If strPreName <> "" Then
                mobjLogControl.Add strPreName, arrSets
            End If
            strPreName = rsTmp!Name & ""
            ReDim arrSets(5)
            arrSets(Info_LogLevel) = Val(rsTmp!Log_Level & "")
            arrSets(Info_LogMode) = Val(rsTmp!Log_Mode & "")
            lngStart = Info_LogMode + 1
            Set arrSets(lngStart + Info_Component) = AnalyseSet(UCase(rsTmp!Component_Names & ""))
            Set arrSets(lngStart + Info_Function) = AnalyseSet(UCase(rsTmp!Function_Names & ""))
            Set arrSets(lngStart + Info_Module) = AnalyseSet(UCase(rsTmp!Module_Names & ""))
            Set arrSets(lngStart + Info_Call) = AnalyseSet(UCase(rsTmp!Call_Names & ""))
            If arrSets(Info_LogMode) = LogMode_OnlyLocal Then
                blnLocalLog = True
            End If

            If arrSets(Info_LogMode) = LogMode_OnlyDB Then
                blnDBLog = True
            End If

            If arrSets(Info_LogMode) = LogMode_Both Then
                blnDBLog = True
                blnLocalLog = True
            End If
        Else
            If rsTmp!Sortid <> 0 Then    '有其他配置
                lngStart = UBound(arrSets) + 1
                ReDim Preserve arrSets(UBound(arrSets) + 4)
                Set arrSets(lngStart + Info_Component) = AnalyseSet(UCase(rsTmp!Component_Names & ""))
                Set arrSets(lngStart + Info_Function) = AnalyseSet(UCase(rsTmp!Function_Names & ""))
                Set arrSets(lngStart + Info_Module) = AnalyseSet(UCase(rsTmp!Module_Names & ""))
                Set arrSets(lngStart + Info_Call) = AnalyseSet(UCase(rsTmp!Call_Names & ""))
            End If
        End If
        rsTmp.MoveNext
    Loop
    
    If strPreName <> "" Then
        mobjLogControl.Add strPreName, arrSets
    End If
    '获取当前时间，和最大的日志保留天数，根据这些信息清理本地日志。数据库日志通过自动作业处理。
    strSQL = "Select Max(a.Log_Keep_Days) Log_Keep_Days, Sysdate Cur" & vbNewLine & _
            "From Zllogcategory A" & vbNewLine & _
            "Where Sysdate Between a.Begin_Time And a.End_Time And a.Log_Mode > 0 And a.Log_Level > 0"
    Set rsTmp = mdlDataBase.OpenSQLRecord(cnLogDB, strSQL, "ReadLogSet")
    lngKeepDays = Val(rsTmp!Log_Keep_Days & "")
    dtCurrent = rsTmp!Cur
    Call ClearLocalLog(lngKeepDays, dtCurrent)
    
    If blnLocalLog Then
        strLogFile = GetLogName
        Set mobjLocalLog = gobjLogFile.OpenTextFile(strLogFile, ForAppending, True)
        Call LogWriteHeader
        mstrLogFile = strLogFile
    Else
        Call ClearObject(mobjLocalLog)
    End If
    If blnDBLog Then
        Set mcnDbLog = cnLogDB
    ElseIf Not cnLogDB Is gcnOracle Then
        Call ClearObject(cnLogDB)
    End If
    mblnUseLog = blnDBLog Or blnLocalLog
    ReadLogSet = True
    Exit Function
ErrH:
    Err.Clear
End Function

'@方法    GetLogName
'   生成日志名称
'@返回值  String
'
'@参数:
'@备注
'
Private Function GetLogName() As String
    Dim strLogFile      As String, arrTmp       As Variant, i           As Long
    Dim strTmp          As String
    On Error GoTo ErrH
    '一、生成日志文件名。进程名_S会话ID P进程ID _YYYYMMDD
    strLogFile = OS.AppsoftPath & "\Log\日志跟踪\" & gobjLogFile.GetBaseName(OS.StartExePath) & "_S" & OS.SessionID & "P" & OS.CurrentPID
    strLogFile = strLogFile & "_" & Format(Now, "yyyymmdd") & ".log"
    '依次判断文件路径是否存在
    If gobjLogFile.FolderExists(gobjLogFile.GetParentFolderName(strLogFile)) Then
        arrTmp = Split(strLogFile, "\")
        strTmp = arrTmp(0) & "\"
        For i = 1 To UBound(arrTmp) - 1
            strTmp = strTmp & arrTmp(i) & "\"
            If Not gobjLogFile.FolderExists(strTmp) Then
                Call gobjLogFile.CreateFolder(strTmp)
            End If
        Next
    End If
    GetLogName = strLogFile
    Exit Function
ErrH:
    Err.Clear
End Function

'@方法    ClearLocalLog
'
'@返回值
'
'@参数:
'lngKeepDays In Long
'   日志保留天数
'dtCurrent In Date
'   当前时间
'@备注
'
Private Sub ClearLocalLog(ByVal lngKeepDays As Long, ByVal dtCurrent As Date)
    Dim objFile     As File
    Dim strMin      As String
    Dim strCur      As String
    Dim arrFile     As Variant, i       As Long
    
    On Error GoTo ErrH
    If lngKeepDays <= 0 Then lngKeepDays = 7
    strMin = Format(DateAdd("d", lngKeepDays, dtCurrent), "yyyymmdd") & ".LOG"
    arrFile = Array()
    For Each objFile In gobjLogFile.GetFolder(OS.AppsoftPath & "\Log\日志跟踪").Files
        If UCase(objFile.Name) Like "*_S*P*_.LOG" Then
            strCur = Mid(objFile.Name, InStrRev(objFile.Name, "_") + 1)
            If strMin > strCur Then
                ReDim Preserve arrFile(UBound(arrFile) + 1)
                arrFile(UBound(arrFile)) = objFile.Path
            End If
        End If
    Next
    For i = 0 To UBound(arrFile)
        On Error Resume Next
        Call gobjLogFile.DeleteFile(arrFile(i), True)
        If Err.Number <> 0 Then Err.Clear
        On Error GoTo ErrH
    Next
    Exit Sub
ErrH:
    Err.Clear
End Sub

'@方法    AnalyseSet
'   分析配置，返回配置字典
'@返回值  Scripting.Dictionary
'
'@参数:
'strSets In String
'   配置字符串
'@备注
'
Private Function AnalyseSet(strSets As String) As Scripting.Dictionary
    Dim objSet          As Scripting.Dictionary
    Dim arrTmp          As Variant, i           As Long
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

Private Sub ClearObject(objInput As Object)
    On Error Resume Next
    If Not objInput Is Nothing Then
        objInput.Close
        Set objInput = Nothing
    End If
End Sub
'---------------------------------------------------------------------------
'                5、对象方法与事件
'---------------------------------------------------------------------------
Private Sub Class_Initialize()
    mstrLogLevelName(LogLevel_UnDefined + 1) = "[UnDefine]"
    mstrLogLevelName(LogLevel_LogOFF + 1) = "[ LogOFF ]"
    mstrLogLevelName(LogLevel_Error + 1) = "[  Error ]"
    mstrLogLevelName(LogLevel_Warn + 1) = "[  Warn  ]"
    mstrLogLevelName(LogLevel_Info + 1) = "[  Info  ]"
    mstrLogLevelName(LogLevel_Trace + 1) = "[  Trace ]"
    mstrLogLevelName(LogLevel_All + 1) = "[   All  ]"
    mlngMaxLogCount = 2 ^ 31 - 1
    mstrProcessName = gobjLogFile.GetBaseName(OS.StartExePath)
    mlngPID = OS.CurrentPID
    mstrStation = OS.ComputerName
    marrCurLogSet = Array(LogLevel_LogOFF, LogMode_None, Nothing, Nothing, Nothing, Nothing)
    Set mobjLogControl = New Scripting.Dictionary
    If SetBusinessDB(gcnOracle) Then
        Call ConnectLogDB(gcnOracle)
    End If
End Sub
