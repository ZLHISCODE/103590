VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRunScript"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'**************************************************************************
'                一、模块级变量
'**************************************************************************
'==========================================================================
'1、模块变量
'==========================================================================

Private mobjText As TextStream '数据流对象
Private mobjFSO As New FileSystemObject '用来进行文件操作
Private mstrBuild As New clsStringBulider
Private marrDefine() As SQL_DEFINE '替换变量定义
Private mstrFMTSQL As String '规范化的SQL
Private mfrmParent As Object '父窗体
Private mcllStrSQL As Collection '根据字符串解析出的SQL
Private mintStrSQLIdx As Integer '当前SQL位置
'==========================================================================
'                2、枚举、结构体
'==========================================================================
Private Type USER_LIST
    SysUser As String
    SysPass As String
    ToolsPass As String 'ToolsUSer 为zltools
    DBAUser As String
    DBAPass As String
End Type

'***************************************************************************
'                二、保存类的各种属性的模块级变量，以及属性定义
'***************************************************************************
'===========================================================================
'                1、保存属性的变量
'===========================================================================
'EOF属性(R)
Private mblnEOF As Boolean  '是否是文件结尾，该属性标识文件已经没有SQL可读取的状态
'SQLInfo属性(R)
Private mobjSQL As clsSQLInfo   '用来保存当前SQL的信息
'LinesCount属性(R)
Private mlngLinesCount As Long '当前文件总行数
'Line属性(R)
Private mlngLine As Long '当前文本流对应的行号
'File属性(R/w)
Private mstrFile As String  '当前文本的文件路径名
'ProcessValue:进度条位置(R)
Private mdblProcess As Double
'Abort属性(R/W)
Private mlngAbort As Long '从指定的行号开始执行文件
'FormatSQL属性(R)

'ErrCount错误数(R/W)
Private mlngErrCount As Long '出错次数
'LogFile日志文件名(R/W)
Private mstrLogFile As String
'Log日志文件(R)
Private mobjLog As TextStream '日志文件
'SysNo系统号(R/W)
Private mlngSys As Long '系统号
'IsMaskAllErr错误忽略(R/W)
Private mblnMaskAllErr As Boolean
'SQLRecTime 记录SQL的时限(R/W)
Private mlngSQLRecTime As Long
'Connection连接(R/W)
Private mcnThis As ADODB.Connection
'ProcMode存储过程选项(R/W) 0-只执行脚本，1-搜集存储过程，2-执行存储过程
Private mintProcMode As Integer
'Procedures(R/W) 搜集到的存储过程或要执行的存储过程,1=过程之间以?分割,2=过程之间以,分割
Private mstrProcedures As String
'ConnectType(R/W) 连接类型 0=应用系统连接，1=管理工具连接，2=DBA连接
Private mintConnectType As Integer
'Server 服务器名(R/W)
Private mstrServer As String
'IsRecErrSQL 记录错误SQL(R/W)
Private mblnRecErrSQL As Boolean
'AbortInfo 中断信息(R/W)
Private mstrAbortInfo As String
'UserList(R/W)当前操作过程的用户密码列表
Private mulCurrent As USER_LIST
'GatherTables(R/W) 统计信息收集搜集的的表,以分号分割表名
Private mstrGatherTables As String
'IsGather(R/W) 是否进行统计信息搜集
Private mblnGratherSta As Boolean
'IsUseLog(R/W) 是否记录日志
Private mblnUseLog As Boolean
'IsMaskIgnore(R/W)是否忽略可忽略错误
Private mblnMaskIgnore As Boolean
'Roles(R/W)从脚本中收集到的角色授权集合
Private mcllRoles As Collection
'IsRoleCollect (R/W)是否进行角色收集
Private mblnRoleCollect As Boolean
'IsOnlineIndex:是否索引操作使用在线模式。索引在线操作包括：
'                    Create Index * online;
'                    ALter Table ADD Constraint *Primary Key* Using Index * online；
'                    ALter Table ADD Constraint *Unique* Using Index * online；
'                    alter Index * Rebuild OnLine;
Private mblnOnlineIndex As Boolean
Private mintSection As Integer
'SectionNumber:脚本数值
'   初始打开一个新脚本，该属性为False,脚本分为数据结构修正，数据修正等。
'   每执行过一个区段，该数值+1
'===========================================================================
'                2、属性定义
'===========================================================================
'EOF是否达到文件结尾
Public Property Get EOF() As Boolean
    EOF = mblnEOF
End Property
'SQLInfo SQL对象
Public Property Get SQLInfo() As clsSQLInfo
    Set SQLInfo = mobjSQL
End Property
'LinesCount文件总行数
Public Property Get LinesCount() As Long
    LinesCount = mlngLinesCount
End Property
'Line文件当前行号
Public Property Get Line() As Long
    Line = mlngLine
End Property
'SectionNumber当前脚本位置的区段数值
Public Property Get SectionNumber() As Integer
    SectionNumber = mintSection
End Property
'File文件名
Public Property Get File() As String
    File = mstrFile
End Property

Public Property Let File(ByVal strFile As String)
    mstrFile = strFile
End Property
'ProcessValue:进度条位置
Public Property Get ProcessValue() As Double
    ProcessValue = mdblProcess
End Property
'Abort文件开始执行行号
Public Property Get Abort() As Long
    Abort = mlngAbort
End Property

Public Property Let Abort(ByVal lngAbort As Long)
    mlngAbort = lngAbort
End Property
'FormatSQL格式化的SQL
Public Property Get FormatSQL() As String
    FormatSQL = UCase(TrimComment(TrimEx(mobjSQL.SQL)))
End Property
'ErrCount错误数
Public Property Get ErrCount() As Long
    ErrCount = mlngErrCount
End Property

Public Property Let ErrCount(ByVal lngErrCount As Long)
    mlngErrCount = lngErrCount
End Property
'LogFile日志文件名
Public Property Let LogFile(ByVal strFile As String)
    mstrLogFile = strFile
End Property

Public Property Get LogFile() As String
    LogFile = mstrLogFile
End Property
'Log日志文件
Public Property Get Log() As TextStream
    Set Log = mobjLog
End Property
'SysNo 系统号
Public Property Let SysNo(ByVal lngSys As Long)
    mlngSys = lngSys
End Property

Public Property Get SysNo() As Long
    SysNo = mlngSys
End Property
'IsMaskAllErr错误忽略
Public Property Let IsMaskAllErr(ByVal blnMaskAllErr As Boolean)
    mblnMaskAllErr = blnMaskAllErr
End Property

Public Property Get IsMaskAllErr() As Boolean
    IsMaskAllErr = mblnMaskAllErr
End Property
'SQLRecTime 记录SQL的时限
Public Property Let SQLRecTime(ByVal lngSQLRecTime As Long)
    mlngSQLRecTime = lngSQLRecTime
End Property

Public Property Get SQLRecTime() As Long
    SQLRecTime = mlngSQLRecTime
End Property
'Connection连接
Public Property Set Connection(ByVal cnThis As ADODB.Connection)
    Set mcnThis = cnThis
End Property

Public Property Get Connection() As ADODB.Connection
    Set Connection = mcnThis
End Property
'ProcMode 存储过程选项
Public Property Let ProcMode(ByVal intProcMode As Integer)
    mintProcMode = intProcMode
End Property

Public Property Get ProcMode() As Integer
    ProcMode = mintProcMode
End Property
'Procedures 搜集到或要执行的存储过程
Public Property Let Procedures(ByVal strProcedures As String)
    mstrProcedures = strProcedures
End Property

Public Property Get Procedures() As String
    ProcMode = mstrProcedures
End Property

'ConnectType 连接类型
Public Property Let ConnectType(ByVal intConnectType As Integer)
    mintConnectType = intConnectType
End Property

Public Property Get ConnectType() As Integer
    ConnectType = mintConnectType
End Property
'Server 服务器名
Public Property Let Server(ByVal strServer As String)
    mstrServer = strServer
End Property

Public Property Get Server() As String
    Server = mstrServer
End Property
'IsRecErrSQL 记录错误SQL
Public Property Let IsRecErrSQL(ByVal blnRecErrSQL As Boolean)
    mblnRecErrSQL = blnRecErrSQL
End Property

Public Property Get IsRecErrSQL() As Boolean
    IsRecErrSQL = mblnRecErrSQL
End Property
'AbortInfo 中断信息
Public Property Let AbortInfo(ByRef strAbortInfo As String)
    mstrAbortInfo = strAbortInfo
End Property

Public Property Get AbortInfo() As String
    AbortInfo = mstrAbortInfo
End Property
'用户密码列表
'SysUser 系统所有者
Public Property Let SysUser(ByVal strSysUser As String)
    mulCurrent.SysUser = strSysUser
End Property

Public Property Get SysUser() As String
    SysUser = mulCurrent.SysUser
End Property

'SysPass系统所有者密码
Public Property Let SysPass(ByVal strSysPass As String)
    mulCurrent.SysPass = strSysPass
End Property

Public Property Get SysPass() As String
    SysPass = mulCurrent.SysPass
End Property
'ToolsPass ZLTOOLS密码
Public Property Let ToolsPass(ByVal strToolsPass As String)
    mulCurrent.ToolsPass = strToolsPass
End Property

Public Property Get ToolsPass() As String
    ToolsPass = mulCurrent.ToolsPass
End Property
'DBAUser DBA用户名
Public Property Let DBAUser(ByVal strDBAUser As String)
    mulCurrent.DBAUser = strDBAUser
End Property

Public Property Get DBAUser() As String
    DBAUser = mulCurrent.DBAUser
End Property
'DBAPass DBA密码
Public Property Let DBAPass(ByVal strDbaPass As String)
    mulCurrent.DBAPass = strDbaPass
End Property

Public Property Get DBAPass() As String
    DBAPass = mulCurrent.DBAPass
End Property
'GatherTables统计信息收集搜集的的表,以分号分割表名
Public Property Get GatherTables() As String
    GatherTables = mstrGatherTables
End Property

Public Property Let GatherTables(ByVal strGatherTables As String)
    mstrGatherTables = strGatherTables
End Property
'IsGather 是否进行统计信息收集
Public Property Let IsGather(ByVal blnGratherSta As Boolean)
    mblnGratherSta = blnGratherSta
End Property

Public Property Get IsGather() As Boolean
    IsGather = mblnGratherSta
End Property
'IsUseLog 是否记录日志
Public Property Let IsUseLog(ByVal blnUseLog As Boolean)
    mblnUseLog = blnUseLog
End Property

Public Property Get IsUseLog() As Boolean
    IsUseLog = mblnUseLog
End Property
'IsMaskIgnore(R/W)是否忽略可忽略错误
Public Property Let IsMaskIgnore(ByVal blnMaskIgnore As Boolean)
    mblnMaskIgnore = blnMaskIgnore
End Property

Public Property Get IsMaskIgnore() As Boolean
    IsMaskIgnore = mblnMaskIgnore
End Property
'Roles 角色授权收集到的角色集合
Public Property Let Roles(ByRef cllRoles As Collection)
    Set mcllRoles = cllRoles
End Property

Public Property Get Roles() As Collection
    Set Roles = mcllRoles
End Property
'IsRoleCollect 是否进行角色收集
Public Property Let IsRoleCollect(ByRef blnRoleCollect As Boolean)
    mblnRoleCollect = blnRoleCollect
End Property

Public Property Get IsRoleCollect() As Boolean
    IsRoleCollect = mblnRoleCollect
End Property

'IsOnlineIndex:在线操作索引模式（创建索引、创建主键唯一键有Using iNdex 关键字，重建索引如果没有ONline关键字，则自动添加）(R/W)
Public Property Let IsOnlineIndex(ByRef blnOnlineIndex As Boolean)
    mblnOnlineIndex = blnOnlineIndex
End Property

Public Property Get IsOnlineIndex() As Boolean
    IsOnlineIndex = mblnOnlineIndex
End Property
'***************************************************************************
'                三、向外部公布的方法
'***************************************************************************
'===========================================================================
'                1、初始化方法
'===========================================================================

Public Function InitGlobalPara(frmParent As Object, Optional ByVal lngSys As Long, Optional ByVal blnMaskAllErr As Boolean = False, Optional ByVal strLogFile As String, _
                                                    Optional ByVal lngRecSQLTime As Long, Optional ByVal blnGatherSta As Boolean, Optional ByVal blnOnlineIndex As Boolean, _
                                                    Optional ByVal blnRecErrSQL As Boolean = True) As Boolean
'功能：初始化全局参数，如记录日志，错误忽略等,一般只在调用一次，如在整个升级过程中、应用系统创建过程中等至调用一次就可以。
'参数：    lngSys=系统编号
'      blnMaskAllErr=忽略错误
'      strLogFile=日志文件
'      lngRecSQLTime=对执行时间超过lngRecSQLTime分钟的SQL进行记录,=0不进行记录
'      blnGatherSta=是否进行统计信息搜集
'      blnMaskIgnore=自动忽略可忽略错误
'      blnOnlineIndex=是否索引操作采用在线模式
'      blnRecErrSQL=是否只记录出错SQL
'返回：是否启用日志且日志文件创建失败，失败则返回False，否则True
    mlngSys = lngSys
    mblnMaskAllErr = blnMaskAllErr
    Set mfrmParent = frmParent
    mblnGratherSta = blnGatherSta
    mblnMaskIgnore = True '默认屏蔽可以忽略的错误
    mstrGatherTables = ""
    mlngErrCount = 0
    mblnRoleCollect = False '默认不进行角色收集
    Set mcllRoles = New Collection
    mstrServer = gstrServer '默认当前服务器
    Set mobjLog = Nothing
    mblnUseLog = True '默认使用日志记录
    mblnOnlineIndex = blnOnlineIndex
    If strLogFile <> "" Then
        If Not mobjFSO.FileExists(strLogFile) Then
            mobjFSO.CreateTextFile (strLogFile)
        End If
        Set mobjLog = mobjFSO.OpenTextFile(strLogFile, ForAppending)
    End If
    '日志创建失败或不创建日志则清空变量
    If mobjLog Is Nothing Then
        mlngSQLRecTime = 0
        mstrLogFile = ""
        mblnRecErrSQL = False
        mblnUseLog = False
    Else
        mlngSQLRecTime = lngRecSQLTime
        mstrLogFile = strLogFile
        mblnRecErrSQL = blnRecErrSQL '默认只记录出错SQL
    End If

    InitGlobalPara = Not (mstrLogFile = "" And strLogFile <> "")
End Function

Public Sub InitUserList(Optional ByVal strSysUser As String, Optional ByVal strSysPass As String, Optional ByVal strToolsPass As String, Optional ByVal strDBAUser As String, Optional ByVal strDbaPass As String)
'功能：初始化用户列表
'参数：strSysUser=系统所有者名称
'      strSysPass=系统所有者密码
'      strToolsPass=管理工具密码，管理工具用户名称为“ZLTOOLS”
'      strDBAUser=DBA用户名
'      strDBAPass=DBA用户密码
'说明：该函数一般与InitGlobalPara同时调用
    mulCurrent.SysUser = IIf(strSysUser = "", gstrUserName, strSysUser)
    mulCurrent.SysPass = IIf(strSysPass = "", gstrPassword, strSysPass)
    mulCurrent.ToolsPass = IIf(strToolsPass = "", IIf(gstrToolsPwd = "" Or gstrToolsPwd = "-", "ZLTOOLS", gstrToolsPwd), strToolsPass)
    mulCurrent.DBAUser = IIf(strSysUser = "", gstrSysUser, strDBAUser)
    mulCurrent.DBAPass = IIf(strSysUser = "", gstrSysPwd, strDbaPass)
End Sub

Public Function OpenFile(ByVal strFile As String, Optional ByVal lngAbort As Long, Optional ByVal intConnectType As Integer = -1, Optional ByVal cnThis As ADODB.Connection) As Boolean
'功能：打开脚本文件,无需传入连接
'参数：strFile=脚本文件名
'      lngAbort=起始行号
'      cnThis=当前连接
'      intConnectType=当前连接类型
'返回：是否打开文件成功

    mlngLinesCount = 0
    mlngLine = 0
    mblnEOF = True
    Erase marrDefine '清空marrDefine，使之恢复初始状态
    ReDim marrDefine(0) '下标为0的元素不用,用来判断
    mstrFile = strFile
    mlngAbort = lngAbort
    mintSection = 0
    If GetTextStream(strFile, True) Then
        '不传连接，说明已经在外部通过属性指定
        If Not cnThis Is Nothing Then Set mcnThis = cnThis
        '不传连接类型，说明外部已经指定
        If intConnectType <> -1 Then mintConnectType = intConnectType
        Call ReadNextSQL '需读取一个SQL，初始化
        OpenFile = True
    Else
        OpenFile = False
    End If
End Function

Public Function AnalysisSQLString(ByVal strSQL As String, Optional ByVal lngStartLine As Long = 1, Optional ByVal strSplit As String = vbNewLine) As Boolean
'功能：将多个SQL组成的字符串解析成单个SQL
    Dim strLine As String, strFMT As String
    Dim arrLines As Variant, i As Long, j As Long
    Dim arrKeyWord As Variant, blnSQL As Boolean
    Dim objSQL As New clsSQLInfo, strBuild As New clsStringBulider
    Dim blnTip As Boolean '是否读取注释
    Dim lngCount As Long, strTemp As String
    Dim strPartSQL As String
    Dim lngPos As Long
    
    
    On Error GoTo errH
    Set mcllStrSQL = New Collection
    mblnEOF = True
    mintStrSQLIdx = 0
    arrLines = Split(strSQL, strSplit)
    Erase marrDefine '清空marrDefine，使之恢复初始状态
    ReDim marrDefine(0) '下标为0的元素不用,用来判断
    lngCount = 0
    blnTip = True
    For i = LBound(arrLines) To UBound(arrLines)
        strLine = Trim(arrLines(i))
        strFMT = Trim(UCase(TrimComment(TrimEx(strLine))))
        If strFMT = "" And Not objSQL.Block Then GoTo NextLine
        '提示信息行，不处理
        If strFMT Like "PROMPT *" Then GoTo NextLine '进入下次循环
        '处理替换变量,则将替换变量添加到替换变量数据，语句不添加到SQL中
        If strFMT Like "DEFINE *" And Not objSQL.Block Then '定义不会在块中
            ReDim Preserve marrDefine(UBound(marrDefine) + 1)
            strFMT = SplitSQL(strFMT): j = UBound(marrDefine)
            marrDefine(j).varName = Trim(Split(Replace(strFMT, "DEFINE ", ""), "=")(0))
            marrDefine(j).varValue = Trim(Split(Replace(strFMT, "DEFINE ", ""), "=")(1))
            If Left(marrDefine(j).varValue, 1) = "'" And Right(marrDefine(j).varValue, 1) = "'" Then
                marrDefine(j).varValue = Mid(marrDefine(j).varValue, 2, Len(marrDefine(j).varValue) - 2)
            End If
            '进入下次循环
            GoTo NextLine
        ElseIf UBound(marrDefine) > 0 Then '块中的也需要替换,vbTextCompare忽略大小写进行替换
            For j = 1 To UBound(marrDefine)
                If marrDefine(j).varName = "N_SYSTEM" Then
                    strLine = Replace(strLine, "&" & marrDefine(j).varName, mlngSys, , , vbTextCompare)
                Else
                    strLine = Replace(strLine, "&" & marrDefine(j).varName, marrDefine(j).varValue, , , vbTextCompare)
                End If
            Next
        End If
        
        If objSQL.Block Then
            '当前行在PL/SQL Block中
            If strFMT = "/" Then
                blnSQL = True
            Else
                Call strBuild.Append(vbCrLf & strLine)
                If lngCount < 5 Then
                    strPartSQL = strPartSQL & vbCrLf & strLine
                    lngCount = lngCount + 1
                End If
            End If
        ElseIf strFMT = "BEGIN" Or strFMT Like "BEGIN *" Or strFMT = "DECLARE" Or strFMT Like "DECLARE *" _
            Or strFMT Like "CREATE OR REPLACE PROCEDURE *" Or strFMT Like "CREATE PROCEDURE *" _
            Or strFMT Like "CREATE OR REPLACE FUNCTION *" Or strFMT Like "CREATE FUNCTION *" _
            Or strFMT Like "CREATE OR REPLACE TRIGGER *" Or strFMT Like "CREATE TRIGGER *" _
            Or strFMT Like "CREATE OR REPLACE TYPE *" Or strFMT Like "CREATE TYPE *" _
            Or strFMT Like "CREATE OR REPLACE PACKAGE *" Or strFMT Like "CREATE PACKAGE *" Then
            blnTip = False
            '当前行是PL/SQL Block开始行:不包含视图,视图和普通语句一样处理。
            objSQL.FileLine = lngStartLine + i: objSQL.Block = True: objSQL.HaveSQL = True
            Call strBuild.Append(vbCrLf & strLine)

            '取块相关信息
            If InStr(strFMT, "WRAPPED") > 0 Then
                '解析Wrapped过程信息
                arrKeyWord = Split(strFMT, " ")
                For j = UBound(arrKeyWord) To 0 Step -1
                    If arrKeyWord(j) = "WRAPPED" Then
                        objSQL.BlockType = arrKeyWord(j - 2)
                        If objSQL.BlockType = "BODY" Then 'PACKAGE BODY
                            objSQL.BlockType = arrKeyWord(j - 3) & " " & objSQL.BlockType
                        End If
                        objSQL.BlockName = arrKeyWord(j - 1)
                        Exit For
                    End If
                Next
                objSQL.BlockWrap = True
            ElseIf strFMT Like "CREATE *" Then
                '解析创建块的信息
                strTemp = Replace(strFMT, "CREATE OR REPLACE ", "")
                strTemp = Replace(strTemp, "CREATE ", "")
                arrKeyWord = Split(strTemp, " ")

                objSQL.BlockType = arrKeyWord(0)
                If arrKeyWord(1) = "BODY" Then
                    'PACKAGE BODY,TYPE BODY
                    '块类型不包括MATERIALIZED VIEW
                    objSQL.BlockType = objSQL.BlockType & " " & arrKeyWord(1)
                    objSQL.BlockName = arrKeyWord(2)
                Else
                    objSQL.BlockName = arrKeyWord(1)
                End If
            End If
            If InStr(objSQL.BlockName, "(") > 0 Then objSQL.BlockName = Left(objSQL.BlockName, InStr(objSQL.BlockName, "(") - 1)
            If InStr(objSQL.BlockName, ".") > 0 Then objSQL.BlockName = Split(objSQL.BlockName, ".")(1) 'ZLHIS.XXXXX
            If objSQL.BlockName <> "" Then '名称按原始大小写
                lngPos = InStr(1, strLine, objSQL.BlockName, vbTextCompare)
                objSQL.BlockName = Mid(strLine, lngPos, Len(objSQL.BlockName))
            End If
        Else
            '当前行不在PL/SQL Block中
            If strFMT = "" Then
            ElseIf Left(strFMT, 2) = "--" Then
                If blnTip Then '只有在单条语句或存储过程开头的语句才算注释
                    objSQL.Tip = objSQL.Tip & vbCrLf & strLine
                End If
            ElseIf strFMT = "/" Then
                blnSQL = True
            ElseIf Right(strFMT, 1) = ";" Then
                blnTip = False
                If strBuild.Length = 0 Then objSQL.FileLine = lngStartLine + i
                If strFMT <> ";" Then objSQL.HaveSQL = True
                Call strBuild.Append(vbCrLf & SplitSQL(strLine))
                If lngCount < 5 Then
                    strPartSQL = strPartSQL & vbCrLf & SplitSQL(strLine)
                    lngCount = lngCount + 1
                End If
                blnSQL = True
            Else
                blnTip = False
                If strBuild.Length = 0 Then objSQL.FileLine = lngStartLine + i
                Call strBuild.Append(vbCrLf & strLine)
                objSQL.HaveSQL = True
                If lngCount < 5 Then
                    strPartSQL = strPartSQL & vbCrLf & strLine
                    lngCount = lngCount + 1
                End If
            End If
        End If
NextLine:
        '获取可执行的SQL语句
        If blnSQL Then
            objSQL.SQL = strBuild.ToString
            strBuild.Clear
            objSQL.Tip = Mid(objSQL.Tip, 3)
            objSQL.SQL = Mid(objSQL.SQL, 3)
            objSQL.PartSQL = UCase(TrimComment(TrimEx(Mid(strPartSQL, 3))))
            mcllStrSQL.Add objSQL.CopyMe
            objSQL.ClearSQL
            lngCount = 0
            blnTip = True
            blnSQL = False
        End If
    Next
    If mcllStrSQL.Count = 0 Then
        mblnEOF = True
        AnalysisSQLString = True
    Else
        mblnEOF = False
        mintStrSQLIdx = mintStrSQLIdx + 1
        Set mobjSQL = mcllStrSQL(mintStrSQLIdx)
        AnalysisSQLString = True
    End If
    Exit Function
errH:
    If 0 = 1 Then
        Resume
    End If
End Function
'===========================================================================
'                2、业务功能方法
'===========================================================================

Public Function ExecuteFile(ByVal strFile As String, Optional ByVal lngAbort As Long, Optional ByVal strExecProcs As String, Optional ByVal intConnectType As Integer = -1, Optional ByVal cnThis As ADODB.Connection) As Boolean
'功能:直接执行文件
'功能:直接执行文件
'参数：strFile=脚本文件名
'      lngAbort=起始行号
'      cnThis=当前连接
'      strExecProcs=需执行的存储过程
'      strCollectedProcs=手机到的存储过程
'返回：是否打开文件成功
    mstrProcedures = strExecProcs
    mintProcMode = 0
    mblnUseLog = True
    '执行存储过程，说明脚本是可选脚本，可选脚本中是存储过程，执行时不能从中断行号执行。
    '没有被选择的可选脚本，不执行并且不调用存储过程
    If mstrProcedures <> "" Then lngAbort = 0: mintProcMode = 2
    If OpenFile(strFile, lngAbort, intConnectType, cnThis) Then
        Do While Not mblnEOF
            If ExecuteSQL = False Then Exit Function
            Call CollectTables
            Call ReadNextSQL
        Loop
    End If
    ExecuteFile = True
End Function

Public Function CollectProcs(ByVal strFile As String) As String
'功能：搜集存储过程名
    mstrProcedures = ""
    mintProcMode = 1
    If OpenFile(strFile) Then
        Do While Not mblnEOF
            Call ExecuteSQL
            Call ReadNextSQL
        Loop
    End If
    CollectProcs = mstrProcedures
End Function

Public Function ReadNextSQL() As Boolean
'功能：读取下一个完整的可执行SQL
'说明：可能不是从文件的开头执行的，因此需要先跳跃到执行位置读取才算有效SQL
    If mcllStrSQL Is Nothing Then
        If mlngLinesCount = 0 Then
            mdblProcess = 100
        Else
            mdblProcess = IIf(mlngLine > mlngLinesCount, mlngLinesCount, mlngLine) / mlngLinesCount * 100
        End If
        If mlngLine >= mlngAbort Then
            ReadNextSQL = ReadNextOneSQL
        Else
            Do
                If mlngLine < mlngAbort Then
                    ReadNextSQL = ReadNextOneSQL
                Else
                    Exit Do
                End If
            Loop Until mblnEOF
        End If
        '缺失Online,则自动添加
        If Me.SQLInfo.LackOnline And mblnOnlineIndex Then
            Me.SQLInfo.SQL = Me.SQLInfo.SQL & " OnLine"
        End If
    Else
        mintStrSQLIdx = mintStrSQLIdx + 1
        If mintStrSQLIdx <= mcllStrSQL.Count Then
            Set mobjSQL = mcllStrSQL(mintStrSQLIdx)
            mblnEOF = False
        Else
            mblnEOF = True
        End If
    End If
End Function

Public Function ExecuteSQL(Optional ByVal objSQL As clsSQLInfo) As Boolean
'功能：执行当前的读取到的SQL语句
    Dim objSQLExe As clsSQLInfo
    If Not objSQL Is Nothing Then Set mobjSQL = objSQL
    If mobjSQL.HaveSQL Then   'Line是以SQL结束行的为准判断和记录
        'ADO不能使用Execute调用过程,换成Call
        If Not mobjSQL.Block And mobjSQL.PartSQL Like "EXECUTE *" Then  '认为写在一行上的
            mobjSQL.SQL = Replace(mobjSQL.SQL, "Execute", "Call", , 1, vbTextCompare)
            '如果不使用{},则即使过程没有参数也要加()
            If InStr(mobjSQL.SQL, "(") = 0 Then mobjSQL.SQL = mobjSQL.SQL & "()"
        End If

        '处理SQL语句
        If mintProcMode = 1 Then
            'x.返回SQL脚本中的(创建)过程名
            If mobjSQL.Block And mobjSQL.BlockType = "PROCEDURE" Then
                mstrProcedures = mstrProcedures & IIf(mstrProcedures = "", "", "?") & mobjSQL.BlockName & "|" & mobjSQL.Tip
            End If
        ElseIf mintProcMode = 2 Then
            'x.只创建指定的过程，并自动调用执行过程
            If mobjSQL.Block And mobjSQL.BlockType = "PROCEDURE" And InStr("," & mstrProcedures & ",", "," & mobjSQL.BlockName & ",") > 0 Then
                Set objSQLExe = mobjSQL.CopyMe
                objSQLExe.Block = False: objSQLExe.BlockWrap = False
                objSQLExe.BlockName = "": objSQLExe.BlockType = ""
                '创建过程
                If Not RunSQLText() Then GoTo AbortLine
                '执行过程
                objSQLExe.SQL = "Call " & mobjSQL.BlockName & "()" '这些过程都没有参数
                If Not RunSQLText(objSQLExe) Then GoTo AbortLine
                '删除过程
                objSQLExe.SQL = "Drop Procedure " & mobjSQL.BlockName
                If Not RunSQLText(objSQLExe) Then GoTo AbortLine
            End If
        Else
            'x.执行SQL语句
            If Not RunSQLText() Then GoTo AbortLine
        End If

    End If
    ExecuteSQL = True
    Exit Function
AbortLine:

End Function

Public Sub CollectTables()
'功能：统计信息收集的表的搜集
    Dim strTable As String
    
    If Me.IsGather Then
        If mintProcMode = 0 Then '普通SQL脚本执行
            If Not mobjSQL.Block Then '非块类型的SQL
                If mobjSQL.PartSQL Like "ALTER TABLE * ADD *" And Not mobjSQL.PartSQL Like "ALTER TABLE * ADD CONSTRAINT *" Then
                    strTable = Trim(Mid(mobjSQL.PartSQL, Len("ALTER TABLE "), 30))  '一般表名长度不会超过30个数据库长度，因为数据库中长度为30
                    If InStr(strTable, "ADD") > 0 Then
                        strTable = Trim(Mid(strTable, 1, InStr(strTable, "ADD") - 1))
                    End If
                ElseIf mobjSQL.PartSQL Like "CREATE TABLE*" Then
                    strTable = Trim(Mid(mobjSQL.PartSQL, InStr(mobjSQL.PartSQL, "CREATE TABLE ") + Len("CREATE TABLE "), 30))   '一般表名长度不会超过30个数据库长度，因为数据库中长度为30
                    If InStr(strTable, "(") > 0 Then
                        strTable = Trim(Mid(strTable, 1, InStr(strTable, "(") - 1))
                    End If
                End If
                If strTable <> "" Then
                    If InStr("," & mstrGatherTables & ",", "," & strTable & ",") = 0 Then
                        mstrGatherTables = IIf(mstrGatherTables = "", "", mstrGatherTables & ",") & strTable
                    End If
                End If
            End If
        End If
    End If
End Sub

Public Sub CollectRoles()
'功能：搜集角色
    Dim strRole As String
    
    If Not mobjSQL.Block Then
        If mobjSQL.PartSQL Like "CREATE ROLE*" Then
            strRole = Trim(Mid(mobjSQL.PartSQL, 12))
            If strRole <> "" Then
                err = 0: On Error Resume Next
                mcllRoles.Add strRole, "K" & strRole
                If err <> 0 Then
                     err.Clear: err = 0
                End If
                On Error GoTo 0
            End If
        End If
    End If
End Sub

Public Sub WriteSection(ByVal strSectionName As String, Optional ByVal strChar = "=")
    If Not Me.Log Is Nothing And mblnUseLog Then
        Me.Log.WriteLine String(120, strChar)
        Me.Log.WriteLine String(4, strChar) & Lpad(strSectionName, (120 - 4 * 2 + Len(strSectionName)) \ 2)
        Me.Log.WriteLine String(120, strChar)
    End If
End Sub

'日志方法
Public Sub WriteLog(Optional ByVal strText As String, Optional ByVal bytAddTime As Byte, Optional ByVal lngReturnLines As Long)
'参数:strText       =要写入的一行日志文本,如果为空表示写一行换行符
'     bytAddTime    >0时，在日志文本之前加入日志时间，1=日期和时间完整格式,2-仅时间,0-不加入
'     lngReturnLines=在日志文本之后写多少行换行符,0-不写换行符

    If Not mblnUseLog Then Exit Sub
    If Not Me.Log Is Nothing Then
        '添加日志时间
        If bytAddTime > 0 Then strText = LogTime(bytAddTime = 2) & strText
        
        '截取超长内容
        If Len(strText) > 200 Then strText = Mid(strText, 1, 200)
                
        '写日志文本
        If strText = "" Then
             Me.Log.WriteBlankLines 1
        Else
            Me.Log.WriteLine strText
            If lngReturnLines > 0 Then Me.Log.WriteBlankLines lngReturnLines
        End If
    End If
End Sub

Public Sub CloseLog()
'功能:关闭日志对象，防止占用，不能续写日志
    If Not mobjLog Is Nothing Then
        mobjLog.Close
        Set mobjLog = Nothing
    End If
End Sub

'***************************************************************************
'                四、类私有的方法
'***************************************************************************

Private Function GetTextStream(ByVal strFile As String, Optional ByVal blnMsg As Boolean) As Boolean
'功能：打开脚本文件
'参数：strFile=脚本文件
'      blnMsg=是否进行错误提示
'返回：打开成功=True,失败=False
    Dim strTemp As String
    
    If Not mobjFSO.FileExists(strFile) Then
        mlngErrCount = mlngErrCount + 1
        strTemp = "文件""" & strFile & """不存在,操作被迫中止!"
        Me.WriteLog Format(Now, "HH:mm:ss") & "，文件：" & strFile
        Me.WriteLog Format(Now, "HH:mm:ss") & "，结果：" & strTemp
        If blnMsg Then
            MsgBox strTemp, vbCritical, gstrSysName: Exit Function
        End If
    Else
        Me.WriteLog Format(Now, "HH:mm:ss") & "，文件：" & strFile & IIf(mlngAbort <> 0, ",开始行号：" & mlngAbort, "")
    End If
    
    '获取行数:空文件正常退出
    Set mobjText = mobjFSO.OpenTextFile(strFile, ForReading)
    If mobjText.AtEndOfStream Then
        GetTextStream = True: Exit Function
    End If
    mobjText.ReadAll: mlngLinesCount = mobjText.Line: mobjText.Close
    Set mobjText = mobjFSO.OpenTextFile(strFile, ForReading)
    
    GetTextStream = True
End Function

Private Sub ClearAndInitSQL()
'功能：清空并初始化SQL对象
    mobjSQL.ClearSQL
    mobjSQL.FileLines = mlngLinesCount
    mobjSQL.FileLine = mlngLine
    mobjSQL.File = mstrFile
End Sub

Private Function ReadNextOneSQL() As Boolean
'功能：读取下一个SQL，类似RecordSet的MoveNext方法
    Dim strLine As String, strFMT As String
    Dim arrKeyWord As Variant, blnSQL As Boolean
    Dim strTemp As String, i As Long
    Dim blnTip As Boolean '是否读取注释
    Dim lngCount As Long
    Dim strPartSQL As String
    Dim intSecStep As Integer, strSectTmp As String
    Call ClearAndInitSQL
    lngCount = 0
    
    blnTip = True
    Do While Not mobjText.AtEndOfStream
        mlngLine = mobjText.Line '当前行号
        strLine = RTrim(mobjText.ReadLine)
        strFMT = UCase(TrimComment(TrimEx(strLine)))
        If blnTip Then
            strSectTmp = Trim(strLine)
            If Replace(strFMT, "-", "") = "" Then
                If Len(strFMT) > 36 Then '36个注释标号，则可以认为是注释区段的开始或结束
                    If intSecStep = 2 Then '区段结束
                        mintSection = mintSection + 1
                    Else '区段开始
                        intSecStep = 1
                    End If
                End If
            ElseIf intSecStep = 1 Then
                strSectTmp = Trim(strFMT)
                If strSectTmp = "--结构修正部份" Or strSectTmp = "--数据修正部份" Or strSectTmp = "--权限修正部份" Or _
                    strSectTmp = "--报表修正部份" Or strSectTmp = "--过程修正部份" Or strSectTmp = "--更改系统及部件的版本号" Then
                    intSecStep = 2
                End If
            Else
                intSecStep = 0
            End If
        Else
            intSecStep = 0
        End If
        '提示信息行，不处理
        If strFMT Like "PROMPT *" Then GoTo NextLine '进入下次循环
        '处理替换变量,则将替换变量添加到替换变量数据，语句不添加到SQL中
        If strFMT Like "DEFINE *" And Not mobjSQL.Block Then '定义不会在块中
            ReDim Preserve marrDefine(UBound(marrDefine) + 1)
            strFMT = SplitSQL(strFMT): i = UBound(marrDefine)
            marrDefine(i).varName = Trim(Split(Replace(strFMT, "DEFINE ", ""), "=")(0))
            marrDefine(i).varValue = Trim(Split(Replace(strFMT, "DEFINE ", ""), "=")(1))
            If Left(marrDefine(i).varValue, 1) = "'" And Right(marrDefine(i).varValue, 1) = "'" Then
                marrDefine(i).varValue = Mid(marrDefine(i).varValue, 2, Len(marrDefine(i).varValue) - 2)
            End If
            '进入下次循环
            GoTo NextLine
        ElseIf UBound(marrDefine) > 0 Then '块中的也需要替换,vbTextCompare忽略大小写进行替换
            For i = 1 To UBound(marrDefine)
                If marrDefine(i).varName = "N_SYSTEM" Then
                    strLine = Replace(strLine, "&" & marrDefine(i).varName, mlngSys, , , vbTextCompare)
                Else
                    strLine = Replace(strLine, "&" & marrDefine(i).varName, marrDefine(i).varValue, , , vbTextCompare)
                End If
            Next
        End If

        If mobjSQL.Block Then
            '当前行在PL/SQL Block中
            If strFMT = "/" Then
                blnSQL = True: Exit Do
            Else
                Call mstrBuild.Append(vbCrLf & strLine)
                
                If lngCount < 5 Then
                    strPartSQL = strPartSQL & vbCrLf & strLine
                    lngCount = lngCount + 1
                End If
            End If
        ElseIf strFMT = "BEGIN" Or strFMT Like "BEGIN *" Or strFMT = "DECLARE" Or strFMT Like "DECLARE *" _
            Or strFMT Like "CREATE OR REPLACE PROCEDURE *" Or strFMT Like "CREATE PROCEDURE *" _
            Or strFMT Like "CREATE OR REPLACE FUNCTION *" Or strFMT Like "CREATE FUNCTION *" _
            Or strFMT Like "CREATE OR REPLACE TRIGGER *" Or strFMT Like "CREATE TRIGGER *" _
            Or strFMT Like "CREATE OR REPLACE TYPE *" Or strFMT Like "CREATE TYPE *" _
            Or strFMT Like "CREATE OR REPLACE PACKAGE *" Or strFMT Like "CREATE PACKAGE *" Then
            blnTip = False
            '当前行是PL/SQL Block开始行:不包含视图,视图和普通语句一样处理。
            mobjSQL.FileLine = mlngLine: mobjSQL.Block = True: mobjSQL.HaveSQL = True
            Call mstrBuild.Append(vbCrLf & strLine)

            '取块相关信息
            If InStr(strFMT, "WRAPPED") > 0 Then
                '解析Wrapped过程信息
                arrKeyWord = Split(strFMT, " ")
                For i = UBound(arrKeyWord) To 0 Step -1
                    If arrKeyWord(i) = "WRAPPED" Then
                        mobjSQL.BlockType = arrKeyWord(i - 2)
                        If mobjSQL.BlockType = "BODY" Then 'PACKAGE BODY
                            mobjSQL.BlockType = arrKeyWord(i - 3) & " " & mobjSQL.BlockType
                        End If
                        mobjSQL.BlockName = arrKeyWord(i - 1)
                        Exit For
                    End If
                Next
                mobjSQL.BlockWrap = True
            ElseIf strFMT Like "CREATE *" Then
                '解析创建块的信息
                strTemp = Replace(strFMT, "CREATE OR REPLACE ", "")
                strTemp = Replace(strTemp, "CREATE ", "")
                arrKeyWord = Split(strTemp, " ")

                mobjSQL.BlockType = arrKeyWord(0)
                If arrKeyWord(1) = "BODY" Then
                    'PACKAGE BODY,TYPE BODY
                    '块类型不包括MATERIALIZED VIEW
                    mobjSQL.BlockType = mobjSQL.BlockType & " " & arrKeyWord(1)
                    mobjSQL.BlockName = arrKeyWord(2)
                Else
                    mobjSQL.BlockName = arrKeyWord(1)
                End If
            End If
            If InStr(mobjSQL.BlockName, "(") > 0 Then mobjSQL.BlockName = Left(mobjSQL.BlockName, InStr(mobjSQL.BlockName, "(") - 1)
            If InStr(mobjSQL.BlockName, ".") > 0 Then mobjSQL.BlockName = Split(mobjSQL.BlockName, ".")(1) 'ZLHIS.XXXXX
            If mobjSQL.BlockName <> "" Then '名称按原始大小写
                i = InStr(1, strLine, mobjSQL.BlockName, vbTextCompare)
                mobjSQL.BlockName = Mid(strLine, i, Len(mobjSQL.BlockName))
            End If
        Else
            '当前行不在PL/SQL Block中
            If strFMT = "" Then
            ElseIf Left(strFMT, 2) = "--" Then
                If blnTip Then '只有在单条语句或存储过程开头的语句才算注释
                    mobjSQL.Tip = mobjSQL.Tip & vbCrLf & strLine
                End If
            ElseIf strFMT = "/" Then
                blnSQL = True: Exit Do
            ElseIf Right(strFMT, 1) = ";" Then
                blnTip = False
                If mstrBuild.Length = 0 Then mobjSQL.FileLine = mlngLine
                If strFMT <> ";" Then mobjSQL.HaveSQL = True
                Call mstrBuild.Append(vbCrLf & SplitSQL(strLine))
                If lngCount < 5 Then
                    strPartSQL = strPartSQL & vbCrLf & SplitSQL(strLine)
                    lngCount = lngCount + 1
                End If
                blnSQL = True: Exit Do
            Else
                blnTip = False
                If mstrBuild.Length = 0 Then mobjSQL.FileLine = mlngLine
                Call mstrBuild.Append(vbCrLf & strLine)
                mobjSQL.HaveSQL = True
                If lngCount < 5 Then
                    strPartSQL = strPartSQL & vbCrLf & strLine
                    lngCount = lngCount + 1
                End If
            End If
        End If
NextLine:
    Loop
    '获取可执行的SQL语句
    If blnSQL Then
        mobjSQL.SQL = mstrBuild.ToString
        mstrBuild.Clear
        mobjSQL.Tip = Mid(mobjSQL.Tip, 3)
        mobjSQL.SQL = Mid(mobjSQL.SQL, 3)
        mobjSQL.PartSQL = UCase(TrimComment(TrimEx(Mid(strPartSQL, 3))))
    End If
    
    mblnEOF = mobjText.AtEndOfStream And Not blnSQL
    
    ReadNextOneSQL = blnSQL
    
End Function

Private Function RunSQLText(Optional objSQL As clsSQLInfo) As Boolean
'功能：执行当前的SQL语句
'参数：objSQL=当前要执行的SQL语句信息,默认是mobjSQL
'说明：
    Dim rsTemp As ADODB.Recordset
    Dim objWrap As Scripting.TextStream
    Dim strUser As String, strPass As String, strFile As String
    Dim lngCommand As Long, lngProcess As Long
    Dim strSQL As String, lngAffect As Long
    Dim lngErrNum As Long, strErrInfo As String, blnHaveErr As Boolean
    Dim blnRetry As Boolean, strLog As String, blnSucess As Boolean
    Dim vReturn As VbMsgBoxResult, blnSysIgnore As Boolean
    Dim datBegin As Date, datEnd As Date, lngSQLTime As Long
    Dim lngTimes As Long, blnErrAtuoRepaired As Boolean
    
    If objSQL Is Nothing Then
        Set objSQL = mobjSQL
    End If
    
    '记录全部执行SQL
    If Not mblnRecErrSQL Then If Not blnRetry Then Me.WriteLog LogTime(True) & Format(objSQL.FileLine, "0000000") & "：" & GetLogSQL(objSQL)
    '处理SQL语句
    If objSQL.Block And objSQL.BlockWrap Then '加密块编译
    '--------------------------------------------------------------------------
        '生成临时脚本文件
        strFile = mobjFSO.GetSpecialFolder(TemporaryFolder).Path & "\" & mobjFSO.GetTempName
        Set objWrap = mobjFSO.OpenTextFile(strFile, ForWriting, True)
        objWrap.Write mobjSQL.SQL: objWrap.WriteLine "/": objWrap.WriteLine "exit;": objWrap.Close
        '生成SQLPlus命令
        Select Case mintConnectType
            Case 0
                strUser = mulCurrent.SysUser: strPass = mulCurrent.SysPass
            Case 1
                strUser = "zltools": strPass = mulCurrent.ToolsPass
            Case 2
                strUser = mulCurrent.DBAUser: strPass = mulCurrent.DBAPass
        End Select
        strSQL = "sqlplus " & strUser & "/" & strPass & IIf(mstrServer <> "", "@" & mstrServer, "") & " @" & strFile
        '执行Shell命令
        blnRetry = False: blnSucess = False
RetrySQLPlus:
        err.Clear: On Error Resume Next
        datBegin = Now: datEnd = Now
        DoEvents
        lngCommand = Shell(strSQL, vbHide) '53:文件未找到
        datEnd = Now
        If err.Number <> 0 Then
            lngErrNum = err.Number: strErrInfo = err.Description & IIf(lngErrNum = 53, ",请检查 sqlplus.exe 是否正确安装", "")
            If Not blnRetry And mblnRecErrSQL Then Me.WriteLog LogTime(True) & Format(objSQL.FileLine, "0000000") & ":" & GetLogSQL(objSQL)
            Me.WriteLog String(17, " ") & IIf(blnRetry, "重试", "错误") & "：" & strErrInfo
            If Not mblnMaskAllErr Then
                mlngErrCount = mlngErrCount + 1
                '不允许忽略错误
                If MsgBox("错误：" & lngErrNum & vbCrLf & vbCrLf & strErrInfo, vbRetryCancel + vbCritical, gstrSysName) = vbRetry Then
                    Me.WriteLog String(17, " ") & "处理：选择重试"
                    blnRetry = True: GoTo RetrySQLPlus
                Else
                    Me.WriteLog String(17, " ") & "处理：人工中止"
                    Kill strFile: GoTo AbortLine
                End If
            Else
                Me.WriteLog String(17, " ") & "处理：自动忽略"
            End If
        End If
        If lngCommand <> 0 Then
            lngTimes = 0
            lngProcess = OpenProcess(Process_Query_Information, False, lngCommand)
            Do
                lngTimes = lngTimes + 1
                'SQL超时
                If lngTimes > 4000 Then
                    Exit Do
                End If
                Sleep 50
                GetExitCodeProcess lngProcess, lngCommand
                DoEvents
            Loop While lngCommand = Still_Active
            CloseHandle lngProcess
        End If
        Kill strFile '删除临时文件
        err.Clear: On Error GoTo errH
        '检查执行结果
        Set rsTemp = New ADODB.Recordset
        strSQL = "Select Status From User_Objects Where Object_Type = '" & UCase(objSQL.BlockType) & "' And Object_Name='" & UCase(objSQL.BlockName) & "'"
        rsTemp.CursorLocation = adUseClient
        rsTemp.Open strSQL, mcnThis, adOpenKeyset
        If rsTemp.EOF Then
            If Not blnRetry And mblnRecErrSQL Then Me.WriteLog LogTime(True) & Format(objSQL.FileLine, "0000000") & "：" & GetLogSQL(objSQL)
            Me.WriteLog String(17, " ") & IIf(blnRetry, "重试", "执行") & "结果：加密的 " & objSQL.BlockType & " " & objSQL.BlockName & " 未能正确创建。"
            If Not mblnMaskAllErr Then
                mlngErrCount = mlngErrCount + 1
                '不允许忽略错误
                If MsgBox("加密的 " & objSQL.BlockType & " " & objSQL.BlockName & " 未能正确创建。" & _
                    vbCrLf & "请检查 sqlplus.exe 是否正确安装，以及脚本的正确性。", vbRetryCancel + vbCritical, gstrSysName) = vbRetry Then
                    Me.WriteLog String(17, " ") & "处理：选择重试"
                    blnRetry = True: GoTo RetrySQLPlus
                Else
                    Me.WriteLog String(17, " ") & "处理：人工中止"
                    GoTo AbortLine
                End If
            Else
                Me.WriteLog String(17, " ") & "处理：自动忽略"
            End If
        ElseIf Nvl(rsTemp!Status) <> "VALID" Then
            If Not blnRetry And mblnRecErrSQL Then Me.WriteLog LogTime(True) & Format(objSQL.FileLine, "0000000") & "：" & GetLogSQL(objSQL)
            Me.WriteLog String(17, " ") & IIf(blnRetry, "重试", "执行") & "结果：执行成功,但 " & objSQL.BlockType & " " & objSQL.BlockName & " 状态为 " & rsTemp!Status
        Else
            blnSucess = True
        End If
        If blnRetry And blnSucess Then '重试成功，则增加成功日志
            Me.WriteLog String(17, " ") & "重试结果：执行成功" & IIf(lngAffect > 0, "," & lngAffect & " 行数据生效", "")
        ElseIf blnSucess And Not mblnRecErrSQL Then '没有重试，且成功
            Me.WriteLog String(17, " ") & "执行结果：执行成功" & IIf(lngAffect > 0, "," & lngAffect & " 行数据生效", "")
        End If
        If mlngSQLRecTime <> 0 Then
            lngSQLTime = DateDiff("n", datBegin, datEnd)
            If lngSQLTime >= mlngSQLRecTime Then
                If Not blnRetry And mblnRecErrSQL Then Me.WriteLog LogTime(True) & Format(objSQL.FileLine, "0000000") & "：" & GetLogSQL(objSQL)
                Me.WriteLog String(17, " ") & "SQL执行耗时：" & lngSQLTime & "分钟"
            End If
        End If
    Else '其他SQL语句执行
    '--------------------------------------------------------------------------
        blnRetry = False: blnSucess = False
RetrySQLLine:
        blnErrAtuoRepaired = False
        err.Clear: On Error Resume Next
        datBegin = Now: datEnd = Now
        DoEvents
        mcnThis.Execute objSQL.SQL, lngAffect, adCmdText
        datEnd = Now
        If err.Number <> 0 Then
            If Not blnRetry And mblnRecErrSQL Then Me.WriteLog LogTime(True) & Format(objSQL.FileLine, "0000000") & ":" & GetLogSQL(objSQL)
            If mcnThis.Errors.Count > 0 Then
                lngErrNum = mcnThis.Errors(0).NativeError
                strErrInfo = mcnThis.Errors(0).Description
            Else
                lngErrNum = err.Number
                strErrInfo = err.Description
            End If
            strErrInfo = Replace(strErrInfo, "[Microsoft][ODBC driver for Oracle][Oracle]", "") '简化日志文件的重复文本
            Me.WriteLog String(17, " ") & IIf(blnRetry, "重试", "错误") & ":" & strErrInfo
            If Not mblnMaskAllErr Then
                mlngErrCount = mlngErrCount + 1
                blnSysIgnore = False
                '错误分析处理
                vReturn = frmAppUpgradeErr.ShowError(mcnThis, lngErrNum, strErrInfo, objSQL, mfrmParent, mblnMaskIgnore, blnSysIgnore, Me, blnErrAtuoRepaired)
                If vReturn = vbRetry Then '重试
                    Me.WriteLog String(17, " ") & "处理：选择重试"
                    blnRetry = True: GoTo RetrySQLLine
                ElseIf vReturn = vbIgnore Then '忽略
                    If blnErrAtuoRepaired Then
                        Me.WriteLog String(17, " ") & "处理：自动忽略(错误中心已经修正该错误)"
                    Else
                        Me.WriteLog String(17, " ") & "处理：" & IIf(blnSysIgnore, "自动忽略", "人工忽略")
                    End If
                ElseIf vReturn = vbAbort Then '中止
                    Me.WriteLog String(17, " ") & "处理：人工中止"
                    GoTo AbortLine
                End If
            End If
        Else
            blnSucess = True
        End If
        If Not blnErrAtuoRepaired Then
            If blnRetry And blnSucess Then '重试成功，则增加成功日志
                Me.WriteLog String(17, " ") & "重试结果：语句执行成功" & IIf(lngAffect > 0, "," & lngAffect & " 行数据生效", "")
            ElseIf blnSucess And Not mblnRecErrSQL Then '没有重试，且成功
                Me.WriteLog String(17, " ") & "执行结果：语句执行成功" & IIf(lngAffect > 0, "," & lngAffect & " 行数据生效", "")
            End If
            If mlngSQLRecTime <> 0 Then
                lngSQLTime = DateDiff("n", datBegin, datEnd)
                If lngSQLTime >= mlngSQLRecTime Then
                    If Not blnRetry And mblnRecErrSQL Then Me.WriteLog LogTime(True) & Format(objSQL.FileLine, "0000000") & ":" & GetLogSQL(objSQL)
                    Me.WriteLog String(17, " ") & "SQL处理耗时：" & lngSQLTime & "分钟"
                End If
            End If
        End If
        err.Clear: On Error GoTo errH
    End If
    RunSQLText = True
    Exit Function
errH: '函数内部的其他未知异常
    If MsgBox(err.Description, vbRetryCancel + vbCritical, gstrSysName) = vbRetry Then Resume
    Exit Function
AbortLine: '人工中止退出
    mstrAbortInfo = mobjSQL.File & "|" & mobjSQL.FileLine '记录中止文件及行号
End Function

Private Function LogTime(Optional blnOnlyTime As Boolean) As String
    If blnOnlyTime Then
        LogTime = Format(Now, "HH:mm:ss") & "，"
    Else
        LogTime = Format(Now, "yyyy-MM-dd HH:mm:ss") & "，"
    End If
End Function

Private Sub Class_Initialize()
    Set mobjSQL = New clsSQLInfo
End Sub

Private Sub Class_Terminate()
    Set mobjFSO = Nothing
    If Not mobjLog Is Nothing Then
        mobjLog.Close
        Set mobjLog = Nothing
    End If
    
    If Not mobjText Is Nothing Then
        mobjText.Close
        Set mobjText = Nothing
    End If
    
    Set mobjSQL = Nothing
    Erase marrDefine
End Sub

