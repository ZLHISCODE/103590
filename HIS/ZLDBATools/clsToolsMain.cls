VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsToolsMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Function InitConnection(ByVal blnIsDba As Boolean, ByVal cnConnection As ADODB.Connection) As Boolean
    '功能： 初始化数据库连接对象,成功则返回True
    '参数：-IsDba：是否是DBA用户    -cnConnection：数据库连接对象
    Dim blnResult As Boolean
    
    '非DBA或连接对象为空
    If Not blnIsDba Or cnConnection Is Nothing Then
        blnResult = False
    Else
        blnResult = True
    End If
    
    InitConnection = blnResult
End Function


Private Function InitParameters(ByVal blnIsDba As Boolean, ByVal cnConnection As ADODB.Connection, ByVal strUserName As String, _
                                                ByVal strPwd As String) As Boolean
    '功能:初始化相关全局变量。
    '参数：-IsDba：是否是DBA用户    -cnConnection：数据库连接对象   -strUserName：用户名    -strPwd：密码
    Dim strFileName As String, strPName As String * 256
    
    On Error GoTo errH
    
    '首先进行一个简单的判断，如果不是ZL开头的程序或者调试程序，那么直接退出
    Call GetModuleFileName(0, strPName, 256)
    strFileName = Left(strPName, InStr(strPName, Chr(0)) - 1)
    strFileName = UCase(Mid(strFileName, InStrRev(strFileName, "\") + 1))

    If strFileName <> "VB6.EXE" Then
        If (Not UCase(strFileName) Like "ZL*") Or GetSetting("ZLSOFT", "公共全局", "执行文件") = "" Then
            InitParameters = False
            Exit Function
        End If
    End If

    '如果数据库连接初始化失败或非DBA用户,返回False
    If Not InitConnection(blnIsDba, cnConnection) Then
        InitParameters = False
        Exit Function
    End If
    
    If gcnOracle = cnConnection Then
        gcnOracle.CursorLocation = adUseClient
        InitParameters = True
        Exit Function
    Else
        gcnOracle.CursorLocation = adUseClient
        Set gcnOracle = cnConnection
    End If
    
    
    '初始化参数
    gblnRAC = CheckRAC(gintInstId)
    gintCpuCount = GetCpuCount(gintCpuAdvise, gintCpuMax) '获取并行度
    gblnIsZlhis = CheckTblExist("ZLSYSTEMS")
    gstrBigVer = GetOracleVersion
    gstrVerNum = GetOracleVersion(True)
    gstrUserName = strUserName
    gstrPassword = strPwd
    gblnHasBigtables = CheckTblExist("ZLBIGTABLES")
    gblnHasZltables = CheckTblExist("ZLTABLES")
    Call CreateStr2list
    Call CreateList2str
    
    InitParameters = True
    Exit Function
errH:
    ErrCenter
    InitParameters = False
End Function

Public Function GetFrmByMdoudle(ByVal strMoudleNum As String, ByVal blnIsDba As Boolean, ByVal cnConnection As ADODB.Connection, ByVal strUserName As String, _
                                                ByVal strPwd As String) As Object
    Dim objForm As Object
                                                
    If Not InitParameters(blnIsDba, cnConnection, strUserName, strPwd) Then
        Set objForm = Nothing
        Exit Function
    End If
    
    Select Case strMoudleNum
        Case "0601"   '性能监控
            Set objForm = New frmMonitorMain
        Case "0602"  'SQL性能优化
            Set objForm = New frmTunning
        Case "0604"     '会话解锁
            Set objForm = New frmKillBlockers
        Case "0605"     '外键补建
            Set objForm = New frmIdxInfo
        Case "0606"     '空间整理
            Set objForm = New frmReused
    End Select
    
    Set GetFrmByMdoudle = objForm
End Function
