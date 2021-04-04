Attribute VB_Name = "mdlDebug"
Option Explicit

#Const DebugVer = True

Public Const M_STR_MODULE_MENU_TAG As String = "采集"
Public Const G_STR_HINT_TITLE As String = "提示"

Public Const G_STR_REG_PATH_PUBLIC As String = "公共模块\zl9PacsCapture"
Public Const G_STR_REG_PATH_PRIVATE As String = "私有模块\zl9PacsCapture\"


Public Enum TDockState                      '浮动窗口状态
    dsClosed = 0    '关闭
    dsOpen = 1      '打开
    dsClosing = 2   '关闭中
End Enum


Public gcnVideoOracle As ADODB.Connection        '公共数据库连接，特别注意：不能设置为新的实例

Public gobjOwner As Object
Public glngRootHandle As Long
Public glngSys As Long
Public gstrPrivs As String                  '当前用户具有的当前模块的功能
Public gstrSysName As String                '系统名称
Public glngModule As Long                   '模块号
Public glngDepartId As Long                 '当前科室ID

Public gobjCapturePar As clsCaptureParameter    '视频采集相关参数对象

Public gobjNotifyEvent As clsNotifyEvent         '消息通知发布对象
Public gobjVideo As frmWork_Video                '视频采集对象
'Public gobjGlobal As clsGlobal

Public glngCurVideoContainerHwnd As Long        '当前视频所在的容器窗口句柄
'Public glngNextVideoContainerHwnd As Long       '当前视频所在窗口Z序列的下一序列窗口句柄
Public gblnDockingState As TDockState           '是否处于弹出窗口状态

Public glngInstanceCount As Long
Public gblnOpenDebug As Boolean
Public gobjZOrder As Scripting.Dictionary
Public gblnIsQuitModule As Boolean

'debug property
Public gstrHotKeyTest As String




Private Function IsDebugMode() As Boolean

    IsDebugMode = False

    On Error Resume Next

    Debug.Print 1 / 0

    If err.Number <> 0 Then

        IsDebugMode = True

    End If

End Function



Public Sub BUGEX(ByVal strDebug As String, Optional ByVal blnIsForce As Boolean = False)
    If gblnOpenDebug Or blnIsForce Then
        OutputDebugString Format(Now, "mmddhhmmss") & " |-> " & strDebug
    End If
End Sub

Public Sub Main()
BUGEX "Main 1", True
    If UCase(Command()) = "DEBUG" Or IsDebugMode Then
BUGEX "Main Enter Debug", True
        frmTestLogin.Show
    End If

BUGEX "Main End", True




'    If Not IsDebugMode Then Exit Sub
'
'    Set gcnOracle = New ADODB.Connection
'
'    OraDataOpen "", "zlhis", "HIS"
'
'    Set gcnVideoOracle = gcnOracle
'
'    frmTestWindow.Show

End Sub


Public Sub InitCommonLib(cnOracle As ADODB.Connection)
'初始化部件相关(用于进程外项目)
    Dim blnIsEqualDB As Boolean
    
    If cnOracle Is Nothing Then Exit Sub
    
 
    If gobjComLib Is Nothing Then
        'Set gobjComLib = zl9ComLib.clsComLib
        Set gobjComLib = CreateObject("zl9ComLib.clsComLib")
    End If

    blnIsEqualDB = False
    If Not gcnVideoOracle Is Nothing Then
        blnIsEqualDB = IIf(gcnVideoOracle.ConnectionString = cnOracle.ConnectionString, True, False)
    End If
    
    '如果连接不同，则需要重新创建连接
    If Not blnIsEqualDB Then
        Set gcnVideoOracle = Nothing
        
        '当数据库连接改变时，重新创建连接
        Set gcnVideoOracle = New ADODB.Connection
            
        '注：可能由于ActiveExe为单独的进程项目，因此不能使用cnOracle直接对gcnVideoOracle对象赋值，否则产生“参数类型不正确,XXX”的错误
        gcnVideoOracle.ConnectionString = cnOracle.ConnectionString
            
        '打开数据库连接
        gcnVideoOracle.Open
    Else
        Exit Sub
    End If
    
    If gobjComLib.gstrNodeNo <> "" Then Exit Sub
    
    Call zlCL_InitCommon(gcnVideoOracle)
    Call zlCl_RegCheck
End Sub


Public Sub OutputDebug(ByVal strMethob As String, objErr As ErrObject)
    If gblnOpenDebug Then
        OutputDebugString "[" & App.ProductName & "]" & strMethob & "：" & objErr.Description
    End If
End Sub


'Public Sub BUGEX(ByVal strDebug As String, Optional ByVal blnIsForce As Boolean = False)
'    If gblnOpenDebug Or blnIsForce Then
'        OutputDebugString Now & " |---> " & strDebug
'    End If
'End Sub

Public Sub RaiseErr(objErr As ErrObject)
    Call err.Raise(objErr.Number, objErr.Source, objErr.Description, objErr.HelpFile, objErr.HelpContext)
End Sub
