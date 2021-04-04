Attribute VB_Name = "mdlDebug"
Option Explicit

Public Const M_STR_MODULE_MENU_TAG As String = "采集"
Public Const G_STR_HINT_TITLE As String = "提示"
Public Const G_STR_REG_PATH_PUBLIC As String = "公共模块\zl9PacsCapture"
Public Const G_STR_REG_PATH_PRIVATE As String = "私有模块\zl9PacsCapture\"

Public Enum TDockState                      '浮动窗口状态
    dsClosed = 0    '关闭
    dsOpen = 1      '打开
    dsClosing = 2   '关闭中
End Enum

Public Enum ReportType
    电子病历编辑器
    PACS报告编辑器
    报告文档编辑器
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

Public glngCurVideoContainerHwnd As Long        '当前视频所在的容器窗口句柄
Public gblnDockingState As TDockState           '是否处于弹出窗口状态

Public glngInstanceCount As Long
Public gblnOpenDebug As Boolean
Public gobjZOrder As Scripting.Dictionary
Public gblnIsQuitModule As Boolean

Public gstrHotKeyTest As String


Public Sub BUGEX(ByVal strDebug As String, Optional ByVal blnIsForce As Boolean = False)
    If gblnOpenDebug Or blnIsForce Then
        OutputDebugString Format(Now, "mmddhhmmss") & " |-> " & strDebug
    End If
End Sub

Public Sub InitCommonLib(cnOracle As ADODB.Connection)
    Set gcnVideoOracle = cnOracle
End Sub

Public Sub OutputDebug(ByVal strMethob As String, objErr As ErrObject)
    If gblnOpenDebug Then
        OutputDebugString "[" & App.ProductName & "]" & strMethob & "：" & objErr.Description
    End If
End Sub

Public Sub RaiseErr(objErr As ErrObject)
    Call err.Raise(objErr.Number, objErr.Source, objErr.Description, objErr.HelpFile, objErr.HelpContext)
End Sub
