VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPublicPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public objDockPath As zlPublicPath.clsDockPath

Public Function InitCommon(ByVal cnOracle As ADODB.Connection, ByVal lngSys As Long, Optional ByVal str站点 As String, Optional ByRef frmMain As Object, _
        Optional ByVal lngModule As Long, Optional ByVal MainPrivs As String, Optional ByRef objEmr As Object) As Boolean
'功能：初始化医嘱公共部件
'参数：MainPrivs=主界面权限串
'     frmMain =主界面窗体对象
'     mlngModule=主界面模块号
'     objMip=消息平台对象
    Set gcnOracle = cnOracle
    If zlGetComLib = False Then Exit Function
    Call GetUserInfo
    glngSys = lngSys
    glngModule = lngModule
    gMainPrivs = MainPrivs
    gstrNodeNo = str站点

    On Error Resume Next
    If gobjPath Is Nothing Then
        Set gobjPath = New zlCISPath.clsCISPath
        Call gobjPath.InitCISPath(cnOracle, frmMain, lngSys, MainPrivs, objEmr)
    End If
    
    InitCommon = True
End Function

Private Sub Class_Initialize()
    If objDockPath Is Nothing Then Set objDockPath = New zlPublicPath.clsDockPath
End Sub

Public Sub zlExecuteCode(ByVal lngSys As Long, ByVal lngModul As Long, ByVal frmParent As Variant, ByVal strPrivs As String)
    Call gobjPath.zlExecuteCode(lngSys, lngModul, frmParent, strPrivs)
End Sub

Public Function CloseWindows() As Boolean
    If gobjPath Is Nothing Then CloseWindows = True: Exit Function
    CloseWindows = gobjPath.CloseWindows
End Function

Public Function zlImportInPathForEMR(ByVal cnOracle As ADODB.Connection, ByVal lngHwnd As Long, ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Boolean
'功能：供新版病历调用路径导入
'参数：父窗体句柄
    InitCommon cnOracle, 100
    objDockPath.zlRefreshReadOnly lng病人ID, lng主页ID
    objDockPath.zlImportPath lngHwnd
End Function
