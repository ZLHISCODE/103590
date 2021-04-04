VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGlobal"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum LoadCommand
    COMMAND_CAPIMG_UPLOAD           '采集图像上传
    COMMAND_RPTIMG_DOWNLOAD         '图像下载
End Enum

Public Enum FileLoadType    '文件上传方式
    Normal                  '普通
    Service                 'zlPacsServices
End Enum

Public Type TFtpDeviceInf
    strDeviceId As String
    strFTPIP As String
    strFTPUser As String
    strFTPPwd As String
    strFtpDir As String
    strSDDir As String
    strSDUser As String
    strSDPswd As String
End Type

Public Enum TMediaType
    imgTag = 0   '图像标记
    MULFRAMETAG = 1 '多侦图
    VIDEOTAG = 2 '视频标记
    AUDIOTAG = 3 '音频标记
End Enum

Public Sub InitPacsControlWithInPro(cnOracle As ADODB.Connection, _
    ByVal lngOwnerHwnd As Long, _
    ByVal lngSys As Long, _
    ByVal lngMoudle As Long, _
    Optional objOwner As Object = Nothing, _
    Optional blnOpenDebug As Boolean = False)
'初始化部件相关连接（用于同一进程）

'    Call InitCommon(cnOracle)
    
    Set gobjOwner = objOwner
    glngSys = lngSys
    glngMoudle = lngMoudle
    glngOwnerHwnd = lngOwnerHwnd
    gblnOpenDebug = blnOpenDebug
End Sub


Public Sub BUGEX(ByVal strDebug As String, Optional ByVal blnIsForce As Boolean = False)
'    If gblnOpenDebug Or blnIsForce Then
'        OutputDebugString Now & " |---> " & strDebug
'    End If
    Call mdlPublic.BUGEX(strDebug, blnIsForce)
End Sub

Public Function MsgboxEx(objParent As Object, sPrompt As String, Optional dwStyle As Long, Optional sTitle As String) As Long
    Dim lngHwnd As Long

BUGEX "MsgboxEx 1"

    If objParent Is Nothing Then
        lngHwnd = GetActiveWindow
    Else
        lngHwnd = objParent.hWnd
    End If

    If lngHwnd = GetDesktopWindow Or lngHwnd = 0 Then
BUGEX "MsgboxEx 2 GetForegroundWindow" & " DesktopWindowHwnd:" & lngHwnd
        lngHwnd = GetForegroundWindow
    End If

BUGEX "MsgBoxEx 3 Hwnd:" & lngHwnd

    MsgboxEx = mdlMsgBox.MsgboxEx(lngHwnd, sPrompt, dwStyle, sTitle)

    '当打开调试状态后，如果有错误信息，则自动提示
    If err.Number <> 0 And gblnOpenDebug Then
        Call mdlMsgBox.MsgboxEx(lngHwnd, "errSource:" & err.Source & "  errDescription:" & err.Description, vbOKOnly, CON_STR_HINT_TITLE)
    End If

BUGEX "MsgBoxEx End"
End Function


Private Sub Class_Initialize()
'初始化相关资源
    Set gcnOracle = Nothing
    Set gobjOwner = Nothing
    
    glngOwnerHwnd = 0
End Sub


Private Sub Class_Terminate()
'释放相关资源
    
End Sub
