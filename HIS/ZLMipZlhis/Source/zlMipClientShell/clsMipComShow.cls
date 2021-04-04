VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMspClientShell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'######################################################################################################################

Private WithEvents mfrmMspComMessage As frmMspComMessage
Attribute mfrmMspComMessage.VB_VarHelpID = -1

Public Event OpenLink(ByVal bytLinkType As Byte, ByVal strLinkPara As String)

'######################################################################################################################

Public Function Initialize(ByVal lngParentForm As Long) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '说明：
    '******************************************************************************************************************
    
    glngParentForm = lngParentForm
    If mfrmMspComMessage Is Nothing Then Set mfrmMspComMessage = New frmMspComMessage
    Initialize = mfrmMspComMessage.Initialize
    
End Function

Public Sub ShowMessage(ByVal strMessageContent As String)
    '******************************************************************************************************************
    '功能：接收到消息，然后进行弹出显示
    '参数：
    '说明：
    '******************************************************************************************************************
    If Not (mfrmMspComMessage Is Nothing) Then
        Call mfrmMspComMessage.ShowMessage(strMessageContent)
    End If
End Sub

Private Sub Class_Initialize()
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeBlue
End Sub

Private Sub Class_Terminate()
    If Not (mfrmMspComMessage Is Nothing) Then
        Unload mfrmMspComMessage
        Set mfrmMspComMessage = Nothing
    End If
End Sub

Private Sub mfrmMspComMessage_OpenLink(ByVal bytLinkType As Byte, ByVal strLinkPara As String)
    RaiseEvent OpenLink(bytLinkType, strLinkPara)
End Sub
