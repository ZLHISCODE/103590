VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMspSendMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'######################################################################################################################

Private WithEvents mobjCommunicateProxy As zlMspInterface.CommunicateProxy  '消息服务平台接口库
Attribute mobjCommunicateProxy.VB_VarHelpID = -1
Private mblnConnected As Boolean                                            '记录是否连接到消息服务平台
Private mstrMsgUser As String                                               '登录消息服务平台的用户
Private mstrDbUser As String                                                '登录ZLHIS的DB用户
'Private mfrmMspClient As frmMspClient

'######################################################################################################################
Public Function ConnectMsp() As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '说明：
    '******************************************************************************************************************
    On Error GoTo errHand
        
    If mobjCommunicateProxy Is Nothing Then
        Set mobjCommunicateProxy = New zlMspInterface.CommunicateProxy
        Call mobjCommunicateProxy.Init("192.168.111.132", 6066)
    End If
    
    If mobjCommunicateProxy Is Nothing Then Exit Function
    
    mstrMsgUser = "frchen"
    Call mobjCommunicateProxy.Login(mstrMsgUser, "123456", "")
        
    mblnConnected = True
    
    ConnectMsp = True
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
'    If ErrCenter = 1 Then
'        Resume
'    End If
'    Call SaveErrLog
    MsgBox Err.Description & vbCrLf & "源:" & Err.Source, vbOKOnly, "错误" & Err.Number, Err.HelpFile, Err.HelpContext
End Function

Public Function SendMessage(ByVal strMsgIdentity As String, ByVal strMsgContent As String, Optional ByRef strRequestResult As String) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '说明：
    '******************************************************************************************************************
    Dim strQueueName As String
    Dim blnRequest As Boolean
    
    On Error GoTo errHand
        
    If Not (mobjCommunicateProxy Is Nothing) Then
        
        If blnRequest = True Then
            strRequestResult = mobjCommunicateProxy.DoRequest(strMsgContent, strQueueName)
            SendMessage = True
        Else
            SendMessage = mobjCommunicateProxy.SendMessage(strMsgContent, strQueueName)
        End If

    End If

    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    MsgBox Err.Description & vbCrLf & "源:" & Err.Source, vbOKOnly, "错误" & Err.Number, Err.HelpFile, Err.HelpContext
End Function

Public Function DisConnectMsp() As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '说明：
    '******************************************************************************************************************
    Dim intLoop As Integer
    
    On Error GoTo errHand
        
    If Not (mobjCommunicateProxy Is Nothing) Then
        Call mobjCommunicateProxy.Logout
        Set mobjCommunicateProxy = Nothing
    End If
    
    DisConnectMsp = True
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    MsgBox Err.Description & vbCrLf & "源:" & Err.Source, vbOKOnly, "错误" & Err.Number, Err.HelpFile, Err.HelpContext
End Function

Private Sub Class_Initialize()
    '
End Sub

Private Sub mobjCommunicateProxy_Error(ByVal Number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    'MsgBox Description & vbCrLf & "源:" & Source, vbOKOnly, "错误" & Number, HelpFile, HelpContext
'    Call OutputTextLine("#Error#" & Source & "#" & Description)
End Sub

Private Sub mobjCommunicateProxy_ErrorClosed(ByVal Number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    
    mblnConnected = False
    
End Sub

Private Sub mobjCommunicateProxy_ServerConnectStateChanged(ByVal IsConnected As Boolean)
        
    If IsConnected And (Not mobjCommunicateProxy Is Nothing) Then
        Call mobjCommunicateProxy.ReLogin
    End If
    
End Sub




