VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMipSendMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'######################################################################################################################

Private WithEvents mobjCommunicateProxy As zlMipClient.CommunicateProxy  '消息服务平台接口库
Attribute mobjCommunicateProxy.VB_VarHelpID = -1
Private mblnConnected As Boolean                                            '记录是否连接到消息服务平台
Private mstrMspUser As String                                               '登录消息服务平台的用户
Private mstrError As String

'######################################################################################################################
Public Function ConnectMsp(ByVal strMspUser As String, ByVal strMspUserPassword As String, _
                            ByVal strMspServerIP As String, ByVal lngMspServerPort As Long) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '说明：
    '******************************************************************************************************************
    On Error GoTo errHand
        
    mstrMspUser = strMspUser
            
    If mobjCommunicateProxy Is Nothing Then

        Set mobjCommunicateProxy = New zlMipClient.CommunicateProxy
        Call mobjCommunicateProxy.Init(strMspServerIP, lngMspServerPort)
    End If
    
    If mobjCommunicateProxy Is Nothing Then Exit Function
    
    Call mobjCommunicateProxy.Login(strMspUser, strMspUserPassword, "", , True)
        
    mblnConnected = True
    
    ConnectMsp = True
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    mstrError = Err.Description
End Function

Public Function GetLastError() As String
    GetLastError = mstrError
End Function

Public Function SendMessage(ByVal strMsgContent As String, Optional ByVal blnRequest As Boolean = False) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '说明：
    '******************************************************************************************************************
    Dim strQueueName As String
    Dim strRequestResult As String
        
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
    mstrError = Err.Description
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




