VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMipClientShell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'######################################################################################################################

'����
Private mstrModuleObjectKey As String
Private mstrMessageItemKey As String
Private mstrMspServerIP As String
Private mlngMspServerPort As Long
Private mstrMspUser As String
Private mstrMspUserPassword As String
Private mblnConnected As Boolean
Private mobjXML As Object
Private mblnRecordLog As Boolean
Private mintSaveLogDay As Integer
Private mstrLasDeleteDay As String
Private mclsMipRunLog As clsMipRunLog
Private mclsMipSystem As clsMipSystemData

Private mobjLogSQLite As clsDataSQLite
Private WithEvents mobjCommunicateProxy As zlMipClient.CommunicateProxy  '��Ϣ����ƽ̨�ӿڿ�
Attribute mobjCommunicateProxy.VB_VarHelpID = -1
Private WithEvents mfrmMipClientShell As frmMipClientShell
Attribute mfrmMipClientShell.VB_VarHelpID = -1

'�¼�
Public Event ReceiveMessage(ByVal strMsgContent As String)
Public Event ConnectStateChanged(ByVal IsConnected As Boolean)
Public Event OpenLink(ByVal strMsgKey As String, ByVal bytLinkType As Byte, ByVal strLinkPara As String, ByVal strModuleObjectKey As String)

'######################################################################################################################

Public Function Initialize(ByVal lngParentForm As Long, Optional ByVal cnOracle As ADODB.Connection) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '˵����
    '******************************************************************************************************************
    Dim strParentFolder As String
    Dim strSysFile As String
    Dim strDataFile As String
    Dim strLogFile As String
        
    Set gcnOracle = cnOracle
    
    gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "gstrSysName", "")
'    ParamInfo.��Ʒ���� = GetSetting("ZLSOFT", "ע����Ϣ", "��Ʒ����", "����")
    
    If Not (gcnOracle Is Nothing) Then Call InitCommon(gcnOracle)
    
    glngParentForm = lngParentForm
        
    mstrLasDeleteDay = ""
    Set mobjXML = InitXMLDoc
    Set gfrmMipResource = New frmMipResource
    
    If mfrmMipClientShell Is Nothing Then Set mfrmMipClientShell = New frmMipClientShell
            
    '�����ļ�
    '------------------------------------------------------------------------------------------------------------------
    strSysFile = App.Path & "\Data\zlMipClientShell_Sys.db"
    strDataFile = App.Path & "\Data\zlMipClientShell_Data.db"
    strLogFile = App.Path & "\Data\zlMipClientShell_Log.db"
    
    Set mclsMipRunLog = New clsMipRunLog
    If mclsMipRunLog.Initialize(strLogFile) = False Then Exit Function
    If mclsMipRunLog.OpenRunLogFile = False Then Exit Function
        
    Set mclsMipSystem = New clsMipSystemData
    If mclsMipSystem.Initialize(strSysFile) = False Then Exit Function
    If mclsMipSystem.OpenDataFile = False Then Exit Function
    
    Initialize = mfrmMipClientShell.Initialize(strSysFile, strDataFile, strLogFile)
    
    Call GetSystemOption
    
End Function

Public Function ConnectMip(ByVal strMipUser As String, ByVal strMipUserPass As String, ByVal strMipServerIP As String, ByVal lngMipServerPort As Long) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '˵����
    '******************************************************************************************************************
    
    On Error GoTo errHand
    
    mstrMspUser = strMipUser
    mstrMspUserPassword = strMipUserPass
    mstrMspServerIP = strMipServerIP
    mlngMspServerPort = lngMipServerPort
    
    If mobjCommunicateProxy Is Nothing Then
        Set mobjCommunicateProxy = New zlMipClient.CommunicateProxy
        Call mobjCommunicateProxy.Init(mstrMspServerIP, mlngMspServerPort)
    End If
    If mobjCommunicateProxy Is Nothing Then Exit Function
    
    Call mobjCommunicateProxy.Login(mstrMspUser, mstrMspUserPassword, "", , True)
    mblnConnected = mobjCommunicateProxy.IsConnected
    
    ConnectMip = mblnConnected
    
    If Not (mfrmMipClientShell Is Nothing) Then
        Call mfrmMipClientShell.UpdateConnectState(mblnConnected)
    End If
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
        
End Function

Public Function DisConnectMip() As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '˵����
    '******************************************************************************************************************
        
    On Error GoTo errHand
    
    If mblnConnected = False Then Exit Function
        
    If Not (mobjCommunicateProxy Is Nothing) Then
        Call mobjCommunicateProxy.Logout
        Set mobjCommunicateProxy = Nothing
    End If
    
    If Not (mfrmMipClientShell Is Nothing) Then
        Call mfrmMipClientShell.UpdateConnectState(mobjCommunicateProxy.IsConnected)
    End If
    
    DisConnectMip = True
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    
End Function

Public Function SendMessage(ByVal strXML As String, Optional ByVal strQueueName As String, Optional ByVal blnRequest As Boolean = False, Optional ByRef strRequestResult As String) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '˵����
    '******************************************************************************************************************
    If mobjCommunicateProxy Is Nothing Then Exit Function
    
    Call WriteRunLog("��Ϣ", "������Ϣ���ݵ���Ϣ����ƽ̨" & vbCrLf & strXML)
    
    If blnRequest = True Then
        strRequestResult = mobjCommunicateProxy.DoRequest(strXML, strQueueName)
        SendMessage = True
    Else
        SendMessage = mobjCommunicateProxy.SendMessage(strXML, strQueueName)
    End If
    
End Function

Public Sub ShowMessage(ByVal strMessageItemKey As String, ByVal strMessageContent As String, ByVal strModuleObjectKey As String)
    '******************************************************************************************************************
    '���ܣ�����ͨ�����������ִ�У���ð�ݵ������ѣ�
    '������
    '˵����
    '******************************************************************************************************************
    mstrModuleObjectKey = strModuleObjectKey
    mstrMessageItemKey = strMessageItemKey
    
    If Not (mfrmMipClientShell Is Nothing) Then
        Call mfrmMipClientShell.ShowMessage(strMessageContent)
    End If
    
End Sub

Public Property Get IsConnect() As Boolean
    IsConnect = mblnConnected
End Property

Public Function WriteRunLog(ByVal strLogType As String, ByVal strLogDesc As String) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim rsCondition As ADODB.Recordset
    Dim strStart As String
    Dim strEnd As String
    
    If mblnRecordLog = False Then Exit Function
    If Not (mclsMipRunLog Is Nothing) Then
        WriteRunLog = mclsMipRunLog.WriteRunLog(strLogType, strLogDesc)
    
        If (mstrLasDeleteDay = "" Or mstrLasDeleteDay <> Format(Now, "yyyy-MM-dd")) And mintSaveLogDay > 0 Then
            mstrLasDeleteDay = Format(Now, "yyyy-MM-dd")
            
            Set rsCondition = zlCommFun.CreateCondition
            
            strStart = "2014-01-01 00:00:00"
            strEnd = Format(DateAdd("d", 0 - mintSaveLogDay, CDate(Format(Now, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 23:59:59"
            
            Call zlCommFun.SetCondition(rsCondition, "��ʼʱ��", strStart)
            Call zlCommFun.SetCondition(rsCondition, "����ʱ��", strEnd)
                            
            Call mclsMipRunLog.EditRunLog("Delete", rsCondition)
            
        End If
    End If
    
End Function

'######################################################################################################################
Private Sub GetSystemOption()
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim rsCondition As New ADODB.Recordset
    Dim strPara As String
    Dim rsTmp As zlDataSQLite.SQLiteRecordset
    
    mblnRecordLog = False
    mintSaveLogDay = 7
    
    If Not (mclsMipSystem Is Nothing) Then
        Set rsCondition = CreateCondition
        Call SetCondition(rsCondition, "�������", "3")
        rsTmp = mclsMipSystem.GetPara("Filter", rsCondition)
        If rsTmp.DataSet.BOF = False Then
            strPara = NVL(rsTmp.DataSet("Para_Value").Value)
            mblnRecordLog = (Val(strPara) = 1)
        End If
        
        Call SetCondition(rsCondition, "�������", "4")
        rsTmp = mclsMipSystem.GetPara("Filter", rsCondition)
        If rsTmp.DataSet.BOF = False Then
            strPara = NVL(rsTmp.DataSet("Para_Value").Value)
            mintSaveLogDay = Val(strPara)
        End If
    End If
    
End Sub

Private Sub Class_Initialize()
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeBlue
End Sub

Private Sub Class_Terminate()
    
    On Error Resume Next
    
    If Not (mclsMipRunLog Is Nothing) Then Set mclsMipRunLog = Nothing
    If Not (mclsMipSystem Is Nothing) Then Set mclsMipSystem = Nothing
    
    If Not (mfrmMipClientShell Is Nothing) Then
        Unload mfrmMipClientShell
        Set mfrmMipClientShell = Nothing
    End If
    
    If Not (mobjCommunicateProxy Is Nothing) Then
        Call mobjCommunicateProxy.Logout
        Set mobjCommunicateProxy = Nothing
    End If
    
    If Not (gfrmMipResource Is Nothing) Then
        Unload gfrmMipResource
        Set gfrmMipResource = Nothing
    End If
    
End Sub

Private Sub mfrmMipClientShell_OpenLink(ByVal bytLinkType As Byte, ByVal strLinkPara As String)
    RaiseEvent OpenLink(mstrMessageItemKey, bytLinkType, strLinkPara, mstrModuleObjectKey)
End Sub

Private Sub mfrmMipClientShell_OptionChanged()
    '�����Ѿ��仯
    Call GetSystemOption
End Sub

Private Sub mobjCommunicateProxy_Error(ByVal Number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    
    If mblnConnected <> mobjCommunicateProxy.IsConnected Then
        mblnConnected = mobjCommunicateProxy.IsConnected
        If Not (mfrmMipClientShell Is Nothing) Then
            Call mfrmMipClientShell.UpdateConnectState(mblnConnected)
        End If
        RaiseEvent ConnectStateChanged(mblnConnected)
    End If
    
End Sub

Private Sub mobjCommunicateProxy_ErrorClosed(ByVal Number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    
    If mblnConnected <> mobjCommunicateProxy.IsConnected Then
        mblnConnected = mobjCommunicateProxy.IsConnected
        If Not (mfrmMipClientShell Is Nothing) Then
            Call mfrmMipClientShell.UpdateConnectState(mblnConnected)
        End If
        RaiseEvent ConnectStateChanged(mblnConnected)
    End If
    
End Sub

Private Sub mobjCommunicateProxy_MessageReceived(ByVal strMsg As String)
    RaiseEvent ReceiveMessage(strMsg)
End Sub

Private Sub mobjCommunicateProxy_ServerConnectStateChanged(ByVal IsConnected As Boolean)

    '������ӶϿ��ˣ��Զ���������
    
    On Error Resume Next
    If IsConnected = True Then Call mobjCommunicateProxy.ReLogin
    On Error GoTo 0
    
    '���֮ǰ��¼������״̬�͵�ǰ������״̬��һ�������Զ�����ͼ�겢����״̬�仯�¼�
    If mblnConnected <> mobjCommunicateProxy.IsConnected Then
        mblnConnected = mobjCommunicateProxy.IsConnected
        If Not (mfrmMipClientShell Is Nothing) Then
            Call mfrmMipClientShell.UpdateConnectState(mblnConnected)
        End If
        
        RaiseEvent ConnectStateChanged(IsConnected)
    End If
    
End Sub

Private Sub mobjCommunicateProxy_SessionKicked(ByVal strMsg As String)
    '
End Sub