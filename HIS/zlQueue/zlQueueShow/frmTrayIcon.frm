VERSION 5.00
Begin VB.Form frmTrayIcon 
   Caption         =   "����ͼ��"
   ClientHeight    =   1455
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   2070
   Icon            =   "frmTrayIcon.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1455
   ScaleWidth      =   2070
   StartUpPosition =   1  '����������
   Begin VB.Timer tmrBroadCast 
      Left            =   360
      Top             =   360
   End
   Begin VB.Menu mnuRight 
      Caption         =   "�Ҽ��˵�"
      Begin VB.Menu mnuSetup 
         Caption         =   "����"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "�˳�"
         Shortcut        =   ^Q
      End
   End
End
Attribute VB_Name = "frmTrayIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mobjIcon As clsTaskIcon
Attribute mobjIcon.VB_VarHelpID = -1
Private WithEvents mobjMsgCenter As clsQueueMsgCenter
Attribute mobjMsgCenter.VB_VarHelpID = -1

Private mblnUserMsgCenter As Boolean
Private mblnUseSound As Boolean

Private mdtLastVoiceDate As Date
Private mobjQueueManage As New clsQueueOperation
Private mlngLoopInterval As Long
Private mstrComputerName As String
'private

Private Sub Form_Load()
On Error GoTo ErrorHand
    '������ͼ��
    Set mobjIcon = New clsTaskIcon
    mobjIcon.frmHwnd = Me.hwnd ' hwnd
    mobjIcon.Icon = Me.Icon.Handle
    mobjIcon.Message = "�Ŷ���ʾ����"
    mobjIcon.AddIcon
    
    If gstrCompareVersion < "010.034.000" Then
        mblnUserMsgCenter = False
    Else
        mblnUserMsgCenter = Val(GetSetting("ZLSOFT", G_STR_REGPATH, "������Ϣ��������", 1)) = 1
    End If
    
    mblnUseSound = Val(GetSetting("ZLSOFT", G_STR_REGPATH, "������������", 1)) = 1
    
    '���ݲ����ж��Ƿ�������Ϣ����
    If mblnUserMsgCenter Then  '�ж��Ƿ���Ҫʹ����Ϣ����
    
        '������Ϣ����
        Call gobjComLib.ConnectMip(Me.hwnd)
        
        Set mobjMsgCenter = New clsQueueMsgCenter
        Call mobjMsgCenter.setComLib(gobjComLib)
        Call mobjMsgCenter.OpenMsgCenter(100, 1160, glngBusinessType)  '���һ��������Ҫ����ʵ�ʵ�ҵ������
        
        tmrBroadCast.Tag = 1
        tmrBroadCast.Enabled = False
    Else
        If mblnUseSound Then
            tmrBroadCast.Tag = 0
            tmrBroadCast.Enabled = True
        End If
    End If
    
    Call mobjQueueManage.setComLib(gobjComLib)
    
    'Ӧ����������
    Call ApplyVoiceConfig
    
    If Not mblnUserMsgCenter Then StartVoice

    Exit Sub
ErrorHand:
    MsgBox Err.Description, vbExclamation, gstrSysName
    Err.Clear
End Sub

Public Sub setMsgBusinessType(ByVal lngBusinessType As Long)
    If mobjMsgCenter Is Nothing Then Exit Sub
    Call mobjMsgCenter.ConfigMsgBusinessType(lngBusinessType)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo ErrorHand
    mobjIcon.MouseState x
    Exit Sub
ErrorHand:
    MsgBox Err.Description, vbExclamation, gstrSysName
    Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrorHand
    '�����������Ϣ��������Ҫ�Ͽ���Ϣƽ̨������
    If mblnUserMsgCenter Then
        Call gobjComLib.DisConnectMip
        Call mobjMsgCenter.CloseMsgCenter
    End If
    
    '�������ͼ��
    mobjIcon.DelIcon
    Set mobjIcon = Nothing
    Set mobjMsgCenter = Nothing
    Set mobjQueueManage = Nothing
    
    Call mnuQuit_Click
    
    Exit Sub
ErrorHand:
    MsgBox Err.Description, vbExclamation, gstrSysName
    Err.Clear
End Sub

'��ȫ�˳�
Private Sub mnuQuit_Click()
    Dim objForm As Form
On Error GoTo ErrorHand
    'ж�����д���
    For Each objForm In Forms
        Unload objForm
    Next
    
    Exit Sub
ErrorHand:
    MsgBox Err.Description, vbExclamation, gstrSysName
    Err.Clear
End Sub

'�����ô���
Private Sub mnuSetup_Click()
On Error GoTo ErrorHand
    frmMain.zlShowMe
    Exit Sub
ErrorHand:
    MsgBox Err.Description, vbExclamation, gstrSysName
    Err.Clear
End Sub

'�����ô���
Private Sub mobjIcon_MouseLeftDBClick()
On Error GoTo ErrorHand
    Call mnuSetup_Click
    Exit Sub
ErrorHand:
    MsgBox Err.Description, vbExclamation, gstrSysName
    Err.Clear
End Sub

'��ʾ�Ҽ��˵�
Private Sub mobjIcon_MouseRightUp()
On Error GoTo ErrorHand
    PopupMenu mnuRight
    Exit Sub
ErrorHand:
    MsgBox Err.Description, vbExclamation, gstrSysName
    Err.Clear
End Sub

Private Sub mobjMsgCenter_OnRecevieMsg(ByVal strMsgItemIdentity As String, ByVal strXmlContext As String, rsData As ADODB.Recordset)
'������յ�����Ϣ
    Dim strValue As String
    Dim i As Integer
    
    'G_STR_MSG_QUEUE_001:�Ŷ���Ϣ
    'G_STR_MSG_QUEUE_002:�����Ϣ
    'G_STR_MSG_QUEUE_003:״̬�ı���Ϣ
    'G_STR_MSG_QUEUE_004:���������Ϣ
On Error GoTo ErrorHand
        
    Select Case strMsgItemIdentity
        Case G_STR_MSG_QUEUE_001, G_STR_MSG_QUEUE_002, G_STR_MSG_QUEUE_003
            If SafeArrayGetDim(gobjStyleWindow) <= 0 Then Exit Sub
            
            '�Բ�ͬ��ʽ�����յ���Ϣ����ͬ������......
            For i = 1 To UBound(gobjStyleWindow)
                If Not gobjStyleWindow(i) Is Nothing Then
                    Call gobjStyleWindow(i).ISty_MsgProcess(gobjStyleWindow(i).ISty_WindNo, strMsgItemIdentity, strXmlContext, rsData)
                End If
            Next
            
        Case G_STR_MSG_QUEUE_004
            '���к��д���......
            If mblnUseSound Then
                If PlayVoice(rsData) Then
                    tmrBroadCast.Tag = 1
                    tmrBroadCast.Enabled = False
                Else    '����ʧ���������ѯ����
                    tmrBroadCast.Tag = 0
                    tmrBroadCast.Enabled = True
                End If
            End If
            
    End Select
    
    Exit Sub
ErrorHand:
    MsgBox Err.Description, vbExclamation, gstrSysName
    Err.Clear
End Sub

Private Sub tmrBroadCast_Timer()
On Error GoTo ErrorHand
    If Val(tmrBroadCast.Tag) = 1 Then Exit Sub
    
    'ֹͣ��ѵ
    Call AbortCall
    
    '������ѵ����
    Call LoopPlayVoice
    
    If Val(tmrBroadCast.Tag) = 1 Then Exit Sub
    
    '��ʼ��ѵ
    Call StartCall
    
    Exit Sub
ErrorHand:
    MsgBox Err.Description, vbExclamation, gstrSysName
    Err.Clear
End Sub

'''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''�������в��ִ���''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub StartVoice()
'��ʼ��������
    mdtLastVoiceDate = Timer - mlngLoopInterval
    
    tmrBroadCast.Interval = mlngLoopInterval * 1000
    tmrBroadCast.Enabled = True
    tmrBroadCast.Tag = 0
End Sub

Private Sub StopVoice()
'������������
On Error GoTo errHandle
    tmrBroadCast.Tag = 1
    tmrBroadCast.Enabled = False
    
    If Not mobjQueueManage Is Nothing Then Call mobjQueueManage.StopVoice
Exit Sub
errHandle:
    Debug.Print "StopVoice Err:" & Err.Description
End Sub

Private Sub ApplyVoiceConfig()
'Ӧ����������
    Dim str����վ������ As String
    
   '��ȡ�кŷ�ʽ
    If Val(GetSetting("ZLSOFT", G_STR_REGPATH & "\mlngWindowNo", "���ŷ�ʽ", 1)) = 1 Then
         str����վ������ = GetSetting("ZLSOFT", G_STR_REGPATH, "Զ�˺���վ��", "")
         
         If Trim(str����վ������) = "" Then str����վ������ = AnalyseComputer
    Else
        str����վ������ = AnalyseComputer
    End If
    
    mstrComputerName = AnalyseComputer
    
    mobjQueueManage.PlayStation = str����վ������
    mobjQueueManage.LocalStation = mstrComputerName
    mobjQueueManage.BusinessType = glngBusinessType
    mobjQueueManage.PlayTimeLength = Val(GetSetting("ZLSOFT", G_STR_REGPATH, "��������ʱ��", 15))
    mobjQueueManage.PlayCount = Val(GetSetting("ZLSOFT", G_STR_REGPATH, "�������Ŵ���", 2))
    mobjQueueManage.VoiceType = GetSetting("ZLSOFT", G_STR_REGPATH, "��������", "")
    mobjQueueManage.IsPlayHintSound = Val(GetSetting("ZLSOFT", G_STR_REGPATH, "��������ǰ������ʾ��", False))
    mobjQueueManage.PlaySpeed = Val(GetSetting("ZLSOFT", G_STR_REGPATH, "������������", 0))
    mobjQueueManage.UseVbsPlay = IIf(Val(GetSetting("ZLSOFT", G_STR_REGPATH, "����VBS�Զ������", 1)) = 0, False, True)
    mobjQueueManage.CusVoiceScript = GetSetting("ZLSOFT", G_STR_REGPATH, "VBS�ű�", "")
    
    mlngLoopInterval = Val(GetSetting("ZLSOFT", G_STR_REGPATH, "��ѯ���ʱ��", 30))
End Sub

Private Function PlayVoice(ByVal rsData As ADODB.Recordset) As Boolean
'���ܣ�������������
'����ֵ�����гɹ�����True,��֮����False
    Dim lngVoiceId As Long
    Dim lngQueueId As Long
    Dim strVoiceContext As String
    
    PlayVoice = False
    
    If rsData.RecordCount <= 0 Then Exit Function
    
    '�ж��Ƿ��Ǳ�վ���������
    rsData.Filter = "node_name='voice_station'"
    If rsData.RecordCount <= 0 Then Exit Function
    If Nvl(rsData!node_value) <> AnalyseComputer Then Exit Function
    
    rsData.Filter = "node_name='voice_context'"
    If rsData.RecordCount > 0 Then strVoiceContext = Nvl(rsData!node_value)
    
    rsData.Filter = "node_name='voice_id'"
    If rsData.RecordCount > 0 Then lngVoiceId = Val(rsData!node_value)
    
    rsData.Filter = "node_name='queue_id'"
    If rsData.RecordCount > 0 Then lngQueueId = Val(rsData!node_value)
    
    '��������
    PlayVoice = mobjQueueManage.PlayQueueVoice(mobjMsgCenter, lngVoiceId, lngQueueId, False, strVoiceContext)

    '���гɹ���ɾ�����й�������
    Call mobjQueueManage.DelVoiceData(lngVoiceId)
End Function

'��ѯ����
Private Sub LoopPlayVoice()
    Dim i As Integer
    Dim strSql As String
    Dim lngVoiceId As Long
    Dim lngQueueId As Long
    Dim strVoiceContext As String
    Dim blnAllowQuery As Boolean
    Dim rsVoiceContext As ADODB.Recordset  '�����ŵ��������ݼ�
    
    '�ж��Ƿ���Ҫ��ѯ���ݿ�����
    blnAllowQuery = IIf(rsVoiceContext Is Nothing, True, False)
    
    If Not rsVoiceContext Is Nothing Then
        blnAllowQuery = IIf(rsVoiceContext.RecordCount <= 0 Or rsVoiceContext.EOF, True, False)
    End If
    
    If blnAllowQuery Then
        If Timer < mdtLastVoiceDate + mlngLoopInterval Then Exit Sub
        mdtLastVoiceDate = Timer
        
        '��ѯ��Ҫ���ŵ���������
        If gstrCompareVersion < "010.034.000" Then
            strSql = "select id,����ID,�������� from �Ŷ���������  where վ��=[1] order by id"
        Else
            strSql = "select id,����ID,��������,����ʱ�� from �Ŷ���������  where վ��=[1] order by ����ʱ��"
        End If
        
        Set rsVoiceContext = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "��ѯ������������", mstrComputerName)
    End If
    
    If rsVoiceContext Is Nothing Then Exit Sub
    If rsVoiceContext.RecordCount <= 0 Or rsVoiceContext.EOF Then Exit Sub
    
    lngVoiceId = Val(Nvl(rsVoiceContext!ID))
    lngQueueId = Val(Nvl(rsVoiceContext!����ID))
    strVoiceContext = Nvl(rsVoiceContext!��������)
    
    Call rsVoiceContext.MoveNext
    
    'ˢ�½���������Ϣ
    If Not mblnUserMsgCenter Then
        If SafeArrayGetDim(gobjStyleWindow) > 0 Then
            For i = 1 To UBound(gobjStyleWindow)
                Call gobjStyleWindow(i).ISty_RefreshQueueData(lngQueueId)
            Next
        End If
    End If
    
    If lngQueueId <= 0 Then
        '�����Զ���ĺ�������
        Call mobjQueueManage.PlayCustomVoice(lngVoiceId, False, strVoiceContext)
    Else
        '��������
        Call mobjQueueManage.PlayQueueVoice(mobjMsgCenter, lngVoiceId, lngQueueId, False, strVoiceContext)
    End If

    '���гɹ���ɾ�����й�������
    Call mobjQueueManage.DelVoiceData(lngVoiceId)
End Sub

'��ʼ������ѯ��ʾ
Private Sub StartCall()
    tmrBroadCast.Enabled = True
End Sub

'��ֹ������ѯ��ʾ
Private Sub AbortCall()
    tmrBroadCast.Enabled = False
End Sub
