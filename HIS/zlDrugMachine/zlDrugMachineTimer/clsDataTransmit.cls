VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDataTransmit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'-----------------------------------------
'˵������ʱ���ݴ�����
'���ߣ�YuZhiYong
'���䣺yzyong@zlsoft.cn
'-----------------------------------------

'�Զ����¼���������־��Ϣ����������
Public Event AfterTransmit(ByVal strLog As String)

Private WithEvents mfrmMain As frmMain                  '���������
Attribute mfrmMain.VB_VarHelpID = -1
Private ParamsInfo As mdlDefine.TYPE_PARAMS             '����
Private mblnAction As Boolean
Private mblnTransmitting As Boolean
Private mobjComLib As Object
Private mclsLog As clsLog
Private mstrData As String
Private mstrUser As String

Public Sub Transmit(ByVal strData As String)
'���ܣ��������ݵĴ���
'������
'  strData�����͵�����
'    ��ʽ��[�ӿڱ��]|[ҵ�����]|[ҵ������]
'    �ӿڱ�ţ��ӿڱ��1[;�ӿڱ��...]
'    ҵ�����
'        1-������Ϣ��
'        2-��Ա��Ϣ��
'        3-ҩƷĿ¼��
'        4-ҩƷ������λ��
'        5-��ҩ���ڣ�
'    ҵ�����ݣ�
'        1-���ţ���������1;��������2;��
'        2-��Ա����Ա����1;��Ա����2;��
'        3-ҩƷ������1;����2;��
'        4-��棺�ⷿid1;�ⷿid2;��
'        5-���ڣ��ⷿid1;�ⷿid2;��

    Dim strTmp As String

    If mfrmMain.timTransmit.Tag = "1" Then
        'Timer����ִ��
        strTmp = "��ʱ�������ڴ��ͣ��������ݴ��ͽ��Ժ����"
        mclsLog.Add strTmp
        mclsLog.Add strData, , Val("1-��ϸ��־")
        RaiseEvent AfterTransmit(strTmp)
        mfrmMain.BaseData = strData
    Else
        'Timerδ��ִ��
        mclsLog.Add vbNewLine & "" & Now
        
        strTmp = "��ʼ���ͻ������ݣ���ʱ����ͣ"
        mclsLog.Add strTmp
        mclsLog.Add strData, , Val("1-��ϸ��־")
        RaiseEvent AfterTransmit(strTmp)
        
        mblnTransmitting = True
        mfrmMain.timTransmit.Enabled = False    '��ͣTimer�¼�
        
        Call mfrmMain.TransBase(strData)
        
        mfrmMain.timTransmit.Enabled = True     '�ָ�Timer�¼�
        mblnTransmitting = False
        
        strTmp = "��ɴ��ͻ������ݣ���ʱ���ָ�"
        mclsLog.Add strTmp
        RaiseEvent AfterTransmit(strTmp)
    End If

    mclsLog.Save

End Sub

Public Property Get Transmitting() As Boolean
'�������ڴ���
    Transmitting = mblnTransmitting
End Property

Public Property Get SupportData() As String
'��ʱ���͵�ҵ�������б��ַ�����ֻ����
    If mfrmMain Is Nothing Then
        SupportData = ""
    Else
        SupportData = mfrmMain.SupportData
    End If
End Property

Public Property Get TimerAction() As Boolean
'��ʱ������ͣ
    TimerAction = mblnAction
End Property
Public Property Let TimerAction(ByVal blnAction As Boolean)
    mblnAction = blnAction
    
    If mblnAction Then
    
        '������ʱ����
        Set mfrmMain = New frmMain
        If mfrmMain.ShowMe(mstrUser, mobjComLib, mclsLog) = False Then
            mblnAction = False
            GoTo makClose
        End If
        
        Call mfrmMain.ReadParams
        
        '������ʱ��
        With mfrmMain.timTransmit
            .Enabled = False
            .Interval = Val("60000/����")
            .Enabled = True
        End With
        mfrmMain.Tag = Timer()  '��ʼʱ��
        
    Else
    
makClose:
        If Not mfrmMain Is Nothing Then
            mfrmMain.timTransmit.Enabled = False
        End If
        Set mfrmMain = Nothing
        
    End If
        
End Property

Public Function Init(ByVal strUser As String, ByRef objComLib As Object) As Boolean
'������
'  strUser��ZLHIS�û���

    Dim strTmp As String
    
    mstrUser = strUser
    
    If Trim(mstrUser) = "" Then
        strTmp = "��ʼ����ʱ����ʱ��δ����ZLHIS�û���"
        RaiseEvent AfterTransmit(strTmp)
        Exit Function
    End If
    
    If objComLib Is Nothing Then
        strTmp = "��ʼ����ʱ����ʱ����ȷ��zl9ComLib�����Ѵ���"
        RaiseEvent AfterTransmit(strTmp)
        Exit Function
    End If
    
    Set mobjComLib = objComLib
End Function

Public Sub ReadParams()
'���ܣ���ȡ�����������浽����
    
    Dim objXML As New clsXML
    Dim strFile As String, strTmp As String
    
    If LCase(App.Path) Like "*\apply" Then
        strFile = App.Path & "\zlDrugMachine.cfg"
    ElseIf LCase(App.Path) Like "*\apply\*" Then
        strFile = Left(App.Path, InStr(LCase(App.Path), "\apply\") + Len("\apply\") - 1) & "zlDrugMachine.cfg"
    ElseIf LCase(App.Path) Like "*zldrugmachinemanage*" Or LCase(App.Path) Like "*zldrugmachine\*" Or LCase(App.Path) Like "*zldrugmachine" Then
        strFile = Replace(App.Path, App.EXEName, "") & "zlDrugMachineManage\zlDrugMachine.cfg"
    End If
    
    If objXML.OpenXMLFile(strFile) = False Then
        strTmp = "�������ߵĲ����ļ�����ȷ��" & vbNewLine & strFile
        RaiseEvent AfterTransmit(strTmp)
    End If

    With ParamsInfo
        .��ʱ���� = Val(GetParameter(objXML, "cycle", "5"))
        .��Ч���� = Val(GetParameter(objXML, "validdays", "2"))
        .��ʾ������� = Val(GetParameter(objXML, "viewlines", "200"))
        .�����־ = Val(GetParameter(objXML, "output", "0")) = 1
        .��ϸ��־ = Val(GetParameter(objXML, "detailed", "0")) = 1
        .������־���� = Val(GetParameter(objXML, "savedays", "7"))
        .ҵ������ = Trim(GetParameter(objXML, "businessdata", ""))
        
        If .��Ч���� > 7 Then .��Ч���� = 5
    End With
    
    objXML.CloseXMLDocument
    Set objXML = Nothing
    
    'ͬʱ���´���Ĳ���ֵ
    If Not mfrmMain Is Nothing Then
        Call mfrmMain.ReadParams
    End If
End Sub

Private Sub Class_Initialize()
    Set mfrmMain = New frmMain

    Call ReadParams
    Set mclsLog = New clsLog
    With mclsLog
        .Path = App.Path
        .FileName = App.EXEName
        .Action = ParamsInfo.�����־
        .Type_ = IIf(ParamsInfo.��ϸ��־, 1, 0)
        .SaveDays = ParamsInfo.������־����
    End With
End Sub

Private Sub Class_Terminate()
    Dim frmThis As Form
    
    mclsLog.Save
    
    Err = 0: On Error Resume Next
    For Each frmThis In Forms
        Unload frmThis
    Next
    
    Set mfrmMain = Nothing
    Set mclsLog = Nothing
    Set mobjComLib = Nothing
End Sub

Private Sub mfrmMain_AfterTransmit(ByVal strLog As String, ByVal blnTransmitting As Boolean)
    mblnTransmitting = blnTransmitting
    RaiseEvent AfterTransmit(strLog)
End Sub


 