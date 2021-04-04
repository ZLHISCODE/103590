VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsArchive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'���ģ�����
Private mfrmInMedRec As Object
Private mfrmOutMedRec As frmArchiveOutMedRec
Private mstZY As MedPage_Standard                'סԺ��ҳ��׼,��ֹ��ҳ��׼�����������չʾ���ݵ�����
Private mclsProperty As clsProperty                   '�����࣬�������沿�ֱ�������������
'��ҳ�Ƿ��Ǵ�״̬��ֻ������
Public Property Get IsOpen() As Boolean
    IsOpen = mclsProperty.IsOpen
End Property
'��ҳ�Ƿ�ϲ���,ֻ������
Public Property Get IsMerged() As Boolean
    IsMerged = mclsProperty.IsMerged
End Property
'��Ժ��Ҫ����Ƿ����ı�,ֻ������
Public Property Get IsDiagChange() As Boolean
    IsDiagChange = mclsProperty.IsDiagChange
End Property
'IsDiagInput:ָ����ϵ���ҳ����Ƿ�����
Public Property Get IsDiagInput() As Boolean
    IsDiagInput = mclsProperty.IsDiagInput
End Property

Public Sub InitArchiveMedRec(cnMain As ADODB.Connection, ByVal lngSys As Long, Optional ByVal lngModule As Long, Optional ByRef objMip As zl9ComLib.clsMipModule)
'���ܣ���ʼ��
'������objMip=��Ϣ����
    Set gcnOracle = cnMain
    gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "gstrSysName", "")
    gstrProductName = GetSetting("ZLSOFT", "ע����Ϣ", "��Ʒ����", "����")
    gstrUnitName = GetUnitName
    If Not (objMip Is Nothing) Then Set gclsMipModule = objMip
    Set mclsProperty = New clsProperty
    Call mclsProperty.InitProperty(f���Ӳ���, lngSys, lngModule, objMip)
End Sub

Private Sub Class_Terminate()
    If Not mfrmInMedRec Is Nothing Then
        Unload mfrmInMedRec
        Set mfrmInMedRec = Nothing
    End If
    If Not mfrmOutMedRec Is Nothing Then
        Unload mfrmOutMedRec
        Set mfrmOutMedRec = Nothing
    End If
    Set mclsProperty = Nothing
End Sub

Public Function zlGetForm(ByVal intType As Integer) As Object
'���ܣ���ȡ��ҵ�񴰿�,���������ʼ��ʱֻ����һ�λ�ȡ
'������intType 0-���ﲡ������ 1-סԺ��������
    If Not gclsMain Is Me Or gclsMain Is Nothing Then
        Set gclsMain = Me
    End If
    If Not gclsPros Is mclsProperty Or gclsPros Is Nothing Then
        Set gclsPros = mclsProperty
    End If
    If intType + 1 = PF_���� Then
        mclsProperty.PatiType = PF_����
        mclsProperty.MedPageSandard = ST_������ҳ
        If mfrmOutMedRec Is Nothing Then
            Set mfrmOutMedRec = New frmArchiveOutMedRec
        End If
        Set zlGetForm = mfrmOutMedRec
        Set mclsProperty.CurrentForm = mfrmOutMedRec
    Else
        mclsProperty.PatiType = PF_סԺ
        mclsProperty.MedPageSandard = Val(zlDatabase.GetPara("������ҳ��׼", mclsProperty.SysNo, pסԺҽ��վ, "0")) + 1
        If mfrmInMedRec Is Nothing Then
            mstZY = mclsProperty.MedPageSandard
            '������ҳ��׼
            Select Case mstZY
                Case ST_��������׼  '��������׼
                    Set mfrmInMedRec = New frmArchiveInMedRec
                Case ST_�Ĵ�ʡ��׼   '�Ĵ�ʡ��׼
                    Set mfrmInMedRec = New frmArchiveInMedRec_SC
                Case ST_����ʡ��׼    '����ʡ��׼
                    Set mfrmInMedRec = New frmArchiveInMedRec_YN
                Case ST_����ʡ��׼    '����ʡ��׼
                    Set mfrmInMedRec = New frmArchiveInMedRec_HN
            End Select
        End If
        Set zlGetForm = mfrmInMedRec
        Set mclsProperty.CurrentForm = mfrmInMedRec
    End If
End Function

Public Function zlRefresh(ByVal intType As Integer, ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal blnMoved As Boolean) As Boolean
'���ܣ�ˢ�»����ҽ���嵥
'������intType 0-���ﲡ������ 1-סԺ��������
'      lng��ҳID :סԺ��ҳ-��ҳID�� ������ҳ-�Һ�ID
'      blnMoved �Ƿ�ת��
    If intType + 1 = PF_���� Then
        mclsProperty.PatiType = PF_����
        mclsProperty.MedPageSandard = ST_������ҳ
        Set mclsProperty.CurrentForm = mfrmOutMedRec
    Else
        mclsProperty.PatiType = PF_סԺ
        mclsProperty.MedPageSandard = mstZY
        Set mclsProperty.CurrentForm = mfrmInMedRec
    End If
    mclsProperty.����ID = lng����ID
    mclsProperty.��ҳID = lng��ҳID
    Set gclsPros = mclsProperty
    Call ArchivezlRefresh
End Function

Public Function CloseWindows() As Boolean
    '--------------------------------------
    '����:�ر������Ӵ���
    '--------------------------------------
    Dim frmThis As Form
    Err = 0: On Error Resume Next
    For Each frmThis In Forms
        Unload frmThis
    Next
    CloseWindows = (Forms.Count = 0)
End Function








