VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDoctorShift"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Version As String

Public Sub CodeMan(ByVal lngSys As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, frmMain As Object, ByVal strDbUser As String)
    '------------------------------------------------
    '���ܣ� ����������ָ�����ܣ�����ִ����س���
    '������
    '   lngModul:��Ҫִ�еĹ������
    '   cnMain:����������ݿ�����
    '   frmMain:������
    Set gcnOracle = cnMain
    glngSys = lngSys
    gstrDbaUser = strDbUser
    glngModul = lngModul
    gstrPrivs = GetPrivFunc(lngSys, lngModul)
    gstrProductName = GetSetting("ZLSOFT", "ע����Ϣ", "��Ʒ����", "����")
    
    On Error Resume Next
    Set gobjEmr = frmMain.mobjEMR  '�²�����Ҫ����
    
    Call InitObjPublicAdvice
    If Err.Number <> 0 Then
        Set gobjEmr = Nothing
        Err.Clear
    End If
    
    Select Case lngModul
        Case 1242
            frmDoctorShift.Show , frmMain
    End Select
End Sub

Private Sub Class_Initialize()
    Version = App.Major & "." & App.Minor & "." & App.Revision
End Sub


Private Sub Class_Terminate()
    If Not gobjPublicAdvice Is Nothing Then Set gobjPublicAdvice = Nothing
End Sub


Public Sub InitDoctorShift(ByVal lngSys As Long, cnMain As ADODB.Connection)
    '------------------------------------------------
    '���ܣ� ����������ָ�����ܣ�����ִ����س���
    '������
    '   lngModul:��Ҫִ�еĹ������
    '   cnMain:����������ݿ�����
    Set gcnOracle = cnMain
    glngSys = lngSys
    gstrProductName = GetSetting("ZLSOFT", "ע����Ϣ", "��Ʒ����", "����")
End Sub

Public Sub ShowViewShift(frmMain As Object, str�������� As String)
    'Ԥ�����Ӱ����ݸ�ʽ
    ' frmMain�����������
    ' str��������:Ĭ�ϲ鿴�������͵��ַ�������ʽ������,һ������,XXX,XXX
    Dim frm As frmShiftEdit
    
    Set frm = New frmShiftEdit
    frm.gstrԤ������ = str��������
    frm.BorderStyle = FormBorderStyleConstants.vbSizable    '����Ϊ�ޱ߿�
    frm.Caption = "���˽��Ӱ�����Ԥ��"      '�ص�����һ��
    
    frm.Show 1, frmMain
End Sub