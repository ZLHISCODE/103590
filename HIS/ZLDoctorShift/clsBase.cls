VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

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
    If Err.Number <> 0 Then
        Set gobjEmr = Nothing
        Err.Clear
    End If
    
    Select Case lngModul
        Case 1242
            frmDoctorShift.Show , frmMain
    End Select
End Sub