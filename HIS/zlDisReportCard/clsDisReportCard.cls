VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDisReportCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mfrmReport As Object
Private mobjReportEx As Object


Public Sub Init(ByVal cnMain As ADODB.Connection, ByVal lngSys As Long)
'���ܣ�������ʼ��

    On Error GoTo errHand
    Set gcnOracle = cnMain
    'Call InitCommon(gcnOracle)
    gstrSysName = GetSetting(AppName:="ZLSOFT", Section:="ע����Ϣ", Key:=UCase("gstrSysName"), Default:="")
    Call GetUserInfo
    glngSys = lngSys
    
    '��ȡzlDisReportCardEx����
    Call CreateCardExOK(1277)
    Exit Sub
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Err = 0
End Sub

Public Sub OpenDoc(ByVal frmParent As Object, ByVal bytType As Byte, ByVal lngPatiID As Long, ByVal lngPageID As Long, ByVal bytFrom As Byte, ByVal bytBabyNo As Byte, ByVal lngDeptID As Long, ByVal lngFileId As Long, Optional ByVal blnHand As Boolean = False)
'���ܣ����������ĵ�ʱ���ã�ʵ�����ݳ�ʼ���༭���ڴ�
'����˵��:
'frmParent:�༭����������ʾ�ĸ�����
'bytType:�༭��ʽ0-������1-�޸ģ�����������ȡ����
'lngPatiID:����ID
'lngPageID:��ҳID
'bytFrom:������Դ
'lngDeptID:��ǰ����
'lngFileId:bytType=1ʱ�������޸ĵ��ļ�ID,��Դ�ڵ��Ӳ�����¼.ID
'blnHand:�Ƿ�Ϊ�ֶ��򿪴�Ⱦ�����濨
    On Error GoTo errHand
    Call frmMain.ShowMe(frmParent, bytType, lngPatiID, lngPageID, bytFrom, bytBabyNo, lngDeptID, lngFileId, blnHand)
    Exit Sub
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Err = 0
End Sub


Public Sub PrintDoc(ByVal frmParent As Object, ByVal lngPatiID As Long, ByVal lngPageID As Long, ByVal lngFileId As Long, ByVal strPrintDeviceName As String)
'����:��ӡ
'����˵��:
'frmParent:�༭����������ʾ�ĸ�����
'lngPatiID:����ID
'lngPageID:��ҳID
'lngFileId:�ļ�ID
'strPrintDeviceName:����������ӡʱ��ָ����ӡ������ָ��ʱ���������ṩ��ӡ��ѡ��
    Call mfrmReport.mclsReport.PrintReport(frmParent, lngPatiID, lngPageID, lngFileId, strPrintDeviceName)
End Sub

Public Sub zlRefresh(ByVal lngPatiID As Long, ByVal lngPageID As Long, ByVal lngFileId As Long, ByVal blnMoved As Boolean)
'����:ˢ��
'����˵��:
'lngPatiID:����ID
'lngPageID:��ҳID
'lngFileId:�ļ�ID
'blnMoved:�����Ƿ�ת��
    On Error GoTo errHand
    Call mfrmReport.mclsReport.zlRefresh(lngPatiID, lngPageID, lngFileId, blnMoved)
    Exit Sub
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Err = 0
End Sub

Public Function zlGetForm() As Object
    If mfrmReport Is Nothing Then Set mfrmReport = GetReport
    If mfrmReport.hWnd = 0 Then Set mfrmReport = GetReport '�ⲿUnLoad
    Set zlGetForm = mfrmReport
End Function

Public Sub SetCaption����֤()
    Call mfrmReport.mclsReport.SetCaption����֤
End Sub

Private Sub Class_Initialize()
    On Error Resume Next
    '��ȡzlDisReportCardEx����
    Call CreateCardExOK(1277)
    
    Set mfrmReport = GetReport
    mfrmReport.mclsReport.blnHaveStatus = False
    Err.Clear
End Sub

Private Sub Class_Terminate()
'    Set gobjCardEx = Nothing
    If Not mfrmReport Is Nothing Then
        Unload mfrmReport
        Set mfrmReport = Nothing
    End If
End Sub
