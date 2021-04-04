VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDockEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Event AfterSaved(lngRecordId As Long)                        'EMR�ĵ�������¼�

Private WithEvents mobjDoc As cEPRDocument
Attribute mobjDoc.VB_VarHelpID = -1
Public Function GetOutEPREditForm(ByVal frmParent As Object, ByVal lngFileID As Long, ByVal lngPatiID As Long, ByVal lngBillId As Long, ByVal lngDeptId As Long) As Long
'���ݴ�����ļ�ID,����ID���Һ�ID������ID����ѯ��ǰ�����Ƿ񴴽��˲����ļ������δ�������򷵻��¿��˲����Ĵ������
'����Ѵ������򷵻��Ѵ����Ĳ����ļ��༭����Ķ��󡣷��صĴ���Ϊ�ޱ߿�ģʽ����������ģ�飩
Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand

    If mobjDoc Is Nothing Then
        Set mobjDoc = New cEPRDocument
    End If
    
    gstrSQL = "Select ID From ���Ӳ�����¼ Where ����id = [1] And ��ҳid = [2] And �ļ�id + 0 = [3] And ����id = [4]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "���ָ��������д���", lngPatiID, lngBillId, lngFileID, lngDeptId)
    If rsTemp.EOF Then
        mobjDoc.InitEPRDoc cprEM_����, cprET_�������༭, lngFileID, cprPF_����, lngPatiID, lngBillId, , lngDeptId
    Else
        mobjDoc.InitEPRDoc cprEM_�޸�, cprET_�������༭, CLng(rsTemp!ID), cprPF_����, lngPatiID, lngBillId, , lngDeptId
    End If
    mobjDoc.ShowEPREditor frmParent, True, -1
    
    GetOutEPREditForm = mobjDoc.frmEditor.hWnd
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Public Function DocModified() As Boolean
    If mobjDoc Is Nothing Then
        DocModified = False
    Else
        DocModified = mobjDoc.frmEditor.Editor1.Modified
    End If
End Function
Public Sub SaveDoc()
    If mobjDoc Is Nothing Then
        Exit Sub
    Else
        Call mobjDoc.frmEditor.SaveEMRDoc
    End If
End Sub
Private Sub DocObjectClear()
    On Error Resume Next
    Set mobjDoc.Tables = Nothing
    Set mobjDoc.Pictures = Nothing
    Set mobjDoc.Compends = Nothing
    Set mobjDoc.Elements = Nothing
    Set mobjDoc.Diagnosises = Nothing
    Set mobjDoc.Signs = Nothing

    Set mobjDoc.EPRFileInfo = Nothing
    Set mobjDoc.EPRDemoInfo = Nothing
    Set mobjDoc.EPRPatiRecInfo = Nothing

    Unload mobjDoc.frmEditor
    Set mobjDoc.frmEditor = Nothing
    
    Set mobjDoc = Nothing
    Err.Clear
End Sub
Private Sub Class_Terminate()
    Call DocObjectClear
End Sub

Private Sub mobjDoc_AfterSaved(lngRecordId As Long)
    RaiseEvent AfterSaved(lngRecordId)
End Sub