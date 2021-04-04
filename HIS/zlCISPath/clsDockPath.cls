VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDockPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Event ViewEPRReport(ByVal ����ID As Long, ByVal CanPrint As Boolean) 'Ҫ��鿴����
Public Event Activate() '�Ӵ��弤��
Public Event RequestRefresh(ByVal lngPathState As Long) 'Ҫ��������ˢ��,-1��ʾδ����״̬
Public Event StatusTextUpdate(ByVal Text As String) 'Ҫ�����������״̬������

Private WithEvents mfrmPathTable As frmPathTable
Attribute mfrmPathTable.VB_VarHelpID = -1
Private mlngHwnd As Long

Private Sub Class_Initialize()
    Set mfrmPathTable = New frmPathTable
End Sub

Private Sub Class_Terminate()
    Unload mfrmPathTable
    Set mfrmPathTable = Nothing
    Set gobjPlugIn = Nothing
End Sub

Private Sub mfrmPathTable_ViewEPRReport(ByVal ����ID As Long, ByVal CanPrint As Boolean)
    RaiseEvent ViewEPRReport(����ID, CanPrint)
End Sub

Private Sub mfrmPathTable_Activate()
    RaiseEvent Activate
End Sub

Private Sub mfrmPathTable_RequestRefresh(ByVal lngPathState As Long)
    If glngHwnd = 0 Then RaiseEvent RequestRefresh(lngPathState)
End Sub

Private Sub mfrmPathTable_StatusTextUpdate(ByVal Text As String)
    RaiseEvent StatusTextUpdate(Text)
End Sub

Public Function zlGetForm() As Object
'���ܣ���ȡ��ҵ�񴰿�,���������ʼ��ʱֻ����һ�λ�ȡ
    Set zlGetForm = mfrmPathTable
End Function

Public Sub zlDefCommandBars(ByVal frmParent As Object, ByVal CommandBars As Object, ByVal int���� As Integer, Optional ByVal blnInsideTools As Boolean = False)
'���ܣ������Ӵ���Ĳ˵��͹�����(����������Ҫʹ�õĲ˵��͹�����)
'������int����=0-ҽ��վ����,1-��ʿվ����,2-ҽ��վ����(PACS/LIS)
'      CommandBars=�����ڲ鿴ʱ���Բ���(����Nothing)
    Call mfrmPathTable.zlDefCommandBars(frmParent, CommandBars, int����, blnInsideTools)
End Sub

Public Sub zlUpdateCommandBars(ByVal Control As Object)
    Call mfrmPathTable.zlUpdateCommandBars(Control)
End Sub

Public Sub zlExecuteCommandBars(ByVal Control As Object)
    Call mfrmPathTable.zlExecuteCommandBars(Control)
End Sub

Public Sub zlPopupCommandBars(ByVal CommandBar As Object)
    Call mfrmPathTable.zlPopupCommandBars(CommandBar)
End Sub

Public Sub zlRefresh(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal lng����ID As Long, ByVal lng����ID As Long, ByVal int����״̬ As Integer, Optional ByVal blnMoved As Boolean, _
     Optional ByVal blnForceRefresh As Boolean, Optional ByVal lngState As Long, Optional ByVal lngҽ������ID As Long, Optional ByRef objMip As Object)
'���ܣ�ˢ��·��������
'������int��������=������ҳ�е�״̬��0-����סԺ��1-��δ��ƣ�2-����ת�ƣ�3-��Ԥ��Ժ
'      blnMoved=�ò��˵������Ƿ���ת��
'      blnForceRefresh=ǿ��ˢ��·����
'      lngState =���˵�״̬
'      lngҽ������ID=ҽ��վ����ID
'      objMip ��Ϣ����
    Call mfrmPathTable.zlRefresh(lng����ID, lng��ҳID, lng����ID, lng����ID, int����״̬, blnMoved, blnForceRefresh, lngState, , lngҽ������ID, objMip)
End Sub

Public Sub zlRefreshReadOnly(ByVal lng����ID As Long, ByVal lng��ҳID As Long)
'���ܣ��ⲿ�������,��ʾֻ����·������
'ע�⣺�������ҽ��վ��ʿվ���ã�֮ǰ�����ȵ�clsCISPath.InitCISPath
'���ó��� ���Ӳ�������;1566-�������Ĵ�ӡ
    Dim lng����ID As Long, lng����ID As Long
    Dim rsTmp As ADODB.Recordset, strSql As String
    
    strSql = "Select ��ǰ����id,��Ժ����id,״̬,����ת�� From ������ҳ Where ����ID = [1] And ��ҳID = [2]"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "zlRefreshReadOnly", lng����ID, lng��ҳID)
    If rsTmp.RecordCount = 0 Then Exit Sub
    Call mfrmPathTable.zlRefresh(lng����ID, lng��ҳID, Val("" & rsTmp!��ǰ����id), Val("" & rsTmp!��Ժ����id), _
            Val("" & rsTmp!״̬), Val("" & rsTmp!����ת��) = 1, False, , , , , mlngHwnd <> 0)

    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Function zlExecPathSend(Optional ByRef blnIsCancel As Boolean, Optional ByVal lngType As Long) As Boolean
'���ܣ�ִ��·������(ҽ��վ�ͻ�ʿվ����)
'������blnIsCancel��û��·��������ʱ���û��Ƿ�ȡ����������true=ȡ��
'      lngType:1-ҽ���༭������ã��������󲻼������ɣ���Ϊҽ���༭���治���ٵ���ҽ���༭��
    zlExecPathSend = mfrmPathTable.FuncSendItem(blnIsCancel, lngType)
End Function

Public Function zlExecEvaluateCancel(Optional ByVal blnPrompt As Boolean = True, Optional ByVal blnRefresh As Boolean = True) As Boolean
'���ܣ�ִ��ȡ������(ҽ��վ�ͻ�ʿվ����)
'������blnPrompt=�Ƿ񵯳�ѯ����ʾ
    zlExecEvaluateCancel = mfrmPathTable.FuncEvaluateCancel(blnPrompt, blnRefresh)
End Function

Public Function zlExePathAppendItem(ByVal strItemType As String, ByVal strAdviceIDs As String, ByVal lngִ��ID As Long, ByVal datDate As Date) As Boolean
'���ܣ�����·������Ŀ(ҽ��վ�ͻ�ʿվ����)
    zlExePathAppendItem = mfrmPathTable.FuncAppendItem(1, strItemType, strAdviceIDs, lngִ��ID, datDate)
End Function

Public Function zlImportPath(Optional ByVal lngHwnd As Long) As Boolean
'������lngHwnd=�°没�����븸��������Ĭ��Ϊ0
    glngHwnd = lngHwnd
    mlngHwnd = lngHwnd
    zlImportPath = mfrmPathTable.FuncImport(lngHwnd)
End Function

Public Function zlShowAltAdvice(ByVal frmParent As Object, ByVal lng·����ĿID As Long, Optional ByVal strSelectedIDs As String, Optional ByVal intӤ�� As Integer) As String
'���ܣ�·������ʱ�����ݵ�ǰ�е�·����ĿID����ʾ��ѡҽ��
'������strSelectedIds���Ѿ�ѡ�еı�ѡҽ��IDs
'���أ�ѡ����ѡ�еı�ѡҽ��IDs
    zlShowAltAdvice = frmAltAdviceSelect.ShowSelect(frmParent, lng·����ĿID, strSelectedIDs, intӤ��)
End Function

Public Sub SetFontSize(ByVal bytSize As Byte)
'����:����·�����嵥�������С
'���:bytSize��0-С(ȱʡ)��1-��
    Call mfrmPathTable.SetFontSize(bytSize)
End Sub

Public Sub zlAddOutPathItem(ByVal strAdviceIDs As String, ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal bytType As Byte, ByRef colSQL As Collection)
'���ܣ�����ҽ��Ϊ"ֹͣ"״̬��ҽ��ʱ�����ݻ���ҽ��ID��������·������Ŀ
'      ��������ҽ��ʱ,���ݻ���ҽ��ID��������Ӧ��·����Ŀ
'������strAdviceIDs�����˵�ҽ��ID��
'      lng����Id-����ID
'      lng��ҳId-��ҳID
'      bytType=8 ֹͣ;=4=����
'����:
'     varSQL=��ִ�е�SQL���,�����ͬһ�����ύ
    Call AddOutPathItem(strAdviceIDs, 2, lng����ID, lng��ҳID, bytType, colSQL)
End Sub

Public Sub zlFuncPathTableOutPut(ByVal bytStyle As Byte, Optional ByVal blnIsSetup As Boolean, Optional ByVal strPDFFile As String, Optional ByVal lng����ID As Long, Optional ByVal lng��ҳID As Long, Optional ByVal strDeviceName As String)
'���ܣ��ٴ�·��������ӡ
'������bytStyle=1-��ӡ,2-Ԥ��,3-�����Excel,4-�����PDF
'     blnIsSetup-������ӡ�����д�ӡǰ����
'     ��bytStyle=4ʱ����Ҫ����strPDFFile=PDF���Ĭ��·��,�����ļ�������׺
'     ����ⲿδ����zlRefresh��zlRefreshReadOnly�Ļ�������Ҫ���벡��ID,��ҳID
'     strDeviceName-�ⲿָ����ӡ��ʱ�����ӡ������
    If lng����ID <> 0 And lng��ҳID <> 0 Then
        Call zlRefreshReadOnly(lng����ID, lng��ҳID)
    End If
    Call mfrmPathTable.zlPrintOutPut(bytStyle, blnIsSetup, strPDFFile, strDeviceName)
End Sub

Public Function GetFormOperation() As String
    GetFormOperation = mfrmPathTable.GetFormOperation
End Function

Public Function RestoreFormOperation(ByVal strValue As String)
    Call mfrmPathTable.RestoreFormOperation(strValue)
End Function