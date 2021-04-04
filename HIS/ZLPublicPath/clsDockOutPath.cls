VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDockOutPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Event ViewEPRReport(ByVal ����ID As Long, ByVal CanPrint As Boolean) 'Ҫ��鿴����
Public Event Activate()                                                     '�Ӵ��弤��
Public Event RequestRefresh(ByVal lngPathState As Long)                     'Ҫ��������ˢ��,-1��ʾδ����״̬
Public Event StatusTextUpdate(ByVal Text As String)                         'Ҫ�����������״̬������

Private WithEvents mclsDockOutPath As zlCISPath.clsDockOutPath
Attribute mclsDockOutPath.VB_VarHelpID = -1

Private Sub Class_Initialize()
    Set mclsDockOutPath = New zlCISPath.clsDockOutPath
End Sub

Private Sub Class_Terminate()
    Set mclsDockOutPath = Nothing
End Sub

Private Sub mclsDockOutPath_ViewEPRReport(ByVal ����ID As Long, ByVal CanPrint As Boolean)
    RaiseEvent ViewEPRReport(����ID, CanPrint)
End Sub

Private Sub mclsDockOutPath_Activate()
    RaiseEvent Activate
End Sub

Private Sub mclsDockOutPath_RequestRefresh(ByVal lngPathState As Long)
    RaiseEvent RequestRefresh(lngPathState)
End Sub

Private Sub mclsDockOutPath_StatusTextUpdate(ByVal Text As String)
    RaiseEvent StatusTextUpdate(Text)
End Sub

Public Function zlGetForm() As Object
'���ܣ���ȡ��ҵ�񴰿�,���������ʼ��ʱֻ����һ�λ�ȡ
    Set zlGetForm = mclsDockOutPath.zlGetForm
End Function

Public Sub zlDefCommandBars(ByVal frmParent As Object, ByVal CommandBars As Object)
'���ܣ������Ӵ���Ĳ˵��͹�����(����������Ҫʹ�õĲ˵��͹�����)
'������int����=0-ҽ��վ����,1-��ʿվ����,2-ҽ��վ����(PACS/LIS)
'      CommandBars=�����ڲ鿴ʱ���Բ���(����Nothing)
    Call mclsDockOutPath.zlDefCommandBars(frmParent, CommandBars)
End Sub

Public Sub zlUpdateCommandBars(ByVal Control As Object)
    Call mclsDockOutPath.zlUpdateCommandBars(Control)
End Sub

Public Sub zlExecuteCommandBars(ByVal Control As Object)
    Call mclsDockOutPath.zlExecuteCommandBars(Control)
End Sub

Public Sub zlPopupCommandBars(ByVal CommandBar As Object)
    Call mclsDockOutPath.zlPopupCommandBars(CommandBar)
End Sub

Public Sub zlRefresh(ByVal lng����Id As Long, ByVal lng�Һ�ID As Long, ByVal str�Һ�NO As String, ByVal lng����ID As Long, ByVal int����״̬ As Integer, Optional ByVal blnMoved As Boolean, _
     Optional ByVal blnForceRefresh As Boolean, Optional ByRef objMip As Object)
'���ܣ�ˢ��·��������
'������int����״̬= ���˹Һż�¼�����ִ��״̬��0-�ȴ�����,1-��ɾ���,2-���ھ���,-1���Ϊ������
'      blnForceRefresh=ǿ��ˢ��·����
'      lngҽ������ID=ҽ��վ����ID
'      objMip ��Ϣ����
    Call mclsDockOutPath.zlRefresh(lng����Id, lng�Һ�ID, str�Һ�NO, lng����ID, int����״̬, blnMoved, blnForceRefresh)
End Sub

Public Function zlExecPathSend(Optional ByRef blnIsCancel As Boolean, Optional ByVal lngType As Long) As Boolean
'���ܣ�ִ��·������(ҽ��վ�ͻ�ʿվ����)
'������blnIsCancel��û��·��������ʱ���û��Ƿ�ȡ����������true=ȡ��
'      lngType:1-ҽ���༭������ã��������󲻼������ɣ���Ϊҽ���༭���治���ٵ���ҽ���༭��
    zlExecPathSend = mclsDockOutPath.zlExecPathSend(blnIsCancel, lngType)
End Function

Public Function zlExecEvaluateCancel(Optional ByVal blnPrompt As Boolean = True, Optional ByVal blnRefresh As Boolean = True) As Boolean
'���ܣ�ִ��ȡ������(ҽ��վ)
'������blnPrompt=�Ƿ񵯳�ѯ����ʾ
    zlExecEvaluateCancel = mclsDockOutPath.zlExecEvaluateCancel(blnPrompt, blnRefresh)
End Function

Public Function zlExePathAppendItem(ByVal strItemType As String, ByVal strAdviceIDs As String, ByVal lngִ��ID As Long, ByVal datDate As Date) As Boolean
'���ܣ�����·������Ŀ(ҽ��վ)
    zlExePathAppendItem = mclsDockOutPath.zlExePathAppendItem(strItemType, strAdviceIDs, lngִ��ID, datDate)
End Function

Public Function zlImportPath() As Boolean
    zlImportPath = mclsDockOutPath.zlImportPath
End Function

Public Function zlShowAltAdvice(ByVal frmParent As Object, ByVal lng·����ĿID As Long, Optional ByVal strSelectedIDs As String, Optional ByVal intMode As Integer = 0) As String
'���ܣ�·������ʱ�����ݵ�ǰ�е�·����ĿID����ʾ��ѡҽ��
'������strSelectedIds���Ѿ�ѡ�еı�ѡҽ��IDs
'���أ�ѡ����ѡ�еı�ѡҽ��IDs
    zlShowAltAdvice = mclsDockOutPath.zlShowAltAdvice(frmParent, lng·����ĿID, strSelectedIDs, intMode)
End Function

Public Function zlCheckPathDiagOut(ByVal lng�Һ�ID As Long, ByVal lng�����Դ As Long, ByVal lngDiagType As Long, ByVal lngDiag As Long, ByVal lng���ID As Long) As Boolean
'���ܣ��������·����Ӧ������Ƿ������޸�
'������lng�Һ�ID���Һ�ID,lngDiagType���������,lngDiag=����ID
'����ֵ:F-�������޸�;T-�����޸�
    zlCheckPathDiagOut = mclsDockOutPath.zlCheckPathDiagOut(lng�Һ�ID, lng�����Դ, lngDiagType, lngDiag, lng���ID)
End Function

Public Sub SetFontSize(ByVal bytSize As Byte)
'����:����·�����嵥�������С
'���:bytSize��0-С(ȱʡ)��1-��
    Call mclsDockOutPath.SetFontSize(bytSize)
End Sub

Public Sub zlAddOutPathItem(ByVal strAdviceIDs As String, ByVal lng����Id As Long, ByVal lng�Һ�ID As Long)
'���ܣ�����ҽ��Ϊ"ֹͣ"״̬��ҽ��ʱ�����ݻ���ҽ��ID��������·������Ŀ
'������strAdviceIDs�����˵�ҽ��ID��
'      lng����Id-����ID
'      lng�Һ�Id-�Һ�ID
    Call mclsDockOutPath.zlAddOutPathItem(strAdviceIDs, lng����Id, lng�Һ�ID)
End Sub

Public Sub zlFuncPathTableOutPut(ByVal bytStyle As Byte, Optional ByVal blnIsSetup As Boolean, Optional ByVal strPDFFile As String, Optional ByVal lng����Id As Long, Optional ByVal lng�Һ�ID As Long, Optional ByVal strDeviceName As String)
'���ܣ��ٴ�·��������ӡ
'������bytStyle=1-��ӡ,2-Ԥ��,3-�����Excel,4-�����PDF
'     blnIsSetup-������ӡ�����д�ӡǰ����
'     ��bytStyle=4ʱ����Ҫ����strPDFFile=PDF���Ĭ��·��,�����ļ�������׺
'     ����ⲿδ����zlRefresh��zlRefreshReadOnly�Ļ�������Ҫ���벡��ID,�Һ�ID
'     strDeviceName-�ⲿָ����ӡ��ʱ�����ӡ������
    Call mclsDockOutPath.zlFuncPathTableOutPut(bytStyle, blnIsSetup, strPDFFile, lng����Id, lng�Һ�ID, strDeviceName)
End Sub

