VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDockOutAdvices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'������ص��¼�
Public Event Activate() '�Ӵ��弤��
Public Event RequestRefresh() 'Ҫ��������ˢ��
Public Event StatusTextUpdate(ByVal Text As String) 'Ҫ�����������״̬������
Public Event ViewEPRReport(ByVal ����ID As Long, ByVal CanPrint As Boolean) 'Ҫ��鿴����
Public Event PrintEPRReport(ByVal ����ID As Long, ByVal Preview As Boolean) 'Ҫ���ӡ����
Public Event ViewPACSImage(ByVal ҽ��ID As Long) 'Ҫ����й�Ƭ
Public Event EditDiagnose(ParentForm As Object, ByVal �Һŵ� As String, Succeed As Boolean) '�༭�������
Public Event CheckInfectDisease(ByVal blnOnChek As Boolean, ByVal str����ID As String, ByVal str���Id As String, ByRef blnNo As Boolean) '������ϼ���Ƿ���д��Ⱦ�����濨
Public Event VSKeyPress(KeyAscii As Integer)
Private WithEvents mclsOutAdvice As zlCISKernel.clsDockOutAdvices
Attribute mclsOutAdvice.VB_VarHelpID = -1

Private Sub Class_Initialize()
    Set mclsOutAdvice = New zlCISKernel.clsDockOutAdvices
End Sub

Private Sub Class_Terminate()
    Set mclsOutAdvice = Nothing
End Sub

Private Sub mclsOutAdvice_VSKeyPress(KeyAscii As Integer)
    RaiseEvent VSKeyPress(KeyAscii)
End Sub

Private Sub mclsOutAdvice_Activate()
    RaiseEvent Activate
End Sub

Private Sub mclsOutAdvice_CheckInfectDisease(ByVal blnOnChek As Boolean, ByVal str����ID As String, ByVal str���Id As String, ByRef blnNo As Boolean)
    RaiseEvent CheckInfectDisease(blnOnChek, str����ID, str���Id, blnNo)
End Sub

Private Sub mclsOutAdvice_EditDiagnose(ParentForm As Object, ByVal �Һŵ� As String, Succeed As Boolean)
    RaiseEvent EditDiagnose(ParentForm, �Һŵ�, Succeed)
End Sub

Private Sub mclsOutAdvice_RequestRefresh()
    RaiseEvent RequestRefresh
End Sub

Private Sub mclsOutAdvice_StatusTextUpdate(ByVal Text As String)
    RaiseEvent StatusTextUpdate(Text)
End Sub

Private Sub mclsOutAdvice_ViewEPRReport(ByVal ����ID As Long, ByVal CanPrint As Boolean)
    RaiseEvent ViewEPRReport(����ID, CanPrint)
End Sub

Private Sub mclsOutAdvice_PrintEPRReport(ByVal ����ID As Long, ByVal Preview As Boolean)
    RaiseEvent PrintEPRReport(����ID, Preview)
End Sub

Private Sub mclsOutAdvice_ViewPACSImage(ByVal ҽ��ID As Long)
    RaiseEvent ViewPACSImage(ҽ��ID)
End Sub

Public Function zlGetForm() As Object
'���ܣ���ȡ��ҵ�񴰿�,���������ʼ��ʱֻ����һ�λ�ȡ
    '��ʼ��ҽ������:�ɷ�ɢ�������GetForm��ȥ��ʼ��
    Set zlGetForm = mclsOutAdvice.zlGetForm
End Function

Public Sub zlDefCommandBars(ByVal frmParent As Object, ByVal CommandBars As Object, ByVal int���� As Integer, _
    Optional ByRef objPlugIn As Object, Optional ByRef objSquareCard As Object, Optional ByVal blnModalNew As Boolean)
'���ܣ������Ӵ���Ĳ˵��͹�����(����������Ҫʹ�õĲ˵��͹�����)
'������int����=0-ҽ��վ����,1-��ʿվ����,2-ҽ��վ����(PACS/LIS)
'      CommandBars=�����ڲ鿴ʱ���Բ���(����Nothing)
'      objPlugIn=ͨ������ҽ������վ�������ҹ��ܶ���
'      objSquareCard=ͨ������ҽ��վ��ҽ������վ�����һ��ͨ���㲿��
    Call mclsOutAdvice.zlDefCommandBars(frmParent, CommandBars, int����, objPlugIn, objSquareCard, blnModalNew)
End Sub

Public Sub zlUpdateCommandBars(ByVal Control As Object)
    Call mclsOutAdvice.zlUpdateCommandBars(Control)
End Sub

Public Sub zlExecuteCommandBars(ByVal Control As Object)
    Call mclsOutAdvice.zlExecuteCommandBars(Control)
End Sub

Public Sub zlPopupCommandBars(ByVal CommandBar As Object)
    Call mclsOutAdvice.zlPopupCommandBars(CommandBar)
End Sub

Public Sub zlInitPath(ByRef clsDockOutPath As Object)
'���ܣ���ʼ���ٴ�·�����󣬴���ҽ��վ��ʿվ���õĶ���
    Call mclsOutAdvice.zlInitPath(clsDockOutPath)
End Sub

Public Sub zlRefresh(ByVal lng����ID As Long, ByVal str�Һŵ� As String, ByVal blnEditable As Boolean, _
    Optional ByVal blnMoved As Boolean, Optional ByVal lngǰ��ID As Long, Optional ByVal lng�������ID As Long, _
    Optional ByRef objMip As Object, Optional ByVal lngǰ�����ID As Long, Optional ByVal lng·��״̬ As Long = -1, Optional ByVal int�������� As Integer)
'���ܣ�ˢ��סԺҽ������
'������lngǰ��ID=����ҽ��վ����ʱ����
'      blnMoved=�ò��˵������Ƿ���ת��
'      blnEditable=�ɷ�Բ���ҽ�����б༭
'      lng�������ID=�����ҽ��վ���ã���Ϊҽ������ID
'      strǰ��IDsҽ��վ���ڵ�ǰ����ִ�е�����ҽ��
'      objMip ��Ϣ���� zl9ComLib.clsMipModule
'      lng·��״̬=-1:δ����,0-�����ϵ���������1-ִ���У�2-����������3-�������
'      int��������= pt���� = 0��pt���� = 1��pt���� = 2��ptת�� = 3��ptԤԼ = 4��pt���� = 5
    Call mclsOutAdvice.zlRefresh(lng����ID, str�Һŵ�, blnEditable, blnMoved, lngǰ��ID, lng�������ID, objMip, lngǰ�����ID, lng·��״̬, int��������)
End Sub

Public Sub SetFontSize(ByVal bytSize As Byte)
'����:����ҽ���嵥�������С
'���:bytSize��0-С(ȱʡ)��1-��
    Call mclsOutAdvice.SetFontSize(bytSize)
End Sub