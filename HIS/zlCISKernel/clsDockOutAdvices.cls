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
Private WithEvents mfrmOutAdvice As frmDockOutAdvice
Attribute mfrmOutAdvice.VB_VarHelpID = -1
Private mblnPlugInFromOut As Boolean '��ҹ����Ƿ����ⲿ����

Private Sub Class_Initialize()
    Set mfrmOutAdvice = New frmDockOutAdvice
End Sub

Private Sub Class_Terminate()
    Unload mfrmOutAdvice
    Set mfrmOutAdvice = Nothing
    'Set gobjPathOut = Nothing
    If Not mblnPlugInFromOut Then Set gobjPlugIn = Nothing
    Set grsҽ�Ƹ��ʽ = Nothing
End Sub

Private Sub mfrmOutAdvice_VSKeyPress(KeyAscii As Integer)
    RaiseEvent VSKeyPress(KeyAscii)
End Sub

Private Sub mfrmOutAdvice_Activate()
    RaiseEvent Activate
End Sub

Private Sub mfrmOutAdvice_CheckInfectDisease(ByVal blnOnChek As Boolean, ByVal str����ID As String, ByVal str���Id As String, ByRef blnNo As Boolean)
    RaiseEvent CheckInfectDisease(blnOnChek, str����ID, str���Id, blnNo)
End Sub

Private Sub mfrmOutAdvice_EditDiagnose(ParentForm As Object, ByVal �Һŵ� As String, Succeed As Boolean)
    RaiseEvent EditDiagnose(ParentForm, �Һŵ�, Succeed)
End Sub

Private Sub mfrmOutAdvice_RequestRefresh()
    RaiseEvent RequestRefresh
End Sub

Private Sub mfrmOutAdvice_StatusTextUpdate(ByVal Text As String)
    RaiseEvent StatusTextUpdate(Text)
End Sub

Private Sub mfrmOutAdvice_ViewEPRReport(ByVal ����ID As Long, ByVal CanPrint As Boolean)
    RaiseEvent ViewEPRReport(����ID, CanPrint)
End Sub

Private Sub mfrmOutAdvice_PrintEPRReport(ByVal ����ID As Long, ByVal Preview As Boolean)
    RaiseEvent PrintEPRReport(����ID, Preview)
End Sub

Private Sub mfrmOutAdvice_ViewPACSImage(ByVal ҽ��ID As Long)
    RaiseEvent ViewPACSImage(ҽ��ID)
End Sub

Public Function zlGetForm() As Object
'���ܣ���ȡ��ҵ�񴰿�,���������ʼ��ʱֻ����һ�λ�ȡ
    '��ʼ��ҽ������:�ɷ�ɢ�������GetForm��ȥ��ʼ��
    Call gclsInsure.InitOracle(gcnOracle)

    '��ʼ��ȡȨ��,�Լӿ�����ٶ�(���Ǳ�Ҫ��)
    Call GetInsidePrivs(p����ҽ���´�, True)
    
    'һЩ�ӿڳ�ʼ��,ֻ��ʼ��һ��
    '���漰����ģ���Ȩ��,��������ʼ��
    
    '����ǩ����ʼ��:���漰����ģ���Ȩ��,��������ʼ��
    '��Ϊ����ģ�����ͬʱʹ��,�ҹ��ñ���gobjESign����˻���Ҫ�ھ���ģ�����ж�Ȩ��
    If gintCA <> 0 Then
        'If InStr(GetInsidePrivs(p����ҽ���´�), "ҽ������ǩ��") > 0 And gobjESign Is Nothing Then
        If gobjESign Is Nothing Then
            On Error Resume Next
            Set gobjESign = CreateObject("zl9ESign.clsESign")
            err.Clear: On Error GoTo 0
            If Not gobjESign Is Nothing Then
                Call gobjESign.Initialize(gcnOracle, glngSys)
            End If
        End If
    Else
        Set gobjESign = Nothing
    End If
    Set zlGetForm = mfrmOutAdvice
End Function

Public Sub zlDefCommandBars(ByVal frmParent As Object, ByVal CommandBars As Object, ByVal int���� As Integer, _
    Optional ByRef objPlugIn As Object, Optional ByRef objSquareCard As Object, Optional ByVal blnModalNew As Boolean)
'���ܣ������Ӵ���Ĳ˵��͹�����(����������Ҫʹ�õĲ˵��͹�����)
'������int����=0-ҽ��վ����,1-��ʿվ����,2-ҽ��վ����(PACS/LIS)
'      CommandBars=�����ڲ鿴ʱ���Բ���(����Nothing)
'      objPlugIn=ͨ������ҽ������վ�������ҹ��ܶ���
'      objSquareCard=ͨ������ҽ��վ��ҽ������վ�����һ��ͨ���㲿��
    Call mfrmOutAdvice.zlDefCommandBars(frmParent, CommandBars, int����, objPlugIn, objSquareCard, blnModalNew)
    If Not objPlugIn Is Nothing Then
        mblnPlugInFromOut = True
    Else
        mblnPlugInFromOut = False
    End If
End Sub

Public Sub zlUpdateCommandBars(ByVal Control As Object)
    Call mfrmOutAdvice.zlUpdateCommandBars(Control)
End Sub

Public Sub zlExecuteCommandBars(ByVal Control As Object)
    Call mfrmOutAdvice.zlExecuteCommandBars(Control)
End Sub

Public Sub zlPopupCommandBars(ByVal CommandBar As Object)
    Call mfrmOutAdvice.zlPopupCommandBars(CommandBar)
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
    Call mfrmOutAdvice.zlRefresh(lng����ID, str�Һŵ�, blnEditable, blnMoved, lngǰ��ID, lng�������ID, objMip, lngǰ�����ID, lng·��״̬, int��������)
End Sub

Public Sub SetFontSize(ByVal bytSize As Byte)
'����:����ҽ���嵥�������С
'���:bytSize��0-С(ȱʡ)��1-��
    Call mfrmOutAdvice.SetFontSize(bytSize)
End Sub

Public Sub zlInitPath(ByRef clsDockOutPath As Object)
'���ܣ���ʼ���ٴ�·�����󣬴���ҽ��վ���õĶ���
    Set gobjPathOut = clsDockOutPath
End Sub