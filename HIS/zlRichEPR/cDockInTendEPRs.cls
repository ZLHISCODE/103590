VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDockInTendEPRs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private mfrmInTendEPR As frmDockInTendEPR
Attribute mfrmInTendEPR.VB_VarHelpID = -1

Private Sub Class_initialize()
    Dim strPrivs As String
    
    strPrivs = GetPrivFunc(glngSys, 1255)
    Set mfrmInTendEPR = New frmDockInTendEPR
    Call mfrmInTendEPR.InitData(strPrivs)
End Sub

Private Sub Class_Terminate()
    Unload mfrmInTendEPR
    Set mfrmInTendEPR = Nothing
End Sub
Public Sub SetFontSize(ByVal bytSize As Byte)
    Call mfrmInTendEPR.SetFontSize(bytSize)
End Sub
Public Function zlGetForm() As Object
    Set zlGetForm = mfrmInTendEPR
End Function

Public Sub zlDefCommandBars(ByVal CommandBars As Object, Optional ByVal blnInsideTools As Boolean)
'�Խ�������
    Call mfrmInTendEPR.zlDefCommandBars(CommandBars, blnInsideTools)
End Sub

Public Sub zlUpdateCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Call mfrmInTendEPR.zlUpdateCommandBars(Control)
End Sub

Public Sub zlExecuteCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Call mfrmInTendEPR.zlExecuteCommandBars(Control)
End Sub

Public Function zlRefresh(ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal lngDeptId As Long, _
                        ByVal bnEdit As Boolean, ByVal blnForce As Boolean, ByVal blnMoved As Boolean, Optional ByVal blnDoctorStation As Boolean = False) As Long
    '����:����ˢ��ָ�����˵Ĳ������ݣ�����������ṩ�༭����
    '����:  lngPatiId-����id;
    '       lngPageId-��ҳid;
    '       lngDeptId-��ǰ�������ţ�ע�ⲻ�ǲ��˵�ǰ���ڿ��һ��Ժ���ң���Ҫ���ݸò���ȷ�������ӵĲ��˲�����
    '       blnEdit-�Ƿ������༭��ͨ����ǰ�������Ų��ǲ��˵�ǰ���ڿ��һ��Ժ���ң���Ӧ�ò������༭
    '       blnForce-ǰ�����β���ID,��ҳID��ͬʱ,�Ƿ�ǿ��ˢ��
    '       blnMoved-�����Ƿ�ת��
    zlRefresh = mfrmInTendEPR.RefreshData(lngPatiID, lngPageId, lngDeptId, blnDoctorStation, bnEdit, blnForce, blnMoved)
End Function

Public Function GetFormOperation() As String
'��¼����ѡ����Ϣ����Ϊ����վ���л�ҳ��ʱ���ͷ��˶��󣬻�����ʱ���³�ʼ��ˢ�µġ�
    GetFormOperation = mfrmInTendEPR.GetFormOperation
End Function
Public Sub RestoreFormOperation(ByVal strValue As String)
'�ָ�����ѡ����Ϣ������վ��ˢ��֮ǰ����
    Call mfrmInTendEPR.RestoreFormOperation(strValue)
End Sub