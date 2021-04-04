VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPacsApplication"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mfrmApplicationPage As New frmPacsApplication

Public Sub InitComponents(ByVal lngDeptID As Long, objOwner As Object)
'��ʼ�����뵥
'lngDeptId:��ǰ����Id
'objOwner:�������ڴ��ڶ���

    Call mfrmApplicationPage.InitComponents(lngDeptID, objOwner)
End Sub



Public Function ShowApplicationForm(ByVal lngPatientID As Long, _
                                    ByVal lngCallFrom As Long, _
                                    ByVal lngPatientRegId As Long, _
                                    ByVal lngPatientPageId As Long, _
                                    ByVal lngUpdateAppNoOrAdvId As Long, _
                                    ByRef objAppPages() As clsApplicationData, _
                                    Optional ByVal intBabyID As Integer = 0, _
                                    Optional ByVal blnEdit As Boolean = True, _
                                    Optional ByVal lngProjectId As Long = 0) As Boolean
'��ʾ���뵥����
'lngPatientId:      ����Id
'lngCallFrom:    �ӿڵ�����Դ
'lngPatientRegId:   �Һ�Id
'lngPatientPageId:  ��ҳId
'lngUpdateAdviceId: ҽ��Id�����Ϊ�գ�˵��������������룬�����ʾ�޸ļ�����룬������ҽ��ʱ������������Ϊ�գ��򴫵ݶ�Ӧҽ����ҽ��id��Ϊ�˲���
'objAppPages():     ����ҽ����¼�����ݣ�ÿ������Ԫ�ر�ʾһ���������
'blnEdit: Ϊtrueʱ���Խ��б༭����֮���ܽ��б༭
'lngProjectId: ������ĿID     �¿��������ʱ����Ϊ0����ݴ���Ŀ�Զ���λ

    ShowApplicationForm = mfrmApplicationPage.ShowApplicationForm(lngPatientID, lngCallFrom, lngPatientRegId, _
                                                lngPatientPageId, lngUpdateAppNoOrAdvId, objAppPages(), intBabyID, blnEdit, lngProjectId)

End Function

Private Sub Class_Terminate()
    Set mfrmApplicationPage = Nothing
End Sub