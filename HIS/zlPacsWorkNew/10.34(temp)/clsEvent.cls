VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


'����ģ��ص��¼�
Public Enum TWorkEventType
    wetDelAllImg = 0        'ɾ������ͼ��
    wetGetImg = 1           '��ȡͼ��
    wetTechDo = 2           '��ʦִ��
    wetChangeImgType = 3    '����ͼ������
    
    wetLockStudy = 4        '�������
    wetUnLockStudy = 5      '�������
    wetCaptureFirstImg = 6  '�ɼ���һ��ͼ��
    wetUpdateImg = 7        '����ͼ��
    wetChangeUser = 8       '�����û�
    
    wetPatholRequest = 9    '��������
    wetPatholQuality = 10   '��������
    wetPatholBatSlices = 11 '��Ƭ��������
    wetPatholBatSpeExm = 12 '�ؼ���������
    
    wetSpecimenAccept = 13  '�걾����
    wetSpecimenReject = 14  '�걾����
    wetSpecimenSave = 15    '�걾����
    
    wetMaterialSure = 16    'ȡ��ȷ��
    wetMaterialSave = 17    '�Ŀ鱣��
    
    wetSlicesSure = 18      '��Ƭȷ��
    
    wetSpeExamSure = 19     '�ؼ�ȷ��
    
    wetViewEprReport = 20   'Ԥ�����Ӳ�������
    wetViewPacsImage = 21   'Ԥ��Pacsͼ��
    wetRejectReport = 22    '���ر���
End Enum



Public Event OnWork(objEvent As Object, ByVal lngWorkType As TWorkEventType, ByVal lngAdviceID As Long, ByVal other As Variant)



Public Sub DoWork(objEvent As Object, ByVal lngWorkType As TWorkEventType, ByVal lngAdviceID As Long, Optional other As Variant = "")
On Error GoTo ErrHandle
    RaiseEvent OnWork(objEvent, lngWorkType, lngAdviceID, other)
    
ErrHandle:
End Sub