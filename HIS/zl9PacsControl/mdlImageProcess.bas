Attribute VB_Name = "mdlImageProcess"
Option Explicit

Public Enum TImageType
    mtTagImage = 0      '���ͼ
    mtReportImage = 1   '����ͼ
    mtStadyImage = 2    '���ͼ
End Enum

Public gobjImageProcess As frmImageProcess

Public glngColor(10) As Long             '���ͼ��Բ�α��ʹ�õ�9����ɫ

Public Const G_STR_TAG = "Po=Ϣ��[+]E=������[+]M=��Ƕ[+]L=ճĤ�װ�[+]C=ʪ��[+]I=�����԰�[+]W=�����ɫ��Ƥ[+]AT=�쳣ת����[+]V=�ǵ���Ѫ��[+]P=��״Ѫ��[+]Xn=ֱ�ӻ�첿λ"

'ͼ����
Public Const conMenu_Process_Window = 501           '���ȶԱȶ�
Public Const conMenu_Process_Zoom = 502             '����
Public Const conMenu_Process_Corp = 512             '�϶�
Public Const conMenu_Process_RRotate = 503          '˳ʱ����ת
Public Const conMenu_Process_LRotate = 504          '��ʱ����ת
Public Const conMenu_Process_Sharpness = 505        '��
Public Const conMenu_Process_Filter = 506           'ƽ��
Public Const conMenu_Process_Arrow = 507            '��ͷ��ע
Public Const conMenu_Process_Ellipse = 508          'Բ�α�ע
Public Const conMenu_Process_Text = 509             '���ֱ�ע
Public Const conMenu_Process_RectZoom = 510         '�ü��ɼ�
Public Const conMenu_Process_RectCapture = 511      '�ü���ɼ�
Public Const conMenu_Process_Line = 520             'ֱ�߱�ע
Public Const conMenu_Process_Exit = 2613            '�˳�
Public Const conMenu_Process_Save = 3091            '����
Public Const conMenu_Process_SaveToReport = 3941    '���浽���
Public Const conMenu_Process_SaveToStady = 3943     '���浽����
Public Const conMenu_Process_DelAllLabels = 8113    'ɾ��ȫ����ע��ʹ������ϵͳ��ͼ����
Public Const conMenu_Process_MoveLabel = 6891       '�ƶ���ɾ��ѡ�б�ע��ʹ������ϵͳ��ͼ����
Public Const conMenu_Process_LabelSetUp = 10003     '��ע��ť���ã�ʹ������ϵͳ��ͼ����
Public Const conMenu_Process_Restore = 8124         '�ָ�
Public Const conMenu_Process_TextTag = 5010         '�ı����
Public Const conMenu_Process_NumTag = 7405          '���ֱ��
Public Const conMenu_Process_Page = 1001
Public Const conMenu_Process_Num = 96
Public Const conMenu_Process_Word = 97
