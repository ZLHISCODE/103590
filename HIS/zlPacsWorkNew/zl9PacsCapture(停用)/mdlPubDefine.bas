Attribute VB_Name = "mdlPubDefine"
Option Explicit


'�鿴�˵�
Public Const conMenu_View_ToolBar = 701              '������(&T)
Public Const conMenu_View_ToolBar_Button = 7011         '��׼��ť(&S)
Public Const conMenu_View_ToolBar_Text = 7012           '�ı���ǩ(&T)
Public Const conMenu_View_ToolBar_Size = 7013           '��ͼ��(&B)
Public Const conMenu_View_StatusBar = 702            '״̬��(&S)
Public Const conMenu_View_Append = 703               '������Ϣ(&A)
Public Const conMenu_View_Difference = 704              '��ʾ����(&D)
Public Const conMenu_View_Contrast = 705                 '�ԱȲ鿴


'�ɼ��˵�
Public Const conMenu_Cap_Group = 8090           '�ɼ�
Public Const conMenu_Cap_Dynamic = 8100         '��̬��ʾ(&V)
Public Const conMenu_Cap_MarkMap = 8101         'Ӱ��ɼ�(&C)
Public Const conMenu_Cap_Import = 8102          'Ӱ����(&I)
Public Const conMenu_Cap_DevSet = 8103          'Ӱ���豸����(&D)
Public Const comMenu_Cap_Process = 8104         'Ӱ����
Public Const conMenu_Cap_Record = 8105          '¼��(&R)
Public Const conMenu_Cap_DelImg = 8097          'ɾ��ͼ��
Public Const conMenu_Cap_Full_Screen = 8098     'ȫ��(&U)
Public Const conMenu_Cap_Record_Stop = 8099     'ֹͣ¼��(&O)
Public Const conMenu_Cap_Play = 8106            '����(&P)
Public Const conMenu_Cap_Stop = 8107            'ֹͣ(&T)
Public Const conMenu_Cap_Forward = 8108         '���(&F)
Public Const conMenu_Cap_Back = 8109            '����(&B)
Public Const conMenu_Cap_SaveAs = 8126    '8110          '����¼��(&S)
Public Const conMenu_Cap_OpenStudyList = 8122   '�򿪼���б�
Public Const conMenu_Cap_StudySyncState = 8123  'Ӱ����ͬ��״̬
Public Const conMenu_Cap_RecordAudio = 8125     '¼��
Public Const conMenu_Cap_After_Capture = 8140   '��̨�ɼ�
Public Const conMenu_Cap_After_Record = 8141    '��̨¼��
Public Const conMenu_Cap_After_Tag = 8142       '���±��




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
Public Const conMenu_Process_RectCapture = 511         '�ü���ɼ�
Public Const conMenu_Process_Restore = 8124         '�ָ�



'�����˵�
Public Const conMenu_Help_Help = 901        '*��������(&H)
Public Const conMenu_Help_Web = 902         '&WEB�ϵ�����
Public Const conMenu_Help_Web_Home = 9021       '������ҳ(&H)
Public Const conMenu_Help_Web_Forum = 9023      '������̳(&F)
Public Const conMenu_Help_Web_Mail = 9022       '*���ͷ���(&M)
Public Const conMenu_Help_About = 991       '����(&A)��

'������������
'*********************************************************************
'CommandBar���г�������
Public Const XTP_ID_WINDOW_LIST = 35000    '�����б�
Public Const XTP_ID_TOOLBARLIST = 59392    '�������б�
Public Const ID_INDICATOR_CAPS = 59137    '״̬������д��
Public Const ID_INDICATOR_NUM = 59138    '״̬�������֣�
Public Const ID_INDICATOR_SCRL = 59139    '״̬����������

'CommandBar�����ȼ�
Public Const FSHIFT = 4
Public Const FCONTROL = 8
Public Const FALT = 16

'CommandBar�����
Public Const VK_BACK = &H8
Public Const VK_TAB = &H9
Public Const VK_ESCAPE = &H1B
Public Const VK_SPACE = &H20
Public Const VK_PRIOR = &H21
Public Const VK_NEXT = &H22
Public Const VK_END = &H23
Public Const VK_HOME = &H24
Public Const VK_LEFT = &H25
Public Const VK_UP = &H26
Public Const VK_RIGHT = &H27
Public Const VK_DOWN = &H28
Public Const VK_INSERT = &H2D
Public Const VK_DELETE = &H2E
Public Const VK_MULTIPLY = &H6A
Public Const VK_ADD = &H6B
Public Const VK_SEPARATOR = &H6C
Public Const VK_SUBTRACT = &H6D
Public Const VK_DECIMAL = &H6E
Public Const VK_DIVIDE = &H6F
Public Const VK_PAGEUP = &H21
Public Const VK_PAGEDOWN = &H22
Public Const VK_F1 = &H70
Public Const VK_F2 = &H71
Public Const VK_F3 = &H72
Public Const VK_F4 = &H73
Public Const VK_F5 = &H74
Public Const VK_F6 = &H75
Public Const VK_F7 = &H76
Public Const VK_F8 = &H77
Public Const VK_F9 = &H78
Public Const VK_F10 = &H79
Public Const VK_F11 = &H7A
Public Const VK_F12 = &H7B

Public Const VsModiBackColor = &HD6FFCA        'vs�ؼ����ɱ༭��Ԫ�ı���ɫ
'*********************************************************************


